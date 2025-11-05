#!/usr/bin/env python3
"""
ENHANCED RED FLAG EXTRACTION SYSTEM v2.0

New Features:
1. Multi-Model Ensemble (Gemini 2.0 Flash-Lite + Gemini 1.5 Flash + Gemini 1.5 Pro)
2. Context Window Enhancement (surrounding publications for context)
3. Confidence Scoring (model agreement + similarity + keyword density)
4. Named Entity Recognition (entities, organizations, locations)
5. Paragraph Context Improvement (expanded context extraction)
6. Multi-Language Support (French translation for AMF)
7. Batch Processing (parallel API calls)
"""

import os
import csv
import json
import re
import time
import numpy as np
import argparse
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Set
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
from dotenv import load_dotenv
import google.generativeai as genai
from tqdm import tqdm
from sklearn.metrics.pairwise import cosine_similarity
import spacy
from deep_translator import GoogleTranslator
import chromadb
from chromadb.config import Settings

# Load environment
load_dotenv()
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
if not GEMINI_API_KEY:
    raise ValueError("GEMINI_API_KEY not found in .env file")

genai.configure(api_key=GEMINI_API_KEY)

# Load spaCy for NER
try:
    nlp = spacy.load('en_core_web_sm')
except:
    print("⚠️  spaCy model not found. Installing...")
    os.system('python -m spacy download en_core_web_sm')
    nlp = spacy.load('en_core_web_sm')

# Constants
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
SCRAPED_CONTENT_DIR = BASE_DIR / 'scraped_content'
OUTPUT_DIR = BASE_DIR / 'output'
CHROMA_DIR = BASE_DIR / 'chroma_db'
OUTPUT_DIR.mkdir(exist_ok=True)
CHROMA_DIR.mkdir(exist_ok=True)

KEYWORDS_FILE = DATA_DIR / 'keywords.txt'
EXISTING_RED_FLAGS_FILE = DATA_DIR / 'existing_red_flags.csv'

# Multiple models for ensemble
MODEL_CONFIGS = [
    {'name': 'gemini-2.0-flash-lite', 'weight': 0.4, 'speed': 'fast'},
    {'name': 'gemini-1.5-flash', 'weight': 0.35, 'speed': 'fast'},
    {'name': 'gemini-1.5-pro', 'weight': 0.25, 'speed': 'slow'}  # Slower but higher quality
]

models = {cfg['name']: genai.GenerativeModel(cfg['name']) for cfg in MODEL_CONFIGS}
embedding_model = 'models/text-embedding-004'

# Initialize ChromaDB
chroma_client = chromadb.PersistentClient(
    path=str(CHROMA_DIR),
    settings=Settings(anonymized_telemetry=False)
)

# Global caches
existing_red_flags_embeddings = None
existing_red_flags_data = None
translator = GoogleTranslator(source='fr', target='en')


def load_keywords() -> List[str]:
    """Load keywords from file"""
    with open(KEYWORDS_FILE, 'r', encoding='utf-8') as f:
        keywords = [line.strip() for line in f if line.strip()]
    return keywords


def load_existing_red_flags() -> pd.DataFrame:
    """Load existing red flags from CSV"""
    df = pd.read_csv(EXISTING_RED_FLAGS_FILE, encoding='utf-8')
    df = df.dropna(subset=['Risk Indicator'])
    df['Risk Indicator'] = df['Risk Indicator'].str.strip()
    df = df[df['Risk Indicator'].str.len() > 10]
    return df


def get_embedding(text: str, retry_count: int = 3, use_cache: bool = True) -> Optional[np.ndarray]:
    """
    Get embedding vector for text with ChromaDB caching
    Args:
        text: Text to embed
        retry_count: Number of retries for API calls
        use_cache: Whether to use ChromaDB cache (default: True)
    """
    # Generate consistent hash for text
    text_hash = str(hash(text.strip()))
    
    # Try to get from cache first
    if use_cache:
        try:
            collection = chroma_client.get_or_create_collection(
                name="embeddings_cache",
                metadata={"hnsw:space": "cosine"}
            )
            
            # Check if embedding exists in cache
            results = collection.get(ids=[text_hash])
            if results['embeddings'] and len(results['embeddings']) > 0:
                return np.array(results['embeddings'][0])
        except Exception as e:
            print(f"Cache read error: {e}")
    
    # If not in cache or cache disabled, compute embedding
    for attempt in range(retry_count):
        try:
            result = genai.embed_content(
                model=embedding_model,
                content=text,
                task_type="semantic_similarity"
            )
            embedding = np.array(result['embedding'])
            
            # Store in cache for future use
            if use_cache:
                try:
                    collection.add(
                        ids=[text_hash],
                        embeddings=[embedding.tolist()],
                        documents=[text[:500]],  # Store truncated text for reference
                        metadatas=[{"timestamp": datetime.now().isoformat()}]
                    )
                except Exception as e:
                    print(f"Cache write error: {e}")
            
            return embedding
            
        except Exception as e:
            if attempt < retry_count - 1:
                time.sleep(1)
            else:
                print(f"Embedding error after {retry_count} attempts: {e}")
                return None


def compute_red_flag_embeddings(red_flags_df: pd.DataFrame, use_cache: bool = True) -> np.ndarray:
    """
    Compute embeddings for all existing red flags with ChromaDB caching
    This is much faster on subsequent runs as embeddings are cached
    """
    print("🧮 Computing embeddings for existing red flags...")
    embeddings = []
    
    # Try to get all from cache first for maximum speed
    if use_cache:
        try:
            collection = chroma_client.get_or_create_collection(
                name="red_flags_library",
                metadata={"hnsw:space": "cosine"}
            )
            
            # Check if we already have all red flags cached
            cache_count = collection.count()
            if cache_count == len(red_flags_df):
                print(f"✅ Found {cache_count} cached embeddings - loading instantly!")
                flags_list = red_flags_df['Risk Indicator'].tolist()
                
                for text in flags_list:
                    text_hash = str(hash(text.strip()))
                    results = collection.get(ids=[text_hash])
                    if results['embeddings'] and len(results['embeddings']) > 0:
                        embeddings.append(np.array(results['embeddings'][0]))
                    else:
                        # Shouldn't happen, but fallback
                        emb = get_embedding(text, use_cache=True)
                        embeddings.append(emb if emb is not None else np.zeros(768))
                
                return np.array(embeddings)
        except Exception as e:
            print(f"Cache check error: {e}, computing fresh...")
    
    # Process with caching (first run or cache miss)
    batch_size = 10
    flags_list = red_flags_df['Risk Indicator'].tolist()
    
    for i in tqdm(range(0, len(flags_list), batch_size), desc="Computing embeddings"):
        batch = flags_list[i:i+batch_size]
        
        for text in batch:
            emb = get_embedding(text, use_cache=use_cache)
            if emb is not None:
                embeddings.append(emb)
                
                # Also store in red_flags_library collection
                if use_cache:
                    try:
                        collection = chroma_client.get_or_create_collection(
                            name="red_flags_library",
                            metadata={"hnsw:space": "cosine"}
                        )
                        text_hash = str(hash(text.strip()))
                        collection.upsert(
                            ids=[text_hash],
                            embeddings=[emb.tolist()],
                            documents=[text],
                            metadatas=[{"source": "existing_library"}]
                        )
                    except Exception as e:
                        pass  # Don't fail on cache write errors
            else:
                embeddings.append(np.zeros(768))
            time.sleep(0.05)  # Reduced rate limiting with batching
    
    return np.array(embeddings)


def extract_entities(text: str) -> Dict[str, List[str]]:
    """
    Extract named entities using spaCy NER
    Returns: Dict with entity types and their values
    """
    doc = nlp(text[:10000])  # Limit text length for performance
    
    entities = {
        'PERSON': [],
        'ORG': [],
        'GPE': [],  # Geopolitical entities (countries, cities)
        'MONEY': [],
        'DATE': []
    }
    
    for ent in doc.ents:
        if ent.label_ in entities:
            entities[ent.label_].append(ent.text)
    
    # Deduplicate and limit
    for key in entities:
        entities[key] = list(set(entities[key]))[:10]  # Top 10 unique entities
    
    return entities


def translate_french_content(text: str, source: str) -> str:
    """Translate French content to English (for AMF source)"""
    if source.upper() != 'AMF':
        return text
    
    try:
        # Detect if content has significant French text
        french_indicators = ['le ', 'la ', 'les ', 'des ', 'une ', 'du ', 'de la ', 'autorité']
        french_count = sum(1 for indicator in french_indicators if indicator in text.lower())
        
        if french_count < 3:  # Already mostly English
            return text
        
        # Translate in chunks (API limit is 5000 chars)
        chunk_size = 4500
        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        
        translated_chunks = []
        for chunk in chunks[:5]:  # Limit to first 5 chunks to avoid rate limits
            try:
                translated = translator.translate(chunk)
                translated_chunks.append(translated)
                time.sleep(0.5)  # Rate limiting
            except:
                translated_chunks.append(chunk)  # Keep original on error
        
        return ' '.join(translated_chunks)
    except Exception as e:
        print(f"  ⚠️  Translation failed: {str(e)[:50]}")
        return text


def get_surrounding_context(publications: List[Dict], current_idx: int, window: int = 2) -> str:
    """
    Get context from surrounding publications (same source, nearby dates)
    Helps understand broader regulatory trends
    """
    current_pub = publications[current_idx]
    context_pubs = []
    
    # Get publications from same source within window
    for i in range(max(0, current_idx - window), min(len(publications), current_idx + window + 1)):
        if i == current_idx:
            continue
        
        pub = publications[i]
        if pub['source'] == current_pub['source']:
            context_pubs.append(pub)
    
    if not context_pubs:
        return ""
    
    # Build context summary
    context_text = "Related publications context:\n"
    for pub in context_pubs[:3]:  # Max 3 for brevity
        context_text += f"- {pub['date']}: {pub['title'][:80]}...\n"
    
    return context_text


def get_expanded_paragraph_context(content: str, paragraph: str, context_sentences: int = 3) -> str:
    """
    Extract expanded context around the paragraph
    Includes surrounding sentences for better understanding
    """
    if not paragraph or paragraph not in content:
        return paragraph
    
    # Find paragraph position
    start_idx = content.find(paragraph)
    if start_idx == -1:
        return paragraph
    
    # Get text before and after
    before_text = content[:start_idx]
    after_text = content[start_idx + len(paragraph):]
    
    # Split into sentences (rough approximation)
    before_sentences = re.split(r'[.!?]\s+', before_text)[-context_sentences:]
    after_sentences = re.split(r'[.!?]\s+', after_text)[:context_sentences]
    
    # Combine
    expanded = ' '.join(before_sentences) + ' ' + paragraph + ' ' + ' '.join(after_sentences)
    expanded = ' '.join(expanded.split())  # Clean whitespace
    
    return expanded[:800]  # Limit length


def find_similar_red_flags(extracted_text: str, similarity_threshold: float = 0.60) -> List[Dict]:
    """Find similar existing red flags using semantic similarity"""
    global existing_red_flags_embeddings, existing_red_flags_data
    
    if existing_red_flags_embeddings is None or existing_red_flags_data is None:
        return []
    
    query_emb = get_embedding(extracted_text)
    if query_emb is None:
        return []
    
    query_emb = query_emb.reshape(1, -1)
    similarities = cosine_similarity(query_emb, existing_red_flags_embeddings)[0]
    
    matches = []
    for idx, score in enumerate(similarities):
        if score >= similarity_threshold:
            red_flag_id = int(existing_red_flags_data.iloc[idx]['ID'])
            red_flag_text = existing_red_flags_data.iloc[idx]['Risk Indicator']
            matches.append({
                'id': red_flag_id,
                'similarity': float(score),
                'text': red_flag_text
            })
    
    matches.sort(key=lambda x: x['similarity'], reverse=True)
    return matches


def generalize_red_flag(specific_red_flag: str) -> str:
    """Convert specific red flag to generic description"""
    prompt = f"""Convert this specific red flag into a generic, reusable red flag pattern description.

Specific red flag: "{specific_red_flag}"

Make it generic by:
1. Remove specific entities/names (e.g., "Houthis" → "entity", "Sinaloa Cartel" → "criminal organization")
2. Focus on the behavioral pattern, not the specific case
3. Keep it concise (under 100 characters)
4. Make it applicable to similar situations

Return ONLY the generic description, no other text."""

    try:
        response = models['gemini-2.0-flash-lite'].generate_content(
            prompt,
            generation_config={'temperature': 0.3, 'max_output_tokens': 150},
            request_options={'timeout': 10}
        )
        generic = response.text.strip().strip('"\'')
        return generic if len(generic) < 150 else specific_red_flag[:100]
    except:
        return specific_red_flag[:100]


def calculate_confidence_score(red_flag: str, paragraph: str, keywords: List[str], 
                               model_agreement: float, similarity_score: float) -> float:
    """
    Calculate confidence score (0-100) based on multiple factors:
    1. Model agreement (0-1): How many models agreed on this extraction
    2. Similarity score (0-1): Highest similarity to existing red flags
    3. Keyword density (0-1): Ratio of keywords in paragraph
    4. Entity presence (0-1): Presence of named entities
    5. Length quality (0-1): Optimal length score
    """
    scores = []
    
    # 1. Model agreement (weight: 35%)
    scores.append(('model_agreement', model_agreement * 35))
    
    # 2. Similarity score (weight: 25%)
    scores.append(('similarity', similarity_score * 25))
    
    # 3. Keyword density (weight: 20%)
    para_lower = paragraph.lower()
    keyword_count = sum(1 for kw in keywords if kw.lower() in para_lower)
    keyword_density = min(keyword_count / 5, 1.0)  # Max at 5 keywords
    scores.append(('keywords', keyword_density * 20))
    
    # 4. Entity presence (weight: 10%)
    entities = extract_entities(red_flag)
    entity_score = min(sum(len(v) for v in entities.values()) / 3, 1.0)
    scores.append(('entities', entity_score * 10))
    
    # 5. Length quality (weight: 10%)
    length = len(red_flag)
    if 50 <= length <= 120:
        length_score = 1.0
    elif 30 <= length <= 150:
        length_score = 0.7
    else:
        length_score = 0.4
    scores.append(('length', length_score * 10))
    
    total_score = sum(s[1] for s in scores)
    return round(total_score, 2)


def extract_red_flags_ensemble(content: str, title: str, context: str = "") -> Dict:
    """
    Extract red flags using ensemble of multiple Gemini models
    Returns consensus with confidence scores
    """
    
    base_prompt = f"""You are an expert AML analyst. Extract ONLY behavioral red flags from this document.

**Document:** {title}

{context}

**Content:**
{content[:3500]}

**CRITICAL INSTRUCTIONS:**
1. Extract ONLY suspicious behaviors, transaction patterns, or criminal activities
2. DO NOT extract: compliance requirements, regulatory obligations, procedural rules, or generic statements
3. Each red flag must be a complete sentence describing WHAT someone does that is suspicious
4. Focus on: unusual transactions, suspicious patterns, evasion tactics, laundering behaviors, fraud schemes
5. Category must be EXACTLY "AML" or "Transaction Patterns"
6. Keep red flags concise (under 150 characters preferred)

**GOOD Examples (extract these types):**
- "Customer makes frequent overseas transfers not in line with their financial profile."
- "Multiple transactions conducted below the reporting threshold within a short period."
- "Funds deposited into several accounts and then consolidated before transferring abroad."

**BAD Examples (DO NOT extract):**
- "Must report suspicious transactions" (requirement, not behavior)
- "Failure to implement compliance program" (compliance issue, not suspicious behavior)
- "Register as money services business" (regulatory requirement)

**JSON Format:**
[
  {{
    "red_flag": "Complete behavioral red flag sentence",
    "paragraph": "Paragraph containing this red flag with surrounding context",
    "category": "AML" or "Transaction Patterns"
  }}
]

Return ONLY valid JSON array."""

    generation_config = {
        'temperature': 0.25,
        'top_p': 0.85,
        'top_k': 15,
        'max_output_tokens': 2048,
    }
    
    # Call all models in parallel
    def call_model(model_name):
        try:
            model = models[model_name]
            response = model.generate_content(
                base_prompt,
                generation_config=generation_config,
                request_options={'timeout': 30}
            )
            result_text = response.text.strip()
            
            # Extract JSON
            if '```json' in result_text:
                result_text = result_text.split('```json')[1].split('```')[0].strip()
            elif '```' in result_text:
                result_text = result_text.split('```')[1].split('```')[0].strip()
            
            red_flags = json.loads(result_text)
            return {'model': model_name, 'success': True, 'red_flags': red_flags if isinstance(red_flags, list) else []}
        except Exception as e:
            return {'model': model_name, 'success': False, 'red_flags': [], 'error': str(e)[:50]}
    
    # Parallel execution
    results_by_model = []
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {executor.submit(call_model, cfg['name']): cfg for cfg in MODEL_CONFIGS}
        
        for future in as_completed(futures):
            result = future.result()
            results_by_model.append(result)
    
    # Aggregate results - find consensus
    all_red_flags = {}
    
    for result in results_by_model:
        if not result['success']:
            continue
        
        model_weight = next(cfg['weight'] for cfg in MODEL_CONFIGS if cfg['name'] == result['model'])
        
        for rf in result['red_flags']:
            if not isinstance(rf, dict):
                continue
            
            rf_text = rf.get('red_flag', '').strip()
            if len(rf_text) < 25:
                continue
            
            # Normalize for matching
            rf_key = rf_text.lower()[:100]
            
            if rf_key not in all_red_flags:
                all_red_flags[rf_key] = {
                    'red_flag': rf_text,
                    'paragraph': rf.get('paragraph', ''),
                    'category': rf.get('category', 'AML'),
                    'models_agreed': [result['model']],
                    'agreement_weight': model_weight
                }
            else:
                all_red_flags[rf_key]['models_agreed'].append(result['model'])
                all_red_flags[rf_key]['agreement_weight'] += model_weight
    
    # Filter by consensus (at least 2 models or high weight model)
    consensus_flags = []
    for rf_data in all_red_flags.values():
        num_models = len(rf_data['models_agreed'])
        agreement_weight = rf_data['agreement_weight']
        
        # Keep if: 2+ models agreed OR single high-weight model (>=0.35)
        if num_models >= 2 or agreement_weight >= 0.35:
            # Calculate model agreement score (0-1)
            model_agreement = min(agreement_weight / 0.6, 1.0)  # Normalize to max 0.6 weight
            
            rf_data['model_agreement'] = model_agreement
            consensus_flags.append(rf_data)
    
    return {
        'success': True,
        'red_flags': consensus_flags,
        'models_used': len(results_by_model),
        'error': None
    }


def clean_extracted_text(text: str) -> str:
    """Clean and validate extracted text"""
    if not text or not isinstance(text, str):
        return ""
    
    text = ' '.join(text.split())
    text = re.sub(r'^\d+[\.\)]\s*', '', text)
    text = re.sub(r'\s*\d+$', '', text)
    
    if text and not text.endswith(('.', '!', '?')):
        text += '.'
    
    if text:
        text = text[0].upper() + text[1:]
    
    return text.strip()


def validate_category(category: str) -> str:
    """Ensure category is only 'AML' or 'Transaction Patterns'"""
    if not category or not isinstance(category, str):
        return "AML"
    
    category_clean = category.strip().upper()
    
    if any(word in category_clean for word in ['TRANSACTION', 'PATTERN', 'FLOW', 'STRUCTUR', 'TRANSFER']):
        return "Transaction Patterns"
    else:
        return "AML"


def has_relevant_keywords(content: str, keywords: List[str], threshold: int = 3) -> bool:
    """Check if content contains at least threshold number of keywords"""
    content_lower = content.lower()
    count = sum(1 for kw in keywords if kw.lower() in content_lower)
    return count >= threshold


def determine_coverage_status(matches: List[Dict], extracted_red_flag: str, 
                              high_threshold: float = 0.75) -> Tuple[str, str, List[int]]:
    """Determine coverage status and format coverage column"""
    if not matches:
        generic_flag = generalize_red_flag(extracted_red_flag)
        coverage_text = f"New: {generic_flag}"
        return (coverage_text, "Not Covered", [])
    
    high_matches = [m for m in matches if m['similarity'] >= high_threshold]
    
    if high_matches:
        matched_ids = [m['id'] for m in high_matches[:3]]
        coverage_text = ', '.join(str(id_) for id_ in matched_ids)
        return (coverage_text, "Fully Covered", matched_ids)
    else:
        matched_ids = [m['id'] for m in matches[:3]]
        generic_flag = generalize_red_flag(extracted_red_flag)
        coverage_text = f"{', '.join(str(id_) for id_ in matched_ids)}, New: {generic_flag}"
        return (coverage_text, "Partially Covered", matched_ids)


def get_all_publications(year: int, month: int) -> List[Dict]:
    """Gather all publications from scraped content"""
    publications = []
    scrapers = ['amf', 'dfs', 'fca', 'fed', 'fincen', 'fintrac', 'nca', 'ofac', 'sec']
    
    for scraper in scrapers:
        scraper_dir = SCRAPED_CONTENT_DIR / scraper / f"{year}-{month:02d}"
        if not scraper_dir.exists():
            continue
        
        manifest_file = scraper_dir / 'manifest.csv'
        if not manifest_file.exists():
            continue
        
        with open(manifest_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                content = ""
                
                # Check html_text folder
                html_dir = scraper_dir / 'html_text'
                if html_dir.exists():
                    for file in html_dir.glob('*.txt'):
                        if row['Date'] in file.name:
                            with open(file, 'r', encoding='utf-8', errors='ignore') as cf:
                                content = cf.read()
                            break
                
                # Check pdf_text folder
                if not content:
                    pdf_dir = scraper_dir / 'pdf_text'
                    if pdf_dir.exists():
                        for file in pdf_dir.glob('*.txt'):
                            if row['Date'] in file.name:
                                with open(file, 'r', encoding='utf-8', errors='ignore') as cf:
                                    content = cf.read()
                                break
                
                if content:
                    publications.append({
                        'source': scraper.upper(),
                        'date': row['Date'],
                        'title': row['Publication Title'],
                        'link': row['Publication Link'],
                        'content': content
                    })
    
    return publications


def process_publications(publications: List[Dict], keywords: List[str]) -> List[Dict]:
    """Process publications with all enhancements"""
    results = []
    global_seen_flags: Set[str] = set()
    
    # Filter by keyword relevance
    print(f"\n🔍 Filtering by keyword relevance...")
    relevant_pubs = [p for p in publications if has_relevant_keywords(p['content'], keywords, threshold=3)]
    print(f"  ✓ {len(relevant_pubs)}/{len(publications)} publications contain 3+ keywords")
    
    # Translate AMF content
    print(f"\n🌐 Translating French content...")
    for pub in relevant_pubs:
        if pub['source'] == 'AMF':
            pub['content'] = translate_french_content(pub['content'], pub['source'])
    print(f"  ✓ Translation complete")
    
    print(f"\n🤖 Analyzing {len(relevant_pubs)} relevant publications with multi-model ensemble...\n")
    
    for idx, pub in enumerate(tqdm(relevant_pubs, desc="Processing publications")):
        print(f"\n{'='*80}")
        print(f"Source: {pub['source']} | Date: {pub['date']}")
        print(f"Title: {pub['title'][:100]}...")
        
        # Get surrounding context
        context = get_surrounding_context(publications, idx, window=2)
        
        # Extract with ensemble
        try:
            analysis = extract_red_flags_ensemble(pub['content'], pub['title'], context)
        except KeyboardInterrupt:
            print(f"\n⚠️  Interrupted. {len(results)} results collected.")
            raise
        except Exception as e:
            print(f"  ⚠️  Skipped: {str(e)[:60]}")
            continue
        
        if not analysis or not analysis['success']:
            print(f"  ⚠️  No valid extraction")
            continue
        
        red_flags_found = analysis['red_flags']
        print(f"  ✓ Extracted {len(red_flags_found)} red flags (consensus from {analysis['models_used']} models)")
        
        added_count = 0
        for rf in red_flags_found:
            # Clean text
            red_flag_text = clean_extracted_text(rf['red_flag'])
            paragraph_text = clean_extracted_text(rf['paragraph'])
            category = validate_category(rf['category'])
            
            # Validation filters
            if len(red_flag_text) < 25:
                continue
            if not red_flag_text.endswith(('.', '!', '?')):
                continue
            
            # Filter compliance language
            skip_phrases = ['must ', 'should ', 'required to', 'failure to', 'obligat', 
                          'register as', 'implement a']
            if any(phrase in red_flag_text.lower() for phrase in skip_phrases):
                continue
            
            # Global deduplication
            flag_key = red_flag_text.lower().strip()
            if flag_key in global_seen_flags:
                continue
            global_seen_flags.add(flag_key)
            
            # Expand paragraph context
            expanded_paragraph = get_expanded_paragraph_context(pub['content'], paragraph_text, context_sentences=2)
            
            # Extract entities
            entities = extract_entities(red_flag_text)
            entities_str = ', '.join([f"{k}:{','.join(v[:2])}" for k, v in entities.items() if v])
            
            # Find similar flags
            matches = find_similar_red_flags(red_flag_text, similarity_threshold=0.60)
            best_similarity = matches[0]['similarity'] if matches else 0.0
            
            # Calculate confidence score
            confidence = calculate_confidence_score(
                red_flag_text, 
                expanded_paragraph, 
                keywords, 
                rf.get('model_agreement', 0.5),
                best_similarity
            )
            
            # Determine coverage
            coverage_text, coverage_status, matched_ids = determine_coverage_status(
                matches, red_flag_text, high_threshold=0.75
            )
            
            result_row = {
                'Source Link': pub['link'],
                'Date': pub['date'],
                'Extracted Red Flag': red_flag_text,
                'Associated Paragraph': expanded_paragraph,
                'Category': category,
                'Coverage by Existing Red Flag': coverage_text,
                'Coverage Status': coverage_status,
                'Confidence Score': confidence,
                'Named Entities': entities_str[:200] if entities_str else 'None',
                'Model Agreement': f"{rf.get('model_agreement', 0.5):.2f}",
                'Similarity Score': f"{best_similarity:.2f}"
            }
            results.append(result_row)
            added_count += 1
        
        if added_count > 0:
            print(f"  ✓ Added {added_count} unique red flags")
    
    return results


def export_to_excel(results: List[Dict], year: int, month: int) -> str:
    """Export results to Excel with enhanced columns"""
    output_file = OUTPUT_DIR / f"red_flags_analysis_enhanced_{year}_{month:02d}.xlsx"
    
    df = pd.DataFrame(results)
    
    # Column order - original 7 + new analytics columns
    column_order = [
        'Source Link',
        'Date',
        'Extracted Red Flag',
        'Associated Paragraph',
        'Category',
        'Coverage by Existing Red Flag',
        'Coverage Status',
        'Confidence Score',
        'Named Entities',
        'Model Agreement',
        'Similarity Score'
    ]
    
    df = df[column_order]
    
    # Sort by confidence score (highest first)
    df = df.sort_values('Confidence Score', ascending=False)
    
    # Export to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Red Flags Analysis', index=False)
        
        # Adjust column widths
        worksheet = writer.sheets['Red Flags Analysis']
        worksheet.column_dimensions['A'].width = 50  # Source Link
        worksheet.column_dimensions['B'].width = 12  # Date
        worksheet.column_dimensions['C'].width = 60  # Extracted Red Flag
        worksheet.column_dimensions['D'].width = 80  # Associated Paragraph
        worksheet.column_dimensions['E'].width = 20  # Category
        worksheet.column_dimensions['F'].width = 40  # Coverage
        worksheet.column_dimensions['G'].width = 18  # Coverage Status
        worksheet.column_dimensions['H'].width = 15  # Confidence Score
        worksheet.column_dimensions['I'].width = 40  # Named Entities
        worksheet.column_dimensions['J'].width = 15  # Model Agreement
        worksheet.column_dimensions['K'].width = 15  # Similarity Score
    
    return str(output_file)


def main():
    parser = argparse.ArgumentParser(description='Enhanced red flag extraction with ensemble models')
    parser.add_argument('--year', type=int, required=True, help='Year (e.g., 2025)')
    parser.add_argument('--month', type=int, required=True, help='Month (1-12)')
    args = parser.parse_args()
    
    year = args.year
    month = args.month
    
    print("="*80)
    print("ENHANCED RED FLAG EXTRACTION SYSTEM v2.0")
    print("Multi-Model Ensemble + NER + Context Enhancement")
    print("="*80)
    print(f"Models: {', '.join(cfg['name'] for cfg in MODEL_CONFIGS)}")
    print(f"Target: {year}-{month:02d}")
    print(f"Output: {OUTPUT_DIR}")
    print("="*80)
    
    # Load resources
    print("\n📚 Loading resources...")
    keywords = load_keywords()
    print(f"  ✓ Loaded {len(keywords)} keywords")
    
    global existing_red_flags_data, existing_red_flags_embeddings
    existing_red_flags_data = load_existing_red_flags()
    print(f"  ✓ Loaded {len(existing_red_flags_data)} existing red flags")
    
    # Compute embeddings
    existing_red_flags_embeddings = compute_red_flag_embeddings(existing_red_flags_data)
    print(f"  ✓ Computed semantic embeddings")
    
    # Gather publications
    print("\n📄 Gathering publications...")
    publications = get_all_publications(year, month)
    print(f"  ✓ Found {len(publications)} publications")
    
    if len(publications) == 0:
        print("\n⚠️  No publications found.")
        return
    
    # Show source breakdown
    sources = {}
    for pub in publications:
        sources[pub['source']] = sources.get(pub['source'], 0) + 1
    print("\n  By source:")
    for source, count in sorted(sources.items()):
        print(f"    • {source}: {count}")
    
    # Process with enhancements
    results = process_publications(publications, keywords)
    
    if not results:
        print("\n⚠️  No red flags extracted.")
        return
    
    # Export
    print("\n📊 Exporting to Excel...")
    output_file = export_to_excel(results, year, month)
    
    # Summary statistics
    fully_covered = sum(1 for r in results if r['Coverage Status'] == 'Fully Covered')
    partially_covered = sum(1 for r in results if r['Coverage Status'] == 'Partially Covered')
    not_covered = sum(1 for r in results if r['Coverage Status'] == 'Not Covered')
    
    avg_confidence = sum(r['Confidence Score'] for r in results) / len(results)
    high_confidence = sum(1 for r in results if r['Confidence Score'] >= 70)
    
    print("\n" + "="*80)
    print("✅ EXTRACTION COMPLETE")
    print("="*80)
    print(f"Total Red Flags: {len(results)}")
    print(f"\nCoverage Breakdown:")
    print(f"  • Fully Covered: {fully_covered}")
    print(f"  • Partially Covered: {partially_covered}")
    print(f"  • Not Covered: {not_covered}")
    print(f"\nConfidence Metrics:")
    print(f"  • Average Confidence: {avg_confidence:.1f}/100")
    print(f"  • High Confidence (≥70): {high_confidence} ({high_confidence/len(results)*100:.0f}%)")
    print(f"\nOutput: {output_file}")
    print("="*80)


if __name__ == "__main__":
    main()
