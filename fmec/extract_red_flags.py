#!/usr/bin/env python3
"""
Comprehensive Red Flag Extraction System using Gemini 2.5 Pro
Processes all 823 documents from 9 regulatory sources
Extracts: Red Flag, Associated Paragraph, Category, Coverage by Existing Red Flags
"""

import os
import json
import time
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv
from tqdm import tqdm

# Load environment variables
load_dotenv()

# Configure Gemini
genai.configure(api_key=os.getenv('GEMINI_API_KEY'))

# Use Gemini 2.5 Flash (fast and efficient)
MODEL_NAME = "gemini-2.0-flash-exp"  # Gemini 2.5 Flash

def categorize_keywords():
    """Categorize 166 keywords into AML and Transaction Patterns"""
    
    # Read keywords
    with open('data/keywords.txt', 'r', encoding='utf-8') as f:
        keywords = [line.strip() for line in f if line.strip()]
    
    prompt = f"""
You are an AML (Anti-Money Laundering) expert. Categorize these {len(keywords)} keywords into TWO categories:

1. **AML_KEYWORDS**: Keywords related to regulations, compliance, entities, and regulatory concepts
2. **TRANSACTION_PATTERNS**: Keywords related to specific transaction behaviors, patterns, and red flags

Keywords to categorize:
{chr(10).join([f"{i+1}. {kw}" for i, kw in enumerate(keywords)])}

Return ONLY a JSON object with this exact structure:
{{
    "AML_KEYWORDS": ["keyword1", "keyword2", ...],
    "TRANSACTION_PATTERNS": ["pattern1", "pattern2", ...]
}}

Ensure every keyword is assigned to exactly ONE category.
"""
    
    try:
        model = genai.GenerativeModel(MODEL_NAME)
        response = model.generate_content(prompt)
        
        # Parse JSON from response
        response_text = response.text.strip()
        if response_text.startswith('```json'):
            response_text = response_text[7:]
        if response_text.endswith('```'):
            response_text = response_text[:-3]
        
        categorized = json.loads(response_text.strip())
        
        # Save categorized keywords
        with open('data/categorized_keywords.json', 'w', encoding='utf-8') as f:
            json.dump(categorized, f, indent=2, ensure_ascii=False)
        
        print(f"✓ Categorized {len(keywords)} keywords:")
        print(f"  - AML Keywords: {len(categorized['AML_KEYWORDS'])}")
        print(f"  - Transaction Patterns: {len(categorized['TRANSACTION_PATTERNS'])}")
        
        return categorized
    
    except Exception as e:
        print(f"✗ Error categorizing keywords: {e}")
        return None

def create_extraction_prompt(
    document_content: str,
    existing_red_flags: List[Dict],
    aml_keywords: List[str],
    transaction_patterns: List[str],
    source_link: str = "",
    doc_date: str = ""
) -> str:
    """Create the perfect prompt for Gemini 2.5 Pro"""
    
    # Show ALL existing red flags for proper matching
    red_flags_summary = "\n".join([
        f"ID {rf.get('ID', 'N/A')}: {rf['Risk Indicator']}"
        for rf in existing_red_flags  # Show ALL, not just sample
    ])
    
    prompt = f"""
You are an expert AML (Anti-Money Laundering) analyst specializing in identifying financial crime red flags from regulatory documents.

**YOUR TASK:**
Analyze the regulatory document below and extract ALL AML red flags, suspicious activity indicators, and money laundering typologies mentioned.

**DOCUMENT TO ANALYZE:**
---
{document_content[:50000]}  # Limit to ~50K chars to fit in context
---

**REFERENCE MATERIALS:**

**EXISTING RED FLAGS DATABASE - YOU MUST COMPARE AGAINST THESE:**
{red_flags_summary}

**MATCHING EXAMPLES** (how to map document content to existing IDs - USE VARIETY):
- "Forex/foreign exchange" → ID 52
- "Structuring/avoiding thresholds" → ID 261, 847, 850, 1354
- "High-risk countries/jurisdictions" → ID 306, 368, 871, 1388
- "Wire transfers/remittances" → ID 63, 857, 860, 862, 865, 866, 1295
- "Cash deposits/withdrawals" → ID 1, 27, 374
- "Account activity inconsistent" → ID 6, 816, 819, 829, 1460
- "Third party transactions" → ID 317, 463, 951, 1271, 1272
- "Hidden relationships" → ID 63, 966, 1036, 1115, 1625
- "Suspicious outflows" → ID 1691
- "Business transactions" → ID 245, 1057
- "Cards/ATM anomalies" → ID 83, 984, 1082
- "Flow through/layering" → ID 1, 317, 437, 644, 687, 863, 1335
- "Frequent transfers" → ID 829, 1129, 1309

**IMPORTANT**: Don't default to ID 52 or 1691 for everything! Use the full range of IDs above!

**AML Keywords to Look For:**
{', '.join(aml_keywords[:50])}

**Transaction Pattern Keywords to Look For:**
{', '.join(transaction_patterns[:50])}

**EXTRACTION RULES:**

1. **Multiple Paragraphs**: Scan the ENTIRE document. Each paragraph can contain multiple red flags.

2. **What to Extract**:
   - Specific transaction patterns described
   - Behavioral red flags mentioned
   - Money laundering typologies explained
   - Suspicious activity indicators listed
   - Case study examples of illicit behavior
   - Regulatory enforcement findings about AML failures

3. **What NOT to Extract**:
   - Generic compliance requirements
   - Background information about the organization
   - Procedural steps without red flag context
   - Standard definitions without suspicious indicators

4. **Coverage Analysis**: For each red flag found, determine if it's:
   - **COVERED**: Already represented in existing red flags database (similar meaning)
   - **NEW**: Not covered by existing red flags (novel indicator)
   - **PARTIAL**: Partially covered but adds new nuance

**OUTPUT FORMAT:**

Return a JSON array where each element represents ONE red flag extracted:

[
  {{
    "source_link": "URL from document if available",
    "extracted_red_flag": "COPY one complete sentence verbatim from the paragraph below (the key red flag sentence)",
    "associated_paragraph": "COPY the complete paragraph containing the sentence above (must include the extracted_red_flag as a substring)",
    "category": "AML or Transaction Patterns (ONLY these two options)",
    "coverage_by_existing": "Comma-separated IDs from existing red flags if similar (e.g., '860, 1335') OR 'New: Generic red flag description' if completely new. Can have multiple like '860, New: Description, 1335'",
    "coverage_status": "Fully Covered (all IDs) | Partially Covered (mix of IDs and New) | Not Covered (all New)"
  }},
  ...
]

**CRITICAL EXTRACTION RULES:**

1. **extracted_red_flag** (MUST BE A SENTENCE FROM THE PARAGRAPH):
   - Extract ONE complete sentence or key phrase (10-25 words)
   - MUST be copied VERBATIM from the associated_paragraph
   - Should be the most important sentence that indicates the red flag
   - Examples: "Failed to verify customer identity before opening accounts", "Large cash deposits inconsistent with business profile"

2. **associated_paragraph** (MUST BE THE COMPLETE PARAGRAPH):
   - Copy the ENTIRE paragraph where you found the extracted_red_flag sentence (50+ words)
   - The extracted_red_flag MUST appear as a substring within this paragraph
   - Include ALL surrounding sentences for full context
   - This should be the complete regulatory text, enforcement finding, or violation description

3. **category** (STRICT):
   - ONLY two allowed values: "AML" or "Transaction Patterns"
   - AML = regulatory, compliance, entity-related risks
   - Transaction Patterns = specific transaction behaviors and patterns

4. **coverage_by_existing** (CRITICALLY IMPORTANT - USE DIVERSE IDS):
   - **MATCH BROADLY ACROSS ALL 50 IDs**: Look through the ENTIRE database, not just 52 and 1691
   - **Think conceptually**:
     * If document mentions cash → consider IDs 1, 27, 261, 374, 847, 850, 1354
     * If mentions accounts → consider IDs 6, 245, 816, 819, 829, 1057, 1460
     * If mentions transfers → consider IDs 63, 317, 437, 644, 857, 860, 862, 863, 865, 866, 1295, 1309, 1335
     * If mentions geography → consider IDs 83, 306, 368, 871, 965, 1388
     * If mentions relationships → consider IDs 63, 951, 966, 1036, 1115, 1271, 1272, 1625
   - **USE MULTIPLE IDs**: Many red flags match 3-5 existing IDs, not just 1-2
   - Format examples:
     * "6, 816, 1460" (account inconsistencies)
     * "261, 847, 850, 1354, New: Crypto structuring"
     * "306, 368, 871, 1388" (high-risk jurisdictions)
   - **FORBIDDEN**: Stop using only ID 52 and 1691! Use the FULL range of available IDs!

5. **coverage_status** (AUTO-LOGIC):
   - "Fully Covered" = ALL entries are numeric IDs
   - "Partially Covered" = BOTH IDs AND "New:" entries
   - "Not Covered" = ONLY "New:" entries

Return ONLY valid JSON array, no markdown formatting, no explanations.
"""
    
    return prompt

def get_source_link_from_manifest(doc_path: Path) -> str:
    """Get the publication link from manifest.csv for a document"""
    try:
        # Get the organization and date folders
        org_folder = doc_path.parts[-4] if len(doc_path.parts) >= 4 else None
        date_folder = doc_path.parts[-3] if len(doc_path.parts) >= 3 else None
        
        if not org_folder or not date_folder:
            return ""
        
        # Construct manifest path
        manifest_path = Path('scraped_content') / org_folder / date_folder / 'manifest.csv'
        
        if not manifest_path.exists():
            return ""
        
        # Read manifest
        df_manifest = pd.read_csv(manifest_path)
        
        # Match by filename (the document filename should contain the title or date)
        doc_filename = doc_path.stem  # without extension
        
        # Try to find matching row
        for _, row in df_manifest.iterrows():
            if 'Publication Link' in row and pd.notna(row['Publication Link']):
                # Simple match - if this is the first/only match, use it
                # You could add more sophisticated matching here
                return str(row['Publication Link'])
        
        return ""
    except Exception as e:
        return ""

def clean_document_content(content: str) -> str:
    """Remove metadata headers, labels, and duplicate lines from document"""
    import re
    
    lines = content.split('\n')
    cleaned_lines = []
    seen_lines = set()
    
    # Patterns to skip (metadata headers)
    skip_patterns = [
        r'^Warning\s*$',
        r'^Savings protection\s*$',
        r'^News\s*$',
        r'^Press Release\s*$',
        r'^Enforcement\s*$',
        r'^\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\s*$',
        r'^\d{4}-\d{2}-\d{2}\s*$',
        r'^Published:\s*',
        r'^Date:\s*',
    ]
    
    for line in lines:
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
        
        # Skip metadata headers
        if any(re.match(pattern, line, re.IGNORECASE) for pattern in skip_patterns):
            continue
        
        # Skip very short lines (likely headers)
        if len(line) < 20 and line[0].isupper() and not line.endswith('.'):
            continue
        
        # Skip duplicate lines (exact duplicates)
        if line.lower() in seen_lines:
            continue
        
        seen_lines.add(line.lower())
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def extract_red_flags_from_document(
    doc_path: Path,
    existing_red_flags: List[Dict],
    aml_keywords: List[str],
    transaction_patterns: List[str],
    retry_count: int = 2
) -> List[Dict]:
    """Extract red flags from a single document using Gemini"""
    
    try:
        # Read document content
        with open(doc_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Skip if too short
        if len(content) < 200:
            return []
        
        # Clean document content (remove metadata headers and duplicates)
        content = clean_document_content(content)
        
        # Skip if cleaned content is too short
        if len(content) < 100:
            return []
        
        # Get source link from manifest
        source_link = get_source_link_from_manifest(doc_path)
        doc_date = doc_path.parts[-3] if len(doc_path.parts) >= 3 else 'unknown'
        
        # Create prompt
        prompt = create_extraction_prompt(content, existing_red_flags, aml_keywords, transaction_patterns, source_link, doc_date)
        
        # Call Gemini
        model = genai.GenerativeModel(MODEL_NAME)
        
        for attempt in range(retry_count):
            try:
                response = model.generate_content(
                    prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.1,  # Low temperature for consistent extraction
                        max_output_tokens=8000,
                    )
                )
                
                # Parse JSON response
                response_text = response.text.strip()
                if response_text.startswith('```json'):
                    response_text = response_text[7:]
                if response_text.endswith('```'):
                    response_text = response_text[:-3]
                
                red_flags = json.loads(response_text.strip())
                
                # Add metadata and rename keys to match desired format
                for rf in red_flags:
                    # Use the source link we got from manifest
                    rf['Source Link'] = source_link if source_link else rf.get('source_link', '')
                    rf['Date'] = doc_date
                    rf['Extracted Red Flag'] = rf.pop('extracted_red_flag', '')
                    rf['Associated Paragraph'] = rf.pop('associated_paragraph', '')
                    rf['Category'] = rf.pop('category', 'AML')
                    rf['Coverage by Existing Red Flag'] = rf.pop('coverage_by_existing', 'New: Unspecified risk')
                    rf['Coverage Status'] = rf.pop('coverage_status', 'Not Covered')
                    
                    # Remove any extra keys
                    keys_to_remove = [k for k in list(rf.keys()) if k not in [
                        'Source Link', 'Date', 'Extracted Red Flag', 'Associated Paragraph', 
                        'Category', 'Coverage by Existing Red Flag', 'Coverage Status'
                    ]]
                    for k in keys_to_remove:
                        rf.pop(k, None)
                
                return red_flags
            
            except json.JSONDecodeError as e:
                print(f"  ⚠ JSON parse error (attempt {attempt + 1}): {e}")
                if attempt == retry_count - 1:
                    return []
                time.sleep(2)
            
            except Exception as e:
                print(f"  ⚠ API error (attempt {attempt + 1}): {e}")
                if attempt == retry_count - 1:
                    return []
                time.sleep(5)
        
        return []
    
    except Exception as e:
        print(f"  ✗ Error processing {doc_path.name}: {e}")
        return []

def process_all_documents(test_mode=False, test_limit=5, year_month_filter=None):
    """Process all 823 documents across 9 sources
    
    Args:
        test_mode: If True, only process test_limit files for testing
        test_limit: Number of files to process in test mode
        year_month_filter: Filter documents by year-month (e.g., '2025-09')
    """
    
    print("\n" + "="*70)
    print("🔍 COMPREHENSIVE RED FLAG EXTRACTION SYSTEM")
    if test_mode:
        print(f"   🧪 TEST MODE: Processing only {test_limit} files")
    if year_month_filter:
        print(f"   📅 FILTER: {year_month_filter}")
    print("   Powered by Gemini 2.5 Pro")
    print("="*70)
    
    # Step 1: Categorize keywords (or load if already done)
    print("\n📋 Step 1: Categorizing 166 keywords...")
    
    categorized_file = 'data/categorized_keywords.json'
    if os.path.exists(categorized_file):
        print("✓ Loading existing categorization...")
        with open(categorized_file, 'r', encoding='utf-8') as f:
            categorized = json.load(f)
        print(f"✓ Loaded {len(categorized['AML_KEYWORDS'])} AML keywords and {len(categorized['TRANSACTION_PATTERNS'])} transaction patterns")
    else:
        categorized = categorize_keywords()
        if not categorized:
            print("✗ Failed to categorize keywords. Exiting.")
            return
    
    aml_keywords = categorized['AML_KEYWORDS']
    transaction_patterns = categorized['TRANSACTION_PATTERNS']
    
    # Step 2: Load existing red flags
    print("\n📚 Step 2: Loading existing red flags database...")
    df_existing = pd.read_csv('data/existing_red_flags.csv')
    existing_red_flags = df_existing.to_dict('records')
    print(f"✓ Loaded {len(existing_red_flags)} existing red flags")
    
    # Step 3: Find all documents
    print("\n🔎 Step 3: Scanning for documents...")
    scraped_dir = Path('scraped_content')
    
    # Find all txt files (html_text and pdf_text)
    documents = []
    for org_dir in scraped_dir.iterdir():
        if org_dir.is_dir():
            # Find html_text and pdf_text files
            html_files = list(org_dir.rglob('html_text/*.txt'))
            pdf_files = list(org_dir.rglob('pdf_text/*.txt'))
            documents.extend(html_files + pdf_files)
    
    print(f"✓ Found {len(documents)} documents to process")
    
    # Filter by year-month if specified
    if year_month_filter:
        filtered_docs = []
        for doc in documents:
            # Extract date from path (e.g., scraped_content/amf/2025-09/...)
            doc_date = doc.parts[-3] if len(doc.parts) >= 3 else None
            if doc_date == year_month_filter:
                filtered_docs.append(doc)
        
        print(f"📅 Filtered to {len(filtered_docs)} documents for {year_month_filter}")
        documents = filtered_docs
        
        if len(documents) == 0:
            print(f"⚠ No documents found for {year_month_filter}")
            return
    
    # Test mode: limit to first N files
    if test_mode:
        documents = documents[:test_limit]
        print(f"🧪 Test mode: Processing only {len(documents)} files")
    
    # Show breakdown by organization
    org_counts = {}
    for doc in documents:
        org = doc.parts[-4] if len(doc.parts) >= 4 else 'unknown'
        org_counts[org] = org_counts.get(org, 0) + 1
    
    print("\n📊 Documents by organization:")
    for org, count in sorted(org_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"   {org.upper():12s}: {count:4d} documents")
    
    # Step 4: Process documents
    print(f"\n⚡ Step 4: Processing {len(documents)} documents with Gemini 2.5 Pro...")
    print("   (This will take some time - processing in batches with rate limiting)")
    
    all_red_flags = []
    processed_count = 0
    error_count = 0
    skipped_count = 0
    
    batch_size = 50  # Process in larger batches for efficiency
    total_batches = (len(documents) + batch_size - 1) // batch_size
    
    # Check for existing progress
    progress_file = 'output/red_flags_progress.csv'
    processed_files = set()
    if os.path.exists(progress_file):
        try:
            df_existing = pd.read_csv(progress_file)
            if 'source_file' in df_existing.columns:
                all_red_flags = df_existing.to_dict('records')
                processed_files = set(df_existing['source_file'].unique())
                print(f"\n✓ Resuming from previous run: {len(processed_files)} files already processed")
        except:
            pass
    
    # Create progress bar for all documents
    with tqdm(total=len(documents), desc="Processing Documents", unit="doc", 
              bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]') as pbar:
        
        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min(start_idx + batch_size, len(documents))
            batch_docs = documents[start_idx:end_idx]
            
            pbar.set_description(f"📦 Batch {batch_num + 1}/{total_batches}")
            
            batch_red_flags = 0
            for doc in batch_docs:
                doc_relative = str(doc.relative_to('scraped_content'))
                
                # Skip if already processed
                if doc_relative in processed_files:
                    pbar.update(1)
                    pbar.set_postfix({"flags": len(all_red_flags), "status": "skipped"})
                    continue
                
                try:
                    red_flags = extract_red_flags_from_document(
                        doc, existing_red_flags, aml_keywords, transaction_patterns
                    )
                    
                    if red_flags:
                        all_red_flags.extend(red_flags)
                        processed_count += 1
                        batch_red_flags += len(red_flags)
                        pbar.set_postfix({
                            "flags": len(all_red_flags), 
                            "current": f"{len(red_flags)} flags",
                            "status": "✓"
                        })
                    else:
                        skipped_count += 1
                        pbar.set_postfix({"flags": len(all_red_flags), "status": "no flags"})
                    
                    pbar.update(1)
                    
                    # Rate limiting (shorter for Flash model)
                    time.sleep(0.5)
                
                except KeyboardInterrupt:
                    print("\n\n⚠ Interrupted by user. Saving progress...")
                    raise
                except Exception as e:
                    error_count += 1
                    pbar.set_postfix({"flags": len(all_red_flags), "status": "✗ error"})
                    pbar.update(1)
            
            # Save progress after each batch (Excel only)
            if all_red_flags:
                df_progress = pd.DataFrame(all_red_flags)
                # Ensure columns are in the correct order (only required columns)
                column_order = [
                    'Source Link', 'Date', 'Extracted Red Flag', 'Associated Paragraph', 
                    'Category', 'Coverage by Existing Red Flag', 'Coverage Status'
                ]
                # Only keep columns that exist
                existing_cols = [col for col in column_order if col in df_progress.columns]
                df_progress = df_progress[existing_cols]
                # Save as Excel only
                progress_excel = 'output/red_flags_progress.xlsx'
                df_progress.to_excel(progress_excel, index=False, engine='openpyxl')
                tqdm.write(f"      💾 Batch {batch_num + 1} complete: {batch_red_flags} flags extracted, {len(all_red_flags)} total saved")

    
    # Step 5: Save results
    print("\n💾 Step 5: Saving results...")
    
    if all_red_flags:
        df_results = pd.DataFrame(all_red_flags)
        
        # Ensure columns are in the correct order (only required columns)
        column_order = [
            'Source Link', 'Date', 'Extracted Red Flag', 'Associated Paragraph', 
            'Category', 'Coverage by Existing Red Flag', 'Coverage Status'
        ]
        # Only keep columns that exist
        existing_cols = [col for col in column_order if col in df_results.columns]
        df_results = df_results[existing_cols]
        
        # Save comprehensive results in Excel format ONLY
        output_file = 'output/extracted_red_flags_comprehensive.xlsx'
        df_results.to_excel(output_file, index=False, engine='openpyxl')
        print(f"✓ Saved {len(all_red_flags)} red flags to {output_file}")
        
        # Calculate summary statistics for display only
        summary = {
            'total_documents': len(documents),
            'processed_successfully': processed_count,
            'skipped_no_flags': skipped_count,
            'errors': error_count,
            'total_red_flags_extracted': len(all_red_flags),
            'not_covered': len(df_results[df_results['Coverage Status'] == 'Not Covered']),
            'fully_covered': len(df_results[df_results['Coverage Status'] == 'Fully Covered']),
            'partially_covered': len(df_results[df_results['Coverage Status'] == 'Partially Covered']),
            'organizations_processed': list(org_counts.keys()),
            'red_flags_by_category': df_results['Category'].value_counts().to_dict(),
        }
        
        # Print summary
        print("\n" + "="*70)
        print("📊 EXTRACTION SUMMARY")
        print("="*70)
        print(f"Total Documents Scanned:        {len(documents)}")
        print(f"Successfully Processed:         {processed_count}")
        print(f"Skipped (no flags):             {skipped_count}")
        print(f"Errors:                         {error_count}")
        print(f"\n✨ TOTAL RED FLAGS EXTRACTED:    {len(all_red_flags)}")
        print(f"   - Fully Covered:             {summary['fully_covered']}")
        print(f"   - Partially Covered:         {summary['partially_covered']}")
        print(f"   - Not Covered:               {summary['not_covered']}")
        
        print(f"\n📂 Top Categories:")
        for cat, count in sorted(summary['red_flags_by_category'].items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"   {cat:30s}: {count:4d} flags")
        
        print(f"\n🏢 By Organization:")
        org_by_date = df_results['Date'].value_counts().to_dict()
        for org, count in sorted(org_by_date.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"   {org:12s}: {count:4d} flags")
        
        print("\n" + "="*70)
        print("✅ EXTRACTION COMPLETE!")
        print("="*70)
    
    else:
        print("✗ No red flags extracted from any documents")

if __name__ == "__main__":
    import sys
    import argparse
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Extract red flags from regulatory documents')
    parser.add_argument('--test', '-t', action='store_true', help='Test mode (process only 5 files)')
    parser.add_argument('--year', type=str, help='Filter by year (e.g., 2025)')
    parser.add_argument('--month', type=str, help='Filter by month (e.g., 09 or 9)')
    
    args = parser.parse_args()
    
    # Format year-month filter
    year_month_filter = None
    if args.year and args.month:
        # Normalize month to 2-digit format
        month = args.month.zfill(2) if len(args.month) == 1 else args.month
        year_month_filter = f"{args.year}-{month}"
        print(f"\n📅 Filtering documents for: {year_month_filter}")
    
    if args.test:
        print("\n🧪 Running in TEST MODE (5 files only)")
        print("   To process all files, run without --test flag\n")
        process_all_documents(test_mode=True, test_limit=5, year_month_filter=year_month_filter)
    else:
        process_all_documents(test_mode=False, year_month_filter=year_month_filter)
