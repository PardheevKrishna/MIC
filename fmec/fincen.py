#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinCEN News harvester (month/year filter, text extractor)
"""

import csv, os, re, html, time, random, pathlib, datetime as dt
from typing import List, Optional, Dict, Set, Tuple
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from pdfminer.high_level import extract_text as extract_pdf_text

BASE  = "https://www.fincen.gov"
INDEX = BASE + "/news"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "FinancialComplianceBot/1.0 (research@example.com)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
})
# Disable SSL verification warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

REQUEST_TIMEOUT   = 10  # Reduced from 20 for faster timeouts
PAUSE_RANGE       = (0.05, 0.1)  # Reduced from (0.1, 0.3) for faster crawling
RETRY_BACKOFF     = 0.5  # Reduced from 0.8
MAX_FILENAME_LEN  = 150

def snooze(): time.sleep(random.uniform(*PAUSE_RANGE))

def safe_filename(s: str, maxlen: int = MAX_FILENAME_LEN) -> str:
    s = html.unescape(s or "").strip()
    s = re.sub(r"[^\w\-.() ]+", "_", s)
    s = re.sub(r"[ _]{2,}", " ", s)
    if len(s) > maxlen:
        stem, ext = os.path.splitext(s)
        s = stem[: maxlen-len(ext)-1] + "_" + ext
    return s or "file"

def ensure_unique(path: pathlib.Path) -> pathlib.Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists(): return path
    stem, ext = path.stem, path.suffix
    i = 2
    while True:
        cand = path.with_name(f"{stem} ({i}){ext}")
        if not cand.exists(): return cand
        i += 1

def get(url: str, retries: int = 3) -> Optional[requests.Response]:
    for k in range(retries):
        try:
            hdrs = {"Referer": INDEX}
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT, headers=hdrs, allow_redirects=True, verify=False)
            if 200 <= r.status_code < 300: return r
            if r.status_code in (403, 429) or 500 <= r.status_code < 600:
                time.sleep(RETRY_BACKOFF * (2**k) + random.uniform(0, 0.3))
                continue
            return None
        except (requests.RequestException, Exception) as e:
            if k < retries - 1:  # Don't sleep on last retry
                time.sleep(RETRY_BACKOFF * (2**k) + random.uniform(0, 0.3))
            else:
                return None  # Give up on last retry
    return None

def soup_get(url: str) -> Optional[BeautifulSoup]:
    r = get(url)
    if not r: return None
    return BeautifulSoup(r.text, "lxml")

def parse_date(text: str) -> Optional[dt.date]:
    """Parse dates like '10/23/2025' or '10/14/2025'"""
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", text or "")
    if not m: return None
    month = int(m.group(1))
    day = int(m.group(2))
    year = int(m.group(3))
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None

def is_external_link(url: str, base_domain: str = "fincen.gov") -> bool:
    """Check if URL is external to FinCEN (like federalregister.gov)"""
    parsed = urlparse(url)
    return parsed.netloc and base_domain not in parsed.netloc

def extract_links_from_soup(soup: BeautifulSoup, base_url: str) -> Tuple[List[str], List[str]]:
    """
    Extract all HTML page links and PDF links from soup.
    Returns: ([html_urls], [pdf_urls])
    """
    html_links = []
    pdf_links = []
    
    # Domains to skip (known problematic sites)
    skip_domains = ["bsaefiling.fincen.treas.gov"]
    
    for link in soup.find_all("a", href=True):
        href = link.get("href", "").strip()
        if not href or href.startswith("#") or href.startswith("javascript:"):
            continue
            
        abs_url = urljoin(base_url, href)
        
        # Skip problematic domains
        if any(domain in abs_url for domain in skip_domains):
            continue
        
        # Check if it's a PDF
        if ".pdf" in abs_url.lower():
            if abs_url not in pdf_links:
                pdf_links.append(abs_url)
        else:
            # It's an HTML page - check if it's a meaningful link
            link_text = link.get_text(strip=True).lower()
            # Skip navigation, footer, and social media links
            if any(skip in abs_url.lower() for skip in ["twitter.com", "facebook.com", "linkedin.com", 
                                                          "youtube.com", "#", "mailto:", "tel:"]):
                continue
            # Include external regulatory sites
            if any(domain in abs_url.lower() for domain in ["federalregister.gov", "treasury.gov", 
                                                              "fincen.gov/system/files"]):
                if abs_url not in html_links:
                    html_links.append(abs_url)
            # Include links with keywords suggesting important content
            elif any(keyword in link_text for keyword in ["full notice", "full report", "full document", 
                                                            "read more", "view document", "federal register",
                                                            "order", "form", "survey", "guidance", "rule"]):
                if abs_url not in html_links:
                    html_links.append(abs_url)
    
    return html_links, pdf_links

def extract_text_from_soup(soup: BeautifulSoup) -> Optional[str]:
    """Extract main text content from a BeautifulSoup object"""
    # Remove navigation and sidebar elements
    for unwanted in soup.select("nav, .navigation, .sidebar, .menu, header, footer, .breadcrumb, .social-share"):
        unwanted.decompose()
    
    # Try multiple selectors for the main content
    content = None
    
    # Common content selectors
    for selector in [
        "article .field--name-body",
        ".node--type-press-release",
        ".node--type-news",
        "article .content",
        "article",
        ".main-content",
        "main",
        ".body-content",
        "#content-area",
    ]:
        content = soup.select_one(selector)
        if content: break
    
    if not content:
        # Fallback: try to find the main content div
        content = soup.find("div", class_=lambda x: x and ("content" in x.lower() or "body" in x.lower()))
    
    if not content:
        # Last resort: get all paragraph text from main content area
        main = soup.find("main") or soup.find("div", {"role": "main"})
        if main:
            paragraphs = main.find_all("p")
        else:
            paragraphs = soup.find_all("p")
        
        if paragraphs:
            return "\n\n".join(p.get_text(strip=True) for p in paragraphs if p.get_text(strip=True))
        return None
    
    # Clean up the text
    text = content.get_text("\n", strip=True)
    # Remove excessive newlines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text

def crawl_recursive(start_url: str, max_depth: int = 2, visited: Optional[Set[str]] = None, pdf_dir=None, pdf_text_dir=None) -> Dict[str, str]:
    """
    Recursively crawl starting from start_url, following all links and PDFs.
    Returns dict mapping URL -> extracted text content.
    max_depth: how many levels deep to crawl (default 2 - main + sub-pages + their PDFs)
    pdf_dir: if provided, saves actual PDF files here
    pdf_text_dir: if provided, saves extracted PDF text here
    """
    if visited is None:
        visited = set()
    
    results = {}
    
    def _crawl(url: str, depth: int):
        if depth > max_depth or url in visited:
            return
        
        visited.add(url)
        
        # Only print main article
        if depth == 0:
            print(f"  Crawling main article...")
        
        # Check if it's a PDF
        if ".pdf" in url.lower():
            # Download and extract PDF text
            try:
                r = get(url)
                if r and r.content:
                    # Generate filename from URL
                    pdf_filename = url.split("/")[-1]
                    if not pdf_filename.endswith(".pdf"):
                        pdf_filename = f"{pdf_filename}.pdf"
                    
                    # Save PDF file if directory provided
                    if pdf_dir:
                        pdf_path = pdf_dir / pdf_filename
                        # Make unique if exists
                        counter = 1
                        while pdf_path.exists():
                            name_part = pdf_filename.rsplit(".pdf", 1)[0]
                            pdf_path = pdf_dir / f"{name_part}_{counter}.pdf"
                            counter += 1
                        
                        with open(pdf_path, "wb") as f:
                            f.write(r.content)
                        print(f"    [PDF saved] {pdf_path.name}")
                    
                    # Save to temp file and extract
                    import tempfile
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(r.content)
                        tmp_path = tmp.name
                    
                    try:
                        pdf_text = extract_pdf_text(tmp_path)
                        if pdf_text and pdf_text.strip():
                            results[url] = pdf_text
                            
                            # Save extracted text if directory provided
                            if pdf_text_dir:
                                txt_filename = pdf_filename.rsplit(".pdf", 1)[0] + ".txt"
                                txt_path = pdf_text_dir / txt_filename
                                # Make unique if exists
                                counter = 1
                                while txt_path.exists():
                                    name_part = pdf_filename.rsplit(".pdf", 1)[0]
                                    txt_path = pdf_text_dir / f"{name_part}_{counter}.txt"
                                    counter += 1
                                
                                with open(txt_path, "w", encoding="utf-8") as f:
                                    f.write(pdf_text)
                                print(f"    [PDF text saved] {txt_path.name}")
                            
                            if depth == 1:  # Only print for direct PDFs
                                print(f"    [PDF] {len(pdf_text):,} chars")
                    except Exception:
                        pass  # Skip failed PDFs silently
                    finally:
                        os.unlink(tmp_path)
            except Exception:
                pass  # Skip failed downloads silently
            return
        
        # It's an HTML page
        soup = soup_get(url)
        if not soup:
            return
        
        # Extract text from this page
        text = extract_text_from_soup(soup)
        if text and text.strip():
            # Skip CAPTCHA/bot detection pages
            if "captcha" in text.lower()[:500] or "bot test" in text.lower()[:500] or "aggressive automated scraping" in text.lower()[:500]:
                return
            
            results[url] = text
            if depth == 0:
                print(f"    [HTML] {len(text):,} chars")
        
        # Only crawl sub-links from the main article (depth 0) or important external pages (depth 1)
        if depth == 0:
            # Find all links on this page
            html_links, pdf_links = extract_links_from_soup(soup, url)
            
            # Crawl ALL PDFs (no limit)
            for pdf_link in pdf_links:
                if pdf_link not in visited:
                    _crawl(pdf_link, depth + 1)
            
            # Also crawl important external pages that might have PDFs
            # (Federal Register, Treasury.gov documents)
            for html_link in html_links[:5]:  # Limit to first 5 HTML sub-pages
                if any(domain in html_link.lower() for domain in ["federalregister.gov/documents", "treasury.gov"]):
                    if html_link not in visited:
                        # Crawl these pages at depth 1 to get their PDFs
                        _crawl(html_link, depth + 1)
        
        elif depth == 1:
            # For depth 1 pages (Federal Register, etc), also extract their PDFs
            html_links, pdf_links = extract_links_from_soup(soup, url)
            
            # Crawl PDFs found on these pages
            for pdf_link in pdf_links[:5]:  # Limit to first 5 PDFs from sub-pages
                if pdf_link not in visited:
                    _crawl(pdf_link, depth + 1)
    
    _crawl(start_url, 0)
    return results

def extract_article_text(url: str, pdf_dir=None, pdf_text_dir=None) -> tuple[Optional[str], list[str]]:
    """
    Fetch the article detail page and recursively extract all linked content.
    Returns: (combined_text, [all_urls_visited])
    """
    # Recursively crawl the article and all linked PDFs (depth 2 for sub-pages)
    crawled_content = crawl_recursive(url, max_depth=2, pdf_dir=pdf_dir, pdf_text_dir=pdf_text_dir)
    
    if not crawled_content:
        return None, []
    
    # Combine all extracted text with source URLs
    combined_parts = []
    all_urls = list(crawled_content.keys())
    
    for crawled_url, content in crawled_content.items():
        combined_parts.append(f"\n{'='*80}\nSource: {crawled_url}\n{'='*80}\n")
        combined_parts.append(content)
    
    combined_text = "\n\n".join(combined_parts)
    
    # Simplified output
    print(f"  → {len(crawled_content)} sources, {len(combined_text):,} chars total")
    
    return combined_text, all_urls

def yield_news_articles(soup: BeautifulSoup):
    """
    Yield tuples (date_text, title, link, category) from the news listing.
    """
    # Find all news article containers - try multiple selectors
    articles = soup.find_all("div", class_="fincen-news-article")
    
    # If no articles found, try alternative selectors
    if not articles:
        articles = soup.find_all("article", class_=lambda x: x and "news" in str(x).lower())
    
    if not articles:
        # Try finding rows in views
        articles = soup.select(".view-content .views-row")
    
    seen = set()
    for article in articles:
        content = article.find("div", class_="fincen-news-article__content")
        if not content:
            # Try alternative structure
            content = article
        
        # Extract date - try multiple methods
        time_tag = content.find("time")
        date_text = ""
        if time_tag:
            date_text = time_tag.get_text(strip=True)
        else:
            # Look for date in metadata or spans
            date_span = content.find("span", class_=lambda x: x and "date" in str(x).lower())
            if date_span:
                date_text = date_span.get_text(strip=True)
        
        if not date_text:
            # Try to find any date pattern
            text = content.get_text()
            date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
            if date_match:
                date_text = date_match.group(1)
        
        if not date_text:
            continue
        
        # Extract category
        meta = content.find("div", class_="fincen-news-article__meta")
        category = ""
        if meta:
            category_text = meta.get_text("|", strip=True)
            parts = category_text.split("|")
            if len(parts) > 1:
                category = parts[1].strip()
        
        # Extract title and link
        title_div = content.find("div", class_="fincen-news-article__title")
        if not title_div:
            # Try h2, h3 or any heading
            title_div = content.find(["h2", "h3", "h4"])
        
        if not title_div:
            continue
        
        link_tag = title_div.find("a", href=True)
        if not link_tag:
            continue
        
        title = link_tag.get_text(strip=True)
        href = link_tag["href"]
        abs_url = urljoin(BASE, href)
        
        key = (abs_url, title)
        if key in seen:
            continue
        seen.add(key)
        
        yield date_text, title, abs_url, category

def crawl_month(year: int, month: int, out_root: str = "scraped_content") -> Dict:
    out_root = pathlib.Path(out_root)
    month_root = out_root / "fincen" / f"{year}-{month:02d}"
    text_dir = month_root / "html_text"
    pdf_dir = month_root / "pdf"
    pdf_text_dir = month_root / "pdf_text"
    text_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    pdf_text_dir.mkdir(parents=True, exist_ok=True)

    url = INDEX
    seen_pages = set()
    seen_articles = set()
    fetched = 0
    manifest: List[Dict[str, str]] = []
    found_previous_month = False
    page_count = 0
    max_pages = 50  # Safety limit

    print(f"[info] Starting crawl for {year}-{month:02d}")
    
    while url and url not in seen_pages and not found_previous_month and page_count < max_pages:
        seen_pages.add(url)
        page_count += 1
        print(f"\n[page {page_count}] {url}")
        soup = soup_get(url)
        if not soup:
            print("  - [warn] failed to fetch page; stopping.")
            break

        articles_found = 0
        for date_text, title, link_url, category in yield_news_articles(soup):
            articles_found += 1
            
            # Skip duplicates
            if link_url in seen_articles:
                continue
            seen_articles.add(link_url)
            
            d = parse_date(date_text)
            if not d:
                print(f"  - [skip] could not parse date: {date_text}")
                continue
            
            # Stop if we've reached the previous month
            if d.year < year or (d.year == year and d.month < month):
                print(f"  - [info] reached previous month ({d.isoformat()}), stopping pagination.")
                found_previous_month = True
                break
            
            # Skip if not the target month
            if not (d.year == year and d.month == month):
                print(f"  - [skip] not target month: {d.isoformat()}")
                continue

            # Fetch the full article text and recursively crawl all linked pages/PDFs
            combined_text, all_urls = extract_article_text(link_url, pdf_dir=pdf_dir, pdf_text_dir=pdf_text_dir)
            
            if not combined_text:
                print(f"    [skip] could not extract text")
                continue

            # Save combined text file with all crawled content
            safe_title = safe_filename(title)
            fname = f"{d.isoformat()} {safe_title}.txt"
            fpath = ensure_unique(text_dir / fname)

            with open(fpath, "w", encoding="utf-8") as f:
                f.write(f"Title: {title}\n")
                f.write(f"Date: {date_text}\n")
                f.write(f"Category: {category}\n")
                f.write(f"URL: {link_url}\n")
                f.write(f"Related Documents: {', '.join(all_urls)}\n")
                f.write(f"\n{'='*80}\n\n")
                f.write(combined_text)

            fetched += 1
            print(f"    ✓ Saved: {len(combined_text):,} chars, {len(all_urls)} URLs\n")
            
            manifest.append({
                "date": d.isoformat(),
                "title": title,
                "category": category,
                "url": link_url,
                "saved_file": str(fpath.resolve()),
                "chars": str(len(combined_text)),
                "related_docs": ", ".join(all_urls),
            })
            
            # Only snooze between main articles, not between sub-crawls
            time.sleep(0.1)  # Very brief pause
        
        if articles_found > 0:
            print(f"  → {articles_found} articles on this page, {fetched} total saved")
        
        if articles_found == 0:
            print(f"  - [warn] No articles found on this page, stopping pagination.")
            break

        # Pagination - look for "Next" or page numbers
        next_link = None
        
        # Try multiple pagination patterns
        if soup:
            # Pattern 1: nav.pager with next button
            pager = soup.find("nav", class_=lambda x: x and "pager" in str(x).lower())
            if pager:
                next_btn = pager.find("a", class_=lambda x: x and "next" in str(x).lower())
                if next_btn and next_btn.get("href"):
                    next_link = urljoin(url, next_btn["href"])
                else:
                    for a in pager.find_all("a", href=True):
                        if "next" in a.get_text().lower() or "›" in a.get_text():
                            next_link = urljoin(url, a["href"])
                            break
            
            # Pattern 2: Look for page=N parameter
            if not next_link:
                all_links = soup.find_all("a", href=True)
                for a in all_links:
                    if "page=" in a.get("href", ""):
                        href_text = a.get_text().strip().lower()
                        if "next" in href_text or "›" in href_text or "»" in href_text:
                            next_link = urljoin(url, a["href"])
                            break
        
        if next_link:
            print(f"  → Next page: {next_link}")
        else:
            print(f"  → No next page found, ending pagination")
        
        url = next_link

    manifest_path = month_root / "manifest.csv"
    with open(manifest_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Date", "Publication Title", "Publication Link", "Related Documents"])
        writer.writeheader()
        for row in sorted(manifest, key=lambda r: r["date"]):
            writer.writerow({
                'Date': row['date'],
                'Publication Title': row['title'],
                'Publication Link': row['url'],
                'Related Documents': row.get('related_docs', '')
            })

    print(f"\n{'='*80}")
    print(f"[done] Collected {fetched} news articles for {year}-{month:02d}")
    print(f"[done] Scanned {page_count} pages")
    print(f"[done] Output dir: {month_root.resolve()}")
    print(f"[done] Manifest: {manifest_path.resolve()}")
    print(f"{'='*80}")
    return {
        "count": fetched,
        "pages_scanned": page_count,
        "output_dir": str(month_root.resolve()),
        "manifest": str(manifest_path.resolve())
    }

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="FinCEN News harvester (month/year text downloader).")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--out", type=str, default="scraped_content")
    args = ap.parse_args()
    if not (1 <= args.month <= 12):
        raise SystemExit("Month must be 1..12")
    crawl_month(args.year, args.month, args.out)
