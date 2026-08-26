import html
import json
import re
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

# Scrapes Google Scholar into publications-data-raw.json (nothing else).
# Normally started through: python scripts/publications_excel.py refresh
ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "publications-data-raw.json"
SCHOLAR_URL = "https://scholar.google.com/citations?hl=en&user=2YQxaJgAAAAJ&view_op=list_works&sortby=pubdate"

# Canonical journal spellings. Only used to fix Google Scholar's casing slips
# (e.g. "Engineering structures" -> "Engineering Structures"); matching is exact,
# so "Earthquakes and Structures" is never folded into "Structures".
KNOWN_JOURNALS = [
    "Engineering Structures",
    "Journal of Constructional Steel Research",
    "Thin-Walled Structures",
    "Structures",
    "Earthquakes and Structures",
    "Earthquake Engineering & Structural Dynamics",
    "Earthquake Engineering and Resilience",
    "Earthquake Engineering and Engineering Vibration",
    "Journal of Structural Engineering",
    "Journal of Building Engineering",
    "Composite Structures",
    "Construction and Building Materials",
    "Automation in Construction",
    "Advances in Structural Engineering",
    "International Journal of Steel Structures",
    "Engineering Failure Analysis",
    "Materials & Design",
    "Bulletin of the New Zealand Society for Earthquake Engineering",
    "PLOS ONE",
]


def fetch_html(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as response:
        return response.read().decode("utf-8", "ignore")


def strip_tags(value: str) -> str:
    value = re.sub(r"<[^>]+>", " ", value)
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def split_venue(value: str):
    """Split a Scholar venue string into (journal, volume, issue, pages).

    "Thin-Walled Structures 212, 113190 , 2025" -> ("Thin-Walled Structures", "212", "", "113190")
    "Earthquakes and Structures 26 (5), 383 , 2024" -> ("Earthquakes and Structures", "26", "5", "383")
    """
    text = re.sub(r"\s+", " ", value or "").strip()
    text = re.sub(r"\s*,\s*(?:19|20)\d{2}\s*$", "", text).strip()
    if not text:
        return "", "", "", ""

    pages = ""
    match = re.match(r"^(.*),\s*([^,]+)$", text)
    if match and re.fullmatch(r"[\dA-Za-z]+(?:\s*[-\u2013]\s*[\dA-Za-z]+)?", match.group(2).strip()):
        text = match.group(1).strip()
        pages = match.group(2).strip()

    volume = ""
    issue = ""
    match = re.match(r"^(.*?)\s+(\d+)\s*(?:\(([^)]+)\))?$", text)
    if match:
        text = match.group(1).strip()
        volume = match.group(2)
        issue = (match.group(3) or "").strip()

    return text.strip(" ,."), volume, issue, pages


def normalize_journal(value: str) -> str:
    """Journal / venue name only, using the canonical spelling when it is known."""
    name = split_venue(value)[0]
    if not name:
        return ""
    for known in KNOWN_JOURNALS:
        if name.lower() == known.lower():
            return known
    return name


def parse_items(html_text: str):
    rows = re.findall(r'<tr class="gsc_a_tr"[^>]*>(.*?)</tr>', html_text, re.S)
    items = []
    for row in rows:
        title_match = re.search(r'<a[^>]*class="gsc_a_at"[^>]*>(.*?)</a>', row, re.S)
        gray_divs = re.findall(r'<div class="gs_gray">(.*?)</div>', row, re.S)
        year_match = re.search(r'<span[^>]*class="gsc_a_h[^"]*"[^>]*>(\d{4})</span>', row)
        citation_match = re.search(r'<a href="([^"]+)" class="gsc_a_at"', row)
        if not title_match:
            continue
        title = strip_tags(title_match.group(1))
        authors = strip_tags(gray_divs[0]) if gray_divs else ""
        venue = strip_tags(gray_divs[1]) if len(gray_divs) > 1 else ""
        year = int(year_match.group(1)) if year_match else None
        citation_url = f"https://scholar.google.com{citation_match.group(1)}" if citation_match else ""
        journal = normalize_journal(venue)
        items.append({
            "title": title,
            "authors": authors,
            "venue": venue,
            "journal": journal,
            "year": year,
            "citationUrl": citation_url,
        })
    return items


def build_payload(items):
    return {
        "generatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sourceUrl": SCHOLAR_URL,
        "totalCount": len(items),
        "items": items,
    }


def main():
    html_text = fetch_html(f"{SCHOLAR_URL}&cstart=0&pagesize=100")
    items = parse_items(html_text)

    # Safety net: Google Scholar sometimes answers with a captcha or a partial
    # page. Never replace a good snapshot with a suspiciously small one.
    previous = 0
    if OUTPUT.exists():
        try:
            previous = len(json.loads(OUTPUT.read_text(encoding="utf-8")).get("items") or [])
        except Exception:
            previous = 0
    if not items or (previous and len(items) < previous * 0.8):
        raise SystemExit(
            f"! Scholar returned only {len(items)} records (previous snapshot: {previous}).\n"
            f"  {OUTPUT.name} was left untouched - try again later."
        )

    OUTPUT.write_text(json.dumps(build_payload(items), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {OUTPUT.name}: {len(items)} records (previous snapshot: {previous})")


if __name__ == "__main__":
    main()
