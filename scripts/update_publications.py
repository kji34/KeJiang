import html
import json
import re
import urllib.request
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path.cwd()
OUTPUT = ROOT / "publications-data.json"
SCHOLAR_URL = "https://scholar.google.com/citations?hl=en&user=2YQxaJgAAAAJ&view_op=list_works&sortby=pubdate"

DEFAULT_IMPACT_FACTORS = {
    "Engineering Structures": "5.3",
    "Journal of Constructional Steel Research": "4.3",
    "Thin-Walled Structures": "5.0",
    "Structures": "4.4",
    "Earthquake Engineering & Structural Dynamics": "4.0",
    "Journal of Structural Engineering": "3.8",
    "Composite Structures": "7.8",
    "Construction and Building Materials": "7.4",
    "Automation in Construction": "10.3",
    "Advances in Structural Engineering": "2.2",
    "International Journal of Steel Structures": "2.4",
    "Engineering Failure Analysis": "4.5",
    "Materials & Design": "7.6",
    "Journal of Building Engineering": "6.0",
    "Bulletin of the New Zealand Society for Earthquake Engineering": "n/a",
}

KNOWN_JOURNALS = [
    "Engineering Structures",
    "Journal of Constructional Steel Research",
    "Thin-Walled Structures",
    "Structures",
    "Earthquake Engineering & Structural Dynamics",
    "Journal of Structural Engineering",
    "Composite Structures",
    "Construction and Building Materials",
    "Automation in Construction",
    "Advances in Structural Engineering",
    "International Journal of Steel Structures",
    "Engineering Failure Analysis",
    "Materials & Design",
    "Journal of Building Engineering",
    "Bulletin of the New Zealand Society for Earthquake Engineering",
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


def normalize_journal(value: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    if not text:
        return ""
    lowered = text.lower()
    for known in sorted(KNOWN_JOURNALS, key=len, reverse=True):
        if lowered.startswith(known.lower()) or known.lower() in lowered:
            return known
    text = re.sub(r"\s*,\s*\d{4}$", "", text)
    text = re.sub(r"\s+\d+(?:\([^)]+\))?(?:,\s*\d+(?:-\d+)?)?$", "", text)
    text = text.strip(" ,")
    if text.endswith("."):
        text = text[:-1]
    return text


def build_impact_display(journal: str, impact_factors: dict) -> str:
    if not journal:
        return ""
    if journal in impact_factors and impact_factors[journal] not in (None, "", "n/a"):
        return f"(JCR Q1, IF={impact_factors[journal]})"
    return "(JCR Q1, IF=n/a)"


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


def build_payload(items, existing_impact_factors):
    by_year = {}
    for item in items:
        year = item["year"] or "Unknown"
        by_year.setdefault(year, []).append(item)

    sorted_years = sorted(by_year, key=lambda y: (y == "Unknown", -int(y)) if isinstance(y, int) else (1, y))
    grouped = []
    for year in sorted_years:
        grouped.append({"year": year, "items": by_year[year]})

    counts = Counter(item["journal"] for item in items if item["journal"])
    stats = []
    for journal, count in counts.most_common():
        stats.append({
            "journal": journal,
            "count": count,
            "impactDisplay": build_impact_display(journal, existing_impact_factors),
        })

    return {
        "generatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sourceUrl": SCHOLAR_URL,
        "impactFactors": existing_impact_factors,
        "stats": stats,
        "years": grouped,
        "items": items,
    }


def main():
    existing_data = {}
    if OUTPUT.exists():
        try:
            existing_data = json.loads(OUTPUT.read_text(encoding="utf-8"))
        except Exception:
            existing_data = {}
    existing_impact_factors = existing_data.get("impactFactors", {}) or {}
    for journal, factor in DEFAULT_IMPACT_FACTORS.items():
        existing_impact_factors.setdefault(journal, factor)

    html_text = fetch_html(f"{SCHOLAR_URL}&cstart=0&pagesize=100")
    items = parse_items(html_text)
    payload = build_payload(items, existing_impact_factors)
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {OUTPUT}")
    print(f"Parsed {len(items)} items")


if __name__ == "__main__":
    main()
