#!/usr/bin/env python3
"""Scrape the full catalogue of books.toscrape.com (a public scraping sandbox)
into CSV and Excel.

Usage:
    python scrape_books.py                 # all pages -> books.csv + books.xlsx
    python scrape_books.py --max-pages 3   # quick run

Each row: title, price_gbp, rating (1-5), availability, in_stock, category, product_url.
"""
import argparse
import csv
import sys
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

BASE = "https://books.toscrape.com/"
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; ClearcutAutomation-demo/1.0)"}


def fetch(session: requests.Session, url: str, retries: int = 3) -> BeautifulSoup:
    """GET a page with simple retry/backoff and return parsed HTML."""
    for attempt in range(retries):
        try:
            resp = session.get(url, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            return BeautifulSoup(resp.text, "lxml")
        except requests.RequestException as exc:
            if attempt == retries - 1:
                raise
            wait = 2 ** attempt
            print(f"  retry {attempt + 1} for {url} after error: {exc} (waiting {wait}s)", file=sys.stderr)
            time.sleep(wait)


def parse_listing(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extract every book card from a catalogue page."""
    rows = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        rating_cls = [c for c in card.select_one("p.star-rating")["class"] if c in RATING_WORDS]
        price_text = card.select_one("p.price_color").get_text(strip=True)
        availability = card.select_one("p.instock.availability").get_text(strip=True)
        rows.append({
            "title": link["title"],
            "price_gbp": float(price_text.replace("£", "").replace("Â", "")),
            "rating": RATING_WORDS[rating_cls[0]] if rating_cls else None,
            "availability": availability,
            "in_stock": availability.lower().startswith("in stock"),
            "product_url": urljoin(page_url, link["href"]),
        })
    return rows


def parse_category(session: requests.Session, product_url: str) -> str:
    """Open a product page and read its category from the breadcrumb."""
    soup = fetch(session, product_url)
    crumbs = [li.get_text(strip=True) for li in soup.select("ul.breadcrumb li")]
    return crumbs[2] if len(crumbs) >= 3 else ""


def scrape(max_pages: int | None, with_category: bool, delay: float) -> list[dict]:
    session = requests.Session()
    url = urljoin(BASE, "catalogue/page-1.html")
    page = 0
    books: list[dict] = []
    while url and (max_pages is None or page < max_pages):
        page += 1
        print(f"page {page}: {url}")
        soup = fetch(session, url)
        rows = parse_listing(soup, url)
        if with_category:
            for row in rows:
                row["category"] = parse_category(session, row["product_url"])
                time.sleep(delay)
        books.extend(rows)
        nxt = soup.select_one("li.next a")
        url = urljoin(url, nxt["href"]) if nxt else None
        time.sleep(delay)
    # de-duplicate on product URL, keep first occurrence
    seen, unique = set(), []
    for b in books:
        if b["product_url"] not in seen:
            seen.add(b["product_url"])
            unique.append(b)
    return unique


def write_outputs(books: list[dict], stem: str) -> None:
    fields = ["title", "price_gbp", "rating", "availability", "in_stock", "category", "product_url"]
    with open(f"{stem}.csv", "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(books)
    try:
        import pandas as pd
        df = pd.DataFrame(books, columns=fields)
        with pd.ExcelWriter(f"{stem}.xlsx", engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="books")
            ws = xw.sheets["books"]
            for col, width in zip("ABCDEFG", (60, 10, 8, 16, 9, 20, 70)):
                ws.column_dimensions[col].width = width
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    except ImportError:
        print("pandas/openpyxl not installed: skipping Excel output", file=sys.stderr)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--max-pages", type=int, default=None, help="stop after N catalogue pages")
    ap.add_argument("--no-category", action="store_true", help="skip product pages (faster, no category column)")
    ap.add_argument("--delay", type=float, default=0.2, help="seconds between requests (be polite)")
    ap.add_argument("--out", default="books", help="output file stem")
    args = ap.parse_args()
    books = scrape(args.max_pages, not args.no_category, args.delay)
    write_outputs(books, args.out)
    print(f"done: {len(books)} books -> {args.out}.csv / {args.out}.xlsx")


if __name__ == "__main__":
    main()
