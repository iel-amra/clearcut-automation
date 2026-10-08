# Demo: full-catalogue scraper (books.toscrape.com)

Scrapes all 1,000 books from [books.toscrape.com](https://books.toscrape.com/), a public sandbox built for scraping practice, and delivers the data as **CSV and Excel**.

What it shows
- Pagination handling (follows "next" until the last page)
- Per-item detail pages (category read from each product page)
- Retry with backoff, polite delay, de-duplication
- Clean typed columns: `title, price_gbp, rating, availability, in_stock, category, product_url`
- Excel output with frozen header, filters and column widths

## Run it
```bash
pip install requests beautifulsoup4 lxml pandas openpyxl
python scrape_books.py                  # all 1,000 books (~2 min)
python scrape_books.py --max-pages 2    # quick sample
```
Outputs `books.csv` and `books.xlsx` in the current folder. Sample output: [`books.csv`](books.csv).

Built by Clearcut Automation. Want the same for your target site? See the [portfolio page](../site/index.html).
