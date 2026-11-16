# price-watch-scraper

Scrape a paginated product listing into **CSV + Excel** and get a list of **price changes since the last run**. Runs by itself every day on GitHub Actions (free) or cron.

```
INFO page 1: https://books.toscrape.com/catalogue/page-1.html
INFO   20 items
INFO page 2: https://books.toscrape.com/catalogue/page-2.html
INFO   20 items
INFO saved 40 products -> data/latest.csv
INFO 2 price change(s) since last run:
INFO   ▼ £53.74 -> £49.99  Tipping the Velvet
INFO   ▲ £22.60 -> £24.10  Sharp Objects
```

## What it does

- Crawls every page by following the "next" link, with retries, timeouts and a polite delay
- Extracts title, price, currency, rating, stock status and URL
- Writes `data/latest.csv` and a formatted `data/latest.xlsx` (bold header, number formats, filters, frozen header row)
- Compares with the previous CSV and prints every price that moved
- All CSS selectors live in `config.json`: pointing it at another store is a config change, not a code change
- Handles `£51.77`, `$1,234.56` and `R$ 1.234,56` price formats

The default target is [books.toscrape.com](https://books.toscrape.com), a sandbox site built for scraping practice.

## Run it

```bash
pip install -r requirements.txt
python scraper.py                       # 2 pages -> data/latest.csv + data/latest.xlsx
python scraper.py --pages 50            # the whole catalogue
python scraper.py --min-price 30 --no-xlsx
python scraper.py --config shop.json    # another site
```

Sample output: [`examples/sample.csv`](examples/sample.csv) · [`examples/sample.xlsx`](examples/sample.xlsx)

## Run it every day for free

The workflow in `.github/workflows/daily.yml` runs the scraper at 09:00 UTC, commits the new snapshot to `data/` and keeps the diff in the Actions log. Enable it in the repo's **Actions** tab. For a server, the cron equivalent is:

```
0 9 * * * cd /opt/price-watch-scraper && /usr/bin/python3 scraper.py >> scrape.log 2>&1
```

## Adapt it to another site

Edit `config.json`:

```json
{
  "start_url": "https://example-shop.com/products?page=1",
  "selectors": {
    "item":   "div.product-card",
    "title":  "a.product-name",
    "price":  "span.price",
    "rating": "div.stars",
    "stock":  "span.stock",
    "next":   "a.pagination-next"
  }
}
```

Sites that render products with JavaScript need a browser-based fetcher (Playwright); the parsing and output code stays the same.

## Tests

```bash
pytest -q
```

Tests run against a saved HTML fixture, so they pass offline.

## Stack

Python 3.10+, requests, BeautifulSoup (lxml), openpyxl, pytest.

## License

MIT
