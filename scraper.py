#!/usr/bin/env python3
"""
price-watch-scraper
-------------------
Scrape a paginated product listing into CSV / Excel and report price changes
since the last run.

Selectors live in config.json, so pointing this at another site is a config
change, not a code change.

Examples:
    python scraper.py                          # scrape with config.json, write data/latest.csv + .xlsx
    python scraper.py --pages 3 --out data/books
    python scraper.py --config shop.json --min-price 20
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup, UnicodeDammit
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

log = logging.getLogger("scraper")

DEFAULT_CONFIG = Path(__file__).with_name("config.json")
WORD_RATINGS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5}


@dataclass
class Product:
    title: str
    price: float
    currency: str
    rating: int | None
    in_stock: bool
    url: str


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #
def make_session(retries: int = 3, backoff: float = 0.5) -> requests.Session:
    """Session with retries on transient errors and a polite User-Agent."""
    session = requests.Session()
    retry = Retry(
        total=retries,
        backoff_factor=backoff,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
    )
    session.mount("http://", HTTPAdapter(max_retries=retry))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers["User-Agent"] = (
        "price-watch-scraper/1.0 (+https://github.com/; polite bot, 1 req/s)"
    )
    return session


def fetch(session: requests.Session, url: str, timeout: int = 20) -> str:
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    # Decode the raw bytes ourselves instead of trusting resp.text. When a server sends
    # no charset, requests assumes Latin-1 and "£" comes out as "Â£". So: try UTF-8
    # first (strict), then the encoding from the HTTP header, then the page's <meta> tag.
    candidates = ["utf-8"] + ([resp.encoding] if resp.encoding else [])
    return UnicodeDammit(resp.content, candidates, is_html=True).unicode_markup


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def parse_price(text: str) -> tuple[float, str]:
    """'£51.77' -> (51.77, '£'); 'R$ 1.234,56' -> (1234.56, 'R$')."""
    text = text.strip()
    currency = re.match(r"^[^\d\s]+", text)
    cur = currency.group(0) if currency else ""
    number = re.sub(r"[^\d,.\-]", "", text)
    # Brazilian style 1.234,56 -> 1234.56
    if "," in number and (number.rfind(",") > number.rfind(".")):
        number = number.replace(".", "").replace(",", ".")
    else:
        number = number.replace(",", "")
    return float(number), cur


def parse_rating(class_list: Iterable[str] | None) -> int | None:
    for cls in class_list or []:
        if cls.lower() in WORD_RATINGS:
            return WORD_RATINGS[cls.lower()]
    return None


def parse_page(html: str, base_url: str, cfg: dict) -> list[Product]:
    soup = BeautifulSoup(html, "lxml")
    sel = cfg["selectors"]
    products: list[Product] = []
    for card in soup.select(sel["item"]):
        link = card.select_one(sel["title"])
        price_el = card.select_one(sel["price"])
        if not link or not price_el:
            continue
        title = (link.get("title") or link.get_text()).strip()
        price, currency = parse_price(price_el.get_text())
        rating_el = card.select_one(sel["rating"]) if sel.get("rating") else None
        stock_el = card.select_one(sel["stock"]) if sel.get("stock") else None
        in_stock = bool(stock_el and cfg.get("in_stock_text", "in stock") in stock_el.get_text().lower())
        products.append(
            Product(
                title=title,
                price=price,
                currency=currency,
                rating=parse_rating(rating_el.get("class") if rating_el else None),
                in_stock=in_stock,
                url=urljoin(base_url, link.get("href", "")),
            )
        )
    return products


def next_page_url(html: str, current_url: str, cfg: dict) -> str | None:
    soup = BeautifulSoup(html, "lxml")
    nxt = soup.select_one(cfg["selectors"].get("next", ""))
    if nxt and nxt.get("href"):
        return urljoin(current_url, nxt["href"])
    return None


# --------------------------------------------------------------------------- #
# Crawl
# --------------------------------------------------------------------------- #
def crawl(cfg: dict, max_pages: int, delay: float, session: requests.Session | None = None) -> list[Product]:
    session = session or make_session()
    url: str | None = cfg["start_url"]
    results: list[Product] = []
    for page in range(1, max_pages + 1):
        if not url:
            break
        log.info("page %d: %s", page, url)
        html = fetch(session, url)
        items = parse_page(html, url, cfg)
        log.info("  %d items", len(items))
        results.extend(items)
        url = next_page_url(html, url, cfg)
        if url:
            time.sleep(delay)
    return results


# --------------------------------------------------------------------------- #
# Output
# --------------------------------------------------------------------------- #
FIELDS = ["title", "price", "currency", "rating", "in_stock", "url"]


def write_csv(products: list[Product], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for p in products:
            w.writerow(asdict(p))


def write_xlsx(products: list[Product], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append([f.replace("_", " ").title() for f in FIELDS])
    for p in products:
        ws.append([p.title, p.price, p.currency, p.rating, "Yes" if p.in_stock else "No", p.url])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "#,##0.00"
    for i, field in enumerate(FIELDS, start=1):
        width = max(len(field), *(len(str(getattr(p, field))) for p in products)) if products else len(field)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def read_previous(path: Path) -> dict[str, float]:
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8") as f:
        return {row["url"]: float(row["price"]) for row in csv.DictReader(f)}


def diff_prices(previous: dict[str, float], current: list[Product]) -> list[tuple[Product, float, float]]:
    """Return (product, old_price, new_price) for every item whose price moved."""
    changes = []
    for p in current:
        old = previous.get(p.url)
        if old is not None and abs(old - p.price) > 1e-9:
            changes.append((p, old, p.price))
    return changes


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="selectors + start URL (JSON)")
    ap.add_argument("--pages", type=int, default=2, help="max pages to crawl")
    ap.add_argument("--delay", type=float, default=1.0, help="seconds between page requests")
    ap.add_argument("--out", type=Path, default=Path("data/latest"), help="output path without extension")
    ap.add_argument("--min-price", type=float, default=None, help="drop items cheaper than this")
    ap.add_argument("--no-xlsx", action="store_true", help="write CSV only")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s %(message)s")
    cfg = json.loads(args.config.read_text(encoding="utf-8"))

    products = crawl(cfg, args.pages, args.delay)
    if args.min_price is not None:
        products = [p for p in products if p.price >= args.min_price]

    csv_path = args.out.with_suffix(".csv")
    previous = read_previous(csv_path)
    changes = diff_prices(previous, products)

    write_csv(products, csv_path)
    if not args.no_xlsx:
        write_xlsx(products, args.out.with_suffix(".xlsx"))

    log.info("saved %d products -> %s", len(products), csv_path)
    if previous:
        if changes:
            log.info("%d price change(s) since last run:", len(changes))
            for p, old, new in changes:
                arrow = "▲" if new > old else "▼"
                log.info("  %s %s%.2f -> %s%.2f  %s", arrow, p.currency, old, p.currency, new, p.title)
        else:
            log.info("no price changes since last run")
    return 0


if __name__ == "__main__":
    sys.exit(main())
