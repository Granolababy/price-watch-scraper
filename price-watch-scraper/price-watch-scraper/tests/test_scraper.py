import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import scraper  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = (Path(__file__).parent / "fixtures" / "page1.html").read_text(encoding="utf-8")
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
BASE = "https://books.toscrape.com/catalogue/page-1.html"


def test_parse_price_formats():
    assert scraper.parse_price("£51.77") == (51.77, "£")
    assert scraper.parse_price("$1,234.56") == (1234.56, "$")
    assert scraper.parse_price("R$ 1.234,56") == (1234.56, "R$")
    assert scraper.parse_price("19,90 €")[0] == 19.90


def test_parse_page_extracts_all_fields():
    items = scraper.parse_page(FIXTURE, BASE, CFG)
    assert len(items) == 3
    first = items[0]
    assert first.title == "A Light in the Attic"          # full title from the title attribute, not the truncated text
    assert first.price == 51.77 and first.currency == "£"
    assert first.rating == 3
    assert first.in_stock is True
    assert first.url == "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"
    assert items[2].in_stock is False
    assert items[2].rating == 5


def test_next_page_url_is_resolved():
    assert scraper.next_page_url(FIXTURE, BASE, CFG) == "https://books.toscrape.com/catalogue/page-2.html"


def test_csv_roundtrip_and_price_diff(tmp_path):
    items = scraper.parse_page(FIXTURE, BASE, CFG)
    out = tmp_path / "latest.csv"
    scraper.write_csv(items, out)

    with out.open(newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    assert [r["title"] for r in rows] == [i.title for i in items]

    previous = scraper.read_previous(out)
    items[1].price = 49.99                       # simulate a price drop on the second product
    changes = scraper.diff_prices(previous, items)
    assert len(changes) == 1
    product, old, new = changes[0]
    assert product.title == "Tipping the Velvet" and old == 53.74 and new == 49.99


def test_xlsx_is_written(tmp_path):
    from openpyxl import load_workbook

    items = scraper.parse_page(FIXTURE, BASE, CFG)
    out = tmp_path / "latest.xlsx"
    scraper.write_xlsx(items, out)
    ws = load_workbook(out)["Products"]
    assert ws.max_row == 4
    assert ws["A1"].value == "Title" and ws["B2"].value == 51.77
