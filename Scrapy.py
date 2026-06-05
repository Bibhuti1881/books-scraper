"""
=============================================================
  E-Commerce Web Scraper — books.toscrape.com
  Author  : Bibhuti Adhikari
  Output  : books_data.xlsx  (all 1000 books, 50 pages)
  Libraries: requests, beautifulsoup4, openpyxl, tqdm
=============================================================

Install dependencies:
    pip install requests beautifulsoup4 openpyxl tqdm

Run:
    python books_scraper.py
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
import time
import re
import os
from tqdm import tqdm  # progress bar


# ── Config ────────────────────────────────────────────────────────────────────
BASE_URL    = "https://books.toscrape.com/catalogue/"
START_URL   = "https://books.toscrape.com/catalogue/page-1.html"
OUTPUT_FILE = "books_data.xlsx"
DELAY       = 0.5   # seconds between requests (be polite)

# Star rating words → numbers
RATING_MAP = {
    "One": 1, "Two": 2, "Three": 3,
    "Four": 4, "Five": 5
}


# ── Helper: fetch page safely ─────────────────────────────────────────────────
def fetch(url, retries=3):
    headers = {"User-Agent": "Mozilla/5.0 (compatible; BookScraper/1.0)"}
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            resp.raise_for_status()
            return resp
        except requests.RequestException as e:
            print(f"  Retry {attempt+1}/{retries} for {url} — {e}")
            time.sleep(2)
    return None


# ── Scrape single book detail page ───────────────────────────────────────────
def scrape_book_detail(url):
    """
    Fetches the individual book page for extra info:
    description, UPC, number of reviews.
    Returns a dict or empty dict on failure.
    """
    resp = fetch(url)
    if not resp:
        return {}

    soup = BeautifulSoup(resp.text, "html.parser")
    detail = {}

    # Description
    desc_tag = soup.select_one("#product_description ~ p")
    detail["description"] = desc_tag.get_text(strip=True) if desc_tag else "N/A"

    # Product table (UPC, price, availability, reviews)
    table = soup.select("table.table tr")
    for row in table:
        cells = row.find_all("td")
        if len(cells) == 2:
            key   = row.find("th").get_text(strip=True)
            value = cells[0].get_text(strip=True)
            if key == "UPC":
                detail["upc"] = value
            elif key == "Number of reviews":
                detail["reviews"] = value

    return detail


# ── Scrape all 50 pages ───────────────────────────────────────────────────────
def scrape_all_books():
    all_books = []
    url = START_URL
    page_num = 1

    print("\n📚 Starting scrape of books.toscrape.com ...")
    print("─" * 55)

    while url:
        print(f"\n📄 Page {page_num}/50 — {url}")
        resp = fetch(url)
        if not resp:
            print(f"  ❌ Failed to fetch page {page_num}, stopping.")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        book_articles = soup.select("article.product_pod")

        for article in tqdm(book_articles, desc=f"  Books on page {page_num}", leave=False):
            book = {}

            # Title
            title_tag = article.select_one("h3 a")
            book["title"] = title_tag["title"] if title_tag else "N/A"

            # Price
            price_tag = article.select_one("p.price_color")
            price_text = price_tag.get_text(strip=True) if price_tag else "0"
            # Remove currency symbol, convert to float
            book["price_gbp"] = float(re.sub(r"[^\d.]", "", price_text))

            # Star rating
            rating_tag = article.select_one("p.star-rating")
            if rating_tag:
                rating_word = rating_tag["class"][1]  # e.g. "Three"
                book["rating"] = RATING_MAP.get(rating_word, 0)
            else:
                book["rating"] = 0

            # Availability
            avail_tag = article.select_one("p.availability")
            book["availability"] = avail_tag.get_text(strip=True) if avail_tag else "N/A"

            # Book detail page URL
            relative_url = title_tag["href"].replace("../", "") if title_tag else ""
            detail_url   = BASE_URL + relative_url
            book["url"]  = detail_url

            # Fetch extra details from book page
            extra = scrape_book_detail(detail_url)
            book["upc"]         = extra.get("upc", "N/A")
            book["description"] = extra.get("description", "N/A")
            book["reviews"]     = extra.get("reviews", "0")

            all_books.append(book)
            time.sleep(DELAY)

        # ── Next page ──────────────────────────────────────────────────────
        next_btn = soup.select_one("li.next a")
        if next_btn:
            next_href = next_btn["href"]
            # Fix URL for first page vs other pages
            if page_num == 1:
                url = BASE_URL + next_href
            else:
                url = BASE_URL + next_href
            page_num += 1
        else:
            url = None   # No more pages

    print(f"\n✅ Scraped {len(all_books)} books total!")
    return all_books


# ── Export to Excel ───────────────────────────────────────────────────────────
def export_to_excel(books):
    print("\n📊 Creating Excel file ...")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All Books"

    # ── Colors & fonts ────────────────────────────────────────────────────────
    header_fill  = PatternFill("solid", fgColor="1A1A2E")   # dark navy
    accent_fill  = PatternFill("solid", fgColor="FF4500")   # orange
    alt_fill     = PatternFill("solid", fgColor="F7F7F7")   # light gray
    header_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    title_font   = Font(name="Calibri", bold=True, color="1A1A2E", size=11)
    data_font    = Font(name="Calibri", size=10)
    center       = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left         = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin         = Side(style="thin", color="DDDDDD")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "📚  BOOKS.TOSCRAPE.COM — Full Catalogue Scrape"
    title_cell.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    title_cell.fill  = accent_fill
    title_cell.alignment = center
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Total books scraped: {len(books)}   |   Source: books.toscrape.com   |   By: Bibhuti Adhikari"
    sub_cell.font  = Font(name="Calibri", italic=True, color="FFFFFF", size=10)
    sub_cell.fill  = header_fill
    sub_cell.alignment = center
    ws.row_dimensions[2].height = 20

    # ── Column headers ────────────────────────────────────────────────────────
    headers = ["#", "Title", "Price (GBP)", "Rating (★)", "Availability", "UPC", "Reviews", "URL"]
    col_widths = [5, 45, 13, 12, 14, 18, 10, 60]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center
        cell.border    = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[3].height = 22

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_idx, book in enumerate(books, start=1):
        excel_row = row_idx + 3   # offset for title + header rows
        fill = alt_fill if row_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")

        row_data = [
            row_idx,
            book["title"],
            book["price_gbp"],
            "★" * book["rating"] + "☆" * (5 - book["rating"]),
            book["availability"],
            book["upc"],
            int(book["reviews"]) if book["reviews"].isdigit() else 0,
            book["url"],
        ]

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.font      = title_font if col_idx == 2 else data_font
            cell.fill      = fill
            cell.border    = border
            cell.alignment = center if col_idx != 2 else left

        ws.row_dimensions[excel_row].height = 18

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")

    ws2.merge_cells("A1:C1")
    s_title = ws2["A1"]
    s_title.value = "📊  SCRAPE SUMMARY"
    s_title.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    s_title.fill  = accent_fill
    s_title.alignment = center
    ws2.row_dimensions[1].height = 28

    prices  = [b["price_gbp"] for b in books]
    ratings = [b["rating"] for b in books]

    summary_data = [
        ("Total Books Scraped",     len(books)),
        ("Total Pages Scraped",     50),
        ("Average Price (GBP)",     f"£{sum(prices)/len(prices):.2f}"),
        ("Cheapest Book (GBP)",     f"£{min(prices):.2f}"),
        ("Most Expensive (GBP)",    f"£{max(prices):.2f}"),
        ("Average Star Rating",     f"{sum(ratings)/len(ratings):.1f} ★"),
        ("5-Star Books",            sum(1 for r in ratings if r == 5)),
        ("1-Star Books",            sum(1 for r in ratings if r == 1)),
        ("In Stock",                sum(1 for b in books if "In stock" in b["availability"])),
        ("Out of Stock",            sum(1 for b in books if "In stock" not in b["availability"])),
        ("Source",                  "books.toscrape.com"),
        ("Scraped By",              "Bibhuti Adhikari"),
    ]

    for r_idx, (label, value) in enumerate(summary_data, start=2):
        l_cell = ws2.cell(row=r_idx, column=1, value=label)
        v_cell = ws2.cell(row=r_idx, column=2, value=value)
        l_cell.font      = Font(name="Calibri", bold=True, size=11)
        v_cell.font      = Font(name="Calibri", size=11)
        row_fill = PatternFill("solid", fgColor="F0F0F0") if r_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        l_cell.fill      = row_fill
        v_cell.fill      = PatternFill("solid", fgColor="F0F0F0") if r_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        l_cell.alignment = left
        v_cell.alignment = center
        l_cell.border    = border
        v_cell.border    = border
        ws2.row_dimensions[r_idx].height = 20

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 22

    # ── Freeze panes & save ───────────────────────────────────────────────────
    ws.freeze_panes  = "A4"   # freeze title + header
    ws2.freeze_panes = "A2"

    wb.save(OUTPUT_FILE)
    print(f"✅ Excel saved → {OUTPUT_FILE}")
    print(f"   • Sheet 1: All {len(books)} books with full details")
    print(f"   • Sheet 2: Summary statistics")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    books = scrape_all_books()
    if books:
        export_to_excel(books)
        print("\n🎉 Done! Open books_data.xlsx to see your data.")
    else:
        print("\n❌ No books scraped. Check your internet connection.")