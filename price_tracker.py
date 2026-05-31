# ================================================
# Amazon Price Tracker
# Author: fafay | Fiverr.com/fafay
# Description: Tracks Amazon product prices and
#              saves history to an Excel file.
# ================================================

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import time
import random

EXCEL_FILE = "price_history.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive",
}


def get_amazon_price(url: str) -> dict:
    """Scrape product name and price from an Amazon product page."""
    try:
        time.sleep(random.uniform(1, 3))  # polite delay
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        # Product title
        title_tag = soup.find(id="productTitle")
        title = title_tag.get_text(strip=True) if title_tag else "Unknown Product"

        # Price (tries multiple selectors for different page layouts)
        price = None
        selectors = [
            {"class": "a-price-whole"},
            {"id": "priceblock_ourprice"},
            {"id": "priceblock_dealprice"},
            {"class": "a-offscreen"},
        ]
        for sel in selectors:
            tag = soup.find(attrs=sel)
            if tag:
                raw = tag.get_text(strip=True).replace(",", "").replace("$", "").replace("€", "").replace("£", "")
                try:
                    price = float(raw.split(".")[0] + "." + raw.split(".")[1][:2]) if "." in raw else float(raw)
                    break
                except Exception:
                    continue

        return {"title": title[:60], "price": price, "url": url, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")}

    except Exception as e:
        print(f"  [!] Error scraping {url}: {e}")
        return {"title": "Error", "price": None, "url": url, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")}


def style_header_row(ws, row: int, num_cols: int):
    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center


def style_data_row(ws, row: int, num_cols: int, alternate: bool):
    fill_color = "D6E4F0" if alternate else "FFFFFF"
    row_fill = PatternFill("solid", start_color=fill_color)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = row_fill
        cell.alignment = center
        cell.border = border
        cell.font = Font(name="Arial", size=10)


def setup_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Price History"

    headers = ["#", "Product Name", "Price (€/$)", "Change", "Date & Time", "URL"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    ws.row_dimensions[1].height = 22

    col_widths = [5, 45, 14, 12, 20, 50]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    return wb


def load_or_create_workbook():
    if os.path.exists(EXCEL_FILE):
        return load_workbook(EXCEL_FILE)
    return setup_workbook()


def get_last_price(ws, url: str):
    """Find the most recent price logged for a given URL."""
    last_price = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[5] == url and row[2] is not None:
            last_price = row[2]
    return last_price


def append_price_record(ws, counter: int, data: dict, last_price):
    price = data["price"]
    change = ""
    if price is not None and last_price is not None:
        diff = round(price - last_price, 2)
        if diff > 0:
            change = f"▲ +{diff}"
        elif diff < 0:
            change = f"▼ {diff}"
        else:
            change = "→ No change"

    row_data = [counter, data["title"], price if price else "N/A", change, data["timestamp"], data["url"]]
    ws.append(row_data)

    row_num = ws.max_row
    alternate = (row_num % 2 == 0)
    style_data_row(ws, row_num, len(row_data), alternate)

    # Color price cell based on change
    price_cell = ws.cell(row=row_num, column=3)
    if "▲" in change:
        price_cell.font = Font(name="Arial", size=10, color="C00000", bold=True)
    elif "▼" in change:
        price_cell.font = Font(name="Arial", size=10, color="375623", bold=True)


def track_prices(urls: list):
    print("\n🔍 Amazon Price Tracker — fafay | Fiverr\n" + "=" * 45)
    wb = load_or_create_workbook()
    ws = wb.active

    counter = ws.max_row  # continue numbering from last run

    for url in urls:
        print(f"\n📦 Checking: {url[:60]}...")
        data = get_amazon_price(url)
        last_price = get_last_price(ws, url)
        append_price_record(ws, counter, data, last_price)
        counter += 1

        status = f"  ✅ {data['title'][:50]} — Price: {data['price']}"
        print(status)

    wb.save(EXCEL_FILE)
    print(f"\n✅ Done! Results saved to '{EXCEL_FILE}'\n")


# ── EDIT YOUR PRODUCT URLs BELOW ──────────────────
PRODUCT_URLS = [
    "https://www.amazon.com/dp/B08N5WRWNW",  # Echo Dot
    "https://www.amazon.com/dp/B07XJ8C8F5",  # Fire TV Stick
    # Add more Amazon product URLs here
]

if __name__ == "__main__":
    track_prices(PRODUCT_URLS)
