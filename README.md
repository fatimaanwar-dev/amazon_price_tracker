# 🛒 Amazon Price Tracker

A Python automation tool that tracks Amazon product prices over time and logs history to a formatted Excel report — built to demonstrate web scraping, browser automation, and data pipeline skills.

---

## 🧠 Technical Highlights

- **Browser automation with Playwright** — uses a real Chromium instance instead of plain HTTP requests, bypassing Amazon's bot detection that blocks `requests`/`urllib` based scrapers
- **ASIN-based URL normalization** — extracts the core product ID from any URL format (standard, short `amzn.to`, `/gp/product/`, or long tracking URLs) so price history stays consistent regardless of how the URL was pasted
- **Incremental Excel writes** — appends new records to existing history without overwriting, using `openpyxl` for full formatting control (color-coded rows, change indicators, frozen headers)
- **Graceful degradation** — detects CAPTCHA responses and rate limits explicitly rather than silently writing bad data

---

## 📸 Sample Output

| # | Product Name | Price | Change | Date & Time |
|---|---|---|---|---|
| 1 | Amazon Echo Dot (5th Gen) | $49.99 | — | 2026-05-28 09:00 |
| 2 | Fire TV Stick 4K Max | $54.99 | — | 2026-05-28 09:01 |
| 3 | Amazon Echo Dot (5th Gen) | $44.99 | ▼ -5.0 | 2026-05-29 09:00 |
| 4 | Fire TV Stick 4K Max | $59.99 | ▲ +5.0 | 2026-05-29 09:01 |

---

## 🏗️ Project Structure

```
amazon-price-tracker/
│
├── price_tracker.py      # Core scraper and Excel writer
├── requirements.txt      # Dependencies
├── sample_output.xlsx    # Example output file
└── README.md
```

---

## ⚙️ How It Works

1. URLs are normalized to extract the ASIN — ensuring consistent matching across runs regardless of URL format
2. Playwright launches a headless Chromium browser with a real user-agent and viewport, mimicking a genuine browser session
3. A homepage visit runs first to acquire cookies before hitting the product page
4. BeautifulSoup parses the DOM across multiple CSS selectors to handle Amazon's varying page layouts
5. The previous price for each product is looked up by ASIN in the existing Excel history
6. New rows are appended with change indicators (▲ / ▼ / →) and the workbook is saved

---

## 🚀 Getting Started

```bash
git clone https://github.com/fatimaanwar-dev/amazon_price_tracker.git
cd amazon-price-tracker
pip install -r requirements.txt
python -m playwright install chromium
```

Add your product URLs to `PRODUCT_URLS` in `price_tracker.py`:

```python
PRODUCT_URLS = [
    "https://www.amazon.com/dp/B08N5WRWNW",
    "https://amzn.to/3xYzABC",           # short links supported
    "https://www.amazon.com/Some-Long-Title/dp/B07XJ8C8F5?ref=sr_1_1",  # messy URLs supported
]
```

Then run:

```bash
python price_tracker.py
```

---

## 📦 Dependencies

| Library | Purpose |
|---|---|
| `playwright` | Headless browser automation |
| `beautifulsoup4` | HTML parsing |
| `openpyxl` | Excel file creation and formatting |

---

## ⚠️ Limitations & Known Tradeoffs

- Amazon occasionally serves CAPTCHAs to headless browsers — the script detects and reports this explicitly rather than writing bad data
- Amazon's page layout changes periodically; CSS selectors may need updating if `N/A` prices appear consistently
- For production-grade reliability, a rotating proxy service (e.g. ScraperAPI) or the official Amazon Product Advertising API would be more appropriate

---

## 👩‍💻 Author

**Fatima Anwar** — Python automation & scripting
[Fiverr](https://www.fiverr.com/blogsbyfatima) · [GitHub](https://github.com/fatimaanwar-dev/fatimaanwar-dev.git)
