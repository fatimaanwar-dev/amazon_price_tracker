# 🛒 Amazon Price Tracker

A Python script that automatically tracks Amazon product prices and saves the full history to a beautifully formatted Excel file — with price change indicators (▲ up / ▼ down).

---

## ✨ Features

- ✅ Scrapes product name and price from any Amazon product URL
- ✅ Saves history to a formatted `.xlsx` Excel file
- ✅ Shows price change on every run (▲ increase / ▼ decrease / → no change)
- ✅ Color-coded rows and price change highlights
- ✅ Polite random delay to avoid being blocked
- ✅ Supports multiple products in one run
- ✅ Appends to existing history — never overwrites old data

---

## 📸 Sample Output

| # | Product Name | Price | Change | Date & Time |
|---|---|---|---|---|
| 1 | Amazon Echo Dot (5th Gen) | $49.99 | — | 2026-05-28 09:00 |
| 2 | Fire TV Stick 4K Max | $54.99 | — | 2026-05-28 09:01 |
| 3 | Amazon Echo Dot (5th Gen) | $44.99 | ▼ -5.0 | 2026-05-29 09:00 |
| 4 | Fire TV Stick 4K Max | $59.99 | ▲ +5.0 | 2026-05-29 09:01 |

---

## 🚀 How to Run

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/amazon-price-tracker.git
cd amazon-price-tracker
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Add your product URLs
Open `price_tracker.py` and edit the `PRODUCT_URLS` list:
```python
PRODUCT_URLS = [
    "https://www.amazon.com/dp/B08N5WRWNW",
    "https://www.amazon.com/dp/YOUR_PRODUCT_ID",
]
```

### 4. Run the script
```bash
python price_tracker.py
```

Results are saved to `price_history.xlsx` in the same folder.

---

## 🔁 Automate It (Optional)

Run it daily automatically using:

**Windows (Task Scheduler):**
- Open Task Scheduler → Create Basic Task → set trigger to Daily → action: run `python price_tracker.py`

**Mac/Linux (cron):**
```bash
# Run every day at 9am
0 9 * * * cd /path/to/project && python price_tracker.py
```

---

## 📦 Requirements

- Python 3.8+
- requests
- beautifulsoup4
- openpyxl

---

## ⚠️ Disclaimer

This tool is for personal/educational use. Amazon's website structure may change over time, which can affect scraping. Use responsibly and in accordance with Amazon's Terms of Service.

---

## 👨‍💻 About

Built by **fafay** — Python automation freelancer on Fiverr.
Need a custom script? [Hire me on Fiverr →](https://www.fiverr.com/fafay)
