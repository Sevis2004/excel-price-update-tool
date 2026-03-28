# Excel Price Update Tool

A simple Python tool that updates product prices in an Excel file using product page URLs from a website.

## What the tool does

The script:

- reads product URLs from an Excel file
- opens each product page
- extracts the current price from the HTML
- writes the result back into Excel:
  - price
  - update timestamp
  - processing status

## Input Excel structure

The input file must contain these columns:

- Product URL
- Price
- Updated At
- Status

Example:

| Product URL | Price | Updated At | Status |
|-------------|-------|------------|--------|
| http://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html |  |  |  |

## Processing rules

Each row is processed independently.

Possible statuses:

- success
- price not found
- request failed
- empty url

Errors in one row do not stop processing of the remaining rows.

## Tech stack

- Python
- requests
- BeautifulSoup
- openpyxl

## Files

- `main.py` — main script
- `input.xlsx` — input/output Excel file
- `requirements.txt` — Python dependencies
- `run_price_update.bat` — quick start for Windows

## How to run

### Option 1 — batch file
Double-click:

`run_price_update.bat`

### Option 2 — terminal

Run:

```bash
python main.py
```

## Example result

After running the script, the Excel file will be updated with:

- current price
- timestamp
- processing status

Example:

| Product URL | Price | Updated At | Status |
|-------------|-------|------------|--------|
| http://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html | £51.77 | 2026-03-28 16:57:39 | success |
