# mdash_to_mint

A one-off script to sync transaction data from [MoneyDashboard](https://www.moneydashboard.com/) (a UK budgeting tool) into [Mint](https://mint.intuit.com/) (a US budgeting tool), written when household finances were being tracked in both countries simultaneously.

## What it does

1. Reads a CSV export from MoneyDashboard containing transactions with UK-style categories
2. Translates those categories to their Mint equivalents via a hardcoded mapping (e.g. "Broadband" → "Internet", "Council tax" → "Local Tax")
3. Cleans up transaction descriptions and converts amounts from a pre-converted USD column
4. Uses Selenium to log into Mint and enter each transaction through the web UI

## Usage

```bash
python mint.py -s transactions.csv -u <mint_email> -p <mint_password>
```

### Arguments

| Flag | Description |
|------|-------------|
| `-s` / `--spreadsheet_path` | Path to MoneyDashboard CSV export (default: `transactions.xlsx`) |
| `-u` / `--username` | Mint login email |
| `-p` / `--password` | Mint password |
| `-f` / `--failonmissingcategory` | Halt if any category can't be mapped |
| `--load_sheet_only` | Parse and validate the spreadsheet without uploading |
| `--debug` | Enable debug logging |

## Input format

The CSV must contain these columns: `Date`, `OriginalDescription`, `L1Tag`, `L2Tag`, `L3Tag`, `USD`.

The `USD` column should contain amounts already converted to US dollars. Transactions tagged `DNU` (Do Not Upload) — such as credit card repayments and inter-account transfers — are silently skipped.

## Notes

- This is abandonware. Mint shut down in 2024, and MoneyDashboard also ceased operations.
- The Selenium-based upload is fragile and was tied to the Mint DOM at the time of writing.
- There is a `pdb.set_trace()` left in `read_transactions_from_csv` — a relic of debugging.
- Excel input (`read_transactions_from_excel`) was never implemented.
