# Assign Shift

A minimal web-based shift management tool with Google Sheets synchronization.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Create a Google service account and share a spreadsheet with it. Set these environment variables:
   - `GOOGLE_SERVICE_ACCOUNT_FILE`: path to the service account JSON file
   - `SPREADSHEET_ID`: ID of the spreadsheet to use

## Running

```bash
python app.py
```

The app reads from the spreadsheet on each request and writes back on form submissions. Editing the spreadsheet directly is reflected on reload.

### Import/Export

- `/export_excel` downloads the current data as an Excel file with three sheets.
- `/import_csv` accepts a CSV upload with columns `Person,Project,Month,Fraction` and appends rows to the Assignments sheet.
