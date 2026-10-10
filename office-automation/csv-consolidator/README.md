# CSV Consolidator

Merges every CSV file in a folder into a single Excel workbook (`.xlsx`) —
one worksheet per CSV, plus a **Summary** sheet with per-file row/column
counts and a total.

## Run

```bash
pip install openpyxl
python csv_consolidator.py ./samples
python csv_consolidator.py ./my_csvs -o consolidated.xlsx
```

Each CSV's first row is treated as the header. Worksheets get styled
headers, frozen panes, and auto-filters; sheet names are sanitized for
Excel's naming rules.
