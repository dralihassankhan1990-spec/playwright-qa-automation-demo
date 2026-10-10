# Office Automation Toolkit

Small, generic Python utilities that automate everyday office paperwork —
no frameworks, no servers, just `python script.py input`.

| Tool | Does | Needs |
|---|---|---|
| `mom-generator/` | Builds a formatted **Meeting Minutes (.docx)** from a JSON file — title block, metadata table, discussion/decisions, action-items table | `python-docx` |
| `excel-tracker/` | Builds a formatted **Excel task tracker** from JSON — styled headers, status dropdowns, RAG conditional formatting, live summary formulas | `openpyxl` |
| `csv-consolidator/` | Merges a folder of **CSVs into one workbook** — one sheet per file + summary sheet | `openpyxl` |

```bash
pip install python-docx openpyxl
cd mom-generator && python mom_generator.py sample_meeting.json
cd ../excel-tracker && python excel_tracker.py sample_tasks.json
cd ../csv-consolidator && python csv_consolidator.py samples
```

Each folder has its own README and sample input files. Outputs are written
next to the script unless `-o <file>` is given.
