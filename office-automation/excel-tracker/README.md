# Excel Task Tracker Builder

Builds a formatted Excel work tracker (`.xlsx`) from a simple JSON task list —
no manual spreadsheet formatting needed.

## Run

```bash
pip install openpyxl
python excel_tracker.py sample_tasks.json
python excel_tracker.py tasks.json -o tracker_oct.xlsx
```

## Input JSON

`title` plus `tasks[]` — each task: `id`, `task`, `owner`, `status`
(`Not Started` / `In Progress` / `Blocked` / `Done`), `priority`
(`Low` / `Medium` / `High` / `Critical`), `due` (YYYY-MM-DD).

## Output

- **Tasks** sheet: styled headers, frozen panes, auto-filters, a Status
  dropdown on every row, and RAG conditional formatting on Status/Priority.
- **Summary** sheet: live `COUNTIF` formulas — total tasks, per-status
  counts, and overdue items (not Done and past due date).
