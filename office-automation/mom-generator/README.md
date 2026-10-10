# Meeting Minutes (MOM) Generator

Generates a professional, print-ready **Minutes of Meeting** Word document (`.docx`)
from a simple JSON file — no manual formatting needed.

## Run

```bash
pip install python-docx
python mom_generator.py sample_meeting.json
python mom_generator.py my_meeting.json -o minutes_oct.docx
```

## Input JSON

`title`, `date`, `time`, `venue`, `attendees[]`, plus:

- `agenda[]` — each item: `topic`, `discussion[]` bullet points, `decisions[]`
- `actions[]` — each item: `task`, `owner`, `due`

See `sample_meeting.json` for a complete example.

## Output

A formatted `.docx` with a centered title block, a metadata table, numbered
discussion sections with decisions highlighted, and an action-items table
(task / owner / due date) — ready to circulate or print.
