# Playwright QA Automation Demo 🎭✅

**Excel-driven Playwright test automation with video evidence and narrated reporting.**

By **Muhammad Bilal**

Five end-to-end scenarios run against public demo sites, driven entirely by
`test_data.xlsx` — no hard-coded values. Every run produces per-case video
evidence, screenshots, and a formatted Excel report with PASS/FAIL highlighting.

## ✨ What it runs (5 scenarios, 5/5 PASS)

| # | Scenario | Site |
|---|----------|------|
| 1 | Standard user buys two items | saucedemo.com |
| 2 | Standard user buys one item | saucedemo.com |
| 3 | Locked-out user cannot log in (negative case) | saucedemo.com |
| 4 | Catalog KPI report generation — extracts products, avg price | saucedemo.com |
| 5 | Sprint board task automation — create + complete tasks | todomvc.com |

Scenarios 4 and 5 are PM-flavored: KPI reporting and sprint-board workflows —
the kind of flows a QA engineer automates for a delivery team.

## 🎬 Video-friendly execution

- Red outline highlighting on every element being acted on
- Slowed pacing (~700 ms between steps) so runs are easy to follow on video
- One `.webm` video + screenshots per case in `output/`

## 📊 Reporting

`output/results.xlsx` is written every run with:

- Per-case PASS/FAIL with green/red cell formatting
- `Catalog_KPI_Report` sheet (product list, average price)
- `Sprint_Summary` sheet (tasks created vs completed)

## 🛠 Tech stack

- **Python** 3.10+, **Playwright** (Chromium), **openpyxl**

## 🚀 Run it

```bash
pip install -r requirements.txt
python -m playwright install chromium
python demo.py
```

On restricted networks, start the bundled forward proxy first (it adds the
egress-proxy auth for Chromium):

```bash
python forward_proxy.py   # in the background, then run demo.py
```

## 📄 License

MIT — demo sites used are public test fixtures (SauceDemo, TodoMVC).

## 🏢 Office Automation Toolkit

This repo also includes [`office-automation/`](office-automation/) — three small,
generic Python utilities for everyday office paperwork:

- **mom-generator** — builds a formatted Meeting Minutes `.docx` from JSON
- **excel-tracker** — builds a formatted Excel task tracker with status dropdowns,
  RAG conditional formatting, and live summary formulas
- **csv-consolidator** — merges a folder of CSVs into one workbook with a summary sheet

```bash
pip install python-docx openpyxl
```
