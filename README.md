# excel-data-doctor

An AI skill that gives your Excel / CSV files a data-quality checkup: automatically scans for 10+ classes of dirty-data issues, scores the file 0–100, cleans it with one command, and produces a health report. Works with WorkBuddy, OpenClaw, and Claude Code (standard Agent Skills format).

## Usage

**In WorkBuddy** (once the skill is installed, just say — no need to mention "checkup"):

> Give this spreadsheet a checkup (attach the file)

It runs four steps on its own: **diagnose** (scan and score) → **prescribe** (issues split into three buckets: auto-fix / needs your confirmation / report-only) → **clean** (original untouched, writes `*_cleaned.xlsx`) → **report** (Markdown, with fix details and pending items).

**Run the scripts directly** (no WorkBuddy needed):

```bash
pip install pandas openpyxl

# Diagnose: JSON report with health score
python3 scripts/diagnose.py sales.xlsx --out report.json

# Clean: safe fixes only, with an action log
python3 scripts/clean.py sales.xlsx --out sales_cleaned.xlsx --report actions.json
```

## Install

**Option 1: upload in WorkBuddy** — Skills → Add skill → Upload skill → select this repo's zip (`SKILL.md` at the root), then start a new conversation.

**Option 2: copy manually** — unzip this repo to `~/.workbuddy/skills/excel-data-doctor/` (`%USERPROFILE%\.workbuddy\skills\excel-data-doctor\` on Windows) and restart WorkBuddy.

**Option 3: one-line install from GitHub** — tell WorkBuddy: `Install this skill: https://github.com/LucasZ97/excel-data-doctor`

## What it detects

| Category | Issues |
|---|---|
| Nulls | Null-rate stats, graded "a few" vs "too many" |
| Duplicates | Exact duplicate rows; near-duplicates ignoring whitespace |
| Dates | Mixed formats (`2024/1/5` vs `2024年1月7日`), unparseable values |
| Numbers | Currency symbols / commas / `%` mixed into numeric columns; unconvertible values flagged individually |
| Text | Leading/trailing whitespace, full-width characters (incl. U+3000) |
| Structure | Merged cells, suspected multi-row headers, fully empty rows/columns, constant columns |

Safe fixes (applied automatically): strip whitespace, full-width → half-width, drop fully empty rows/columns, normalize dates to `YYYY-MM-DD`, strip currency symbols and convert to numbers.
Needs your confirmation: deleting duplicate rows, null-fill strategy, unmerging cells — shown to you first, applied only on approval.

## Measured result

On a 6-row dirty spreadsheet (merged cells, `"1,200元"`, three mixed date formats, full-width spaces): health score **55 → 74** after cleaning. 5 of 9 issues auto-fixed; the remaining 4 (whether to delete the duplicate row, what "unknown" should be) listed for human decision — **it never deletes data on its own**.

## Caveats

- The original file is never modified. Cleaning writes a new file, so every fix is traceable.
- Chinese CSVs are auto-decoded (utf-8-sig → gbk → gb18030), no manual re-encoding needed.
- For files over ~500k rows, sample first: `diagnose.py --sample 5000`.
- Out of scope: tables inside images/scanned PDFs, and judgments about business correctness.

## Layout

```
├── SKILL.md                  # Skill instructions and workflow
├── scripts/diagnose.py       # Diagnosis: 10+ issue checks + health score
├── scripts/clean.py          # Cleaning: one-shot safe fixes + action log
├── references/               # Cleaning recipes (8 field-tested gotchas) + troubleshooting
├── assets/report-template.md # Health report template
└── tests/smoke_test.py       # 17 smoke tests
```

Self-check: `python3 tests/smoke_test.py`

## Changelog

- **v1.1.0**: added `clean.py` one-shot safe fixes; CSV encoding auto-detection; severity counts; troubleshooting guide; 17 smoke tests
- **v1.0.0**: initial release (diagnosis script + cleaning recipes + report template)

## License

[MIT](LICENSE)
