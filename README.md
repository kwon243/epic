# Epic Grouping Toolkit

Python toolkit for making (un)biased groups for ebd.

## What this repo does
- You know what it does.

## Files
- `epic_cli.py`: menu-driven entry point <- run this unless you enjoy pain or a cs major (those two things are the same)
- `make_groups.py`: grouping script
- `epic_excel_check.py`: Excel validation because people in Epic enjoy making life difficult for me.
- `init_db_from_excel.py`: bootstraps `db/people.csv` and `db/pair_scores.csv`
- `db/people.csv`: person records keyed by phone
- `db/pair_scores.csv`: pair recency scores + locked-pair flags

## Requirements
- Python 3.10+
- Packages:
  - `pandas`
  - `openpyxl`

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage
1) Run the CLI:

```bash
python epic_cli.py
```

## TODO
- Make groups editable before final output in `make_groups.py`.
