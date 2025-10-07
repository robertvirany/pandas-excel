# pandas-excel

## Workflow
1. Run `python3 manual_sql_loader.py --directory path/to/sql --connection-string "DSN=netezza"` to emit `generated_sql_loader.py` (edit creds if needed).
2. Run `python3 translate_formulas.py workbook.xlsx` to emit `generated_formulas.py` (use `--verbose` for per-sheet stats).
3. Load data: `dfs = load_data()` from the generated loader, align keys with sheet names, then import `generated_formulas.py` and evaluate the assignments.

## Tests
`python3 -m unittest discover -s tests`
