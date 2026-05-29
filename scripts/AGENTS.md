# SCRIPTS KNOWLEDGE BASE

## OVERVIEW
Python utilities prepare Excel data for the later VBA workflow; keep them root-relative, log-heavy, and compatible with Chinese workbook names.

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Main merge/import | `merge_excel_files.py` | Reads source directory, writes merged workbook, updates `.xlsm` database sheet |
| Structure report | `tools/analyze_excel_structure.py` | Single-workbook metadata extractor |
| Quick field listing | `tools/list_sheet_fields.py` | Prints every sheet's field names and non-null record counts to console |
| Future VBA modules | `vba/` | Child scope has VBA-specific rules |

## CONVENTIONS
- Script entry points accept optional positional paths; defaults are embedded near each `main()`.
- Keep log/user messages in Chinese to match operator workflow.
- Preserve the duplicate-handling contract in `merge_excel_files.py`: hash business columns, write unique rows to `合并数据`, duplicate rows to `重复数据`.
- Preserve source traceability columns `source_file` and `source_sheet` in merged data.
- Header detection scans up to 30 rows; if changing it, update both merge and analysis scripts or document why they diverge.
- Type conversion is conservative: date-like column names convert to datetime; other object columns convert to numeric only when at least 80% parse.
- For `.xlsm` writes, use `openpyxl.load_workbook(..., keep_vba=True)` and replace only the target data sheet.

## COMMANDS
```powershell
uv sync
uv run python scripts/merge_excel_files.py
uv run python scripts/tools/analyze_excel_structure.py
uv run python scripts/tools/list_sheet_fields.py
uv run python scripts/tools/list_sheet_fields.py "上层产品净值数据库.xlsm"
uv run python -m py_compile scripts/merge_excel_files.py scripts/tools/analyze_excel_structure.py
```

## ANTI-PATTERNS
- Do not hard-code absolute paths; current scripts are portable from repo root.
- Do not drop `source_file`/`source_sheet`; they are the audit trail for daily Excel inputs.
- Do not change the database sheet name `净值数据` without updating downstream VBA assumptions.
- Do not make `outputs/` a hand-edited input; it is generated analysis/log output.
