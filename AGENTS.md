# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-25
**Commit:** 249bd11
**Branch:** master

## OVERVIEW
Excel automation repo for multi-account NAV data: Python merges daily source workbooks into a macro-enabled database, then VBA text scripts are intended to refresh sheets, calculate 7/28-day annualized returns, derive open dates, and flag anomalies.

## STRUCTURE
```
nav_open_day_vba/
├── 181-固收上层/              # daily HS-181 多账套净值查询 source workbooks; hundreds of dated .xlsx files
├── scripts/                  # Python utilities and future VBA text scripts
│   ├── merge_excel_files.py   # main import/merge pipeline
│   ├── tools/                 # workbook structure analysis helper
│   └── vba/                   # plain-text VBA modules copied manually into VBE
├── outputs/
│   ├── analysis/              # JSON workbook structure reports
│   └── logs/                  # timestamped merge logs
├── drafts/                    # recordings and working Excel drafts, not runtime inputs
├── 多套账净值查询.xlsx          # merged workbook output
└── 上层产品净值数据库.xlsm       # macro-enabled database updated by Python; VBA work continues here
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Merge daily 181 source files | `scripts/merge_excel_files.py` | Default input `181-固收上层`, default output `多套账净值查询.xlsx` |
| Inspect workbook layout | `scripts/tools/analyze_excel_structure.py` | Emits JSON under `outputs/analysis` |
| Confirm source columns | `outputs/analysis/HS-181_多账套净值查询_20260220.structure.json` | Known headers: 账套名称, 日期, 资产净值, 单位净值, 总资产, 总负债, 实收信托, 总估值增值, 信托计划代码, 信托计划累计净值 |
| Add VBA logic | `scripts/vba/*.txt` | Store as text only; user manually copies into Excel VBE |
| Runtime logs | `outputs/logs/merge_*.log` | Merge script writes INFO/ERROR lines here |
| Business workbook | `上层产品净值数据库.xlsm` | `merge_excel_files.py` rewrites sheet `净值数据` with `keep_vba=True` |
| Product open cycle | sheet `产品分类` column H `理论间隔` | 1=日开, 7=周开, 61=两个月开, 152=五个月开, 183=六个月开, 其余值为理论日数间隔 |

## CODE MAP
| Symbol | Type | Location | Role |
|--------|------|----------|------|
| `merge_excels` | function | `scripts/merge_excel_files.py` | Read all `.xlsx/.xls` source files, normalize headers/types, dedupe rows, write merged workbook and `上层产品净值数据库.xlsm` |
| `_detect_header_row` | function | `scripts/merge_excel_files.py`, `scripts/tools/analyze_excel_structure.py` | Scan first 30 rows and choose the densest likely header row |
| `_make_unique_columns` | function | both Python scripts | Normalize blank/duplicate Excel headers to stable names |
| `analyze_excel_structure` | function | `scripts/tools/analyze_excel_structure.py` | Produce JSON metadata for sheet names, header row, shape, dtypes, null ratios |

## CONVENTIONS
- Python project uses `uv`; `.python-version` is `3.13`, while script runtime guards only require Python `>=3.8`.
- Dependencies are intentionally small: `pandas`, `openpyxl`, `xlrd` in `pyproject.toml`.
- Paths are project-root relative in script arguments; Python helpers compute the repo root from their own file location.
- Excel output preserves macros only for `上层产品净值数据库.xlsm`; do not convert it to `.xlsx` when updating database content.
- Chinese business field names are canonical. Avoid translating sheet names, column headers, workbook names, or log messages unless adding separate comments.
- `scripts/vba` is the authoritative location for VBA source; code is stored as `.txt`, not embedded-only in workbook files.

## BUSINESS RULES FOR VBA WORK
- Data source sheets: imported `净值数据`/`多套账净值查询` data plus the “7日28日年化收益率” worksheet once present.
- Avoid data pulls on the first workday after holidays; holiday “未确认” valuations can distort day counts and annualized returns.
- New products with age below 7 or 28 days use “成立以来年化收益率”, not period annualization.
- Day-count divisor for inception calculations is `(结束日期 - 开始日期) + 1`.
- Next open date is `当前开放日 + 间隔天数`; if the result is `<=` operation date, treat the row as abnormal and stop/flag it.
- VBA should compare three nodes for cycle return work: next open date, current cycle start, previous cycle start.
- V0.1 does not auto-fetch performance benchmarks; business users enter benchmark data manually after product-manager confirmation.

## COMMANDS
```powershell
uv sync
uv run python scripts/merge_excel_files.py
uv run python scripts/merge_excel_files.py "181-固收上层" "多套账净值查询.xlsx"
uv run python scripts/tools/analyze_excel_structure.py
uv run python scripts/tools/analyze_excel_structure.py "181-固收上层\HS-181_多账套净值查询_20260220.xlsx" "outputs\analysis\HS-181_多账套净值查询_20260220.structure.json"
uv run python -m py_compile scripts/merge_excel_files.py scripts/tools/analyze_excel_structure.py
```

## ANTI-PATTERNS (THIS PROJECT)
- Do not commit or rely on Excel lock files like `~$*.xlsm`; one is currently present as a working artifact.
- Do not edit generated `outputs/logs` as source of truth; regenerate by running scripts.
- Do not add packaging around VBA unless explicitly requested; current operating model is copy `.txt` into VBE manually.
- Do not assume holidays are solved in code yet; the rule is a business control pending a confirmed holiday dictionary/manual gate.
- Do not overwrite unrelated workbook deletions/modifications in the worktree; multiple Excel artifacts are already changed outside this task.

## NOTES
- `README.md` is empty; this file is the current project map.
- `rg` is not installed in this environment; use PowerShell search or AST tools for local discovery.
- No test suite, CI workflow, Makefile, batch script, or `requirements*.txt` was found; validation is script/manual-output based.
- Existing git state before this knowledge-base update included modified `pyproject.toml`, `uv.lock`, `上层产品净值数据库.xlsm`, deleted Excel/VBA artifacts, and untracked `.omo/`, `drafts/`, and an Excel lock file.
