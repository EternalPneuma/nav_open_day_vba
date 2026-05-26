# VBA KNOWLEDGE BASE

## OVERVIEW
VBA source belongs here as plain `.txt` modules copied manually into Excel VBE; this directory is for report refresh, return calculation, open-date derivation, and anomaly marking.

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Add VBA module text | `scripts/vba/*.txt` | Keep copy-paste friendly for VBE; no packaging step |
| Read merged NAV data | workbook sheet `净值数据` | Created by `scripts/merge_excel_files.py` in `上层产品净值数据库.xlsm` |
| Compare source layout | `outputs/analysis/*.structure.json` | Confirms imported NAV headers and row shape |
| Business requirements | root `AGENTS.md` | Holds current 7/28-day, open-day, holiday, and benchmark rules |

## CONVENTIONS
- Store VBA as `.txt`, not `.bas`/embedded-only, unless the user changes the execution model.
- Put sheet names and column indexes/names in constants at the top of the module; layout mapping is still pending and must be easy to adjust.
- Use Chinese procedure comments/messages for operator-facing warnings.
- Keep the main macro copy-paste runnable: a business user should be able to open the `.txt`, copy into VBE, and run the top-level `Sub`.

## BUSINESS LOGIC CHECKLIST
- Before calculation, warn/stop when the operation date is a holiday-after-first-workday risk date once the holiday rule is formalized.
- For each product row, calculate `下一次开放日 = 当前开放日 + 间隔天数` and reject/flag if `下一次开放日 <= Date`.
- Fetch/compare the three cycle nodes needed for single-period return: 下一次开放日, 当前周期开始日, 上一个周期开始日.
- For age under 7 or under 28 days, write “成立以来年化收益率” instead of 7/28-day annualized return.
- In inception annualization, use inclusive days: `(End_Date - Start_Date) + 1`.
- If calculated return is above peer/average threshold such as `Average_Value * 1.1`, flag the row for manual review or rerun day-count correction.
- Do not implement benchmark auto-fetch in V0.1; operators manually enter confirmed benchmark data.

## ANTI-PATTERNS
- Do not hide VBA only inside `.xlsm`; text source is required for maintainability.
- Do not run automated VBA/Excel COM tests for these text modules; VBE/COM encoding can corrupt Chinese identifiers, sheet names, and messages, so validate by reading/static review unless the user explicitly requests a manual Excel run.
- Do not silently continue after impossible open-date math; stop or visibly flag the product row.
- Do not calculate new-product 7/28-day annualization before the product has enough days.
- Do not omit the `+1` day-count rule; it prevents early-life annualized return spikes.
