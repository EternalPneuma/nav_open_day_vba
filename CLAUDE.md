# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel 自动化仓库：Python 合并每日净值源文件，VBA 完成数据计算、分类报表、图表与展示文件输出。面向固收上层产品的净值数据管理。

## Tech Stack

- **Python 3.13** with `pandas`, `openpyxl`, `xlrd` (managed via `uv`)
- **VBA** — 文本模块文件（`.bas` 标准模块 / `.cls` 类模块 / `.frm` UserForm），复制到 Excel VBE 运行；也可通过 `sync-vba.ps1` 自动注入
- 工作簿／sheet／字段名、提示文案、注释保持**中文原文**

## Code Architecture

### Python 管线（scripts/）

| 文件 | 职责 |
|------|------|
| `scripts/merge_excel_files.py` | 主合并入口：从 `181-固收上层/` 扫描 `.xlsx/.xls` 合并到 `多套账净值查询.xlsx`，同时以 `keep_vba=True` 更新 `上层产品净值数据库.xlsm` 的 `净值数据` sheet |
| `scripts/tools/analyze_excel_structure.py` | 生成工作簿结构 JSON → `outputs/analysis/` |
| `scripts/tools/list_sheet_fields.py` | 列出数据库工作簿各 sheet 字段名和非空记录数 |
| `scripts/tools/diagnose_nav_match.py` | 净值数据核对诊断（未正式纳入主流程） |

Python 各脚本通过自身位置推导项目根目录，路径参数按根目录相对路径传入。`outputs/logs/` 和 `outputs/analysis/` 是生成物，不作为源数据。

### VBA 数据管线（scripts/vba/data/）

顺序运行：`01_auto_input` → `02_calculate_open_date` → `03_output_date` → `04_output_report`

| 模块 | 入口（新名称） | 旧名（兼容别名） | 产出 |
|------|------|------|------|
| `01_auto_input.bas` | `Data01_ImportNav181` | `Step01AutoInput`, `S1ImportNAV` | 从 `181-固收上层` 增量导入 `上层产品净值数据(181)` |
| `02_calculate_open_date.bas` | `Data02_CalculateOpenDate` | `Step02CalculateOpenDate`, `S2CalculateOpenDays` | 写入下一/上一/上上一开放日、实际间隔、基准日期净值 |
| `03_output_date.bas` | `Data03_ExportProductReport` | `Step03OutputDate`, `S3ProductReport` | `yyyyMMdd-上层产品分类表现.xlsx` |
| `04_output_report.bas` | `Data04_ExportDisplayReport` | `Step04OutputReport` | `yyyyMMdd-展示.xlsx`（回填展示模板） |

### VBA 图表管线（scripts/vba/chart/）

顺序运行：`01_auto_input` → `02_output_data` → `03_output_chart` → `04_output_image`

| 模块 | 入口（新名称） | 旧名（兼容别名） | 产出 |
|------|------|------|------|
| `01_auto_input.bas` | `Chart01_ImportNavData` | `CSTEP1AutoInputNAV` | 从 `净值数据浏览表 yyyy-mm-dd至yyyy-mm-dd.xlsx` 导入 `绘图净值数据` |
| `02_output_data.bas` | `Chart02_ExportProductSummary` | `CSTEP2OutputData` | `产品净值汇总_yyyymmdd.xlsx`（每个产品一个 sheet） |
| `03_output_chart.bas` | `Chart03_GenerateCharts` | `CSTEP3OutputChart` | 使用 `净值图表_红/蓝.crtx` 和 `收益率图表_红/蓝.crtx` 模板生成图表 |
| `04_output_image.bas` | `Chart04_ExportImages` | `CSTEP4OutputImage` | 导出并拼接图片 → `产品图表_yyyymmdd/`（含 `raw/`） |

### 其他 VBA

- `scripts/vba/weekly_recommendation/` — `Weekly01_UpdateDependencies` → `Weekly02_GenerateReport`（旧名 `RSTEP1UpdateWeeklyRecommendationDependencies`、`RSTEP2GenerateWeeklyRecommendation`），输出 `推荐材料-yyyymmdd.xlsx`
- `scripts/vba/tool/` — 维护工具：`Tool01_CleanDuplicateData`（旧名 `清洗重复数据`）、`Tool02_DeleteByProductId`（旧名 `批量删除数据`）、`Tool03_FillNextOpenDate`（旧名 `TOOL2FillNextOpenDateByInterval`）、`Tool04_CheckNavData`（旧名 `核对净值数据`）
- `scripts/vba/optional_panel/` — 操作面板：`00_operation_panel.bas`（标准模块）+ `00_operation_panel_button.cls`（类模块）+ `00_operation_panel_form.frm`（UserForm）
- `scripts/vba/chart/archive/` — 旧 chart 流程归档（均已加 `.bas` 后缀），仅参考

### 工作簿依赖关系

```
181-固收上层/ (每日源文件)
    ↓ Python: merge_excel_files.py
多套账净值查询.xlsx  ──→  上层产品净值数据库.xlsm (净值数据 sheet)
                                ↓ VBA data 管线
                           yyyyMMdd-上层产品分类表现.xlsx
                                ↓ VBA data 04
                           yyyyMMdd-展示.xlsx
                           
上层产品净值数据库.xlsm (绘图净值数据 sheet)
    ↓ VBA chart 管线
产品净值汇总_yyyymmdd.xlsx  →  产品图表_yyyymmdd/ (图片)
```

## Key Business Rules

- **理论间隔**: `1=日开`, `7=周开`, `61=两个月开`, `152=五个月开`, `183=六个月开`；其他值按实际日数处理
- **新产品不足 7/28 天**时使用"成立以来年化收益率"，不强算 7 日/28 日年化；成立以来年化分母 = `(结束日期 - 开始日期) + 1`
- **节假日影响**：节后首个工作日可能有"未确认"估值，尚无假日字典前不得静默修正
- **开放日计算不可能时**必须显式停止或标记，不要静默继续
- **V0.1 不自动抓取业绩基准**，由业务人员手工录入
- weekly 推荐找不到净值或收益率时显示 `-`，不得静默写入 0

## Common VBA Patterns

- `Scripting.Dictionary` 建索引／去重／分组
- 批量数组读写（`Variant` 数组整体赋值 `Range.Value`）
- 运行前关闭 `ScreenUpdating`/`Calculation`/`EnableEvents`/`DisplayAlerts`，结束后恢复并用 `MsgBox` 汇总
- 增量导入依赖 `source_file`/`source_sheet` 审计列和 30 行表头扫描

## Commands

```bash
# Python 环境
uv sync

# 合并净值数据（默认输入 181-固收上层 → 多套账净值查询.xlsx）
uv run python scripts/merge_excel_files.py

# 指定输入/输出
uv run python scripts/merge_excel_files.py "181-固收上层" "多套账净值查询.xlsx"

# 分析工作簿结构
uv run python scripts/tools/analyze_excel_structure.py

# 查看数据库 sheet 字段
uv run python scripts/tools/list_sheet_fields.py
uv run python scripts/tools/list_sheet_fields.py "上层产品净值数据库.xlsm"

# 静态检查 Python 语法
uv run python -m py_compile scripts/merge_excel_files.py scripts/tools/analyze_excel_structure.py scripts/tools/list_sheet_fields.py
```

## VBA 模块同步工具

`sync-vba.ps1` — PowerShell 脚本，通过 Excel COM 将 `scripts/vba/` 中的模块自动注入到目标 `.xlsm` 工作簿。

```powershell
# 从 WSL2 同步全部模块到默认工作簿
powershell.exe -File sync-vba.ps1

# 仅同步 data 和 optional_panel 模块
powershell.exe -File sync-vba.ps1 -ModuleGroups "data,optional_panel"

# 指定工作簿
powershell.exe -File sync-vba.ps1 -WorkbookPath "上层产品净值数据库.xlsm"
```

**前置条件**：Excel 信任中心 → 宏设置 → 勾选"信任对 VBA 工程对象模型的访问"。运行前关闭目标工作簿。

## VBA 模块操作注意事项

- 模块文件后缀：`.bas`（标准模块）、`.cls`（类模块）、`.frm`（UserForm），统一 CRLF 行尾
- 复制内容到 VBE 时注意中文编码；或者用 `sync-vba.ps1` 自动注入
- 改文件名时不要顺手改内部过程名、注释或 MsgBox 文案（Excel 运行时依赖过程名和 `ThisWorkbook` 引用）
- `.xlsm` 工作簿经常处于外部编辑状态，不要覆盖改动中的二进制文件

## Git & Workspace

- 当前分支 `master`，`.gitignore` 已配置 `~$*` 锁定文件忽略、`outputs/logs/` 和 `outputs/analysis/` 生成目录
- `.gitattributes` 强制 VBA 模块 CRLF、Python LF，Excel 二进制文件标记为 binary
- `vba/` 目录（根目录）是旧资料，当前重构以 `scripts/vba/` 为准
