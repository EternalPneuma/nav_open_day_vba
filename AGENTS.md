# 仓库工作指引

## 项目定位
- 这是 Excel 自动化仓库：Python 负责把 `181-固收上层` 每日源文件合并进 `多套账净值查询.xlsx` 和 `上层产品净值数据库.xlsm`；VBA 负责后续数据计算、分类报表、图表与展示文件输出。
- `README.md` 目前为空；本文件和 `scripts/AGENTS.md`、`scripts/vba/data/AGENTS.md` 是主要说明来源。若说明与代码冲突，以可执行脚本和 VBA 模块为准。

## 当前目录与职责
- `scripts/merge_excel_files.py`：主合并管线；默认输入 `181-固收上层`，默认输出 `多套账净值查询.xlsx`，并用 `keep_vba=True` 更新 `上层产品净值数据库.xlsm` 的 `净值数据` sheet。
- `scripts/tools/analyze_excel_structure.py`：生成单个工作簿结构 JSON，默认输出到 `outputs/analysis/`。
- `scripts/tools/list_sheet_fields.py`：快速列出数据库工作簿每个 sheet 的字段名和非空记录数。
- `scripts/vba/data/`：新版数据侧 VBA，按 `01_auto_input` → `02_calculate_open_date` → `03_output_date` 顺序运行。
- `scripts/vba/chart/`：新版图表侧 VBA，按 `01_auto_input` → `02_check_data` → `03_output_data` → `04_output_chart` → `05_output_image` → `06_output_report` 顺序运行；`t_01_*`、`t_02_*` 是维护/临时工具脚本。
- `vba/` 是旧/备用说明与打包资料；理解当前重构优先看 `scripts/vba/data` 和 `scripts/vba/chart`。

## Python 数据合并约定
- 使用 `uv`；项目配置和 `.python-version` 要求 Python `3.13`，虽然 `merge_excel_files.py` 的运行时保护只检查 `>=3.8`。
- 合并脚本扫描 `.xlsx/.xls`，前 30 行自动识别表头，规范空/重复列名，保留审计列 `source_file`、`source_sheet`。
- 去重只基于业务列哈希；唯一数据写入 `合并数据`，重复数据写入 `重复数据`。
- 写 `.xlsm` 时只替换目标数据 sheet `净值数据`，必须保留宏：`openpyxl.load_workbook(..., keep_vba=True)`。
- `outputs/logs/merge_*.log` 和 `outputs/analysis/*.json` 是生成物；不要手工当作源数据编辑。

## VBA 重构约定
- VBA 源码是可复制到 VBE 的文本模块；当前文件多为**无扩展名**，不要再按旧的 `.txt` 通配路径查找。
- 文件命名：主流程用 `NN_动作`，工具/临时模块用 `t_NN_动作`；内部入口 Sub 多为 `STEP1...`、`STEP2...`，`data` 线还保留 `S1ImportNAV`、`S2CalculateOpenDays`、`S3ProductReport` 英文别名。
- 不要只改文件名而忽略模块内部中文注释/MsgBox/过程名；用户已说明前期重命名尽量不改内部内容以避免 Excel 运行风险。
- Operator-facing 文案、sheet 名、字段名、工作簿名保持中文原文；除非另加注释，不要翻译业务字段。
- 常见实现模式：`Scripting.Dictionary` 建索引/去重/分组，批量数组读写，运行前关闭 `ScreenUpdating/Calculation/Events`，结束恢复设置并用 `MsgBox` 汇总统计。
- 复制到 VBE 前后注意中文编码；除非用户明确要求，不要用自动 Excel COM/VBE 测试去改写这些模块，避免中文标识、sheet 名或提示语被编码破坏。

## 两条 VBA 流水线的输出规范
- `scripts/vba/data/01_auto_input`：从同级工作簿目录下的 `181-固收上层` 增量导入到 `净值数据`，依赖 `HS-181_多账套净值查询_yyyyMMdd.xlsx` 命名、`source_file/source_sheet` 审计列和 30 行表头扫描。
- `scripts/vba/data/02_calculate_open_date`：基于 `净值数据` 最大日期、`开放日`、`产品分类` 写入下一/上一/上上一开放日、实际间隔、基准日期净值等列。
- `scripts/vba/data/03_output_date`：生成 `yyyyMMdd-上层产品分类表现.xlsx`，按分类 sheet 输出周期年化、7 日/28 日年化、成立以来年化等字段。
- `scripts/vba/chart/03_output_data`：从 `Sheet1` 维度和 `Sheet2` 净值数据生成 `产品净值汇总_yyyymmdd.xlsx`，每个产品一个 sheet。
- `scripts/vba/chart/04_output_chart`、`05_output_image`、`06_output_report`：依次生成图表、导出/拼接图片到 `产品图表_yyyymmdd\`（含 `raw\`），再回填展示模板并按日期另存。

## 业务规则不要误改
- `产品分类` 的 `理论间隔`：`1=日开`、`7=周开`、`61=两个月开`、`152=五个月开`、`183=六个月开`，其他值按理论日数间隔处理。
- 节假日后第一个工作日的“未确认”估值可能扭曲天数和年化收益率；未有确认假日字典前，不要假装已自动解决。
- 新产品未满 7/28 天时用“成立以来年化收益率”，不要强算 7 日/28 日年化。
- 成立以来年化的天数分母为 `(结束日期 - 开始日期) + 1`。
- 开放周期计算至少关注三个节点：下一开放日、上一开放日、上上一开放日；如果开放日数学结果不可能，必须显式停止或标记，不能静默继续。
- V0.1 不自动抓取业绩基准；基准由业务人员确认后手工录入。

## 常用命令
```powershell
uv sync
uv run python scripts/merge_excel_files.py
uv run python scripts/merge_excel_files.py "181-固收上层" "多套账净值查询.xlsx"
uv run python scripts/tools/analyze_excel_structure.py
uv run python scripts/tools/list_sheet_fields.py
uv run python scripts/tools/list_sheet_fields.py "上层产品净值数据库.xlsm"
uv run python -m py_compile scripts/merge_excel_files.py scripts/tools/analyze_excel_structure.py scripts/tools/list_sheet_fields.py
```

## 工作区注意事项
- 不要提交或依赖 Excel 锁文件 `~$*.xlsm`；当前 `.gitignore` 未覆盖所有 Excel 临时文件。
- 不要覆盖用户/其他 agent 已改动的工作簿、草稿、`.omo/` 或生成物；本仓库常有 Excel 二进制文件处于外部编辑状态。
- 本环境没有 `rg`；本地搜索优先用可用工具或 PowerShell/AST 工具。
