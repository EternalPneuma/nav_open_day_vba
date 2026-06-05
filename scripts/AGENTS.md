# scripts 目录指引

## 职责
- `merge_excel_files.py` 是 Python 主入口：从 `181-固收上层` 合并每日 Excel，写 `多套账净值查询.xlsx`，并更新 `上层产品净值数据库.xlsm` 的 `净值数据` sheet。
- `tools/analyze_excel_structure.py` 用于生成工作簿结构 JSON；`tools/list_sheet_fields.py` 用于快速查看数据库各 sheet 字段和非空记录数。
- `vba/` 下已有两条重构后的 VBA 流水线：`data/`（导入、开放日计算、分类表现报表）和 `chart/`（导入、检查、汇总、图表、图片、展示回填）。

## Python 约定
- 路径参数按项目根目录相对路径传入；脚本内部通过自身位置推导根目录，不要硬编码绝对路径。
- 日志和操作者可见信息保持中文。
- 保留 `merge_excel_files.py` 的去重契约：业务列哈希去重，唯一行写 `合并数据`，重复行写 `重复数据`。
- 保留 `source_file`、`source_sheet`；它们是每日输入文件审计线索，也被 VBA 增量导入逻辑借鉴。
- 表头探测默认扫描前 30 行；若改动，同步检查 merge、analysis 和 VBA data 导入逻辑。
- 类型转换保持保守：日期列按日期尝试，其他 object 列只有 80% 以上可解析才转数值。
- 写 `.xlsm` 必须 `openpyxl.load_workbook(..., keep_vba=True)`，只替换目标数据 sheet。

## 命令
```powershell
uv sync
uv run python scripts/merge_excel_files.py
uv run python scripts/tools/analyze_excel_structure.py
uv run python scripts/tools/list_sheet_fields.py
uv run python scripts/tools/list_sheet_fields.py "上层产品净值数据库.xlsm"
uv run python -m py_compile scripts/merge_excel_files.py scripts/tools/analyze_excel_structure.py scripts/tools/list_sheet_fields.py
```

## 不要做
- 不要删除或改名 `净值数据`，下游 VBA 依赖它。
- 不要把 `outputs/` 当手工维护输入；它是分析和日志输出。
- 不要把 `scripts/vba` 当“未来目录”；当前已有可运行模块，且文件多为无扩展名文本。
