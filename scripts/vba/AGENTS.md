# `scripts/vba` 工作指引

## 定位与维护方式

- 此目录中的文件是可复制到 Excel VBE 的 VBA 文本模块；当前多数文件没有扩展名。
- 不要将这些模块改写为 Excel COM/VBE 自动导入流程。复制到 VBE 后由用户手动运行、验证，避免破坏中文编码、工作表名称和提示文案。
- 面向业务人员的工作簿名、sheet 名、字段名、注释和 MsgBox 文案保持中文原文；不要为重命名或规范化而翻译。
- 主流程文件采用 `NN_动作` 命名；维护工具采用 `t_NN_动作` 命名。改名时不要顺带修改过程名或内部文案，除非任务明确要求。

## 目录与运行顺序

| 目录 | 用途 | 主流程/入口 |
| --- | --- | --- |
| `data/` | 上层产品净值数据计算与分类展示 | `01_auto_input` → `02_calculate_open_date` → `03_output_date` → `04_output_report`；别名入口为 `S1ImportNAV`、`S2CalculateOpenDays`、`S3ProductReport` |
| `chart/` | 当前重构后的图表流程 | `01_auto_input` → `02_output_data` → `03_output_chart` → `04_output_image`；入口为 `CSTEP1AutoInputNAV` 至 `CSTEP4OutputImage` |
| `chart/archive/` | 旧 chart 流程归档，仅供回溯和迁移参考 | 不作为当前运行入口；包含原 `01`–`06` 模块 |
| `tool/` | 非主流程的维护工具 | `t_01_clean_data`、`t_02_del_data`、`t_03_next_open_date_by_interval`；运行前确认目标工作簿、sheet 和影响范围 |
| `optional_panel/` | 可选的 VBA 操作面板 | `00_operation_panel`（标准模块）、`00_operation_panel_button`（类模块）、`00_operation_panel_form`（UserForm 代码）；需分别按对应的 VBE 模块类型导入 |
| `weekly_recommendation/` | 每周推荐材料输出 | `01_update_weekly_recommendation_dependencies` → `02_generate_weekly_recommendation`；入口为 `RSTEP1UpdateWeeklyRecommendationDependencies`、`RSTEP2GenerateWeeklyRecommendation` |

## 数据侧规则

- `data/01_auto_input` 从同级工作簿目录中的 `181-固收上层` 增量导入 `净值数据`；依赖 `HS-181_多账套净值查询_yyyyMMdd.xlsx` 命名、`source_file/source_sheet` 审计列和前 30 行表头扫描。
- `data/02_calculate_open_date` 以 `净值数据` 最大日期、`开放日` 和 `产品分类` 为依据，写入下一/上一/上上一开放日、实际间隔和基准日期净值等列。
- `data/03_output_date` 输出 `yyyyMMdd-上层产品分类表现.xlsx`；`data/04_output_report` 以该文件为输入，输出 `yyyyMMdd-展示.xlsx`。
- 产品分类的理论间隔固定含义为：`1=日开`、`7=周开`、`61=两个月开`、`152=五个月开`、`183=六个月开`；其他数值按理论日数间隔处理。
- 新产品不足 7 或 28 天时使用成立以来年化收益率，不强算 7 日或 28 日年化；成立以来年化的天数分母为 `(结束日期 - 开始日期) + 1`。
- 节假日后的首个工作日可能存在“未确认”估值，尚未具备确认假日字典前不得静默修正该影响。开放日计算不可能时必须显式停止或标记。
- V0.1 不自动抓取业绩基准，基准由业务人员确认后手工录入。

## 图表与推荐材料规则

- `chart/` 根目录是唯一当前 chart 主流程。它生成产品净值汇总、图表和图片素材；最终展示文件由 `data/04_output_report` 负责，不要扩展归档的 `chart/archive/06_output_report`。
- `weekly_recommendation/01_update_weekly_recommendation_dependencies` 先更新依赖数据，再运行 `02_generate_weekly_recommendation` 输出 `推荐材料-yyyymmdd.xlsx`。
- weekly 推荐材料的基准日期取 `产品分类` sheet 的 `基准日期`，所有非空值必须一致；找不到净值或收益率时显示 `-`，不得静默写入 0。

## 实现与验证边界

- 优先沿用 `Scripting.Dictionary`、批量数组读写，并在运行中妥善恢复 `ScreenUpdating`、`Calculation`、`Events` 等 Excel 设置。
- 涉及删除、清洗或覆盖数据的工具，先核对目标 sheet、主键和影响行数；不要把维护工具加入主流程或面板的一键流程。
- 修改后以静态检查、字段核对和用户在 Excel 中的手动运行验证为主。除非用户明确要求，不读取或改写大型 `.xlsm` 工作簿来验证 VBA。
