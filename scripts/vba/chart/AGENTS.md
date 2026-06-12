# scripts/vba/chart 迁移指引

## 迁移目标
- 旧环境：`vba/package/产品净值数据库.xlsm`。
- 新环境：仓库根目录 `上层产品净值数据库.xlsm`。
- 本目录 VBA 仍按文本模块维护，复制到 Excel VBE 后运行；除非用户明确要求，不要自动用 Excel COM/VBE 改写模块。
- 两个数据库文件都较大，后续排查和修改以 VBA 文本源码、用户提供的 sheet/字段信息为准，不要为了确认结构直接读取整本工作簿。

## 最终集成目标
- 最终展示文件以 `scripts/vba/data/04_output_report` 为主线生成；该模块已经在展示报表中预留图表/图片插入空位。
- `scripts/vba/chart/04_output_chart` 和 `05_output_image` 后续应只负责生成图表和图片素材，供 data 线 `04_output_report` 插入。
- `scripts/vba/chart/06_output_report` 是旧图表侧展示回填流程，迁移完成后应由 `scripts/vba/data/04_output_report` 的预留空位插图逻辑替代，不再作为最终展示报告入口。
- 后续修改不要继续扩展 `chart/06_output_report` 的展示模板能力；需要新增或调整展示版式时，优先落在 `data/04_output_report`。
- 集成后的目标流程为：data 线生成分类表现和展示框架，chart 线生成产品净值汇总/图表/图片，最后由 `data/04_output_report` 在预留位置插入图表素材并输出最终展示文件。

## 已确认 sheet 对应关系
| 旧数据库对象 | 新数据库对象 | 用途 | 迁移要求 |
| --- | --- | --- | --- |
| `ThisWorkbook` = `产品净值数据库.xlsm` | `ThisWorkbook` = `上层产品净值数据库.xlsm` | chart 线宏所在数据库工作簿 | 代码仍可使用 `ThisWorkbook`，但语义变为新数据库。 |
| `Sheet1` | `产品信息` | 产品维度/信息表，常用 A:C 列读取产品编号、产品名称、产品简称 | 数据库端所有 `wbDB.Sheets("Sheet1")` 应迁移为 `wbDB.Sheets("产品信息")`。 |
| `Sheet2` | `绘图净值数据` | 绘图用净值明细表，常用 A:J 列；主键通常为日期 + 产品编号 | 数据库端所有 `wbDB.Sheets("Sheet2")` 或针对 Sheet2 的注释/提示语应迁移为 `wbDB.Sheets("绘图净值数据")` 和对应中文文案。 |
| `Sheet3` | 导出工作簿中的 `数据摘要` | 数据完整度/填充统计摘要 | 不再写回数据库；`02_check_data` 的报告逻辑并入 `03_output_data`，在 `产品净值汇总_yyyymmdd.xlsx` 中新增 `数据摘要` sheet。 |

## 已确认流程调整
- `02_check_data` 不再作为单独写回数据库 `Sheet3` 的模块使用；其数据完整度报告能力应合并到 `03_output_data`。
- `03_output_data` 生成 `产品净值汇总_yyyymmdd.xlsx` 时，应在导出工作簿中新增 `数据摘要` sheet，承载原 `02_check_data` 和 `WriteFillCountsToSheet3` 的摘要/统计结果。
- `数据摘要` 是输出文件中的普通 sheet，不是 `上层产品净值数据库.xlsm` 中的数据库 sheet；运行宏时不要创建、清空或覆盖数据库内的 `Sheet3`。
- `01_auto_input` 读取每日导入源文件时，默认读取源工作簿的第一个 sheet。该源 sheet 不映射到数据库 `产品信息`，后续代码应优先使用 `wbSrc.Worksheets(1)` 或等价逻辑，而不是硬编码源文件 `Sheet1`。

## 暂未确认或不属于本次映射
- `04_output_chart`、`05_output_image` 主要处理 `03_output_data` 生成的 `产品净值汇总_yyyymmdd.xlsx` 和图片目录，当前未发现直接读取数据库 `Sheet1/Sheet2` 的逻辑。迁移时仍需检查工作簿路径和输出命名，但不应套用本次 sheet 映射。
- `06_output_report` 属于旧图表侧展示回填模块，后续方向是被 `scripts/vba/data/04_output_report` 替代；除非为了过渡兼容，不应继续作为最终展示入口维护。

## 当前源码影响范围
- `01_auto_input`
  - 数据库目标：`Dim wsDB As Worksheet: Set wsDB = wbDB.Sheets("Sheet2")`，应改为 `绘图净值数据`。
  - 注释和 MsgBox 中出现的 `Sheet2` 指数据库目标表，应同步改成 `绘图净值数据`。
  - 源文件读取：`Set wsSrc = wbSrc.Sheets("Sheet1")` 应改为默认读取源工作簿第一个 sheet，例如 `wbSrc.Worksheets(1)`。
- `02_check_data`
  - 数据库维度表：`wbDB.Sheets("Sheet1")`，应改为 `产品信息`。
  - 数据库净值表：`wbDB.Sheets("Sheet2")`，应改为 `绘图净值数据`。
  - 独立写数据库 `Sheet3` 的逻辑应迁移进 `03_output_data`，后续不再向数据库插入或清空 `Sheet3`。
- `03_output_data`
  - 数据库维度表：`wbDB.Sheets("Sheet1")`，应改为 `产品信息`。
  - 数据库净值表：`wbDB.Sheets("Sheet2")`，应改为 `绘图净值数据`。
  - `WriteFillCountsToSheet3` 应改造为向导出工作簿新增/写入 `数据摘要` sheet，不再接收数据库工作簿作为写入目标。
  - 原 `02_check_data` 的产品级统计摘要应合并到本模块输出流程中，和产品明细 sheet 一起写入同一个导出工作簿。
  - 相关提示语中“维度表(Sheet1)”“Sheet2中存在/无数据”等应随逻辑名同步更新，避免运行时提示误导。
- `t_01_clean_data`
  - 清洗目标：`wbDB.Sheets("Sheet2")`，应改为 `绘图净值数据`。
  - 删除重复的主键仍是“净值日期 + 产品编号”，未确认字段顺序改变前保持现有列位逻辑。
- `t_02_del_data`
  - 删除目标：`wbDB.Sheets("Sheet2")`，应改为 `绘图净值数据`。
  - 用户提示中的 `Sheet2` 应同步改为 `绘图净值数据`。
- `04_output_chart` / `05_output_image`
  - 保留为图表和图片素材生成步骤，输出应服务于 `data/04_output_report` 的插图需求。
  - 后续需要确认图片命名、日期目录和产品/系列匹配规则，确保 `data/04_output_report` 能稳定定位素材。
- `06_output_report`
  - 旧职责是把图表图片回填展示模板并另存展示文件。
  - 新职责应逐步取消；其可复用的图片查找、插入、尺寸适配逻辑可以迁移到 `data/04_output_report`，但最终入口不再放在 chart 线。

## 字段与列位假设
- `产品信息` 暂按旧 `Sheet1` 的 A:C 列语义迁移：
  - A 列：产品编号。
  - B 列：产品名称。
  - C 列：产品简称。
- `绘图净值数据` 暂按旧 `Sheet2` 的 A:J 列语义迁移：
  - A 列：净值日期/日期。
  - B 列：产品编号。
  - D 列：净值，供并入 `03_output_data` 的摘要统计使用。
  - A:J 整段供 `01_auto_input` 增量写入和 `03_output_data` 按产品导出使用。
- 如果新表字段名相同但列顺序不同，不能只改 sheet 名；需要先把代码改为按表头定位字段，再替换当前固定列位读取。

## 后续修改原则
- 优先新增集中常量，例如 `SHEET_PRODUCT_INFO = "产品信息"`、`SHEET_CHART_NAV_DATA = "绘图净值数据"`、`SHEET_DATA_SUMMARY = "数据摘要"`，再替换硬编码 sheet 名。
- 用户可见的中文 sheet 名、字段名、MsgBox 文案保持中文原文，不要翻译。
- 每次只改一到两个模块，先做静态搜索核对，再由用户在 Excel 中手动运行验证。
- 不要把 `scripts/vba/data` 线的 `净值数据`、`产品分类`、`开放日` 逻辑混入 chart 线；本目录迁移只处理图表侧数据、图表、图片、展示输出流水线。
- 涉及最终展示文件的改动，以 `scripts/vba/data/04_output_report` 为汇总入口；chart 线只提供素材，不再单独完成最终报告。
