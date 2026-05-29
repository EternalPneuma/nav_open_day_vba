# Step 3: 产品分类表现报告生成

## TL;DR

> **Quick Summary**: 创建 VBA 模块，从 `上层产品净值数据库.xlsm` 读取 Step 2 输出，按产品分类（稳享长期限/直销/交行代销）生成三页 `.xlsx` 报告，含净值日期、年化收益率计算。
> 
> **Deliverables**:
> - `scripts/vba/03_产品分类表现报告.txt` — VBA 源码模块
> - `yyyyMMdd-上层产品分类表现.xlsx` — 生成的三页报告
> 
> **Estimated Effort**: Medium
> **Parallel Execution**: YES — 5 waves + FINAL
> **Critical Path**: Task 1 → Task 3 → Task 7

---

## Context

### Original Request
用户希望在 Step 1（净值导入）和 Step 2（开放日测算）之后，编写 Step 3 VBA 程序：根据产品分类将产品数据分为三类，生成一个包含三个 sheet 的新 `.xlsx` 文件，每类产品有不同的扩展字段和年化收益率计算。

### Interview Summary

**Key Discussions**:
- **输出结构**: 共三个 sheet — "稳享长期限"(63个产品)、"直销"(10个)、"交行代销"(110个)
- **通用字段**: 序号、信托计划代码、系列、产品名称、上一开放日、上一开放日净值、基准日期、基准日期净值、下一开放日
- **分类特有字段**: 每个分类有不同的额外列和年化计算需求
- **年化公式**: 所有公式不使用 +1 日计数规则；7日/28日用精确日历天查找，无数据则留空
- **缺失数据处理**: 保留产品行，计算字段留空；日开/周开产品包含但留空
- **输出路径**: 与 `上层产品净值数据库.xlsm` 同级目录；序号保留原始值
- **AGENTS.md 显式覆盖**: 不使用 +1 规则；不应用新产品成立以来年化替代逻辑（用户决定留空）

**Research Findings**:
- 产品分类工作表已有完整 Step 2 输出（19 列），可直接读取
- 净值数据有 48,039 行，支持按信托计划代码+日期查询
- 现有 VBA 模式：Dictionary 查找、BuildHeaderMap、错误处理包装器
- 48,039 行中约 37% 的单位净值为空（非每日估值产品）

### Metis Review

**Identified Gaps** (addressed):
- 覆盖行为未定义 → 默认自动覆盖（与 Step 1/2 一致）
- 列匹配策略 → 按表头名称匹配（沿用现有 BuildHeaderMap 模式）
- 重复 NAV 处理 → 保留第一条（沿用 Step 2 模式）
- 输出格式化 → 年化收益存小数，NumberFormat "0.00%"
- 排序规则 → 保持产品分类原始行序
- AGENTS.md 规则冲突 → 用户显式覆盖：不用 +1，不用新产品规则
- 测试策略未定义 → 声明为 none + agent QA

---

## Work Objectives

### Core Objective
在 `scripts/vba/` 下创建单个 VBA 文本模块 `03_产品分类表现报告.txt`，该模块读取 `上层产品净值数据库.xlsm` 的产品分类数据，按三个分类筛选产品并计算各自的年化收益率，生成三页 `.xlsx` 报告文件。

### Concrete Deliverables
- `scripts/vba/03_产品分类表现报告.txt` — 完整 VBA 模块源码
- 运行后在数据库同级目录生成 `yyyyMMdd-上层产品分类表现.xlsx`

### Definition of Done
- [ ] VBA 模块可复制到 VBE 中运行，无编译错误
- [ ] 运行后在正确路径生成 `.xlsx` 文件
- [ ] 生成的 `.xlsx` 包含恰好三个 sheet：稳享长期限、直销、交行代销
- [ ] 每个 sheet 包含正确的列（通用 + 分类特有）
- [ ] 年化收益率使用确认的公式计算
- [ ] 缺失数据的产品行保留、计算字段留空
- [ ] 日开/周开产品包含在输出中，开放日相关字段留空
- [ ] 源工作簿未被修改

### Must Have
- 按表头名称定位列（BuildHeaderMap 模式）
- NAV 查找支持按任意日期查询（不仅限于基准日期）
- 7日/28日使用精确 N 日历天查找，无匹配则留空
- 年化计算使用确认的无 +1 公式
- 缺失净值时计算结果留空（不崩溃、不显示 #DIV/0!）
- 输出文件不含宏（纯 .xlsx）
- 序号保留产品分类原始值

### Must NOT Have (Guardrails)
- 不修改 `上层产品净值数据库.xlsm`（只读）
- 不修改 Step 1 或 Step 2 的 VBA 模块
- 不添加节假日/日历逻辑
- 不使用 +1 日计数规则（用户显式覆盖 AGENTS.md）
- 不对新产品应用成立以来年化替代逻辑（用户显式覆盖）
- 7日/28日不使用最近日期回退（精确日历天）
- 不添加报告样式美化（Logo、合并单元格、打印布局）
- 不添加自动发送/邮件功能
- 不添加 Ribbon/菜单集成
- 不重复计算或更改产品分类中的已有字段
- 不实现基准收益率自动获取

---

## Verification Strategy (MANDATORY)

> **ZERO HUMAN INTERVENTION** - ALL verification is agent-executed.

### Test Decision
- **Infrastructure exists**: NO
- **Automated tests**: None (VBA .txt 文件，手动复制到 VBE)
- **Framework**: N/A
- **Agent-Executed QA**: MANDATORY — 使用 Python/openpyxl 自动化验证生成的 .xlsx

### QA Policy
Every task MUST include agent-executed QA scenarios.
Evidence saved to `.omo/evidence/task-{N}-{scenario-slug}.{ext}`.

- **VBA 源码验证**: 使用 Bash 检查 .txt 文件内容（语法模式、表头常量、公式正确性）
- **输出 .xlsx 验证**: 使用 Python/openpyxl 读取生成的文件，验证 sheet 名、列名、行数、公式结果、边界情况

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 1 (Start Immediately - foundation):
└── Task 1: Module scaffolding + constants + shared helpers [quick]
    ↓ (Task 1 must complete first — it provides shared code all others depend on)

Wave 2 (After Task 1 - infrastructure, MAX PARALLEL):
├── Task 2: NAV lookup infrastructure (date-based Dictionary) [quick]
└── Task 3: Output workbook scaffolding (create .xlsx, write headers) [quick]

Wave 3 (After Waves 1+2 - data writing):
└── Task 4: Common field writer (read 产品分类, write data to sheets) [quick]

Wave 4 (After Wave 3 - category computation, MAX PARALLEL):
├── Task 5: 稳享长期限 computation (当前周期年化) [quick]
├── Task 6: 直销 computation (基准收益率, 7日/28日年化, 成立以来年化留空) [quick]
└── Task 7: 交行代销 computation (上一周期年化, 当前周期年化) [quick]

Wave 5 (After Wave 4 - integration + polish):
└── Task 8: Main orchestration subroutine + formatting + save [quick]

Wave FINAL (After ALL tasks — 4 parallel reviews, then user okay):
├── F1: Plan compliance audit (oracle)
├── F2: Code quality review (unspecified-high)
├── F3: Real manual QA — .xlsx verification (unspecified-high + Python)
└── F4: Scope fidelity check (deep)
-> Present results -> Get explicit user okay

Critical Path: Task 1 → Task 3 → Task 4 → Task 5/6/7 → Task 8 → F1-F4 → user okay
Parallel Speedup: ~40% faster than sequential
Max Concurrent: 3 (Waves 4 & FINAL)
```

### Dependency Matrix

| Task | Depends On | Blocks | Wave |
|------|-----------|--------|------|
| 1 | - | 2, 3, 4, 5, 6, 7, 8 | 1 |
| 2 | 1 | 5, 6, 7 | 2 |
| 3 | 1 | 4, 5, 6, 7, 8 | 2 |
| 4 | 1, 3 | 5, 6, 7, 8 | 3 |
| 5 | 1, 2, 3, 4 | 8 | 4 |
| 6 | 1, 2, 3, 4 | 8 | 4 |
| 7 | 1, 2, 3, 4 | 8 | 4 |
| 8 | 5, 6, 7 | F1-F4 | 5 |
| F1-F4 | 8 | - | FINAL |

### Agent Dispatch Summary
- **1**: **1** — T1 → `quick`
- **2**: **2** — T2-T3 → `quick`
- **3**: **1** — T4 → `quick`
- **4**: **3** — T5-T7 → `quick`
- **5**: **1** — T8 → `quick`
- **FINAL**: **4** — F1 → `oracle`, F2 → `unspecified-high`, F3 → `unspecified-high`, F4 → `deep`

---

## TODOs

- [x] 1. Module scaffolding: constants, shared helpers, error handling wrapper

  **What to do**:
  - Create `scripts/vba/03_产品分类表现报告.txt`
  - Add `Option Explicit` at top
  - Define constants for all sheet names, column headers, and category names
  - Copy shared helper functions from 02 module: `BuildHeaderMap`, `NormalizeText`, `ParseYYYYMMDD`, `LastUsedRow`, `LastUsedColumn`, `TryReadDate`
  - Add error handling wrapper pattern (Application settings save/restore, On Error GoTo)
  - Add entry sub stubs: `Public Sub STEP3产品分类表现报告()` and `Public Sub S3ProductReport()`
  - Define output column order constants for each of the 3 sheets

  **Constants to define**:
  ```vb
  Private Const TARGET_SHEET_NAME As String = "产品分类"
  Private Const NAV_SHEET_NAME As String = "净值数据"
  Private Const COL_TRUST_CODE As String = "信托计划代码"
  Private Const COL_NAV_DATE As String = "日期"
  Private Const COL_UNIT_NAV As String = "单位净值"
  Private Const COL_SEQ As String = "序号"
  Private Const COL_CATEGORY As String = "分类"
  Private Const COL_SERIES As String = "系列"
  Private Const COL_PRODUCT_NAME As String = "产品名称"
  ' ... Step 2 output columns (already on 产品分类)
  Private Const COL_PREV_OPEN As String = "上一开放日"
  Private Const COL_PREV_PREV_OPEN As String = "上上一开放日"
  Private Const COL_NEXT_OPEN As String = "下一开放日"
  Private Const COL_BASELINE_DATE As String = "基准日期"
  Private Const COL_BASELINE_NAV As String = "基准日期净值"
  Private Const COL_PREV_NAV As String = "上一开放日净值"
  Private Const COL_PREV_PREV_NAV As String = "上上一开放日净值"
  Private Const COL_INTERVAL As String = "实际间隔"
  Private Const COL_PREV_INTERVAL As String = "上次开放实际间隔"
  Private Const COL_ELAPSED As String = "运作时间"
  Private Const COL_BENCHMARK_RATE As String = "基准收益率"
  ' Output-specific constants
  Private Const COL_CURRENT_PERIOD_ANNUAL As String = "当前周期年化"
  Private Const COL_PREV_PERIOD_ANNUAL As String = "上一周期年化"
  Private Const COL_7DAY_ANNUAL As String = "7日年化"
  Private Const COL_28DAY_ANNUAL As String = "28日年化"
  Private Const COL_INCEPTION_ANNUAL As String = "成立以来年化"
  ```

  **Must NOT do**:
  - Do not modify 01 or 02 .txt files
  - Do not duplicate code verbatim from 01 module (only 02 shared helpers)
  - Do not add helper functions not needed for this module

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Well-defined scaffolding task with existing patterns to copy
  - **Skills**: None
  - **Skills Evaluated but Omitted**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4)
  - **Blocks**: Tasks 2, 3, 4, 5, 6, 7, 8
  - **Blocked By**: None (can start immediately)

  **References**:
  - `scripts/vba/02_测算开放日.txt:1-65` — Module structure pattern (Option Explicit, constants, helper functions)
  - `scripts/vba/02_测算开放日.txt:569-587` — `BuildHeaderMap` implementation to copy
  - `scripts/vba/02_测算开放日.txt:611-623` — `NormalizeText` implementation to copy
  - `scripts/vba/02_测算开放日.txt:539-567` — `ParseYYYYMMDD` implementation to copy
  - `scripts/vba/02_测算开放日.txt:589-609` — `LastUsedRow` and `LastUsedColumn` to copy
  - `scripts/vba/01_增量导入净值数据.txt:418-438` — `TryReadDate` (alternative date parsing) to copy
  - `scripts/vba/02_测算开放日.txt:52-65` — Error handling wrapper pattern (Application settings save/restore)

  **Acceptance Criteria**:
  - [ ] File exists: `scripts/vba/03_产品分类表现报告.txt`
  - [ ] Contains `Option Explicit` as first non-comment line
  - [ ] All constants defined (verified by grep for each constant name)
  - [ ] Entry sub signatures: `Public Sub STEP3产品分类表现报告()` and `Public Sub S3ProductReport()`
  - [ ] All 5 shared helper functions present: BuildHeaderMap, NormalizeText, ParseYYYYMMDD, LastUsedRow, LastUsedColumn
  - [ ] Error handling wrapper pattern present

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify module scaffolding is complete and syntactically sound
    Tool: Bash (Python)
    Preconditions: File created at scripts/vba/03_产品分类表现报告.txt
    Steps:
      1. Read file content: uv run python -c "with open('scripts/vba/03_产品分类表现报告.txt','r',encoding='utf-8') as f: c=f.read()"
      2. Assert 'Option Explicit' in first 5 lines
      3. Assert 'Public Sub STEP3' or 'Public Sub S3' present
      4. Assert each required constant name present (grep for COL_TRUST_CODE, COL_CATEGORY, etc.)
      5. Assert each helper function signature present (BuildHeaderMap, NormalizeText, ParseYYYYMMDD, LastUsedRow, LastUsedColumn)
    Expected Result: All assertions pass
    Failure Indicators: Missing Option Explicit, missing entry sub, missing any constant or helper function
    Evidence: .omo/evidence/task-1-scaffolding.txt
  ```

  **Evidence to Capture**:
  - [ ] Evidence file: task-1-scaffolding.txt — contains grep results for all constants and functions

  **Commit**: NO (grouped with Task 8)

- [x] 2. NAV lookup infrastructure: date-based Dictionary for arbitrary date queries

  **What to do**:
  - Copy and adapt `BuildNAVDateLookup` from 02 module (lines 443-502)
  - Copy and adapt `LookupNAV` from 02 module (lines 508-533)
  - Ensure the lookup supports querying by ANY date (not just baseline date) — this is critical for 7/28日 queries
  - The existing implementation already builds a Dictionary(信托计划代码 → Dictionary(yyyymmdd → 单位净值)), which supports arbitrary dates
  - Add a helper `GetNAVDate` that looks up a specific date string (yyyy-mm-dd or Date) from the inner Dictionary
  - Document that duplicate (code, date) rows keep first occurrence (existing behavior)
  - Handle edge cases: empty trust code, non-date values, missing NAV sheet

  **Must NOT do**:
  - Do not modify the existing BuildNAVDateLookup in 02 module
  - Do not add nearest-date fallback logic (exact date match only)
  - Do not change the "first wins" duplicate policy

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Copy-adapt from existing, well-understood pattern
  - **Skills**: None
  - **Skills Evaluated but Omitted**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 3, 4) — all depend on Task 1
  - **Blocks**: Tasks 5, 6, 7
  - **Blocked By**: Task 1

  **References**:
  - `scripts/vba/02_测算开放日.txt:443-502` — `BuildNAVDateLookup` implementation (exact pattern to copy)
  - `scripts/vba/02_测算开放日.txt:508-533` — `LookupNAV` implementation (exact pattern to copy)
  - `scripts/vba/02_测算开放日.txt:539-567` — `ParseYYYYMMDD` (already copied in Task 1, used by NAV lookup)

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Function BuildNAVDateLookup() As Object`
  - [ ] Function signature: `Private Function LookupNAV(ByVal navLookup As Object, ByVal trustCodeValue As Variant, ByVal targetDate As Variant) As Variant`
  - [ ] Lookup key structure: outer dict key = 信托计划代码, inner dict key = yyyymmdd
  - [ ] Correct handling of duplicate rows (first wins)
  - [ ] Correct handling of empty/error inputs (returns Empty)

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify NAV lookup function signatures and logic are correct
    Tool: Bash (Python regex)
    Preconditions: Task 1 completed (helpers available in file)
    Steps:
      1. Read file: uv run python -c "with open('scripts/vba/03_产品分类表现报告.txt','r',encoding='utf-8') as f: c=f.read()"
      2. Assert BuildNAVDateLookup signature with 'As Object' return type
      3. Assert LookupNAV signature with proper parameters (navLookup, trustCodeValue, targetDate)
      4. Assert inner Dictionary uses Format$(parsedDate, 'yyyymmdd') as date key
      5. Assert 'first wins' logic: If Not outerDict(codeKey).Exists(dateKey) Then ... Add
      6. Assert LookupNAV returns Empty for missing code or date (not raises error)
    Expected Result: All pattern assertions pass
    Failure Indicators: Missing return type, wrong date key format, missing duplicate check, error-raising on miss
    Evidence: .omo/evidence/task-2-nav-lookup.txt
  ```

  **Evidence to Capture**:
  - [ ] Evidence file: task-2-nav-lookup.txt — grep matches for function signatures and key logic patterns

  **Commit**: NO (grouped with Task 8)

- [x] 3. Output workbook scaffolding: create .xlsx, define sheet structure, write headers

  **What to do**:
  - Write function `CreateOutputWorkbook` that:
    - Determines baseline date: reads the 基准日期 from 产品分类 (same global value computed by Step 2 for all rows; take the first valid value from any data row)
    - Constructs output filename: `Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现"` (note: no .xlsx in VBA variable — added at SaveAs)
    - Creates a new `Workbook` object via `Workbooks.Add`
    - Renames the default sheets to "稳享长期限", "直销", "交行代销" (delete any extra default sheets)
    - Writes header row (row 1) for each sheet with the exact column list
  - Write helper `WriteSheetHeaders` that takes a worksheet and an array of headers
  - Each sheet's headers (in order):

  **稳享长期限**: 序号 | 信托计划代码 | 系列 | 产品名称 | 上一开放日 | 上一开放日净值 | 基准日期 | 基准日期净值 | 下一开放日 | 实际间隔 | 运作时间 | 当前周期年化

  **直销**: 序号 | 信托计划代码 | 系列 | 产品名称 | 上一开放日 | 上一开放日净值 | 基准日期 | 基准日期净值 | 下一开放日 | 基准收益率 | 7日年化 | 28日年化 | 成立以来年化

  **交行代销**: 序号 | 信托计划代码 | 系列 | 产品名称 | 上一开放日 | 上一开放日净值 | 基准日期 | 基准日期净值 | 下一开放日 | 上上一开放日 | 上上一开放日净值 | 运作时间 | 上一周期年化 | 当前周期年化

  **Must NOT do**:
  - Do not write any data rows yet (Task 4 handles that)
  - Do not add extra default sheets (delete them)
  - Do not save the workbook prematurely

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Straightforward workbook creation and header writing
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 2, 4) — all depend on Task 1
  - **Blocks**: Tasks 4, 5, 6, 7, 8
  - **Blocked By**: Task 1

  **References**:
  - `scripts/vba/02_测算开放日.txt:358-389` — `EnsureOutputColumns` pattern for writing column headers
  - `scripts/vba/02_测算开放日.txt:318-334` — `ClearExistingOutputColumns` pattern (reverse iteration, though not needed here)
  - `上层产品净值数据库.xlsm` — Source for baseline date determination (产品分类 sheet, 基准日期 column)

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Function CreateOutputWorkbook(ByRef baselineDate As Date) As Workbook`
  - [ ] Filename constructed as `yyyyMMdd-上层产品分类表现` (baseline date from 产品分类)
  - [ ] Three sheets created with exact names: "稳享长期限", "直销", "交行代销"
  - [ ] Each sheet has correct headers in row 1
  - [ ] No extra default sheets remain
  - [ ] Workbook not saved yet (save deferred to Task 8)

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify workbook creation logic is correctly coded
    Tool: Bash (Python regex)
    Preconditions: Task 1 completed
    Steps:
      1. Read file and grep for '稳享长期限', '直销', '交行代销' — assert exactly 3 occurrences each in sheet creation context
      2. Assert header arrays contain correct column names for each category
      3. Assert filename construction uses Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现"
      4. Assert default sheet deletion logic exists (Delete or Application.DisplayAlerts = False)
     Expected Result: All pattern assertions pass
     Failure Indicators: Wrong sheet names, wrong header order, default sheets not cleaned up
    Evidence: .omo/evidence/task-3-workbook-scaffolding.txt
  ```

  **Evidence to Capture**:
  - [ ] Evidence file: task-3-workbook-scaffolding.txt — grep results for sheet names, header arrays, SaveAs

  **Commit**: NO (grouped with Task 8)

- [x] 4. Common field writer: read 产品分类, filter by category, write common fields to output sheets

  **What to do**:
  - Write function `WriteCommonFields` that:
    - Reads 产品分类 sheet from `ThisWorkbook`
    - Builds header map via BuildHeaderMap
    - Iterates all data rows (row 2 to last)
    - For each row, reads the 分类 column value
    - Matches against the 3 category names (normalized trim compare)
    - Writes to the corresponding output sheet using a "next row" counter per sheet
    - Copies: 序号, 信托计划代码, 系列, 产品名称, 上一开放日, 上一开放日净值, 基准日期, 基准日期净值, 下一开放日
  - Maps source columns to output columns (both found by header name)
  - Products outside the 3 categories are silently skipped
  - Blank rows in 产品分类 are skipped
  - Date columns get `NumberFormat = "yyyy-mm-dd"`
  - NAV columns keep source format (number)

  **Must NOT do**:
  - Do not compute any annualized returns yet (Tasks 5/6/7 handle that)
  - Do not write category-specific fields (Task 5/6/7 handle that)
  - Do not modify source workbook
  - Do not skip products with missing data (keep all rows per user decision)

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Row-by-row data copy with category filtering
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: NO (depends on Task 3 completing first)
  - **Parallel Group**: Wave 3 (sequential — blocked by Task 3)
  - **Blocks**: Tasks 5, 6, 7, 8
  - **Blocked By**: Tasks 1, 3

  **References**:
  - `scripts/vba/02_测算开放日.txt:125-219` — Row iteration pattern in CalculateOpenDaysCore (For r = 2 To lastRow)
  - `scripts/vba/02_测算开放日.txt:569-587` — BuildHeaderMap for header-based column lookup
  - `scripts/vba/01_增量导入净值数据.txt:452-464` — NormalizeText for category matching
  - 产品分类 sheet data analyzed: categories are exactly "稳享长期限", "直销", "交行代销" (no spaces/variants needed)

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Sub WriteCommonFields(ByVal wbOutput As Workbook, ByRef rowCounters As Object)`
  - [ ] Reads 产品分类 from ThisWorkbook (not from a new workbook)
  - [ ] Category matching via normalized trim comparison on 分类 column
  - [ ] 9 common fields written to correct output sheet
  - [ ] Date columns formatted as "yyyy-mm-dd"
  - [ ] Row counters tracked per sheet for subsequent tasks to use
  - [ ] Products outside 3 categories skipped silently

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify common field writer logic matches specification
    Tool: Bash (Python regex + code review)
    Preconditions: Tasks 1, 3 completed
    Steps:
      1. Read file and grep for '产品分类' — assert ThisWorkbook.Worksheets("产品分类") pattern
      2. Assert category comparison logic exists (If/Select Case on normalized 分类 value)
      3. Assert all 9 common columns are written (grep for each: 序号, 信托计划代码, 系列, 产品名称, 上一开放日, 上一开放日净值, 基准日期, 基准日期净值, 下一开放日)
      4. Assert NumberFormat = "yyyy-mm-dd" applied to date columns
      5. Assert row counter tracking (per-sheet counter increment)
    Expected Result: All pattern assertions pass
    Failure Indicators: Wrong source sheet reference, missing common column, wrong date format, missing category filter
    Evidence: .omo/evidence/task-4-common-writer.txt
  ```

  **Evidence to Capture**:
  - [ ] Evidence file: task-4-common-writer.txt — grep results for category matching and column writing

  **Commit**: NO (grouped with Task 8)

- [x] 5. 稳享长期限 computation: 当前周期年化

  **What to do**:
  - Write function `ComputeStableLongTerm` that iterates the "稳享长期限" output sheet
  - For each data row (row 2 to lastRow), read:
    - `上一开放日净值` — may be Empty
    - `基准日期净值` — may be Empty
    - `运作时间` — may be Empty
  - If ALL three values present AND 运作时间 > 0:
    - Compute: `当前周期年化 = (基准日期净值 / 上一开放日净值 - 1) * (365 / 运作时间)`
    - Write to `当前周期年化` column, format as `"0.00%"`
  - If ANY value missing or 运作时间 = 0: leave cell blank
  - Check for 上一开放日净值 = 0 to avoid division by zero

  **Must NOT do**:
  - Do NOT use +1 day-count rule
  - Do NOT recompute 实际间隔 or 运作时间 (from Step 2)
  - Do NOT add 上上一开放日 or 上一周期年化 for this category

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Single formula computation with well-defined inputs
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 4 (with Tasks 6, 7)
  - **Blocks**: Task 8
  - **Blocked By**: Tasks 1, 2, 3, 4

  **References**:
  - Formula: `当前周期年化 = (基准日期净值 / 上一开放日净值 - 1) × (365 / 运作时间)` — no +1 rule
  - `scripts/vba/02_测算开放日.txt:184` — 运作时间 pattern: `CLng(baselineDate - prevDate)`
  - Product data: 63 products in 稳享长期限, all have 运作时间 from Step 2

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Sub ComputeStableLongTerm(ByVal wbOutput As Workbook)`
  - [ ] Formula: `(baselineNav / prevNav - 1) * (365 / elapsedDays)`
  - [ ] Division-by-zero guard: prevNav <> 0 AND elapsedDays > 0
  - [ ] Missing data → blank cell
  - [ ] Result formatted as "0.00%"

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify 稳享长期限 formula correctness
    Tool: Bash (Python)
    Preconditions: Tasks 1-4 completed
    Steps:
      1. Search for annualized return formula in context of '稳享长期限' sheet
      2. Assert: (baselineNav / prevNav - 1) * (365 / elapsed)
      3. Assert NO "+ 1" in denominator
      4. Assert guard: If prevNav <> 0 And elapsed > 0
      5. Assert NumberFormat = "0.00%"
    Expected Result: Formula matches spec, guards present
    Failure Indicators: +1 present, missing guard, different formula
    Evidence: .omo/evidence/task-5-stable-formula.txt
  ```

  **Evidence to Capture**:
  - [ ] task-5-stable-formula.txt — grep for formula and guard pattern

  **Commit**: NO (grouped with Task 8)

- [x] 6. 直销 computation: 基准收益率, 7日/28日年化, 成立以来年化留空

  **What to do**:
  - Write function `ComputeDirectSales` iterating the "直销" output sheet
  - For each row:
    a. **基准收益率**: Copy from 产品分类 source data (read in same pass or from pre-read)
    b. **7日年化**:
       - `targetDate = DateAdd("d", -7, baselineDate)`
       - NAV via `LookupNAV(navLookup, trustCode, targetDate)`
       - If found AND not zero: `7日年化 = (baselineNav / targetNav - 1) * (365 / 7)`
       - Not found or zero: leave blank
    c. **28日年化**: Same pattern with `DateAdd("d", -28, baselineDate)` and `365 / 28`
    d. **成立以来年化**: Always leave blank (per user decision)
  - Apply NumberFormat = "0.00%" to 7日/28日 columns
  - **IMPORTANT**: Use `DateAdd` for exact calendar day offset — NO nearest-date fallback loop

  **Must NOT do**:
  - Do NOT use nearest-date fallback (exact date only)
  - Do NOT compute 成立以来年化
  - Do NOT use +1 rule
  - Do NOT apply AGENTS.md young-product inception rule

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: NAV date lookup + formula computation
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 4 (with Tasks 5, 7)
  - **Blocks**: Task 8
  - **Blocked By**: Tasks 1, 2, 3, 4

  **References**:
  - Formula: `7日年化 = (基准日期净值 / 前7日历天净值 - 1) × (365 / 7)`
  - Formula: `28日年化 = (基准日期净值 / 前28日历天净值 - 1) × (365 / 28)`
  - VBA `DateAdd("d", -7, baselineDate)` for exact calendar offset
  - `scripts/vba/02_测算开放日.txt:508-533` — LookupNAV call pattern
  - Product data: 10 products in 直销

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Sub ComputeDirectSales(ByVal wbOutput As Workbook, ByVal navLookup As Object)`
  - [ ] 基准收益率 copied from source data
  - [ ] 7日: `DateAdd("d", -7)`, lookup, formula `(baselineNav / targetNav - 1) * (365 / 7)`
  - [ ] 28日: `DateAdd("d", -28)`, lookup, formula `(baselineNav / targetNav - 1) * (365 / 28)`
  - [ ] Exact date match only — no fallback loop
  - [ ] Division-by-zero guards
  - [ ] 成立以来年化 column always blank

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify 直销 7日/28日 date offset is exact (no fallback)
    Tool: Bash (Python)
    Preconditions: Tasks 1-4 completed
    Steps:
      1. Search for DateAdd("d", -7 near 7日年化 context
      2. Assert DateAdd result passed directly to LookupNAV without While/Wend loop
      3. Assert no nearest-date iteration pattern
      4. Repeat for DateAdd("d", -28 near 28日年化 context
    Expected Result: Exact DateAdd offsets, no fallback loops
    Failure Indicators: While/Wend loop, looping through dates, nearest-match logic
    Evidence: .omo/evidence/task-6-exact-date.txt

  Scenario: Verify 成立以来年化 left blank
    Tool: Bash (Python)
    Preconditions: Tasks 1-4 completed
    Steps:
      1. Search for '成立以来年化' in file
      2. Assert NO computation logic (no formula, no value assignment beyond Empty)
      3. Assert column header written but data cells untouched
    Expected Result: Column exists but no data written
    Failure Indicators: Formula or value assignment for 成立以来年化
    Evidence: .omo/evidence/task-6-inception-blank.txt
  ```

  **Evidence to Capture**:
  - [ ] task-6-exact-date.txt — grep for DateAdd usage without fallback
  - [ ] task-6-inception-blank.txt — grep confirming no 成立以来年化 computation

  **Commit**: NO (grouped with Task 8)

- [x] 7. 交行代销 computation: 上上一开放日数据 + 两个年化

  **What to do**:
  - Write function `ComputeBankAgent` iterating the "交行代销" output sheet
  - For each row, FIRST write the extra data fields (not yet written by Task 4):
    - Write `上上一开放日` (date format "yyyy-mm-dd"), `上上一开放日净值`, `运作时间` from source 产品分类
  - THEN compute:
    a. **当前周期年化**: `(基准日期净值 / 上一开放日净值 - 1) * (365 / 运作时间)` — same as 稳享长期限
    b. **上一周期年化**: `(上一开放日净值 / 上上一开放日净值 - 1) * (365 / 上次开放实际间隔)`
  - Guards: check each field non-empty, non-zero for denominator values
  - Missing data → leave computed cell blank
  - Apply NumberFormat = "0.00%" to annualized columns

  **Must NOT do**:
  - Do NOT use +1 rule
  - Do NOT compute fields belonging to other categories

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Data copy + two formula computations
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 4 (with Tasks 5, 6)
  - **Blocks**: Task 8
  - **Blocked By**: Tasks 1, 2, 3, 4

  **References**:
  - 当前周期年化: `(bnav / pnav - 1) * (365 / elapsed)`
  - 上一周期年化: `(pnav / ppnav - 1) * (365 / prevInterval)`
  - `scripts/vba/02_测算开放日.txt:200-215` — 上上一开放日 handling pattern
  - Product data: 110 products, 77 have 上一开放日净值, 63 have 上上一开放日净值

  **Acceptance Criteria**:
  - [ ] Function signature: `Private Sub ComputeBankAgent(ByVal wbOutput As Workbook)`
  - [ ] 上上一开放日, 上上一开放日净值, 运作时间 written from source
  - [ ] 当前周期年化: correct formula with guard
  - [ ] 上一周期年化: correct formula with guard (uses 上次开放实际间隔 not 运作时间)
  - [ ] Both formatted "0.00%"

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify 交行代销 previous period formula uses correct denominator
    Tool: Bash (Python)
    Preconditions: Tasks 1-4 completed
    Steps:
      1. Search for 上一周期年化 formula in 交行代销 context
      2. Assert: (prevNav / prevPrevNav - 1) * (365 / prevInterval)
      3. Assert denominator is 上次开放实际间隔 (NOT 运作时间)
      4. Assert guard: prevPrevNav <> 0 And prevInterval > 0
    Expected Result: Formula uses 上次开放实际间隔 as denominator
    Failure Indicators: Using 运作时间 instead of 上次开放实际间隔
    Evidence: .omo/evidence/task-7-prev-formula.txt
  ```

  **Evidence to Capture**:
  - [ ] task-7-prev-formula.txt — grep for prev period formula with correct denominator

  **Commit**: NO (grouped with Task 8)

- [x] 8. Main orchestration subroutine + formatting + save

  **What to do**:
  - Implement the `CalculateClassReportCore` private sub (called by public entry points)
  - Wire everything together:
    1. Save/restore Application settings (ScreenUpdating, Calculation, EnableEvents)
    2. Get baseline date from 产品分类
    3. Build NAV lookup: `Set navLookup = BuildNAVDateLookup()`
    4. Create output workbook: `Set wbOut = CreateOutputWorkbook(baselineDate)`
    5. Write common fields: `WriteCommonFields wbOut, rowCounters` (requires reading 产品分类)
    6. Compute category fields (parallel-safe — each operates on separate sheet):
       - `ComputeStableLongTerm wbOut`
       - `ComputeDirectSales wbOut, navLookup`
       - `ComputeBankAgent wbOut`
    7. Auto-fit columns for all sheets
    8. Save workbook: `wbOut.SaveAs outputPath, xlOpenXMLWorkbook` (51 = .xlsx)
    9. Close output workbook
    10. Restore Application settings
    11. MsgBox with summary: products per sheet, baseline date
  - Public entry points: `Public Sub STEP3产品分类表现报告()` and `Public Sub S3ProductReport()`
  - Output path: `ThisWorkbook.Path & Application.PathSeparator & Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现.xlsx"`

  **Must NOT do**:
  - Do NOT modify ThisWorkbook
  - Do NOT leave output workbook open
  - Do NOT save as .xlsm (must be .xlsx — no macros in output)

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Wiring existing functions, straightforward orchestration
  - **Skills**: None

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 5 (sequential — depends on all previous tasks)
  - **Blocks**: F1-F4 (Final Verification)
  - **Blocked By**: Tasks 5, 6, 7

  **References**:
  - `scripts/vba/02_测算开放日.txt:52-238` — Full orchestration pattern (settings, build lookups, iterate, restore, MsgBox)
  - `scripts/vba/01_增量导入净值数据.txt:15-119` — Error handling wrapper + MsgBox summary pattern
  - VBA `xlOpenXMLWorkbook` = 51 (constant for .xlsx format)

  **Acceptance Criteria**:
  - [ ] Public entry subs: `STEP3产品分类表现报告` and `S3ProductReport`
  - [ ] Application settings saved/restored
  - [ ] Calls all 4 component functions in correct order
  - [ ] Saves as .xlsx (xlOpenXMLWorkbook or 51)
  - [ ] Output path: `{ThisWorkbook.Path}\{yyyymmdd}-上层产品分类表现.xlsx`
  - [ ] MsgBox shows summary with product counts per sheet
  - [ ] Output workbook closed after save
  - [ ] Error handler with CleanFail label for settings restoration

  **QA Scenarios (MANDATORY)**:

  ```
  Scenario: Verify main orchestration wires all components correctly
    Tool: Bash (Python)
    Preconditions: All tasks 1-7 completed
    Steps:
      1. Read complete .txt file
      2. Assert STEP3 entry sub calls into a Core function
      3. Assert BuildNAVDateLookup is called (NAV lookup built)
      4. Assert CreateOutputWorkbook is called (workbook created)
      5. Assert WriteCommonFields is called (data written)
      6. Assert ComputeStableLongTerm, ComputeDirectSales, ComputeBankAgent all called
      7. Assert SaveAs uses xlOpenXMLWorkbook (51) not xlOpenXMLWorkbookMacroEnabled (52)
      8. Assert output path: ThisWorkbook.Path & "\" & Format$(baselineDate, "yyyymmdd") & "-上层产品分类表现.xlsx"
    Expected Result: All functions wired, .xlsx format, correct path
    Failure Indicators: Missing function call, .xlsm format, hardcoded path
    Evidence: .omo/evidence/task-8-orchestration.txt
  ```

  **Evidence to Capture**:
  - [ ] task-8-orchestration.txt — grep for all function calls and SaveAs

  **Commit**: YES
  - Message: `feat(vba): add step 3 product classification report module`
  - Files: `scripts/vba/03_产品分类表现报告.txt`
  - Pre-commit: verify file exists and has correct encoding

---

## Final Verification Wave (MANDATORY — after ALL implementation tasks)

> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.
>
> **Do NOT auto-proceed after verification. Wait for user's explicit approval before marking work complete.**
> **Never mark F1-F4 as checked before getting user's okay.** Rejection or user feedback -> fix -> re-run -> present again -> wait for okay.

- [ ] F1. **Plan Compliance Audit** — `oracle`
  Read the plan end-to-end. For each "Must Have": verify implementation exists (read the .txt file, check constants, check logic). For each "Must NOT Have": search .txt for forbidden patterns — reject with file:line if found. Check evidence files exist in .omo/evidence/. Compare deliverables against plan.
  **Acceptance Criteria**:
  - [ ] All 13 Must Have items verified present in .txt via grep
  - [ ] All 11 Must NOT Have items verified absent in .txt
  - [ ] All 9 task evidence files present in .omo/evidence/
  - [ ] VERDICT must be APPROVE (not REJECT)
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [ ] F2. **Code Quality Review** — `unspecified-high`
  Review the .txt file for: VBA compilation issues (undeclared variables, missing End Sub/Function), correct error handling pattern (On Error GoTo), Chinese comment consistency, code duplication vs existing modules (01/02), correct formula implementation. Check AI slop: excessive comments, over-abstraction, unused procedures.
  **Acceptance Criteria**:
  - [ ] Every Sub/Function has matching End Sub/End Function
  - [ ] Every variable declared (Option Explicit compatibility verified manually)
  - [ ] Error handler pattern: On Error GoTo CleanFail present with label
  - [ ] No `as any`, `@ts-ignore`, or Excel-specific hacks
  - [ ] Formula `(x / y - 1) * (365 / z)` verified for each computation task
  - [ ] No unused procedures (every Sub/Function called by orchestration)
  Output: `Syntax [PASS/FAIL] | Pattern [PASS/FAIL] | Formula [N correct/N total] | VERDICT`

- [ ] F3. **Real Manual QA — .xlsx Verification** — `unspecified-high`
  Use Python/openpyxl to verify the generated .xlsx. Check: file exists at correct path, exactly 3 sheets with correct names, each sheet has correct column headers in correct order, row counts match expected, sample formula results computed manually, edge cases (missing NAV, missing open date) produce blanks not errors.
  **Acceptance Criteria**:
  - [ ] File exists at: `{db dir}\yyyyMMdd-上层产品分类表现.xlsx`
  - [ ] Workbook has exactly 3 sheets: "稳享长期限", "直销", "交行代销"
  - [ ] 稳享长期限: 12 columns, headers match spec order, rows ≤ 63
  - [ ] 直销: 13 columns, headers match spec order, rows ≤ 10
  - [ ] 交行代销: 14 columns, headers match spec order, rows ≤ 110
  - [ ] At least 1 sample row per sheet with annualized return verified by manual calculation
  - [ ] Rows with missing input data have blank (None) computed cells, not #DIV/0! or errors
  - [ ] Date columns formatted as yyyy-mm-dd (not serial numbers)
  Output: `Scenarios [N/N pass] | Integration [N/N] | Edge Cases [N tested] | VERDICT`

- [ ] F4. **Scope Fidelity Check** — `deep`
  For each task: read "What to do", read actual .txt content. Verify 1:1 — everything in spec was implemented (no missing), nothing beyond spec was implemented (no creep). Check "Must NOT do" compliance. Flag unaccounted procedures or constants.
  **Acceptance Criteria**:
  - [ ] Task 1 (scaffolding): all 5 shared helpers + all constants present
  - [ ] Task 2 (NAV lookup): 2 functions present, exact date match only, first-wins duplicate
  - [ ] Task 3 (workbook): 3 sheets created, SaveAs .xlsx, no default sheets
  - [ ] Task 4 (common writer): 9 fields written, category filter correct, date format applied
  - [ ] Task 5 (稳享长期限): formula correct, no +1, guard present
  - [ ] Task 6 (直销): 7日/28日 via DateAdd, no fallback, 成立以来年化 blank
  - [ ] Task 7 (交行代销): both formulas correct, 上次开放实际间隔 used correctly
  - [ ] Task 8 (orchestration): all functions wired, Application settings restored, MsgBox summary
  - [ ] No unaccounted procedures, constants, or modules beyond plan scope
  - [ ] Zero cross-contamination with 01/02 modules
  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | VERDICT`

---

## Commit Strategy
- **1**: `feat(vba): add step 3 product classification report module` — `scripts/vba/03_产品分类表现报告.txt`

---

## Success Criteria

### Verification Commands
```powershell
# Verify VBA source exists and has expected structure
uv run python -c "
import re
with open('scripts/vba/03_产品分类表现报告.txt', 'r', encoding='utf-8') as f:
    content = f.read()
checks = {
    'Option Explicit': 'Option Explicit' in content,
    'STEP3 entry sub': 'Public Sub STEP3' in content or 'Public Sub S3' in content,
    'NAV lookup func': 'BuildNAVDateLookup' in content or 'LookupNAV' in content,
    'Annualized formula': '365' in content and '/ 运作时间' in content or '/ 运作时间' not in content,
    'Error handler': 'On Error GoTo' in content,
    'Sheet names': all(s in content for s in ['稳享长期限', '直销', '交行代销']),
}
for k, v in checks.items():
    print(f'  {k}: {\"PASS\" if v else \"FAIL\"}')"
```

### Final Checklist
- [ ] All "Must Have" present in .txt
- [ ] All "Must NOT Have" absent from .txt
- [ ] VBA syntax valid (Option Explicit, no undeclared vars)
- [ ] Formula implementations match confirmed specifications
- [ ] Python/openpyxl verification of .xlsx output passes
