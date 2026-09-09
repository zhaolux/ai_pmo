# Weekly Iteration Plan Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the current text-heavy weekly iteration sheet with a reconciled weekly summary and task-stage detail view driven by the master control ledger.

**Architecture:** Extract weekly-iteration calculations into a focused pure-Python module, keep workbook styling and integration in the existing V2 builder, and preserve human-entered review fields by stable keys before rebuilding the sheet. The summary is calculated from the generated stage-level detail so weekly counts always reconcile.

**Tech Stack:** Python 3.12, `openpyxl`, `pytest`, LibreOffice headless rendering, existing `.artifact-work/build_ai_pmo_v2.py` workbook generator.

---

## File Structure

- Create `.artifact-work/weekly_iteration.py`: pure functions for week extraction, stage assignment, exception classification, weekly aggregation and stable review keys.
- Create `tests/test_weekly_iteration.py`: unit tests for week assignment, completion-rate rules, future-week behavior, missing dates and paused tasks.
- Modify `.artifact-work/build_ai_pmo_v2.py`: preserve existing review inputs, call the new calculation module, rebuild the optimized sheet, apply formatting and update the operating guide.
- Create `.artifact-work/verify_weekly_iteration.py`: workbook-level reconciliation, formula-error and layout checks.

### Task 1: Build the weekly-iteration calculation model

**Files:**
- Create: `.artifact-work/weekly_iteration.py`
- Create: `tests/test_weekly_iteration.py`

- [ ] **Step 1: Write failing tests for week assignment and aggregation**

```python
from datetime import date

from weekly_iteration import StageItem, Week, assign_stage_items, summarize_week


def test_assigns_stage_by_planned_finish_date():
    weeks = [Week(1, date(2026, 9, 7), date(2026, 9, 11))]
    item = StageItem(
        task_id="T001", task_name="船舶查询", batch="第二批次",
        stage="后端", owner="常浩", planned_start=date(2026, 9, 7),
        planned_finish=date(2026, 9, 11), stage_status="已完成",
        overall_status="已完成", actual_finish=date(2026, 9, 10),
    )
    assigned, unscheduled = assign_stage_items(weeks, [item])
    assert assigned[1] == [item]
    assert unscheduled == []


def test_summary_reconciles_completed_delayed_and_paused():
    week = Week(1, date(2026, 9, 7), date(2026, 9, 11))
    items = [
        StageItem("T001", "A", "一批", "前端", "张三", date(2026, 9, 7), date(2026, 9, 11), "已完成", "已完成", date(2026, 9, 10)),
        StageItem("T002", "B", "一批", "后端", "李四", date(2026, 9, 7), date(2026, 9, 11), "进行中", "未完成", None),
        StageItem("T003", "C", "一批", "测试", "王五", date(2026, 9, 7), date(2026, 9, 11), "暂停", "暂停", None),
    ]
    result = summarize_week(week, items, snapshot_date=date(2026, 9, 15))
    assert result.planned_count == 3
    assert result.completed_count == 1
    assert result.completion_rate == 1 / 3
    assert result.delayed_count == 1
    assert result.paused_count == 1
```

- [ ] **Step 2: Run the tests and confirm they fail**

Run:

```bash
PYTHONPATH=.artifact-work pytest tests/test_weekly_iteration.py -v
```

Expected: collection fails because `weekly_iteration` does not exist.

- [ ] **Step 3: Implement the data model and core calculations**

```python
from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True)
class Week:
    number: int
    start: date
    end: date


@dataclass(frozen=True)
class StageItem:
    task_id: str
    task_name: str
    batch: str | None
    stage: str
    owner: str | None
    planned_start: date | None
    planned_finish: date | None
    stage_status: str | None
    overall_status: str | None
    actual_finish: date | None


@dataclass(frozen=True)
class WeekSummary:
    planned_count: int
    completed_count: int
    completion_rate: float | None
    delayed_count: int
    paused_count: int
    state: str


def assign_stage_items(weeks, items):
    assigned = {week.number: [] for week in weeks}
    unscheduled = []
    for item in items:
        if item.planned_finish is None:
            unscheduled.append(item)
            continue
        match = next((week for week in weeks if week.start <= item.planned_finish <= week.end), None)
        if match is None:
            unscheduled.append(item)
        else:
            assigned[match.number].append(item)
    return assigned, unscheduled


def summarize_week(week, items, snapshot_date):
    planned = len(items)
    completed = sum(item.stage_status == "已完成" for item in items)
    paused = sum(item.overall_status in {"暂停", "挂起"} or item.stage_status in {"暂停", "挂起"} for item in items)
    delayed = sum(
        item.stage_status != "已完成"
        and item.overall_status not in {"暂停", "挂起"}
        and week.end < snapshot_date
        for item in items
    )
    state = "未开始" if week.start > snapshot_date else ("进行中" if week.end >= snapshot_date else "已结束")
    rate = None if state == "未开始" or planned == 0 else completed / planned
    return WeekSummary(planned, completed, rate, delayed, paused, state)
```

- [ ] **Step 4: Add tests for missing dates, future weeks and missing actual completion dates**

```python
def test_missing_planned_finish_is_unscheduled_and_excluded():
    item = StageItem("T004", "D", None, "测试", None, None, None, "未开始", "未开始", None)
    assigned, unscheduled = assign_stage_items([Week(1, date(2026, 9, 7), date(2026, 9, 11))], [item])
    assert assigned[1] == []
    assert unscheduled == [item]


def test_future_week_has_no_early_completion_rating():
    week = Week(2, date(2026, 9, 14), date(2026, 9, 18))
    result = summarize_week(week, [], snapshot_date=date(2026, 9, 9))
    assert result.state == "未开始"
    assert result.completion_rate is None
```

- [ ] **Step 5: Run tests and commit**

Run:

```bash
PYTHONPATH=.artifact-work pytest tests/test_weekly_iteration.py -v
```

Expected: all tests pass.

Commit:

```bash
git add .artifact-work/weekly_iteration.py tests/test_weekly_iteration.py
git commit -m "feat: add weekly iteration calculation model"
```

### Task 2: Preserve review inputs and extract source data

**Files:**
- Modify: `.artifact-work/build_ai_pmo_v2.py`
- Modify: `tests/test_weekly_iteration.py`

- [ ] **Step 1: Add a failing test for stable review keys**

```python
from weekly_iteration import detail_review_key, week_review_key


def test_review_keys_are_stable():
    assert week_review_key(12) == "W012"
    assert detail_review_key(12, "T134", "后端") == "W012|T134|后端"
```

- [ ] **Step 2: Implement stable keys**

```python
def week_review_key(week_number):
    return f"W{int(week_number):03d}"


def detail_review_key(week_number, task_id, stage):
    return f"{week_review_key(week_number)}|{task_id}|{stage}"
```

- [ ] **Step 3: Capture existing manual fields before clearing the sheet**

Add to `.artifact-work/build_ai_pmo_v2.py` before rebuilding `周迭代计划`:

```python
existing_week_reviews = {}
existing_detail_reviews = {}
if "周迭代计划" in wb.sheetnames:
    old_weekly = wb["周迭代计划"]
    for row in range(5, old_weekly.max_row + 1):
        key = old_weekly.cell(row, 13).value
        if isinstance(key, str) and key.startswith("W"):
            existing_week_reviews[key] = (
                old_weekly.cell(row, 4).value,
                old_weekly.cell(row, 12).value,
            )
        detail_key = old_weekly.cell(row, 15).value
        if isinstance(detail_key, str) and "|" in detail_key:
            existing_detail_reviews[detail_key] = old_weekly.cell(row, 14).value
```

- [ ] **Step 4: Extract `Week` and `StageItem` records from the current master ledger**

Use the four established master column groups:

```python
stage_columns = [
    ("业务", 12, 13, 14, 15),
    ("后端", 16, 17, 18, 19),
    ("前端", 20, 21, 22, 23),
    ("测试", 24, 25, 26, 27),
]
```

Create one `StageItem` only when the stage has at least one of owner, status, planned start or planned finish. Use master column `K` for actual completion date and columns `A`, `E/D/C`, `F`, `G` for task ID, preferred task name, batch and overall status.

- [ ] **Step 5: Run tests and commit**

Run:

```bash
PYTHONPATH=.artifact-work pytest tests/test_weekly_iteration.py -v
```

Expected: all tests pass.

Commit:

```bash
git add .artifact-work/weekly_iteration.py .artifact-work/build_ai_pmo_v2.py tests/test_weekly_iteration.py
git commit -m "feat: preserve weekly iteration review inputs"
```

### Task 3: Rebuild the weekly summary view

**Files:**
- Modify: `.artifact-work/build_ai_pmo_v2.py`

- [ ] **Step 1: Replace the existing nine-column weekly copy block**

Create a new sheet with title rows 1-2, summary header at row 4 and data beginning at row 5. Use these columns:

```python
summary_headers = [
    "迭代周", "开始日期", "结束日期", "周目标", "原周计划摘要",
    "计划完成阶段数", "实际完成阶段数", "承诺完成率", "延期数",
    "暂停数", "周状态", "复盘结论", "_week_key",
]
```

- [ ] **Step 2: Populate summary rows from the detail aggregation**

For each `Week`, write counts from `summarize_week`, restore column D and L from `existing_week_reviews`, and store `week_review_key` in hidden column M. Format dates as `yyyy-mm-dd`, completion rate as `0%`, and show `无计划任务` instead of `0%` when the denominator is zero.

- [ ] **Step 3: Preserve and consolidate the original weekly plan text**

Join nonblank source values from original columns D:G with stage labels:

```python
parts = []
for label, col in (("功能设计", 4), ("前端", 5), ("后端", 6), ("测试", 7)):
    value = latest_weekly.cell(source_row, col).value
    if value not in (None, "", "/"):
        parts.append(f"{label}：{value}")
original_summary = "\n".join(parts)
```

- [ ] **Step 4: Apply management formatting**

Use blue headers, full grid borders, yellow fill for columns D and L, wrapped text for D/E/L, conditional fills for delayed and paused counts, and a data bar on column H. Set freeze panes to `A5` until the detail header is added in Task 4.

- [ ] **Step 5: Build and inspect the summary**

Run:

```bash
PRESERVE_MASTER_ASSIGNMENTS=1 python .artifact-work/build_ai_pmo_v2.py
```

Expected: the dated V2 workbook contains one summary row per source iteration week and retains existing review values where stable keys match.

- [ ] **Step 6: Commit**

```bash
git add .artifact-work/build_ai_pmo_v2.py
git commit -m "feat: add weekly iteration summary"
```

### Task 4: Add the task-stage detail and exceptions

**Files:**
- Modify: `.artifact-work/build_ai_pmo_v2.py`
- Modify: `tests/test_weekly_iteration.py`

- [ ] **Step 1: Add failing tests for exception messages**

```python
from weekly_iteration import exception_messages


def test_exception_messages_are_explicit():
    item = StageItem("T005", "E", None, "后端", None, date(2026, 9, 12), date(2026, 9, 11), "已完成", "已完成", None)
    assert exception_messages(item) == ["负责人待补充", "日期异常", "完成日期待补充"]
```

- [ ] **Step 2: Implement exception classification**

```python
def exception_messages(item):
    messages = []
    if not item.owner:
        messages.append("负责人待补充")
    if item.planned_start and item.planned_finish and item.planned_start > item.planned_finish:
        messages.append("日期异常")
    if item.stage_status == "已完成" and item.actual_finish is None:
        messages.append("完成日期待补充")
    return messages
```

- [ ] **Step 3: Create the detail table beneath the summary**

Use a two-row separation, section title, then this header:

```python
detail_headers = [
    "迭代周", "任务ID", "任务名称", "批次", "专业阶段", "负责人",
    "计划开始", "计划完成", "阶段状态", "总体状态",
    "实际完成日期", "是否延期", "异常提示", "复盘说明", "_detail_key",
]
```

- [ ] **Step 4: Populate scheduled and unscheduled items**

Write scheduled items by week, stage and task ID. Append missing-date items under iteration value `待排期`. Restore review notes using `detail_review_key`; fill review cells yellow and hide column O.

- [ ] **Step 5: Add filtering, freezing and conditional formatting**

Set the detail AutoFilter to columns A:N, freeze at the row immediately below the detail header, highlight paused rows amber and delayed rows light red, and apply complete borders to all summary and detail cells.

- [ ] **Step 6: Run tests and commit**

Run:

```bash
PYTHONPATH=.artifact-work pytest tests/test_weekly_iteration.py -v
```

Expected: all tests pass.

Commit:

```bash
git add .artifact-work/weekly_iteration.py .artifact-work/build_ai_pmo_v2.py tests/test_weekly_iteration.py
git commit -m "feat: add weekly task-stage detail"
```

### Task 5: Update guidance and add workbook verification

**Files:**
- Modify: `.artifact-work/build_ai_pmo_v2.py`
- Create: `.artifact-work/verify_weekly_iteration.py`

- [ ] **Step 1: Update the operating guide text**

Replace the current weekly-iteration guide entry with:

```python
(
    "周迭代计划",
    "上方按周汇总承诺完成、延期和暂停情况；下方按总控任务的业务、前端、后端、测试阶段展开。周目标和复盘说明由项目经理填写，其余字段每次更新V2模板时根据总控台账重建。",
)
```

- [ ] **Step 2: Implement workbook-level verification**

```python
from openpyxl import load_workbook


def verify(path):
    wb = load_workbook(path, data_only=False)
    ws = wb["周迭代计划"]
    assert "周度汇总" in {cell.value for row in ws.iter_rows() for cell in row}
    assert "周任务明细" in {cell.value for row in ws.iter_rows() for cell in row}
    assert ws.auto_filter.ref is not None
    formulas = [cell.value for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")]
    assert not any("#REF!" in formula for formula in formulas)
```

Add this reconciliation inside `verify` after locating the summary and detail headers:

```python
from collections import defaultdict

detail_totals = defaultdict(lambda: {"planned": 0, "completed": 0, "delayed": 0, "paused": 0})
for row in range(detail_header + 1, ws.max_row + 1):
    week_number = ws.cell(row, 1).value
    if not isinstance(week_number, int):
        continue
    totals = detail_totals[week_number]
    totals["planned"] += 1
    totals["completed"] += ws.cell(row, 9).value == "已完成"
    totals["delayed"] += ws.cell(row, 12).value == "是"
    totals["paused"] += ws.cell(row, 10).value in {"暂停", "挂起"}

for row in range(summary_header + 1, detail_title_row - 2):
    week_number = ws.cell(row, 1).value
    if not isinstance(week_number, int):
        continue
    totals = detail_totals[week_number]
    assert ws.cell(row, 6).value == totals["planned"]
    assert ws.cell(row, 7).value == totals["completed"]
    assert ws.cell(row, 9).value == totals["delayed"]
    assert ws.cell(row, 10).value == totals["paused"]
```

- [ ] **Step 3: Run the complete build and verification**

Run:

```bash
PRESERVE_MASTER_ASSIGNMENTS=1 python .artifact-work/build_ai_pmo_v2.py
python .artifact-work/verify_weekly_iteration.py outputs/ai-pmo-v2/华能港口交易平台-AI-PMO-V2模板-20260909.xlsx
```

Expected: verifier exits successfully and prints weekly row count, detail row count, reconciliation result and zero formula errors.

- [ ] **Step 4: Render and visually inspect the changed sheet**

Run LibreOffice headless PDF conversion and inspect the pages containing `周度汇总` and `周任务明细`. Confirm readable widths, wrapped text, visible input fills, full borders, correct freeze position and no truncated dates.

- [ ] **Step 5: Commit final integration**

```bash
git add .artifact-work/build_ai_pmo_v2.py .artifact-work/verify_weekly_iteration.py
git commit -m "test: verify weekly iteration workbook output"
```
