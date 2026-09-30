"""按总控台账"迭代周"列生成周迭代计划（默认只读，显式执行才写工作簿）。

联动口径：明细区以任务ID为键，任务名称/批次/负责人/计划日期/状态全部用
INDEX/MATCH 按任务ID查项目总控台账——总控更新后明细自动跟随，插行不错位。
人工列（专业阶段、实际完成日期、复盘说明）按任务保留；汇总区周目标/复盘结论按周保留。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

WEEK_RE = re.compile(r"^(\d{4})-W(\d{2})$")
ITERATION_WEEK_HEADER = "迭代周"
CONTROL_SHEET = "项目总控台账"
ITERATION_SHEET = "周迭代计划"
SUMMARY_START = 5
SUMMARY_END = 28
DETAIL_START = 31
DETAIL_END = 610
WEEK_COLUMN = 32  # 项目总控台账 AF 列
DATE_FORMAT = "yyyy\\-mm\\-dd"

# 明细列 → 项目总控台账列字母（按表头语义：负责人=主责人H，计划开始/完成=I/J）
DETAIL_SOURCES = {"C": "E", "D": "F", "F": "H", "G": "I", "H": "J", "I": "M", "J": "G"}
DETAIL_MANUAL_COLUMNS = (5, 11, 14)  # E 专业阶段、K 实际完成日期、N 复盘说明
SUMMARY_MANUAL_COLUMNS = (4, 12)  # D 周目标、L 复盘结论


@dataclass
class GenerationReport:
    weeks: list[str] = field(default_factory=list)
    detail_count: int = 0
    skipped: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    applied: bool = False

    def to_dict(self) -> dict:
        return {
            "weeks": self.weeks,
            "detail_count": self.detail_count,
            "skipped": self.skipped,
            "errors": self.errors,
            "applied": self.applied,
        }


def _week_dates(week: str) -> tuple[date, date]:
    match = WEEK_RE.match(week)
    year, number = int(match.group(1)), int(match.group(2))
    return date.fromisocalendar(year, number, 1), date.fromisocalendar(year, number, 7)


def read_assigned_tasks(plan_path: Path) -> tuple[list[dict], list[str], list[str]]:
    """返回（按迭代周分组的任务列表、格式非法被跳过的任务ID、错误列表）。"""
    workbook = load_workbook(plan_path)
    try:
        control = workbook[CONTROL_SHEET]
        if control.cell(row=3, column=WEEK_COLUMN).value != ITERATION_WEEK_HEADER:
            return [], [], [f"{CONTROL_SHEET} 第{WEEK_COLUMN}列应为表头'{ITERATION_WEEK_HEADER}'"]
        tasks: list[dict] = []
        skipped: list[str] = []
        for row in control.iter_rows(min_row=4, values_only=True):
            task_id = row[0]
            if not task_id:
                continue
            week = row[WEEK_COLUMN - 1] if len(row) >= WEEK_COLUMN else None
            if week in (None, ""):
                continue
            week = str(week).strip()
            if not WEEK_RE.match(week):
                skipped.append(str(task_id))
                continue
            tasks.append({"task_id": str(task_id), "week": week})
        return tasks, skipped, []
    finally:
        workbook.close()


def plan_generation(plan_path: Path) -> GenerationReport:
    tasks, skipped, errors = read_assigned_tasks(plan_path)
    weeks = sorted({task["week"] for task in tasks})
    capacity = SUMMARY_END - SUMMARY_START + 1
    if len(weeks) > capacity:
        errors.append(f"迭代周数 {len(weeks)} 超过汇总区预留 {capacity} 行（{SUMMARY_START}-{SUMMARY_END} 行）")
    return GenerationReport(weeks=weeks, detail_count=len(tasks), skipped=skipped, errors=errors)


def _index_formula(source_column: str, row: int) -> str:
    return (
        f'=IFERROR(INDEX(项目总控台账!{source_column}:{source_column},'
        f'MATCH($B{row},项目总控台账!$A:$A,0)),"")'
    )


def _summary_formulas(row: int) -> dict[int, str]:
    return {
        6: f"=COUNTIF($A${DETAIL_START}:$A${DETAIL_END},A{row})",
        7: f"=COUNTIFS($A${DETAIL_START}:$A${DETAIL_END},A{row},$I${DETAIL_START}:$I${DETAIL_END},\"已完成\")",
        8: f"=IFERROR(G{row}/F{row},0)",
        9: f"=COUNTIFS($A${DETAIL_START}:$A${DETAIL_END},A{row},$L${DETAIL_START}:$L${DETAIL_END},\"是\")",
        10: f"=COUNTIFS($A${DETAIL_START}:$A${DETAIL_END},A{row},$I${DETAIL_START}:$I${DETAIL_END},\"暂停\")",
        11: f"=IF(J{row}>0,\"暂停\",IF(I{row}>0,\"延期\",IF(H{row}=1,\"已完成\",\"进行中\")))",
    }


def _copy_style(template_cell, target_cell) -> None:
    from copy import copy
    target_cell._style = copy(template_cell._style)


def apply_generation(plan_path: Path) -> GenerationReport:
    report = plan_generation(plan_path)
    if report.errors:
        return report

    workbook = load_workbook(plan_path)
    try:
        control = workbook[CONTROL_SHEET]
        iteration = workbook[ITERATION_SHEET]

        # 旧明细人工列按任务保留；旧汇总人工列按周保留
        detail_manual: dict[str, dict[int, object]] = {}
        for row in iteration.iter_rows(min_row=DETAIL_START, max_row=DETAIL_END, max_col=15):
            task_id = row[1].value
            if task_id:
                detail_manual[str(task_id)] = {
                    column: row[column - 1].value for column in DETAIL_MANUAL_COLUMNS
                }
        summary_manual: dict[str, dict[int, object]] = {}
        for row in iteration.iter_rows(min_row=SUMMARY_START, max_row=SUMMARY_END, max_col=12):
            week = row[0].value
            if week:
                saved = {column: row[column - 1].value for column in SUMMARY_MANUAL_COLUMNS}
                original = row[4].value
                if isinstance(original, str) and not original.startswith("="):
                    saved[5] = original
                summary_manual[str(week)] = saved

        sorted_tasks = sorted(
            ((task["week"], task["task_id"]) for task in
             read_assigned_tasks(plan_path)[0]),
        )

        # 汇总区：每周一行，写入周一/周日日期与统计公式
        summary_template = iteration.cell(row=SUMMARY_START, column=1)
        for index, week in enumerate(report.weeks):
            row_number = SUMMARY_START + index
            monday, sunday = _week_dates(week)
            if row_number != SUMMARY_START:
                _copy_style(summary_template, iteration.cell(row=row_number, column=1))
            iteration.cell(row=row_number, column=1, value=week)
            cell_b = iteration.cell(row=row_number, column=2, value=monday)
            cell_c = iteration.cell(row=row_number, column=3, value=sunday)
            cell_b.number_format = DATE_FORMAT
            cell_c.number_format = DATE_FORMAT
            saved = summary_manual.get(week, {})
            for column in (4, 12):
                iteration.cell(row=row_number, column=column).value = saved.get(column)
            if 5 in saved:
                iteration.cell(row=row_number, column=5).value = saved[5]
            for column, formula in _summary_formulas(row_number).items():
                iteration.cell(row=row_number, column=column, value=formula)
        for row_number in range(SUMMARY_START + len(report.weeks), SUMMARY_END + 1):
            for column in range(1, 13):
                iteration.cell(row=row_number, column=column).value = None

        # 明细区：任务ID为键，其余列 INDEX/MATCH 查总控台账
        detail_templates = {column: iteration.cell(row=DETAIL_START, column=column) for column in range(1, 16)}
        for index, (week, task_id) in enumerate(sorted_tasks):
            row_number = DETAIL_START + index
            if row_number != DETAIL_START:
                for column in range(1, 16):
                    _copy_style(detail_templates[column], iteration.cell(row=row_number, column=column))
            iteration.cell(row=row_number, column=1, value=week)
            iteration.cell(row=row_number, column=2, value=task_id)
            for column_letter, source in DETAIL_SOURCES.items():
                column = ord(column_letter) - 64
                iteration.cell(row=row_number, column=column, value=_index_formula(source, row_number))
            saved = detail_manual.get(task_id, {})
            for column in DETAIL_MANUAL_COLUMNS:
                iteration.cell(row=row_number, column=column).value = saved.get(column)
            iteration.cell(row=row_number, column=12, value=(
                f'=IF(AND(H{row_number}<TODAY(),I{row_number}<>"已完成"),"是","否")'
            ))
            iteration.cell(row=row_number, column=13, value=(
                f'=IF(I{row_number}="暂停","阶段暂停",'
                f'IF(L{row_number}="是","计划完成日期已过",""))'
            ))
            iteration.cell(row=row_number, column=15, value=(
                f'=A{row_number}&"|"&B{row_number}&"|"&E{row_number}'
            ))
        for row_number in range(DETAIL_START + len(sorted_tasks), DETAIL_END + 1):
            for column in range(1, 16):
                iteration.cell(row=row_number, column=column).value = None

        workbook.save(plan_path)
        report.applied = True
        return report
    finally:
        workbook.close()
