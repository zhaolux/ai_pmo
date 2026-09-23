#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from collections import Counter
from uuid import uuid4

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ai_pmo.current_workbooks import resolve_current


@dataclass(frozen=True)
class ProjectMetrics:
    total_tasks: int
    completed_tasks: int
    in_progress_tasks: int
    not_started_or_pending: int
    paused_tasks: int
    average_progress: float
    overdue_tasks: int
    due_in_14_days: int
    total_risks: int
    major_residual_risks: int
    high_residual_risks: int
    total_gates: int
    required_gates_not_passed: int
    unclosed_gates: int
    pending_decisions: int
    key_stakeholders: int
    unsupportive_stakeholders: int
    communication_overdue: int
    total_milestones: int
    completed_milestones: int


MODULE_DIRECTORIES = [
    "00_使用说明与配置", "01_项目启动与治理", "02_计划与进度管理",
    "03_需求与范围管理", "04_系统集成管理", "05_成本与合同管理",
    "06_风险问题与变更", "07_质量测试与验收", "08_沟通会议与报告",
    "09_上线移交与运营", "10_复盘与经验沉淀", "11_AI_PMO中心",
]


def _quality_issue(
    source: str, record_id: str, name: str, owner: str, start, finish,
    problem: str, status: str, action: str, source_row: int | str,
    check_code: str,
) -> dict[str, object]:
    return {
        "source": source, "record_id": record_id, "name": name, "owner": owner,
        "start": start, "finish": finish, "problem": problem, "status": status,
        "action": action, "source_row": source_row, "check_code": check_code,
    }


def collect_data_quality_issues(suite: Path, as_of: date) -> list[dict[str, object]]:
    issues: list[dict[str, object]] = []
    for directory_name in MODULE_DIRECTORIES:
        directory = suite / directory_name
        if not directory.exists() or not any(directory.glob("*.xlsx")):
            issues.append(_quality_issue(
                directory_name, "FILE-MISSING", "模块工作簿完整性", "项目经理", None, None,
                "模块目录或工作簿缺失", "阻断", "补齐模块工作簿后重新运行刷新", "-", "FILE-MISSING",
            ))
            continue
        if directory_name == "04_系统集成管理":
            continue
        workbooks = [p for p in directory.glob("*.xlsx") if not p.name.startswith("~$")]
        if directory_name == "11_AI_PMO中心":
            workbooks = [p for p in workbooks if p.stem.startswith("AI PMO中心-")]
        for workbook_path in workbooks:
            workbook = load_workbook(workbook_path, data_only=True, read_only=True)
            try:
                if "_数据接口" not in workbook.sheetnames:
                    issues.append(_quality_issue(
                        workbook_path.name, "MODULE-CODE", "模块编号检查", "项目经理", None, None,
                        "缺少_data接口工作表", "阻断", "补充_data接口并填写模块编号", "-", "MODULE-CODE",
                    ))
                else:
                    expected = directory_name[:2]
                    actual = str(workbook["_数据接口"]["B4"].value or "").zfill(2)
                    if actual != expected:
                        issues.append(_quality_issue(
                            workbook_path.name, "MODULE-CODE", "模块编号检查", "项目经理", None, None,
                            f"模块编号应为{expected}，当前为{actual or '空'}", "阻断",
                            "修正_data接口B4后重新运行刷新", 4, "MODULE-CODE",
                        ))
            finally:
                workbook.close()

    plan_path = resolve_current(suite, "plan")
    workbook = load_workbook(plan_path, data_only=True, read_only=True)
    try:
        sheet = workbook["项目总控台账"]
        rows = [(row_number, row) for row_number, row in enumerate(
            sheet.iter_rows(min_row=4, values_only=True), start=4
        ) if row[0]]
        counts = Counter(str(row[0]).strip() for _, row in rows)
        allowed_statuses = {"未开始", "待确认", "进行中", "已完成", "暂停", "挂起", "未分配"}
        for row_number, row in rows:
            task_id = str(row[0]).strip()
            task_name = str(row[4] or "")
            status = str(row[6] or "")
            owner = str(row[7] or "")
            start = _as_date(row[8])
            finish = _as_date(row[9])
            if counts[task_id] > 1:
                issues.append(_quality_issue(
                    "项目总控台账", task_id, task_name, owner, start, finish,
                    "任务ID重复", "阻断", "为重复记录分配唯一任务ID", row_number, "DUPLICATE-ID",
                ))
            if status not in allowed_statuses:
                issues.append(_quality_issue(
                    "项目总控台账", task_id, task_name, owner, start, finish,
                    f"状态值不在标准枚举中：{status or '空'}", "需整改",
                    "使用标准状态：未开始、待确认、进行中、已完成、暂停、挂起或未分配",
                    row_number, "INVALID-STATUS",
                ))
            if not owner or finish is None:
                missing = "负责人" if not owner else "计划完成日期"
                issues.append(_quality_issue(
                    "项目总控台账", task_id, task_name, owner, start, finish,
                    f"缺少{missing}", "需整改", f"在源台账补充{missing}", row_number, "MISSING-FIELD",
                ))
            if start and finish and start > finish:
                issues.append(_quality_issue(
                    "项目总控台账", task_id, task_name, owner, start, finish,
                    "计划开始日期晚于计划完成日期", "阻断", "核对并修正计划日期", row_number,
                    "INVALID-DATE",
                ))
            if finish and finish < as_of and status != "已完成" and not (row[29] or "").strip():
                issues.append(_quality_issue(
                    "项目总控台账", task_id, task_name, owner, start, finish,
                    "计划完成日已过，执行状态和实际证据未确认", "需整改",
                    "补充实际状态、实际日期、问题/阻塞及完成证据", row_number,
                    "TASK-OVERDUE-EVIDENCE",
                ))
    finally:
        workbook.close()
    return issues


def next_run_id(existing_ids: list[str], run_date: date) -> str:
    prefix = f"RUN-{run_date:%Y%m%d}-"
    numbers = []
    for value in existing_ids:
        if value and value.startswith(prefix):
            try:
                numbers.append(int(value[len(prefix):]))
            except ValueError:
                pass
    return f"{prefix}{max(numbers, default=0) + 1:03d}"


def build_run_log_record(
    run_id: str, run_date: date, metrics: ProjectMetrics,
    issues: list[dict[str, object]],
) -> dict[str, object]:
    blocking = sum(item["status"] == "阻断" for item in issues)
    corrections = sum(item["status"] == "需整改" for item in issues)
    risk_level = "高" if blocking or metrics.major_residual_risks or metrics.overdue_tasks else (
        "中" if corrections or metrics.high_residual_risks else "低"
    )
    return {
        "run_id": run_id, "run_time": datetime.now().replace(microsecond=0),
        "analysis_type": "全量刷新与质检", "data_date": run_date,
        "scope": "里程碑、总控、门禁、风险、决策、相关方及模块完整性",
        "finding": f"发现{len(issues)}项数据质量问题，其中{blocking}项阻断、{corrections}项需整改；{metrics.overdue_tasks}项任务已延期",
        "risk_level": risk_level,
        "advice": (
            "优先处理阻断项，并由任务负责人补齐逾期任务的实际状态、日期和完成证据"
            if blocking else "优先处理需整改事项，并由任务负责人补齐逾期任务的实际状态、日期和完成证据"
        ),
        "human_conclusion": "待项目经理确认", "handler": "项目经理",
    }


def preview_sync(suite: Path, as_of: date) -> dict[str, object]:
    return {
        "data_date": as_of.isoformat(),
        "target": resolve_current(suite, "ai_pmo"),
        "sources": {
            "plan": resolve_current(suite, "plan"),
            "gate": resolve_current(suite, "gate"),
            "risk": resolve_current(suite, "risk"),
            "change": resolve_current(suite, "change"),
        },
    }


def backup_workbook(target: Path) -> Path:
    archive = target.parent / "_archive" / "自动刷新前备份"
    archive.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = archive / f"{target.stem}-{timestamp}-{uuid4().hex[:8]}{target.suffix}"
    shutil.copy2(target, backup)
    digest = lambda path: hashlib.sha256(path.read_bytes()).digest()
    if digest(target) != digest(backup):
        raise OSError(f"备份校验失败：{backup}")
    return backup


def find_next_log_row(path: Path) -> tuple[int, list[str]]:
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        sheet = workbook["AI运行日志"]
        existing_ids = [str(sheet.cell(row, 1).value or "") for row in range(5, 105)]
        row = next((
            number for number in range(5, 105)
            if all(sheet.cell(number, column).value in (None, "") for column in range(1, 11))
        ), None)
    finally:
        workbook.close()
    if row is None:
        raise RuntimeError("AI运行日志已满，请先归档历史运行记录")
    return row, existing_ids


def validate_refresh_commands(commands: list[dict], log_row: int) -> None:
    """Keep refresh writes inside declared generated cells."""
    dashboard = {f"{column}6" for column in "ABCDEFGHIJKLMN"} | {
        "B3", "A9", "E9", "F9", "G9", "H9",
        "B12", "B13", "B14", "B15", "B16",
        "F12", "F13", "F15", "F16",
        "I18", "J18", "K18", "L18", "I19", "J19", "K19", "L19",
    } | {f"{column}{row}" for column in "IJKLM" for row in range(12, 17)} \
      | {f"{column}{row}" for column in "ABCDE" for row in range(20, 25)}
    diagnosis = {"A2", "B4", "C11", "C12", "C13", "C14", "C16"} | {
        f"{column}7" for column in "ABCDEFGH"
    }
    fixed = {
        "使用说明": {"B9"},
        "项目驾驶舱": dashboard,
        "AI项目诊断": diagnosis,
        "PMO周报": {"B3", "B29", "B30", "B31", "B32", "B33", "B34"},
    }
    for command in commands:
        path = command.get("path", "")
        parts = path.split("/")
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", parts[2]) if len(parts) == 3 else None
        if command.get("command") != "set" or not match:
            raise ValueError(f"刷新命令触及人工维护区域或格式无效：{path}")
        sheet, cell = parts[1], parts[2]
        column, row = match.group(1), int(match.group(2))
        allowed = cell in fixed.get(sheet, set())
        if sheet == "_数据接口":
            allowed = (column == "A" and 22 <= row <= 42) or (
                column == "B" and (row == 5 or 7 <= row <= 15 or 22 <= row <= 42)
            )
        elif sheet == "两周督办":
            allowed = column in set("ABCDEFGHIJK") and row >= 5
        elif sheet == "数据质量检查":
            allowed = cell == "A2" or (column in set("ABCDEFGHIJ") and row >= 5)
        elif sheet == "AI运行日志":
            allowed = column in set("ABCDEFGHIJ") and row == log_row
        if not allowed:
            raise ValueError(f"刷新命令触及人工维护区域：{path}")


def _as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def extract_project_metrics(suite: Path, as_of: date) -> ProjectMetrics:
    plan_path = resolve_current(suite, "plan")
    gate_path = resolve_current(suite, "gate")
    risk_path = resolve_current(suite, "risk")
    change_path = resolve_current(suite, "change")

    plan_book = load_workbook(plan_path, data_only=True, read_only=True)
    try:
        sheet = plan_book["项目总控台账"]
        rows = [row for row in sheet.iter_rows(min_row=4, values_only=True) if row[0]]
        statuses = [row[6] for row in rows]
        total_tasks = len(rows)
        completed = statuses.count("已完成")
        in_progress = statuses.count("进行中")
        pending = sum(status in {"未开始", "待确认", "未分配"} for status in statuses)
        paused = sum(status in {"暂停", "挂起"} for status in statuses)
        progress_values = [float(row[27] or 0) for row in rows]
        average_progress = sum(progress_values) / total_tasks if total_tasks else 0
        overdue = sum(row[28] == "已延期" for row in rows)
        due_soon = 0
        for row in rows:
            due = _as_date(row[9])
            if due and row[6] != "已完成" and 0 <= (due - as_of).days <= 14:
                due_soon += 1
        milestone_sheet = plan_book["里程碑计划"]
        milestone_rows = [row for row in milestone_sheet.iter_rows(min_row=4, values_only=True) if row[0]]
        total_milestones = len(milestone_rows)
        completed_milestones = sum(row[10] == "已完成" for row in milestone_rows)
    finally:
        plan_book.close()

    gate_book = load_workbook(gate_path, data_only=True, read_only=True)
    try:
        gate_sheet = gate_book["启动门禁清单"]
        gate_rows = [row for row in gate_sheet.iter_rows(min_row=8, values_only=True) if row[0]]
        total_gates = len(gate_rows)
        required_not_passed = sum(row[8] == "是" and row[7] != "已通过" for row in gate_rows)
        unclosed_gates = sum(row[7] != "已通过" for row in gate_rows)
        stakeholder_sheet = gate_book["相关方管理"]
        stakeholder_rows = [row for row in stakeholder_sheet.iter_rows(min_row=9, values_only=True) if row[0]]
        key_stakeholders = sum(row[13] == "重点管理" for row in stakeholder_rows)
        unsupportive = sum(row[9] == "不支持" for row in stakeholder_rows)
        communication_overdue = sum(
            (_as_date(row[19]) is not None and _as_date(row[19]) < as_of)
            for row in stakeholder_rows
        )
    finally:
        gate_book.close()

    risk_book = load_workbook(risk_path, data_only=True, read_only=True)
    try:
        risk_sheet = risk_book["风险登记册"]
        risk_rows = [row for row in risk_sheet.iter_rows(min_row=9, values_only=True) if row[0]]
        total_risks = len(risk_rows)
        major = sum(row[33] == "重大" for row in risk_rows)
        high = sum(row[33] == "高" for row in risk_rows)
    finally:
        risk_book.close()

    change_book = load_workbook(change_path, data_only=True, read_only=True)
    try:
        decision_sheet = change_book["待决策事项"]
        decision_rows = [row for row in decision_sheet.iter_rows(min_row=5, values_only=True) if row[0]]
        pending_decisions = sum(row[8] not in {"已决策", "已关闭"} for row in decision_rows)
    finally:
        change_book.close()

    return ProjectMetrics(
        total_tasks, completed, in_progress, pending, paused, average_progress,
        overdue, due_soon, total_risks, major, high, total_gates,
        required_not_passed, unclosed_gates, pending_decisions, key_stakeholders,
        unsupportive, communication_overdue, total_milestones, completed_milestones,
    )


def extract_two_week_watchlist(suite: Path, as_of: date) -> list[dict[str, object]]:
    plan_path = resolve_current(suite, "plan")
    workbook = load_workbook(plan_path, data_only=True, read_only=True)
    try:
        sheet = workbook["项目总控台账"]
        result = []
        for row in sheet.iter_rows(min_row=4, values_only=True):
            if not row[0] or row[6] == "已完成":
                continue
            due = _as_date(row[9])
            if due is None:
                continue
            days = (due - as_of).days
            if days > 14:
                continue
            delay_status = "已延期" if days < 0 else ("临近到期" if days <= 3 else "正常")
            action = "确认实际状态；完成、纠偏或提交基线变更" if days < 0 else "按计划提交可验收证据"
            result.append({
                "task_id": row[0], "task_name": row[4], "milestone": row[5],
                "status": row[6], "owner": row[7], "due": due,
                "days_to_due": days, "delay_status": delay_status,
                "issue": row[29] or "", "action": action,
            })
        return sorted(result, key=lambda item: (item["due"], item["task_id"]))
    finally:
        workbook.close()


def extract_milestone_overview(suite: Path, limit: int = 5) -> list[dict[str, object]]:
    """Read the first `limit` milestones (by plan row order) for the dashboard block."""
    plan_path = resolve_current(suite, "plan")
    workbook = load_workbook(plan_path, data_only=True, read_only=True)
    try:
        sheet = workbook["里程碑计划"]
        result = []
        for row in sheet.iter_rows(min_row=4, values_only=True):
            if not row[0] or not str(row[0]).startswith("MS"):
                continue
            result.append({
                "name": row[2], "due": _as_date(row[7]), "status": row[10],
                "health": row[12], "owner": row[4],
            })
            if len(result) >= limit:
                break
        return result
    finally:
        workbook.close()


def extract_management_focus(suite: Path, limit: int = 5) -> list[dict[str, object]]:
    """管理关注点：重大级在前（剩余重大/高风险、待决策事项），同级按日期升序取前 `limit` 项。"""
    focus: list[dict[str, object]] = []
    risk_path = resolve_current(suite, "risk")
    risk_book = load_workbook(risk_path, data_only=True, read_only=True)
    try:
        for row in risk_book["风险登记册"].iter_rows(min_row=9, values_only=True):
            if not row[0] or str(row[0]).startswith("RSK") is False:
                continue
            if row[33] not in ("重大", "高") or row[39] not in (None, ""):
                continue
            due = _as_date(row[27])
            strategy = str(row[24] or "").strip()
            next_step = (f"最迟 {due.isoformat()} 完成应对" if due else "完成应对") \
                + (f"（策略：{strategy}）" if strategy else "")
            focus.append({
                "item": str(row[1] or row[0]), "next": next_step,
                "owner": str(row[22] or ""), "level": str(row[33]),
                "due": due, "grade_rank": 0 if row[33] == "重大" else 1,
                "source_id": str(row[0]),
            })
    finally:
        risk_book.close()
    change_path = resolve_current(suite, "change")
    change_book = load_workbook(change_path, data_only=True, read_only=True)
    try:
        for row in change_book["待决策事项"].iter_rows(min_row=5, values_only=True):
            if not row[0] or row[8] in ("已决策", "已关闭"):
                continue
            due = _as_date(row[7])
            proposal = str(row[5] or "").strip()
            next_step = (f"{due.isoformat()} 前完成决策" if due else "完成决策") \
                + (f"：{proposal}" if proposal else "")
            focus.append({
                "item": str(row[2] or row[0]), "next": next_step,
                "owner": str(row[6] or ""), "level": "重大",
                "due": due, "grade_rank": 0, "source_id": str(row[0]),
            })
    finally:
        change_book.close()
    focus.sort(key=lambda x: (x["grade_rank"], x["due"] or date.max, x["source_id"]))
    return focus[:limit]


def extract_migration_metrics(suite: Path) -> dict[str, object] | None:
    """Read migration inventory metrics from the registered migration workbook interface."""
    try:
        migration_path = resolve_current(suite, "migration")
    except (ValueError, FileNotFoundError):
        return None
    workbook = load_workbook(migration_path, data_only=True, read_only=True)
    try:
        if "_数据接口" not in workbook.sheetnames:
            return None
        interface = {
            str(row[0]): row[1]
            for row in workbook["_数据接口"].iter_rows(values_only=True)
            if len(row) > 1 and row[0] not in (None, "")
        }
        gate_date = interface.get("G15 目标日期")
        if isinstance(gate_date, datetime):
            gate_date = gate_date.date().isoformat()
        return {
            "来源线索数": interface.get("来源线索数"),
            "已盘点来源系统数": interface.get("已盘点来源系统数"),
            "迁移门禁数": interface.get("迁移门禁数"),
            "G15目标日期": gate_date,
        }
    finally:
        workbook.close()


def ensure_watchlist_manual_alignment(path: Path, watchlist: list[dict[str, object]]) -> None:
    """Reject a refresh that would attach manually entered L:M values to another task."""
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        sheet = workbook["两周督办"]
        for row_number in range(5, sheet.max_row + 1):
            if all(sheet[f"{column}{row_number}"].value in (None, "") for column in "LM"):
                continue
            current_id = str(sheet[f"B{row_number}"].value or "").strip()
            index = row_number - 5
            incoming_id = str(watchlist[index]["task_id"]).strip() if index < len(watchlist) else ""
            if not current_id or current_id != incoming_id:
                raise RuntimeError(
                    f"两周督办第{row_number}行人工列与任务ID {current_id or '空'} 错配；"
                    f"新任务ID为 {incoming_id or '空'}。请先按任务ID处理L:M列，再刷新。"
                )
    finally:
        workbook.close()


def find_stale_watchlist_rows(path: Path, *, new_count: int) -> list[int]:
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        sheet = workbook["两周督办"]
        return [
            row_number for row_number in range(5 + new_count, sheet.max_row + 1)
            if any(sheet[f"{column}{row_number}"].value not in (None, "") for column in "ABCDEFGHIJK")
        ]
    finally:
        workbook.close()


def sync_ai_pmo(
    suite: Path, as_of: date, *, expected_target: Path | None = None,
) -> tuple[Path, ProjectMetrics, Path]:
    ai_pmo = resolve_current(suite, "ai_pmo")
    if expected_target is not None and ai_pmo != expected_target:
        raise RuntimeError(f"刷新目标在预览后已变化：{expected_target} -> {ai_pmo}")
    metrics = extract_project_metrics(suite, as_of)
    watchlist = extract_two_week_watchlist(suite, as_of)
    milestone_overview = extract_milestone_overview(suite)
    management_focus = extract_management_focus(suite)
    ensure_watchlist_manual_alignment(ai_pmo, watchlist)
    stale_watchlist_rows = find_stale_watchlist_rows(ai_pmo, new_count=len(watchlist))
    quality_issues = collect_data_quality_issues(suite, as_of)
    migration_metrics = extract_migration_metrics(suite)
    log_row, existing_ids = find_next_log_row(ai_pmo)
    log_record = build_run_log_record(next_run_id(existing_ids, as_of), as_of, metrics, quality_issues)
    sources = {
        "plan": resolve_current(suite, "plan").relative_to(suite).as_posix(),
        "gate": resolve_current(suite, "gate").relative_to(suite).as_posix(),
        "risk": resolve_current(suite, "risk").relative_to(suite).as_posix(),
    }
    formulas = {
        "A6": 11, "B6": 22, "C6": 23, "D6": 24, "E6": 25, "F6": 26,
        "G6": 27, "H6": 28, "I6": 33, "J6": 34, "K6": 35, "L6": 14,
        "A9": 26, "E9": 29, "F9": 30, "G9": 31, "H9": 32,
        "B12": 22, "B13": 23, "B14": 24, "B15": 25, "B16": 27,
        "F12": 33, "F13": 34, "F15": 29, "F16": 14,
    }
    commands = [
        {"command": "set", "path": f"/项目驾驶舱/{cell}", "props": {"formula": f"'_数据接口'!B{row}"}}
        for cell, row in formulas.items()
    ]
    commands.extend([
        {"command": "set", "path": "/项目驾驶舱/M6", "props": {"formula": "'_数据接口'!B37&\"/\"&'_数据接口'!B36"}},
        {"command": "set", "path": "/项目驾驶舱/N6", "props": {"formula": "'_数据接口'!B38"}},
    ])
    values = {
        5: as_of.isoformat(), 7: sources["plan"], 8: sources["gate"], 9: sources["risk"],
        10: "源台账更新后运行 automation/scripts/sync_all_to_ai_pmo.py",
        11: metrics.total_tasks, 12: metrics.total_risks, 13: metrics.total_gates,
        14: metrics.pending_decisions, 15: as_of.isoformat(), 22: metrics.completed_tasks,
        23: metrics.in_progress_tasks, 24: metrics.not_started_or_pending,
        25: metrics.paused_tasks, 26: metrics.average_progress, 27: metrics.overdue_tasks,
        28: metrics.due_in_14_days, 29: metrics.required_gates_not_passed,
        30: metrics.key_stakeholders, 31: metrics.unsupportive_stakeholders,
        32: metrics.communication_overdue, 33: metrics.major_residual_risks,
        34: metrics.high_residual_risks, 35: metrics.unclosed_gates,
        36: metrics.total_milestones, 37: metrics.completed_milestones,
        38: "红" if (metrics.major_residual_risks or metrics.required_gates_not_passed or metrics.overdue_tasks) else ("黄" if (metrics.high_residual_risks or metrics.pending_decisions) else "绿"),
    }
    if migration_metrics is not None:
        values.update({
            39: migration_metrics.get("来源线索数"), 40: migration_metrics.get("已盘点来源系统数"),
            41: migration_metrics.get("G15目标日期"), 42: migration_metrics.get("迁移门禁数"),
        })
    labels = {
        22: "已完成任务", 23: "进行中任务", 24: "未开始/待确认任务",
        25: "暂停/挂起任务", 26: "平均进度", 27: "已延期任务",
        28: "14天内到期", 29: "必过门禁未通过", 30: "重点相关方",
        31: "不支持相关方", 32: "沟通逾期", 33: "重大剩余风险",
        34: "高剩余风险", 35: "未关闭门禁", 36: "里程碑总数",
        37: "已完成里程碑", 38: "项目健康度",
        39: "迁移来源线索数", 40: "已盘点来源系统数",
        41: "G15目标日期", 42: "迁移门禁数",
    }
    if migration_metrics is not None:
        commands.extend([
            {"command": "set", "path": "/项目驾驶舱/I18", "props": {"value": "迁移来源线索"}},
            {"command": "set", "path": "/项目驾驶舱/I19", "props": {"formula": "'_数据接口'!B39"}},
            {"command": "set", "path": "/项目驾驶舱/J18", "props": {"value": "已盘点来源"}},
            {"command": "set", "path": "/项目驾驶舱/J19", "props": {"formula": "'_数据接口'!B40"}},
            {"command": "set", "path": "/项目驾驶舱/K18", "props": {"value": "G15目标日期"}},
            {"command": "set", "path": "/项目驾驶舱/K19", "props": {"formula": "'_数据接口'!B41"}},
            {"command": "set", "path": "/项目驾驶舱/L18", "props": {"value": "迁移门禁"}},
            {"command": "set", "path": "/项目驾驶舱/L19", "props": {"formula": "'_数据接口'!B42"}},
        ])
    for index, item in enumerate(milestone_overview, start=12):
        commands.extend([
            {"command": "set", "path": f"/项目驾驶舱/I{index}", "props": {"value": item["name"] or ""}},
            {"command": "set", "path": f"/项目驾驶舱/J{index}", "props": {
                "value": item["due"].isoformat() if item["due"] else "", "type": "date", "numberformat": "yyyy-mm-dd"}},
            {"command": "set", "path": f"/项目驾驶舱/K{index}", "props": {"value": item["status"] or ""}},
            {"command": "set", "path": f"/项目驾驶舱/L{index}", "props": {"value": item["health"] or ""}},
            {"command": "set", "path": f"/项目驾驶舱/M{index}", "props": {"value": item["owner"] or ""}},
        ])
    for index in range(12 + len(milestone_overview), 17):
        for column in "IJKLM":
            commands.append({"command": "set", "path": f"/项目驾驶舱/{column}{index}", "props": {"value": ""}})
    for index, item in enumerate(management_focus, start=20):
        commands.extend([
            {"command": "set", "path": f"/项目驾驶舱/A{index}", "props": {"value": index - 19, "type": "number"}},
            {"command": "set", "path": f"/项目驾驶舱/B{index}", "props": {"value": item["item"]}},
            {"command": "set", "path": f"/项目驾驶舱/C{index}", "props": {"value": item["next"]}},
            {"command": "set", "path": f"/项目驾驶舱/D{index}", "props": {"value": item["owner"]}},
            {"command": "set", "path": f"/项目驾驶舱/E{index}", "props": {"value": item["level"]}},
        ])
    for index in range(20 + len(management_focus), 25):
        for column in "ABCDE":
            commands.append({"command": "set", "path": f"/项目驾驶舱/{column}{index}", "props": {"value": ""}})
    for row, label in labels.items():
        commands.append({"command": "set", "path": f"/_数据接口/A{row}", "props": {"value": label}})
    for row, value in values.items():
        kind = "number" if isinstance(value, (int, float)) else "string"
        commands.append({"command": "set", "path": f"/_数据接口/B{row}", "props": {"value": value, "type": kind}})
    commands.extend([
        {"command": "set", "path": "/项目驾驶舱/B3", "props": {"value": as_of.isoformat(), "type": "date", "numberformat": "yyyy-mm-dd"}},
        {"command": "set", "path": "/使用说明/B9", "props": {"value": "执行标准刷新程序 sync_all_to_ai_pmo.py，将进度、门禁、风险、决策、相关方、前 5 项里程碑及管理关注点同步至本工作簿_data接口；成本指标请使用单独成本同步程序"}},
        {"command": "set", "path": "/AI项目诊断/A2", "props": {"formula": "\"数据截止\"&'_数据接口'!B5&\"。客观状态根据总控、风险和启动门禁快照形成，管理结论须由项目经理确认。\""}},
        {"command": "set", "path": "/AI项目诊断/B4", "props": {"formula": "'_数据接口'!B5"}},
        {"command": "set", "path": "/AI项目诊断/A7", "props": {"formula": "'_数据接口'!B11"}},
        {"command": "set", "path": "/AI项目诊断/B7", "props": {"formula": "'_数据接口'!B22"}},
        {"command": "set", "path": "/AI项目诊断/C7", "props": {"formula": "'_数据接口'!B27"}},
        {"command": "set", "path": "/AI项目诊断/D7", "props": {"formula": "'_数据接口'!B28"}},
        {"command": "set", "path": "/AI项目诊断/E7", "props": {"formula": "'_数据接口'!B33"}},
        {"command": "set", "path": "/AI项目诊断/F7", "props": {"formula": "'_数据接口'!B34"}},
        {"command": "set", "path": "/AI项目诊断/G7", "props": {"formula": "'_数据接口'!B29"}},
        {"command": "set", "path": "/AI项目诊断/H7", "props": {"formula": "'_数据接口'!B14"}},
        {"command": "set", "path": "/AI项目诊断/C11", "props": {"formula": "'_数据接口'!B27&\"项任务已延期或计划完成日已过\""}},
        {"command": "set", "path": "/AI项目诊断/C12", "props": {"formula": "'_数据接口'!B33&\"项重大、\"&'_数据接口'!B34&\"项高剩余风险\""}},
        {"command": "set", "path": "/AI项目诊断/C13", "props": {"formula": "'_数据接口'!B29&\"项必过门禁尚未通过\""}},
        {"command": "set", "path": "/AI项目诊断/C14", "props": {"formula": "'_数据接口'!B14&\"项事项等待管理层决策\""}},
        {"command": "set", "path": "/AI项目诊断/C16", "props": {"formula": "'_数据接口'!B30&\"名重点管理相关方，\"&'_数据接口'!B31&\"名不支持\""}},
        {"command": "set", "path": "/PMO周报/B3", "props": {"formula": "'_数据接口'!B5"}},
        {"command": "set", "path": "/PMO周报/B29", "props": {"formula": "IF('_数据接口'!B11=0,0,'_数据接口'!B22/'_数据接口'!B11)"}},
        {"command": "set", "path": "/PMO周报/B30", "props": {"formula": "'_数据接口'!B27"}},
        {"command": "set", "path": "/PMO周报/B31", "props": {"formula": "'_数据接口'!B33+'_数据接口'!B34"}},
        {"command": "set", "path": "/PMO周报/B32", "props": {"formula": "'_数据接口'!B29"}},
        {"command": "set", "path": "/PMO周报/B33", "props": {"formula": "'_数据接口'!B14"}},
        {"command": "set", "path": "/PMO周报/B34", "props": {"formula": "'_数据接口'!B27"}},
        {"command": "set", "path": "/数据质量检查/A2", "props": {"value": "自动检查模块文件、模块编号、重复ID、状态、负责人、日期及逾期任务执行证据。请在源台账修正后重新运行刷新。"}},
    ])
    for index, item in enumerate(watchlist, start=5):
        row_values = [
            "是", item["task_id"], item["task_name"], item["milestone"], item["status"],
            item["owner"], item["due"].isoformat(), item["days_to_due"],
            item["delay_status"], item["issue"], item["action"],
        ]
        for column, value in zip("ABCDEFGHIJK", row_values):
            props = {"value": value}
            if column == "G":
                props.update({"type": "date", "numberformat": "yyyy-mm-dd"})
            elif column == "H":
                props["type"] = "number"
            commands.append({"command": "set", "path": f"/两周督办/{column}{index}", "props": props})
    for row_number in stale_watchlist_rows:
        for column in "ABCDEFGHIJK":
            commands.append({
                "command": "set", "path": f"/两周督办/{column}{row_number}",
                "props": {"value": ""},
            })
    for row in range(5, 105):
        for column in "ABCDEFGHIJ":
            commands.append({"command": "set", "path": f"/数据质量检查/{column}{row}", "props": {"value": ""}})
    for row, item in enumerate(quality_issues, start=5):
        values = [
            item["source"], item["record_id"], item["name"], item["owner"],
            item["start"].isoformat() if item["start"] else "",
            item["finish"].isoformat() if item["finish"] else "",
            item["problem"], item["status"], item["action"], item["source_row"],
        ]
        for column, value in zip("ABCDEFGHIJ", values):
            props = {"value": value}
            if column in {"E", "F"} and value:
                props.update({"type": "date", "numberformat": "yyyy-mm-dd"})
            elif column == "J" and isinstance(value, int):
                props["type"] = "number"
            commands.append({"command": "set", "path": f"/数据质量检查/{column}{row}", "props": props})
    log_values = [
        log_record["run_id"], log_record["run_time"].isoformat(sep=" "),
        log_record["analysis_type"], log_record["data_date"].isoformat(),
        log_record["scope"], log_record["finding"], log_record["risk_level"],
        log_record["advice"], log_record["human_conclusion"], log_record["handler"],
    ]
    for column, value in zip("ABCDEFGHIJ", log_values):
        props = {"value": value}
        if column == "B":
            props.update({"type": "date", "numberformat": "yyyy-mm-dd hh:mm"})
        elif column == "D":
            props.update({"type": "date", "numberformat": "yyyy-mm-dd"})
        commands.append({"command": "set", "path": f"/AI运行日志/{column}{log_row}", "props": props})
    validate_refresh_commands(commands, log_row)
    officecli = shutil.which("officecli")
    if not officecli:
        raise RuntimeError("未找到 officecli")
    backup = backup_workbook(ai_pmo)
    with tempfile.NamedTemporaryFile("w", suffix=".json", encoding="utf-8") as stream:
        json.dump(commands, stream, ensure_ascii=False)
        stream.flush()
        subprocess.run([officecli, "batch", str(ai_pmo), "--input", stream.name, "--stop-on-error"], check=True)
    return ai_pmo, metrics, backup


def main() -> None:
    parser = argparse.ArgumentParser(description="预览或同步项目进度、门禁、风险、决策和相关方指标到 AI PMO")
    parser.add_argument("--suite", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--date", type=date.fromisoformat, default=date.today())
    parser.add_argument("--apply", action="store_true", help="备份目标后执行原位刷新；默认仅预览")
    args = parser.parse_args()
    suite = args.suite.resolve()
    preview = preview_sync(suite, args.date)
    print(f"数据日期：{preview['data_date']}")
    for name, path in preview["sources"].items():
        print(f"读取 {name}：{path}")
    print(f"目标工作簿：{preview['target']}")
    print("将更新：数据接口、项目驾驶舱、AI项目诊断、PMO周报、两周督办、数据质量检查、AI运行日志")
    if not args.apply:
        print("仅预览，未修改工作簿。执行刷新请添加 --apply。")
        return
    ai_pmo, metrics, backup = sync_ai_pmo(suite, args.date, expected_target=preview["target"])
    print(f"刷新前备份：{backup}")
    print(f"已同步至 {ai_pmo.name}：任务{metrics.total_tasks}项，延期{metrics.overdue_tasks}项，重大/高风险{metrics.major_residual_risks + metrics.high_residual_risks}项；已写入数据质量检查和运行日志")


if __name__ == "__main__":
    main()
