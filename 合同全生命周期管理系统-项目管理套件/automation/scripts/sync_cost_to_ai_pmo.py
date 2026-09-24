#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ai_pmo.current_workbooks import resolve_current


@dataclass(frozen=True)
class CostMetrics:
    person_days: float
    total_cost: float
    average_day_rate: float
    total_budget: float


def extract_cost_metrics(path: Path) -> CostMetrics:
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook["人工成本预算"]
        total_row = None
        for row in sheet.iter_rows(min_col=1, max_col=1):
            if row[0].value == "合计":
                total_row = row[0].row
                break
        if total_row is None:
            raise ValueError("人工成本预算 缺少'合计'行，无法定位总人天")
        person_days = float(sheet.cell(total_row, 4).value or 0)
        total_cost = float(sheet.cell(total_row, 6).value or 0)
        if person_days <= 0:
            raise ValueError("人工成本预算!D14 总人天必须大于0")
        budget_sheet = workbook["预算基线"]
        total_budget = 0.0
        for row in budget_sheet.iter_rows(min_row=5, max_row=104, min_col=1, max_col=10):
            if row[0].value == "合计":
                continue
            amount = row[9].value
            if isinstance(amount, (int, float)):
                total_budget += float(amount)
        if total_budget <= 0:
            raise ValueError("预算基线 J5:J104 已量化总预算必须大于0")
        return CostMetrics(person_days, total_cost, total_cost / person_days, total_budget)
    finally:
        workbook.close()


def preview_sync(suite: Path) -> dict[str, object]:
    cost = resolve_current(suite, "cost")
    ai_pmo = resolve_current(suite, "ai_pmo")
    return {"source": cost, "target": ai_pmo, "metrics": extract_cost_metrics(cost)}


def backup_workbook(target: Path) -> Path:
    archive = target.parent / "_archive" / "自动刷新前备份"
    archive.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = archive / f"{target.stem}-{timestamp}-{uuid4().hex[:8]}{target.suffix}"
    shutil.copy2(target, backup)
    if hashlib.sha256(target.read_bytes()).digest() != hashlib.sha256(backup.read_bytes()).digest():
        raise OSError(f"备份校验失败：{backup}")
    return backup


def sync(
    suite: Path, *, expected_target: Path | None = None,
) -> tuple[Path, Path, CostMetrics, Path]:
    cost = resolve_current(suite, "cost")
    ai_pmo = resolve_current(suite, "ai_pmo")
    if expected_target is not None and ai_pmo != expected_target:
        raise RuntimeError(f"刷新目标在预览后已变化：{expected_target} -> {ai_pmo}")
    metrics = extract_cost_metrics(cost)
    today = date.today().isoformat()
    source = cost.relative_to(suite).as_posix()
    commands = [
        {"command": "set", "path": "/项目驾驶舱/B9", "props": {"formula": "'_数据接口'!B17"}},
        {"command": "set", "path": "/项目驾驶舱/C9", "props": {"formula": "'_数据接口'!B18"}},
        {"command": "set", "path": "/项目驾驶舱/D9", "props": {"formula": "'_数据接口'!B19"}},
        {"command": "set", "path": "/项目驾驶舱/I9", "props": {"formula": "'_数据接口'!B43"}},
        {"command": "set", "path": "/_数据接口/B16", "props": {"value": source}},
        {"command": "set", "path": "/_数据接口/B17", "props": {"value": metrics.person_days, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B18", "props": {"value": metrics.total_cost, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B19", "props": {"value": metrics.average_day_rate, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B20", "props": {"value": today, "type": "string"}},
        {"command": "set", "path": "/_数据接口/B21", "props": {"value": "已同步"}},
        {"command": "set", "path": "/_数据接口/A43", "props": {"value": "已量化总预算"}},
        {"command": "set", "path": "/_数据接口/B43", "props": {"value": metrics.total_budget, "type": "number"}},
    ]
    officecli = shutil.which("officecli")
    if not officecli:
        raise RuntimeError("未找到 officecli，无法在保留图表的前提下更新 AI PMO")
    backup = backup_workbook(ai_pmo)
    with tempfile.NamedTemporaryFile("w", suffix=".json", encoding="utf-8") as stream:
        json.dump(commands, stream, ensure_ascii=False)
        stream.flush()
        subprocess.run(
            [officecli, "batch", str(ai_pmo), "--input", stream.name, "--stop-on-error"],
            check=True,
        )
    return cost, ai_pmo, metrics, backup


def main() -> None:
    default_suite = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description="将成本模块指标同步到 AI PMO 数据接口")
    parser.add_argument("--suite", type=Path, default=default_suite)
    parser.add_argument("--apply", action="store_true", help="备份目标后执行原位刷新；默认仅预览")
    args = parser.parse_args()
    suite = args.suite.resolve()
    preview = preview_sync(suite)
    print(f"读取成本工作簿：{preview['source']}")
    print(f"目标工作簿：{preview['target']}")
    print("将更新：项目驾驶舱 B9:D9 与 I9、_数据接口 B16:B21 与 A43:B43")
    metrics = preview["metrics"]
    print(f"计划人天={metrics.person_days:g}，计划人工成本={metrics.total_cost:g}，平均人日单价={metrics.average_day_rate:.2f}，已量化总预算={metrics.total_budget:g}")
    if not args.apply:
        print("仅预览，未修改工作簿。执行刷新请添加 --apply。")
        return
    cost, ai_pmo, metrics, backup = sync(suite, expected_target=preview["target"])
    print(f"刷新前备份：{backup}")
    print(f"已从 {cost.name} 同步至 {ai_pmo.name}")


if __name__ == "__main__":
    main()
