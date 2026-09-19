#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import load_workbook


@dataclass(frozen=True)
class CostMetrics:
    person_days: float
    total_cost: float
    average_day_rate: float


def extract_cost_metrics(path: Path) -> CostMetrics:
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook["人工成本预算"]
        person_days = float(sheet["D14"].value or 0)
        total_cost = float(sheet["F14"].value or 0)
        if person_days <= 0:
            raise ValueError("人工成本预算!D14 总人天必须大于0")
        return CostMetrics(person_days, total_cost, total_cost / person_days)
    finally:
        workbook.close()


def newest_workbook(directory: Path, prefix: str) -> Path:
    candidates = [
        path for path in directory.glob(f"{prefix}-*.xlsx")
        if not path.name.startswith("~$") and "_archive" not in path.parts
    ]
    if not candidates:
        raise FileNotFoundError(f"未找到工作簿：{directory}/{prefix}-*.xlsx")
    return max(candidates, key=lambda path: (path.stat().st_mtime, path.name))


def sync(suite: Path) -> tuple[Path, Path, CostMetrics]:
    cost = newest_workbook(suite / "05_成本与合同管理", "成本与合同管理")
    ai_pmo = newest_workbook(suite / "11_AI_PMO中心", "AI PMO中心")
    metrics = extract_cost_metrics(cost)
    today = date.today().isoformat()
    source = cost.relative_to(suite).as_posix()
    commands = [
        {"command": "set", "path": "/项目驾驶舱/B9", "props": {"formula": "'_数据接口'!B17"}},
        {"command": "set", "path": "/项目驾驶舱/C9", "props": {"formula": "'_数据接口'!B18"}},
        {"command": "set", "path": "/项目驾驶舱/D9", "props": {"formula": "'_数据接口'!B19"}},
        {"command": "set", "path": "/_数据接口/B16", "props": {"value": source}},
        {"command": "set", "path": "/_数据接口/B17", "props": {"value": metrics.person_days, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B18", "props": {"value": metrics.total_cost, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B19", "props": {"value": metrics.average_day_rate, "type": "number"}},
        {"command": "set", "path": "/_数据接口/B20", "props": {"value": today, "type": "string"}},
        {"command": "set", "path": "/_数据接口/B21", "props": {"value": "已同步"}},
    ]
    officecli = shutil.which("officecli")
    if not officecli:
        raise RuntimeError("未找到 officecli，无法在保留图表的前提下更新 AI PMO")
    with tempfile.NamedTemporaryFile("w", suffix=".json", encoding="utf-8") as stream:
        json.dump(commands, stream, ensure_ascii=False)
        stream.flush()
        subprocess.run(
            [officecli, "batch", str(ai_pmo), "--input", stream.name, "--stop-on-error"],
            check=True,
        )
    return cost, ai_pmo, metrics


def main() -> None:
    default_suite = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description="将成本模块指标同步到 AI PMO 数据接口")
    parser.add_argument("--suite", type=Path, default=default_suite)
    args = parser.parse_args()
    cost, ai_pmo, metrics = sync(args.suite.resolve())
    print(f"已从 {cost.name} 同步至 {ai_pmo.name}")
    print(f"计划人天={metrics.person_days:g}，计划人工成本={metrics.total_cost:g}，平均人日单价={metrics.average_day_rate:.2f}")


if __name__ == "__main__":
    main()
