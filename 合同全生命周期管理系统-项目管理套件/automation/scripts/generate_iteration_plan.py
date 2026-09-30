#!/usr/bin/env python3
"""按项目总控台账"迭代周"列生成周迭代计划（默认仅预览；--apply 执行并备份）。

联动口径：明细区以任务ID为键，其余列 INDEX/MATCH 查总控台账，总控更新自动跟随。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

AUTOMATION = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(AUTOMATION / "src"))

from ai_pmo.current_workbooks import resolve_current
from ai_pmo.iteration_plan import apply_generation, plan_generation


def backup_workbook(target: Path) -> Path:
    import hashlib
    import shutil
    from datetime import datetime
    from uuid import uuid4
    archive = target.parent / "_archive" / "自动刷新前备份"
    archive.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = archive / f"{target.stem}-{timestamp}-{uuid4().hex[:8]}{target.suffix}"
    shutil.copy2(target, backup)
    if hashlib.sha256(target.read_bytes()).digest() != hashlib.sha256(backup.read_bytes()).digest():
        raise OSError(f"备份校验失败：{backup}")
    return backup


def main() -> int:
    parser = argparse.ArgumentParser(
        description="按总控台账迭代周列生成周迭代计划（默认仅预览；--apply 执行并备份）",
    )
    parser.add_argument("--suite", type=Path, default=AUTOMATION.parent)
    parser.add_argument("--apply", action="store_true", help="执行生成；缺省只预览")
    args = parser.parse_args()
    suite = args.suite.resolve()

    plan = resolve_current(suite, "plan")

    if not args.apply:
        report = plan_generation(plan)
        print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
        print("仅预览，未修改工作簿。执行生成请添加 --apply。")
        return 1 if report.errors else 0

    backup = backup_workbook(plan)
    report = apply_generation(plan)
    result = report.to_dict()
    result["backup"] = str(backup.relative_to(suite))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if report.errors:
        return 1
    print(f"已生成周迭代计划：{report.detail_count} 条任务明细 / {len(report.weeks)} 个迭代周"
          f"（备份：{result['backup']}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
