#!/usr/bin/env python3
"""资源计划周任务清单与项目总控台账联动刷新（默认仅预览；--apply 执行并备份）。

任务清单（人员→任务→周分配）保持不变，按总控当前名称/状态刷新清单文本、
任务数、周工时与 master 汇总；引用总控不存在的任务ID时拒绝写入。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path
from uuid import uuid4

AUTOMATION = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(AUTOMATION / "src"))

from ai_pmo.current_workbooks import resolve_current
from ai_pmo.resource_week_sync import apply_resource_sync, plan_resource_sync


def backup_workbook(target: Path) -> Path:
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
        description="资源计划周任务清单与总控台账联动刷新（默认仅预览；--apply 执行并备份）",
    )
    parser.add_argument("--suite", type=Path, default=AUTOMATION.parent)
    parser.add_argument("--apply", action="store_true", help="执行刷新；缺省只预览")
    args = parser.parse_args()
    suite = args.suite.resolve()

    plan = resolve_current(suite, "plan")

    if not args.apply:
        report = plan_resource_sync(plan)
        print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
        print("仅预览，未修改工作簿。执行刷新请添加 --apply。")
        return 1 if report.errors else 0

    backup = backup_workbook(plan)
    report = apply_resource_sync(plan)
    result = report.to_dict()
    result["backup"] = str(backup.relative_to(suite))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if report.errors:
        return 1
    print(f"已刷新资源计划：{report.rows_refreshed}/{report.rows_scanned} 行周任务清单、"
          f"{len(report.persons)} 名人员 master 汇总（备份：{result['backup']}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
