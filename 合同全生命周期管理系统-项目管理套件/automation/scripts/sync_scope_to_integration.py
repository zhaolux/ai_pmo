#!/usr/bin/env python3
"""03 需求与范围管理 ↔ 04 系统集成管理 接口快照双向同步。

默认仅预览差异；确认后添加 --apply 执行。执行前备份两本工作簿。
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
from ai_pmo.scope_sync import apply_sync, plan_sync


def digest(path: Path) -> bytes:
    return hashlib.sha256(path.read_bytes()).digest()


def backup_workbook(target: Path) -> Path:
    archive = target.parent / "_archive" / "自动刷新前备份"
    archive.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = archive / f"{target.stem}-{timestamp}-{uuid4().hex[:8]}{target.suffix}"
    shutil.copy2(target, backup)
    if digest(target) != digest(backup):
        raise OSError(f"备份校验失败：{backup}")
    return backup


def main() -> int:
    parser = argparse.ArgumentParser(
        description="按列归属双向同步 03 接口范围与 04 执行快照（默认仅预览）",
    )
    parser.add_argument("--suite", type=Path, default=AUTOMATION.parent)
    parser.add_argument("--apply", action="store_true", help="执行同步；缺省只预览差异")
    args = parser.parse_args()
    suite = args.suite.resolve()

    scope = resolve_current(suite, "scope")
    integration = resolve_current(suite, "integration")

    if not args.apply:
        report = plan_sync(scope, integration)
        print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
        print("仅预览，未修改工作簿。执行同步请添加 --apply。")
        return 1 if any(p.errors for p in report.plans) else 0

    backups = [backup_workbook(scope), backup_workbook(integration)]
    report = apply_sync(scope, integration)
    result = report.to_dict()
    result["backups"] = [str(path.relative_to(suite)) for path in backups]
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if any(p.errors for p in report.plans):
        print("同步后复核存在错误，请检查备份并人工处理。", file=sys.stderr)
        return 1
    print("同步完成。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
