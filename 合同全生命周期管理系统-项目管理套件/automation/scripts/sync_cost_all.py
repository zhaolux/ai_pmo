#!/usr/bin/env python3
"""成本联动统一入口：一次预览/执行同时完成驾驶舱指标与 G15 材料口径。

聚合两个既有命令（各自保留独立可用）：
- sync_cost_to_ai_pmo：AI PMO 项目驾驶舱 B9:D9 与 _数据接口 B16:B21
- sync_cost_caliber_to_g15：G15 会前签认跟踪与两份 G15 md 的口径数字

默认仅预览；--apply 先跑驾驶舱同步（officecli，备份后写入），
再跑 G15 口径联动（zipfile/md，备份后写入）。两段均带目标防切换守卫。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from scripts import sync_cost_caliber_to_g15 as g15_sync
from scripts import sync_cost_to_ai_pmo as cost_sync


def preview_sync(suite: Path) -> dict[str, object]:
    dashboard = cost_sync.preview_sync(suite)
    g15 = g15_sync.preview_sync(suite)
    return {
        "source": dashboard["source"],
        "metrics": dashboard["metrics"],
        "dashboard": dashboard,
        "g15": g15,
    }


def sync(suite: Path, *, preview: dict[str, object] | None = None):
    if preview is None:
        preview = preview_sync(suite)
    result_dashboard = cost_sync.sync(
        suite, expected_target=preview["dashboard"]["target"]
    )
    backups_g15 = g15_sync.sync(
        suite, expected_source=preview["g15"]["source"]
    )
    return result_dashboard, backups_g15


def main() -> None:
    default_suite = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(
        description="成本联动统一入口：驾驶舱指标 + G15 材料口径一次预览/执行"
    )
    parser.add_argument("--suite", type=Path, default=default_suite)
    parser.add_argument("--apply", action="store_true", help="备份目标后执行两段联动；默认仅预览")
    args = parser.parse_args()
    suite = args.suite.resolve()
    preview = preview_sync(suite)
    m = preview["metrics"]
    g15 = g15_sync
    print(f"读取成本工作簿：{preview['source']}")
    print(
        f"计划人天={g15.fmt_person_days(m.person_days)}，"
        f"人工成本={g15.fmt_wan(m.total_cost)} 万元，"
        f"平均人日单价={g15.fmt_rate(m.average_day_rate)} 元"
    )
    print(f"[1] 驾驶舱：将更新 项目驾驶舱 B9:D9、_数据接口 B16:B21（{preview['dashboard']['target'].relative_to(suite)}）")
    plan = preview["g15"]["plan"]
    if plan:
        for item in plan:
            print(f"[2] G15：将更新 {item['file'].relative_to(suite)}（命中 {item['replacements']} 处）")
    else:
        print("[2] G15：未找到会前材料目标，跳过")
    if not args.apply:
        print("仅预览，未修改任何文件。执行两段联动请添加 --apply。")
        return
    result_dashboard, backups_g15 = sync(suite, preview=preview)
    cost, ai_pmo, metrics, backup = result_dashboard
    print(f"[1] 已同步至 {ai_pmo.name}，刷新前备份：{backup}")
    for backup_path in backups_g15:
        print(f"[2] G15 备份：{backup_path}")


if __name__ == "__main__":
    main()
