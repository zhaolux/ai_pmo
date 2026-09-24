from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from scripts import sync_cost_all


def make_suite(root: Path) -> tuple[Path, Path]:
    cost_dir = root / "05_成本与合同管理"
    target_dir = root / "11_AI_PMO中心"
    g15_dir = root / "01_项目启动与治理"
    cost_dir.mkdir(parents=True)
    target_dir.mkdir(parents=True)
    g15_dir.mkdir(parents=True)
    cost = cost_dir / "成本与合同管理-20260918-V5.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "人工成本预算"
    sheet["A15"] = "合计"
    sheet["D15"] = 4721
    sheet["F15"] = 6525700
    budget = book.create_sheet("预算基线")
    budget["J5"] = 6525700
    budget["J6"] = 9000000
    book.save(cost)
    target = target_dir / "AI PMO中心-20260920-V1.xlsx"
    target.write_bytes(b"human-edited-workbook")
    registry = root / "automation" / "config" / "current_workbooks.json"
    registry.parent.mkdir(parents=True)
    registry.write_text(json.dumps({
        "cost": cost.relative_to(root).as_posix(),
        "ai_pmo": target.relative_to(root).as_posix(),
    }, ensure_ascii=False), encoding="utf-8")
    return cost, target


class CostAllPreviewTests(unittest.TestCase):
    def test_preview_aggregates_dashboard_and_g15(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            make_suite(root)
            preview = sync_cost_all.preview_sync(root)
            self.assertEqual(preview["metrics"].person_days, 4721)
            self.assertIn("dashboard", preview)
            self.assertIn("g15", preview)

    def test_preview_does_not_write_anything(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            _, target = make_suite(root)
            sync_cost_all.preview_sync(root)
            self.assertEqual(target.read_bytes(), b"human-edited-workbook")
            self.assertFalse(list(root.rglob("_archive")))


class CostAllApplyTests(unittest.TestCase):
    def test_apply_runs_both_syncs_with_guards(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            make_suite(root)
            calls: list[str] = []

            def fake_dashboard(suite, *, expected_target=None):
                calls.append("dashboard")
                return None

            def fake_g15(suite, *, expected_source=None):
                calls.append("g15")
                return []

            with patch.object(sync_cost_all.cost_sync, "sync", fake_dashboard), \
                 patch.object(sync_cost_all.g15_sync, "sync", fake_g15):
                preview = sync_cost_all.preview_sync(root)
                sync_cost_all.sync(root, preview=preview)

            self.assertEqual(calls, ["dashboard", "g15"])

    def test_apply_propagates_guard_failure_from_either_stage(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            make_suite(root)

            def guarded_dashboard(suite, *, expected_target=None):
                raise RuntimeError("驾驶舱目标在预览后已变化")

            calls: list[str] = []

            def fake_g15(suite, *, expected_source=None):
                calls.append("g15")
                return []

            with patch.object(sync_cost_all.cost_sync, "sync", guarded_dashboard), \
                 patch.object(sync_cost_all.g15_sync, "sync", fake_g15):
                preview = sync_cost_all.preview_sync(root)
                with self.assertRaises(RuntimeError):
                    sync_cost_all.sync(root, preview=preview)
            # 第一段守卫失败时第二段不应执行
            self.assertEqual(calls, [])


if __name__ == "__main__":
    unittest.main()
