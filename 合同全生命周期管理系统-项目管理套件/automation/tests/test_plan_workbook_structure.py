"""项目总控台账结构护栏：总体状态公式回算全覆盖、预测完成日期列、基线偏差天数列。

针对真实登记的 02 计划工作簿；环境缺失时跳过，不构造合成替身——
本测试守护的是正式台账结构本身。
"""
from __future__ import annotations

import unittest
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ai_pmo.current_workbooks import resolve_current

SUITE = Path(__file__).resolve().parents[2]
DATA_START, DATA_END = 4, 138  # 项目总控台账数据行（含表头第 3 行）


def _load_plan():
    try:
        path = resolve_current(SUITE, "plan")
    except (FileNotFoundError, ValueError):
        return None
    return path if path.is_file() else None


@unittest.skipIf(_load_plan() is None, "02 计划工作簿未登记，跳过结构护栏")
class PlanWorkbookStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.path = _load_plan()
        cls.wb = load_workbook(cls.path)
        cls.ws = cls.wb["项目总控台账"]

    def test_g_column_status_formula_covers_all_data_rows(self):
        """总体状态（G 列）全部由四专业状态公式回算，不允许静态值残留。"""
        for r in range(DATA_START, DATA_END + 1):
            v = self.ws.cell(r, 7).value
            self.assertIsInstance(v, str, f"G{r} 应为公式字符串")
            self.assertTrue(v.startswith("=IF(OR(M"), f"G{r} 不是回算公式: {v[:60]}")
            for token in ('"暂停"', '"未分配"', '"已完成"', '"进行中"', '"待确认"', '"未开始"'):
                self.assertIn(token, v, f"G{r} 公式缺少分支 {token}")

    def test_g_column_has_no_dropdown_validation(self):
        """G 列已由公式接管，不应再挂人工下拉验证。"""
        for dv in self.ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                for row in range(DATA_START, DATA_END + 1):
                    self.assertFalse(
                        rng.min_col <= 7 <= rng.max_col
                        and rng.min_row <= row <= rng.max_row,
                        f"G 列仍被数据验证覆盖: {rng}",
                    )

    def test_forecast_finish_and_baseline_variance_columns(self):
        """AG=预测完成日期（人工维护）、AH=基线偏差天数（公式）。"""
        self.assertEqual(self.ws.cell(3, 33).value, "预测完成日期")
        self.assertEqual(self.ws.cell(3, 34).value, "基线偏差天数")
        for r in range(DATA_START, DATA_END + 1):
            ah = self.ws.cell(r, 34).value
            self.assertIsInstance(ah, str, f"AH{r} 应为公式字符串")
            self.assertIn(f"AG{r}", ah)
            self.assertIn(f"J{r}", ah)

    def test_g_formula_results_cached_and_sane(self):
        """LO 重算后的缓存值必须与四专业状态口径一致（抽查已知行）。"""
        wbv = load_workbook(self.path, data_only=True)
        wsv = wbv["项目总控台账"]
        g4 = wsv.cell(4, 7).value
        if g4 is None:  # 尚无公式缓存（未重算），不误判失败
            self.skipTest("G 列无公式缓存，请先运行 LO 重算")
        self.assertEqual(g4, "已完成")
        allowed = {"已完成", "进行中", "未开始", "待确认", "暂停", "未分配", ""}
        for r in range(DATA_START, DATA_END + 1):
            v = wsv.cell(r, 7).value
            self.assertIn(v, allowed, f"G{r} 缓存值异常: {v!r}")

    def test_filter_view_mirrors_new_columns(self):
        """总控筛选视图同步 AG/AH 联动列，且自动筛选范围覆盖到 AH。"""
        vw = self.wb["总控筛选视图"]
        self.assertEqual(vw.cell(3, 33).value, "预测完成日期")
        self.assertEqual(vw.cell(3, 34).value, "基线偏差天数")
        self.assertIn("项目总控台账!AG4", vw.cell(4, 33).value)
        self.assertIn("项目总控台账!AH4", vw.cell(4, 34).value)
        self.assertEqual(vw.auto_filter.ref, f"A3:{get_column_letter(34)}{DATA_END}")
