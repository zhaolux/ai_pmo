from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from ai_pmo.database import connect, initialize
from ai_pmo.finding_review import import_reviews, review_summary


def seed_run(database: Path) -> None:
    initialize(database, {"project_id": "P1", "project_name": "合同系统", "industry": "能源"})
    with connect(database) as connection:
        connection.execute(
            "INSERT INTO agent_run VALUES(?,?,?,?,?,?,?)",
            ("RUN-1", "P1", "2026-09-23", "2026-09-23T10:00:00", "黄", "READ_ONLY", "r.json"),
        )
        for finding_id in ("F-1", "F-2"):
            connection.execute(
                "INSERT INTO agent_finding VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (finding_id, "RUN-1", "test", finding_id, "中", "标题", "详情", "建议", "PM", 0, "[]"),
            )


def save_review(path: Path, rows: list[tuple[str, str, str]], run_id: str = "RUN-1") -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "发现清单"
    sheet.append([])
    sheet.append([])
    sheet.append([])
    sheet.append(["发现ID", "Agent", "严重度", "对象ID", "标题", "详细说明", "建议动作", "责任人", "需审批", "数据日期", "确认状态", "人工结论"])
    for finding_id, status, conclusion in rows:
        sheet.append([finding_id, "Agent", "中", "O", "标题", "", "", "PM", "否", "2026-09-23", status, conclusion])
    instructions = book.create_sheet("使用说明")
    instructions.append(["运行ID", run_id])
    book.save(path)


class FindingReviewTests(unittest.TestCase):
    def test_imports_non_pending_reviews_and_summarizes(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            database, workbook = root / "pmo.db", root / "review.xlsx"
            seed_run(database)
            save_review(workbook, [("F-1", "已接受", "按建议处理"), ("F-2", "待确认", "")])
            result = import_reviews(database, workbook, "项目经理")
            summary = review_summary(database, "RUN-1")
        self.assertEqual(result["imported"], 1)
        self.assertEqual(summary["已接受"], 1)
        self.assertEqual(summary["待确认"], 1)

    def test_unknown_finding_rejects_whole_import(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            database, workbook = root / "pmo.db", root / "review.xlsx"
            seed_run(database)
            save_review(workbook, [("F-1", "已接受", "有效"), ("F-X", "已关闭", "无效")])
            with self.assertRaisesRegex(ValueError, "不存在"):
                import_reviews(database, workbook, "项目经理")
            with connect(database) as connection:
                count = connection.execute("SELECT COUNT(*) FROM agent_finding_review").fetchone()[0]
        self.assertEqual(count, 0)

    def test_invalid_status_or_empty_conclusion_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            database, workbook = root / "pmo.db", root / "review.xlsx"
            seed_run(database)
            save_review(workbook, [("F-1", "已接受", "")])
            with self.assertRaisesRegex(ValueError, "人工结论"):
                import_reviews(database, workbook, "项目经理")


if __name__ == "__main__":
    unittest.main()
