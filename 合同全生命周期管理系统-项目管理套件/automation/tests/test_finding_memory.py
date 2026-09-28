from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from ai_pmo.database import connect, initialize
from ai_pmo.agent_models import AgentResult, Evidence, Finding
from ai_pmo.finding_memory import annotate_results


def make_finding(finding_id: str, severity: str = "高") -> Finding:
    return Finding(
        finding_id=finding_id, agent="schedule_resource", object_id="T001",
        severity=severity, title="任务已延期", detail="计划完成日已过。",
        recommendation="确认实际状态。", owner="项目经理", requires_approval=True,
        evidence=(Evidence("plan.xlsx", "项目总控台账", "T001", 4),),
    )


def make_result(*findings: Finding) -> AgentResult:
    return AgentResult("schedule_resource", "摘要", tuple(findings))


def make_database(root: Path) -> Path:
    database = root / "ai_pmo.db"
    initialize(database, {"project_id": "ECM-2026", "project_name": "示例", "industry": "能源"})
    return database


def add_run(database: Path, run_id: str, generated_at: str, findings: dict[str, str]) -> None:
    with connect(database) as connection:
        connection.execute(
            """INSERT INTO agent_run
               (run_id, project_id, data_date, generated_at, health, write_policy, report_path)
               VALUES(?,?,?,?,?,?,?)""",
            (run_id, "ECM-2026", generated_at[:10], generated_at, "红", "READ_ONLY_RECOMMENDATIONS", ""),
        )
        for finding_id, severity in findings.items():
            connection.execute(
                """INSERT INTO agent_finding
                   (finding_id, run_id, agent, object_id, severity, title, detail,
                    recommendation, owner, requires_approval, evidence_json)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (finding_id, run_id, "schedule_resource", "T001", severity,
                 "任务已延期", "历史", "确认实际状态。", "项目经理", 1, "[]"),
            )


def add_review(database: Path, run_id: str, finding_id: str, status: str) -> None:
    with connect(database) as connection:
        connection.execute(
            """INSERT INTO agent_finding_review
               (run_id, finding_id, status, conclusion, reviewer, reviewed_at, source_path)
               VALUES(?,?,?,?,?,?,?)""",
            (run_id, finding_id, status, "人工结论", "项目经理",
             "2026-09-20T00:00:00+00:00", "review.xlsx"),
        )


class FindingMemoryTests(unittest.TestCase):
    def test_first_run_without_history_stays_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            result = make_result(make_finding("F1"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.severity, "高")
            self.assertEqual(finding.age_runs, 1)
            self.assertEqual(finding.review_state, "")
            self.assertEqual(finding.detail, "计划完成日已过。")

    def test_escalates_when_persisting_three_runs(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            add_run(database, "AGENT-20260923", "2026-09-23T09:00:00", {"F1": "中"})
            add_run(database, "AGENT-20260924", "2026-09-24T09:00:00", {"F1": "中"})
            result = make_result(make_finding("F1", severity="高"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.age_runs, 3)
            self.assertEqual(finding.severity, "重大")
            self.assertTrue(finding.detail.startswith("【连续第3期出现】"))

    def test_streak_resets_after_absence(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            add_run(database, "AGENT-20260922", "2026-09-22T09:00:00", {"F1": "高"})
            add_run(database, "AGENT-20260923", "2026-09-23T09:00:00", {"F9": "高"})
            add_run(database, "AGENT-20260924", "2026-09-24T09:00:00", {"F1": "高"})
            result = make_result(make_finding("F1"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.age_runs, 2)
            self.assertEqual(finding.severity, "高")
            self.assertTrue(finding.detail.startswith("【连续第2期出现】"))

    def test_accepted_review_marks_ongoing_without_escalation(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            add_run(database, "AGENT-20260923", "2026-09-23T09:00:00", {"F1": "中"})
            add_run(database, "AGENT-20260924", "2026-09-24T09:00:00", {"F1": "中"})
            add_run(database, "AGENT-20260925", "2026-09-25T09:00:00", {"F1": "中"})
            add_review(database, "AGENT-20260925", "F1", "已接受")
            result = make_result(make_finding("F1", severity="高"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.review_state, "处置中")
            self.assertEqual(finding.severity, "高")
            self.assertTrue(finding.detail.startswith("【处置中·第4期】"))

    def test_closed_review_reappearing_is_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            add_run(database, "AGENT-20260924", "2026-09-24T09:00:00", {"F1": "高"})
            add_run(database, "AGENT-20260925", "2026-09-25T09:00:00", {"F1": "高"})
            add_review(database, "AGENT-20260925", "F1", "已关闭")
            result = make_result(make_finding("F1", severity="高"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.review_state, "已关闭后复现")
            self.assertEqual(finding.severity, "重大")
            self.assertTrue(finding.detail.startswith("【已确认关闭仍复现·第3期】"))

    def test_ongoing_review_uses_latest_status(self):
        with tempfile.TemporaryDirectory() as folder:
            database = make_database(Path(folder))
            add_run(database, "AGENT-20260924", "2026-09-24T09:00:00", {"F1": "高"})
            add_run(database, "AGENT-20260925", "2026-09-25T09:00:00", {"F1": "高"})
            add_review(database, "AGENT-20260924", "F1", "已接受")
            add_review(database, "AGENT-20260925", "F1", "需调整")
            result = make_result(make_finding("F1", severity="高"))
            with connect(database) as connection:
                annotated = annotate_results([result], connection)
            finding = annotated[0].findings[0]
            self.assertEqual(finding.review_state, "处置中")
            self.assertEqual(finding.severity, "高")


if __name__ == "__main__":
    unittest.main()
