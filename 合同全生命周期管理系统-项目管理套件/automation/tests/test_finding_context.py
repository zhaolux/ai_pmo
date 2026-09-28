from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path

from ai_pmo.agent_models import AgentResult, Evidence, Finding
from ai_pmo.database import connect, initialize
from ai_pmo.finding_context import annotate_results


def make_finding(finding_id: str, agent: str, title: str, owner: str, object_id: str) -> Finding:
    return Finding(
        finding_id=finding_id, agent=agent, object_id=object_id,
        severity="中", title=title, detail="详情", recommendation="建议",
        owner=owner, requires_approval=False,
        evidence=(Evidence("source.xlsx", "Sheet1", object_id, 4),),
    )


def make_result(agent: str, findings: list[Finding]) -> AgentResult:
    return AgentResult(agent, f"{len(findings)}项", tuple(findings))


class CurrentRunContextTests(unittest.TestCase):
    def test_current_run_counts_rule_owner_and_object(self):
        findings = [
            make_finding("F-1", "schedule_resource", "任务已延期", "项目经理", "T001"),
            make_finding("F-2", "schedule_resource", "任务已延期", "项目经理", "T002"),
            make_finding("F-3", "schedule_resource", "延期任务未关联风险登记", "项目经理", "T001"),
            make_finding("F-4", "risk_issue", "未关闭问题", "产品经理", "O2"),
        ]
        results = annotate_results([make_result("schedule_resource", findings[:3]), make_result("risk_issue", findings[3:])])

        by_id = {f.finding_id: f for r in results for f in r.findings}
        self.assertEqual(by_id["F-1"].context["rule_count"], 2)
        self.assertEqual(by_id["F-1"].context["owner_load"], 3)
        self.assertEqual(by_id["F-1"].context["object_hits"], 1)
        self.assertEqual(by_id["F-2"].context["object_hits"], 0)
        self.assertEqual(by_id["F-4"].context["rule_count"], 1)
        self.assertEqual(by_id["F-4"].context["owner_load"], 1)
        self.assertEqual(by_id["F-4"].context["object_hits"], 0)
        self.assertEqual(by_id["F-1"].context["history_runs"], 0)
        self.assertEqual(by_id["F-1"].context["history_first"], "")

    def test_context_serializes_to_json(self):
        findings = [make_finding("F-1", "schedule_resource", "任务已延期", "项目经理", "T001")]
        results = annotate_results([make_result("schedule_resource", findings)])
        data = results[0].findings[0].to_dict()
        self.assertEqual(data["context"]["rule_count"], 1)
        json.dumps(data, ensure_ascii=False)


class HistoryContextTests(unittest.TestCase):
    def save_database(self, folder: str) -> Path:
        database = Path(folder) / "pmo.db"
        initialize(database, {"project_id": "P1", "project_name": "合同系统", "industry": "能源"})
        with connect(database) as connection:
            for run_id, data_date in (("R1", "2026-09-20"), ("R2", "2026-09-21")):
                connection.execute(
                    "INSERT INTO agent_run (run_id, project_id, data_date, generated_at, health, write_policy)"
                    " VALUES (?, 'P1', ?, ?, '绿', 'READ_ONLY_RECOMMENDATIONS')",
                    (run_id, data_date, f"{data_date}T09:00:00"),
                )
                connection.execute(
                    "INSERT INTO agent_finding (finding_id, run_id, agent, object_id, severity, title, detail,"
                    " recommendation, owner, requires_approval, evidence_json)"
                    " VALUES (?, ?, 'schedule_resource', 'T001', '高', '任务已延期', '详情', '建议', '项目经理', 1, '[]')",
                    (f"SCH-{run_id}", run_id),
                )
            connection.execute(
                "INSERT INTO agent_finding (finding_id, run_id, agent, object_id, severity, title, detail,"
                " recommendation, owner, requires_approval, evidence_json)"
                " VALUES ('SCH-OLD', 'R1', 'schedule_resource', 'T009', '高', '任务已延期', '详情', '建议', '项目经理', 1, '[]')"
            )
            connection.commit()
        return database

    def test_history_context_counts_runs_and_dates(self):
        with tempfile.TemporaryDirectory() as folder:
            database = self.save_database(folder)
            findings = [make_finding("F-NEW", "schedule_resource", "任务已延期", "项目经理", "T003")]
            with connect(database) as connection:
                results = annotate_results([make_result("schedule_resource", findings)], connection)
            context = results[0].findings[0].context
            self.assertEqual(context["history_runs"], 2)
            self.assertEqual(context["history_first"], "2026-09-20")
            self.assertEqual(context["history_last"], "2026-09-21")

    def test_unknown_rule_has_empty_history(self):
        with tempfile.TemporaryDirectory() as folder:
            database = self.save_database(folder)
            findings = [make_finding("F-NEW", "risk_issue", "未关闭问题", "产品经理", "O2")]
            with connect(database) as connection:
                results = annotate_results([make_result("risk_issue", findings)], connection)
            context = results[0].findings[0].context
            self.assertEqual(context["history_runs"], 0)
            self.assertEqual(context["history_first"], "")
            self.assertEqual(context["history_last"], "")


if __name__ == "__main__":
    unittest.main()
