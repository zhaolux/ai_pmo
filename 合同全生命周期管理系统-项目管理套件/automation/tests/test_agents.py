from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

from ai_pmo.agent_models import AgentResult, Evidence, Finding
from ai_pmo.agents.risk_issue import analyze_risk_issue
from ai_pmo.agents.schedule_resource import analyze_schedule_resource
from ai_pmo.orchestrator import consolidate
from ai_pmo.database import connect, initialize


def save_plan(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "项目总控台账"
    sheet.append([""] * 30)
    sheet.append([""] * 30)
    sheet.append([""] * 30)
    overdue = [None] * 30
    overdue[0], overdue[4], overdue[6], overdue[7] = "T001", "逾期任务", "进行中", "项目经理"
    overdue[8], overdue[9], overdue[29] = datetime(2026, 9, 1), datetime(2026, 9, 10), ""
    sheet.append(overdue)
    due_soon = [None] * 30
    due_soon[0], due_soon[4], due_soon[6], due_soon[7] = "T002", "近期任务", "未开始", "产品经理"
    due_soon[8], due_soon[9] = datetime(2026, 9, 18), datetime(2026, 9, 25)
    sheet.append(due_soon)
    book.save(path)


def save_risk(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "风险登记册"
    for _ in range(8):
        sheet.append([])
    row = [None] * 44
    row[0], row[1], row[22], row[28], row[33] = "RSK-001", "重大风险", "技术经理", "处理中", "重大"
    sheet.append(row)
    book.save(path)


def save_change(path: Path) -> None:
    book = Workbook()
    issue = book.active
    issue.title = "问题台账"
    for _ in range(3):
        issue.append([])
    issue.append(["ISS-001", "问题", "描述", "高", "处理中", "项目经理"])
    change = book.create_sheet("变更台账")
    for _ in range(3):
        change.append([])
    change.append(["CHG-001", datetime(2026, 9, 18), "项目经理", "范围变更", "原因", "高", "待审批"])
    decision = book.create_sheet("待决策事项")
    for _ in range(4):
        decision.append([])
    row = [None] * 13
    row[0], row[1], row[4], row[8] = "DEC-001", "架构决策", "项目经理", "待决策"
    decision.append(row)
    book.save(path)


class AgentTests(unittest.TestCase):
    def test_finding_serializes_with_evidence_and_approval(self):
        finding = Finding(
            finding_id="F-001", agent="schedule_resource", object_id="T001",
            severity="高", title="任务延期", detail="计划完成日已过",
            recommendation="确认状态", owner="项目经理", requires_approval=True,
            evidence=(Evidence("plan.xlsx", "项目总控台账", "T001", 4),),
        )
        data = finding.to_dict()
        self.assertEqual(data["evidence"][0]["record_id"], "T001")
        self.assertTrue(data["requires_approval"])
        json.dumps(data, ensure_ascii=False)

    def test_agents_find_schedule_risk_issue_change_and_decision(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            plan, risk, change = root / "plan.xlsx", root / "risk.xlsx", root / "change.xlsx"
            save_plan(plan)
            save_risk(risk)
            save_change(change)
            schedule = analyze_schedule_resource(plan, date(2026, 9, 18))
            risk_result = analyze_risk_issue(risk, change, date(2026, 9, 18))
            self.assertEqual({f.object_id for f in schedule.findings}, {"T001", "T002"})
            self.assertEqual(
                {f.object_id for f in risk_result.findings},
                {"RSK-001", "ISS-001", "CHG-001", "DEC-001"},
            )

    def test_controller_returns_red_health_and_read_only_boundary(self):
        report = consolidate("ECM-2026", date(2026, 9, 18), [])
        self.assertEqual(report["health"], "绿")
        self.assertEqual(report["write_policy"], "READ_ONLY_RECOMMENDATIONS")
        self.assertEqual(report["management_summary"]["status_judgment"], "绿")
        self.assertEqual(report["management_summary"]["finding_count"], 0)

    def test_management_summary_counts_all_approvals(self):
        findings = tuple(
            Finding(
                finding_id=f"F-{index}", agent="test", object_id=f"O-{index}",
                severity="中", title="待审批", detail="详情", recommendation="审批",
                owner="项目经理", requires_approval=True,
                evidence=(Evidence("source.xlsx", "Sheet1", f"O-{index}", index),),
            )
            for index in range(11)
        )
        report = consolidate(
            "ECM-2026", date(2026, 9, 18),
            [AgentResult("test", "11项待审批", findings)],
        )
        self.assertEqual(report["management_summary"]["approval_count"], 11)
        self.assertEqual(len(report["management_summary"]["decisions_needed"]), 10)

    def test_database_contains_agent_audit_tables(self):
        with tempfile.TemporaryDirectory() as folder:
            database = Path(folder) / "pmo.db"
            initialize(database, {"project_id": "ECM-2026", "project_name": "合同系统", "industry": "能源"})
            with connect(database) as connection:
                tables = {row[0] for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )}
            self.assertIn("agent_run", tables)
            self.assertIn("agent_finding", tables)


if __name__ == "__main__":
    unittest.main()
