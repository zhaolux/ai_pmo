from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from ai_pmo.agent_models import AgentResult, Evidence, Finding
from ai_pmo.agents.risk_issue import analyze_risk_issue
from ai_pmo.agents.schedule_resource import analyze_schedule_resource
from ai_pmo.orchestrator import consolidate, run_agents
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
    def test_run_agents_uses_registered_sources_not_newer_workbooks(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            files = {
                "plan": ("02_计划与进度管理", "计划与进度管理"),
                "risk": ("06_风险问题与变更", "风险管理"),
                "change": ("06_风险问题与变更", "问题与变更管理"),
                "cost": ("05_成本与合同管理", "成本与合同管理"),
                "quality": ("07_质量测试与验收", "质量测试与验收"),
                "deliverable": ("07_质量测试与验收", "验收交付物"),
                "communication": ("08_沟通会议与报告", "沟通会议与报告"),
            }
            current = {}
            for key, (directory, prefix) in files.items():
                path = suite / directory / f"{prefix}-20260918-V1.xlsx"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(b"registered")
                (path.parent / f"{prefix}-20260921-V1.xlsx").write_bytes(b"unregistered")
                current[key] = path.relative_to(suite).as_posix()
            registry = suite / "automation" / "config" / "current_workbooks.json"
            registry.parent.mkdir(parents=True)
            registry.write_text(json.dumps(current, ensure_ascii=False), encoding="utf-8")
            empty = AgentResult("test", "无发现", ())
            with patch("ai_pmo.orchestrator.analyze_schedule_resource", return_value=empty) as schedule, patch(
                "ai_pmo.orchestrator.analyze_risk_issue", return_value=empty
            ) as risk_issue, patch("ai_pmo.orchestrator.analyze_cost_contract", return_value=empty) as cost, patch(
                "ai_pmo.orchestrator.analyze_quality_acceptance", return_value=empty
            ) as quality, patch("ai_pmo.orchestrator.analyze_communication_report", return_value=empty) as communication:
                run_agents(suite, "P1", date(2026, 9, 20), suite / "outputs")
            self.assertEqual(schedule.call_args.args[0].name, "计划与进度管理-20260918-V1.xlsx")
            self.assertEqual(risk_issue.call_args.args[0].name, "风险管理-20260918-V1.xlsx")
            self.assertEqual(cost.call_args.args[0].name, "成本与合同管理-20260918-V1.xlsx")
            self.assertEqual(quality.call_args.args[0].name, "质量测试与验收-20260918-V1.xlsx")
            self.assertEqual(quality.call_args.args[1].name, "验收交付物-20260918-V1.xlsx")
            self.assertEqual(communication.call_args.args[0].name, "沟通会议与报告-20260918-V1.xlsx")

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
            schedule = analyze_schedule_resource(plan, risk, date(2026, 9, 18))
            risk_result = analyze_risk_issue(risk, change, date(2026, 9, 18))
            self.assertEqual({f.object_id for f in schedule.findings}, {"T001", "T002"})
            self.assertEqual(
                {f.object_id for f in risk_result.findings},
                {"RSK-001", "ISS-001", "CHG-001", "DEC-001"},
            )

class DelayedTaskRiskCrossCheckTests(unittest.TestCase):
    def save_plan_overdue(self, path: Path) -> None:
        book = Workbook()
        sheet = book.active
        sheet.title = "项目总控台账"
        for _ in range(3):
            sheet.append([None] * 30)
        overdue = [None] * 30
        overdue[0], overdue[4], overdue[6], overdue[7] = "T099", "逾期任务", "进行中", "项目经理"
        overdue[8], overdue[9] = datetime(2026, 9, 1), datetime(2026, 9, 10)
        sheet.append(overdue)
        book.save(path)

    def save_risk_with_text(self, path: Path, text: str) -> None:
        book = Workbook()
        sheet = book.active
        sheet.title = "风险登记册"
        for _ in range(8):
            sheet.append([])
        row = [None] * 44
        row[0], row[1], row[22], row[28], row[33] = "RSK-100", text, "技术经理", "处理中", "中"
        sheet.append(row)
        book.save(path)

    def test_delayed_task_covered_by_risk_register_not_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            plan, risk = root / "plan.xlsx", root / "risk.xlsx"
            self.save_plan_overdue(plan)
            self.save_risk_with_text(risk, "逾期风险关联T099")
            result = analyze_schedule_resource(plan, risk, date(2026, 9, 18))
            self.assertNotIn(
                "SCH-T099-NORISK",
                {f.finding_id for f in result.findings},
            )

    def test_delayed_task_without_risk_coverage_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            plan, risk = root / "plan.xlsx", root / "risk.xlsx"
            self.save_plan_overdue(plan)
            self.save_risk_with_text(risk, "无关风险")
            result = analyze_schedule_resource(plan, risk, date(2026, 9, 18))
            finding = next(f for f in result.findings if f.finding_id == "SCH-T099-NORISK")
            self.assertEqual(finding.severity, "中")
            self.assertEqual(finding.title, "延期任务未关联风险登记")
            self.assertEqual(finding.owner, "项目经理")
            self.assertFalse(finding.requires_approval)
            self.assertIn("T099", finding.detail)


class RiskIssueGradingTests(unittest.TestCase):
    def save_risk_rows(self, path: Path, rows: list[dict]) -> None:
        book = Workbook()
        sheet = book.active
        sheet.title = "风险登记册"
        for _ in range(8):
            sheet.append([])
        for item in rows:
            row = [None] * 44
            row[0], row[1] = item["id"], item["name"]
            row[22], row[28], row[33] = item.get("owner", "技术经理"), item.get("status", "处理中"), item.get("level", "中")
            if item.get("response_due"):
                row[27] = item["response_due"]
            sheet.append(row)
        book.save(path)

    def save_decisions(self, path: Path, rows: list[dict]) -> None:
        book = Workbook()
        issue = book.active
        issue.title = "问题台账"
        change = book.create_sheet("变更台账")
        decision = book.create_sheet("待决策事项")
        for _ in range(4):
            decision.append([])
        for item in rows:
            row = [None] * 13
            row[0], row[2], row[6], row[8] = item["id"], item["topic"], item.get("decider", "指导委员会"), item.get("status", "待决策")
            if item.get("due"):
                row[7] = item["due"]
            decision.append(row)
        book.save(path)

    def test_decision_overdue_escalates_to_major(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [])
            self.save_decisions(change, [{"id": "DEC-009", "topic": "产品选型", "due": datetime(2026, 9, 20)}])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            finding = next(f for f in result.findings if f.object_id == "DEC-009")
            self.assertEqual(finding.severity, "重大")
            self.assertEqual(finding.title, "决策事项已逾期")
            self.assertEqual(finding.owner, "指导委员会")
            self.assertIn("产品选型", finding.detail)

    def test_decision_due_within_7_days_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [])
            self.save_decisions(change, [{"id": "DEC-010", "topic": "预算批准", "due": datetime(2026, 10, 2)}])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            finding = next(f for f in result.findings if f.object_id == "DEC-010")
            self.assertEqual(finding.severity, "高")
            self.assertEqual(finding.title, "决策事项7天内到期")

    def test_decision_without_date_keeps_generic_finding(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [])
            self.save_decisions(change, [{"id": "DEC-011", "topic": "范围确认"}])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            finding = next(f for f in result.findings if f.object_id == "DEC-011")
            self.assertEqual(finding.title, "待决策事项")
            self.assertEqual(finding.severity, "高")

    def test_decided_item_not_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [])
            self.save_decisions(change, [{"id": "DEC-012", "topic": "已决项", "due": datetime(2026, 9, 1), "status": "已决策"}])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            self.assertNotIn("DEC-012", {f.object_id for f in result.findings})

    def test_risk_response_overdue_flagged_regardless_of_level(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [{"id": "RSK-009", "name": "资源不足", "level": "中", "response_due": datetime(2026, 9, 10)}])
            self.save_decisions(change, [])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            finding = next(f for f in result.findings if f.finding_id == "RSK-RSK-009-RESPONSE")
            self.assertEqual(finding.severity, "高")
            self.assertEqual(finding.title, "风险应对已逾期")
            self.assertEqual(finding.owner, "技术经理")
            self.assertTrue(finding.requires_approval)

    def test_closed_risk_response_not_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            risk, change = root / "risk.xlsx", root / "change.xlsx"
            self.save_risk_rows(risk, [{"id": "RSK-010", "name": "已关闭风险", "level": "高", "status": "已关闭", "response_due": datetime(2026, 9, 1)}])
            self.save_decisions(change, [])
            result = analyze_risk_issue(risk, change, date(2026, 9, 28))
            self.assertNotIn("RSK-RSK-010-RESPONSE", {f.finding_id for f in result.findings})

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
