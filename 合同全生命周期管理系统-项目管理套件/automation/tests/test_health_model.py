from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from ai_pmo.agent_models import AgentResult, Evidence, Finding
from ai_pmo.agents.schedule_resource import analyze_schedule_resource
from ai_pmo.database import connect, initialize
from ai_pmo.health_model import calculate_health
from ai_pmo.orchestrator import consolidate, persist_report


def finding(finding_id: str, severity: str, agent: str = "schedule_resource") -> Finding:
    return Finding(
        finding_id=finding_id, agent=agent, object_id="OBJ-1", severity=severity,
        title="事项", detail="详情", recommendation="处理", owner="项目经理",
        requires_approval=False, evidence=(Evidence("a.xlsx", "S", "OBJ-1", 4),),
    )


class HealthModelTests(unittest.TestCase):
    def test_weighted_dimensions_are_explainable(self):
        result = calculate_health([
            AgentResult("schedule_resource", "", (
                finding("SCH-1-OVERDUE", "高"),
                finding("RES-1-OVERLOAD", "中"),
            )),
        ])
        self.assertEqual(result["model_version"], "2026-09-23")
        self.assertLess(result["score"], 100)
        self.assertEqual(result["dimensions"]["progress"]["penalty"], 15)
        self.assertEqual(result["dimensions"]["resource"]["penalty"], 8)
        self.assertIn(result["state"], {"绿", "黄", "红"})

    def test_major_finding_is_red_hard_gate(self):
        result = calculate_health([
            AgentResult("risk_issue", "", (finding("RSK-1", "重大", "risk_issue"),)),
        ])
        self.assertEqual(result["state"], "红")

    def test_health_snapshot_is_persisted_with_dimensions(self):
        with tempfile.TemporaryDirectory() as folder:
            database = Path(folder) / "pmo.db"
            initialize(database, {"project_id": "P1", "project_name": "合同系统", "industry": "能源"})
            report = consolidate("P1", date(2026, 9, 23), [])
            persist_report(database, report, Path(folder) / "report.json")
            with connect(database) as connection:
                snapshot = connection.execute("SELECT score, state FROM health_snapshot").fetchone()
                dimensions = connection.execute("SELECT COUNT(*) FROM health_dimension_snapshot").fetchone()[0]
            self.assertEqual((snapshot["score"], snapshot["state"]), (100, "绿"))
            self.assertEqual(dimensions, 6)

    def test_current_week_overload_is_detected_from_formula_result(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "plan.xlsx"
            book = Workbook()
            control = book.active
            control.title = "项目总控台账"
            for _ in range(3):
                control.append([])
            resource = book.create_sheet("资源计划")
            for _ in range(38):
                resource.append([])
            row = [None] * 9
            row[0], row[1], row[2], row[5], row[6], row[7], row[8] = (
                "张三", date(2026, 9, 21), date(2026, 9, 27), 48, 1.2, "超负荷", "T001；T002",
            )
            resource.append(row)
            book.save(path)
            result = analyze_schedule_resource(path, date(2026, 9, 23))
        overload = [item for item in result.findings if item.finding_id.startswith("RES-")]
        self.assertEqual(len(overload), 1)
        self.assertEqual(overload[0].severity, "高")


if __name__ == "__main__":
    unittest.main()
