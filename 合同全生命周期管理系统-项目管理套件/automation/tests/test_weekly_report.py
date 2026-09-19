from __future__ import annotations

import unittest
from datetime import date
from pathlib import Path

from ai_pmo.weekly_report import weekly_report_paths, weekly_report_payload


class WeeklyReportTests(unittest.TestCase):
    def test_payload_contains_management_judgment_and_decisions(self):
        report = {
            "data_date": "2026-09-18",
            "health": "红",
            "management_summary": {
                "finding_count": 66, "major_count": 4, "high_count": 50,
                "medium_count": 12, "approval_count": 37,
                "top_priorities": [{
                    "severity": "重大", "title": "重大风险", "object_id": "RSK-001",
                    "owner": "项目经理", "recommendation": "完成应对",
                }],
                "decisions_needed": [{
                    "severity": "中", "title": "POC选型", "object_id": "DEC-001",
                    "owner": "治理委员会",
                }],
                "next_action": "先处理重大事项。",
            },
            "agent_results": [{"agent": "communication_report", "finding_count": 3, "summary": "3项决策"}],
            "findings": [{
                "severity": "中", "title": "决策事项14天内到期", "object_id": "DEC-001",
                "owner": "治理委员会", "recommendation": "提交决策材料", "agent": "communication_report",
            }],
        }

        payload = weekly_report_payload(report, "能源行业合同全生命周期管理系统")

        self.assertEqual(payload["status"], "红")
        self.assertEqual(payload["metrics"]["finding_count"], 66)
        self.assertEqual(payload["decisions"][0]["object_id"], "DEC-001")

    def test_payload_prefers_real_decision_findings_over_delayed_tasks(self):
        report = {
            "data_date": "2026-09-18", "health": "红", "agent_results": [],
            "management_summary": {
                "finding_count": 2, "major_count": 0, "high_count": 1,
                "medium_count": 1, "approval_count": 2, "top_priorities": [],
                "decisions_needed": [{"severity": "高", "title": "任务已延期", "object_id": "T001", "owner": "项目经理"}],
                "next_action": "推进决策。",
            },
            "findings": [{
                "severity": "中", "title": "决策事项14天内到期", "object_id": "DEC-001",
                "owner": "治理委员会", "recommendation": "提交决策材料", "agent": "communication_report",
            }],
        }
        payload = weekly_report_payload(report, "合同系统")
        self.assertEqual([item["object_id"] for item in payload["decisions"]], ["DEC-001"])

    def test_daily_paths_are_created_under_communication_directory(self):
        suite = Path("/tmp/suite")
        xlsx, docx = weekly_report_paths(suite, date(2026, 9, 18))
        expected = suite / "08_沟通会议与报告" / "周报"
        self.assertEqual(xlsx, expected / "项目周报-20260918-V1.xlsx")
        self.assertEqual(docx, expected / "项目周报-20260918-V1.docx")


if __name__ == "__main__":
    unittest.main()
