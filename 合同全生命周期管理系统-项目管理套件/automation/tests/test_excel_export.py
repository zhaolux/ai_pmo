from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from ai_pmo.excel_export import agent_excel_path, build_agent_excel


class AgentExcelExportTests(unittest.TestCase):
    def test_daily_result_path_is_stable(self):
        suite = Path("/tmp/project-suite")
        self.assertEqual(
            agent_excel_path(suite, date(2026, 9, 18)),
            suite / "11_AI_PMO中心" / "AI PMO-Agent运行结果-20260918-V1.xlsx",
        )

    def test_builder_receives_report_and_output_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            report = root / "report.json"
            output = root / "result.xlsx"
            script = root / "builder.mjs"
            report.write_text("{}", encoding="utf-8")
            script.write_text("", encoding="utf-8")
            calls: list[list[str]] = []

            def runner(command: list[str], check: bool) -> None:
                self.assertTrue(check)
                calls.append(command)
                output.write_bytes(b"xlsx")

            result = build_agent_excel(
                report, output, script, node_bin=Path("/opt/node"), runner=runner
            )

            self.assertEqual(result, output)
            self.assertEqual(
                calls,
                [["/opt/node", str(script), str(report), str(output)]],
            )


if __name__ == "__main__":
    unittest.main()
