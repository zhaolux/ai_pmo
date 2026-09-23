from __future__ import annotations

import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

from ai_pmo import builder
from ai_pmo.cli import execute
from ai_pmo.decision_brief import build_decision_brief, decision_brief_paths
from ai_pmo.output_guard import ensure_outputs_available
from ai_pmo.orchestrator import run_agents
from ai_pmo.weekly_report import build_weekly_reports, weekly_report_paths


class OutputGuardTests(unittest.TestCase):
    def test_agents_target_preview_is_read_only_and_marks_existing_output(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            existing = suite / "11_AI_PMO中心" / "AI PMO-Agent运行结果-20260918-V1.xlsx"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"reviewed")
            paths = SimpleNamespace(suite=suite, outputs=suite / "outputs", database=suite / "pmo.db")
            output = StringIO()
            with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                "ai_pmo.cli.load_json", return_value={"project_id": "P1"}
            ), patch("ai_pmo.cli.initialize") as initialize, patch(
                "ai_pmo.cli._log"
            ) as log, redirect_stdout(output):
                self.assertEqual(execute("agents", date(2026, 9, 18), preview_targets=True), 0)
                initialize.assert_not_called()
                log.assert_not_called()
            self.assertIn("已存在", output.getvalue())
            self.assertIn(str(existing), output.getvalue())
            self.assertIn("agent-report-20260918.json", output.getvalue())
            self.assertEqual(existing.read_bytes(), b"reviewed")
            self.assertFalse(paths.outputs.exists())

    def test_preview_reflects_explicit_generated_replacement_flag(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            existing = suite / "11_AI_PMO中心" / "AI PMO-Agent运行结果-20260918-V1.xlsx"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"reviewed")
            paths = SimpleNamespace(suite=suite, outputs=suite / "outputs", database=suite / "pmo.db")
            output = StringIO()
            with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                "ai_pmo.cli.load_json", return_value={"project_id": "P1"}
            ), redirect_stdout(output):
                self.assertEqual(execute(
                    "agents", date(2026, 9, 18),
                    replace_generated=True, preview_targets=True,
                ), 0)
            self.assertIn("已存在，显式允许替换", output.getvalue())
            self.assertEqual(existing.read_bytes(), b"reviewed")

    def test_build_target_preview_lists_all_workbooks_without_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            paths = SimpleNamespace(suite=suite, outputs=suite / "outputs", database=suite / "pmo.db")
            target = suite / "01_治理" / "治理-V1.xlsx"
            output = StringIO()
            with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                "ai_pmo.cli.load_json", return_value={"project_id": "P1"}
            ), patch("ai_pmo.cli.build_output_paths", return_value={"01": target}), patch(
                "ai_pmo.cli.initialize"
            ) as initialize, redirect_stdout(output):
                self.assertEqual(execute("run", preview_targets=True), 0)
                initialize.assert_not_called()
            self.assertIn(str(target), output.getvalue())
            self.assertIn("将新建", output.getvalue())
            self.assertFalse(target.exists())

    def test_weekly_and_decision_previews_list_both_targets_without_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            paths = SimpleNamespace(suite=suite, outputs=suite / "outputs", database=suite / "pmo.db")
            for command, targets in (
                ("weekly-report", weekly_report_paths(suite, date(2026, 9, 18))),
                ("decision-brief", decision_brief_paths(suite, date(2026, 9, 18))),
            ):
                output = StringIO()
                with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                    "ai_pmo.cli.load_json", return_value={"project_id": "P1"}
                ), patch("ai_pmo.cli.initialize") as initialize, patch(
                    "ai_pmo.cli._log"
                ) as log, redirect_stdout(output):
                    self.assertEqual(execute(command, date(2026, 9, 18), preview_targets=True), 0)
                    initialize.assert_not_called()
                    log.assert_not_called()
                for target in targets:
                    self.assertIn(str(target), output.getvalue())
                    self.assertFalse(target.exists())
            self.assertFalse(paths.outputs.exists())

    def test_rejects_any_existing_target_before_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            existing = root / "report.xlsx"
            existing.write_bytes(b"human-edited")
            missing = root / "report.docx"

            with self.assertRaisesRegex(FileExistsError, "report.xlsx"):
                ensure_outputs_available((existing, missing))

            self.assertEqual(existing.read_bytes(), b"human-edited")
            self.assertFalse(missing.exists())

    def test_replace_generated_must_be_explicit(self):
        with tempfile.TemporaryDirectory() as folder:
            existing = Path(folder) / "report.json"
            existing.write_text("old", encoding="utf-8")

            ensure_outputs_available((existing,), replace_generated=True)

    def test_build_rejects_all_targets_before_first_workbook_is_saved(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            reference = suite / "inputs" / "参考资料"
            reference.mkdir(parents=True)
            template = Workbook()
            template.active.title = "SheetA"
            template.save(reference / "template.xlsx")
            today = date.today().strftime("%Y%m%d")
            existing = suite / "01_治理" / f"治理-{today}-V1.xlsx"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"human-edited")
            paths = SimpleNamespace(suite=suite, inputs=suite / "inputs")
            config = {
                "project.json": {"template_filename": "template.xlsx", "project_id": "P1", "project_name": "项目"},
                "directories.json": {"00": "00_说明", "01": "01_治理"},
                "sheet_mapping.json": {"00": ["SheetA"], "01": ["SheetA"]},
            }

            with patch.object(builder, "MODULE_NAMES", {"00": "说明", "01": "治理"}), patch.object(
                builder, "load_json", side_effect=config.__getitem__
            ):
                with self.assertRaises(FileExistsError):
                    builder.build(paths)

            self.assertEqual(existing.read_bytes(), b"human-edited")
            self.assertFalse((suite / "00_说明" / f"说明-{today}-V1.xlsx").exists())

    def test_agents_rejects_existing_json_before_reading_sources(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            output = suite / "outputs" / "agents"
            output.mkdir(parents=True)
            existing = output / "agent-report-20260918.json"
            existing.write_bytes(b"reviewed")

            with self.assertRaises(FileExistsError):
                run_agents(suite, "P1", date(2026, 9, 18), output)

            self.assertEqual(existing.read_bytes(), b"reviewed")

    def test_weekly_report_rejects_existing_file_before_reading_report(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            existing, _ = weekly_report_paths(suite, date(2026, 9, 18))
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"reviewed")

            with self.assertRaises(FileExistsError):
                build_weekly_reports(suite / "missing.json", "项目", suite, date(2026, 9, 18), suite / "missing.mjs")

            self.assertEqual(existing.read_bytes(), b"reviewed")

    def test_decision_brief_rejects_existing_file_before_reading_sources(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            existing, _ = decision_brief_paths(suite, date(2026, 9, 18))
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"reviewed")

            with self.assertRaises(FileExistsError):
                build_decision_brief(suite, [], date(2026, 9, 18))

            self.assertEqual(existing.read_bytes(), b"reviewed")

    def test_agents_command_preflights_json_and_excel_together(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            excel = suite / "11_AI_PMO中心" / "AI PMO-Agent运行结果-20260918-V1.xlsx"
            excel.parent.mkdir(parents=True)
            excel.write_bytes(b"human-edited")
            paths = SimpleNamespace(
                suite=suite, outputs=suite / "outputs", database=suite / "pmo.db",
                automation=suite / "automation",
            )
            with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                "ai_pmo.cli.load_json", return_value={"project_id": "P1", "project_name": "项目"}
            ), patch("ai_pmo.cli.initialize"), patch("ai_pmo.cli._log"), patch(
                "ai_pmo.cli.run_agents"
            ) as run:
                self.assertEqual(execute("agents", date(2026, 9, 18)), 1)
                run.assert_not_called()
            self.assertEqual(excel.read_bytes(), b"human-edited")

    def test_run_command_rejects_workbook_collision_before_database_init(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            existing = suite / "02_计划与进度管理" / "计划与进度管理-20260920-V1.xlsx"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"human-edited")
            paths = SimpleNamespace(suite=suite, database=suite / "pmo.db")
            with patch("ai_pmo.cli.get_paths", return_value=paths), patch(
                "ai_pmo.cli.load_json", return_value={"project_id": "P1", "project_name": "项目"}
            ), patch("ai_pmo.cli.build_output_paths", return_value={"02": existing}), patch(
                "ai_pmo.cli.initialize"
            ) as initialize, patch("ai_pmo.cli._log"):
                self.assertEqual(execute("run"), 1)
                initialize.assert_not_called()

            self.assertEqual(existing.read_bytes(), b"human-edited")


if __name__ == "__main__":
    unittest.main()
