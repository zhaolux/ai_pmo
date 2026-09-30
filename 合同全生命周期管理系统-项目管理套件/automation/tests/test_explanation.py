from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from ai_pmo.explanation import (
    OpenAICompatibleClient, build_prompt, explain_report, load_llm_config,
    select_explainable, write_explanations,
)


def make_finding(finding_id: str, severity: str, requires_approval: bool) -> dict:
    return {
        "finding_id": finding_id, "agent": "schedule_resource", "object_id": "T001",
        "severity": severity, "title": "任务已延期",
        "detail": "计划完成日为2026-09-10。", "recommendation": "确认纠偏计划。",
        "owner": "项目经理", "requires_approval": requires_approval,
        "evidence": [{"file": "plan.xlsx", "sheet": "项目总控台账", "record_id": "T001", "row": 8}],
        "context": {"rule_count": 3, "owner_load": 5, "object_hits": 1,
                    "history_runs": 4, "history_first": "2026-09-01", "history_last": "2026-09-20"},
    }


class FakeClient:
    model = "fake-model"

    def __init__(self, fail: bool = False):
        self.fail = fail
        self.prompts: list[str] = []

    def complete(self, prompt: str) -> str:
        self.prompts.append(prompt)
        if self.fail:
            raise RuntimeError("模型服务不可用")
        return "原因分析：示例。情景建议：示例。风险提示：示例。"


class SelectionTests(unittest.TestCase):
    def test_select_prioritizes_approval_then_severity_with_cap(self):
        findings = [
            make_finding("F-低", "低", False),
            make_finding("F-中高", "中", True),
            make_finding("F-高", "高", False),
            make_finding("F-重大", "重大", True),
            make_finding("F-中", "中", False),
        ]
        selected = select_explainable(findings, limit=3)
        self.assertEqual([f["finding_id"] for f in selected], ["F-重大", "F-中高", "F-高"])

    def test_select_ties_break_by_finding_id(self):
        findings = [make_finding("F-B", "中", False), make_finding("F-A", "中", False)]
        selected = select_explainable(findings, limit=1)
        self.assertEqual(selected[0]["finding_id"], "F-A")


class PromptTests(unittest.TestCase):
    def test_prompt_contains_facts_context_and_boundary(self):
        prompt = build_prompt(make_finding("F-1", "高", True))
        self.assertIn("任务已延期", prompt)
        self.assertIn("项目经理", prompt)
        self.assertIn("plan.xlsx", prompt)
        self.assertIn("同规则本期命中3项", prompt)
        self.assertIn("历史上出现4期", prompt)
        self.assertIn("2026-09-01", prompt)
        self.assertIn("建议", prompt)
        self.assertIn("不得臆造", prompt)


class ExplainReportTests(unittest.TestCase):
    def report(self) -> dict:
        return {
            "project_id": "P1", "data_date": "2026-09-25",
            "findings": [make_finding("F-1", "重大", True), make_finding("F-2", "低", False)],
        }

    def test_explain_with_client_marks_ok_and_records_model(self):
        client = FakeClient()
        payload = explain_report(self.report(), client, limit=5)
        self.assertEqual(len(payload["explanations"]), 2)
        entry = payload["explanations"][0]
        self.assertEqual(entry["finding_id"], "F-1")
        self.assertEqual(entry["status"], "ok")
        self.assertIn("原因分析", entry["explanation"])
        self.assertEqual(payload["model"], "fake-model")
        self.assertIn("不构成审批结论", payload["boundary"])

    def test_explain_without_client_marks_skipped(self):
        payload = explain_report(self.report(), None, limit=5)
        self.assertEqual(
            {entry["status"] for entry in payload["explanations"]},
            {"skipped_no_key"},
        )
        self.assertEqual(payload["model"], "")

    def test_client_error_marked_and_not_raised(self):
        client = FakeClient(fail=True)
        payload = explain_report(self.report(), client, limit=5)
        self.assertEqual(payload["explanations"][0]["status"], "error")
        self.assertIn("模型服务不可用", payload["explanations"][0]["error"])

    def test_write_explanations_roundtrip(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            report_path = root / "agent-report-20260925.json"
            report_path.write_text(json.dumps(self.report(), ensure_ascii=False), encoding="utf-8")
            output_path = root / "agent-explanations-20260925.json"
            payload = write_explanations(report_path, output_path, client=FakeClient())
            self.assertTrue(output_path.exists())
            loaded = json.loads(output_path.read_text(encoding="utf-8"))
            self.assertEqual(loaded["source_report"], "agent-report-20260925.json")
            self.assertEqual(loaded["data_date"], "2026-09-25")
            self.assertEqual(len(loaded["explanations"]), len(payload["explanations"]))


class LlmConfigTests(unittest.TestCase):
    def test_load_llm_config_from_env(self):
        env = {"AI_PMO_LLM_API_KEY": "k", "AI_PMO_LLM_BASE_URL": "https://x/v1", "AI_PMO_LLM_MODEL": "m"}
        config = load_llm_config(env)
        self.assertEqual(config["api_key"], "k")
        self.assertEqual(config["base_url"], "https://x/v1")
        self.assertEqual(config["model"], "m")

    def test_load_llm_config_defaults_and_missing_key(self):
        config = load_llm_config({})
        self.assertEqual(config["api_key"], "")
        self.assertTrue(config["base_url"].startswith("https://"))
        self.assertTrue(config["model"])

    def test_openai_compatible_client_posts_chat_completions(self):
        from unittest.mock import MagicMock
        client = OpenAICompatibleClient("https://x/v1", "k", "m", timeout=5)
        with patch("ai_pmo.explanation.urlopen") as urlopen:
            response = MagicMock()
            response.read.return_value = json.dumps(
                {"choices": [{"message": {"content": "解释文本"}}]}
            ).encode()
            response.__enter__.return_value = response
            urlopen.return_value = response
            text = client.complete("提示词")
        self.assertEqual(text, "解释文本")
        request = urlopen.call_args.args[0]
        self.assertEqual(request.full_url, "https://x/v1/chat/completions")
        body = json.loads(request.data.decode())
        self.assertEqual(body["model"], "m")
        self.assertEqual(body["messages"][0]["content"], "提示词")
        self.assertNotIn("temperature", body)


class CliDispatchTests(unittest.TestCase):
    def test_cli_dispatches_explain(self):
        from ai_pmo.cli import execute
        paths = SimpleNamespace(
            suite=Path("/tmp/suite"),
            outputs=Path("/tmp/suite/outputs"),
            logs=Path("/tmp/suite/logs"),
            database=Path("/tmp/suite/pmo.db"),
            automation=Path("/tmp/suite/automation"),
        )
        report_dir = paths.outputs / "agents"
        report_dir.mkdir(parents=True, exist_ok=True)
        (report_dir / "agent-report-20260925.json").write_text(
            json.dumps({"data_date": "2026-09-25", "findings": []}, ensure_ascii=False),
            encoding="utf-8",
        )
        with patch("ai_pmo.cli.get_paths", return_value=paths), \
            patch("ai_pmo.cli.load_json", return_value={"project_id": "P1", "project_name": "合同系统"}), \
            patch("ai_pmo.cli._log"), \
            patch("ai_pmo.cli.write_explanations", return_value={"explanations": [{"status": "ok"}]}) as build:
            code = execute("explain", date(2026, 9, 25), replace_generated=True)
        self.assertEqual(code, 0)
        args = build.call_args
        self.assertEqual(args.args[0].name, "agent-report-20260925.json")
        self.assertEqual(args.args[1].name, "agent-explanations-20260925.json")
        self.assertNotIn("client", args.kwargs)


if __name__ == "__main__":
    unittest.main()
