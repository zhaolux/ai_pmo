from __future__ import annotations

import json
import os
from urllib.request import Request, urlopen
from datetime import date
from pathlib import Path

BOUNDARY = "AI解释仅为建议，不构成审批结论，不修改任何源台账"
SEVERITY_RANK = {"重大": 0, "高": 1, "中": 2, "低": 3}
DEFAULT_LIMIT = 5


def select_explainable(findings: list[dict], limit: int = DEFAULT_LIMIT) -> list[dict]:
    """优先解释需审批发现，再按严重度排序，同档按 finding_id 稳定取前 limit 条。"""
    ranked = sorted(
        findings,
        key=lambda item: (
            0 if item.get("requires_approval") else 1,
            SEVERITY_RANK.get(item.get("severity", ""), 9),
            item.get("finding_id", ""),
        ),
    )
    return ranked[:limit]


def build_prompt(finding: dict) -> str:
    """把已验证事实与数据上下文组装为解释层提示词，要求只基于事实、输出三段建议。"""
    context = finding.get("context") or {}
    evidence_lines = "\n".join(
        f"- {item.get('file')} / {item.get('sheet')} / 记录{item.get('record_id')}"
        for item in finding.get("evidence", [])
    )
    return (
        "你是项目管理顾问。请基于下列已核实的项目事实与数据上下文，"
        "为该关注事项给出解释。只许基于给定事实推理，不得臆造数据或结论。\n\n"
        f"【事实】{finding.get('title', '')}（{finding.get('severity', '')}）\n"
        f"【明细】{finding.get('detail', '')}\n"
        f"【责任人】{finding.get('owner', '')}\n"
        f"【证据】\n{evidence_lines or '- 无'}\n"
        "【上下文】"
        f"同规则本期命中{context.get('rule_count', 0)}项；"
        f"责任人本期共{context.get('owner_load', 0)}项；"
        f"该对象被命中{context.get('object_hits', 0)}次；"
        f"该规则历史上出现{context.get('history_runs', 0)}期"
        f"（{context.get('history_first', '')}至{context.get('history_last', '')}）。\n\n"
        "请严格按以下三段输出，每段不超过150字：\n"
        "原因分析：…\n情景建议：…\n风险提示：…\n"
        "注意：你的输出是建议，供项目经理参考，不构成审批结论。"
    )


def load_llm_config(env: dict | None = None) -> dict:
    source = os.environ if env is None else env
    return {
        "api_key": source.get("AI_PMO_LLM_API_KEY", ""),
        "base_url": source.get("AI_PMO_LLM_BASE_URL", "https://api.moonshot.cn/v1").rstrip("/"),
        "model": source.get("AI_PMO_LLM_MODEL", "moonshot-v1-8k"),
    }


class OpenAICompatibleClient:
    """OpenAI 兼容 chat/completions 客户端（仅标准库，不引入新依赖）。"""

    def __init__(self, base_url: str, api_key: str, model: str, timeout: int = 60):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.model = model
        self.timeout = timeout

    def complete(self, prompt: str) -> str:
        request = Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps({
                "model": self.model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
            }).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}",
            },
        )
        with urlopen(request, timeout=self.timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return payload["choices"][0]["message"]["content"].strip()


def explain_report(report: dict, client, *, limit: int = DEFAULT_LIMIT) -> dict:
    """对报告中发现生成解释条目；无客户端时全部标记 skipped_no_key，单条出错不中断。"""
    selected = select_explainable(report.get("findings", []), limit=limit)
    entries = []
    for finding in selected:
        entry = {
            "finding_id": finding.get("finding_id", ""),
            "title": finding.get("title", ""),
            "severity": finding.get("severity", ""),
            "status": "ok",
            "explanation": "",
            "error": "",
        }
        if client is None:
            entry["status"] = "skipped_no_key"
        else:
            try:
                entry["explanation"] = client.complete(build_prompt(finding))
            except Exception as exc:  # 模型服务不可用时降级为错误标注，不影响其余条目
                entry["status"] = "error"
                entry["error"] = str(exc)
        entries.append(entry)
    return {
        "source_report": "",
        "data_date": report.get("data_date", ""),
        "model": getattr(client, "model", "") if client is not None else "",
        "boundary": BOUNDARY,
        "explanations": entries,
    }


def write_explanations(report_path: Path, output_path: Path, client=None, *, limit: int = DEFAULT_LIMIT) -> dict:
    """读取Agent报告并写出解释JSON；client为None时按环境变量构建，无密钥则全部跳过。"""
    report = json.loads(report_path.read_text(encoding="utf-8"))
    if client is None:
        config = load_llm_config()
        if config["api_key"]:
            client = OpenAICompatibleClient(config["base_url"], config["api_key"], config["model"])
    payload = explain_report(report, client, limit=limit)
    payload["source_report"] = report_path.name
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload
