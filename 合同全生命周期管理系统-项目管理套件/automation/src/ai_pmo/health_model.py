from __future__ import annotations

from pathlib import Path
import json

from .agent_models import AgentResult, Finding


DEFAULT_MODEL = {
    "version": "2026-09-23",
    "dimensions": {
        "progress": {"label": "进度", "weight": 25},
        "cost": {"label": "成本", "weight": 20},
        "risk_issue": {"label": "风险与问题", "weight": 20},
        "quality": {"label": "质量", "weight": 15},
        "resource": {"label": "资源", "weight": 10},
        "stakeholder": {"label": "相关方与沟通", "weight": 10},
    },
    "severity_penalties": {"重大": 25, "高": 15, "中": 8, "低": 3},
    "thresholds": {"green": 85, "yellow": 70},
    "major_is_red": True,
}


def load_health_model(path: Path | None = None) -> dict:
    if path is None:
        return DEFAULT_MODEL
    return json.loads(path.read_text(encoding="utf-8"))


def _dimension(finding: Finding) -> str:
    if finding.finding_id.startswith("RES-"):
        return "resource"
    if finding.finding_id.startswith("SCH-"):
        return "progress"
    return {
        "risk_issue": "risk_issue",
        "cost_contract": "cost",
        "quality_acceptance": "quality",
        "communication_report": "stakeholder",
    }.get(finding.agent, "stakeholder")


def calculate_health(results: list[AgentResult], model: dict | None = None) -> dict:
    model = model or DEFAULT_MODEL
    dimensions = {
        code: {
            "label": settings["label"], "weight": settings["weight"],
            "score": 100, "penalty": 0, "finding_count": 0,
        }
        for code, settings in model["dimensions"].items()
    }
    major = False
    for result in results:
        for finding in result.findings:
            code = _dimension(finding)
            if code not in dimensions:
                continue
            penalty = model["severity_penalties"].get(finding.severity, 0)
            dimensions[code]["penalty"] += penalty
            dimensions[code]["finding_count"] += 1
            major = major or finding.severity == "重大"
    for item in dimensions.values():
        item["score"] = max(0, 100 - item["penalty"])
    weight_total = sum(item["weight"] for item in dimensions.values()) or 1
    score = round(sum(item["score"] * item["weight"] for item in dimensions.values()) / weight_total, 1)
    if major and model.get("major_is_red", True):
        state = "红"
    elif score >= model["thresholds"]["green"]:
        state = "绿"
    elif score >= model["thresholds"]["yellow"]:
        state = "黄"
    else:
        state = "红"
    return {
        "model_version": model["version"], "score": score, "state": state,
        "dimensions": dimensions,
    }
