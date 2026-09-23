from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from .agent_models import AgentResult
from .agents.risk_issue import analyze_risk_issue
from .agents.schedule_resource import analyze_schedule_resource
from .agents.cost_contract import analyze_cost_contract
from .agents.quality_acceptance import analyze_quality_acceptance
from .agents.communication_report import analyze_communication_report
from .database import connect
from .current_workbooks import resolve_current
from .output_guard import ensure_outputs_available
from .health_model import calculate_health, load_health_model


def consolidate(project_id: str, as_of: date, results: list[AgentResult], health_model: dict | None = None) -> dict:
    findings = [finding for result in results for finding in result.findings]
    major = sum(item.severity == "重大" for item in findings)
    high = sum(item.severity == "高" for item in findings)
    health_result = calculate_health(results, health_model)
    health = health_result["state"]
    severity_rank = {"重大": 0, "高": 1, "中": 2, "低": 3}
    priorities = sorted(findings, key=lambda item: (severity_rank.get(item.severity, 9), item.finding_id))[:5]
    approval_findings = [item for item in findings if item.requires_approval]
    decisions = approval_findings[:10]
    management_summary = {
        "status_judgment": health,
        "finding_count": len(findings),
        "major_count": major,
        "high_count": high,
        "medium_count": sum(item.severity == "中" for item in findings),
        "approval_count": len(approval_findings),
        "top_priorities": [
            {"severity": item.severity, "title": item.title, "owner": item.owner,
             "object_id": item.object_id, "recommendation": item.recommendation}
            for item in priorities
        ],
        "decisions_needed": [
            {"severity": item.severity, "title": item.title, "owner": item.owner,
             "object_id": item.object_id}
            for item in decisions
        ],
        "next_action": "先处理重大和高严重度事项，再由项目经理确认待审批结论。",
    }
    return {
        "run_id": f"AGENT-{as_of:%Y%m%d}-{datetime.now():%H%M%S}",
        "project_id": project_id,
        "data_date": as_of.isoformat(),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "health": health,
        "health_score": health_result["score"],
        "health_model_version": health_result["model_version"],
        "health_dimensions": health_result["dimensions"],
        "write_policy": "READ_ONLY_RECOMMENDATIONS",
        "approval_boundary": "项目基线、预算、风险等级、任务状态和关闭结论必须由项目经理或授权人确认。",
        "management_summary": management_summary,
        "agent_results": [result.to_dict() for result in results],
        "finding_count": len(findings),
        "findings": [finding.to_dict() for finding in findings],
    }


def persist_report(database: Path, report: dict, report_path: Path) -> None:
    with connect(database) as connection:
        connection.execute(
            """INSERT INTO agent_run
               (run_id, project_id, data_date, generated_at, health, write_policy, report_path)
               VALUES(?,?,?,?,?,?,?)""",
            (report["run_id"], report["project_id"], report["data_date"],
             report["generated_at"], report["health"], report["write_policy"], str(report_path)),
        )
        for finding in report["findings"]:
            connection.execute(
                """INSERT INTO agent_finding
                   (finding_id, run_id, agent, object_id, severity, title, detail,
                    recommendation, owner, requires_approval, evidence_json)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (finding["finding_id"], report["run_id"], finding["agent"], finding["object_id"],
                 finding["severity"], finding["title"], finding["detail"],
                 finding["recommendation"], finding["owner"], int(finding["requires_approval"]),
                json.dumps(finding["evidence"], ensure_ascii=False)),
            )
        connection.execute(
            """INSERT INTO health_snapshot
               (run_id, project_id, data_date, model_version, score, state)
               VALUES(?,?,?,?,?,?)""",
            (report["run_id"], report["project_id"], report["data_date"],
             report["health_model_version"], report["health_score"], report["health"]),
        )
        for dimension, item in report["health_dimensions"].items():
            connection.execute(
                """INSERT INTO health_dimension_snapshot
                   (run_id, dimension, label, weight, score, penalty, finding_count)
                   VALUES(?,?,?,?,?,?,?)""",
                (report["run_id"], dimension, item["label"], item["weight"], item["score"],
                 item["penalty"], item["finding_count"]),
            )


def run_agents(
    suite: Path, project_id: str, as_of: date, output_dir: Path,
    database: Path | None = None, *, replace_generated: bool = False,
) -> Path:
    path = output_dir / f"agent-report-{as_of:%Y%m%d}.json"
    ensure_outputs_available((path,), replace_generated=replace_generated)
    plan = resolve_current(suite, "plan")
    risk = resolve_current(suite, "risk")
    change = resolve_current(suite, "change")
    cost = resolve_current(suite, "cost")
    quality = resolve_current(suite, "quality")
    deliverable = resolve_current(suite, "deliverable")
    communication = resolve_current(suite, "communication")
    results = [
        analyze_schedule_resource(plan, as_of),
        analyze_risk_issue(risk, change, as_of),
        analyze_cost_contract(cost, as_of),
        analyze_quality_acceptance(quality, deliverable, as_of),
        analyze_communication_report(communication, as_of),
    ]
    health_config = suite / "automation" / "config" / "health_model.json"
    report = consolidate(project_id, as_of, results, load_health_model(health_config) if health_config.exists() else None)
    output_dir.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if database is not None:
        persist_report(database, report, path)
    return path
