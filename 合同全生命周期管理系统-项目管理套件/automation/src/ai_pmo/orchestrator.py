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


def newest(directory: Path, contains: str) -> Path:
    files = [p for p in directory.glob("*.xlsx") if contains in p.stem and not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"未找到工作簿：{directory}/{contains}*.xlsx")
    return max(files, key=lambda p: (p.stat().st_mtime, p.name))


def consolidate(project_id: str, as_of: date, results: list[AgentResult]) -> dict:
    findings = [finding for result in results for finding in result.findings]
    major = sum(item.severity == "重大" for item in findings)
    high = sum(item.severity == "高" for item in findings)
    health = "红" if major or high >= 3 else ("黄" if high or findings else "绿")
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


def run_agents(
    suite: Path, project_id: str, as_of: date, output_dir: Path,
    database: Path | None = None,
) -> Path:
    plan = newest(suite / "02_计划与进度管理", "计划与进度管理")
    risk = newest(suite / "06_风险问题与变更", "风险管理")
    change = newest(suite / "06_风险问题与变更", "问题与变更管理")
    cost = newest(suite / "05_成本与合同管理", "成本与合同管理")
    quality = newest(suite / "07_质量测试与验收", "质量测试与验收")
    communication = newest(suite / "08_沟通会议与报告", "沟通会议与报告")
    results = [
        analyze_schedule_resource(plan, as_of),
        analyze_risk_issue(risk, change, as_of),
        analyze_cost_contract(cost, as_of),
        analyze_quality_acceptance(quality, as_of),
        analyze_communication_report(communication, as_of),
    ]
    report = consolidate(project_id, as_of, results)
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"agent-report-{as_of:%Y%m%d}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if database is not None:
        persist_report(database, report, path)
    return path
