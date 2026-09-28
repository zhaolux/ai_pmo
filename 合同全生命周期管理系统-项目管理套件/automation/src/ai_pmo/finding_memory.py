from __future__ import annotations

import sqlite3

from .agent_models import AgentResult, Finding

ESCALATION_THRESHOLD = 3
ESCALATION_LADDER = {"低": "中", "中": "高", "高": "重大"}
ONGOING_STATUSES = {"已接受", "需调整"}


def _load_streaks(connection: sqlite3.Connection) -> dict[str, int]:
    """每个发现ID在此之前的连续出现期数（紧邻最近的运行往回数）。"""
    runs = [row[0] for row in connection.execute(
        "SELECT run_id FROM agent_run ORDER BY generated_at, run_id"
    )]
    if not runs:
        return {}
    presence: dict[str, set[str]] = {}
    for row in connection.execute("SELECT finding_id, run_id FROM agent_finding"):
        presence.setdefault(row[0], set()).add(row[1])
    streaks: dict[str, int] = {}
    for finding_id, seen in presence.items():
        streak = 0
        for run_id in reversed(runs):
            if run_id not in seen:
                break
            streak += 1
        streaks[finding_id] = streak
    return streaks


def _load_latest_reviews(connection: sqlite3.Connection) -> dict[str, str]:
    """每个发现ID最新一次人工复核状态（按复核时间）。"""
    latest: dict[str, str] = {}
    for row in connection.execute(
        """SELECT r.finding_id, r.status, r.reviewed_at
           FROM agent_finding_review r
           ORDER BY r.reviewed_at"""
    ):
        latest[row[0]] = row[1]
    return latest


def annotate_finding(finding: Finding, streak: int, review_status: str | None) -> Finding:
    age_runs = streak + 1
    review_state = ""
    if review_status in ONGOING_STATUSES:
        review_state = "处置中"
    elif review_status == "已关闭":
        review_state = "已关闭后复现"

    if review_state == "处置中":
        detail = f"【处置中·第{age_runs}期】{finding.detail}"
    elif review_state == "已关闭后复现":
        detail = f"【已确认关闭仍复现·第{age_runs}期】{finding.detail}"
    elif age_runs >= 2:
        detail = f"【连续第{age_runs}期出现】{finding.detail}"
    else:
        detail = finding.detail

    severity = finding.severity
    if age_runs >= ESCALATION_THRESHOLD and review_state != "处置中":
        severity = ESCALATION_LADDER.get(severity, severity)

    return Finding(
        finding_id=finding.finding_id, agent=finding.agent, object_id=finding.object_id,
        severity=severity, title=finding.title, detail=detail,
        recommendation=finding.recommendation, owner=finding.owner,
        requires_approval=finding.requires_approval, evidence=finding.evidence,
        age_runs=age_runs, review_state=review_state,
    )


def annotate_results(results: list[AgentResult], connection: sqlite3.Connection) -> list[AgentResult]:
    """为各 Agent 发现标注跨期记忆：持续期数、处置状态、久拖升级。"""
    streaks = _load_streaks(connection)
    reviews = _load_latest_reviews(connection)
    annotated: list[AgentResult] = []
    for result in results:
        findings = tuple(
            annotate_finding(finding, streaks.get(finding.finding_id, 0), reviews.get(finding.finding_id))
            for finding in result.findings
        )
        annotated.append(AgentResult(result.agent, result.summary, findings))
    return annotated
