from __future__ import annotations

import sqlite3

from .agent_models import AgentResult, Finding


def _load_rule_history(connection: sqlite3.Connection) -> dict[tuple[str, str], tuple[int, str, str]]:
    """每个规则（agent+title）在既往运行中出现的期数与最早/最晚数据日期。"""
    history: dict[tuple[str, str], tuple[int, str, str]] = {}
    for row in connection.execute(
        """SELECT f.agent, f.title, COUNT(DISTINCT f.run_id), MIN(r.data_date), MAX(r.data_date)
           FROM agent_finding f JOIN agent_run r ON f.run_id = r.run_id
           GROUP BY f.agent, f.title"""
    ):
        history[(row[0], row[1])] = (row[2], row[3], row[4])
    return history


def build_context(
    finding: Finding,
    current: list[Finding],
    history: dict[tuple[str, str], tuple[int, str, str]],
) -> dict:
    """单条发现的数据上下文：本期同类规模、责任人压力、对象多维暴露、历史期数。"""
    rule_count = sum(1 for item in current if item.agent == finding.agent and item.title == finding.title)
    owner_load = sum(1 for item in current if item.owner == finding.owner)
    object_hits = sum(
        1 for item in current
        if item.object_id == finding.object_id and item.finding_id != finding.finding_id
    )
    runs, first, last = history.get((finding.agent, finding.title), (0, "", ""))
    return {
        "rule_count": rule_count,
        "owner_load": owner_load,
        "object_hits": object_hits,
        "history_runs": runs,
        "history_first": first,
        "history_last": last,
    }


def annotate_results(
    results: list[AgentResult], connection: sqlite3.Connection | None = None,
) -> list[AgentResult]:
    """为各 Agent 发现附加数据上下文包，供解释层与大模型使用。"""
    history = _load_rule_history(connection) if connection is not None else {}
    current = [finding for result in results for finding in result.findings]
    annotated: list[AgentResult] = []
    for result in results:
        findings = tuple(
            Finding(
                finding_id=finding.finding_id, agent=finding.agent, object_id=finding.object_id,
                severity=finding.severity, title=finding.title, detail=finding.detail,
                recommendation=finding.recommendation, owner=finding.owner,
                requires_approval=finding.requires_approval, evidence=finding.evidence,
                age_runs=finding.age_runs, review_state=finding.review_state,
                context=build_context(finding, current, history),
            )
            for finding in result.findings
        )
        annotated.append(AgentResult(result.agent, result.summary, findings))
    return annotated
