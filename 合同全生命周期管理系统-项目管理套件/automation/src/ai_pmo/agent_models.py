from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(frozen=True)
class Evidence:
    file: str
    sheet: str
    record_id: str
    row: int | None = None

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(frozen=True)
class Finding:
    finding_id: str
    agent: str
    object_id: str
    severity: str
    title: str
    detail: str
    recommendation: str
    owner: str
    requires_approval: bool
    evidence: tuple[Evidence, ...]

    def to_dict(self) -> dict:
        data = asdict(self)
        data["evidence"] = [item.to_dict() for item in self.evidence]
        return data


@dataclass(frozen=True)
class AgentResult:
    agent: str
    summary: str
    findings: tuple[Finding, ...]

    def to_dict(self) -> dict:
        return {
            "agent": self.agent,
            "summary": self.summary,
            "finding_count": len(self.findings),
            "findings": [item.to_dict() for item in self.findings],
        }
