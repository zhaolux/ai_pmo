from __future__ import annotations

import json
from pathlib import Path


WORKBOOKS = {
    "gate": ("01_项目启动与治理", "项目启动与治理"),
    "plan": ("02_计划与进度管理", "计划与进度管理"),
    "scope": ("03_需求与范围管理", "需求与范围管理"),
    "integration": ("04_系统集成管理", "系统集成管理"),
    "cost": ("05_成本与合同管理", "成本与合同管理"),
    "risk": ("06_风险问题与变更", "风险管理"),
    "change": ("06_风险问题与变更", "问题与变更管理"),
    "quality": ("07_质量测试与验收", "质量测试与验收"),
    "deliverable": ("07_质量测试与验收", "验收交付物"),
    "communication": ("08_沟通会议与报告", "沟通会议与报告"),
    "migration": ("09_上线移交与运营", "旧合同系统迁移管理"),
    "ai_pmo": ("11_AI_PMO中心", "AI PMO中心"),
}


def resolve_current(suite: Path, key: str) -> Path:
    """Return an explicitly registered business workbook, never an mtime fallback."""
    if key not in WORKBOOKS:
        raise ValueError(f"未知工作簿类型：{key}")
    registry = suite / "automation" / "config" / "current_workbooks.json"
    if not registry.is_file():
        raise FileNotFoundError(f"缺少当前版本登记表：{registry}")
    entries = json.loads(registry.read_text(encoding="utf-8"))
    relative = entries.get(key) if isinstance(entries, dict) else None
    if not isinstance(relative, str) or not relative:
        raise ValueError(f"当前版本登记表缺少有效条目：{key}")
    directory, prefix = WORKBOOKS[key]
    candidate = Path(relative)
    if candidate.is_absolute() or ".." in candidate.parts or "_archive" in candidate.parts:
        raise ValueError(f"当前版本登记路径越界或指向归档：{key}={relative}")
    if candidate.parent != Path(directory) or not candidate.name.startswith(f"{prefix}-") or candidate.suffix != ".xlsx" or candidate.name.startswith("~$"):
        raise ValueError(f"当前版本登记文件名或目录不匹配：{key}={relative}")
    resolved = (suite / candidate).resolve()
    if resolved.parent != (suite / directory).resolve():
        raise ValueError(f"当前版本登记路径指向预期目录之外：{key}={relative}")
    if not resolved.is_file():
        raise FileNotFoundError(f"当前版本登记文件不存在：{key}={resolved}")
    return suite / candidate
