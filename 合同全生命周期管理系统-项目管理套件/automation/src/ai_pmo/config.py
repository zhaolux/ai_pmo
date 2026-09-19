from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Paths:
    automation: Path
    suite: Path
    config: Path
    inputs: Path
    outputs: Path
    data_center: Path
    database: Path
    logs: Path


def get_paths() -> Paths:
    automation = Path(__file__).resolve().parents[2]
    suite = automation.parent
    return Paths(
        automation=automation,
        suite=suite,
        config=automation / "config",
        inputs=suite / "inputs",
        outputs=suite / "outputs",
        data_center=suite / "90_项目数据中心",
        database=suite / "90_项目数据中心" / "ai_pmo.db",
        logs=automation / "logs",
    )


def load_json(name: str) -> dict:
    path = get_paths().config / name
    return json.loads(path.read_text(encoding="utf-8"))

