from __future__ import annotations

import os
import subprocess
from datetime import date
from pathlib import Path
from typing import Callable


DEFAULT_NODE = Path(
    "/Users/zhaolu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"
)


def agent_excel_path(suite: Path, as_of: date) -> Path:
    return (
        suite
        / "11_AI_PMO中心"
        / f"AI PMO-Agent运行结果-{as_of:%Y%m%d}-V1.xlsx"
    )


def build_agent_excel(
    report_path: Path,
    output_path: Path,
    script_path: Path,
    *,
    node_bin: Path | None = None,
    runner: Callable[..., object] = subprocess.run,
) -> Path:
    node = node_bin or Path(os.environ.get("AI_PMO_NODE", str(DEFAULT_NODE)))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    runner(
        [str(node), str(script_path), str(report_path), str(output_path)],
        check=True,
    )
    if not output_path.exists():
        raise RuntimeError(f"Agent结果Excel未生成：{output_path}")
    return output_path
