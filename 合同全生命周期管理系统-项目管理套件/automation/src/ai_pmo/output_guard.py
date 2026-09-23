from __future__ import annotations

from pathlib import Path
from typing import Iterable


def ensure_outputs_available(
    paths: Iterable[Path], *, replace_generated: bool = False,
) -> None:
    """Reject collisions before a command starts producing its output set."""
    if replace_generated:
        return
    existing = [path for path in paths if path.exists()]
    if existing:
        names = ", ".join(str(path) for path in existing)
        raise FileExistsError(f"目标文件已存在，未覆盖：{names}")
