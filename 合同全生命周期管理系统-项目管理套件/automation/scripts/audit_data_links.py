from __future__ import annotations

import argparse
import json
from pathlib import Path

from ai_pmo.current_workbooks import resolve_current
from ai_pmo.link_audit import audit_suite_links


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="只读核对接口快照、实施任务与总控/资源关联")
    parser.add_argument("--suite", type=Path, default=Path(__file__).resolve().parents[2])
    args = parser.parse_args()
    suite = args.suite.resolve()
    result = audit_suite_links(
        resolve_current(suite, "scope"),
        resolve_current(suite, "integration"),
        resolve_current(suite, "plan"),
    )
    result["sources"] = {
        key: str(resolve_current(suite, key).relative_to(suite))
        for key in ("scope", "integration", "plan")
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if result["errors"] else 0)
