from __future__ import annotations

import argparse
import json
from pathlib import Path

from ai_pmo.source_ownership import audit_source_ownership


def main() -> None:
    parser = argparse.ArgumentParser(description="审计AI PMO对象的唯一业务主源")
    parser.add_argument("--suite", type=Path, required=True)
    args = parser.parse_args()
    result = audit_source_ownership(args.suite.resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if result["errors"] else 0)


if __name__ == "__main__":
    main()
