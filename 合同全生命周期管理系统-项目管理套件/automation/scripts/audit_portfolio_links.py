from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

AUTOMATION = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(AUTOMATION / 'src'))

from ai_pmo.current_workbooks import resolve_current
from ai_pmo.portfolio_audit import audit_portfolio


def main() -> int:
    parser = argparse.ArgumentParser(
        description='只读核对里程碑、风险、问题、决策、质量证据、交付物、成本与驾驶舱快照')
    parser.add_argument('--suite', type=Path, default=AUTOMATION.parent)
    args = parser.parse_args()
    suite = args.suite.resolve()
    paths = {key: resolve_current(suite, key) for key in
             ('plan', 'risk', 'change', 'quality', 'deliverable', 'cost', 'ai_pmo', 'migration')}
    result = audit_portfolio(
        paths['plan'], paths['risk'], paths['change'], paths['quality'],
        paths['deliverable'], paths['cost'], paths['ai_pmo'],
        {key: paths[key].relative_to(suite).as_posix() for key in ('plan', 'risk')},
        migration_path=paths['migration'],
    )
    result['sources'] = {key: path.relative_to(suite).as_posix() for key, path in paths.items()}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result['errors'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
