from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

AUTOMATION = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(AUTOMATION / 'src'))

from ai_pmo.daily_run import execute_daily


def main() -> int:
    parser = argparse.ArgumentParser(
        description='AI PMO统一日常运行入口；默认只读预览，遇错即停',
    )
    parser.add_argument('--suite', type=Path, default=AUTOMATION.parent)
    parser.add_argument('--date', type=date.fromisoformat, default=date.today())
    parser.add_argument(
        '--generate-derived', action='store_true',
        help='显式生成Agent分析和周报派生文件；仍不写回业务工作簿',
    )
    args = parser.parse_args()
    return execute_daily(
        args.suite.resolve(), args.date,
        generate_derived=args.generate_derived,
    )


if __name__ == '__main__':
    raise SystemExit(main())
