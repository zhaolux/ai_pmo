from __future__ import annotations

import subprocess
import sys
import json
import os
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .current_workbooks import WORKBOOKS, resolve_current


@dataclass(frozen=True)
class DailyStep:
    name: str
    command: tuple[str, ...]


def validate_daily_inputs(suite: Path) -> dict[str, Path]:
    resolved = {key: resolve_current(suite, key) for key in WORKBOOKS}
    project_path = suite / 'automation' / 'config' / 'project.json'
    project = json.loads(project_path.read_text(encoding='utf-8'))
    hours_per_day = project.get('hours_per_day')
    person_day_rate = project.get('person_day_rate')
    hourly_rate = project.get('hourly_rate')
    if not isinstance(hours_per_day, (int, float)) or hours_per_day <= 0:
        raise ValueError('每日工时必须大于0')
    if hourly_rate != person_day_rate / hours_per_day:
        raise ValueError('小时单价与人天单价不一致')
    if project.get('hours_per_week') != 40:
        raise ValueError('每人每周可用工时不是40小时')
    return resolved


def build_daily_steps(
    suite: Path,
    report_date: date,
    python_executable: str = sys.executable,
    *,
    generate_derived: bool = False,
) -> list[DailyStep]:
    automation = suite / 'automation'
    date_text = report_date.isoformat()
    agent_arguments = () if generate_derived else ('--preview-targets',)
    weekly_arguments = () if generate_derived else ('--preview-targets',)
    return [
        DailyStep('数据主源审计', (
            python_executable, str(automation / 'scripts' / 'audit_source_ownership.py'),
            '--suite', str(suite),
        )),
        DailyStep('数据联动审计', (
            python_executable, str(automation / 'scripts' / 'audit_data_links.py'),
            '--suite', str(suite),
        )),
        DailyStep('接口快照同步预览', (
            python_executable, str(automation / 'scripts' / 'sync_scope_to_integration.py'),
            '--suite', str(suite),
        )),
        DailyStep('项目组合联动审计', (
            python_executable, str(automation / 'scripts' / 'audit_portfolio_links.py'),
            '--suite', str(suite),
        )),
        DailyStep('AI PMO全量刷新预览', (
            python_executable, str(automation / 'scripts' / 'sync_all_to_ai_pmo.py'),
            '--suite', str(suite), '--date', date_text,
        )),
        DailyStep('成本刷新预览', (
            python_executable, str(automation / 'scripts' / 'sync_cost_to_ai_pmo.py'),
            '--suite', str(suite),
        )),
        DailyStep('Agent分析' if generate_derived else 'Agent输出预览', (
            python_executable, '-m', 'ai_pmo', 'agents', '--date', date_text,
            *agent_arguments,
        )),
        DailyStep('周报生成' if generate_derived else '周报输出预览', (
            python_executable, '-m', 'ai_pmo', 'weekly-report', '--date', date_text,
            *weekly_arguments,
        )),
    ]


def run_daily_steps(
    steps: Sequence[DailyStep],
    *,
    runner: Callable[..., subprocess.CompletedProcess] = subprocess.run,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
) -> int:
    for index, step in enumerate(steps, 1):
        print(f'[{index}/{len(steps)}] {step.name}')
        result = runner(step.command, cwd=cwd, env=env, check=False)
        if result.returncode:
            print(f'日常运行已停止：{step.name}失败（退出码{result.returncode}）')
            return result.returncode
    print('日常运行完成：全部步骤通过')
    return 0


def execute_daily(
    suite: Path,
    report_date: date,
    *,
    generate_derived: bool = False,
    runner: Callable[..., subprocess.CompletedProcess] = subprocess.run,
    python_executable: str = sys.executable,
) -> int:
    try:
        sources = validate_daily_inputs(suite)
    except Exception as exc:
        print(f'日常运行已停止：输入校验失败：{exc}')
        return 1
    print(f'当前版本登记校验通过：{len(sources)}类工作簿')
    for key, path in sources.items():
        print(f'- {key}: {path.relative_to(suite)}')
    automation = suite / 'automation'
    environment = os.environ.copy()
    source_path = str(automation / 'src')
    environment['PYTHONPATH'] = source_path + (
        os.pathsep + environment['PYTHONPATH'] if environment.get('PYTHONPATH') else ''
    )
    steps = build_daily_steps(
        suite, report_date, python_executable,
        generate_derived=generate_derived,
    )
    return run_daily_steps(steps, runner=runner, cwd=automation, env=environment)
