from __future__ import annotations

import subprocess
import tempfile
import unittest
from datetime import date
from pathlib import Path

from ai_pmo.current_workbooks import WORKBOOKS
from ai_pmo.daily_run import DailyStep, build_daily_steps, execute_daily, run_daily_steps, validate_daily_inputs


class DailyRunTests(unittest.TestCase):
    def make_suite(self, root: Path, *, hourly_rate: float = 150) -> Path:
        suite = root / 'suite'
        entries = {}
        for key, (directory, prefix) in WORKBOOKS.items():
            target = suite / directory / f'{prefix}-20260921-V1.xlsx'
            target.parent.mkdir(parents=True, exist_ok=True)
            target.touch()
            entries[key] = target.relative_to(suite).as_posix()
        config = suite / 'automation' / 'config'
        config.mkdir(parents=True, exist_ok=True)
        (config / 'current_workbooks.json').write_text(__import__('json').dumps(entries), encoding='utf-8')
        (config / 'project.json').write_text(__import__('json').dumps({
            'hours_per_day': 8, 'hours_per_week': 40,
            'person_day_rate': 1200, 'hourly_rate': hourly_rate,
        }), encoding='utf-8')
        return suite

    def test_validates_all_registered_sources_and_cost_parameters(self):
        with tempfile.TemporaryDirectory() as directory:
            result = validate_daily_inputs(self.make_suite(Path(directory)))
        self.assertEqual(set(result), set(WORKBOOKS))

    def test_rejects_inconsistent_hourly_rate_before_running_steps(self):
        with tempfile.TemporaryDirectory() as directory:
            suite = self.make_suite(Path(directory), hourly_rate=149)
            with self.assertRaisesRegex(ValueError, '小时单价'):
                validate_daily_inputs(suite)

    def test_execute_daily_does_not_start_steps_when_inputs_are_invalid(self):
        with tempfile.TemporaryDirectory() as directory:
            suite = self.make_suite(Path(directory), hourly_rate=149)
            calls = []
            result = execute_daily(
                suite, date(2026, 9, 21),
                runner=lambda command, **kwargs: calls.append(command),
            )
        self.assertEqual(result, 1)
        self.assertEqual(calls, [])

    def test_default_steps_are_preview_only(self):
        suite = Path('/tmp/suite')
        steps = build_daily_steps(suite, date(2026, 9, 21), 'python3')
        self.assertEqual(
            [step.name for step in steps],
            ['数据主源审计', '数据联动审计', '接口快照同步预览', '项目组合联动审计', 'AI PMO全量刷新预览', '成本刷新预览', '成本口径G15联动预览', 'Agent输出预览', '周报输出预览'],
        )
        flattened = [argument for step in steps for argument in step.command]
        self.assertNotIn('--apply', flattened)
        self.assertNotIn('--replace-generated', flattened)
        self.assertEqual(flattened.count('--preview-targets'), 2)
        self.assertTrue(all(
            step.command[step.command.index('-m') + 1] == 'ai_pmo'
            for step in steps if '-m' in step.command
        ))

    def test_generate_derived_runs_agents_and_weekly_report_without_apply(self):
        steps = build_daily_steps(
            Path('/tmp/suite'), date(2026, 9, 21), 'python3', generate_derived=True,
        )
        agent = next(step for step in steps if step.name == 'Agent分析')
        weekly = next(step for step in steps if step.name == '周报生成')
        self.assertNotIn('--preview-targets', agent.command)
        self.assertNotIn('--preview-targets', weekly.command)
        self.assertNotIn('--apply', [arg for step in steps for arg in step.command])

    def test_stops_after_first_failed_step(self):
        calls = []

        def runner(command, **kwargs):
            calls.append(command)
            return subprocess.CompletedProcess(command, 1 if len(calls) == 2 else 0)

        steps = [
            DailyStep('一', ('cmd-1',)),
            DailyStep('二', ('cmd-2',)),
            DailyStep('三', ('cmd-3',)),
        ]
        result = run_daily_steps(steps, runner=runner)
        self.assertEqual(result, 1)
        self.assertEqual(calls, [('cmd-1',), ('cmd-2',)])


if __name__ == '__main__':
    unittest.main()
