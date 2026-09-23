import unittest

from ai_pmo.portfolio_audit import (
    check_decision_detail, check_deliverable_consistency, check_issue_detail,
    check_migration_progress, check_milestone_links, check_milestone_owner_alignment,
    check_milestone_owner_drift, check_quality_evidence, check_risk_followup,
    check_weekly_report, parse_weekly_report, compare_dashboard_snapshot,
    dashboard_metrics_from_cells,
)


class PortfolioAuditTests(unittest.TestCase):
    def test_milestone_links_report_missing_and_uncovered_ids(self):
        result = check_milestone_links(['MS001', 'MS002'], ['MS001', 'MS999'])
        self.assertEqual(result['errors'], ['总控任务引用不存在的里程碑 MS999'])
        self.assertEqual(result['warnings'], ['里程碑 MS002 没有总控任务'])

    def test_dashboard_snapshot_reports_source_and_metric_drift_as_warnings(self):
        result = compare_dashboard_snapshot(
            expected_sources={'plan': '02/new.xlsx'},
            actual_sources={'plan': '02/old.xlsx'},
            expected_metrics={'总任务数': 135, '风险总数': 12},
            actual_metrics={'总任务数': 135, '风险总数': 11},
        )
        self.assertEqual(result['errors'], [])
        self.assertIn('驾驶舱来源快照过期：plan 登记为02/new.xlsx，快照为02/old.xlsx', result['warnings'])
        self.assertIn('驾驶舱指标不一致：风险总数 源台账=12，快照=11', result['warnings'])

    def test_missing_dashboard_metric_is_pending_refresh_not_zero(self):
        result = compare_dashboard_snapshot({}, {}, {'计划人工成本': 100}, {})
        self.assertEqual(result['warnings'], ['驾驶舱指标待刷新：计划人工成本 源台账=100，快照缺失'])

    def test_ignores_insignificant_floating_point_difference(self):
        result = compare_dashboard_snapshot({}, {}, {'人日单价': 1454.18502202643},
                                            {'人日单价': 1454.1850220264316})
        self.assertEqual(result['warnings'], [])

    def test_dashboard_metrics_use_fixed_interface_cells_when_labels_are_absent(self):
        values = {14: 4, 17: 1135, 18: 1650500, 19: 1454.18, 36: 17}
        self.assertEqual(dashboard_metrics_from_cells(values), {
            '待决策事项': 4, '里程碑总数': 17, '计划投入人天': 1135,
            '计划人工成本': 1650500, '人日单价': 1454.18,
        })

    def test_migration_progress_warns_until_all_sources_inventoried(self):
        result = check_migration_progress(
            {'来源线索数': 6, '已盘点来源系统数': 2, 'G15 目标日期': '2026-10-13'})
        self.assertEqual(result['counts'], {'迁移来源线索': 6, '已盘点来源': 2})
        self.assertEqual(result['warnings'], [
            '迁移来源系统盘点未完成：已盘点 2/6，G15 目标日期 2026-10-13'])

    def test_migration_progress_is_silent_when_inventory_complete(self):
        result = check_migration_progress({'来源线索数': 6, '已盘点来源系统数': 6})
        self.assertEqual(result['counts'], {'迁移来源线索': 6, '已盘点来源': 6})
        self.assertEqual(result['warnings'], [])

    def test_migration_progress_flags_missing_interface_stats(self):
        result = check_migration_progress({})
        self.assertEqual(result['counts'], {})
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('缺少_数据接口', result['warnings'][0])

    def test_milestone_owner_drift_flags_owner_changed_after_dashboard_snapshot(self):
        result = check_milestone_owner_drift(
            {'可行性研究完成': '项目经理', '可研工作启动': '项目经理'},
            {'可行性研究完成': '项目集经理', '可研工作启动': '项目经理'},
        )
        self.assertEqual(result['errors'], [])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('可行性研究完成', result['warnings'][0])
        self.assertIn('项目集经理', result['warnings'][0])
        self.assertIn('项目经理', result['warnings'][0])

    def test_milestone_owner_drift_ignores_milestones_absent_from_dashboard(self):
        result = check_milestone_owner_drift({'正式启动': '项目集经理'}, {})
        self.assertEqual(result['warnings'], [])

    def test_milestone_owner_alignment_warns_when_owner_own_no_assigned_task(self):
        milestones = [{'id': 'MS001', 'name': '可研工作启动', 'owner': '项目集经理',
                       'status': '已完成'}]
        result = check_milestone_owner_alignment(milestones, {'MS001': ['项目经理', '产品经理']})
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('MS001', result['warnings'][0])

    def test_milestone_owner_alignment_skips_unassigned_or_not_started_batches(self):
        milestones = [
            {'id': 'MS001', 'name': 'A', 'owner': '项目集经理', 'status': '已完成'},
            {'id': 'MS002', 'name': 'B', 'owner': '技术经理', 'status': '未开始'},
        ]
        result = check_milestone_owner_alignment(
            milestones, {'MS001': ['项目经理', ''], 'MS002': ['技术经理']})
        self.assertEqual(result['warnings'], [])

    def test_risk_followup_warns_when_open_high_risk_lacks_action(self):
        risks = [{'id': 'RSK-001', 'level': '高', 'status': '处理中',
                  'strategy': '减轻', 'deadline': '2026-09-25'}]
        result = check_risk_followup(risks, {})
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('RSK-001', result['warnings'][0])
        self.assertIn('应对计划', result['warnings'][0])

    def test_risk_followup_warns_when_action_missing_owner_or_plan_date(self):
        risks = [{'id': 'RSK-004', 'level': '重大', 'status': '待处理',
                  'strategy': '规避', 'deadline': '2026-10-13'}]
        result = check_risk_followup(risks, {'RSK-004': [{'owner': '', 'plan_date': None}]})
        self.assertEqual(len(result['warnings']), 2)
        self.assertTrue(any('行动责任人' in w for w in result['warnings']))
        self.assertTrue(any('计划完成日期' in w for w in result['warnings']))

    def test_risk_followup_silent_for_closed_or_medium_risks(self):
        risks = [
            {'id': 'RSK-001', 'level': '高', 'status': '已关闭',
             'strategy': '', 'deadline': None},
            {'id': 'RSK-010', 'level': '中', 'status': '待处理',
             'strategy': '接受', 'deadline': None},
        ]
        result = check_risk_followup(risks, {})
        self.assertEqual(result['warnings'], [])

    def test_issue_detail_warns_when_open_high_issue_lacks_response_or_plan(self):
        issues = [{'id': 'G10', 'level': '高', 'status': '待确认',
                   'response': '', 'plan_close': None, 'actual_close': None}]
        result = check_issue_detail(issues)
        self.assertEqual(len(result['warnings']), 2)
        self.assertTrue(any('应对措施' in w for w in result['warnings']))
        self.assertTrue(any('计划关闭' in w for w in result['warnings']))

    def test_issue_detail_warns_when_actual_close_filled_but_status_open(self):
        issues = [{'id': 'G10', 'level': '高', 'status': '待确认',
                   'response': '有措施', 'plan_close': '2026-09-25',
                   'actual_close': '2026-09-20'}]
        result = check_issue_detail(issues)
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('实际关闭', result['warnings'][0])

    def test_issue_detail_silent_for_closed_issue(self):
        issues = [{'id': 'G10', 'level': '高', 'status': '已关闭',
                   'response': '', 'plan_close': None, 'actual_close': '2026-09-20'}]
        self.assertEqual(check_issue_detail(issues)['warnings'], [])

    def test_decision_detail_warns_when_pending_decision_lacks_proposal(self):
        decisions = [{'id': 'D-001', 'status': '待决策', 'proposal': '',
                      'need_date': None, 'decision_maker': '委员会'}]
        result = check_decision_detail(decisions)
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('D-001', result['warnings'][0])
        self.assertIn('建议方案', result['warnings'][0])
        self.assertIn('需要日期', result['warnings'][0])

    def test_decision_detail_silent_when_pending_decision_is_complete(self):
        decisions = [{'id': 'D-001', 'status': '待决策', 'proposal': '建议方案',
                      'need_date': '2026-09-25', 'decision_maker': '委员会'}]
        self.assertEqual(check_decision_detail(decisions)['warnings'], [])

    def test_quality_evidence_warns_when_coverage_complete_without_evidence(self):
        result = check_quality_evidence(
            coverage=[{'id': 'REQ-001', 'exec_result': '', 'evidence': '',
                       'coverage': '完整'}])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('REQ-001', result['warnings'][0])

    def test_quality_evidence_warns_when_uat_passed_without_evidence(self):
        result = check_quality_evidence(
            uat=[{'id': 'UAT-001', 'exec_result': '通过', 'evidence': ''}])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('UAT-001', result['warnings'][0])

    def test_quality_evidence_warns_when_open_severe_defect_lacks_plan(self):
        result = check_quality_evidence(
            defects=[{'id': 'DEF-001', 'severity': 'Critical', 'status': '处理中',
                      'plan_done': None, 'verify_evidence': ''}])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('DEF-001', result['warnings'][0])

    def test_quality_evidence_warns_when_closed_defect_lacks_verify_evidence(self):
        result = check_quality_evidence(
            defects=[{'id': 'DEF-002', 'severity': 'Minor', 'status': '已关闭',
                      'plan_done': '2027-01-01', 'verify_evidence': ''}])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('验证证据', result['warnings'][0])

    def test_quality_evidence_silent_for_not_started_and_clean_rows(self):
        result = check_quality_evidence(
            coverage=[{'id': 'REQ-001', 'exec_result': '', 'evidence': '',
                       'coverage': '不完整'}],
            uat=[{'id': 'UAT-001', 'exec_result': '未开始', 'evidence': ''}])
        self.assertEqual(result['warnings'], [])

    def test_deliverable_consistency_warns_on_interface_count_mismatch(self):
        rows = [{'id': 'D-001', 'status': '未开始', 'actual_done': None,
                 'review_result': '待评审', 'location': ''}]
        result = check_deliverable_consistency(rows, interface_total=13)
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('13', result['warnings'][0])

    def test_deliverable_consistency_warns_when_accepted_lacks_location(self):
        rows = [{'id': 'D-001', 'status': '已验收', 'actual_done': '2026-10-31',
                 'review_result': '通过', 'location': ''}]
        result = check_deliverable_consistency(rows, interface_total=1)
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('存放位置', result['warnings'][0])

    def test_deliverable_consistency_silent_for_not_started_rows(self):
        rows = [{'id': f'D-{i:03d}', 'status': '未开始', 'actual_done': None,
                 'review_result': '待评审', 'location': ''} for i in range(1, 14)]
        self.assertEqual(
            check_deliverable_consistency(rows, interface_total=13)['warnings'], [])

    def test_weekly_report_flags_severity_nesting_violation_as_error(self):
        report = {'file_date': '2026-09-23', 'data_date': '2026-09-23',
                  'status': '红', 'finding_count': 5, 'major': 4, 'high': 3,
                  'medium': 2, 'approval_count': 1}
        result = check_weekly_report(report, dashboard_data_date='2026-09-23',
                                     today='2026-09-24')
        self.assertEqual(len(result['errors']), 1)
        self.assertIn('严重度口径', result['errors'][0])

    def test_weekly_report_flags_data_date_after_file_date_as_error(self):
        report = {'file_date': '2026-09-23', 'data_date': '2026-09-24',
                  'status': '红', 'finding_count': 9, 'major': 4, 'high': 3,
                  'medium': 2, 'approval_count': 1}
        result = check_weekly_report(report, dashboard_data_date='2026-09-24',
                                     today='2026-09-24')
        self.assertEqual(len(result['errors']), 1)
        self.assertIn('晚于发布日期', result['errors'][0])

    def test_weekly_report_warns_when_caliber_lags_dashboard(self):
        report = {'file_date': '2026-09-20', 'data_date': '2026-09-20',
                  'status': '黄', 'finding_count': 9, 'major': 4, 'high': 3,
                  'medium': 2, 'approval_count': 1}
        result = check_weekly_report(report, dashboard_data_date='2026-09-23',
                                     today='2026-09-24')
        self.assertEqual(result['errors'], [])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('口径滞后', result['warnings'][0])

    def test_weekly_report_warns_when_weekly_cadence_broken(self):
        report = {'file_date': '2026-09-10', 'data_date': '2026-09-10',
                  'status': '红', 'finding_count': 9, 'major': 4, 'high': 3,
                  'medium': 2, 'approval_count': 1}
        result = check_weekly_report(report, dashboard_data_date='2026-09-10',
                                     today='2026-09-24')
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('超过一周未更新', result['warnings'][0])

    def test_weekly_report_silent_when_fresh_and_consistent(self):
        report = {'file_date': '2026-09-23', 'data_date': '2026-09-23',
                  'status': '红', 'finding_count': 58, 'major': 4, 'high': 43,
                  'medium': 11, 'approval_count': 30}
        result = check_weekly_report(report, dashboard_data_date='2026-09-23',
                                     today='2026-09-24')
        self.assertEqual(result['errors'], [])
        self.assertEqual(result['warnings'], [])

    def test_parse_weekly_report_reads_metrics_and_data_date(self):
        from datetime import datetime
        from openpyxl import Workbook
        import tempfile
        from pathlib import Path
        book = Workbook()
        sheet = book.active
        sheet.title = '周报摘要'
        sheet['A2'] = '能源行业合同全生命周期管理系统｜数据截止日 2026-09-23'
        sheet.append([])
        sheet.append(['项目状态', '发现总数', '重大', '高', '中', '需人工确认'])
        sheet.append(['红', 58, 4, 43, 11, 30])
        usage = book.create_sheet('使用说明')
        usage.append(['项目周报使用说明'])
        usage.append(['数据截止日', datetime(2026, 9, 23)])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / '项目周报-20260923-V1.xlsx'
            book.save(path)
            parsed = parse_weekly_report(path)
        self.assertEqual(parsed['file_date'], '2026-09-23')
        self.assertEqual(parsed['data_date'], '2026-09-23')
        self.assertEqual(parsed['status'], '红')
        self.assertEqual(parsed['finding_count'], 58)
        self.assertEqual(parsed['major'], 4)
        self.assertEqual(parsed['high'], 43)
        self.assertEqual(parsed['medium'], 11)
        self.assertEqual(parsed['approval_count'], 30)


if __name__ == '__main__':
    unittest.main()
