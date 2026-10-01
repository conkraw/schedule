"""Report-boundary regressions: invented data and simulated GitHub/Streamlit."""
from copy import deepcopy
from datetime import date
from io import BytesIO
import json
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from helpers import run_app
from test_learner_reach import scan_cells, make_reach_fixture
from test_reporting_dates import all_doc_text
from test_strict_reach_charts import ui_values, CONFLICT_CELLS
from schedule_app.services.report_diagnostics import (
    ReportDataError, REPORT_BUILD_ID, REPORT_OUTPUT_VERSION, REPORT_ISSUE_COLUMNS,
    checked_report_reach, report_step,
)
from schedule_app.services.learner_reach import reach_totals, reach_percent
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.teaching_validation import TeachingConflictError, validate_teaching_report
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.reports.chair_summary import teaching_chair_summary_data, teaching_make_chair_summary
from schedule_app.reports.learner_reach_charts import learner_reach_pie
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME

NAME = 'Example, Avery'
COUNTS = {'recorded_clinical_shifts':10, 'shifts_with_students':8, 'shifts_without_students':2}


class ReportMetricChecks(unittest.TestCase):
    def test_missing_count_is_an_error_not_fake_zero(self):
        for field in COUNTS:
            bad = dict(COUNTS); bad.pop(field)
            with self.subTest(field=field), self.assertRaisesRegex(ReportDataError, field):
                reach_totals([bad], report='Chair summary', preceptor_name=NAME)

    def test_negative_fractional_and_nonfinite_counts_stop(self):
        for value in (-1, 0.5, float('nan'), float('inf'), True, '10'):
            with self.subTest(value=value), self.assertRaises(ReportDataError):
                reach_totals([dict(COUNTS, recorded_clinical_shifts=value)])

    def test_inconsistent_counts_stop_with_specific_preceptor(self):
        with self.assertRaises(ReportDataError) as caught:
            checked_report_reach(dict(COUNTS, shifts_with_students=12),
                report='Individual preceptor report', preceptor_name=NAME, work_type='Ward A')
        row = caught.exception.rows[0]
        self.assertEqual(row['preceptor_name'], NAME)
        self.assertEqual(row['work_type'], 'Ward A')
        self.assertIn('do not reconcile', row['issue'])

    def test_missing_percentage_rebuilt_from_counts(self):
        for pct in ('absent', None):
            row = dict(COUNTS)
            if pct is None: row['learner_reach_pct'] = None
            with self.subTest(pct=pct):
                self.assertEqual(checked_report_reach(row)['learner_reach_pct'], 80.0)

    def test_missing_hours_rebuilt_without_mutating_source(self):
        source = dict(COUNTS, learner_reach_pct=None)
        before = deepcopy(source)
        checked = checked_report_reach(source)
        self.assertEqual((checked['recorded_clinical_hours'],checked['hours_with_students'],checked['hours_without_students']),(40,32,8))
        self.assertEqual(source,before)

    def test_wrong_nonfinite_or_text_percentage_not_replaced(self):
        for value in (75, -1, 101, float('nan'), float('inf'), '80%', True):
            with self.subTest(value=value), self.assertRaises(ReportDataError):
                checked_report_reach(dict(COUNTS, learner_reach_pct=value))

    def test_wrong_hours_do_not_pass(self):
        for field in ('recorded_clinical_hours','hours_with_students','hours_without_students'):
            with self.subTest(field=field), self.assertRaisesRegex(ReportDataError,field):
                checked_report_reach(dict(COUNTS, **{field:999}))

    def test_empty_total_remains_undefined_and_cannot_publish(self):
        empty=reach_totals([])
        self.assertIsNone(empty['learner_reach_pct'])
        with self.assertRaises(OPDArchiveError):reach_percent(empty['learner_reach_pct'])
        with self.assertRaisesRegex(ReportDataError,'no denominator'):checked_report_reach(empty)

    def test_known_zero_and_one_hundred_are_valid(self):
        for n in (0,10):
            result=checked_report_reach(dict(COUNTS,shifts_with_students=n,shifts_without_students=10-n))
            self.assertEqual(result['learner_reach_pct'],n*10)

    def test_aggregate_is_ratio_of_counts(self):
        other={'recorded_clinical_shifts':2,'shifts_with_students':1,'shifts_without_students':1}
        self.assertEqual(reach_totals([COUNTS,other])['learner_reach_pct'],75.0)

    def test_unresolved_review_flag_still_blocks(self):
        with self.assertRaisesRegex(ReportDataError,'conflict'):
            checked_report_reach(dict(COUNTS,availability_review_shifts=1))

    def test_plain_message_is_reproduced_at_old_formatter(self):
        # The exact screenshot text is the old formatter's response to None.
        with self.assertRaisesRegex(OPDArchiveError,'Learner Reach is undefined or invalid'):
            reach_percent(None)

    def test_chart_recovers_absent_display_fields_from_complete_counts(self):
        png=learner_reach_pie('Academic Pediatrics',dict(COUNTS,learner_reach_pct=None))
        self.assertTrue(png.startswith(b'\x89PNG'))

    def test_chart_failure_identifies_area(self):
        with self.assertRaises(ReportDataError) as caught:
            learner_reach_pie('Academic Pediatrics',dict(COUNTS,shifts_without_students=4),'26-27')
        row=caught.exception.rows[0]
        self.assertEqual(row['report'],'Clinical experience pie chart')
        self.assertEqual(row['work_type'],'Academic Pediatrics')
        self.assertEqual(row['academic_year'],'26-27')


class ReportBoundaryChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scan,cls.repo,cls.client,cls.secrets,_=scan_cells(make_reach_fixture())

    def test_context_identifies_failed_individual_and_no_partial_zip(self):
        with patch('schedule_app.reports.teaching_export.teaching_make_docx',side_effect=OPDArchiveError('Learner Reach is undefined or invalid.')):
            with self.assertRaises(ReportDataError) as caught:
                teaching_build_zip(self.scan,[2026])
        row=caught.exception.rows[0]
        self.assertEqual(row['report'],'Individual preceptor Word report')
        self.assertEqual(row['preceptor_name'],NAME)

    def test_context_identifies_failed_chair_stage(self):
        with patch('schedule_app.reports.teaching_export.teaching_make_chair_summary',side_effect=OPDArchiveError('Learner Reach is undefined or invalid.')):
            with self.assertRaises(ReportDataError) as caught:teaching_build_zip(self.scan,[2026])
        self.assertEqual(caught.exception.rows[0]['report'],'Chair Word report')

    def test_context_identifies_failed_chart_stage(self):
        with patch('schedule_app.reports.teaching_export.teaching_clinical_charts',side_effect=OPDArchiveError('Learner Reach is undefined or invalid.')):
            with self.assertRaises(ReportDataError) as caught:teaching_build_zip(self.scan,[2026])
        self.assertEqual(caught.exception.rows[0]['report'],'Clinical experience pie charts')

    def test_unexpected_exception_text_is_not_exposed(self):
        with self.assertRaises(ReportDataError) as caught:
            with report_step('Individual report',preceptor_name=NAME):
                raise ValueError('Private Learner Name; token=must-not-be-exported')
        text=json.dumps(caught.exception.rows)+str(caught.exception)
        self.assertNotIn('Private Learner',text)
        self.assertNotIn('must-not-be-exported',text)
        self.assertIn('ValueError',text)

    def test_conflict_source_details_are_not_replaced(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        with self.assertRaises(TeachingConflictError) as caught:
            with report_step('Chair summary'):validate_teaching_report(scan,[2026])
        self.assertTrue(caught.exception.rows[0]['cell'])

    def test_chair_handles_missing_derived_percentages(self):
        summaries=teaching_chair_summary_data(self.scan,[2026])
        for item in summaries:
            item['learner_reach_pct']=None
            for group in item['work_types']:
                group['learner_reach_pct']=None
                for row in group['named_preceptors']+group['unresolved_labels']:
                    row['learner_reach_pct']=None
        before=deepcopy(summaries)
        with patch('schedule_app.reports.chair_summary.teaching_chair_summary_data',return_value=summaries):
            raw=teaching_make_chair_summary(self.scan,[2026])
        text=all_doc_text(raw)
        self.assertIn('80.0%',text)
        self.assertNotIn('N/A',text)
        self.assertEqual(summaries,before)

    def test_invalid_chair_percentage_is_not_hidden(self):
        summaries=teaching_chair_summary_data(self.scan,[2026])
        row=summaries[0]['work_types'][0]['named_preceptors'][0]
        row['learner_reach_pct']=101
        with patch('schedule_app.reports.chair_summary.teaching_chair_summary_data',return_value=summaries):
            with self.assertRaises(ReportDataError) as caught:teaching_make_chair_summary(self.scan,[2026])
        self.assertEqual(caught.exception.rows[0]['preceptor_name'],NAME)
        self.assertEqual(caught.exception.rows[0]['work_type'],'Academic Pediatrics')

    def test_empty_optional_subgroups_have_no_fake_percentage(self):
        scan,*_=scan_cells({('SJR_HOSP','B6'):'SJR_1 ~ Private Learner'})
        raw=teaching_make_chair_summary(scan,[2026])
        text=all_doc_text(raw)
        self.assertNotIn('Named preceptors subtotal',text)
        self.assertIn('100.0%',text)

    def test_ui_gives_safe_diagnostic_csv_instead_of_partial_download(self):
        first=run_app(ui_values(),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        exc=ReportDataError('No percentage could be generated.',metrics=COUNTS,
                           report='Chair summary',preceptor_name=NAME,work_type='Academic Pediatrics')
        # Force a new build: unchanged valid ZIPs are intentionally reused.
        first['state'].pop('teaching_zip', None)
        with patch('schedule_app.sections.preceptor_teaching_summary.teaching_build_zip',side_effect=exc):
            bad=run_app(ui_values(teaching_load_archives=False),secrets=self.secrets,repo=self.repo,state=first['state'], evaluation_login=True)
        self.assertNotIn('teaching_zip',bad['state'])
        self.assertFalse(any(name.endswith(('.zip','.docx')) for name in bad['downloads']))
        csv=bad['downloads']['Learner_Reach_Report_Issues.csv'].decode('utf-8-sig')
        self.assertIn(NAME,csv)
        self.assertIn('Academic Pediatrics',csv)
        self.assertNotIn('Confidential',csv)
        self.assertNotIn('offline-test-token',csv)
        self.assertTrue(any(REPORT_BUILD_ID in text for _,text in bad['messages']))

    def test_ui_refresh_not_needed_for_current_scan_or_date_inputs(self):
        first=run_app(ui_values(),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        state=deepcopy(first['state']); state['teaching_zip_signature']=('old output',)
        second=run_app(ui_values(teaching_load_archives=False,teaching_build_zip=False),secrets=self.secrets,repo=self.repo,state=state, evaluation_login=True)
        self.assertEqual(second['state']['teaching_scan'],first['state']['teaching_scan'])
        self.assertEqual(second['state']['teaching_period_start'],first['state']['teaching_period_start'])
        self.assertEqual(second['state']['teaching_period_label'],first['state']['teaching_period_label'])
        self.assertNotIn('teaching_zip',second['state'])
        self.assertEqual(second['state']['teaching_zip_signature'][0][-1],REPORT_OUTPUT_VERSION)

    def test_custom_dates_priority_and_unique_students_survive_export(self):
        cells={('NYES','B8'):'Example, Avery ~ Private Clinic',
               ('PSHCH_NURSERY','B8'):'Example, Avery ~ Private Nursery',
               ('NYES','C6'):'Example, Avery ~ Private Clinic',
               ('NYES','E8'):'Example, Avery ~ Private Clinic',
               ('NYES','F8'):'Example, Avery ~ '}
        scan,*_=scan_cells(cells)
        view=teaching_filter_date_range(scan,ReportingPeriod('Custom',date(2026,9,7),date(2026,9,13)))
        raw,rows=teaching_build_zip(view,[2026])
        self.assertEqual((rows[0]['no_of_shifts'],rows[0]['learner_reach_pct']),(3,75.0))
        with ZipFile(BytesIO(raw)) as z:
            chair=all_doc_text(z.read(TEACHING_CHAIR_SUMMARY_FILENAME))
            self.assertIn('Students meeting minimum shifts',chair)
            self.assertIn('75.0%',chair)
            individual=next(n for n in z.namelist() if n.startswith('Preceptor_Reports/'))
            self.assertIn('Unique students assigned: 1',all_doc_text(z.read(individual)))
            self.assertNotIn('Private Clinic',chair)
            self.assertIn('Outpatient_Priority_Adjustments.csv',z.namelist())


if __name__=='__main__':unittest.main()
