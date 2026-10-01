"""Report inclusion policy: skip zero-only contributors, never inflate denominators.

All fixtures and GitHub calls in this suite are synthetic and in memory.
"""
import copy
import csv
from datetime import date
from io import BytesIO, StringIO
import unittest
from zipfile import ZipFile
from unittest.mock import patch

from helpers import run_app
from test_learner_reach import scan_cells, make_reach_fixture
from test_reporting_dates import small_opd, all_doc_text
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import (
    teaching_scan_archives, teaching_filter_date_range, teaching_annual_rows, teaching_work_type_rows,
)
from schedule_app.services.learner_reach import (
    learner_reach_rows, participating_reach_rows, teaching_participation_keys, PARTICIPATION_REPORT_VERSION,
)
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.chair_summary import teaching_chair_summary_data, teaching_make_chair_summary
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME


def csv_rows(zf, name):
    return list(csv.DictReader(StringIO(zf.read(name).decode('utf-8-sig'))))


class TeachingContributorsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cells = make_reach_fixture()
        cells.update({('ADOLMED', 'B6'): 'Excluded, Adolescent ~ ',
                      ('ADOLMED', 'C6'): 'Excluded, Adolescent ~ ',
                      ('SJR_HOSP', 'B6'): 'SJR_1 ~ ',
                      ('WARD A', 'C6'): 'Ward, Casey ~ Confidential Ward Learner'})
        cls.scan, cls.repo, cls.client, cls.secrets, _ = scan_cells(cells)
        cls.blob, cls.annual = teaching_build_zip(cls.scan, [2026])

    def test_zero_services_absent_from_chair_data(self):
        summary = teaching_chair_summary_data(self.scan, [2026])[0]
        self.assertEqual({g['work_type'] for g in summary['work_types']}, {'Academic Pediatrics', 'Ward A'})
        self.assertEqual(summary['named_preceptor_count'], 2)
        self.assertEqual(summary['unresolved_labels'], [])

    def test_word_reports_hide_zero_names_and_services(self):
        with ZipFile(BytesIO(self.blob)) as zf:
            docs = [name for name in zf.namelist() if name.endswith('.docx')]
            self.assertEqual(len(docs), 3)  # chair + two teaching contributors
            for name in docs:
                text = all_doc_text(zf.read(name))
                for unwanted in ('Excluded, Adolescent', 'ADOLMED', 'Available, Blake', 'SJR_1', 'Confidential'):
                    self.assertNotIn(unwanted, text)

    def test_csvs_hide_zero_names_and_services(self):
        with ZipFile(BytesIO(self.blob)) as zf:
            for name in ('preceptor_teaching_summary.csv', 'preceptor_teaching_by_work_type.csv',
                         'preceptor_learner_reach_monthly.csv'):
                rows = csv_rows(zf, name)
                self.assertEqual({r['preceptor_name'] for r in rows}, {'Example, Avery', 'Ward, Casey'})
                self.assertNotIn('ADOLMED', {r.get('work_type') for r in rows})
                if name != 'preceptor_learner_reach_monthly.csv':
                    self.assertTrue(all(int(r['teaching_shifts']) > 0 for r in rows))

    def test_notes_do_not_list_nonparticipants_or_zero_mapping(self):
        with ZipFile(BytesIO(self.blob)) as zf:
            text = zf.read('Report_Notes.txt').decode()
            for name in ('ADOLMED', 'Excluded, Adolescent', 'Available, Blake', 'SJR_1'):
                self.assertNotIn(name, text)
            self.assertIn('including shifts without students', text)

    def test_non_teaching_shifts_keep_eighty_percent(self):
        row = next(r for r in self.annual if r['preceptor_name'] == 'Example, Avery')
        self.assertEqual((row['shifts_with_students'], row['recorded_clinical_shifts'], row['learner_reach_pct']), (8, 10, 80.0))
        self.assertEqual((row['no_of_shifts'], row['educational_hours']), (9, 32))

    def test_raw_scan_contains_excluded_people_and_is_not_mutated(self):
        before = copy.deepcopy(self.scan)
        teaching_build_zip(self.scan, [2026])
        participating_reach_rows(self.scan, [2026], by_work_type=True, monthly=True)
        self.assertEqual(before, self.scan)
        self.assertIn('Excluded, Adolescent', {r['preceptor_name'] for r in learner_reach_rows(self.scan, [2026])})

    def test_inactive_type_for_active_person_hidden_but_overall_hours_kept(self):
        scan, *_ = scan_cells({('NYES', 'B6'): 'Example, Avery ~ Learner',
                              ('NYES', 'C6'): 'Example, Avery ~ ',
                              ('ADOLMED', 'D6'): 'Example, Avery ~ '})
        overall = teaching_annual_rows(scan, [2026])[0]
        typed = teaching_work_type_rows(scan, [2026])
        self.assertEqual((overall['recorded_clinical_hours'], overall['learner_reach_pct']), (12, 33.3))
        self.assertEqual(len(typed), 1)
        self.assertEqual((typed[0]['recorded_clinical_hours'], typed[0]['learner_reach_pct']), (8, 50.0))
        blob, _ = teaching_build_zip(scan, [2026])
        with ZipFile(BytesIO(blob)) as zf:
            for name in [f for f in zf.namelist() if f.endswith('.docx')]:
                text = all_doc_text(zf.read(name))
                self.assertNotIn('ADOLMED', text)
                self.assertIn('33.3%', text)
                self.assertIn('may be lower than the overall total', text)

    def test_zero_person_within_active_category_hidden(self):
        typed = teaching_work_type_rows(self.scan, [2026])
        academic = [r for r in typed if r['work_type'] == 'Academic Pediatrics']
        self.assertEqual([r['preceptor_name'] for r in academic], ['Example, Avery'])
        self.assertEqual(academic[0]['recorded_clinical_shifts'], 10)  # excludes Blake's denominator

    def test_adolmed_returns_automatically_when_teaching_occurs(self):
        scan, *_ = scan_cells({('ADOLMED', 'B6'): 'Example, Avery ~ Learner',
                              ('ADOLMED', 'C6'): 'Example, Avery ~ '})
        typed = teaching_work_type_rows(scan, [2026])
        self.assertEqual((typed[0]['work_type'], typed[0]['learner_reach_pct']), ('ADOLMED', 50.0))
        self.assertIn('ADOLMED', all_doc_text(teaching_make_chair_summary(scan, [2026])))

    def test_month_without_students_retained_for_participating_provider_type(self):
        scan, repo, client, _, _ = scan_cells({('NYES', 'C6'): 'Example, Avery ~ Learner',
                                               ('ETOWN', 'D6'): 'Example, Avery ~ '}, start=date(2026, 6, 29))
        view = teaching_filter_date_range(scan, ReportingPeriod('Custom', date(2026, 6, 30), date(2026, 7, 1)))
        blob, _ = teaching_build_zip(view, [2026])
        with ZipFile(BytesIO(blob)) as zf:
            monthly = csv_rows(zf, 'preceptor_learner_reach_monthly.csv')
            self.assertEqual([r['month'] for r in monthly], ['2026-06-01', '2026-07-01'])
            self.assertEqual([r['learner_reach_pct'] for r in monthly], ['100.0', '0.0'])
            self.assertEqual(csv_rows(zf, 'preceptor_teaching_summary.csv')[0]['learner_reach_pct'], '50.0')

    def test_eligibility_uses_selected_dates_not_ever_elsewhere(self):
        scan, *_ = scan_cells({('NYES', 'B6'): 'Example, Avery ~ Learner', ('NYES', 'C6'): 'Example, Avery ~ '})
        view = teaching_filter_date_range(scan, ReportingPeriod('Blank day only', date(2026, 9, 8), date(2026, 9, 8)))
        self.assertEqual(teaching_annual_rows(view, [2026]), [])
        self.assertEqual(participating_reach_rows(view, [2026], by_work_type=True, monthly=True), [])
        self.assertEqual(len(learner_reach_rows(view, [2026])), 1)

    def test_no_empty_individual_year_or_chair_year(self):
        scan, *_ = scan_cells({('NYES', 'C6'): 'Example, Avery ~ Learner', ('NYES', 'D6'): 'Example, Avery ~ '}, start=date(2026, 6, 29))
        blob, _ = teaching_build_zip(scan, [2025, 2026])
        with ZipFile(BytesIO(blob)) as zf:
            for filename in [f for f in zf.namelist() if f.endswith('.docx')]:
                text = all_doc_text(zf.read(filename))
                self.assertIn('Academic year 25-26', text)
                self.assertNotIn('Academic year 26-27', text)
        self.assertEqual([g['academic_start_year'] for g in teaching_chair_summary_data(scan, [2025, 2026])], [2025])

    def test_direct_zero_report_calls_rejected(self):
        scan, *_ = scan_cells({('ADOLMED', 'B6'): 'Example, Avery ~ '})
        for fn in (lambda: teaching_make_docx('Example, Avery', [], scan),
                   lambda: teaching_make_chair_summary(scan, [2026]), lambda: teaching_build_zip(scan, [2026])):
            with self.assertRaisesRegex(OPDArchiveError, 'No student assignments'):
                fn()

    def test_full_app_previews_and_downloads_match(self):
        result = run_app({'schedule_app_mode': 'PTS', 'teaching_load_archives': True,
                          'teaching_build_zip': True, 'teaching_period_start': date(2026, 9, 7),
                          'teaching_period_end': date(2026, 9, 13), 'teaching_period_label': '26-27'},
                         secrets=self.secrets, repo=self.repo, evaluation_login=True)
        self.assertFalse(any(kind == 'error' for kind, text in result['messages']))
        table_text = '\n'.join(text for kind, text in result['messages'] if kind == 'dataframe')
        self.assertNotIn('Available, Blake', table_text)
        self.assertNotIn('ADOLMED', table_text)
        self.assertEqual(result['state']['teaching_zip_signature'][0][-4], PARTICIPATION_REPORT_VERSION)
        self.assertIn(TEACHING_CHAIR_SUMMARY_FILENAME, result['downloads'])

    def test_old_download_invalidated_without_rescanning(self):
        values = {'schedule_app_mode': 'PTS', 'teaching_period_start': date(2026, 9, 7),
                  'teaching_period_end': date(2026, 9, 13), 'teaching_period_label': '26-27'}
        first = run_app(dict(values, teaching_load_archives=True, teaching_build_zip=True), secrets=self.secrets, repo=self.repo, evaluation_login=True)
        old = copy.deepcopy(first['state'])
        old['teaching_zip_signature'] = old['teaching_zip_signature'][:-1]
        old['teaching_zip'] = b'old unfiltered export'
        with patch.object(self.repo, 'request', wraps=self.repo.request) as request_mock:
            result = run_app(values, secrets=self.secrets, repo=self.repo, state=old, evaluation_login=True)
            request_mock.assert_not_called()
        self.assertNotIn('teaching_zip', result['state'])
        self.assertIsNotNone(result['state'].get('teaching_scan'))
        self.assertFalse(any(name.endswith('.zip') for name in result['downloads']))
        self.assertEqual(old['teaching_scan'], result['state']['teaching_scan'])

    def test_reverse_name_order_same_exclusion(self):
        scan, *_ = scan_cells({('NYES', 'B6'): 'Learner ~ Example, Avery',
                              ('NYES', 'C6'): ' ~ Example, Avery',
                              ('ADOLMED', 'B6'): ' ~ Excluded, Adolescent'}, order='Student ~ Preceptor')
        rows = teaching_annual_rows(scan, [2026])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['learner_reach_pct'], 50.0)
        self.assertEqual([r['work_type'] for r in teaching_work_type_rows(scan, [2026])], ['Academic Pediatrics'])

    def test_conflict_validation_does_not_hide_zero_only_provider_conflicts(self):
        from schedule_app.services.teaching_validation import TeachingConflictError
        scan, *_ = scan_cells({('NYES', 'B6'): 'Example, Avery ~ Learner',
                              ('WARD A', 'B6'): 'Example, Avery ~ ',
                              ('NYES', 'C6'): 'Excluded, Adolescent ~ ',
                              ('WARD A', 'C6'): 'Excluded, Adolescent ~ '})
        with self.assertRaises(TeachingConflictError) as caught:
            teaching_build_zip(scan, [2026])
        self.assertEqual(caught.exception.conflict_count,2)
        self.assertEqual({r['preceptor_name'] for r in caught.exception.rows},
                         {'Example, Avery','Excluded, Adolescent'})


    def test_annual_educational_totals_unchanged_by_visibility(self):
        raw_total = sum(r['no_of_shifts'] for r in self.scan['monthly'])
        self.assertEqual(sum(r['no_of_shifts'] for r in self.annual), raw_total)
        self.assertEqual(sum(r['no_of_shifts'] for r in teaching_work_type_rows(self.scan, [2026])), raw_total)
        # Inclusion logic can iterate over a generator without discarding it early.
        self.assertEqual(len(participating_reach_rows(self.scan, (y for y in [2026]))), 2)


if __name__ == '__main__':
    unittest.main()
