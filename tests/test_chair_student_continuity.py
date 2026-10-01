"""Chair continuity additions: offline fixtures, exact dates, no real students."""
from copy import deepcopy
from datetime import date
from io import BytesIO
import json
import unittest
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from helpers import run_app
from test_learner_reach import scan_cells
from test_reporting_dates import all_doc_text
from test_student_continuity import assigned, exact_example, NAME
from test_strict_reach_charts import ui_values
from schedule_app.reports.chair_summary import teaching_chair_summary_data, teaching_make_chair_summary
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.services.student_continuity import student_continuity_counts
from schedule_app.services.teaching_analysis import teaching_filter_date_range, teaching_annual_rows
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_validation import TeachingConflictError
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME


def summary_row(scan, name=NAME, year=2026):
    data = teaching_chair_summary_data(scan, [year])[0]
    return next(row for row in data['named_preceptors'] + data['unresolved_labels']
                if row['preceptor_name'] == name)


def continuity_tables(raw):
    return [table for table in Document(BytesIO(raw)).tables
            if [cell.text for cell in table.rows[0].cells][1:]
            == ['Unique students', 'Students meeting minimum shifts']]


class ChairStudentContinuityTests(unittest.TestCase):
    def test_matches_individual_counts(self):
        scan, *_ = scan_cells(exact_example())
        row = summary_row(scan)
        for field, count in student_continuity_counts(scan, NAME, 2026).items():
            self.assertEqual(row[field], count)

    def test_three_nonconsecutive_dates_qualify(self):
        scan, *_ = scan_cells(exact_example())
        row = summary_row(scan)
        self.assertEqual((row['unique_students'], row['unique_students_3plus_days']), (1, 1))

    def test_am_pm_not_two_dates(self):
        scan, *_ = scan_cells({('NYES', 'B6'): assigned(), ('NYES', 'B8'): assigned(),
                              ('NYES', 'C6'): assigned()})
        self.assertEqual(summary_row(scan)['unique_students_3plus_days'], 0)

    def test_simultaneous_students_remain_unique(self):
        scan, *_ = scan_cells({key: assigned('Private Alpha; Private Beta') for key in exact_example()})
        row = summary_row(scan)
        self.assertEqual((row['unique_students'], row['unique_students_3plus_days']), (2, 2))

    def test_weekend_dates_count(self):
        scan, *_ = scan_cells({('NYES', 'F8'): assigned(), ('NYES', 'G6'): assigned(),
                              ('NYES', 'H8'): assigned()})
        self.assertEqual(summary_row(scan)['unique_students_3plus_days'], 1)

    def test_counts_combine_settings_not_sum_of_type_unique_counts(self):
        scan, *_ = scan_cells({('NYES', 'B6'): assigned(), ('WARD A', 'C6'): assigned(),
                              ('COMPLEX', 'D6'): assigned()})
        row = summary_row(scan)
        self.assertEqual((row['unique_students'], row['unique_students_3plus_days']), (1, 1))
        # One overall row for this person despite three work-type appearances.
        tables = continuity_tables(teaching_make_chair_summary(scan, [2026]))
        self.assertEqual(len(tables), 1)
        self.assertEqual(len(tables[0].rows), 2)

    def test_boundary_days_are_filtered_before_threshold(self):
        scan, *_ = scan_cells(exact_example())
        view = teaching_filter_date_range(scan, ReportingPeriod('Partial', date(2026, 9, 8), date(2026, 9, 10)))
        row = summary_row(view)
        self.assertEqual((row['unique_students'], row['unique_students_3plus_days']), (1, 0))

    def test_custom_year_spanning_july_is_one_period(self):
        scan, *_ = scan_cells(exact_example(), start=date(2026, 6, 29))
        view = teaching_filter_date_range(scan, ReportingPeriod('Cohort', date(2026, 6, 29), date(2026, 7, 2)))
        self.assertEqual(summary_row(view)['unique_students_3plus_days'], 1)
        self.assertEqual(len(continuity_tables(teaching_make_chair_summary(view, [2026]))), 1)

    def test_standard_year_sections_do_not_share_threshold_days(self):
        scan, *_ = scan_cells(exact_example(), start=date(2026, 6, 29))
        data = teaching_chair_summary_data(scan, [2025, 2026])
        self.assertEqual([row['named_preceptors'][0]['unique_students_3plus_days'] for row in data], [0, 0])
        self.assertEqual(len(continuity_tables(teaching_make_chair_summary(scan, [2025, 2026]))), 2)

    def test_outpatient_priority_excludes_nursery_students(self):
        cells = exact_example()
        cells['PSHCH_NURSERY', 'B8'] = assigned('Private Nursery Only')
        scan, *_ = scan_cells(cells)
        self.assertEqual(summary_row(scan)['unique_students'], 1)

    def test_blanks_do_not_inherit_nursery_student(self):
        scan, *_ = scan_cells({('PSHCH_NURSERY', 'B8'): assigned(), ('NYES', 'B8'): assigned(' '),
                              ('PSHCH_NURSERY', 'C6'): assigned(), ('PSHCH_NURSERY', 'E8'): assigned()})
        self.assertEqual(summary_row(scan)['unique_students_3plus_days'], 0)

    def test_zero_assignment_preceptors_not_in_continuity(self):
        cells = exact_example()
        cells['ADOLMED', 'G6'] = 'Unassigned, Preceptor ~ '
        scan, *_ = scan_cells(cells)
        raw = teaching_make_chair_summary(scan, [2026])
        self.assertNotIn('Unassigned, Preceptor', all_doc_text(raw))
        self.assertNotIn('ADOLMED', all_doc_text(raw))

    def test_placeholders_get_separate_label_table(self):
        cells = exact_example()
        cells['SJR_HOSP', 'B6'] = 'SJR_1 ~ Private Other Student'
        scan, *_ = scan_cells(cells)
        raw = teaching_make_chair_summary(scan, [2026])
        tables = continuity_tables(raw)
        self.assertEqual([t.rows[0].cells[0].text for t in tables], ['Preceptor', 'Provider label'])
        self.assertEqual(tables[1].rows[1].cells[0].text, 'SJR_1')

    def test_no_sum_of_unique_students_across_preceptors(self):
        cells = exact_example()
        cells['NYES', 'F8'] = 'Other, Preceptor ~ Private Student Alpha'
        scan, *_ = scan_cells(cells)
        raw = teaching_make_chair_summary(scan, [2026])
        table = continuity_tables(raw)[0]
        self.assertEqual(len(table.rows), 3)  # heading + two people, no sum row
        self.assertIn('Do not add student counts across preceptors', all_doc_text(raw))

    def test_both_columns_and_clear_all_work_types_scope(self):
        scan, *_ = scan_cells(exact_example())
        raw = teaching_make_chair_summary(scan, [2026])
        text = all_doc_text(raw)
        self.assertIn('Students assigned by preceptor', text)
        self.assertIn('All work types combined for each preceptor', text)
        self.assertNotIn('Students assigned on 3+ days', text)
        table = continuity_tables(raw)[0]
        self.assertEqual([c.text for c in table.rows[1].cells], [NAME, '1', 'Not checked'])
        self.assertIsNotNone(table.rows[0]._tr.find('w:trPr/w:tblHeader', namespaces=table.rows[0]._tr.nsmap))
        self.assertTrue(all(row._tr.find('w:trPr/w:cantSplit', namespaces=row._tr.nsmap) is not None for row in table.rows))

    def test_no_students_or_date_groups_in_chair(self):
        scan, *_ = scan_cells(exact_example())
        text = all_doc_text(teaching_make_chair_summary(scan, [2026]))
        self.assertNotIn('Private Student Alpha', text)
        self.assertNotIn('student_day_groups', text)
        self.assertNotIn('student_key', text)

    def test_stale_or_incomplete_scan_not_silently_zero(self):
        scan, *_ = scan_cells(exact_example())
        for field in ('student_continuity_version', 'student_day_groups'):
            bad = deepcopy(scan)
            bad.pop(field)
            with self.subTest(field=field), self.assertRaisesRegex(OPDArchiveError, 'refresh archived OPDs'):
                teaching_make_chair_summary(bad, [2026])

    def test_conflicts_still_block_chair(self):
        scan, *_ = scan_cells({('NYES', 'B6'): assigned(), ('WARD A', 'B6'): assigned()})
        with self.assertRaises(TeachingConflictError):
            teaching_make_chair_summary(scan, [2026])

    def test_pies_and_existing_totals_unchanged(self):
        scan, *_ = scan_cells(exact_example())
        data = teaching_chair_summary_data(scan, [2026])[0]
        self.assertEqual((data['no_of_shifts'], data['educational_hours'], data['learner_reach_pct']), (3, 12, 100.0))
        raw = teaching_make_chair_summary(scan, [2026])
        self.assertEqual(len(Document(BytesIO(raw)).inline_shapes), 1)

    def test_zip_includes_chair_continuity_and_simple_csv_counts(self):
        scan, *_ = scan_cells(exact_example())
        payload, annual = teaching_build_zip(scan, [2026])
        self.assertEqual(annual, teaching_annual_rows(scan, [2026]))
        with ZipFile(BytesIO(payload)) as z:
            self.assertEqual(z.testzip(), None)
            self.assertEqual(len(continuity_tables(z.read(TEACHING_CHAIR_SUMMARY_FILENAME))), 1)
            self.assertIn('unique_students', z.read('preceptor_teaching_summary.csv').decode('utf-8-sig').splitlines()[0])
            self.assertIn('CHAIR AND INDIVIDUAL REPORTS', z.read('Report_Notes.txt').decode())

    def test_generation_does_not_mutate_scan_or_archive(self):
        scan, repo, *_ = scan_cells(exact_example())
        before = deepcopy(scan)
        tree, writes = deepcopy(repo.tree), repo.write_count
        teaching_make_chair_summary(scan, [2026])
        self.assertEqual(scan, before)
        self.assertEqual(repo.tree, tree)
        self.assertEqual(repo.write_count, writes)

    def test_ui_old_download_clears_but_current_scan_is_kept(self):
        scan, repo, client, secrets, _ = scan_cells(exact_example())
        first = run_app(ui_values(), secrets=secrets, repo=repo, evaluation_login=True)
        self.assertFalse([m for kind, m in first['messages'] if kind == 'error'])
        state = deepcopy(first['state'])
        state['teaching_zip_signature'] = state['teaching_zip_signature'][:-1]  # old output format
        state['teaching_zip'] = b'old report that must not be shown'
        second = run_app(ui_values(teaching_load_archives=False, teaching_build_zip=False),
                         secrets=secrets, repo=repo, state=state, evaluation_login=True)
        self.assertEqual(second['state']['teaching_scan'], scan)
        self.assertNotIn('teaching_zip', second['state'])
        self.assertFalse(any(n.endswith(('.zip', '.docx')) for n in second['downloads']))

    def test_ui_download_and_zip_contain_the_same_updated_word_report(self):
        scan, repo, client, secrets, _ = scan_cells(exact_example())
        result = run_app(ui_values(), secrets=secrets, repo=repo, evaluation_login=True)
        word = result['downloads'][TEACHING_CHAIR_SUMMARY_FILENAME]
        zipped = next(raw for name, raw in result['downloads'].items() if name.endswith('.zip'))
        self.assertEqual(len(continuity_tables(word)), 1)
        with ZipFile(BytesIO(zipped)) as z:
            self.assertEqual(word, z.read(TEACHING_CHAIR_SUMMARY_FILENAME))


if __name__ == '__main__':
    unittest.main()
