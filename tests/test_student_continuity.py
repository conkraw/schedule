"""Unique learners on distinct dates; synthetic OPDs and offline GitHub only."""
from copy import deepcopy
from datetime import date
from io import BytesIO
import json
import unittest
from zipfile import ZipFile

from helpers import st, FakeGitHub, secret_settings, run_app, make_opd
from test_reporting_dates import small_opd, all_doc_text
from test_learner_reach import scan_cells
from test_strict_reach_charts import ui_values
from schedule_app.services.opd_archive import GitHubOPDArchive, get_opd_archive_config, OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import (
    teaching_scan_archives, teaching_filter_date_range, teaching_annual_rows,
)
from schedule_app.services.student_continuity import (
    require_student_continuity_data, student_continuity_counts,
)
from schedule_app.services.teaching_validation import TeachingConflictError
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.teaching_export import teaching_build_zip

NAME = 'Example, Avery'

def assigned(student='Private Student Alpha'):
    return f'{NAME} ~ {student}'


def exact_example():
    return {('HOPE_DRIVE', 'B8'): assigned(),    # Monday PM
            ('HOPE_DRIVE', 'C6'): assigned(),    # Tuesday AM
            ('HOPE_DRIVE', 'E8'): assigned()}    # Thursday PM


def document(scan, name=NAME):
    return teaching_make_docx(name, [r for r in scan['monthly'] if r['preceptor_name']==name], scan)


def count_pair(scan, name=NAME, year=2026):
    counts = student_continuity_counts(scan, name, year)
    return counts['unique_students'], counts['unique_students_3plus_days']


class StudentContinuityTests(unittest.TestCase):
    def test_three_students_one_date_each_do_not_reach_three_days(self):
        scan,*_=scan_cells({('NYES','B8'):assigned('Alpha'),('NYES','C6'):assigned('Beta'),('NYES','E8'):assigned('Gamma')})
        self.assertEqual(count_pair(scan),(3,0))

    def test_ui_stale_student_details_clear_old_downloads(self):
        scan,repo,client,secrets,_=scan_cells(exact_example())
        first=run_app(ui_values(),secrets=secrets,repo=repo, evaluation_login=True)
        state=deepcopy(first['state'])
        state['teaching_scan'].pop('student_continuity_version')
        result=run_app(ui_values(teaching_load_archives=False,teaching_build_zip=False),
                       secrets=secrets,repo=repo,state=state, evaluation_login=True)
        self.assertNotIn('teaching_scan',result['state'])
        self.assertFalse(any(name.endswith(('.zip','.docx')) for name in result['downloads']))
        self.assertTrue(any('unique-student' in message for _,message in result['messages']))

    def test_user_monday_pm_tuesday_am_thursday_pm_example(self):
        scan,*_=scan_cells(exact_example())
        self.assertEqual(count_pair(scan), (1,1))

    def test_three_shifts_over_two_days_do_not_qualify(self):
        scan,*_=scan_cells({('NYES','B6'): assigned(), ('NYES','B8'): assigned(), ('NYES','C6'): assigned()})
        self.assertEqual(count_pair(scan), (1,0))

    def test_both_am_and_pm_still_one_day(self):
        scan,*_=scan_cells({('NYES','B6'): assigned(), ('NYES','B8'): assigned()})
        self.assertEqual(count_pair(scan), (1,0))
        self.assertEqual(scan['student_day_groups'][0]['dates'], ['2026-09-07'])

    def test_two_students_in_same_shift_are_two_unique_students(self):
        scan,*_=scan_cells({('NYES','B6'): assigned('Private Alpha; Private Beta')})
        self.assertEqual(count_pair(scan), (2,0))
        self.assertEqual(len(scan['student_day_groups']), 2)
        # Identical date groups represent different people and must NOT collapse.
        self.assertEqual(scan['student_day_groups'][0], scan['student_day_groups'][1])

    def test_two_students_on_identical_three_dates_both_qualify(self):
        scan,*_=scan_cells({key: assigned('Private Alpha; Private Beta') for key in exact_example()})
        self.assertEqual(count_pair(scan), (2,2))

    def test_multiple_preceptors_count_their_own_unique_students(self):
        cells=exact_example();cells['ETOWN','F6']='Other, Preceptor ~ Private Student Alpha'
        scan,*_=scan_cells(cells)
        self.assertEqual(count_pair(scan), (1,1))
        self.assertEqual(count_pair(scan,'Other, Preceptor'), (1,0))

    def test_distinct_days_across_work_types_combine(self):
        scan,*_=scan_cells({('HOPE_DRIVE','B8'):assigned(),('WARD A','C6'):assigned(),('COMPLEX','E8'):assigned()})
        self.assertEqual(count_pair(scan), (1,1))

    def test_same_date_am_pm_in_different_types_remains_one_day(self):
        scan,*_=scan_cells({('NYES','B6'):assigned(),('WARD A','B8'):assigned(),('COMPLEX','C6'):assigned()})
        self.assertEqual(count_pair(scan), (1,0))

    def test_weekends_count_as_distinct_days(self):
        scan,*_=scan_cells({('NYES','F8'):assigned(),('NYES','G6'):assigned(),('NYES','H8'):assigned()})
        self.assertEqual(count_pair(scan), (1,1))

    def test_duplicate_rows_and_academic_sites_not_extra_people(self):
        cells=exact_example();cells.update({('ETOWN','B8'):assigned(),('NYES','B8'):assigned(),('HOPE_DRIVE','C7'):assigned()})
        scan,*_=scan_cells(cells)
        self.assertEqual(count_pair(scan), (1,1))
        self.assertEqual(len(scan['student_day_groups'][0]['dates']),3)

    def test_blank_fields_do_not_create_students(self):
        cells={('NYES','B6'):assigned(' '),('NYES','C6'):assigned('None'),('NYES','D6'):assigned('TBD')}
        scan,*_=scan_cells(cells)
        self.assertEqual(count_pair(scan),(0,0))
        self.assertEqual(scan['student_day_groups'],[])

    def test_name_normalization_ignores_case_spaces_comma_spacing(self):
        scan,*_=scan_cells({('NYES','B6'):assigned('Student, Alpha'),('NYES','C6'):assigned(' student ,  alpha '),('NYES','D8'):assigned('STUDENT,ALPHA')})
        self.assertEqual(count_pair(scan),(1,1))

    def test_different_student_names_not_fuzzy_merged(self):
        scan,*_=scan_cells({('NYES','B6'):assigned('Student, Alpha'),('NYES','C6'):assigned('Student, A.')})
        self.assertEqual(count_pair(scan),(2,0))

    def test_reversed_name_order(self):
        cells={key:'Private Student Alpha ~ Example, Avery' for key in exact_example()}
        scan,*_=scan_cells(cells, order='Student ~ Preceptor')
        self.assertEqual(count_pair(scan),(1,1))

    def test_custom_filter_applies_before_threshold(self):
        scan,*_=scan_cells(exact_example())
        view=teaching_filter_date_range(scan,ReportingPeriod('Selected',date(2026,9,8),date(2026,9,10)))
        self.assertEqual(count_pair(view),(1,0))
        self.assertEqual(count_pair(scan),(1,1))

    def test_custom_filter_removes_student_with_only_outside_dates(self):
        cells=exact_example();cells['NYES','H6']=assigned('Private Outside Student')
        scan,*_=scan_cells(cells)
        view=teaching_filter_date_range(scan,ReportingPeriod('Selected',date(2026,9,7),date(2026,9,10)))
        self.assertEqual(count_pair(scan),(2,1))
        self.assertEqual(count_pair(view),(1,1))

    def test_exact_start_end_inclusive(self):
        scan,*_=scan_cells(exact_example())
        view=teaching_filter_date_range(scan,ReportingPeriod('Selected',date(2026,9,7),date(2026,9,10)))
        self.assertEqual(count_pair(view),(1,1))

    def test_one_day_range_never_reaches_three_days(self):
        scan,*_=scan_cells(exact_example())
        view=teaching_filter_date_range(scan,ReportingPeriod('One day',date(2026,9,8),date(2026,9,8)))
        self.assertEqual(count_pair(view),(1,0))

    def test_custom_dates_cross_july_no_year_reset(self):
        scan,*_=scan_cells(exact_example(),start=date(2026,6,29))
        view=teaching_filter_date_range(scan,ReportingPeriod('26-27',date(2026,6,29),date(2026,7,2)))
        self.assertEqual(count_pair(view),(1,1))
        self.assertEqual(count_pair(scan,year=2025),(1,0))
        self.assertEqual(count_pair(scan,year=2026),(1,0))

    def test_across_rotations_months_and_more_than_one_year(self):
        scan,repo,client,*_=scan_cells({('NYES','B6'):assigned()},start=date(2026,2,2))
        client.save(small_opd(date(2026,8,3),{('ETOWN','C8'):assigned()}))
        client.save(small_opd(date(2027,3,1),{('HOPE_DRIVE','D6'):assigned()}))
        scan=teaching_scan_archives(client)
        view=teaching_filter_date_range(scan,ReportingPeriod('Cohort',date(2026,2,1),date(2027,3,31)))
        self.assertEqual(count_pair(view),(1,1))

    def test_overlapping_archives_count_same_learner_once(self):
        scan,repo,client,*_=scan_cells({('NYES','B6'):assigned()},start=date(2026,9,14))
        client.save(make_opd(date(2026,9,7),{('NYES','B30'):assigned()}))
        scan=teaching_scan_archives(client)
        self.assertEqual(count_pair(scan),(1,0))
        self.assertEqual(next(r for r in scan['student_day_groups'] if r['preceptor_name']==NAME)['dates'],['2026-09-14'])

    def test_nursery_overlap_excludes_nursery_student(self):
        cells=exact_example();cells['PSHCH_NURSERY','B8']=assigned('Private Nursery Student')
        scan,*_=scan_cells(cells)
        self.assertEqual(count_pair(scan),(1,1))

    def test_excluded_nursery_date_cannot_push_learner_over_threshold(self):
        scan,*_=scan_cells({('PSHCH_NURSERY','B8'):assigned(),('NYES','B8'):assigned(' '),
                           ('PSHCH_NURSERY','C6'):assigned(),('PSHCH_NURSERY','E8'):assigned()})
        self.assertEqual(count_pair(scan),(1,0))

    def test_blank_clinic_does_not_inherit_nursery_student(self):
        scan,*_=scan_cells({('PSHCH_NURSERY','B8'):assigned(),('NYES','B8'):assigned(' ')})
        self.assertEqual(count_pair(scan),(0,0))
        self.assertEqual(scan['student_day_groups'],[])

    def test_other_conflicts_still_block_individual_and_zip(self):
        scan,*_=scan_cells({('NYES','B8'):assigned(),('WARD A','B8'):assigned()})
        with self.assertRaises(TeachingConflictError):document(scan)
        with self.assertRaises(TeachingConflictError):teaching_build_zip(scan,[2026])

    def test_missing_schema_requests_rescan_not_zero(self):
        scan,*_=scan_cells(exact_example());scan.pop('student_continuity_version')
        with self.assertRaisesRegex(OPDArchiveError,'refresh archived OPDs'):document(scan)
        with self.assertRaisesRegex(OPDArchiveError,'refresh archived OPDs'):teaching_build_zip(scan,[2026])

    def test_malformed_or_inconsistent_date_groups_rejected(self):
        scan,*_=scan_cells(exact_example())
        bad_dates=[[],['2026-09-07','2026-09-07'],['2026-09-10','2026-09-07'],['not-date'],['2026-10-01']]
        for days in bad_dates:
            with self.subTest(days=days):
                bad=deepcopy(scan);bad['student_day_groups'][0]['dates']=days
                with self.assertRaises(OPDArchiveError):require_student_continuity_data(bad)

    def test_counts_and_date_filter_do_not_mutate_full_scan(self):
        scan,*_=scan_cells(exact_example());before=deepcopy(scan)
        count_pair(scan)
        view=teaching_filter_date_range(scan,ReportingPeriod('Selected',date(2026,9,8),date(2026,9,10)))
        count_pair(view)
        self.assertEqual(scan,before)
        view['student_day_groups'][0]['dates'].clear()
        self.assertEqual(scan,before)

    def test_scan_retains_neither_names_nor_linkable_student_ids(self):
        scan,*_=scan_cells(exact_example())
        self.assertNotIn('Private Student Alpha',json.dumps(scan))
        for group in scan['student_day_groups']:
            self.assertEqual(set(group),{'preceptor_name','dates'})
        self.assertNotIn('student_key',json.dumps(scan))

    def test_rescan_rebuilds_same_counts_despite_fresh_salt(self):
        scan,_,client,*_=scan_cells(exact_example());second=teaching_scan_archives(client)
        self.assertEqual(scan['student_day_groups'],second['student_day_groups'])
        self.assertEqual(count_pair(scan),count_pair(second))

    def test_individual_doc_contains_both_metrics_and_definition(self):
        scan,*_=scan_cells(exact_example())
        text=all_doc_text(document(scan))
        self.assertIn('Unique students assigned: 1',text)
        self.assertNotIn('Students assigned on 3+ days',text)
        self.assertIn('AM and PM on the same date are two shifts',text)
        self.assertNotIn('Private Student Alpha',text)

    def test_standard_year_doc_uses_each_year_counts_not_whole_archive(self):
        scan,*_=scan_cells(exact_example(),start=date(2026,6,29))
        text=all_doc_text(document(scan))
        self.assertEqual(text.count('Unique students assigned: 1'),2)
        self.assertEqual(text.count('Students assigned on 3+ days'),0)

    def test_zip_contains_metrics_only_not_student_data_or_date_groups(self):
        scan,*_=scan_cells(exact_example());payload,_=teaching_build_zip(scan,[2026])
        with ZipFile(BytesIO(payload)) as zf:
            for path in zf.namelist():
                raw=zf.read(path)
                text=all_doc_text(raw) if path.endswith('.docx') else raw.decode('utf-8','ignore')
                self.assertNotIn('Private Student Alpha',text)
                self.assertNotIn('student_day_groups',text)
                if path.startswith('Preceptor_Reports/'):
                    self.assertIn('Unique students assigned: 1',text)
                    self.assertNotIn('Students assigned on 3+ days',text)

    def test_saving_reports_does_not_write_to_github(self):
        scan,repo,*_=scan_cells(exact_example());before=repo.write_count;tree=dict(repo.tree)
        teaching_build_zip(scan,[2026]);self.assertEqual(repo.write_count,before);self.assertEqual(repo.tree,tree)

    def test_ui_load_build_custom_date_workflow(self):
        scan,repo,client,secrets,_=scan_cells(exact_example())
        result=run_app(ui_values(),secrets=secrets,repo=repo, evaluation_login=True)
        self.assertFalse([m for kind,m in result['messages'] if kind == 'error'])
        payload=next(raw for name,raw in result['downloads'].items() if name.endswith('.zip'))
        with ZipFile(BytesIO(payload)) as zf:
            text=all_doc_text(zf.read(next(n for n in zf.namelist() if n.startswith('Preceptor_Reports/'))))
            self.assertIn('Unique students assigned: Not checked',text)
            self.assertIn('Students assigned for 3+ shifts: Not checked',text)

if __name__=='__main__':
    unittest.main()
