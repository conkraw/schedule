"""Simplified-time contract: earned hours do not depend on simultaneous learners."""
import copy
import csv
from datetime import date
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile
import unittest
from docx import Document

from helpers import st, run_app, secret_settings, FakeGitHub
from test_learner_reach import scan_cells
from test_reporting_dates import small_opd, all_doc_text
from test_strict_reach_charts import ui_values
from schedule_app.services.opd_archive import GitHubOPDArchive, get_opd_archive_config, OPDArchiveError
from schedule_app.services.teaching_analysis import teaching_scan_archives, teaching_filter_date_range, teaching_annual_rows, teaching_work_type_rows
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.educational_time import (with_educational_time, teaching_time_rows, TIME_CSV_COLUMNS,
    TIME_WORK_TYPE_CSV_COLUMNS, TIME_MONTHLY_CSV_COLUMNS)
from schedule_app.services.report_diagnostics import REPORT_BUILD_ID, ReportDataError
from schedule_app.services.teaching_validation import TeachingConflictError
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.chair_summary import teaching_chair_summary_data
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME

NAME = 'Example, Avery'
START = date(2026,9,7)

def ward_cells(weekend=False):
    cells = {('WARD A',f'{col}{row}'): f'{NAME} ~ Learner A; Learner B'
             for col in 'BCDEF' for row in (6,8)}
    if weekend:
        cells['WARD A','G6'] = f'{NAME} ~ '
        cells['WARD A','H8'] = f'{NAME} ~ Learner A; Learner B'
    return cells

def csv_rows(blob, filename):
    with ZipFile(BytesIO(blob)) as z:
        return list(csv.DictReader(StringIO(z.read(filename).decode('utf-8-sig'))))

class SimpleTimeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scan,*_=scan_cells(ward_cells())
        cls.row=teaching_annual_rows(cls.scan,[2026])[0]
        cls.blob,_=teaching_build_zip(cls.scan,[2026])

    def test_two_students_five_full_days_is_forty_educational_hours(self):
        self.assertEqual(self.row['educational_hours'],40)
        self.assertEqual(self.row['total_scheduled_availability_hours'],40)
        self.assertEqual(self.row['no_of_shifts'],20)  # internal assignment audit only
        self.assertEqual(self.row['teaching_shifts'],10)

    def test_continuity_counts_stay_two_and_two(self):
        row=teaching_time_rows(self.scan,[2026])[0]
        self.assertEqual(row['unique_students'],2)
        self.assertNotIn('unique_students_3plus_days',row)
        self.assertEqual(row['learner_reach_pct'],100.0)

    def test_all_student_counts_have_same_hours_for_one_halfday(self):
        for n in (1,2,3,10):
            with self.subTest(n=n):
                names='; '.join(f'Learner {i}' for i in range(n))
                scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ {names}'})
                row=teaching_time_rows(scan,[2026])[0]
                self.assertEqual((row['educational_hours'],row['scheduled_shifts'],row['unique_students']),(4,1,n))

    def test_saturday_and_sunday_are_included(self):
        scan,*_=scan_cells(ward_cells(weekend=True))
        row=teaching_time_rows(scan,[2026])[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours'],row['learner_reach_pct']), (48,44,91.7))

    def test_weekend_only_range_uses_both_endpoints(self):
        scan,*_=scan_cells(ward_cells(weekend=True))
        view=teaching_filter_date_range(scan,ReportingPeriod('Weekend',date(2026,9,12),date(2026,9,13)))
        row=teaching_time_rows(view,[2026])[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours'],row['learner_reach_pct']), (8,4,50.0))

    def test_two_halfdays_are_eight_hours_but_one_continuity_day(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A; B', ('WARD A','B8'):f'{NAME} ~ A; B'})
        row=teaching_time_rows(scan,[2026])[0]
        self.assertEqual((row['educational_hours'],row['unique_students']),(8,2))
        self.assertNotIn('unique_students_3plus_days',row)

    def test_duplicate_provider_rows_do_not_increase_hours(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A',('WARD A','B7'):f'{NAME} ~ B'})
        row=teaching_time_rows(scan,[2026])[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours'],row['unique_students']),(4,4,2))

    def test_overlapping_rotation_snapshots_do_not_duplicate_hours(self):
        st.reset(secrets=secret_settings());repo=FakeGitHub()
        client=GitHubOPDArchive(get_opd_archive_config(),transport=repo)
        raw=small_opd(START,{('WARD A','B6'):f'{NAME} ~ A; B'})
        client.save(raw);client.save(raw)
        scan=teaching_scan_archives(client)
        self.assertEqual(teaching_time_rows(scan,[2026])[0]['educational_hours'],4)

    def test_outpatient_priority_excludes_nursery_learner_hours(self):
        scan,*_=scan_cells({('PSHCH_NURSERY','B6'):f'{NAME} ~ Nursery A; Nursery B',
                           ('NYES','B6'):f'{NAME} ~ Clinic A; Clinic B'})
        rows=teaching_time_rows(scan,[2026],by_work_type=True)
        self.assertEqual([(r['work_type'],r['educational_hours']) for r in rows],[('Academic Pediatrics',4)])
        self.assertEqual(teaching_time_rows(scan,[2026])[0]['unique_students'],2)

    def test_empty_outpatient_does_not_inherit_nursery_students(self):
        scan,*_=scan_cells({('PSHCH_NURSERY','B6'):f'{NAME} ~ Nursery A',
                           ('NYES','B6'):f'{NAME} ~ ',('NYES','C6'):f'{NAME} ~ Clinic A'})
        row=teaching_time_rows(scan,[2026])[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours'],row['learner_reach_pct'],row['unique_students']),(8,4,50,1))

    def test_other_conflicts_still_block(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A',('NYES','B6'):f'{NAME} ~ B'})
        with self.assertRaises(TeachingConflictError):teaching_build_zip(scan,[2026])

    def test_both_period_boundaries_and_cross_july(self):
        st.reset(secrets=secret_settings());repo=FakeGitHub();client=GitHubOPDArchive(get_opd_archive_config(),transport=repo)
        client.save(small_opd(date(2026,6,29),{('WARD A','B6'):f'{NAME} ~ A; B',('WARD A','C6'):f'{NAME} ~ A; B',('WARD A','D6'):f'{NAME} ~ A; B',('WARD A','E6'):f'{NAME} ~ A; B'}))
        view=teaching_filter_date_range(teaching_scan_archives(client),ReportingPeriod('Custom',date(2026,6,30),date(2026,7,1)))
        row=teaching_time_rows(view,[2026])[0]
        self.assertEqual((row['academic_year'],row['educational_hours']),('Custom',8))
        monthly=teaching_time_rows(view,[2026],by_work_type=True,monthly=True)
        self.assertEqual([r['educational_hours'] for r in monthly],[4,4])

    def test_zero_teaching_month_retained_for_included_provider(self):
        st.reset(secrets=secret_settings());repo=FakeGitHub();client=GitHubOPDArchive(get_opd_archive_config(),transport=repo)
        client.save(small_opd(START,{('WARD A','B6'):f'{NAME} ~ A; B'}))
        client.save(small_opd(date(2026,10,5),{('WARD A','B6'):f'{NAME} ~ '}))
        scan=teaching_scan_archives(client)
        months=teaching_time_rows(scan,[2026],by_work_type=True,monthly=True)
        self.assertEqual([(r['educational_hours'],r['total_scheduled_availability_hours'],r['learner_reach_pct']) for r in months],[(4,4,100),(0,4,0)])
        self.assertEqual(teaching_time_rows(scan,[2026])[0]['learner_reach_pct'],50)

    def test_zero_teaching_people_and_settings_hidden(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A',('ADOLMED','B6'):'Unused, Provider ~ '})
        rows=teaching_time_rows(scan,[2026],by_work_type=True)
        self.assertEqual({r['preceptor_name'] for r in rows},{NAME})
        self.assertEqual({r['work_type'] for r in rows},{'Ward A'})

    def test_hidden_service_availability_kept_for_active_person(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A; B',('ADOLMED','C6'):f'{NAME} ~ '})
        row=teaching_time_rows(scan,[2026])[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours']),(8,4))
        self.assertEqual(len(teaching_time_rows(scan,[2026],by_work_type=True)),1)

    def test_work_type_educational_hours_add_to_overall(self):
        cells=ward_cells();cells['NYES','G6']=f'{NAME} ~ A; B';cells['COMPLEX','H6']=f'{NAME} ~ A; B; C'
        scan,*_=scan_cells(cells)
        all_=teaching_time_rows(scan,[2026])[0];types=teaching_time_rows(scan,[2026],by_work_type=True)
        self.assertEqual(all_['educational_hours'],48)
        self.assertEqual(sum(r['educational_hours'] for r in types),48)

    def test_area_totals_sum_preceptor_time_not_global_date_union(self):
        scan,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ A; B',('WARD A','B7'):'Other, Blair ~ C'})
        data=teaching_chair_summary_data(scan,[2026])[0]
        self.assertEqual((data['educational_hours'],data['recorded_clinical_hours']),(8,8))

    def test_report_csvs_match_the_simple_schema(self):
        with ZipFile(BytesIO(self.blob)) as z:
            for fn,cols in [('preceptor_teaching_summary.csv',TIME_CSV_COLUMNS),('preceptor_teaching_by_work_type.csv',TIME_WORK_TYPE_CSV_COLUMNS),('preceptor_learner_reach_monthly.csv',TIME_MONTHLY_CSV_COLUMNS)]:
                with self.subTest(file=fn):
                    reader=csv.DictReader(StringIO(z.read(fn).decode('utf-8-sig')))
                    self.assertEqual(tuple(reader.fieldnames),cols)
                    row=next(reader)
                    self.assertEqual(row['educational_hours'],'40')
                    self.assertEqual(row['total_scheduled_availability_hours'],'40')
                    self.assertNotIn('no_of_shifts',row)

    def test_chart_csv_uses_same_values_as_the_report(self):
        row=csv_rows(self.blob,'clinical_experience_learner_reach.csv')[0]
        self.assertEqual((row['total_scheduled_availability_hours'],row['educational_hours'],row['learner_reach_pct']),('40','40','100.0'))

    def test_docs_show_no_student_weighted_hour_metric(self):
        with ZipFile(BytesIO(self.blob)) as z:
            for fn in z.namelist():
                if fn.endswith('.docx'):
                    text=all_doc_text(z.read(fn))
                    for old in ('OPD hours','Student-shifts','student-weighted','80 educational hours'):
                        self.assertNotIn(old,text)
                    self.assertIn('Total scheduled',text)
                    self.assertIn('Educational',text)
                    self.assertIn('weekends',text)
                    self.assertNotIn('Learner A',text)
                    self.assertNotIn('Learner B',text)

    def test_individual_top_table_is_three_metrics(self):
        with ZipFile(BytesIO(self.blob)) as z:
            fn=next(f for f in z.namelist() if f.startswith('Preceptor_Reports/') and f.endswith('.docx'))
            table=Document(BytesIO(z.read(fn))).tables[0]
        self.assertEqual([[c.text for c in r.cells] for r in table.rows], [
            ['Measure','Result'],['Total scheduled availability','40 hours'],['Educational hours','40 hours'],['Learner Reach','100.0%']])

    def test_missing_or_inconsistent_counts_are_not_guessed(self):
        for row in ({'educational_hours':40},{'recorded_clinical_shifts':10,'shifts_with_students':11,'shifts_without_students':0}):
            with self.assertRaises(ReportDataError):with_educational_time(row)

    def test_legacy_weighted_field_is_recalculated_from_valid_distinct_counts(self):
        row=with_educational_time({'recorded_clinical_shifts':10,'shifts_with_students':8,'shifts_without_students':2,'educational_hours':64})
        self.assertEqual((row['educational_hours'],row['total_scheduled_availability_hours'],row['learner_reach_pct']),(32,40,80))

    def test_generation_does_not_mutate_source_scan(self):
        before=copy.deepcopy(self.scan)
        teaching_build_zip(self.scan,[2026])
        self.assertEqual(before,self.scan)

    def test_zip_integrity_and_shared_totals(self):
        with ZipFile(BytesIO(self.blob)) as z:self.assertIsNone(z.testzip())
        monthly=csv_rows(self.blob,'preceptor_learner_reach_monthly.csv')
        annual=csv_rows(self.blob,'preceptor_teaching_summary.csv')
        self.assertEqual(sum(int(r['educational_hours']) for r in monthly),sum(int(r['educational_hours']) for r in annual))

    def test_interface_has_simple_metrics_no_weighted_metric(self):
        scan,repo,client,secrets,_=scan_cells(ward_cells())
        result=run_app(ui_values(pts_show_diagnostics=True),secrets=secrets,repo=repo, evaluation_login=True)
        msgs='\n'.join(text for _,text in result['messages'])
        self.assertNotIn('Student-weighted educational hours',msgs)
        self.assertIn('Total scheduled availability',msgs)
        self.assertIn(REPORT_BUILD_ID,msgs)

    def test_settings_archive_and_oasis_runtime_are_unchanged(self):
        root=Path(__file__).resolve().parents[1];base=root.parent/'base'
        if not base.exists():self.skipTest('Local baseline packaging comparison only')
        paths=['app_sch_2026.py','requirements.txt','schedule_app/settings.py']
        paths += [str(p.relative_to(root)) for p in (root/'schedule_app').rglob('*.py')
                  if p.name.startswith(('oasis_','preceptor_oasis')) or p.name in ('opd_archive.py','preceptor_evaluations.py','teaching_priority.py')]
        # The later missing-username update intentionally changes this UI only.
        # All persistence, OASIS calculation, archive, and priority modules remain
        # covered by the original immutability check.
        paths = [rel for rel in paths if rel != 'schedule_app/sections/preceptor_oasis_links.py']
        for rel in paths:
            with self.subTest(path=rel):
                current=(root/rel).read_bytes()
                previous=(base/rel).read_bytes()
                if rel=='schedule_app/reports/preceptor_evaluations.py':
                    # The existing feedback renderer is unchanged; the release
                    # appends a missing-feedback explanation helper only.
                    current=current.split(b'\n\ndef append_feedback_unavailable')[0].rstrip()+b'\n'
                    previous=previous.rstrip()+b'\n'
                self.assertEqual(current,previous)

if __name__=='__main__':unittest.main()
