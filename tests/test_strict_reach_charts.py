"""Strict conflict validation and area-only pies: synthetic data, no live calls."""
import copy
import csv
from datetime import date
from io import BytesIO, StringIO
import json
from pathlib import Path
from unittest.mock import patch
import unittest
from zipfile import ZipFile

from helpers import st, run_app, FakeGitHub, secret_settings, make_opd
from test_learner_reach import scan_cells, make_reach_fixture
from test_reporting_dates import small_opd, all_doc_text
from schedule_app.services.teaching_validation import (
    TeachingConflictError, teaching_conflict_rows, validate_teaching_report, require_conflict_source_data,
)
from schedule_app.services.teaching_analysis import (
    teaching_scan_archives, teaching_filter_date_range, teaching_annual_rows, teaching_work_type_rows,
)
from schedule_app.services.opd_archive import OPDArchiveError, GitHubOPDArchive, get_opd_archive_config
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.learner_reach import reach_percent, reach_totals
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.chair_summary import teaching_make_chair_summary, teaching_chair_summary_data
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.learner_reach_charts import learner_reach_pie
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME

CONFLICT_CELLS={('NYES','B6'):'Example, Avery ~ Secret Learner',
                ('WARD A','B6'):'Example, Avery ~ ',
                ('NYES','C6'):'Example, Avery ~ Secret Learner',
                ('NYES','D6'):'Example, Avery ~ '}

def ui_values(**overrides):
    return dict({'schedule_app_mode':'PTS', 'teaching_load_archives':True,
                'teaching_build_zip':True, 'teaching_period_start':date(2026,9,7),
                'teaching_period_end':date(2026,9,13), 'teaching_period_label':'26-27'},**overrides)


class StrictReachTests(unittest.TestCase):
    def test_conflict_has_every_exact_source_cell_and_no_student_names(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        rows=teaching_conflict_rows(scan,[2026])
        self.assertEqual(len(rows),2)
        self.assertEqual({r['cell'] for r in rows},{'B6'})
        self.assertEqual({r['worksheet'] for r in rows},{'NYES','WARD A'})
        self.assertEqual({r['archive_file'] for r in rows},{'OPD_2026-09-07.xlsx.enc'})
        self.assertEqual({r['rotation_start'] for r in rows},{'2026-09-07'})
        self.assertEqual({r['student_assigned'] for r in rows},{'YES','NO'})
        self.assertNotIn('Secret Learner',json.dumps(scan))
        self.assertNotIn('Secret Learner',json.dumps(rows))

    def test_conflicts_outside_custom_period_do_not_block(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        view=teaching_filter_date_range(scan,ReportingPeriod('Only clear days',date(2026,9,8),date(2026,9,9)))
        validate_teaching_report(view,[2026])
        self.assertEqual(teaching_annual_rows(view,[2026])[0]['learner_reach_pct'],50.0)

    def test_start_and_end_dates_include_conflicts(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        for start,end in [(date(2026,9,7),date(2026,9,7)),(date(2026,9,1),date(2026,9,7))]:
            view=teaching_filter_date_range(scan,ReportingPeriod('Selected',start,end))
            with self.assertRaises(TeachingConflictError):validate_teaching_report(view,[2026])

    def test_conflicts_in_other_academic_years_do_not_block(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        self.assertEqual(teaching_conflict_rows(scan,[2025]),[])
        with self.assertRaises(TeachingConflictError):validate_teaching_report(scan,[2026])

    def test_blocks_all_outputs_before_word_creation(self):
        scan,*_=scan_cells(CONFLICT_CELLS)
        for operation in (
            lambda:teaching_annual_rows(scan,[2026]), lambda:teaching_work_type_rows(scan,[2026]),
            lambda:teaching_make_chair_summary(scan,[2026]),
            lambda:teaching_make_docx('Example, Avery',scan['monthly'],scan),
            lambda:teaching_build_zip(scan,[2026])):
            with self.subTest(operation=operation), self.assertRaises(TeachingConflictError):operation()
        with patch('schedule_app.reports.chair_summary.Document') as document:
            with self.assertRaises(TeachingConflictError):teaching_make_chair_summary(scan,[2026])
            document.assert_not_called()

    def test_both_blank_does_not_hide_a_cross_type_conflict(self):
        scan,*_=scan_cells({('NYES','B6'):'Blank, Provider ~ ',('COMPLEX','B6'):'Blank, Provider ~ ',
                           ('WARD A','C6'):'Teaching, Provider ~ Student'})
        with self.assertRaises(TeachingConflictError):teaching_build_zip(scan,[2026])

    def test_two_learners_same_time_not_a_conflict(self):
        scan,*_=scan_cells({('NYES','B6'):'Example, Avery ~ First; Second'})
        validate_teaching_report(scan,[2026])
        row=teaching_work_type_rows(scan,[2026])[0]
        self.assertEqual((row['no_of_shifts'],row['recorded_clinical_hours'],row['learner_reach_pct']),(2,4,100))

    def test_repeated_academic_sites_are_one_category(self):
        scan,*_=scan_cells({('NYES','B6'):'Example, Avery ~ First',
                           ('HOPE_DRIVE','B6'):'Example, Avery ~ First',('ETOWN','B6'):'Example, Avery ~ '})
        self.assertEqual(teaching_conflict_rows(scan,[2026]),[])
        self.assertEqual(teaching_work_type_rows(scan,[2026])[0]['learner_reach_pct'],100)

    def test_different_half_days_or_dates_are_not_conflicts(self):
        scan,*_=scan_cells({('NYES','B6'):'Example, Avery ~ First',('WARD A','B8'):'Example, Avery ~ Second',
                           ('COMPLEX','C6'):'Example, Avery ~ Third'})
        validate_teaching_report(scan,[2026])
        self.assertEqual(len(teaching_work_type_rows(scan,[2026])),3)

    def test_sources_can_span_overlapping_rotations(self):
        settings=secret_settings();st.reset(secrets=settings)
        repo=FakeGitHub();client=GitHubOPDArchive(get_opd_archive_config(),transport=repo)
        client.save(make_opd(date(2026,9,7),{('HOPE_DRIVE','B32'):'Overlap, Example ~ First'}))
        client.save(small_opd(date(2026,9,14),{('WARD A','B6'):'Overlap, Example ~ Second'}))
        scan=teaching_scan_archives(client)
        rows=teaching_conflict_rows(scan,[2026])
        self.assertEqual({r['archive_file'] for r in rows},{'OPD_2026-09-07.xlsx.enc','OPD_2026-09-14.xlsx.enc'})
        self.assertEqual({r['cell'] for r in rows},{'B32','B6'})

    def test_ui_blocks_reports_and_pies_and_only_offers_diagnostics(self):
        scan,repo,client,secrets,_=scan_cells(CONFLICT_CELLS)
        writes=repo.write_count
        result=run_app(ui_values(),secrets=secrets,repo=repo, evaluation_login=True)
        self.assertEqual(set(result['downloads']),{'OPD_Conflict_Review.csv'})
        self.assertNotIn('teaching_zip',result['state'])
        self.assertFalse(any(kind in ('metric','image') for kind,text in result['messages']))
        self.assertTrue(any(kind=='error' and 'Reports blocked' in text for kind,text in result['messages']))
        self.assertNotIn('Secret Learner',result['downloads']['OPD_Conflict_Review.csv'].decode('utf-8-sig'))
        self.assertEqual(repo.write_count,writes)

    def test_corrected_upload_then_refresh_succeeds(self):
        scan,repo,client,secrets,_=scan_cells(CONFLICT_CELLS)
        blocked=run_app(ui_values(),secrets=secrets,repo=repo, evaluation_login=True)
        fixed=dict(CONFLICT_CELLS);fixed[('WARD A','B6')]=None
        client.save(small_opd(date(2026,9,7),fixed))
        result=run_app(ui_values(pts_preview_charts=True),state=blocked['state'],secrets=secrets,repo=repo, evaluation_login=True)
        self.assertFalse(any(kind=='error' for kind,text in result['messages']))
        self.assertTrue(any(name.endswith('.zip') for name in result['downloads']))
        self.assertTrue(any(kind=='image' for kind,text in result['messages']))
        self.assertNotIn('OPD_Conflict_Review.csv',result['downloads'])

    def test_previous_good_zip_not_shown_after_conflict_loaded(self):
        clear=dict(CONFLICT_CELLS);clear[('WARD A','B6')]=None
        scan,repo,client,secrets,_=scan_cells(clear)
        good=run_app(ui_values(),secrets=secrets,repo=repo, evaluation_login=True)
        self.assertIn('teaching_zip',good['state'])
        client.save(small_opd(date(2026,9,7),CONFLICT_CELLS))
        bad=run_app(ui_values(),state=good['state'],secrets=secrets,repo=repo, evaluation_login=True)
        self.assertFalse(any(n.endswith(('.zip','.docx')) for n in bad['downloads']))
        self.assertNotIn('teaching_zip',bad['state'])

    def test_stale_scan_requires_detailed_rescan(self):
        scan,*_=scan_cells({('NYES','B6'):'Example, Avery ~ First'})
        scan.pop('strict_conflict_source_version')
        with self.assertRaisesRegex(OPDArchiveError,'refresh'):teaching_build_zip(scan,[2026])

    def test_unknown_denominator_is_blocked_not_na(self):
        for value in (None,float('nan'),float('inf'),101,-1):
            with self.assertRaises(OPDArchiveError):reach_percent(value)
        with self.assertRaises(OPDArchiveError):reach_totals([{'availability_review_shifts':1}])


class PieChartTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cells=make_reach_fixture()
        cells.update({('WARD A','B6'):'Ward, Example ~ Learner',('WARD A','B8'):'Ward, Example ~ ',
                      ('ADOLMED','C6'):'Unused, Example ~ '})
        cls.scan,cls.repo,cls.client,cls.secrets,_=scan_cells(cells)
        cls.blob,cls.rows=teaching_build_zip(cls.scan,[2026])

    def test_one_pie_per_area_not_person_and_zero_only_area_omitted(self):
        with ZipFile(BytesIO(self.blob)) as z:
            files=[n for n in z.namelist() if n.startswith('Learner_Reach_Charts/')]
            self.assertEqual(len(files),2)
            self.assertTrue(any('Academic_Pediatrics' in n for n in files))
            self.assertTrue(any('Ward_A' in n for n in files))
            self.assertFalse(any('ADOLMED' in n for n in files))
            self.assertTrue(all(z.read(n).startswith(b'\x89PNG\r\n\x1a\n') for n in files))

    def test_chart_data_matches_chair_area_totals(self):
        groups=teaching_chair_summary_data(self.scan,[2026])[0]['work_types']
        with ZipFile(BytesIO(self.blob)) as z:
            data=list(csv.DictReader(StringIO(z.read('clinical_experience_learner_reach.csv').decode('utf-8-sig'))))
        for row in data:
            group=next(g for g in groups if g['work_type']==row['work_type'])
            aliases = {'learner_reach_pct':'learner_reach_pct',
                       'total_scheduled_availability_hours':'recorded_clinical_hours',
                       'educational_hours':'hours_with_students', 'hours_without_students':'hours_without_students'}
            for public_key, internal_key in aliases.items():
                self.assertEqual(float(row[public_key]),group[internal_key])
        academic=next(r for r in data if r['work_type']=='Academic Pediatrics')
        self.assertEqual(float(academic['learner_reach_pct']),80.0)
        self.assertEqual(int(academic['educational_hours']),32) # NOT student weighted 36

    def test_docx_contains_the_same_chart_images(self):
        with ZipFile(BytesIO(self.blob)) as z:
            images={z.read(n) for n in z.namelist() if n.startswith('Learner_Reach_Charts/')}
            with ZipFile(BytesIO(z.read(TEACHING_CHAIR_SUMMARY_FILENAME))) as doc:
                self.assertEqual({doc.read(n) for n in doc.namelist() if n.startswith('word/media/')},images)
                xml=doc.read('word/document.xml').decode()
                self.assertIn('descr="Academic Pediatrics.',xml)
            for name in z.namelist():
                if name.startswith('Preceptor_Reports/'):
                    with ZipFile(BytesIO(z.read(name))) as doc:
                        self.assertFalse(any(n.startswith('word/media/') for n in doc.namelist()))

    def test_all_reported_percentages_are_numeric_not_na(self):
        with ZipFile(BytesIO(self.blob)) as z:
            for name in ('preceptor_teaching_summary.csv','preceptor_teaching_by_work_type.csv',
                         'preceptor_learner_reach_monthly.csv','clinical_experience_learner_reach.csv'):
                rows=list(csv.DictReader(StringIO(z.read(name).decode('utf-8-sig'))))
                self.assertTrue(rows)
                self.assertTrue(all(0<=float(r['learner_reach_pct'])<=100 for r in rows))
            for name in z.namelist():
                if name.endswith('.docx'):
                    self.assertNotIn('N/A',all_doc_text(z.read(name)))

    def test_zero_one_hundred_and_small_slices_render(self):
        for total,with_students in ((10,0),(10,10),(10000,1),(10000,9999)):
            counts={'recorded_clinical_shifts':total,'shifts_with_students':with_students,
                    'shifts_without_students':total-with_students}
            metrics=reach_totals([counts])
            image=learner_reach_pie('Clinical experience',metrics)
            self.assertTrue(image.startswith(b'\x89PNG'))

    def test_zero_denominator_and_mismatched_hours_block_chart(self):
        with self.assertRaises(OPDArchiveError):learner_reach_pie('Empty',reach_totals([]))
        metrics=reach_totals([{'recorded_clinical_shifts':10,'shifts_with_students':8,'shifts_without_students':2}])
        metrics['hours_with_students']=36
        with self.assertRaises(OPDArchiveError):learner_reach_pie('Wrong hours',metrics)

    def test_long_area_names_and_custom_date_labels_render(self):
        metrics=reach_totals([{'recorded_clinical_shifts':10,'shifts_with_students':8,'shifts_without_students':2}])
        png=learner_reach_pie('A long clinical experience name for the readability test',metrics,'26-27','February 17, 2026 - March 16, 2027')
        self.assertTrue(png.startswith(b'\x89PNG'))

    def test_package_has_no_credentials_or_learner_names(self):
        with ZipFile(BytesIO(self.blob)) as z:
            texts=[z.read(n) for n in z.namelist() if n.endswith(('.csv','.txt'))]
            for text in texts:
                self.assertNotIn(b'Confidential',text)
                self.assertNotIn(b'offline-test-token',text)

if __name__=='__main__':unittest.main()
