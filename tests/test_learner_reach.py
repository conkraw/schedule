"""Learner Reach tests: invented data, in-memory GitHub, no live requests."""
from datetime import date, timedelta
from io import BytesIO, StringIO
from zipfile import ZipFile
from unittest.mock import patch
import copy
import csv
import json
import unittest

from helpers import st, secret_settings, FakeGitHub, run_app
from test_reporting_dates import small_opd, all_doc_text
from schedule_app.services.opd_archive import GitHubOPDArchive, get_opd_archive_config, OPDArchiveError
from schedule_app.services.teaching_analysis import (teaching_scan_archives, teaching_filter_date_range,
    teaching_annual_rows, teaching_work_type_rows)
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.learner_reach import (require_learner_reach_data, learner_reach_rows,
    reach_totals, reach_percent, nonclinical_provider, LEARNER_REACH_COLUMNS)
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.chair_summary import teaching_chair_summary_data
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME, TEACHING_CSV_COLUMNS
from schedule_app.services.educational_time import TIME_CSV_COLUMNS


def make_reach_fixture():
    cells = {}
    for day, col in enumerate('BCDEF'):
        for period, row in [('AM', 6), ('PM', 8)]:
            student = f'Confidential Learner {day} {period}' if day < 4 else ' '
            cells['HOPE_DRIVE', f'{col}{row}'] = f'Example, Avery ~ {student}'
    # An additional learner and a duplicate of the first learner, same half-day.
    cells['HOPE_DRIVE', 'B7'] = 'Example, Avery ~ Confidential Second'
    cells['ETOWN', 'B6'] = 'example,  avery ~ Confidential Learner 0 AM'
    # A duplicate blank must NOT make an assigned shift unassigned.
    cells['NYES', 'B6'] = 'Example, Avery ~ '
    cells['NYES', 'G6'] = 'Available, Blake ~ '
    return cells


def scan_cells(cells, start=date(2026, 9, 7), order='Preceptor ~ Student'):
    secrets=secret_settings(); st.reset(secrets=secrets)
    repo=FakeGitHub(); client=GitHubOPDArchive(get_opd_archive_config(), transport=repo)
    raw=small_opd(start, cells)
    client.save(raw)
    scan=teaching_scan_archives(client, order)
    return scan, repo, client, secrets, raw


class LearnerReachTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scan,cls.repo,cls.client,cls.secrets,cls.raw=scan_cells(make_reach_fixture())
        cls.rows=teaching_annual_rows(cls.scan, [2026])
        cls.avery=next(r for r in cls.rows if r['preceptor_name']=='Example, Avery')
        cls.zip_bytes,_=teaching_build_zip(cls.scan, [2026])

    def test_eighty_percent_uses_unique_clinical_shifts(self):
        self.assertEqual(self.avery['recorded_clinical_shifts'],10)
        self.assertEqual(self.avery['shifts_with_students'],8)
        self.assertEqual(self.avery['learner_reach_pct'],80.0)

    def test_hours_use_four_hour_equivalents(self):
        self.assertEqual(self.avery['recorded_clinical_hours'],40)
        self.assertEqual(self.avery['hours_with_students'],32)
        self.assertEqual(self.avery['hours_without_students'],8)

    def test_multiple_students_do_not_multiply_educational_hours(self):
        self.assertEqual(self.avery['no_of_shifts'],9)
        self.assertEqual(self.avery['educational_hours'],32)

    def test_percentage_does_not_exceed_one_hundred(self):
        cells={('HOPE_DRIVE','B6'):'One, Provider ~ First; Second; Third'}
        scan,*_=scan_cells(cells)
        row=teaching_annual_rows(scan,[2026])[0]
        self.assertEqual(row['educational_hours'],4)
        self.assertEqual(row['learner_reach_pct'],100)
        self.assertEqual(row['hours_with_students'],4)

    def test_blank_student_includes_provider_with_zero_teaching(self):
        # The source stays complete; display no longer includes zero-only people.
        self.assertNotIn('Available, Blake', {r['preceptor_name'] for r in self.rows})
        row=next(r for r in learner_reach_rows(self.scan,[2026]) if r['preceptor_name']=='Available, Blake')
        self.assertEqual((row['recorded_clinical_hours'],row['learner_reach_pct']),(4,0))
        self.assertEqual(row['months_scheduled'],'September 2026')

    def test_same_assignment_multiple_academic_sites_one_shift(self):
        typed=[r for r in teaching_work_type_rows(self.scan,[2026]) if r['preceptor_name']=='Example, Avery']
        self.assertEqual(len(typed),1)
        self.assertEqual(typed[0]['work_type'],'Academic Pediatrics')
        self.assertEqual(typed[0]['recorded_clinical_shifts'],10)
        self.assertEqual(set(typed[0]['source_sites'].split('; ')),{'HOPE_DRIVE','ETOWN','NYES'})

    def test_repeated_blank_does_not_erase_student(self):
        view=teaching_filter_date_range(self.scan,ReportingPeriod('Monday',date(2026,9,7),date(2026,9,7)))
        row=teaching_annual_rows(view,[2026])[0]
        self.assertEqual((row['recorded_clinical_shifts'],row['shifts_with_students']),(2,2))

    def test_report_overall_is_ratio_not_mean(self):
        cells=make_reach_fixture()
        cells.update({('WARD A','B6'):'Other, Dana ~ Learner Dana',('WARD A','B8'):'Other, Dana ~ '})
        scan,*_=scan_cells(cells)
        summary=teaching_chair_summary_data(scan,[2026])[0]
        # Included contributors: 8/10 and 1/2. Exclude zero-only Blake.
        self.assertEqual(summary['learner_reach_pct'],75.0) # 9/12, NOT mean(80,50)
        self.assertEqual(summary['recorded_clinical_hours'],48)
        self.assertEqual(summary['hours_with_students'],36)

    def test_clinical_fields_match_no_student_in_memory(self):
        encoded=json.dumps(self.scan)
        self.assertNotIn('Confidential',encoded)
        require_learner_reach_data(self.scan)

    def test_zero_denominator_is_not_zero_percent(self):
        self.assertIsNone(reach_totals([])['learner_reach_pct'])
        with self.assertRaises(OPDArchiveError):
            reach_percent(None)

    def test_no_tilde_is_not_guessed(self):
        scan,*_=scan_cells({('NYES','B6'):'Maybe, Provider',('NYES','C6'):'Real, Provider ~ '})
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual([r['preceptor_name'] for r in learner_reach_rows(scan,[2026])],['Real, Provider'])
        self.assertTrue(any("No '~'" in row['issue'] for row in scan['warnings']))
        self.assertNotIn('Maybe, Provider',json.dumps(scan))

    def test_nonclinical_labels_not_capacity(self):
        for label in ('CLOSED','OFF','VACATION','PTO','NO CLINIC','ADMIN','CLINIC CANCELLED'):
            with self.subTest(label=label):
                scan,*_=scan_cells({('NYES','B6'):label+' ~ ',('NYES','C6'):'Real, Provider ~ '})
                self.assertEqual(teaching_annual_rows(scan,[2026]),[])
                self.assertEqual(len(learner_reach_rows(scan,[2026])),1)
        self.assertFalse(nonclinical_provider('Offutt, Avery'))

    def test_dash_student_is_blank(self):
        scan,*_=scan_cells({('NYES','B6'):'Real, Provider ~ --'})
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual(learner_reach_rows(scan,[2026])[0]['learner_reach_pct'],0)

    def test_reversed_order_handles_empty_student(self):
        scan,*_=scan_cells({('NYES','B6'):'Learner One ~ Real, Provider',('NYES','C6'):' ~ Real, Provider'}, order='Student ~ Preceptor')
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['learner_reach_pct'],50)

    def test_missing_provider_logged_not_inferred(self):
        scan,*_=scan_cells({('NYES','B6'):' ~ Private Student',('NYES','C6'):'Real, Provider ~ '})
        self.assertEqual(scan['sources'][0]['missing_provider_cells'],1)
        self.assertNotIn('Private Student',json.dumps(scan))
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual(len(learner_reach_rows(scan,[2026])),1)

    def test_availability_only_archive_generates_no_empty_reports(self):
        scan,*_=scan_cells({('NYES','B6'):'Real, Provider ~ ',('NYES','C6'):'Real, Provider ~ '})
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual(learner_reach_rows(scan,[2026])[0]['recorded_clinical_hours'],8)
        with self.assertRaisesRegex(OPDArchiveError,'No student assignments'):
            teaching_build_zip(scan,[2026])

    def test_exact_custom_dates_both_numerator_and_denominator(self):
        view=teaching_filter_date_range(self.scan,ReportingPeriod('Friday',date(2026,9,11),date(2026,9,11)))
        self.assertEqual(teaching_annual_rows(view,[2026]),[])
        row=learner_reach_rows(view,[2026])[0]
        self.assertEqual((row['recorded_clinical_hours'],row['learner_reach_pct']),(8,0))
        self.assertEqual(len(view['clinical_daily']),1)
        self.assertEqual(row['academic_year'],'Friday')

    def test_custom_dates_keep_blanks_across_july(self):
        scan,*_=scan_cells({('NYES','C6'):'Real, Provider ~ X',('NYES','D6'):'Real, Provider ~ '},start=date(2026,6,29))
        period=ReportingPeriod('Custom',date(2026,6,30),date(2026,7,1))
        view=teaching_filter_date_range(scan,period)
        row=teaching_annual_rows(view,[2026])[0]
        self.assertEqual(row['learner_reach_pct'],50)
        self.assertEqual(row['months_scheduled'],'June 2026; July 2026')

    def test_july_mode_separates_years(self):
        scan,*_=scan_cells({('NYES','C6'):'Real, Provider ~ X',('NYES','D6'):'Real, Provider ~ '},start=date(2026,6,29))
        self.assertEqual(teaching_annual_rows(scan,[2025])[0]['learner_reach_pct'],100)
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual(learner_reach_rows(scan,[2026])[0]['learner_reach_pct'],0)

    def test_conflicting_types_block_instead_of_omitting_percentage(self):
        from schedule_app.services.teaching_validation import TeachingConflictError
        scan,*_=scan_cells({('NYES','B6'):'Real, Provider ~ X',('WARD A','B6'):'Real, Provider ~ Y'})
        for operation in (lambda:teaching_annual_rows(scan,[2026]), lambda:teaching_work_type_rows(scan,[2026]),
                          lambda:teaching_build_zip(scan,[2026])):
            with self.assertRaises(TeachingConflictError):
                operation()
        self.assertEqual(len(scan['clinical_shift_conflicts']),1)


    def test_conflict_sources_keep_the_actual_category_labels(self):
        from schedule_app.services.teaching_validation import teaching_conflict_rows
        # The nursery/clinic pair is now an explicit exception; unrelated work types still block.
        scan,*_=scan_cells({('NYES','B6'):'Real, Provider ~ X',('COMPLEX','B6'):'Real, Provider ~ '})
        rows=teaching_conflict_rows(scan,[2026])
        self.assertEqual({(r['worksheet'],r['listed_work_type']) for r in rows},
                         {('NYES','Academic Pediatrics'),('COMPLEX','Complex Care')})


    def test_blank_conflicting_type_is_also_blocked(self):
        from schedule_app.services.teaching_validation import TeachingConflictError
        scan,*_=scan_cells({('NYES','B6'):'Real, Provider ~ X',('COMPLEX','B6'):'Real, Provider ~ '})
        self.assertEqual(len(scan['clinical_shift_conflicts']),1)
        with self.assertRaises(TeachingConflictError):
            teaching_annual_rows(scan,[2026])


    def test_stale_schema_requires_refresh_before_export(self):
        old=dict(self.scan);old.pop('learner_reach_version')
        with self.assertRaisesRegex(OPDArchiveError,'refresh'):teaching_build_zip(old,[2026])

    def test_invalid_clinical_totals_rejected(self):
        scan=copy.deepcopy(self.scan)
        scan['clinical_daily'][0]['recorded_clinical_shifts']+=1
        with self.assertRaises(OPDArchiveError):require_learner_reach_data(scan)

    def test_false_learner_counts_rejected(self):
        scan=copy.deepcopy(self.scan)
        for key in ('clinical_daily','clinical_daily_by_work_type'):
            for row in scan[key]:
                if row['preceptor_name']=='Available, Blake':
                    row['shifts_with_students']=1;row['shifts_without_students']=0
        with self.assertRaises(OPDArchiveError):require_learner_reach_data(scan)

    def test_filters_do_not_mutate_full_clinical_source(self):
        before=copy.deepcopy(self.scan)
        view=teaching_filter_date_range(self.scan,ReportingPeriod('Monday',date(2026,9,7),date(2026,9,7)))
        view['clinical_daily'][0]['source_sites'].append('TEST')
        self.assertEqual(self.scan,before)

    def test_current_rotation_replacement_no_history_counted(self):
        scan,repo,client,_,raw=scan_cells({('NYES','B6'):'Real, Provider ~ X'})
        client.save(small_opd(date(2026,9,7),{('NYES','B6'):'Real, Provider ~ '}))
        updated=teaching_scan_archives(client)
        self.assertEqual(teaching_annual_rows(updated,[2026]),[])
        row=learner_reach_rows(updated,[2026])[0]
        self.assertEqual((row['recorded_clinical_shifts'],row['learner_reach_pct']),(1,0))

    def test_scan_never_writes_to_github(self):
        writes=self.repo.write_count
        teaching_scan_archives(self.client)
        self.assertEqual(writes,self.repo.write_count)

    def test_csvs_use_simple_hour_headers_and_keep_reach(self):
        with ZipFile(BytesIO(self.zip_bytes)) as z:
            reader=csv.DictReader(StringIO(z.read('preceptor_teaching_summary.csv').decode('utf-8-sig')))
            self.assertEqual(tuple(reader.fieldnames),TIME_CSV_COLUMNS)
            rows=list(reader);avery=next(r for r in rows if r['preceptor_name']=='Example, Avery')
            self.assertEqual(avery['learner_reach_pct'],'80.0')
            self.assertEqual(avery['educational_hours'],'32')
            self.assertIn('preceptor_learner_reach_monthly.csv',z.namelist())

    def test_monthly_csv_excludes_entirely_unassigned_preceptors(self):
        with ZipFile(BytesIO(self.zip_bytes)) as z:
            reader=csv.DictReader(StringIO(z.read('preceptor_learner_reach_monthly.csv').decode('utf-8-sig')))
            rows=list(reader)
            self.assertNotIn('Available, Blake',{r['preceptor_name'] for r in rows})
            self.assertEqual(rows[0]['month'],'2026-09-01')
            self.assertEqual(rows[0]['learner_reach_pct'],'80.0')

    def test_word_reports_use_learner_reach_and_omit_students(self):
        with ZipFile(BytesIO(self.zip_bytes)) as z:
            for filename in [f for f in z.namelist() if f.endswith('.docx')]:
                text=all_doc_text(z.read(filename))
                self.assertIn('Learner Reach',text)
                self.assertNotIn('Confidential',text)
                self.assertIn('not verified total clinical work',text)

    def test_generic_labels_remain_separate(self):
        scan,*_=scan_cells({('SJR_HOSP','B6'):'SJR_1 ~ Learner'})
        data=teaching_chair_summary_data(scan,[2026])[0]
        self.assertEqual(data['named_preceptor_count'],0)
        self.assertEqual(data['unresolved_labels'][0]['preceptor_name'],'SJR_1')

    def test_full_interface_explains_no_student_assignments(self):
        scan,repo,client,secrets,raw=scan_cells({('NYES','B6'):'Real, Provider ~ '})
        result=run_app({'schedule_app_mode':'PTS',
            'teaching_load_archives':True,'teaching_build_zip':True,
            'teaching_period_start':date(2026,9,7),'teaching_period_end':date(2026,9,13),
            'teaching_period_label':'26-27'},secrets=secrets,repo=repo, evaluation_login=True)
        self.assertFalse(any(name.endswith('.zip') for name in result['downloads']))
        self.assertTrue(any(level=='info' and 'No student assignments' in text for level,text in result['messages']))
        self.assertFalse(any(level=='error' for level,text in result['messages']))

    def test_cross_rotation_duplicate_provider_slot_once(self):
        _,repo,client,_,_=scan_cells({('NYES','B6'):'Real, Provider ~ '})
        # Same snapshot file cannot appear twice through list_rotations, but the
        # distinct registry still handles any overlap of source calendar weeks.
        from schedule_app.services.learner_reach import ClinicalShiftAccumulator
        c=ClinicalShiftAccumulator();row={'day':date(2026,9,7),'shift':'AM','has_student':False}
        c.add('real',row,'Academic Pediatrics','NYES');c.add('real',row,'Academic Pediatrics','ETOWN')
        self.assertEqual(c.finish({'real':'Real, Provider'})['clinical_daily'][0]['recorded_clinical_shifts'],1)

if __name__=='__main__':unittest.main()
