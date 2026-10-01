"""Outpatient-over-PSHCH-Nursery regression tests; offline invented fixtures only."""
import copy
import csv
from datetime import date
from io import BytesIO, StringIO
import json
from unittest.mock import patch
import unittest
from zipfile import ZipFile

from helpers import st, secret_settings, FakeGitHub, run_app, make_opd
from test_reporting_dates import small_opd, all_doc_text
from test_learner_reach import scan_cells
from test_strict_reach_charts import ui_values
from schedule_app.services.opd_archive import GitHubOPDArchive, get_opd_archive_config, OPDArchiveError
from schedule_app.services.teaching_analysis import (teaching_scan_archives, teaching_filter_date_range,
    teaching_annual_rows, teaching_work_type_rows, teaching_require_date_range_data)
from schedule_app.services.learner_reach import learner_reach_rows, require_learner_reach_data
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_priority import (outpatient_priority_audit_rows,
    selected_priority_adjustments, outpatient_priority_report_note)
from schedule_app.services.teaching_validation import TeachingConflictError, teaching_conflict_rows
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.chair_summary import teaching_chair_summary_data
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME

PERSON = 'Example, Avery'

def assignment(text='Clinic Learner'):
    return f'{PERSON} ~ {text}'

def overlap(clinic='NYES', student='Clinic Learner', nursery='Nursery Learner'):
    return {(clinic,'B8'):assignment(student),('PSHCH_NURSERY','B8'):assignment(nursery)}

def example_cells():
    return {('PSHCH_NURSERY','B6'):assignment('Private Nursery Learner'),
            ('PSHCH_NURSERY','B8'):assignment('Private Nursery Learner'),
            ('HOPE_DRIVE','B8'):assignment('Private Clinic One; Private Clinic Two'),
            ('PSHCH_NURSERY','C6'):assignment(' '),
            ('HOPE_DRIVE','C8'):assignment(' '),
            ('ETOWN','D8'):assignment('Private Clinic Three')}

class OutpatientPriorityTests(unittest.TestCase):
    def test_all_three_clinics_take_priority(self):
        for clinic in ('HOPE_DRIVE','ETOWN','NYES'):
            with self.subTest(clinic=clinic):
                scan,*_=scan_cells(overlap(clinic))
                self.assertEqual(teaching_conflict_rows(scan,[2026]),[])
                rows=teaching_work_type_rows(scan,[2026])
                self.assertEqual(len(rows),1)
                self.assertEqual((rows[0]['work_type'],rows[0]['no_of_shifts'],rows[0]['recorded_clinical_hours'],rows[0]['learner_reach_pct']),
                                 ('Academic Pediatrics',1,4,100))
                self.assertEqual(rows[0]['source_sites'],clinic)

    def test_both_am_and_pm_use_the_rule(self):
        for row in (6,8):
            cells={(s,f'B{row}'):v for (s,_),v in overlap().items()}
            scan,*_=scan_cells(cells)
            self.assertEqual(teaching_annual_rows(scan,[2026])[0]['no_of_shifts'],1)
            self.assertEqual(scan['outpatient_priority_adjustments'][0]['shift'],'AM' if row==6 else 'PM')

    def test_nursery_only_student_is_removed_not_transferred(self):
        scan,*_=scan_cells(overlap(student=' '))
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        row=learner_reach_rows(scan,[2026])[0]
        self.assertEqual((row['recorded_clinical_shifts'],row['shifts_with_students'],row['learner_reach_pct']),(1,0,0))
        self.assertEqual(scan['student_shifts_removed_by_outpatient_priority'],1)

    def test_blank_outpatient_overlap_stays_in_included_preceptors_denominator(self):
        cells=overlap(student=' ');cells['NYES','C8']=assignment()
        scan,*_=scan_cells(cells)
        row=teaching_annual_rows(scan,[2026])[0]
        self.assertEqual((row['no_of_shifts'],row['recorded_clinical_hours'],row['hours_with_students'],row['learner_reach_pct']),(1,8,4,50))

    def test_clinic_assigned_nursery_blank(self):
        scan,*_=scan_cells(overlap(nursery=' '))
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['learner_reach_pct'],100)
        self.assertEqual(scan['student_shifts_removed_by_outpatient_priority'],0)

    def test_both_blank_is_resolved_but_hidden(self):
        scan,*_=scan_cells(overlap(student=' ',nursery=' '))
        self.assertEqual(teaching_conflict_rows(scan,[2026]),[])
        self.assertEqual(teaching_annual_rows(scan,[2026]),[])
        self.assertEqual(learner_reach_rows(scan,[2026])[0]['recorded_clinical_hours'],4)

    def test_am_nursery_and_pm_clinic_are_preserved(self):
        cells=overlap();cells['PSHCH_NURSERY','B6']=assignment('Nursery Learner')
        scan,*_=scan_cells(cells)
        rows={r['work_type']:r for r in teaching_work_type_rows(scan,[2026])}
        self.assertEqual(set(rows),{'Academic Pediatrics','PSHCH Nursery'})
        self.assertEqual(rows['PSHCH Nursery']['recorded_clinical_hours'],4)
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['recorded_clinical_hours'],8)

    def test_different_date_or_preceptor_does_not_apply(self):
        for cells in (
            {('NYES','B8'):assignment(),('PSHCH_NURSERY','C8'):assignment('Nursery Learner')},
            {('NYES','B8'):assignment(),('PSHCH_NURSERY','B8'):'Other, Provider ~ Nursery Learner'}):
            scan,*_=scan_cells(cells)
            self.assertEqual(scan['outpatient_priority_adjustments'],[])
            self.assertEqual(sum(r['no_of_shifts'] for r in teaching_annual_rows(scan,[2026])),2)

    def test_nursery_without_clinic_is_unchanged(self):
        scan,*_=scan_cells({('PSHCH_NURSERY','B6'):assignment(),('PSHCH_NURSERY','B8'):assignment(' ')})
        row=teaching_work_type_rows(scan,[2026])[0]
        self.assertEqual((row['work_type'],row['recorded_clinical_hours'],row['learner_reach_pct']),('PSHCH Nursery',8,50))
        self.assertEqual(scan['outpatient_priority_adjustments'],[])

    def test_two_clinic_students_earn_one_shift_of_educational_hours(self):
        scan,*_=scan_cells(overlap(student='Clinic One; Clinic Two'))
        row=teaching_annual_rows(scan,[2026])[0]
        self.assertEqual((row['no_of_shifts'],row['educational_hours'],row['recorded_clinical_hours'],row['hours_with_students'],row['learner_reach_pct']),
                         (2,4,4,4,100))

    def test_same_student_on_both_sheets_keeps_clinic_credit_once(self):
        scan,*_=scan_cells(overlap(student='Same Learner',nursery='Same Learner'))
        row=teaching_work_type_rows(scan,[2026])[0]
        self.assertEqual((row['no_of_shifts'],row['source_sites']),(1,'NYES'))
        self.assertEqual(scan['student_shifts_removed_by_outpatient_priority'],0)

    def test_duplicate_nursery_rows_never_add_hours_or_credit(self):
        cells={('NYES','B6'):assignment(),('PSHCH_NURSERY','B6'):assignment('Nursery Learner'),
               ('PSHCH_NURSERY','B7'):assignment('Nursery Learner')}
        scan,*_=scan_cells(cells)
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['recorded_clinical_hours'],4)
        self.assertEqual(scan['sources'][0]['nursery_student_assignment_listings_excluded'],2)
        self.assertEqual(scan['nursery_clinical_listings_excluded'],2)

    def test_multiple_academic_sites_still_one_clinical_shift(self):
        cells=overlap();cells['HOPE_DRIVE','B8']=assignment();cells['ETOWN','B8']=assignment('Second Clinic Learner')
        scan,*_=scan_cells(cells)
        row=teaching_work_type_rows(scan,[2026])[0]
        self.assertEqual((row['no_of_shifts'],row['recorded_clinical_hours']),(2,4))
        self.assertEqual(len(scan['outpatient_priority_adjustments']),1)

    def test_other_outpatient_and_nursery_remain_conflicts(self):
        for site in ('COMPLEX','LANCASTER','ADOLMED'):
            scan,*_=scan_cells({(site,'B8'):assignment(),('PSHCH_NURSERY','B8'):assignment('Nursery Learner')})
            self.assertEqual(scan['outpatient_priority_adjustments'],[])
            with self.assertRaises(TeachingConflictError):teaching_build_zip(scan,[2026])

    def test_three_way_conflict_still_blocks_with_remaining_cells(self):
        cells=overlap();cells['WARD A','B8']=assignment('Ward Learner')
        scan,*_=scan_cells(cells)
        with self.assertRaises(TeachingConflictError) as caught:teaching_build_zip(scan,[2026])
        self.assertEqual({r['worksheet'] for r in caught.exception.rows},{'NYES','WARD A'})
        self.assertEqual(len(selected_priority_adjustments(scan,[2026])),1)

    def test_surname_normalization_and_explicit_aliases_apply(self):
        cells={('NYES','B8'):'example,  avery ~ Clinic Learner',('PSHCH_NURSERY','B8'):'Example, A. ~ Nursery Learner'}
        with patch('schedule_app.services.teaching_analysis.TEACHING_PRECEPTOR_NAME_MAP',{'Example, A.':'Example, Avery'}):
            scan,*_=scan_cells(cells)
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['no_of_shifts'],1)

    def test_reverse_tilde_order(self):
        cells={key:' ~ '.join(reversed([s.strip() for s in value.split('~')])) for key,value in example_cells().items()}
        scan,*_=scan_cells(cells,order='Student ~ Preceptor')
        self.assertEqual(teaching_annual_rows(scan,[2026])[0]['learner_reach_pct'],60)

    def test_source_sheet_order_does_not_change_counts(self):
        cells=example_cells();a,*_=scan_cells(cells);b,*_=scan_cells(dict(reversed(list(cells.items()))))
        self.assertEqual(teaching_annual_rows(a,[2026]),teaching_annual_rows(b,[2026]))
        self.assertEqual(teaching_work_type_rows(a,[2026]),teaching_work_type_rows(b,[2026]))

    def test_overlapping_rotations_in_either_read_order(self):
        secrets=secret_settings();st.reset(secrets=secrets)
        repo=FakeGitHub();client=GitHubOPDArchive(get_opd_archive_config(),transport=repo)
        # Nursery on Monday of week 2; clinic is in a separately archived OPD for that Monday.
        client.save(make_opd(date(2026,9,7),{('PSHCH_NURSERY','B40'):assignment('Nursery Learner')}))
        client.save(small_opd(date(2026,9,14),{('NYES','B8'):assignment()}))
        original=client.list_rotations
        a=teaching_scan_archives(client)
        with patch.object(client,'list_rotations',side_effect=lambda commit=None:list(reversed(original(commit=commit)))):
            b=teaching_scan_archives(client)
        for scan in (a,b):
            row=next(r for r in teaching_work_type_rows(scan,[2026]) if r['preceptor_name']==PERSON)
            self.assertEqual((row['work_type'],row['no_of_shifts']),('Academic Pediatrics',1))
            self.assertEqual({r['archive_file'] for r in outpatient_priority_audit_rows(scan,[2026])},
                             {'OPD_2026-09-07.xlsx.enc','OPD_2026-09-14.xlsx.enc'})

    def test_exact_date_filter_does_not_mutate_scan(self):
        scan,*_=scan_cells(example_cells());before=copy.deepcopy(scan)
        only_monday=teaching_filter_date_range(scan,ReportingPeriod('Custom',date(2026,9,7),date(2026,9,7)))
        later=teaching_filter_date_range(scan,ReportingPeriod('Later',date(2026,9,8),date(2026,9,9)))
        self.assertEqual(teaching_annual_rows(only_monday,[2026])[0]['learner_reach_pct'],100)
        self.assertEqual(len(selected_priority_adjustments(only_monday,[2026])),1)
        self.assertEqual(selected_priority_adjustments(later,[2026]),[])
        self.assertEqual(scan,before)

    def test_july_boundary_and_custom_label(self):
        cells=overlap();scan,*_=scan_cells(cells,start=date(2026,6,29))
        self.assertEqual(len(selected_priority_adjustments(scan,[2025])),1)
        self.assertEqual(selected_priority_adjustments(scan,[2026]),[])
        custom=teaching_filter_date_range(scan,ReportingPeriod('26-27',date(2026,2,1),date(2027,3,1)))
        self.assertEqual(teaching_annual_rows(custom,[2026])[0]['academic_year'],'26-27')
        self.assertEqual(len(selected_priority_adjustments(custom,[2026])),1)

    def test_audit_contains_no_students_and_both_source_actions(self):
        scan,*_=scan_cells(overlap())
        rows=outpatient_priority_audit_rows(scan,[2026])
        self.assertEqual(len(rows),2)
        self.assertEqual({r['cell'] for r in rows},{'B8'})
        self.assertEqual({r['worksheet'] for r in rows},{'NYES','PSHCH_NURSERY'})
        self.assertEqual({r['report_action'].split(':')[0] for r in rows},{'RETAINED','EXCLUDED'})
        self.assertNotIn('Clinic Learner',json.dumps(scan));self.assertNotIn('Nursery Learner',json.dumps(rows))

    def test_source_manifest_reconciles_after_suppression(self):
        cells=example_cells();cells['PSHCH_NURSERY','B7']=assignment('Private Nursery Learner')
        scan,*_=scan_cells(cells)
        for row in scan['sources']:
            self.assertEqual(row['assigned_student_shifts_read'],row['assigned_student_shifts_counted']+
                             row['duplicate_student_shifts_removed']+row['nursery_student_assignment_listings_excluded'])
        self.assertEqual(sum(row['assigned_student_shifts_counted'] for row in scan['sources']),sum(row['no_of_shifts'] for row in scan['monthly']))

    def test_all_percentages_and_totals_reconcile(self):
        scan,*_=scan_cells(example_cells())
        require_learner_reach_data(scan);teaching_require_date_range_data(scan)
        overall=teaching_annual_rows(scan,[2026])[0]
        types={row['work_type']:row for row in teaching_work_type_rows(scan,[2026])}
        self.assertEqual(overall['learner_reach_pct'],60)
        self.assertEqual(types['Academic Pediatrics']['learner_reach_pct'],66.7)
        self.assertEqual(types['PSHCH Nursery']['learner_reach_pct'],50)
        for column in ('recorded_clinical_hours','hours_with_students','no_of_shifts','educational_hours'):
            self.assertEqual(overall[column],sum(row[column] for row in types.values()))

    def test_scanning_and_reporting_never_change_original_archives(self):
        scan,repo,client,_,raw=scan_cells(example_cells());before=dict(repo.tree);writes=repo.write_count
        teaching_build_zip(scan,[2026])
        self.assertEqual(repo.tree,before);self.assertEqual(repo.write_count,writes)
        self.assertEqual(client.load(date(2026,9,7))['raw'],raw)

    def test_zip_charts_csvs_and_word_reports_match(self):
        scan,*_=scan_cells(example_cells());blob,_=teaching_build_zip(scan,[2026])
        with ZipFile(BytesIO(blob)) as z:
            self.assertIsNone(z.testzip())
            self.assertIn('Outpatient_Priority_Adjustments.csv',z.namelist())
            csvrows=list(csv.DictReader(StringIO(z.read('clinical_experience_learner_reach.csv').decode('utf-8-sig'))))
            self.assertEqual({r['work_type']:float(r['learner_reach_pct']) for r in csvrows},
                             {'Academic Pediatrics':66.7,'PSHCH Nursery':50})
            self.assertEqual(len([n for n in z.namelist() if n.startswith('Learner_Reach_Charts/')]),2)
            for path in [TEACHING_CHAIR_SUMMARY_FILENAME]+[n for n in z.namelist() if n.startswith('Preceptor_Reports/')]:
                text=all_doc_text(z.read(path))
                self.assertIn('Outpatient priority',text)
                self.assertIn('66.7%',text);self.assertIn('50.0%',text)
                self.assertNotIn('N/A',text);self.assertNotIn('Private Nursery Learner',text)

    def test_nursery_category_hidden_when_all_teaching_was_overlapped(self):
        scan,*_=scan_cells(overlap())
        self.assertEqual([x['work_type'] for x in teaching_chair_summary_data(scan,[2026])[0]['work_types']],['Academic Pediatrics'])
        blob,_=teaching_build_zip(scan,[2026])
        with ZipFile(BytesIO(blob)) as z:
            self.assertEqual(len([n for n in z.namelist() if n.startswith('Learner_Reach_Charts/')]),1)

    def test_ui_resolves_and_generates_reports(self):
        scan,repo,client,secrets,_=scan_cells(example_cells())
        result=run_app(ui_values(pts_show_diagnostics=True,pts_preview_charts=True),secrets=secrets,repo=repo, evaluation_login=True)
        self.assertFalse(any(kind=='error' for kind,text in result['messages']))
        self.assertTrue(any(n.endswith('.zip') for n in result['downloads']))
        self.assertIn('Outpatient_Priority_Adjustments.csv',result['downloads'])
        self.assertNotIn('OPD_Conflict_Review.csv',result['downloads'])
        self.assertTrue(any(kind=='image' for kind,text in result['messages']))

    def test_ui_three_way_conflict_keeps_diagnostics_not_reports(self):
        cells=overlap();cells['WARD A','B8']=assignment('Ward Learner')
        scan,repo,client,secrets,_=scan_cells(cells)
        result=run_app(ui_values(pts_show_diagnostics=True,pts_preview_charts=True),secrets=secrets,repo=repo, evaluation_login=True)
        self.assertIn('OPD_Conflict_Review.csv',result['downloads'])
        self.assertNotIn('Outpatient_Priority_Adjustments.csv',result['downloads'])
        self.assertFalse(any(n.endswith(('.zip','.docx')) for n in result['downloads']))
        self.assertFalse(any(kind in ('metric','image') for kind,text in result['messages']))

    def test_old_scan_requires_rescan(self):
        scan,*_=scan_cells(example_cells());scan.pop('outpatient_priority_version')
        with self.assertRaisesRegex(OPDArchiveError,'refresh'):teaching_build_zip(scan,[2026])

    def test_ui_stale_signature_clears_scan_and_downloads(self):
        scan,repo,client,secrets,_=scan_cells(example_cells())
        state={'teaching_options_signature':'old-rule','teaching_scan':scan,'teaching_zip':b'OLD'}
        result=run_app(ui_values(teaching_load_archives=False),state=state,secrets=secrets,repo=repo, evaluation_login=True)
        self.assertNotIn('teaching_scan',result['state']);self.assertNotIn('teaching_zip',result['state'])
        self.assertFalse(any(n.endswith(('.zip','.docx')) for n in result['downloads']))

    def test_person_report_note_is_scoped_to_person(self):
        scan,*_=scan_cells(example_cells())
        self.assertTrue(outpatient_priority_report_note(scan,[2026],preceptor_name=PERSON))
        self.assertEqual(outpatient_priority_report_note(scan,[2026],preceptor_name='Other, Person'),'')

if __name__ == '__main__': unittest.main()
