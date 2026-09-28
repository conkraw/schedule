"""Run from the project root: python -m unittest discover -s tests -v

Uses invented data and simulated Streamlit/GitHub. Does not touch a live archive.
"""
import ast
from datetime import date, timedelta
from importlib import import_module
import io
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch
from zipfile import ZipFile
from openpyxl import load_workbook

from helpers import ROOT, FakeGitHub, st, secret_settings, make_opd, make_roster, make_qgenda, Upload, run_app
from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveConfig, OPDArchiveError, inspect_opd_rotation
from schedule_app.services.teaching_analysis import (teaching_scan_archives, teaching_academic_label,
    teaching_academic_start, teaching_work_type, teaching_work_type_rows)
from schedule_app.services.student_schedules import collect_opd_assignments, create_ms_schedule_template, populate_ms_schedule
from schedule_app.services.primary_preceptors import build_preceptor_assignment_report, build_preceptor_report_workbook
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.settings import REPORT_COLUMNS, TEACHING_CHAIR_SUMMARY_FILENAME

MODES = {
    'Instructions':'instructions', 'Format OPD + Summary':'format_opd_summary',
    'Create Student Schedule':'create_student_schedule', 'OPD Check':'opd_check',
    'Create Individual Schedules':'create_individual_schedules', 'OPD Archive':'opd_archive',
    'Preceptor Teaching Summary':'preceptor_teaching_summary',
    'OPD MD PA Conflict Detector':'opd_md_pa_conflict_detector',
    'Shift Availability Tracker':'shift_availability_tracker',
}

class ModularTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raw = make_opd()
        cls.config_values=secret_settings()
        cls.repo=FakeGitHub()
        cls.config=OPDArchiveConfig(**cls.config_values['opd_archive'])
        cls.client=GitHubOPDArchive(cls.config,transport=cls.repo)
        cls.client.save(cls.raw)
        cls.scan=teaching_scan_archives(cls.client)
        names=['Learner '+word for word in ('One','Two','Three','Four','Five')]
        cls.assignments,_=collect_opd_assignments(cls.raw,names)
        dates=[date(2026,8,3)+timedelta(days=i) for i in range(28)]
        blank=create_ms_schedule_template(names,dates)
        cls.master=populate_ms_schedule(blank.getvalue(),cls.assignments)

    def test_all_python_files_compile(self):
        for p in ROOT.rglob('*.py'):
            compile(p.read_text(), str(p), 'exec')

    def test_entrypoint_small(self):
        self.assertLess(len((ROOT/'app_sch_2026.py').read_text().splitlines()),50)

    def test_sections_import_without_rendering(self):
        st.reset()
        for short in MODES.values():
            module=import_module('schedule_app.sections.'+short)
            self.assertTrue(callable(module.render))
        self.assertEqual(st.events,[])

    def test_all_nine_modes_start_without_files(self):
        for mode in MODES:
            with self.subTest(mode=mode):
                result=run_app({'schedule_app_mode':mode})
                self.assertEqual(result['events'][0],'set_page_config')

    def test_instruction_download(self):
        result=run_app({'schedule_app_mode':'Instructions','Start date (m/d/yyyy)':'8/3/2026'})
        self.assertIn('Qgenda_Report_Instructions.docx',result['downloads'])

    def test_opd_check_download(self):
        result=run_app({'schedule_app_mode':'OPD Check',
                       'baseline':Upload(self.raw,'base.xlsx'),'assigned':Upload(self.raw,'assigned.xlsx')})
        self.assertIn('change_report.docx',result['downloads'])

    def test_existing_archive_round_trip(self):
        self.assertEqual(self.client.load(date(2026,8,3))['raw'],self.raw)
        self.assertTrue(next(iter(self.repo.tree.values())).startswith(b'gAAAA'))

    def test_identical_save_does_not_commit(self):
        n=self.repo.write_count
        receipt=self.client.save(self.raw)
        self.assertEqual(receipt['action'],'unchanged')
        self.assertEqual(self.repo.write_count,n)

    def test_replacement_and_other_rotation(self):
        repo=FakeGitHub(); client=GitHubOPDArchive(self.config,transport=repo)
        client.save(self.raw)
        revised=make_opd(changes={('HOPE_DRIVE','B8'):'New, Name ~ Learner One'})
        self.assertEqual(client.save(revised)['action'],'replaced')
        client.save(make_opd(start=date(2026,8,31)))
        self.assertEqual(len(client.list_rotations()),2)
        self.assertEqual(client.load(date(2026,8,3))['raw'],revised)

    def test_wrong_key_stops_load(self):
        other=OPDArchiveConfig(**secret_settings()['opd_archive'])
        with self.assertRaises(OPDArchiveError):
            GitHubOPDArchive(other,transport=self.repo).load(date(2026,8,3))

    def test_archive_ui_reload(self):
        result=run_app({'schedule_app_mode':'OPD Archive','archive_page_load':True},
                       secrets=self.config_values,repo=self.repo)
        self.assertEqual(result['downloads']['OPD_2026-08-03.xlsx'],self.raw)

    def test_create_student_schedule_from_upload(self):
        result=run_app({'schedule_app_mode':'Create Student Schedule',
                       'opd_main':Upload(self.raw,'Updated_OPD.xlsx'),
                       'rot_main':Upload(make_roster(),'rotation.csv'),'opd_build_master':True},
                       secrets=self.config_values,repo=self.repo)
        self.assertIn('MS_Schedule.xlsx',result['downloads'])

    def test_create_student_schedule_from_archive(self):
        result=run_app({'schedule_app_mode':'Create Student Schedule',
                       'opd_source_choice':'Reload archived OPD','schedule_archive_load':True,
                       'rot_main':Upload(make_roster(),'rotation.csv'),'opd_build_master':True},
                       secrets=self.config_values,repo=self.repo)
        self.assertIn('MS_Schedule.xlsx',result['downloads'])

    def test_bad_roster_date_blocks_schedule(self):
        result=run_app({'schedule_app_mode':'Create Student Schedule',
                       'opd_main':Upload(self.raw,'OPD.xlsx'),
                       'rot_main':Upload(make_roster(start=date(2026,9,7)),'rotation.csv'),'opd_build_master':True},
                       secrets=self.config_values,repo=self.repo)
        self.assertNotIn('MS_Schedule.xlsx',result['downloads'])
        self.assertTrue(any('does not match' in text for _,text in result['messages']))

    def test_individual_schedule_zip_and_report(self):
        result=run_app({'schedule_app_mode':'Create Individual Schedules',
                       'individual_schedule_master':Upload(self.master,'MS_Schedule.xlsx'),
                       'Build individual schedules + preceptor report':True})
        self.assertIn('Preceptor_Assignment_Report.xlsx',result['downloads'])
        with ZipFile(io.BytesIO(result['downloads']['individual_schedules_with_preceptor_report.zip'])) as z:
            self.assertEqual(len(z.namelist()),6)
            self.assertIn('Preceptor_Assignment_Report.xlsx',z.namelist())

    def test_three_session_primary_and_fallback_flags(self):
        report=build_preceptor_assignment_report(load_workbook(io.BytesIO(self.master)))
        self.assertEqual(report.columns.tolist(),REPORT_COLUMNS)
        primary=report[report['primary_preceptor']=='YES']
        self.assertTrue((primary.groupby(['student_name','monday_date']).size()==1).all())
        self.assertTrue((primary.loc[primary['no_of_sessions']<3,'primary_preceptor_flag']=='YES').all())
        self.assertTrue((report.loc[report['no_of_sessions']==3,'fragmented_preceptor']=='NO').all())

    def test_excel_report_table(self):
        report=build_preceptor_assignment_report(load_workbook(io.BytesIO(self.master)))
        output,_=build_preceptor_report_workbook(report)
        wb=load_workbook(output)
        self.assertIn('PreceptorAssignmentTable',wb['Preceptor Assignments'].tables)
        self.assertEqual([c.value for c in wb['Preceptor Assignments'][1]],REPORT_COLUMNS)

    def test_stale_preceptor_preview_no_keyerror(self):
        import pandas as pd,hashlib
        source=('MS_Schedule.xlsx',hashlib.sha256(self.master).hexdigest())
        result=run_app({'schedule_app_mode':'Create Individual Schedules',
                       'individual_schedule_master':Upload(self.master,'MS_Schedule.xlsx')},
                       state={'individual_report_schema_version':3,'individual_schedule_source':source,
                              'individual_preceptor_preview':pd.DataFrame({'primary_preceptor':['YES']})})
        self.assertNotIn('individual_preceptor_preview',result['state'])

    def test_work_types_preserved(self):
        self.assertEqual([teaching_work_type(s) for s in ['HOPE_DRIVE','ETOWN','NYES']],['Academic Pediatrics']*3)
        self.assertEqual(teaching_work_type('WARD A'),'Ward A')
        self.assertEqual(teaching_work_type('COMPLEX'),'Complex Care')

    def test_academic_year_boundary(self):
        self.assertEqual(teaching_academic_label(teaching_academic_start(date(2026,7,1))),'26-27')
        self.assertEqual(teaching_academic_label(teaching_academic_start(date(2026,6,30))),'25-26')

    def test_double_student_count_and_duplicates(self):
        # June fixture includes 2 different students for the same teacher in B8/B9.
        rows=[r for r in self.scan['monthly_by_work_type'] if r['preceptor_name']=='Adams, Alex' and r['work_type']=='Academic Pediatrics']
        self.assertEqual(sum(r['no_of_shifts'] for r in rows),5)
        self.assertEqual(self.scan['duplicate_assignments_removed'],1)

    def test_teaching_zip_has_all_reports(self):
        data,annual=teaching_build_zip(self.scan,[2026])
        with ZipFile(io.BytesIO(data)) as z:
            self.assertIsNone(z.testzip())
            self.assertIn(TEACHING_CHAIR_SUMMARY_FILENAME,z.namelist())
            self.assertIn('preceptor_teaching_summary.csv',z.namelist())
            self.assertIn('preceptor_teaching_by_work_type.csv',z.namelist())
            for name in z.namelist():
                contents=z.read(name)
                if name.endswith('.docx'):
                    with ZipFile(io.BytesIO(contents)) as doc:
                        text=doc.read('word/document.xml')
                        self.assertNotIn(b'Learner One',text)
                else:
                    self.assertNotIn(b'Learner One',contents)

    def test_teaching_page_builds_chair_and_zip(self):
        result=run_app({'schedule_app_mode':'Preceptor Teaching Summary',
                       'teaching_load_archives':True,'teaching_build_zip':True,'teaching_selected_years':[2026]},
                       secrets=self.config_values,repo=self.repo)
        self.assertIn('Preceptor_Teaching_26-27.zip',result['downloads'])
        self.assertIn(TEACHING_CHAIR_SUMMARY_FILENAME,result['downloads'])

    def test_shift_availability_export(self):
        raw=make_opd(availability_only=True)
        result=run_app({'schedule_app_mode':'Shift Availability Tracker','Upload md_opd.xlsx':Upload(raw,'md.xlsx')})
        self.assertIn('HOPE_DRIVE_weekly_grid.csv',result['downloads'])
        self.assertIn('HOPE_DRIVE_weekly_grid.xlsx',result['downloads'])

    def test_md_pa_conflicts_and_annotations(self):
        pa=make_opd(changes={('HOPE_DRIVE','B8'):'Adams, Alex ~ PA Learner'})
        result=run_app({'schedule_app_mode':'OPD MD PA Conflict Detector','md':Upload(self.raw,'md.xlsx'),
                       'pa':Upload(pa,'pa.xlsx'),'tog_annotated_downloads':True,'tog_same_site':True,
                       'tog_other_site':True,'tog_suggestions':True},secrets=self.config_values)
        self.assertIn('opd_double_bookings.csv',result['downloads'])
        self.assertIn('md_opd_annotated.xlsx',result['downloads'])
        self.assertIn('pa_opd_annotated.xlsx',result['downloads'])

    def test_format_opd_and_summary_zip(self):
        result=run_app({'schedule_app_mode':'Format OPD + Summary',
            '1) Upload one or more QGenda calendar Excel(s)':[Upload(make_qgenda(),'qgenda.xlsx')],
            "2) Upload Redcap Rotation list CSV (must have a 'legal_name' and 'start_date' column)":Upload(make_roster(),'rotation.csv'),
            'Generate OPD File For Sarah to Load Students':True})
        with ZipFile(io.BytesIO(result['downloads']['Batch_Output.zip'])) as z:
            self.assertEqual(set(z.namelist()),{'Updated_OPD.xlsx','Assignment_Summary.docx'})

if __name__=='__main__':
    unittest.main()
