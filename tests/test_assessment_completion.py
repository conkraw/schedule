"""Offline completion denominators, unique-form counting, links and soft alerts."""
from copy import deepcopy
from datetime import date
from io import BytesIO, StringIO
import csv
import json
from pathlib import Path
import unittest
from unittest.mock import patch
from zipfile import ZipFile
from cryptography.fernet import Fernet
from helpers import st, run_app
from test_learner_reach import scan_cells
from test_strict_reach_charts import ui_values
from test_reporting_dates import all_doc_text
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_scan_archives, teaching_filter_date_range
from schedule_app.services.student_continuity import student_continuity_counts
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks, student_name_key
from schedule_app.services.assessment_completion import (
    CLINICAL, HP, prepare_student_assessments, load_completion_inputs, build_completion_bundle,
    completion_rows, completion_display, unverified_bundle, completion_signature,
)
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.sections.assessment_completion import render_assessment_completion, P

NAME = 'Example, Avery'
HEADER = ['Course ID','Start Date','End Date','Student','Student External ID','Evaluator',
          'Evaluator Username','Evaluator Email','Evaluation','Form Record','Question ID','Question',
          'Answer text','Multiple Choice Value','Submit Date']


def record(fid='1', sid='id-alpha', student='Learner, Alpha; MD2028', form=CLINICAL,
           email='avery@example.edu', submitted='2026-09-10 23:59:59', **kw):
    row = dict(zip(HEADER, ['DEMO-101','2026-09-07','2026-09-13',student,sid,NAME,
                    'different_export_username',email,form,fid,'101','Ignored question',
                    'CONFIDENTIAL_COMMENT','5',submitted]))
    row.update(kw)
    return row


def raw_csv(rows):
    out = StringIO(newline=''); w=csv.DictWriter(out,fieldnames=HEADER);w.writeheader();w.writerows(rows)
    return out.getvalue().encode('utf-8-sig')


def cells(two=False):
    learner = 'Learner, Alpha; Learner, Beta' if two else 'Learner, Alpha'
    return {('NYES','B6'):f'{NAME} ~ {learner}', ('NYES','B8'):f'{NAME} ~ {learner}',
            ('NYES','C6'):f'{NAME} ~ {learner}', ('NYES','D6'):f'{NAME} ~ '}


class ParserTests(unittest.TestCase):
    def test_rows_questions_and_snapshots_do_not_multiply_forms(self):
        a=record(); b=record(**{'Question ID':'202','Answer text':'different ignored question'})
        d=prepare_student_assessments([('one',raw_csv([a,b])),('two',raw_csv([a,b]))])
        self.assertEqual(len(d['forms']),1)
        self.assertEqual(d['forms'][0]['username'],'avery')
        self.assertNotIn('CONFIDENTIAL_COMMENT',repr(d))
        self.assertNotIn('Ignored question',repr(d))

    def test_distinct_form_records_still_distinct_in_parser(self):
        d=prepare_student_assessments([('x',raw_csv([record('1'),record('2')]))])
        self.assertEqual(len(d['forms']),2)

    def test_only_two_requested_types_count_handoff_supplies_crosswalk(self):
        d=prepare_student_assessments([('x',raw_csv([record(form=HP),record('2',form='*PEDS Handoff',sid='id-z',student='Learner, Z')]))])
        self.assertEqual([r['target'] for r in d['forms']],['hp'])
        self.assertEqual(d['name_ids']['learner, z'],['id-z'])

    def test_empty_submission_not_completed(self):
        d=prepare_student_assessments([('x',raw_csv([record(submitted='')]))])
        self.assertEqual(d['forms'],[]); self.assertEqual(d['unsubmitted_forms'],1)

    def test_unsubmitted_then_submitted_snapshot_is_one_completed(self):
        d=prepare_student_assessments([('a',raw_csv([record(submitted='')])),('b',raw_csv([record()]))])
        self.assertEqual(len(d['forms']),1)

    def test_invalid_date_is_an_issue_not_zero(self):
        d=prepare_student_assessments([('x',raw_csv([record(submitted='not-a-date')]))])
        self.assertFalse(d['forms']);self.assertIn('Submit Date',d['issues'][0]['issue'])

    def test_missing_email_does_not_use_exported_evaluator_username(self):
        d=prepare_student_assessments([('x',raw_csv([record(email='')]))])
        self.assertFalse(d['forms']);self.assertIn('Evaluator Email',d['issues'][0]['issue'])
        self.assertEqual(d['issues'][0]['username'],'')

    def test_missing_external_id_flagged_not_replaced_by_name(self):
        d=prepare_student_assessments([('x',raw_csv([record(sid='')]))])
        self.assertFalse(d['forms']);self.assertIn('Student External ID',d['issues'][0]['issue'])

    def test_conflicting_form_student_identity_flagged(self):
        d=prepare_student_assessments([('x',raw_csv([record(),record(sid='id-other')]))])
        self.assertFalse(d['forms']);self.assertIn('Student External ID',d['issues'][0]['issue'])

    def test_conflicting_dates_same_form_flagged(self):
        d=prepare_student_assessments([('a',raw_csv([record()])),('b',raw_csv([record(submitted='2026-09-11')]))])
        self.assertFalse(d['forms']);self.assertIn('Submit Date',d['issues'][0]['issue'])

    def test_case_spaces_class_suffix_normalized_no_fuzzy(self):
        self.assertEqual(student_name_key(' Learner ,  Alpha; MD2028 '),'learner, alpha')
        self.assertNotEqual(student_name_key('Learner, Al'),student_name_key('Learner, Alpha'))
        self.assertEqual(student_name_key('Learner, Alpha; something'), 'learner, alpha; something')

    def test_external_ids_preserve_leading_zero(self):
        d=prepare_student_assessments([('x',raw_csv([record(sid='001234')]))])
        self.assertEqual(d['forms'][0]['external_id'],'001234')

    def test_metadata_changes_in_question_answers_ignored(self):
        d=prepare_student_assessments([('a',raw_csv([record()])),('b',raw_csv([record(**{'Answer text':'Other','Multiple Choice Value':'1'})]))])
        self.assertFalse(d['issues']);self.assertEqual(len(d['forms']),1)

    def test_same_form_number_different_form_types_separate(self):
        d=prepare_student_assessments([('x',raw_csv([record(),record(form=HP)]))])
        self.assertEqual(len(d['forms']),2)

    def test_exact_form_title_formatting_only(self):
        d=prepare_student_assessments([('x',raw_csv([record(form=' * CLINICAL  ASSESSMENT OF STUDENT ')]))])
        self.assertEqual(d['forms'][0]['target'],'clinical')


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.scan,self.repo,self.archive,self.secrets,self.raw=scan_cells(cells(two=True))
        self.period=ReportingPeriod('26-27',date(2026,9,7),date(2026,9,13))
        self.view=teaching_filter_date_range(self.scan,self.period)
        self.links=GitHubPreceptorOASISLinks(self.archive)
        self.links.save_username(NAME,'avery',expected=self.links.load())
        self.sources=GitHubOASISStudentEvaluations(self.archive)

    def source(self, rows=None):
        return self.sources.save(raw_csv(rows or [record(),record('2',sid='id-beta',student='Learner, Beta',form=HP)]))

    def bundle(self, rows=None):
        self.source(rows)
        self.inputs=load_completion_inputs(self.archive,self.view,[2026])
        b,u=build_completion_bundle(self.inputs,self.view,[2026]);return b,u,b['rows'][0]

    def test_three_shifts_over_two_days_qualifies_but_three_day_count_stays_zero(self):
        b,u,r=self.bundle()
        self.assertEqual(r['eligible_students'],2)
        self.assertEqual(student_continuity_counts(self.view,NAME,2026)['unique_students_3plus_days'],0)
        self.assertEqual(r['clinical_completion_pct'],50)
        self.assertEqual(r['hp_completion_pct'],50)
        self.assertEqual(r['either_completion_pct'],100)

    def test_repeated_completed_forms_for_same_student_count_once(self):
        b,u,r=self.bundle([record(),record('2'),record('3',form=HP),record('4',sid='id-beta',student='Learner, Beta',form='*PEDS Handoff')])
        self.assertEqual(r['clinical_students_evaluated'],1);self.assertEqual(r['clinical_forms_submitted'],2)
        self.assertEqual(r['either_completion_pct'],50)

    def test_duplicate_exports_do_not_inflate_forms_or_percent(self):
        self.source()
        rows=[record(),record('2',sid='id-beta',student='Learner, Beta',form=HP)]
        self.sources.save(raw_csv(rows+[record('3',sid='outside',student='Outside, Student')]))
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,_=build_completion_bundle(inputs,self.view,[2026])
        r=b['rows'][0];self.assertEqual(r['clinical_completion_pct'],50);self.assertEqual(r['eligible_students'],2)

    def test_unmatched_eligible_student_not_dropped_for_better_percentage(self):
        b,u,r=self.bundle([record()])
        self.assertEqual(r['eligible_students'],2)
        self.assertEqual(r['clinical_completion_pct'],50);self.assertEqual(len(u),1)
        self.assertEqual(r['assessment_status'],'Calculated')
        self.assertEqual(r['either_students_without_assessment'],1)

    def test_manual_external_id_makes_missing_assessment_zero_not_unverified(self):
        self.source([record()]);store=GitHubStudentAssessmentLinks(self.archive)
        store.save_link('Learner, Beta','id-beta',expected=store.load())
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,u=build_completion_bundle(inputs,self.view,[2026])
        self.assertFalse(u);self.assertEqual(b['rows'][0]['clinical_completion_pct'],50)
        self.assertEqual(b['rows'][0]['hp_completion_pct'],0)

    def test_missing_username_alerts_both_directions_not_stop(self):
        self.links.remove_username(NAME,expected=self.links.load())
        b,u,r=self.bundle();self.assertIn('username missing',r['assessment_status'])
        self.assertEqual({x['direction'] for x in b['warnings']}, {'Student → educator','Preceptor → student'})
        view=dict(self.view,assessment_completion=b)
        z,_=teaching_build_zip(view,[2026]);self.assertTrue(ZipFile(BytesIO(z)).testzip() is None)

    def test_no_student_sources_is_not_reported_as_zero(self):
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,_=build_completion_bundle(inputs,self.view,[2026])
        self.assertIsNone(b['rows'][0]['clinical_completion_pct'])
        self.assertIn('Not checked',b['rows'][0]['assessment_status'])

    def test_saved_feedback_absent_alert_not_zero_assumption(self):
        b,u,r=self.bundle();self.assertIsNone(r['student_feedback_evaluations'])
        self.assertTrue(any(x['direction']=='Student → educator' and 'Not checked' in x['issue'] for x in b['warnings']))

    def test_verified_summary_without_preceptor_flags_no_feedback(self):
        self.source();inputs=load_completion_inputs(self.archive,self.view,[2026])
        inputs['summaries']['2026']={'rows_by_id':{}}
        b,_=build_completion_bundle(inputs,self.view,[2026]);self.assertEqual(b['rows'][0]['student_feedback_evaluations'],0)
        self.assertTrue(any('No educator evaluation' in x['issue'] for x in b['warnings']))

    def test_feedback_present_clears_that_direction(self):
        self.source();inputs=load_completion_inputs(self.archive,self.view,[2026])
        inputs['summaries']['2026']={'rows_by_id':{'avery':{'evaluation_count':3}}}
        b,_=build_completion_bundle(inputs,self.view,[2026]);self.assertEqual(b['rows'][0]['student_feedback_evaluations'],3)
        self.assertFalse(any(x['direction']=='Student → educator' for x in b['warnings']))

    def test_no_preceptor_forms_found_alert(self):
        b,u,r=self.bundle([record(email='other@example.edu'),record('2',sid='id-beta',student='Learner, Beta',email='other@example.edu')])
        self.assertEqual(r['clinical_completion_pct'],0)
        self.assertTrue(any('No submitted Clinical' in x['issue'] for x in b['warnings']))

    def test_late_last_day_submission_included(self):
        b,u,r=self.bundle([record(submitted='2026-09-13 23:59:59'),record('2',sid='id-beta',student='Learner, Beta',form=HP)])
        self.assertEqual(r['clinical_completion_pct'],50)

    def test_outside_submit_date_not_counted_but_identity_can_match(self):
        b,u,r=self.bundle([record(submitted='2026-09-14'),record('2',sid='id-beta',student='Learner, Beta',submitted='2026-09-06')])
        self.assertEqual(r['clinical_completion_pct'],0)

    def test_exact_assignment_dates_filter_eligibility(self):
        self.source();self.view=teaching_filter_date_range(self.scan,ReportingPeriod('partial',date(2026,9,8),date(2026,9,13)))
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,u=build_completion_bundle(inputs,self.view,[2026])
        self.assertEqual(b['rows'][0]['eligible_students'],0)
        self.assertEqual(completion_display(b['rows'][0],'clinical'),'No eligible students')

    def test_other_preceptor_and_unassigned_students_not_in_numerator(self):
        b,u,r=self.bundle([record(),record('2',sid='id-beta',student='Learner, Beta',email='other@example.edu'),
                           record('3',sid='id-out',student='Out, Student')])
        self.assertEqual(r['clinical_students_evaluated'],1)
        self.assertEqual(r['eligible_students'],2)

    def test_new_snapshot_extends_union(self):
        self.source([record()]);self.source([record('2',sid='id-beta',student='Learner, Beta')])
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,u=build_completion_bundle(inputs,self.view,[2026])
        self.assertEqual(b['rows'][0]['clinical_completion_pct'],100)

    def test_student_external_id_unifies_explicit_name_aliases(self):
        self.source([record(sid='same-id'),record('2',sid='same-id',student='Learner, Beta')])
        inputs=load_completion_inputs(self.archive,self.view,[2026]);b,u=build_completion_bundle(inputs,self.view,[2026])
        self.assertEqual(b['rows'][0]['eligible_students'],1)
        self.assertEqual(b['rows'][0]['clinical_completion_pct'],100)

    def test_ambiguous_exact_name_multiple_ids_not_guessed(self):
        b,u,r=self.bundle([record(),record('2',sid='DIFFERENT'),record('3',sid='id-beta',student='Learner, Beta')])
        self.assertEqual(r['clinical_completion_pct'],50);self.assertTrue(u)
        self.assertTrue(r['assessment_status'].startswith('Provisional:'))

    def test_original_student_hours_not_doubled(self):
        b,_,_=self.bundle();from schedule_app.services.educational_time import teaching_time_rows
        old=teaching_time_rows(self.view,[2026]);new=teaching_time_rows(dict(self.view,assessment_completion=b),[2026])
        for field in ('educational_hours','total_scheduled_availability_hours','learner_reach_pct','scheduled_shifts','teaching_shifts'):
            self.assertEqual(old[0][field],new[0][field])
        self.assertEqual(new[0]['educational_hours'],12)

    def test_reports_and_csv_no_student_names_ids_scores_comments(self):
        b,u,r=self.bundle();view=dict(self.view,assessment_completion=b)
        z,_=teaching_build_zip(view,[2026])
        with ZipFile(BytesIO(z)) as out:
            self.assertIn('preceptor_student_assessment_completion.csv',out.namelist())
            for f in out.namelist():
                if f.endswith('.docx'):
                    text=all_doc_text(out.read(f))
                    self.assertIn('Documented assessment completion',text)
                    self.assertIn('1 / 2 (50.0%)',text)
                elif f.endswith('.csv'):
                    text=out.read(f).decode('utf-8-sig')
                else:continue
                for forbidden in ('Learner, Alpha','Learner, Beta','id-alpha','id-beta','CONFIDENTIAL_COMMENT'):
                    self.assertNotIn(forbidden,text)

    def test_scan_normal_output_has_no_new_student_data(self):
        text=repr(self.scan)
        self.assertNotIn('_assessment_student',text);self.assertNotIn('Learner, Alpha',text)
        rows=[];replay=teaching_scan_archives(self.archive,commit=self.scan['commit'],assessment_collector=rows.extend)
        self.assertTrue(rows);self.assertNotIn('Learner, Alpha',repr(replay))

    def test_replay_reads_old_displayed_opd_commit_after_revision(self):
        self.source();old=self.scan['commit']
        inputs=load_completion_inputs(self.archive,self.view,[2026])
        self.assertEqual(inputs['opd_commit'],old)
        self.assertEqual(len(inputs['assignments']),6)

    def test_context_change_prevents_stale_report(self):
        b,_,_=self.bundle()
        other=teaching_filter_date_range(self.scan,ReportingPeriod('other',date(2026,9,7),date(2026,9,12)))
        with self.assertRaises(OPDArchiveError): completion_rows(b,other,2026)

    def test_missing_forms_metadata_makes_unverified_not_misleading_zero(self):
        b,u,r=self.bundle([record(email=''),record('2',sid='id-beta',student='Learner, Beta')])
        self.assertIsNone(r['clinical_completion_pct']);self.assertIn('metadata',r['assessment_status'])

    def test_unrelated_missing_email_does_not_suppress_all_other_rates(self):
        b,u,r=self.bundle([record(),record('2',sid='id-beta',student='Learner, Beta',form=HP),
                           record('3',email='',**{'Evaluator':'Unrelated, Evaluator','Evaluator Username':'other'})])
        self.assertEqual(r['clinical_completion_pct'],50)
        self.assertTrue(any(x['preceptor_name']=='Unattributed / other evaluator' for x in b['warnings']))

    def test_multi_course_requires_selection(self):
        self.source([record(),record('2',sid='id-beta',student='Learner, Beta',**{'Course ID':'OTHER'})])
        inputs=load_completion_inputs(self.archive,self.view,[2026])
        with self.assertRaisesRegex(OPDArchiveError,'course'):build_completion_bundle(inputs,self.view,[2026])
        b,_=build_completion_bundle(inputs,self.view,[2026],courses=['DEMO-101'])
        self.assertEqual(b['rows'][0]['clinical_completion_pct'],50)

    def test_missing_data_ui_does_not_stop_zip(self):
        result=run_app(ui_values(**{P+'refresh':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        self.assertTrue(any(f.endswith('.zip') for f in result['downloads']))
        self.assertTrue(any('not' in msg.lower() for kind,msg in result['messages'] if kind=='warning'))

    def test_ui_end_to_end(self):
        self.source()
        result=run_app(ui_values(**{P+'refresh':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        self.assertTrue(any(f.endswith('.zip') for f in result['downloads']))
        z=next(v for k,v in result['downloads'].items() if k.endswith('.zip'))
        with ZipFile(BytesIO(z)) as archive:
            rows=list(csv.DictReader(StringIO(archive.read('preceptor_student_assessment_completion.csv').decode('utf-8-sig'))))
            self.assertEqual(rows[0]['clinical_completion_pct'],'50.0')

    def test_ui_disabled_no_added_data_or_requests(self):
        writes=self.repo.write_count
        with patch('schedule_app.sections.assessment_completion.load_completion_inputs') as mock:
            st.reset(values={P+'enabled':False})
            self.assertIsNone(render_assessment_completion(self.archive,self.view,[2026]));mock.assert_not_called()
        self.assertEqual(self.repo.write_count,writes)


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.scan,self.repo,self.archive,self.secrets,_=scan_cells(cells())
        self.service=GitHubStudentAssessmentLinks(self.archive)

    def test_save_reload_verified_only_ciphertext(self):
        first=self.service.load();saved=self.service.save_link('Private, Learner','000045',expected=first)
        self.assertEqual(self.service.load()['entries']['private, learner']['external_id'],'000045')
        self.assertNotIn(b'Private, Learner',self.repo.tree[self.service.path])
        self.assertNotIn(b'000045',self.repo.tree[self.service.path])

    def test_duplicate_value_no_extra_commit(self):
        state=self.service.save_link('Private, Learner','000045',expected=self.service.load())
        n=self.repo.write_count;self.service.save_link('Private, Learner','000045',expected=state)
        self.assertEqual(self.repo.write_count,n)

    def test_stale_save_rejected(self):
        state=self.service.load();self.service.save_link('Private, Learner','1',expected=state)
        with self.assertRaisesRegex(OPDArchiveError,'another session'):
            self.service.save_link('Other, Learner','2',expected=state)

    def test_remove_does_not_remove_other_link_or_opd(self):
        s=self.service.save_link('A','1',expected=self.service.load());s=self.service.save_link('B','2',expected=s)
        s=self.service.remove_link('A',expected=s)
        self.assertEqual(set(s['entries']),{'b'});self.assertEqual(len(self.archive.list_rotations()),1)

    def test_aliases_can_explicitly_share_one_external_id(self):
        s=self.service.save_link('A','1',expected=self.service.load());s=self.service.save_link('B','1',expected=s)
        self.assertEqual(len(s['entries']),2)

    def test_invalid_text_no_write(self):
        state=self.service.load();n=self.repo.write_count
        for sid in ('','  ','a\nb'):
            with self.assertRaises(OPDArchiveError): self.service.save_link('A',sid,expected=state)
        self.assertEqual(self.repo.write_count,n)

    def test_wrong_key_does_not_overwrite(self):
        from dataclasses import replace
        from schedule_app.services.opd_archive import GitHubOPDArchive
        self.service.save_link('A','1',expected=self.service.load());n=self.repo.write_count
        config=replace(self.archive.config,encryption_key=Fernet.generate_key().decode())
        other=GitHubStudentAssessmentLinks(GitHubOPDArchive(config,transport=self.repo))
        with self.assertRaises(OPDArchiveError):other.load()
        self.assertEqual(self.repo.write_count,n)


class PriorityTests(unittest.TestCase):
    def test_excluded_nursery_learners_do_not_enter_three_shift_denominator(self):
        rows={}
        for pos in ['B6','B8','C6']:
            rows['NYES',pos]=f'{NAME} ~ Clinic, Student'
            rows['PSHCH_NURSERY',pos]=f'{NAME} ~ Nursery, Student'
        scan,repo,a,secret,_=scan_cells(rows)
        links=GitHubPreceptorOASISLinks(a);links.save_username(NAME,'avery',expected=links.load())
        GitHubOASISStudentEvaluations(a).save(raw_csv([record(student='Clinic, Student',sid='clinic')]))
        inputs=load_completion_inputs(a,scan,[2026]);b,u=build_completion_bundle(inputs,scan,[2026])
        self.assertEqual(b['rows'][0]['eligible_students'],1)
        self.assertEqual(b['rows'][0]['clinical_completion_pct'],100);self.assertFalse(u)
        self.assertNotIn('Nursery, Student',repr(inputs['assignments']))

    def test_weekends_count_and_duplicates_do_not_inflate_threshold(self):
        rows={('NYES','G6'):f'{NAME} ~ Learner, Alpha',('NYES','G8'):f'{NAME} ~ Learner, Alpha',
              ('NYES','H6'):f'{NAME} ~ Learner, Alpha',('ETOWN','G6'):f'{NAME} ~ Learner, Alpha'}
        scan,repo,a,secret,_=scan_cells(rows)
        links=GitHubPreceptorOASISLinks(a);links.save_username(NAME,'avery',expected=links.load())
        GitHubOASISStudentEvaluations(a).save(raw_csv([record(submitted='2026-09-13')]))
        inputs=load_completion_inputs(a,scan,[2026]);b,u=build_completion_bundle(inputs,scan,[2026])
        self.assertEqual(len(inputs['assignments']),3)
        self.assertEqual(b['rows'][0]['clinical_completion_pct'],100)

    def test_two_identical_sessions_stay_below_threshold(self):
        rows={('NYES','B6'):f'{NAME} ~ Learner, Alpha',('ETOWN','B6'):f'{NAME} ~ Learner, Alpha',
              ('NYES','B8'):f'{NAME} ~ Learner, Alpha'}
        scan,repo,a,secret,_=scan_cells(rows)
        links=GitHubPreceptorOASISLinks(a);links.save_username(NAME,'avery',expected=links.load())
        GitHubOASISStudentEvaluations(a).save(raw_csv([record()]))
        inputs=load_completion_inputs(a,scan,[2026]);b,u=build_completion_bundle(inputs,scan,[2026])
        self.assertEqual(b['rows'][0]['eligible_students'],0)
        self.assertIsNone(b['rows'][0]['clinical_completion_pct'])
