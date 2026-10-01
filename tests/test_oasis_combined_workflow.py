"""Combined OASIS workflow regressions: synthetic data and simulated GitHub only."""
import ast
import base64
from dataclasses import replace
from datetime import date
import csv
import hashlib
import io
from pathlib import Path
import unittest
from unittest.mock import patch

from cryptography.fernet import Fernet
from helpers import login_for_test, st, FakeGitHub, FakeResponse, Upload, secret_settings, run_app, StopRun, ROOT
from test_oasis_educator_reports import row, comments, make_csv
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive, OPDArchiveError
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations
from schedule_app.services.oasis_educator_reports import OASISReportError
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames
from schedule_app.services.oasis_workflow import (
    WORKFLOW_VERSION, OASISOutputScope, GitHubOASISSummaries,
    load_cumulative_evaluations, make_period_summary, summary_csv, inspect_summary_csv,
)
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.reporting_presets import GitHubReportingPresets
from schedule_app.sections import oasis_workflow as ui, oasis_date_controls as dates


def csv_rows(raw):
    return list(csv.DictReader(io.StringIO(raw.decode('utf-8-sig'))))


class Fixture(unittest.TestCase):
    def setUp(self):
        self.secrets = secret_settings()
        st.reset(secrets=self.secrets)
        self.repo = FakeGitHub()
        self.config = OPDArchiveConfig(**self.secrets['opd_archive'])
        self.archive = GitHubOPDArchive(self.config, transport=self.repo)
        self.sources = GitHubOASISEvaluations(self.archive)
        self.outputs = GitHubOASISSummaries(self.archive)
        self.usernames = GitHubOASISUsernames(self.archive)
        self.period = ReportingPeriod('26-27', date(2026,3,1), date(2027,3,1))
        self.scope = OASISOutputScope(self.period, ('DEMO-101',), ('*Clinical Teaching Eval',))

    def source(self, rows=None):
        return self.sources.save(make_csv(rows))

    def prepared(self):
        return load_cumulative_evaluations(self.sources)

    def report(self, prepared=None, scope=None):
        prepared = prepared or self.prepared()
        names = self.usernames.load()
        report = make_period_summary(prepared, scope or self.scope, names['entries'])
        return prepared, names, report

    def save(self, prepared=None, scope=None):
        prepared, names, report = self.report(prepared, scope)
        return self.outputs.save(scope or self.scope, summary_csv(report), prepared=prepared, username_catalog=names)


class CumulativeTests(Fixture):
    def test_old_and_new_forms_accumulate(self):
        self.source([row('1'), row('2')]); self.source([row('2'), row('3')])
        _, _, result = self.report()
        self.assertEqual(result['evaluation_count'], 3)
        self.assertEqual(result['rows'][0]['q587_n'], 3)

    def test_prior_educator_not_in_new_file_retained(self):
        self.source([row('1')])
        self.source([row('2', **{'Evaluator':'Example, Bailey','Evaluator Email':'bb@example.edu',
                                'Evaluator Username':'bb','Evaluator External ID':'staff-002'})])
        self.assertEqual({r['record_id'] for r in self.report()[2]['rows']}, {'abc12','bb'})

    def test_does_not_average_prior_summary_averages(self):
        self.source([row('1',value='1'), row('2',value='1')]); self.save()
        self.source([row('3',value='5')]); saved=self.save()
        self.assertEqual(csv_rows(saved['raw'])[0]['q587_mean'], '2.33')

    def test_changed_filename_same_contents_no_new_snapshot(self):
        a=self.source(); self.assertEqual(self.source()['action'],'unchanged')
        self.assertEqual(self.sources.list_exports()['filenames'],[a['filename']])

    def test_repeated_comment_from_same_form_once(self):
        self.source([row('1')]+comments('1'))
        self.source([row('1')]+comments('1')+[row('2')])
        self.assertEqual(self.report()[2]['rows'][0]['strengths_comments'], '1. Clear explanations.')

    def test_same_comment_different_forms_both_retained(self):
        self.source([row('1')]+comments('1')); self.source([row('2')]+comments('2'))
        self.assertEqual(self.report()[2]['rows'][0]['strengths_comments'].count('Clear explanations.'),2)

    def test_conflicting_existing_answer_blocks_not_guessed(self):
        self.source([row('1',value='1')]); self.source([row('1',value='5')])
        with self.assertRaises(OASISReportError) as err: self.prepared()
        self.assertTrue(err.exception.issues)
        self.assertNotIn('SYNTHETIC LEARNER', repr(err.exception.issues))
        self.assertFalse(any('/oasis_reports/' in k for k in self.repo.tree))

    def test_archive_reads_pinned_and_does_not_resave(self):
        self.source(); before=self.repo.write_count
        prepared=self.prepared()
        self.assertEqual(prepared['archive_commit'],self.repo.head)
        self.assertEqual(self.repo.write_count,before)

    def test_previous_source_rows_preserved_when_new_file_omits_them(self):
        self.source([row('1'),row('2')]); self.source([row('3')])
        self.assertEqual(self.report()[2]['evaluation_count'],3)

    def test_all_question_rows_one_evaluation(self):
        self.source([row('1','587'), row('1','589')]+comments('1'))
        self.assertEqual(self.report()[2]['rows'][0]['evaluation_count'],1)

    def test_unsubmitted_forms_excluded(self):
        self.source([row('1'),row('2',**{'Submit Date':''})])
        p=self.prepared(); self.assertEqual(p['unsubmitted_forms_excluded'],1)
        self.assertEqual(self.report(p)[2]['evaluation_count'],1)

    def test_invalid_submit_date_blocks(self):
        self.source([row(**{'Submit Date':'bad-date'})])
        with self.assertRaises(OASISReportError):self.prepared()

    def test_question_number_ignored_for_matching(self):
        self.source([row('1','589',**{'Question Number':'3'})])
        self.source([row('1','589',**{'Question Number':'2'}),row('2','589')])
        self.assertEqual(self.report()[2]['rows'][0]['q589_n'],2)

    def test_no_archive_available(self):
        with self.assertRaisesRegex(OASISReportError,'No OASIS'):self.prepared()


class DateFilterTests(Fixture):
    def test_includes_whole_start_and_end_dates(self):
        values=['2026-02-28 23:59:59','2026-03-01 00:00:00','2027-03-01 23:59:59','2027-03-02 00:00:00']
        self.source([row(str(i),**{'Submit Date':v}) for i,v in enumerate(values)])
        self.assertEqual(self.report()[2]['evaluation_count'],2)

    def test_uses_submit_not_rotation_start(self):
        self.source([row('1',**{'Start Date':'2025-01-01','End Date':'2025-01-31','Submit Date':'2026-06-01 11:00:00'}),
                     row('2',**{'Submit Date':'2027-05-01 00:00:00'})])
        self.assertEqual(self.report()[2]['evaluation_count'],1)

    def test_crossing_july_stays_one_educator_row(self):
        self.source([row('1',**{'Submit Date':'2026-06-30'}),row('2',**{'Submit Date':'2026-07-01'})])
        s=self.report()[2]
        self.assertEqual(len(s['rows']),1);self.assertEqual(s['rows'][0]['academic_year'],'26-27')

    def test_weekends_included(self):
        self.source([row('1',**{'Submit Date':'2026-03-07'}),row('2',**{'Submit Date':'2026-03-08'})])
        self.assertEqual(self.report()[2]['evaluation_count'],2)

    def test_same_day_period(self):
        self.source([row('1',**{'Submit Date':'2026-04-11 23:59:59'})])
        scope=OASISOutputScope(ReportingPeriod('day',date(2026,4,11),date(2026,4,11)),self.scope.courses,self.scope.evaluation_types)
        self.assertEqual(self.report(scope=scope)[2]['evaluation_count'],1)

    def test_empty_period_not_saved(self):
        self.source()
        scope=OASISOutputScope(ReportingPeriod('future',date(2029,1,1),date(2029,2,1)),self.scope.courses,self.scope.evaluation_types)
        with self.assertRaisesRegex(OASISReportError,'No submitted'):self.save(scope=scope)
        self.assertEqual(len(self.repo.tree),1)

    def test_date_metadata_appended_original_record_id_first(self):
        self.source(); report=self.report()[2]
        self.assertEqual(report['columns'][0],'record_id')
        self.assertEqual(report['columns'][-4:],['academic_year','report_start_date','report_end_date','date_basis'])
        data=csv_rows(summary_csv(report))[0]
        self.assertEqual(data['date_basis'],'Submit Date')
        self.assertEqual(data['report_end_date'],'2027-03-01')

    def test_missing_username_outside_period_does_not_block(self):
        self.source([row('1'),row('2',**{'Evaluator':'No, Email','Evaluator Email':'','Evaluator Username':'missing',
                                       'Evaluator External ID':'missing','Submit Date':'2029-04-11'})])
        self.assertFalse(self.report()[2]['issues']); self.save()


class OutputTests(Fixture):
    def test_only_summary_csv_output_and_encrypted_original(self):
        original=self.source(); saved=self.save()
        self.assertEqual(set(self.repo.tree),{original['path'],saved['path']})
        self.assertTrue(saved['path'].endswith('.csv.enc'))
        for token in self.repo.tree.values():
            self.assertNotIn(b'Example, Avery',token)
            self.assertNotIn(b'Clear explanations.',token)
        self.assertEqual(self.config.cipher().decrypt(self.repo.tree[saved['path']]),saved['raw'])

    def test_same_scope_replaces_one_current_output(self):
        self.source([row('1')]); a=self.save()
        self.source([row('2')]); b=self.save()
        self.assertEqual(a['filename'],b['filename']);self.assertEqual(b['action'],'updated')
        self.assertEqual(csv_rows(b['raw'])[0]['evaluation_count'],'2')
        self.assertEqual(len(self.outputs.list_outputs()['filenames']),1)

    def test_identical_result_no_unnecessary_commit(self):
        self.source(); a=self.save(); before=self.repo.write_count; b=self.save()
        self.assertEqual(b['action'],'unchanged');self.assertEqual(before,self.repo.write_count)
        self.assertEqual(a['raw'],b['raw'])

    def test_other_period_separate_and_old_preserved(self):
        self.source(); a=self.save()
        scope=OASISOutputScope(ReportingPeriod('other',date(2026,4,1),date(2026,5,1)),self.scope.courses,self.scope.evaluation_types)
        b=self.save(scope=scope)
        self.assertNotEqual(a['path'],b['path']);self.assertEqual(len(self.outputs.list_outputs()['filenames']),2)

    def test_label_change_same_period_replaces(self):
        self.source(); a=self.save()
        scope=OASISOutputScope(ReportingPeriod('new label',self.period.start,self.period.end),self.scope.courses,self.scope.evaluation_types)
        b=self.save(scope=scope)
        self.assertEqual(a['path'],b['path']);self.assertEqual(csv_rows(b['raw'])[0]['academic_year'],'new label')

    def test_selector_order_not_filename_changes(self):
        p=self.period
        self.assertEqual(OASISOutputScope(p,('b','a'),('b','a')).filename,OASISOutputScope(p,('a','b'),('a','b')).filename)

    def test_different_course_forms_different_output_path(self):
        self.assertNotEqual(self.scope.filename,OASISOutputScope(self.period,('other',),self.scope.evaluation_types).filename)
        self.assertNotEqual(self.scope.filename,OASISOutputScope(self.period,self.scope.courses,('other',)).filename)

    def test_label_formula_safe(self):
        self.source()
        scope=OASISOutputScope(ReportingPeriod('=not a formula',self.period.start,self.period.end),self.scope.courses,self.scope.evaluation_types)
        saved=self.save(scope=scope)
        self.assertEqual(csv_rows(saved['raw'])[0]['academic_year'],"'=not a formula")

    def test_wrong_label_rejected(self):
        self.source(); prepared,names,report=self.report()
        report['rows'][0]['academic_year']='wrong'
        with self.assertRaisesRegex(OASISReportError,'label'):
            self.outputs.save(self.scope,summary_csv(report),prepared=prepared,username_catalog=names)

    def test_reload_exact_original_summary_bytes(self):
        self.source(); saved=self.save()
        self.assertEqual(self.outputs.load(saved['filename'])['raw'],saved['raw'])

    def test_previous_key_can_recover_output(self):
        self.source(); saved=self.save()
        changed=replace(self.config,encryption_key=Fernet.generate_key().decode(),previous_encryption_keys=(self.config.encryption_key,))
        other=GitHubOASISSummaries(GitHubOPDArchive(changed,transport=self.repo))
        self.assertEqual(other.load(saved['filename'])['raw'],saved['raw'])

    def test_wrong_key_no_overwrite(self):
        self.source(); saved=self.save();before=self.repo.write_count
        changed=replace(self.config,encryption_key=Fernet.generate_key().decode())
        with self.assertRaisesRegex(OASISReportError,'decrypt'):
            GitHubOASISSummaries(GitHubOPDArchive(changed,transport=self.repo)).load(saved['filename'])
        self.assertEqual(before,self.repo.write_count)

    def test_corrupt_output_not_replaced(self):
        self.source(); a=self.save();self.repo.tree[a['path']]=b'corrupt';self.repo._commit()
        before=self.repo.write_count
        with self.assertRaises(OASISReportError):self.save()
        self.assertEqual(before,self.repo.write_count)

    def test_large_file_raw_fallback(self):
        from test_oasis_evaluations import RecordingGitHub
        repo=RecordingGitHub();repo.raw_fallback=True
        archive=GitHubOPDArchive(self.config,transport=repo)
        sources=GitHubOASISEvaluations(archive);sources.save(make_csv())
        p=load_cumulative_evaluations(sources);n=GitHubOASISUsernames(archive).load()
        outputs=GitHubOASISSummaries(archive)
        saved=outputs.save(self.scope,summary_csv(make_period_summary(p,self.scope)),prepared=p,username_catalog=n)
        self.assertEqual(outputs.load(saved['filename'])['raw'],saved['raw'])

    def test_invalid_summary_header_rejected_before_write(self):
        self.source();p,n,s=self.report();before=self.repo.write_count
        with self.assertRaises(OASISReportError):
            self.outputs.save(self.scope,b'invalid,headers\na,b\n',prepared=p,username_catalog=n)
        self.assertEqual(before,self.repo.write_count)

    def test_duplicate_ids_rejected_before_write(self):
        self.source();p,n,s=self.report();s['rows'].append(dict(s['rows'][0]))
        before=self.repo.write_count
        with self.assertRaises(OASISReportError):
            self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)
        self.assertEqual(before,self.repo.write_count)

    def test_wrong_archive_scope_blocks(self):
        self.source();p,n,s=self.report();p['archive_scope']='changed'
        with self.assertRaisesRegex(OASISReportError,'settings changed'):
            self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)

    def test_concurrent_source_arrival_during_output_write_not_confirmed(self):
        self.source();p,n,s=self.report()
        extra=make_csv([row('2')]); details=__import__('schedule_app.services.oasis_evaluations',fromlist=['inspect_oasis_csv']).inspect_oasis_csv(extra)
        name=self.sources._candidate_names(extra,details)[0]
        original=self.repo.request
        def request(method,url,**kwargs):
            if method=='PUT' and '/oasis_reports/' in url:
                self.repo.tree[self.sources.path_for(name)]=self.config.cipher().encrypt(extra)
                self.repo._commit()
            return original(method,url,**kwargs)
        with patch.object(self.repo,'request',side_effect=request):
            with self.assertRaisesRegex(OASISReportError,'source archive changed'):
                self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)

    def test_no_path_injection(self):
        for name in ('../../x.csv.enc','OPD_2026-03-02.xlsx.enc','oasis_educator_usernames.json.enc'):
            with self.assertRaises(OASISReportError):self.outputs.load(name)

    def test_no_record_ids_no_output(self):
        self.source([row(**{'Evaluator Email':''})]);p,n,s=self.report()
        self.assertTrue(s['issues'])
        with self.assertRaises(OASISReportError):summary_csv(s)

    def test_username_fix_persisted_and_used(self):
        self.source([row(**{'Evaluator Email':''})]);p,n,s=self.report()
        educator=s['educators'][0]
        self.usernames.save(educator['educator_key'],educator['educator_name'],'newname',expected=n)
        saved=self.save();r=csv_rows(saved['raw'])[0]
        self.assertEqual((r['record_id'],r['email_missing'],r['record_id_source']),('newname','YES','manual_username'))

    def test_changed_sources_block_stale_summary(self):
        self.source([row('1')]);p,n,s=self.report();self.source([row('2')])
        with self.assertRaisesRegex(OASISReportError,'source archive changed'):
            self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)
        self.assertEqual(self.outputs.list_outputs()['filenames'],[])

    def test_changed_usernames_block_stale_summary(self):
        self.source();p,n,s=self.report();e=s['educators'][0]
        self.usernames.save(e['educator_key'],e['educator_name'],'newid',expected=n)
        with self.assertRaisesRegex(OASISReportError,'usernames changed'):
            self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)

    def test_opd_or_preset_updates_do_not_invalidate_oasis_sources(self):
        self.source();p,n,s=self.report()
        self.repo.tree['opd_archive/OPD_2026-03-02.xlsx.enc']=b'irrelevant';self.repo._commit()
        saved=self.outputs.save(self.scope,summary_csv(s),prepared=p,username_catalog=n)
        self.assertEqual(saved['action'],'created')

    def test_input_sources_unchanged_by_summary(self):
        a=self.source(); token=self.repo.tree[a['path']];self.save()
        self.assertEqual(token,self.repo.tree[a['path']])


class PresetTests(Fixture):
    def setUp(self):
        super().setUp()
        login_for_test()

    def test_shared_presets_reused_and_deleted_without_deleting_data(self):
        self.source(); saved=self.save(); service=GitHubReportingPresets(self.archive)
        result=service.save('Evaluation year',self.period,service.load())
        st.session_state[dates.P+'snapshot']=result['snapshot']
        st.session_state[dates.P+'choice']=result['preset_id']
        dates._load_preset()
        self.assertEqual(st.session_state[dates.P+'applied'],self.period)
        key=dates._confirmation('delete',result['snapshot'],result['preset_id']);st.session_state[key]=True
        dates._delete_preset(service,result['preset_id'],key)
        self.assertEqual(service.load()['presets'],[])
        self.assertIn(saved['path'],self.repo.tree)
        self.assertEqual(st.session_state[dates.P+'applied'],self.period)

    def test_apply_dates_invalidates_report_not_loaded_data(self):
        st.session_state.update({dates.P+'start':self.period.start,dates.P+'end':self.period.end,
                                dates.P+'label':'new',ui.P+'receipt':{'old':True},ui.P+'prepared':{'stay':True}})
        dates._apply_period()
        self.assertNotIn(ui.P+'receipt',st.session_state)
        self.assertEqual(st.session_state[ui.P+'prepared'],{'stay':True})

    def test_invalid_dates_remove_previous_active_period(self):
        st.session_state.update({dates.P+'applied':self.period,dates.P+'label':'bad',
                                dates.P+'start':self.period.end,dates.P+'end':self.period.start})
        dates._apply_period();self.assertNotIn(dates.P+'applied',st.session_state)
        self.assertIn(dates.P+'error',st.session_state)

    def test_delete_requires_confirmation(self):
        service=GitHubReportingPresets(self.archive);result=service.save('a',self.period,service.load())
        st.session_state[dates.P+'snapshot']=result['snapshot']
        dates._delete_preset(service,result['preset_id'],'wrong')
        self.assertEqual(len(service.load()['presets']),1)


class InterfaceTests(Fixture):
    def page(self, values=None, state=None, period=True):
        scope=(WORKFLOW_VERSION,self.config.signature())
        initial={ui.P+'scope':scope,dates.P+'scope':self.config.signature()}
        if period:initial[dates.P+'applied']=self.period
        initial.update(state or {})
        return run_app({'schedule_app_mode':'OER',**(values or {})},
                       secrets=self.secrets,state=initial,repo=self.repo, evaluation_login=True)

    def test_upload_auto_archives_and_saves_only_output_csv(self):
        result=self.page({ui.P+'upload':Upload(make_csv(),'input.csv')})
        self.assertEqual(len(self.repo.tree),2)
        self.assertTrue(any('Output CSV created' in m for kind,m in result['messages']))
        self.assertEqual(len(result['downloads']),1)
        self.assertTrue(all(k.endswith('.csv') for k in result['downloads']))
        self.assertFalse(any(k.endswith('.zip') for k in result['downloads']))

    def test_missing_username_archived_but_report_waits(self):
        result=self.page({ui.P+'upload':Upload(make_csv([row(**{'Evaluator Email':''})]),'input.csv')})
        self.assertEqual(len(self.repo.tree),1)
        self.assertNotIn(ui.P+'receipt',result['state'])
        self.assertTrue(any('Action needed' in m for _,m in result['messages']))

    def test_no_dates_still_archive_no_summary(self):
        result=self.page({ui.P+'upload':Upload(make_csv(),'input.csv')},period=False)
        self.assertEqual(len(self.repo.tree),1);self.assertFalse(result['downloads'])

    def test_upload_additions_rebuilds_cumulative_summary(self):
        first=self.page({ui.P+'upload':Upload(make_csv([row('1')]),'first.csv')})
        state=dict(first['state'])
        result=self.page({ui.P+'upload':Upload(make_csv([row('2')]),'second.csv')},state)
        self.assertEqual(csv_rows(result['state'][ui.P+'receipt']['raw'])[0]['evaluation_count'],'2')
        self.assertEqual(len(self.outputs.list_outputs()['filenames']),1)

    def test_rerun_no_duplicate_commits(self):
        upload=Upload(make_csv(),'original.csv')
        first=self.page({ui.P+'upload':upload});before=self.repo.write_count
        self.page({ui.P+'upload':upload},dict(first['state']))
        self.assertEqual(self.repo.write_count,before)

    def test_fresh_session_reads_all_old_sources(self):
        self.source([row('1')]);self.source([row('2')])
        result=self.page()
        self.assertEqual(csv_rows(result['state'][ui.P+'receipt']['raw'])[0]['evaluation_count'],'2')

    def test_failed_upload_clears_old_download(self):
        first=self.page({ui.P+'upload':Upload(make_csv(),'good.csv')})
        result=self.page({ui.P+'upload':Upload(b'bad file','bad.csv')},dict(first['state']))
        self.assertFalse(result['downloads']);self.assertNotIn(ui.P+'receipt',result['state'])

    def test_source_conflict_no_false_output_success(self):
        self.source([row(value='1')]);self.source([row(value='5')])
        result=self.page()
        self.assertFalse(result['downloads'])
        self.assertFalse(any('Output CSV created' in m for _,m in result['messages']))

    def test_optional_saved_output_decrypt_download(self):
        self.source();saved=self.save()
        result=self.page({ui.P+'list_outputs':True,ui.P+'load_output':True})
        self.assertIn(saved['filename'].removesuffix('.enc'),result['downloads'])

    def test_missing_username_fix_triggers_save_after_rerun(self):
        upload=Upload(make_csv([row(**{'Evaluator Email':''})]),'missing.csv')
        first=self.page({ui.P+'upload':upload})
        key=hashlib.sha256(b'ext:staff-001').hexdigest()[:16]
        fixed=self.page({ui.P+'upload':upload,ui.P+'username_'+key:'verified_name',ui.P+'apply_username':True},first['state'])
        final=self.page({ui.P+'upload':upload},fixed['state'])
        self.assertEqual(csv_rows(final['state'][ui.P+'receipt']['raw'])[0]['record_id'],'verified_name')
        self.assertFalse(any('Action needed' in text for kind,text in final['messages']))

    def test_failed_output_does_not_repeat_until_retry(self):
        upload=Upload(make_csv(),'input.csv')
        with patch.object(GitHubOASISSummaries,'save',side_effect=OASISReportError('Temporary save failure')) as save:
            first=self.page({ui.P+'upload':upload})
            self.assertFalse(first['downloads'])
            second=self.page({ui.P+'upload':upload},first['state'])
            self.assertEqual(save.call_count,1)
        final=self.page({ui.P+'upload':upload,ui.P+'retry_output':True},second['state'])
        self.assertIn(ui.P+'receipt',final['state'])

    def test_original_recovery_kept_in_combined_page(self):
        a=self.source()
        result=self.page({ui.P+'list_originals':True,ui.P+'load_original':True},period=False)
        self.assertEqual(result['downloads'][a['filename'].removesuffix('.enc')],__import__('schedule_app.services.oasis_privacy',fromlist=['minimize_oasis_csv']).minimize_oasis_csv(make_csv(),'educator')['raw'])

    def test_new_empty_period_clears_old_download_not_saved_file(self):
        first=self.page({ui.P+'upload':Upload(make_csv(),'input.csv')})
        state=first['state'];path=state[ui.P+'receipt']['path']
        state[dates.P+'applied']=ReportingPeriod('other',date(2029,1,1),date(2029,2,1))
        result=self.page(state=state)
        self.assertFalse(result['downloads']);self.assertIn(path,self.repo.tree)

    def test_old_selected_menu_state_migrated(self):
        result=run_app({},secrets=self.secrets,repo=self.repo,state={'schedule_app_mode':'OASIS Educator Reports'})
        self.assertEqual(result['state']['schedule_app_mode'],'OER')

    def test_launcher_has_combined_menu_not_two_oasis_choices(self):
        source=(ROOT/'app_sch_2026.py').read_text()
        tree=ast.parse(source)
        sections=next(ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign)
                      and any(isinstance(t,ast.Name) and t.id=='SECTIONS' for t in n.targets))
        self.assertEqual([k for k in sections if k == 'OER'],['OER'])


if __name__ == '__main__':unittest.main()
