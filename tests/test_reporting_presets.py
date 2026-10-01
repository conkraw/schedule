"""Encrypted GitHub date presets; offline transport and UI callback regression tests."""
import base64
from copy import deepcopy
from datetime import date
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from cryptography.fernet import Fernet
import requests
from helpers import login_for_test, FakeGitHub, FakeResponse, st, secret_settings, run_app
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.reporting_presets import (
    GitHubReportingPresets, ReportingPresetError, PRESET_FILENAME,
    MAX_ENCRYPTED_BYTES, MAX_PRESETS, normalize_preset_name, preset_name_key,
    period_from_preset, _validate_catalog, _empty_catalog,
)
from schedule_app.sections import reporting_date_controls as ui


class RecordingGitHub(FakeGitHub):
    def __init__(self):
        self.calls=[]
        super().__init__()

    def request(self, method, url, **kwargs):
        self.calls.append((method,url,deepcopy(kwargs.get('json')),deepcopy(kwargs.get('params'))))
        if method == 'PUT':
            path=url.split('/contents/',1)[1]
            body=kwargs['json']
            if path not in self.tree and body.get('sha') is not None:
                return FakeResponse(409)
        return super().request(method,url,**kwargs)


class PresetFixture(unittest.TestCase):
    def setUp(self):
        self.secrets=secret_settings()
        st.reset(secrets=self.secrets)
        login_for_test()
        self.repo=RecordingGitHub()
        self.config=OPDArchiveConfig(**self.secrets['opd_archive'])
        self.archive=GitHubOPDArchive(self.config,transport=self.repo)
        self.service=GitHubReportingPresets(self.archive)
        self.period=ReportingPeriod('26-27',date(2026,2,17),date(2027,3,16))

    def create(self,name='Teaching year',period=None,snapshot=None):
        return self.service.save(name,period or self.period,snapshot or self.service.load())

    def raw_catalog(self):
        token=self.repo.tree[self.service.path]
        return json.loads(self.config.cipher().decrypt(token))

    def put_catalog(self,data,*,encrypt=True):
        raw=json.dumps(data).encode() if not isinstance(data,bytes) else data
        self.repo.tree[self.service.path]=self.config.cipher().encrypt(raw) if encrypt else raw
        self.repo._commit()

    def seed_state(self,receipt=None):
        receipt=receipt or self.create()
        snapshot=receipt['snapshot']
        st.reset(secrets=self.secrets,state={
            ui.SCOPE:self.config.signature(),ui.SNAPSHOT:snapshot,ui.CHOICE:receipt['preset_id'],
            ui.NAME:'Teaching year','teaching_reporting_mode':'Custom dates',
            'teaching_period_start':self.period.start,'teaching_period_end':self.period.end,
            'teaching_period_label':self.period.label,
        })
        login_for_test()
        return receipt

    def request_patch(self):
        return patch('schedule_app.services.opd_archive.requests.request',self.repo.request)


class PresetStorageTests(PresetFixture):
    def test_empty_repository_requires_no_write(self):
        result=self.service.load()
        self.assertEqual(result['presets'],[])
        self.assertIsNone(result['sha'])
        self.assertEqual(self.repo.write_count,0)

    def test_encrypted_roundtrip_exact_dates_and_label(self):
        saved=self.create()
        loaded=self.service.load()
        self.assertEqual(loaded,saved['snapshot'])
        self.assertEqual(period_from_preset(loaded['presets'][0]),self.period)
        self.assertNotIn(b'Teaching year',self.repo.tree[self.service.path])
        self.assertNotIn(b'26-27',self.repo.tree[self.service.path])
        self.assertEqual(self.service.path,'opd_archive/'+PRESET_FILENAME)

    def test_create_multiple_presets_independently(self):
        first=self.create()
        other=ReportingPeriod('27-28',date(2027,3,17),date(2028,4,15))
        second=self.create('Next cohort',other,first['snapshot'])
        self.assertEqual(len(second['snapshot']['presets']),2)
        self.assertEqual({period_from_preset(p) for p in second['snapshot']['presets']},{self.period,other})

    def test_another_session_reads_all_saved_presets(self):
        first=self.create()
        second=GitHubReportingPresets(GitHubOPDArchive(self.config,transport=self.repo))
        self.assertEqual(second.load()['presets'],first['snapshot']['presets'])

    def test_same_name_requires_explicit_replace(self):
        saved=self.create()
        with self.assertRaisesRegex(ReportingPresetError,'already exists'):
            self.service.save(' teaching  YEAR ',self.period,saved['snapshot'])
        self.assertEqual(self.repo.write_count,1)

    def test_same_values_no_unnecessary_commit(self):
        saved=self.create()
        again=self.service.save('Teaching year',self.period,saved['snapshot'],replace_id=saved['preset_id'])
        self.assertEqual(again['action'],'unchanged')
        self.assertEqual(self.repo.write_count,1)

    def test_explicit_replace_retains_other_presets(self):
        first=self.create()
        second=self.create('Other',snapshot=first['snapshot'])
        new=ReportingPeriod('Revised cohort',date(2026,2,18),date(2027,4,1))
        updated=self.service.save('Teaching year',new,second['snapshot'],replace_id=first['preset_id'])
        by_id={p['id']:p for p in updated['snapshot']['presets']}
        self.assertEqual(len(by_id),2)
        self.assertEqual(period_from_preset(by_id[first['preset_id']]),new)
        self.assertEqual(period_from_preset(by_id[second['preset_id']]),self.period)
        self.assertEqual(by_id[first['preset_id']]['created_at'],first['snapshot']['presets'][0]['created_at'])

    def test_delete_only_selected_preset(self):
        first=self.create()
        second=self.create('Other',snapshot=first['snapshot'])
        result=self.service.delete(first['preset_id'],second['snapshot'])
        self.assertEqual([p['id'] for p in result['snapshot']['presets']],[second['preset_id']])

    def test_delete_last_retains_empty_encrypted_catalog(self):
        first=self.create()
        self.service.delete(first['preset_id'],first['snapshot'])
        self.assertEqual(self.service.load()['presets'],[])
        self.assertIn(self.service.path,self.repo.tree)
        self.assertEqual(self.raw_catalog()['presets'],[])

    def test_delete_never_changes_opds(self):
        path='opd_archive/OPD_2026-09-07.xlsx.enc'
        raw=b'Unchanged encrypted OPD fixture'
        self.repo.tree[path]=raw; self.repo._commit()
        first=self.create()
        self.service.delete(first['preset_id'],first['snapshot'])
        self.assertEqual(self.repo.tree[path],raw)
        self.assertTrue(all(url.endswith('/'+PRESET_FILENAME) for method,url,_,_ in self.repo.calls if method=='PUT'))
        self.assertFalse(any(method=='DELETE' for method,*_ in self.repo.calls))

    def test_catalog_not_listed_as_opd_rotation(self):
        self.create()
        self.repo.tree['opd_archive/OPD_2026-09-07.xlsx.enc']=b'fixture';self.repo._commit()
        self.assertEqual(self.archive.list_rotations(),[date(2026,9,7)])

    def test_stale_save_does_not_overwrite_newer_catalog(self):
        old=self.service.load()
        first=self.create(snapshot=old)
        with self.assertRaisesRegex(ReportingPresetError,'changed in GitHub'):
            self.service.save('Another',self.period,old)
        self.assertEqual(self.service.load()['presets'],first['snapshot']['presets'])

    def test_stale_delete_does_not_remove_newer_preset(self):
        first=self.create()
        other=self.create('Other',snapshot=first['snapshot'])
        with self.assertRaisesRegex(ReportingPresetError,'changed in GitHub'):
            self.service.delete(first['preset_id'],first['snapshot'])
        self.assertEqual(len(self.service.load()['presets']),2)

    def test_unrelated_opd_commit_does_not_block_save(self):
        first=self.create()
        self.repo.tree['opd_archive/OPD_2026-09-07.xlsx.enc']=b'fixture';self.repo._commit()
        self.create('Other',snapshot=first['snapshot'])
        self.assertEqual(len(self.service.load()['presets']),2)

    def test_put_conflict_is_not_retried(self):
        first=self.create()
        actual=self.repo.request
        puts=[]
        def conflict(method,url,**kwargs):
            if method=='PUT':
                puts.append(url);return FakeResponse(409)
            return actual(method,url,**kwargs)
        self.archive.transport=type('T',(),{'request':staticmethod(conflict)})()
        with self.assertRaisesRegex(ReportingPresetError,'Refresh saved presets'):
            self.service.save('Other',self.period,first['snapshot'])
        self.assertEqual(len(puts),1)
        self.assertEqual(self.repo.write_count,1)

    def test_success_is_verified_at_returned_commit(self):
        self.create()
        last=self.repo.calls[-1]
        self.assertEqual(last[0],'GET')
        self.assertEqual(last[3],{'ref':self.repo.head})

    def test_unconfirmed_write_not_reported_successful(self):
        original=self.repo.request
        def no_commit(method,url,**kwargs):
            return FakeResponse(201,{}) if method=='PUT' else original(method,url,**kwargs)
        self.archive.transport=type('T',(),{'request':staticmethod(no_commit)})()
        with self.assertRaisesRegex(ReportingPresetError,'did not confirm'):
            self.create()

    def test_wrong_key_refuses_read_or_overwrite(self):
        self.create()
        other=OPDArchiveConfig(**secret_settings()['opd_archive'])
        service=GitHubReportingPresets(GitHubOPDArchive(other,transport=self.repo))
        with self.assertRaisesRegex(ReportingPresetError,'could not be decrypted'):service.load()
        self.assertEqual(self.repo.write_count,1)

    def test_previous_key_can_read_then_save_with_current_key(self):
        self.create()
        new=Fernet.generate_key().decode()
        cfg=OPDArchiveConfig(**{**self.secrets['opd_archive'],'encryption_key':new},
                            previous_encryption_keys=(self.config.encryption_key,))
        service=GitHubReportingPresets(GitHubOPDArchive(cfg,transport=self.repo))
        service.save('Other',self.period,service.load())
        self.assertEqual(len(json.loads(Fernet(new.encode()).decrypt(self.repo.tree[self.service.path]))['presets']),2)

    def test_catalog_schema_and_dates_validated(self):
        self.create()
        original=self.raw_catalog()
        variants=[]
        a=deepcopy(original);a['schema_version']=2;variants.append(a)
        a=deepcopy(original);a['schema_version']=True;variants.append(a)
        a=deepcopy(original);a['presets'][0]['period']['start_date']='wrong';variants.append(a)
        a=deepcopy(original);a['presets'][0]['period']['end_date']='2020-01-01';variants.append(a)
        a=deepcopy(original);a['presets'][0]['period']['label']='';variants.append(a)
        a=deepcopy(original);a['presets'][0]['updated_at']='not a date';variants.append(a)
        a=deepcopy(original);a['presets'][0]['student_name']='MUST NOT SAVE';variants.append(a)
        a=deepcopy(original);a['presets'].append(deepcopy(a['presets'][0]));variants.append(a)
        for data in variants+[[],{},b'bad json']:
            with self.subTest(data=str(data)[:40]):
                self.put_catalog(data)
                with self.assertRaises(ReportingPresetError):self.service.load()
        self.assertEqual(self.repo.write_count,1)

    def test_tampered_ciphertext_rejected(self):
        self.create()
        raw=bytearray(self.repo.tree[self.service.path]);raw[80]=ord('A') if raw[80]!=ord('A') else ord('B')
        self.repo.tree[self.service.path]=bytes(raw);self.repo._commit()
        with self.assertRaisesRegex(ReportingPresetError,'decrypted'):self.service.load()

    def test_wrong_blob_sha_rejected(self):
        self.create()
        original=self.repo.request
        def wrong_sha(method,url,**kwargs):
            response=original(method,url,**kwargs)
            if response.status_code==200 and isinstance(response.data,dict) and response.data.get('type')=='file':
                response.data['sha']='0'*40
            return response
        self.archive.transport=type('T',(),{'request':staticmethod(wrong_sha)})()
        with self.assertRaisesRegex(ReportingPresetError,'identifier'):self.service.load()

    def test_symlink_or_bad_file_metadata_rejected(self):
        self.create()
        original=self.repo.request
        for changed in ({'type':'symlink'},{'size':0},{'size':MAX_ENCRYPTED_BYTES+1},
                        {'encoding':'none'},{'content':'@@@'},{'path':'other/file'}):
            def invalid(method,url,**kwargs):
                response=original(method,url,**kwargs)
                if isinstance(response.data,dict) and response.data.get('type')=='file':response.data.update(changed)
                return response
            self.archive.transport=type('T',(),{'request':staticmethod(invalid)})()
            with self.subTest(changed=changed),self.assertRaises(ReportingPresetError):self.service.load()

    def test_repository_scope_change_rejected(self):
        snapshot=self.service.load()
        other=OPDArchiveConfig(**{**self.secrets['opd_archive'],'repo':'a-different-repo'})
        service=GitHubReportingPresets(GitHubOPDArchive(other,transport=self.repo))
        with self.assertRaisesRegex(ReportingPresetError,'Refresh'):
            service.save('Example',self.period,snapshot)
        self.assertEqual(self.repo.write_count,0)

    def test_read_failure_is_not_empty_list(self):
        def offline(*args,**kwargs):raise requests.ConnectionError('private diagnostic')
        self.archive.transport=type('T',(),{'request':staticmethod(offline)})()
        with self.assertRaises(ReportingPresetError) as caught:self.service.load()
        self.assertNotIn('private diagnostic',str(caught.exception))
        self.assertNotIn(self.config.github_token,str(caught.exception))

    def test_names_unicode_spaces_and_invalid_controls(self):
        self.assertEqual(normalize_preset_name('  26-27   Review  '),'26-27 Review')
        self.assertEqual(preset_name_key('Cohort'),preset_name_key('cohort'))
        for name in ('', 'x'*81,'x\nname','x\x00name','\ud800',None):
            with self.subTest(name=repr(name)),self.assertRaises(ReportingPresetError):normalize_preset_name(name)
        result=self.create('Révision cohorte 26–27')
        self.assertEqual(result['snapshot']['presets'][0]['name'],'Révision cohorte 26–27')

    def test_preset_name_cannot_change_github_path(self):
        self.create('../../not-a-path')
        self.assertEqual(set(self.repo.tree),{self.service.path})

    def test_leap_day_long_period_and_equal_dates(self):
        for number,period in enumerate((ReportingPeriod('Leap',date(2028,2,29),date(2028,2,29)),
                                       ReportingPeriod('Long',date(2026,1,1),date(2028,4,1)))):
            receipt=self.create(str(number),period)
            self.assertIn(period,[period_from_preset(p) for p in receipt['snapshot']['presets']])

    def test_missing_target_and_duplicate_rename_rejected(self):
        first=self.create()
        second=self.create('Other',snapshot=first['snapshot'])
        with self.assertRaises(ReportingPresetError):
            self.service.save('Other',self.period,second['snapshot'],replace_id=first['preset_id'])
        with self.assertRaises(ReportingPresetError):self.service.delete('f'*32,second['snapshot'])
        self.assertEqual(len(self.service.load()['presets']),2)

    def test_no_credentials_or_opd_content_in_catalog(self):
        self.create()
        content=json.dumps(self.raw_catalog())
        self.assertNotIn(self.config.github_token,content)
        self.assertNotIn(self.config.encryption_key,content)
        self.assertEqual(set(self.raw_catalog()['presets'][0]),{'id','name','period','created_at','updated_at'})
        for method,url,body,_ in self.repo.calls:
            if method=='PUT':
                self.assertEqual(body['message'],'Update encrypted reporting-date presets')
                self.assertNotIn('Teaching year',json.dumps(body))


class PresetUITests(PresetFixture):
    def test_first_visit_loads_only_small_catalog_and_never_saves(self):
        with self.request_patch():ui.render_period_controls()
        self.assertEqual(self.repo.write_count,0)
        self.assertEqual(len(self.repo.calls),2)
        self.assertIsNone(st.session_state['teaching_period_start'])
        self.assertEqual(st.session_state[ui.SNAPSHOT]['presets'],[])
        self.assertFalse(any('OPD_' in url for _,url,_,_ in self.repo.calls))

    def test_dropdown_in_new_session_shows_saved_preset(self):
        receipt=self.create()
        st.reset(secrets=self.secrets)
        login_for_test()
        with self.request_patch():ui.render_period_controls()
        self.assertEqual(st.session_state[ui.SNAPSHOT]['presets'],receipt['snapshot']['presets'])
        self.assertIn(('Saved date presets',ui.CHOICE),st.widget_keys)

    def test_load_callback_sets_dates_label_mode_and_clears_only_outputs(self):
        receipt=self.seed_state()
        st.session_state['teaching_zip']=b'old zip'
        st.session_state['teaching_scan']={'sentinel':'unchanged'}
        st.session_state['teaching_reporting_mode']='Standard July-June academic years'
        with self.request_patch():ui.load_selected_preset(receipt['preset_id'])
        self.assertEqual(st.session_state['teaching_period_start'],self.period.start)
        self.assertEqual(st.session_state['teaching_period_end'],self.period.end)
        self.assertEqual(st.session_state['teaching_period_label'],self.period.label)
        self.assertEqual(st.session_state['teaching_reporting_mode'],'Custom dates')
        self.assertNotIn('teaching_zip',st.session_state)
        self.assertEqual(st.session_state['teaching_scan'],{'sentinel':'unchanged'})

    def test_explicit_save_callback_persists_dates(self):
        with self.request_patch():ui.render_period_controls()
        st.session_state.update({ui.NAME:'New','teaching_period_label':'26-27',
                                'teaching_period_start':self.period.start,'teaching_period_end':self.period.end})
        with self.request_patch():ui.save_current_preset()
        self.assertEqual(self.repo.write_count,1)
        self.assertEqual(self.service.load()['presets'][0]['name'],'New')
        self.assertNotIn(ui.ERROR,st.session_state)

    def test_replacement_requires_confirmation(self):
        receipt=self.seed_state()
        with self.request_patch():ui.save_current_preset(receipt['preset_id'],'not-confirmed')
        self.assertEqual(self.repo.write_count,1)
        self.assertIn('Confirm replacement',st.session_state[ui.ERROR])

    def test_valid_replace_callback_updates_dates(self):
        receipt=self.seed_state()
        changed=ReportingPeriod('Revised',self.period.start,date(2027,4,30))
        st.session_state['teaching_period_label']=changed.label
        st.session_state['teaching_period_end']=changed.end
        key=ui._confirmation_key('replace',receipt['snapshot'],receipt['preset_id'],'Teaching year',changed.as_dict())
        st.session_state[key]=True
        with self.request_patch():ui.save_current_preset(receipt['preset_id'],key)
        self.assertEqual(period_from_preset(self.service.load()['presets'][0]),changed)

    def test_old_confirmation_cannot_authorize_changed_dates(self):
        receipt=self.seed_state()
        key=ui._confirmation_key('replace',receipt['snapshot'],receipt['preset_id'],'Teaching year',self.period.as_dict())
        st.session_state[key]=True
        st.session_state['teaching_period_end']=date(2028,1,1)
        with self.request_patch():ui.save_current_preset(receipt['preset_id'],key)
        self.assertEqual(self.repo.write_count,1)

    def test_delete_callback_requires_confirmation(self):
        receipt=self.seed_state()
        key=ui._confirmation_key('delete',receipt['snapshot'],receipt['preset_id'])
        with self.request_patch():ui.delete_selected_preset(receipt['preset_id'],key)
        self.assertEqual(len(self.service.load()['presets']),1)

    def test_confirmed_delete_keeps_current_date_inputs(self):
        receipt=self.seed_state()
        key=ui._confirmation_key('delete',receipt['snapshot'],receipt['preset_id'])
        st.session_state[key]=True
        with self.request_patch():ui.delete_selected_preset(receipt['preset_id'],key)
        self.assertEqual(self.service.load()['presets'],[])
        self.assertEqual(st.session_state['teaching_period_start'],self.period.start)
        self.assertEqual(st.session_state['teaching_period_label'],'26-27')
        self.assertIsNone(st.session_state[ui.CHOICE])

    def test_confirmation_bound_to_selected_preset(self):
        first=self.create();second=self.create('Other',snapshot=first['snapshot'])
        self.seed_state(second)
        key=ui._confirmation_key('delete',second['snapshot'],first['preset_id'])
        st.session_state[key]=True
        with self.request_patch():ui.delete_selected_preset(second['preset_id'],key)
        self.assertEqual(len(self.service.load()['presets']),2)

    def test_manual_dates_still_work_when_github_unavailable(self):
        st.session_state.update({'teaching_period_label':'Manual','teaching_period_start':self.period.start,
                                'teaching_period_end':self.period.end})
        with patch('schedule_app.services.opd_archive.requests.request',side_effect=requests.ConnectionError()):
            mode,period,issue=ui.render_period_controls()
        self.assertEqual(period.label,'Manual');self.assertIsNone(issue)
        self.assertIn(ui.ERROR,st.session_state)

    def test_read_error_requires_refresh_instead_of_retrying_every_widget(self):
        with patch('schedule_app.services.opd_archive.requests.request',side_effect=requests.ConnectionError()) as call:
            ui.render_period_controls();n=call.call_count
            ui.render_period_controls();self.assertEqual(call.call_count,n)
        with self.request_patch():ui.refresh_saved_presets()
        self.assertNotIn(ui.ERROR,st.session_state)
        self.assertIn(ui.SNAPSHOT,st.session_state)

    def test_rerenders_do_not_write_or_reread_catalog(self):
        with self.request_patch():
            ui.render_period_controls();n=len(self.repo.calls)
            ui.render_period_controls()
        self.assertEqual(len(self.repo.calls),n);self.assertEqual(self.repo.write_count,0)

    def test_no_json_upload_download_controls(self):
        with self.request_patch():ui.render_period_controls()
        self.assertNotIn('teaching_period_json_upload',[key for _,key in st.widget_keys])
        self.assertNotIn('teaching_save_period',[key for _,key in st.widget_keys])
        self.assertFalse(any(name.endswith('.json') for name in st.downloads))

    def test_loading_deleted_preset_does_not_restore_old_dates(self):
        receipt=self.seed_state()
        self.service.delete(receipt['preset_id'],receipt['snapshot'])
        st.session_state['teaching_period_label']='Leave alone'
        with self.request_patch():ui.load_selected_preset(receipt['preset_id'])
        self.assertEqual(st.session_state['teaching_period_label'],'Leave alone')
        self.assertIn('deleted',st.session_state[ui.ERROR])

    def test_remote_changed_preset_needs_review_before_load(self):
        receipt=self.seed_state()
        changed=ReportingPeriod('Changed remotely',self.period.start,date(2027,4,1))
        self.service.save('Teaching year',changed,receipt['snapshot'],replace_id=receipt['preset_id'])
        with self.request_patch():ui.load_selected_preset(receipt['preset_id'])
        self.assertEqual(st.session_state['teaching_period_label'],'26-27')
        self.assertIn('changed in GitHub',st.session_state[ui.NOTICE])
        with self.request_patch():ui.load_selected_preset(receipt['preset_id'])
        self.assertEqual(st.session_state['teaching_period_label'],'Changed remotely')

    def test_archive_config_switch_does_not_leak_old_presets(self):
        self.seed_state()
        other=secret_settings()
        other['opd_archive']['repo']='other-repo'
        st.secrets=other
        login_for_test()
        other_repo=FakeGitHub()
        with patch('schedule_app.services.opd_archive.requests.request',other_repo.request):
            ui.render_period_controls()
        self.assertEqual(st.session_state[ui.SNAPSHOT]['presets'],[])
        self.assertIsNone(st.session_state[ui.CHOICE])
        self.assertEqual(st.session_state[ui.NAME],'')

    def test_names_and_confirmations_are_only_presets_not_opd_permissions(self):
        source=(Path(__file__).parents[1]/'schedule_app/sections/reporting_date_controls.py').read_text()
        self.assertNotIn('app_password',source)
        self.assertNotIn('file_uploader',source)
        self.assertIn('on_click=delete_selected_preset',source)

    def test_refresh_does_not_change_active_dates(self):
        self.seed_state()
        with self.request_patch():ui.refresh_saved_presets()
        self.assertEqual(st.session_state['teaching_period_start'],self.period.start)
        self.assertEqual(st.session_state['teaching_period_end'],self.period.end)


if __name__=='__main__':unittest.main()
