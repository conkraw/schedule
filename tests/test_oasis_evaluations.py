"""Offline regression tests. Synthetic evaluations only; never contact GitHub."""
import base64
from dataclasses import replace
from datetime import date
import hashlib
import io
import csv
from importlib import import_module
from pathlib import Path
import unittest
from unittest.mock import patch
from urllib.parse import unquote

from schedule_app.services.oasis_privacy import minimize_oasis_csv

def minimal(raw):
    return minimize_oasis_csv(raw, "educator")['raw']

from cryptography.fernet import Fernet
from helpers import FakeGitHub, FakeResponse, Upload, st, secret_settings, run_app, ROOT
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive, OPDArchiveError
from schedule_app.services.oasis_evaluations import (
    GitHubOASISEvaluations, OASISArchiveError, OASIS_MAX_BYTES, OASIS_MAX_ENCRYPTED_BYTES,
    OASIS_MAX_FIELD_CHARS, inspect_oasis_csv, oasis_export_label,
)
from schedule_app.sections import oasis_evaluation_archive as ui

HEADERS = ['Course ID', 'Start Date', 'End Date', 'Student', 'Evaluator', 'Evaluation',
           'Form Record', 'Question ID', 'Question', 'Answer text', 'Submit Date']

def make_csv(*, answer='Useful feedback', start='2026-03-16', end='2026-04-10',
             rows=1, encoding='utf-8', newline='\r\n'):
    out=io.StringIO(newline='')
    writer=csv.writer(out,lineterminator=newline)
    writer.writerow(HEADERS)
    for i in range(rows):
        writer.writerow(['DEMO-101',start,end,'Example Learner','Example Teacher','Teaching Evaluation',
                         str(100+i),'99','Example question',answer,'2026-04-11 10:00:00'])
    return out.getvalue().encode(encoding)

class RecordingGitHub(FakeGitHub):
    def __init__(self):
        super().__init__()
        self.calls=[]
        self.raw_fallback=False
        self.put_status=None
        self.break_verification=False
    def request(self,method,url,**kwargs):
        self.calls.append((method,url,kwargs))
        if method=='PUT' and self.put_status:
            return FakeResponse(self.put_status)
        response=super().request(method,url,**kwargs)
        if method=='PUT' and self.break_verification:
            path=unquote(url.split('/contents/',1)[1])
            self.tree[path]=b'not-valid-ciphertext'
            self._commit()
            return FakeResponse(201,{'commit':{'sha':self.head}})
        if self.raw_fallback and method=='GET' and isinstance(response.data,dict) and response.data.get('type')=='file':
            response.data['encoding']='none';response.data['content']=''
        return response

class StorageFixture(unittest.TestCase):
    def setUp(self):
        self.secrets=secret_settings()
        st.reset(secrets=self.secrets)
        self.config=OPDArchiveConfig(**self.secrets['opd_archive'])
        self.repo=RecordingGitHub()
        self.archive=GitHubOPDArchive(self.config,transport=self.repo)
        self.service=GitHubOASISEvaluations(self.archive)
        self.raw=make_csv()

class CSVValidationTests(unittest.TestCase):
    def test_metadata_is_aggregate_only(self):
        result=inspect_oasis_csv(make_csv(rows=4))
        self.assertEqual(result['row_count'],4)
        self.assertEqual(result['column_count'],len(HEADERS))
        self.assertEqual(result['coverage'],'2026-03-16_to_2026-04-10')
        self.assertNotIn('Example Teacher',repr(result))
        self.assertNotIn('Example Learner',repr(result))
    def test_quotes_commas_and_multiline_comments(self):
        self.assertEqual(inspect_oasis_csv(make_csv(answer='First, "quoted" line\r\nsecond line'))['row_count'],1)
    def test_utf8_bom(self):
        self.assertEqual(inspect_oasis_csv(make_csv(encoding='utf-8-sig'))['encoding'],'UTF-8')
    def test_utf16_with_bom(self):
        self.assertEqual(inspect_oasis_csv(make_csv(encoding='utf-16'))['encoding'],'UTF-16')
    def test_windows1252(self):
        self.assertEqual(inspect_oasis_csv(make_csv(answer='Caf\u00e9 — excellent',encoding='cp1252'))['encoding'],'Windows-1252')
    def test_unrelated_csv_rejected(self):
        with self.assertRaisesRegex(OASISArchiveError,'Missing columns'):
            inspect_oasis_csv(b'name,date\nSomeone,2026-01-01\n')
    def test_empty_rejected(self):
        for raw in (b'',None,'not bytes'):
            with self.subTest(raw=raw),self.assertRaises(OASISArchiveError):
                inspect_oasis_csv(raw)
    def test_header_only_rejected(self):
        with self.assertRaisesRegex(OASISArchiveError,'no response rows'):
            inspect_oasis_csv(make_csv(rows=0))
    def test_binary_and_xlsx_rejected(self):
        for raw in (b'PK\x03\x04bad',b'%PDF bad',b'abc\x00def'):
            with self.subTest(raw=raw),self.assertRaises(OASISArchiveError):
                inspect_oasis_csv(raw)
    def test_bad_row_width_rejected_without_values(self):
        raw=make_csv()+b'PRIVATE VALUE,wrong\r\n'
        with self.assertRaisesRegex(OASISArchiveError,'has 2 fields') as caught:
            inspect_oasis_csv(raw)
        self.assertNotIn('PRIVATE VALUE',str(caught.exception))
    def test_duplicate_headers_rejected(self):
        raw=make_csv().replace(b'Submit Date',b'Question ID')
        with self.assertRaisesRegex(OASISArchiveError,'duplicate'):
            inspect_oasis_csv(raw)
    def test_blank_headers_rejected(self):
        raw=make_csv().replace(b'Submit Date',b'')
        with self.assertRaisesRegex(OASISArchiveError,'blank'):
            inspect_oasis_csv(raw)
    def test_malformed_quotes_rejected(self):
        with self.assertRaises(OASISArchiveError):
            inspect_oasis_csv(make_csv()+b'"unterminated')
    def test_date_formats(self):
        result=inspect_oasis_csv(make_csv(start='03/16/2026',end='04/10/2026'))
        self.assertEqual(result['coverage'],'2026-03-16_to_2026-04-10')
    def test_date_problems_preserve_original_as_undated(self):
        for start,end in (('',''),('not a date','2026-04-10'),('2026-04-11','2026-04-10')):
            with self.subTest(start=start,end=end):
                result=inspect_oasis_csv(make_csv(start=start,end=end))
                self.assertFalse(result['date_range_complete'])
                self.assertEqual(result['coverage'],'undated')
    def test_course_dates_not_submission_dates(self):
        self.assertEqual(inspect_oasis_csv(make_csv())['course_end'],'2026-04-10')
    def test_large_comment_allowed_and_csv_limit_restored(self):
        before=csv.field_size_limit()
        self.assertEqual(inspect_oasis_csv(make_csv(answer='x'*160000))['row_count'],1)
        self.assertEqual(csv.field_size_limit(),before)
    def test_oversized_field_rejected_and_csv_limit_restored(self):
        before=csv.field_size_limit()
        with self.assertRaises(OASISArchiveError):
            inspect_oasis_csv(make_csv(answer='x'*(OASIS_MAX_FIELD_CHARS+1)))
        self.assertEqual(csv.field_size_limit(),before)
    def test_size_limit(self):
        with self.assertRaisesRegex(OASISArchiveError,'10 MiB'):
            inspect_oasis_csv(b'x'*(OASIS_MAX_BYTES+1))
    def test_literal_formula_preserved_not_evaluated(self):
        result=inspect_oasis_csv(make_csv(answer='=SUM(1,2)'))
        self.assertEqual(result['row_count'],1)

class ArchiveStorageTests(StorageFixture):
    def test_empty_list_does_not_write(self):
        self.assertEqual(self.service.list_exports()['filenames'],[])
        self.assertEqual(self.repo.write_count,0)
    def test_exact_original_bytes_roundtrip(self):
        saved=self.service.save(self.raw)
        loaded=self.service.load(saved['filename'])
        self.assertEqual(loaded['raw'],minimal(self.raw))
        self.assertEqual(saved['action'],'created')
        self.assertEqual(self.config.cipher().decrypt(self.repo.tree[saved['path']]),minimal(self.raw))
    def test_no_plaintext_or_identity_in_public_filename_or_ciphertext(self):
        saved=self.service.save(self.raw)
        public=saved['path'].encode()+self.repo.tree[saved['path']]
        for value in (b'Example Teacher',b'Example Learner',b'Useful feedback',hashlib.sha256(self.raw).hexdigest().encode()):
            self.assertNotIn(value,public)
    def test_destination_same_repository_separate_folder(self):
        saved=self.service.save(self.raw)
        self.assertTrue(saved['path'].startswith('opd_archive/oasis_evaluations/'))
        self.assertTrue(all('/example-account/opd-test-archive/' in url for _,url,_ in self.repo.calls))
    def test_same_export_is_not_rewritten(self):
        first=self.service.save(self.raw); second=self.service.save(self.raw)
        self.assertEqual(second['filename'],first['filename'])
        self.assertEqual(second['action'],'unchanged')
        self.assertEqual(self.repo.write_count,1)
    def test_identical_file_in_other_session_is_not_duplicated(self):
        first=self.service.save(self.raw)
        other=GitHubOASISEvaluations(GitHubOPDArchive(self.config,transport=self.repo))
        self.assertEqual(other.save(self.raw)['action'],'unchanged')
        self.assertEqual(other.load(first['filename'])['raw'],minimal(self.raw))
    def test_different_content_same_dates_retained_as_separate_export(self):
        first=self.service.save(self.raw); changed=make_csv(answer='Revised response')
        second=self.service.save(changed)
        self.assertNotEqual(first['filename'],second['filename'])
        self.assertEqual(len(self.service.list_exports()['filenames']),2)
        self.assertEqual(self.service.load(first['filename'])['raw'],minimal(self.raw))
    def test_different_dates_retained_and_dropdown_ordered_by_coverage(self):
        first=self.service.save(self.raw)
        second=self.service.save(make_csv(start='2026-08-03',end='2026-08-28'))
        self.assertEqual(self.service.list_exports()['filenames'],[second['filename'],first['filename']])
    def test_raw_download_fallback_for_large_exports(self):
        self.repo.raw_fallback=True
        saved=self.service.save(self.raw)
        self.assertEqual(self.service.load(saved['filename'])['raw'],minimal(self.raw))
        self.assertTrue(any(kw.get('headers',{}).get('Accept')=='application/vnd.github.raw+json'
                            for method,url,kw in self.repo.calls))
    def test_old_key_recovery_and_duplicate_detection(self):
        first=self.service.save(self.raw)
        config=replace(self.config,encryption_key=Fernet.generate_key().decode(),
                       previous_encryption_keys=(self.config.encryption_key,))
        new=GitHubOASISEvaluations(GitHubOPDArchive(config,transport=self.repo))
        self.assertEqual(new.load(first['filename'])['raw'],minimal(self.raw))
        self.assertEqual(new.save(self.raw)['action'],'unchanged')
    def test_wrong_key_refuses_to_decrypt(self):
        saved=self.service.save(self.raw)
        wrong=replace(self.config,encryption_key=Fernet.generate_key().decode())
        with self.assertRaisesRegex(OASISArchiveError,'cannot be decrypted'):
            GitHubOASISEvaluations(GitHubOPDArchive(wrong,transport=self.repo)).load(saved['filename'])
    def test_corrupted_ciphertext_not_overwritten(self):
        saved=self.service.save(self.raw)
        self.repo.tree[saved['path']]=b'altered'; self.repo._commit()
        with self.assertRaises(OASISArchiveError):
            self.service.save(self.raw)
        self.assertEqual(self.repo.write_count,1)
    def test_renamed_ciphertext_rejected(self):
        saved=self.service.save(self.raw)
        altered=saved['filename'].replace('_to_2026-04-10','_to_2026-04-11')
        self.repo.tree[self.service.path_for(altered)]=self.repo.tree[saved['path']]; self.repo._commit()
        with self.assertRaisesRegex(OASISArchiveError,'does not match'):
            self.service.load(altered)
    def test_invalid_filename_cannot_access_opd_or_presets(self):
        for name in ('../OPD_2026-03-16.xlsx.enc','../../secrets','reporting_date_presets.json.enc','OASIS_bad.csv.enc'):
            with self.subTest(name=name),self.assertRaises(OASISArchiveError):
                self.service.load(name)
        self.assertEqual(self.repo.calls,[])
    def test_missing_export_reports_error(self):
        name=self.service._candidate_names(self.raw,inspect_oasis_csv(self.raw))[0]
        with self.assertRaisesRegex(OASISArchiveError,'not found'):
            self.service.load(name)
    def test_invalid_csv_does_not_contact_github(self):
        with self.assertRaises(OASISArchiveError):self.service.save(b'bad\n')
        self.assertEqual(self.repo.calls,[])
    def test_failed_save_is_not_success(self):
        self.repo.put_status=403
        with self.assertRaisesRegex(OASISArchiveError,'not confirmed'):
            self.service.save(self.raw)
        self.assertEqual(self.repo.write_count,0)
    def test_concurrent_creation_stops_and_retry_verifies(self):
        self.repo.put_status=409
        with self.assertRaises(OASISArchiveError):self.service.save(self.raw)
        self.repo.put_status=None
        self.service.save(self.raw)
        self.assertEqual(self.service.save(self.raw)['action'],'unchanged')
    def test_post_save_verification_must_pass(self):
        self.repo.break_verification=True
        with self.assertRaises(OASISArchiveError):self.service.save(self.raw)
    def test_verification_response_missing_commit_then_retry_no_duplicate(self):
        original=self.service._call
        def call(method,route,**kwargs):
            result=original(method,route,**kwargs)
            return {} if method=='PUT' else result
        with patch.object(self.service,'_call',side_effect=call):
            with self.assertRaisesRegex(OASISArchiveError,'did not confirm'):
                self.service.save(self.raw)
        self.assertEqual(self.service.save(self.raw)['action'],'unchanged')
        self.assertEqual(self.repo.write_count,1)
    def test_receipt_contains_no_csv_or_plaintext_hash(self):
        receipt=self.service.save(self.raw)
        self.assertNotIn('raw',receipt)
        self.assertNotIn('sha256',receipt)
        self.assertNotIn('Example Teacher',repr(receipt))
    def test_public_listing_has_no_data_or_original_filename(self):
        self.service.save(self.raw)
        listing=self.service.list_exports()
        self.assertNotIn('Example Teacher',repr(listing))
        self.assertNotIn('Question',repr(listing))
    def test_load_only_never_writes(self):
        saved=self.service.save(self.raw); before=self.repo.write_count
        self.service.list_exports();self.service.load(saved['filename'])
        self.assertEqual(self.repo.write_count,before)
    def test_opd_and_preset_bytes_untouched_and_oasis_not_an_opd(self):
        path='opd_archive/OPD_2026-09-07.xlsx.enc'
        preset='opd_archive/reporting_date_presets.json.enc'
        self.repo.tree.update({path:b'KEEP_OPD',preset:b'KEEP_PRESET'});self.repo._commit()
        self.service.save(self.raw)
        self.assertEqual(self.repo.tree[path],b'KEEP_OPD');self.assertEqual(self.repo.tree[preset],b'KEEP_PRESET')
        self.assertEqual(self.archive.list_rotations(),[date(2026,9,7)])
        self.assertTrue(all('/oasis_evaluations/' in url for method,url,_ in self.repo.calls if method=='PUT'))
    def test_unknown_files_ignored(self):
        self.repo.tree[self.service.folder+'/README.md']=b'not an export';self.repo._commit()
        self.assertEqual(self.service.list_exports()['filenames'],[])
    def test_1000_entry_limit_reported_not_partial(self):
        response={'entries':[{'type':'file','name':'other'}]*1000}
        with patch.object(self.service,'_call',return_value=response):
            with self.assertRaisesRegex(OASISArchiveError,'complete list cannot'):
                self.service.list_exports()
    def test_non_file_metadata_rejected(self):
        saved=self.service.save(self.raw)
        original=self.service._call
        def altered(method,route,**kwargs):
            result=original(method,route,**kwargs)
            if isinstance(result,dict):result['type']='symlink'
            return result
        with patch.object(self.service,'_call',side_effect=altered):
            with self.assertRaisesRegex(OASISArchiveError,'regular file'):
                self.service.load(saved['filename'])
    def test_invalid_date_original_is_still_encrypted_recoverable(self):
        raw=make_csv(start='unknown'); saved=self.service.save(raw)
        self.assertIn('_undated_',saved['filename'])
        self.assertEqual(self.service.load(saved['filename'])['raw'],minimal(raw))
    def test_label_shows_dates_and_identifier(self):
        saved=self.service.save(self.raw)
        label=oasis_export_label(saved['filename'])
        self.assertIn('2026-03-16 to 2026-04-10',label)
        self.assertIn('export',label)
    def test_read_pinned_to_requested_commit(self):
        saved=self.service.save(self.raw); head=saved['commit']
        self.repo.tree[saved['path']]=b'changed outside app';self.repo._commit()
        self.assertEqual(self.service.load(saved['filename'],commit=head)['raw'],minimal(self.raw))
        with self.assertRaises(OASISArchiveError):self.service.load(saved['filename'])

class InterfaceTests(StorageFixture):
    def page(self,values=None,state=None):
        return run_app({'schedule_app_mode':'OASIS Evaluation Archive',**(values or {})},
                       secrets=self.secrets,state=state,repo=self.repo, evaluation_login=True, original=ROOT/"tests"/"legacy_oasis_entrypoint.py")
    def test_new_menu_option_loads_without_upload(self):
        result=self.page()
        self.assertIn(('subheader','OASIS Evaluation Archive'),result['messages'])
        self.assertEqual(self.repo.write_count,0)
    def test_auto_save_upload_and_unchanged_rerun(self):
        upload=Upload(self.raw,'oasis_eval_export (1).csv')
        first=self.page({ui.UPLOAD:upload})
        self.assertIn('receipt',first['state'][ui.SAVE])
        second=self.page({ui.UPLOAD:upload},state=first['state'])
        self.assertEqual(self.repo.write_count,1)
        self.assertIn('receipt',second['state'][ui.SAVE])
    def test_changed_contents_same_local_filename_saved_separately(self):
        first=self.page({ui.UPLOAD:Upload(self.raw,'oasis.csv')})
        second=self.page({ui.UPLOAD:Upload(make_csv(answer='updated'),'oasis.csv')},state=first['state'])
        self.assertEqual(self.repo.write_count,2)
        self.assertNotEqual(first['state'][ui.SAVE]['receipt']['filename'],second['state'][ui.SAVE]['receipt']['filename'])
    def test_renamed_identical_file_in_fresh_session_no_duplicate(self):
        self.page({ui.UPLOAD:Upload(self.raw,'original.csv')})
        second=self.page({ui.UPLOAD:Upload(self.raw,'renamed.csv')})
        self.assertEqual(second['state'][ui.SAVE]['receipt']['action'],'unchanged')
        self.assertEqual(self.repo.write_count,1)
    def test_separate_session_reload_download_matches_original(self):
        first=self.service.save(self.raw)
        result=self.page({ui.CHOICE:first['filename'],'oasis_load_export':True})
        self.assertEqual(result['downloads'][first['filename'][:-4]],minimal(self.raw))
        self.assertEqual(self.repo.write_count,1)
    def test_save_failure_no_success_or_download(self):
        self.repo.put_status=403
        result=self.page({ui.UPLOAD:Upload(self.raw,'oasis.csv')})
        self.assertIn('error',result['state'][ui.SAVE])
        self.assertEqual(result['downloads'],{})
        self.assertFalse(any(kind=='success' for kind,text in result['messages']))
    def test_callback_allows_explicit_retry(self):
        self.repo.put_status=403
        upload=Upload(self.raw,'oasis.csv');first=self.page({ui.UPLOAD:upload})
        ui._upload_changed(); state=dict(st.session_state)
        self.repo.put_status=None
        result=self.page({ui.UPLOAD:upload},state=state)
        self.assertIn('receipt',result['state'][ui.SAVE])
    def test_error_not_retried_automatically(self):
        self.repo.put_status=403
        upload=Upload(self.raw,'oasis.csv');first=self.page({ui.UPLOAD:upload})
        calls=len(self.repo.calls)
        self.page({ui.UPLOAD:upload},state=first['state'])
        self.assertEqual(len(self.repo.calls),calls)
    def test_no_evaluation_details_echoed_on_page(self):
        result=self.page({ui.UPLOAD:Upload(self.raw,'oasis.csv')})
        all_messages=repr(result['messages'])
        self.assertNotIn('Example Learner',all_messages)
        self.assertNotIn('Example Teacher',all_messages)
        self.assertNotIn('Useful feedback',all_messages)
    def test_switch_selection_clears_old_download(self):
        first=self.service.save(self.raw)
        second=self.service.save(make_csv(answer='another'))
        loaded=self.page({ui.CHOICE:first['filename'],'oasis_load_export':True})
        changed=self.page({ui.CHOICE:second['filename']},state=loaded['state'])
        self.assertEqual(changed['downloads'],{})
    def test_scope_change_drops_oasis_data_not_teaching_reports(self):
        first=self.service.save(self.raw)
        loaded=self.page({ui.CHOICE:first['filename'],'oasis_load_export':True})
        state=loaded['state'];state['teaching_zip']=b'UNCHANGED';state[ui.SCOPE]='different'
        result=self.page(state=state)
        self.assertNotIn(ui.LOADED,result['state'])
        self.assertEqual(result['state']['teaching_zip'],b'UNCHANGED')
    def test_import_has_no_side_effects(self):
        st.reset()
        module=import_module('schedule_app.sections.oasis_evaluation_archive')
        self.assertTrue(callable(module.render));self.assertEqual(st.events,[])
    def test_only_combined_workflow_runtime_files_differ_from_baseline(self):
        # The small update leaves all settings, OPD, learner reach and reports alone.
        base=ROOT.parent/'base'
        if not base.exists():self.skipTest('Baseline comparison is a packaging check')
        if (base/'schedule_app/services/oasis_workflow.py').exists():
            self.skipTest('Historical OASIS-only packaging check: this baseline already contains the combined workflow')
        changed={p.relative_to(ROOT).as_posix() for p in [ROOT/'app_sch_2026.py',*(ROOT/'schedule_app').rglob('*.py')]
                 if not (base/p.relative_to(ROOT)).exists() or p.read_bytes()!=(base/p.relative_to(ROOT)).read_bytes()}
        self.assertEqual(changed,{'app_sch_2026.py','schedule_app/services/oasis_evaluations.py',
                                 'schedule_app/services/oasis_workflow.py',
                                 'schedule_app/sections/oasis_workflow.py',
                                 'schedule_app/sections/oasis_date_controls.py'})

if __name__=='__main__':unittest.main()
