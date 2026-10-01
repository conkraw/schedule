"""Synthetic-only report tests; no real student/educator information or network."""
import base64
from dataclasses import replace
from datetime import date
import csv
import io
import json
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from cryptography.fernet import Fernet
from helpers import st, FakeGitHub, FakeResponse, Upload, secret_settings, run_app, ROOT
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive, OPDArchiveError
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations
from schedule_app.services.oasis_educator_reports import (
    QUESTION_TEXTS, COMMENT_TEXTS, OASISReportError, prepare_reports, educator_summary,
    csv_bytes, build_report_downloads, validate_username, email_username, question_key_rows,
)
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames

FIELDS = ['Course ID','Start Date','End Date','Student','Student Email','Who Completed',
          'Evaluator','Evaluator Email','Evaluator Username','Evaluator External ID',
          'Evaluation','Form Record','Question Number','Question ID','Question','Answer text',
          'Multiple Choice Order','Multiple Choice Value','Multiple Choice Label','Submit Date']


def row(form='1', qid='587', value='5', label='Strongly Agree', answer='', **changes):
    data = dict(zip(FIELDS, ['DEMO-101','2026-03-16','2026-04-10','SYNTHETIC LEARNER',
            'learner@example.edu','SYNTHETIC COMPLETER','Example, Avery','abc12@example.edu','other_username','staff-001',
            '*Clinical Teaching Eval',form,'1',qid,(QUESTION_TEXTS | COMMENT_TEXTS).get(qid,'Additional question'),answer,
            '999',value,label,'2026-04-11 10:00:00']))
    data.update(changes)
    return data


def comments(form='1', **changes):
    return [row(form,'172','','','Clear explanations.',**changes), row(form,'437','','','More feedback.',**changes)]


def make_csv(rows=None, *, encoding='utf-8', omit=()):
    output=io.StringIO(newline='')
    fields=[f for f in FIELDS if f not in omit]
    writer=csv.DictWriter(output,fields,extrasaction='ignore',lineterminator='\r\n')
    writer.writeheader();writer.writerows(rows or [row()]+comments())
    return output.getvalue().encode(encoding)


def prepared(rows=None):
    return prepare_reports([('test.csv',make_csv(rows))])


class SummaryTests(unittest.TestCase):
    def test_one_form_not_three_rows(self):
        s=educator_summary(prepared());self.assertEqual(s['rows'][0]['evaluation_count'],1)
        self.assertEqual(s['rows'][0]['record_id'],'abc12')
    def test_username_from_email_not_username_column(self):
        self.assertEqual(educator_summary(prepared())['rows'][0]['record_id'],'abc12')
    def test_means_use_value_not_order(self):
        r=educator_summary(prepared([row('1',value='5'),row('2',value='3')]))['rows'][0]
        self.assertEqual(r['q587_mean'],'4.00');self.assertEqual(r['q587_n'],2)
    def test_question_number_changes_do_not_split(self):
        r=educator_summary(prepared([row('1','589',**{'Question Number':'3'}),row('2','589',value='3',**{'Question Number':'2'})]))['rows'][0]
        self.assertEqual(r['q589_mean'],'4.00');self.assertEqual(r['q589_n'],2)
    def test_na_value_blank_ignored(self):
        r=educator_summary(prepared([row('1'),row('2',value='',label='N/A')]))['rows'][0]
        self.assertEqual((r['q587_mean'],r['q587_n'],r['evaluation_count']),('5.00',1,2))
    def test_na_numeric_code_ignored(self):
        r=educator_summary(prepared([row('1'),row('2',value='0',label='N/A')]))['rows'][0]
        self.assertEqual((r['q587_mean'],r['q587_n']),('5.00',1))
    def test_na_in_value_also_excluded(self):
        r=educator_summary(prepared([row(value="N/A",label="")]))["rows"][0]
        self.assertEqual((r["q587_mean"],r["q587_n"]),("",0))
    def test_negative_numeric_score_remains_numeric_csv(self):
        p=prepared([row(value="-1",label="Example code")]);d=build_report_downloads(p,educator_summary(p))
        r=list(csv.DictReader(io.StringIO(d["csv"].decode("utf-8-sig"))))[0]
        self.assertEqual(r["q587_mean"],"-1.00")
    def test_all_blank_not_zero(self):
        r=educator_summary(prepared([row(value='',label='N/A')]))['rows'][0]
        self.assertEqual((r['q587_mean'],r['q587_n']),('',0))
    def test_omitted_question_n_zero(self):
        r=educator_summary(prepared())['rows'][0]
        self.assertEqual((r['q588_mean'],r['q588_n']),('',0))
    def test_duplicate_exports_count_once(self):
        raw=make_csv();p=prepare_reports([('one',raw),('two',raw)])
        r=educator_summary(p)['rows'][0]
        self.assertEqual((r['evaluation_count'],p['duplicates_removed']),(1,3))
        self.assertEqual(r['strengths_comments'],'1. Clear explanations.')
    def test_overlapping_exports_not_double_counted(self):
        p=prepare_reports([('a',make_csv([row('1'),row('2')])),('b',make_csv([row('2'),row('3')]))])
        self.assertEqual(educator_summary(p)['evaluation_count'],3)
    def test_conflicting_score_blocks(self):
        with self.assertRaises(OASISReportError) as err:
            prepared([row('1',value='5'),row('1',value='4')])
        self.assertTrue(err.exception.issues);self.assertNotIn('SYNTHETIC LEARNER',str(err.exception.issues))
    def test_conflicting_comment_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row('1','172','','','A'),row('1','172','','','B')])
    def test_conflicting_form_date_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row('1'),row('1','589',**{'Submit Date':'2026-05-01 10:00:00'})])
    def test_same_comment_distinct_forms_retained(self):
        r=educator_summary(prepared(comments('1')+comments('2')))['rows'][0]
        self.assertEqual(r['strengths_comments'],'1. Clear explanations.\n\n2. Clear explanations.')
    def test_literal_na_comments_retained(self):
        r=educator_summary(prepared([row('1','437','','','n/a')]))['rows'][0]
        self.assertEqual(r['areas_for_improvement_comments'],'1. n/a')
    def test_comments_multiline_quotes_roundtrip(self):
        text='First, "quoted" line\nsecond line'
        p=prepared([row('1','172','','',text)]);s=educator_summary(p)
        decoded=list(csv.DictReader(io.StringIO(build_report_downloads(p,s)['csv'].decode('utf-8-sig'))))
        self.assertEqual(decoded[0]['strengths_comments'],'1. '+text)
    def test_comment_order_is_chronological(self):
        p=prepared([row('2','172','','','Later',**{'Submit Date':'2026-05-02 10:00:00'}),row('1','172','','','Earlier')])
        self.assertEqual(educator_summary(p)['rows'][0]['strengths_comments'],'1. Earlier\n\n2. Later')
    def test_only_two_requested_comment_questions(self):
        p=prepared([row('1','10000','','','Not requested',Question='Another text question')]+comments())
        b=build_report_downloads(p,educator_summary(p))
        self.assertNotIn(b'Not requested',b['csv'])
    def test_source_student_columns_not_retained(self):
        p=prepared();self.assertNotIn('SYNTHETIC LEARNER',repr(p));self.assertNotIn('learner@example.edu',repr(p))
        self.assertNotIn('SYNTHETIC COMPLETER',repr(p))
    def test_missing_email_flagged_not_auto_filled(self):
        s=educator_summary(prepared([row(**{'Evaluator Email':''})]))
        self.assertTrue(s['issues']);self.assertEqual(s['rows'][0]['record_id'],'')
        self.assertEqual(s['rows'][0]['email_missing'],'YES')
    def test_manual_username_resolves_keeps_email_missing(self):
        p=prepared([row(**{'Evaluator Email':''})]);s=educator_summary(p,{'ext:staff-001':{'record_id':'new12'}})
        self.assertEqual(s['issues'],[]);self.assertEqual(s['rows'][0]['record_id'],'new12')
        self.assertEqual(s['rows'][0]['email_missing'],'YES');self.assertEqual(s['rows'][0]['evaluator_email'],'')
    def test_email_header_optional(self):
        p=prepare_reports([('a',make_csv(omit=['Evaluator Email']))]);self.assertTrue(educator_summary(p)['issues'])
    def test_invalid_email_flagged(self):
        self.assertTrue(educator_summary(prepared([row(**{'Evaluator Email':'invalid email'})]))['issues'])
    def test_duplicate_record_ids_block(self):
        rows=[row('1'),row('2',**{'Evaluator':'Other, Taylor','Evaluator Email':'abc12@other.edu','Evaluator Username':'xyz','Evaluator External ID':'staff002'})]
        s=educator_summary(prepared(rows));self.assertEqual(len(s['issues']),2)
    def test_full_email_is_not_username_override(self):
        with self.assertRaises(OASISReportError):validate_username('abc@example.edu')
    def test_username_validation(self):
        for s in ('','x x','../x','@abc','x\nkey','=cmd','+foo','a'*65):
            with self.subTest(s=s),self.assertRaises(OASISReportError):validate_username(s)
        self.assertEqual(validate_username('  Ab_1.z-2  '),'ab_1.z-2')
    def test_missing_ids_blocks(self):
        for field in ('Form Record','Question ID','Course ID','Evaluation'):
            with self.subTest(field=field),self.assertRaises(OASISReportError):prepared([row(**{field:''})])
    def test_non_numeric_question_id_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row(qid='A1')])
    def test_bad_scores_block(self):
        for value in ('words','NaN','Infinity','-Infinity','1e9999999'):
            with self.subTest(value=value),self.assertRaises(OASISReportError):prepared([row(value=value)])
    def test_same_qid_different_wording_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row(Question='Different measure')])
    def test_unknown_mc_ids_appended(self):
        p=prepared([row('1','9999',Question='New rating')]);s=educator_summary(p)
        self.assertEqual(s['rows'][0]['q9999_mean'],'5.00')
        self.assertEqual(question_key_rows(p)[-3]['question_id'],'9999')
    def test_duration_mean_is_kept_separate_and_explained(self):
        p=prepared([row('1','1286',value='1',label='less than 1/2 week'),row('2','1286',value='3',label='1 week')])
        s=educator_summary(p);self.assertEqual(s['rows'][0]['q1286_mean'],'2.00')
        self.assertNotIn('overall_mean',s['columns'])
        self.assertIn('not weeks',next(r for r in question_key_rows(p) if r['question_id']=='1286')['note'])
    def test_two_forms_same_student_still_count_two(self):
        self.assertEqual(educator_summary(prepared([row('1'),row('2')]))['rows'][0]['evaluation_count'],2)
    def test_unsubmitted_excluded(self):
        p=prepared([row('1'),row('2',**{'Submit Date':''})])
        self.assertEqual(p['unsubmitted_forms_excluded'],1);self.assertEqual(len(p['forms']),1)
    def test_invalid_submit_date_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row(**{'Submit Date':'bad'})])
    def test_inclusive_submit_dates(self):
        p=prepared([row('1',**{'Submit Date':'2026-03-01 23:59:00'}),row('2'),row('3',**{'Submit Date':'2026-03-02 01:00:00'})])
        s=educator_summary(p,start_date=date(2026,3,1),end_date=date(2026,3,2))
        self.assertEqual(s['evaluation_count'],2)
    def test_different_date_basis(self):
        p=prepared();s=educator_summary(p,start_date=date(2026,3,16),end_date=date(2026,3,16),date_field='Start Date')
        self.assertEqual(s['evaluation_count'],1)
    def test_missing_selected_date_blocks_not_dropped(self):
        p=prepared([row(**{'Start Date':''})])
        with self.assertRaises(OASISReportError):educator_summary(p,start_date=date(2026,1,1),end_date=date(2026,12,31),date_field='Start Date')
    def test_reverse_dates_block(self):
        with self.assertRaises(OASISReportError):educator_summary(prepared(),start_date=date(2026,5,1),end_date=date(2026,1,1))
    def test_course_form_filters(self):
        p=prepared([row('1'),row('2',**{'Evaluation':'Other form'}),row('3',**{'Course ID':'Other course'})])
        s=educator_summary(p,courses=['DEMO-101'],evaluation_types=['*Clinical Teaching Eval'])
        self.assertEqual(s['evaluation_count'],1)
    def test_no_forms_selected_no_report(self):
        p=prepared();s=educator_summary(p,evaluation_types=[])
        with self.assertRaises(OASISReportError):build_report_downloads(p,s)
    def test_missing_records_no_final_csv(self):
        p=prepared([row(**{'Evaluator Email':''})])
        with self.assertRaises(OASISReportError):build_report_downloads(p,educator_summary(p))
    def test_identity_when_external_id_missing_on_one_form(self):
        p=prepared([row('1'),row('2',**{'Evaluator External ID':''})]);self.assertEqual(len(educator_summary(p)['rows']),1)
    def test_identity_when_email_missing_on_one_form(self):
        p=prepared([row('1'),row('2',**{'Evaluator Email':''})]);s=educator_summary(p)
        self.assertEqual(s['issues'],[]);self.assertEqual(s['rows'][0]['evaluation_count'],2)
    def test_identity_ambiguous_external_ids_block(self):
        with self.assertRaises(OASISReportError):prepared([row('1'),row('2',**{'Evaluator External ID':'other-002'})])
    def test_changed_name_with_same_external_id_one_educator(self):
        p=prepared([row('1'),row('2',Evaluator='Example, A.')]);self.assertEqual(len(educator_summary(p)['rows']),1)
    def test_different_educators_same_name_not_merged(self):
        p=prepared([row('1'),row('2',**{'Evaluator External ID':'002','Evaluator Username':'002','Evaluator Email':'different@example.edu'})])
        self.assertEqual(len(educator_summary(p)['rows']),2)
    def test_missing_name_blocks(self):
        with self.assertRaises(OASISReportError):prepared([row(Evaluator='')])
    def test_empty_or_oversized_source_list_blocks(self):
        with self.assertRaises(OASISReportError):prepare_reports([])
        with self.assertRaises(OASISReportError):prepare_reports([('x',b'x')]*101)
    def test_encoding_bom_utf16_cp1252(self):
        for enc in ('utf-8-sig','utf-16','cp1252'):
            p=prepare_reports([('x',make_csv([row('1','172','','','Caf\u00e9')],encoding=enc))])
            self.assertEqual(educator_summary(p)['rows'][0]['strengths_comments'],'1. Caf\u00e9')
    def test_csv_bom_and_formula_safety(self):
        data=csv_bytes([{'a':'=1+2','b':'@cmd','c':'value'}],['a','b','c'])
        self.assertTrue(data.startswith(b'\xef\xbb\xbf'));self.assertIn(b"'=1+2",data)
    def test_zip_structure_and_no_raw_sources(self):
        p=prepared();d=build_report_downloads(p,educator_summary(p))
        with ZipFile(io.BytesIO(d['zip'])) as z:
            self.assertEqual(z.testzip(),None)
            self.assertEqual(set(z.namelist()),{'oasis_educator_summary.csv','oasis_question_key.csv','oasis_educator_question_detail.csv','OASIS_Report_Notes.txt'})
            for n in z.namelist():self.assertNotIn(b'SYNTHETIC LEARNER',z.read(n))
    def test_record_id_first_column(self):
        p=prepared();d=build_report_downloads(p,educator_summary(p))
        self.assertTrue(d['csv'].decode('utf-8-sig').startswith('record_id,'))
    def test_csv_all_rows_same_columns(self):
        p=prepared();s=educator_summary(p);d=build_report_downloads(p,s)
        rows=list(csv.reader(io.StringIO(d['csv'].decode('utf-8-sig'))))
        self.assertTrue(all(len(r)==42 for r in rows))


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.secrets=secret_settings();self.config=OPDArchiveConfig(**self.secrets['opd_archive'])
        self.repo=FakeGitHub();self.archive=GitHubOPDArchive(self.config,transport=self.repo)
        self.service=GitHubOASISUsernames(self.archive)
    def save(self,rid='abc12',key='ext:staff-001',expected=None):
        return self.service.save(key,'Example, Avery',rid,expected=expected or self.service.load())
    def test_empty_catalog_no_write(self):
        self.assertEqual(self.service.load()['entries'],{});self.assertEqual(self.repo.write_count,0)
    def test_encrypt_roundtrip(self):
        saved=self.save();self.assertEqual(self.service.load()['entries'],saved['entries'])
        raw=self.repo.tree[self.service.path];self.assertNotIn(b'abc12',raw)
        self.assertEqual(json.loads(self.config.cipher().decrypt(raw))['entries']['ext:staff-001']['record_id'],'abc12')
    def test_save_identical_no_extra_commit(self):
        a=self.save();self.save(expected=a);self.assertEqual(self.repo.write_count,1)
    def test_updated_username(self):
        a=self.save();b=self.save('abc13',expected=a);self.assertEqual(b['entries']['ext:staff-001']['record_id'],'abc13')
    def test_remove_keeps_other_entry(self):
        a=self.save();b=self.save('other12',key='ext:002',expected=a)
        c=self.service.remove('ext:staff-001',expected=b)
        self.assertEqual(set(c['entries']),{'ext:002'})
    def test_stale_revision_no_write(self):
        empty=self.service.load();self.save(expected=empty);count=self.repo.write_count
        with self.assertRaisesRegex(OASISReportError,'another session'):self.save('abc13',expected=empty)
        self.assertEqual(count,self.repo.write_count)
    def test_wrong_scope_blocks(self):
        expected=self.service.load();expected['scope']='other'
        with self.assertRaises(OASISReportError):self.save(expected=expected)
    def test_wrong_key_blocks(self):
        self.save();other=GitHubOASISUsernames(GitHubOPDArchive(replace(self.config,encryption_key=Fernet.generate_key().decode()),transport=self.repo))
        with self.assertRaisesRegex(OASISReportError,'decrypted'):other.load()
    def test_duplicate_override_ids_block(self):
        a=self.save()
        with self.assertRaisesRegex(OASISReportError,'duplicate'):self.save(key='ext:002',expected=a)
    def test_other_archives_untouched(self):
        self.repo.tree['opd_archive/OPD_2026-03-16.xlsx.enc']=b'keep';self.repo._commit()
        self.save();self.assertEqual(self.repo.tree['opd_archive/OPD_2026-03-16.xlsx.enc'],b'keep')
    def test_invalid_username_no_write(self):
        with self.assertRaises(OASISReportError):self.save('not valid')
        self.assertEqual(self.repo.write_count,0)
    def test_corrupt_json_blocks_no_overwrite(self):
        self.repo.tree[self.service.path]=self.config.cipher().encrypt(b'not json');self.repo._commit()
        with self.assertRaises(OASISReportError):self.service.load()
    def test_corrupted_token_blocks(self):
        self.repo.tree[self.service.path]=b'not cipher';self.repo._commit()
        with self.assertRaises(OASISReportError):self.service.load()
    def test_path_cannot_be_controlled_by_educator(self):
        self.save(key='name:../name')
        self.assertEqual(list(self.repo.tree),[self.service.path])


class InterfaceTests(unittest.TestCase):
    def setUp(self):
        self.secrets=secret_settings();self.repo=FakeGitHub()
    def inputs(self,raw=None):
        return {'schedule_app_mode':'OASIS Educator Reports','oer_source':'Upload CSV for this report only',
                'oer_uploads':[Upload(raw or make_csv(),'example.csv')],'oer_read':True}
    def runui(self,values,state=None):
        return run_app(values,secrets=self.secrets,state=state,repo=self.repo,evaluation_login=True,
                       original=ROOT/"tests"/"legacy_oasis_entrypoint.py")
    def test_menu_opens_new_section(self):
        r=self.runui({'schedule_app_mode':'OASIS Educator Reports'})
        self.assertTrue(any('OASIS Educator Reports' in text for kind,text in r['messages']))
    def test_local_upload_report_and_zip(self):
        values=self.inputs();values['oer_build_csv']=True
        r=self.runui(values)
        self.assertIn('oasis_educator_summary.csv',r['downloads'])
        self.assertIn('OASIS_Educator_Reports.zip',r['downloads'])
        self.assertEqual(self.repo.write_count,0)
    def test_missing_email_blocks_download_shows_issues(self):
        v=self.inputs(make_csv([row(**{'Evaluator Email':''})]));v['oer_build_csv']=True
        r=self.runui(v)
        self.assertNotIn('oasis_educator_summary.csv',r['downloads']);self.assertIn('OASIS_Username_Issues.csv',r['downloads'])
    def test_manual_id_saved_then_csv_enabled(self):
        v=self.inputs(make_csv([row(**{'Evaluator Email':''})]));key=__import__('hashlib').sha256(b'ext:staff-001').hexdigest()[:16]
        v['oer_username_'+key]='abc12';v['oer_apply_username']=True;v['oer_build_csv']=True
        r=self.runui(v)
        self.assertIn('oasis_educator_summary.csv',r['downloads']);self.assertEqual(self.repo.write_count,1)
        self.assertIn('opd_archive/oasis_educator_usernames.json.enc',self.repo.tree)
    def test_saved_override_reloads_next_session(self):
        service=GitHubOASISUsernames(GitHubOPDArchive(OPDArchiveConfig(**self.secrets['opd_archive']),transport=self.repo))
        service.save('ext:staff-001','Example, Avery','abc12',expected=service.load())
        v=self.inputs(make_csv([row(**{'Evaluator Email':''})]));v['oer_build_csv']=True
        r=self.runui(v);self.assertIn('oasis_educator_summary.csv',r['downloads'])
    def test_source_change_clears_old_download(self):
        v=self.inputs();v['oer_build_csv']=True;r=self.runui(v)
        v2=self.inputs(make_csv([row('2')]));v2['oer_read']=False;r2=self.runui(v2,r['state'])
        self.assertNotIn('oasis_educator_summary.csv',r2['downloads'])
    def test_conflicting_source_yields_diagnostic(self):
        r=self.runui(self.inputs(make_csv([row('1'),row('1',value='4')])))
        self.assertIn('OASIS_Source_Issues.csv',r['downloads'])
    def test_saved_archived_csv_workflow(self):
        service=GitHubOASISEvaluations(GitHubOPDArchive(OPDArchiveConfig(**self.secrets['opd_archive']),transport=self.repo))
        service.save(make_csv());writes=self.repo.write_count
        r=self.runui({'schedule_app_mode':'OASIS Educator Reports','oer_read':True,'oer_build_csv':True})
        self.assertIn('oasis_educator_summary.csv',r['downloads']);self.assertEqual(writes,self.repo.write_count)

if __name__=='__main__':unittest.main()
