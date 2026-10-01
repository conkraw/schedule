"""Teaching/OASIS linkage tests. All data and remote services are invented/offline."""
import copy
import csv
from dataclasses import replace
from datetime import date
from io import BytesIO, StringIO
import json
import re
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from cryptography.fernet import Fernet
from helpers import st, run_app, FakeGitHub, secret_settings, canonical
from test_learner_reach import scan_cells, make_reach_fixture
from test_reporting_dates import all_doc_text
from test_strict_reach_charts import ui_values
from test_oasis_educator_reports import row, comments, make_csv
from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames
from schedule_app.services.oasis_educator_reports import QUESTION_TEXTS, csv_bytes
from schedule_app.services.oasis_workflow import (
    OASISOutputScope, GitHubOASISSummaries, load_cumulative_evaluations,
    make_period_summary, summary_csv, PERIOD_COLUMNS,
)
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks, period_key
from schedule_app.services.teaching_evaluations import (
    parse_saved_summary, join_feedback, feedback_for_preceptor, feedback_signature, load_feedback_bundle,
)
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.sections.preceptor_oasis_links import render_teaching_oasis_links, P
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME


class Fixture(unittest.TestCase):
    def setUp(self):
        cells = make_reach_fixture()
        cells['ETOWN', 'E8'] = 'Unlinked, Bailey ~ Private B'
        cells['ADOLMED', 'B6'] = 'No Student, Casey ~ '
        self.scan, self.repo, self.archive, self.secrets, _ = scan_cells(cells)
        self.period = ReportingPeriod('26-27', date(2026,9,7), date(2026,9,13))
        self.view = teaching_filter_date_range(self.scan, self.period)
        self.links = GitHubPreceptorOASISLinks(self.archive)
        self.sources = GitHubOASISEvaluations(self.archive)
        self.summaries = GitHubOASISSummaries(self.archive)
        self.scope = OASISOutputScope(self.period, ('DEMO-101',), ('*Clinical Teaching Eval',))
        data = []
        for fid, score in [('1','4'),('2','5')]:
            for qid in QUESTION_TEXTS:
                data.append(row(fid,qid,score, **{'Submit Date':'2026-09-10 10:00:00'}))
            data.append(row(fid,'172','',label='',answer='Clear feedback and helpful bedside explanations.'))
            data.append(row(fid,'437','',label='',answer='More time to discuss the differential.'))
        for r in data:
            r['Submit Date'] = '2026-09-10 10:00:00'
            r['Evaluator'] = 'Different OASIS Label, Avery'
        data.append(row('3', **{'Submit Date':'2026-09-11', 'Evaluator':'OASIS Only, Resident',
                          'Evaluator Email':'resident@example.edu','Evaluator Username':'resident',
                          'Evaluator External ID':'staff-resident'}))
        self.sources.save(make_csv(data))
        self.prepared = load_cumulative_evaluations(self.sources)
        self.names = GitHubOASISUsernames(self.archive).load()
        self.summary = make_period_summary(self.prepared, self.scope, self.names['entries'])
        self.receipt = self.summaries.save(self.scope, summary_csv(self.summary),
                                          prepared=self.prepared, username_catalog=self.names)
        self.parsed = parse_saved_summary(self.receipt)

    def linked(self):
        catalog = self.links.save_username('Example, Avery','abc12', expected=self.links.load())
        catalog = self.links.save_report(self.period.start,self.period.end,self.scope.filename,expected=catalog)
        return catalog, join_feedback(self.view,[2026],catalog,{2026:self.parsed})

    def mutate_csv(self, change=None, drop=()):
        reader = csv.DictReader(StringIO(self.receipt['raw'].decode('utf-8-sig')))
        fields = [f for f in reader.fieldnames if f not in drop]
        rows = list(reader)
        if change:
            change(rows, fields)
        return {**self.receipt, 'raw':csv_bytes(rows,fields)}


class CatalogTests(Fixture):
    def test_empty_catalog_is_not_written(self):
        before = self.repo.write_count
        self.assertEqual(self.links.load()['entries'],{})
        self.assertEqual(self.repo.write_count,before)

    def test_ciphertext_roundtrip_and_originals_unchanged(self):
        before = dict(self.repo.tree)
        catalog,_=self.linked()
        self.assertEqual(self.links.load()['entries'],catalog['entries'])
        token = self.repo.tree[self.links.path]
        self.assertNotIn(b'abc12', token)
        self.assertNotIn(b'Example, Avery',token)
        self.assertIn(b'abc12',self.archive.config.cipher().decrypt(token))
        for path,data in before.items(): self.assertEqual(data,self.repo.tree[path])

    def test_username_normalization(self):
        c=self.links.save_username('Example, Avery',' ABC12 ',expected=self.links.load())
        self.assertEqual(c['entries']['example, avery']['record_id'],'abc12')

    def test_same_username_link_is_noop(self):
        c,_=self.linked();before=self.repo.write_count
        self.links.save_username('Example, Avery','abc12',expected=c)
        self.assertEqual(before,self.repo.write_count)

    def test_repeated_report_selection_is_noop(self):
        c,_=self.linked();before=self.repo.write_count
        self.links.save_report(self.period.start,self.period.end,self.scope.filename,expected=c)
        self.assertEqual(before,self.repo.write_count)

    def test_two_preceptors_cannot_receive_same_username(self):
        c,_=self.linked()
        with self.assertRaisesRegex(OPDArchiveError,'already linked'):
            self.links.save_username('Unlinked, Bailey','abc12',expected=c)

    def test_full_email_rejected(self):
        with self.assertRaisesRegex(OPDArchiveError,'username only'):
            self.links.save_username('Example, Avery','abc12@example.edu',expected=self.links.load())

    def test_stale_catalog_cannot_overwrite(self):
        old=self.links.load(); self.linked()
        with self.assertRaisesRegex(OPDArchiveError,'another session'):
            self.links.save_username('Unlinked, Bailey','bb',expected=old)

    def test_wrong_scope_cannot_write(self):
        c=self.links.load();c['scope']='bad'
        with self.assertRaises(OPDArchiveError): self.links.save_username('Example, Avery','abc12',expected=c)

    def test_remove_username_preserves_report_selection_and_other_links(self):
        c,_=self.linked()
        c=self.links.save_username('Unlinked, Bailey','bb',expected=c)
        c=self.links.remove_username('Example, Avery',expected=c)
        self.assertNotIn('example, avery',c['entries']);self.assertIn('unlinked, bailey',c['entries'])
        self.assertEqual(len(c['report_links']),1)

    def test_remove_report_does_not_remove_username(self):
        c,_=self.linked();c=self.links.remove_report(self.period.start,self.period.end,expected=c)
        self.assertEqual(c['report_links'],{});self.assertEqual(len(c['entries']),1)

    def test_wrong_key_never_overwrites(self):
        self.linked();before=dict(self.repo.tree)
        bad=replace(self.archive.config,encryption_key=Fernet.generate_key().decode())
        wrong=GitHubPreceptorOASISLinks(GitHubOPDArchive(bad,transport=self.repo))
        with self.assertRaisesRegex(OPDArchiveError,'decrypt'): wrong.load()
        self.assertEqual(before,self.repo.tree)

    def test_report_period_mismatch_blocked(self):
        with self.assertRaisesRegex(OPDArchiveError,'same exact'):
            self.links.save_report(date(2026,9,8),self.period.end,self.scope.filename,expected=self.links.load())

    def test_corrupt_catalog_blocks_load(self):
        self.linked(); self.repo.tree[self.links.path]=b'bad';self.repo._commit()
        with self.assertRaises(OPDArchiveError): self.links.load()


class ReaderAndJoinTests(Fixture):
    def test_new_csv_embeds_full_source_wording(self):
        self.assertIn('q588_question',self.summary['columns'])
        self.assertEqual(self.summary['rows'][0]['q588_question'],'Built on my knowledge and skill base.')
        self.assertEqual(tuple(self.summary['columns'][-4:]),PERIOD_COLUMNS)

    def test_old_known_question_summary_supported(self):
        old=self.mutate_csv(drop=[c for c in self.summary['columns'] if c.endswith('_question')])
        result=parse_saved_summary(old)
        self.assertIn('Built on my knowledge and skill base.',[q['question'] for q in result['rows_by_id']['abc12']['questions']])

    def test_unknown_question_without_text_not_guessed(self):
        def edit(rows,fields):
            fields.extend(['q9999_mean','q9999_n'])
            for row_ in rows: row_.update(q9999_mean='4',q9999_n='1')
        with self.assertRaisesRegex(OPDArchiveError,'no full wording'):
            parse_saved_summary(self.mutate_csv(edit))

    def test_new_question_with_full_text_supported(self):
        def edit(rows,fields):
            fields.extend(['q9999_mean','q9999_n','q9999_question'])
            for row_ in rows: row_.update(q9999_mean='4',q9999_n='1',q9999_question='Invented additional teaching question.')
        result=parse_saved_summary(self.mutate_csv(edit))
        self.assertEqual(result['rows_by_id']['abc12']['questions'][-1]['question'],'Invented additional teaching question.')

    def test_changed_known_question_blocked(self):
        def edit(rows,fields):
            for r in rows:r['q588_question']='Different question'
        with self.assertRaisesRegex(OPDArchiveError,'changed wording'):parse_saved_summary(self.mutate_csv(edit))

    def test_mismatched_question_wording_between_rows_blocked(self):
        def edit(rows,fields):rows[0]['q588_question']='Different question'
        with self.assertRaisesRegex(OPDArchiveError,'inconsistent'):parse_saved_summary(self.mutate_csv(edit))

    def test_nonfinite_mean_blocked(self):
        def edit(rows,fields): rows[0]['q588_mean']='NaN'
        with self.assertRaisesRegex(OPDArchiveError,'average is invalid'):parse_saved_summary(self.mutate_csv(edit))

    def test_missing_response_count_column_blocked(self):
        with self.assertRaisesRegex(OPDArchiveError,'missing its response'):parse_saved_summary(self.mutate_csv(drop=['q588_n']))

    def test_count_exceeds_evaluation_count_blocked(self):
        def edit(rows,fields):rows[0]['q588_n']='999'
        with self.assertRaisesRegex(OPDArchiveError,'response count'):parse_saved_summary(self.mutate_csv(edit))

    def test_missing_comments_blocked(self):
        with self.assertRaisesRegex(OPDArchiveError,'comment'):parse_saved_summary(self.mutate_csv(drop=['strengths_comments']))

    def test_zero_response_blank_mean_is_not_scored(self):
        def edit(rows,fields):
            for r in rows:r.update(q588_mean='',q588_n='0')
        result=parse_saved_summary(self.mutate_csv(edit))
        q=next(q for q in result['rows_by_id']['abc12']['questions'] if q['question']==QUESTION_TEXTS['588'])
        self.assertEqual((q['mean'],q['response_count']),('Not scored',0))

    def test_username_match_does_not_require_name_match(self):
        _,bundle=self.linked()
        entry=feedback_for_preceptor(bundle,self.view,'example,   AVERY',2026)
        self.assertEqual(entry['educator_name'],'Different OASIS Label, Avery')
        self.assertEqual(entry['evaluation_count'],2)

    def test_oasis_only_educator_excluded(self):
        _,bundle=self.linked()
        self.assertNotIn('resident',{r['record_id'] for g in bundle['periods'].values() for r in g['preceptors'].values()})
        self.assertNotIn('OASIS Only',json.dumps(bundle))

    def test_unmapped_preceptor_gets_no_feedback(self):
        _,bundle=self.linked()
        self.assertIsNone(feedback_for_preceptor(bundle,self.view,'Unlinked, Bailey',2026))

    def test_unmatched_username_not_reported_as_zero_evaluations(self):
        c,_=self.linked();c=self.links.save_username('Unlinked, Bailey','noeval',expected=c)
        bundle=join_feedback(self.view,[2026],c,{2026:self.parsed})
        r=next(x for x in bundle['status'] if x['preceptor_name']=='Unlinked, Bailey')
        self.assertEqual(r['evaluation_count'],'');self.assertIn('not in selected',r['status'])

    def test_generic_slot_not_matched_even_with_saved_map(self):
        scan=copy.deepcopy(self.view);scan['unresolved_preceptor_labels']=['Example, Avery']
        c,_=self.linked();bundle=join_feedback(scan,[2026],c,{2026:self.parsed})
        self.assertIsNone(feedback_for_preceptor(bundle,scan,'Example, Avery',2026))

    def test_period_mismatch_blocked_even_when_labels_match(self):
        c,_=self.linked();summary=copy.deepcopy(self.parsed);summary['details']['end_date']=date(2027,9,13)
        with self.assertRaisesRegex(OPDArchiveError,'dates do not exactly'):join_feedback(self.view,[2026],c,{2026:summary})

    def test_no_summary_period_does_not_reuse_other_year(self):
        c,_=self.linked()
        with self.assertRaisesRegex(OPDArchiveError,'Choose and save'):join_feedback(self.view,[2026],c,{})

    def test_read_preflight_unchanged_no_writes(self):
        c,b=self.linked();before=self.repo.write_count
        fresh=load_feedback_bundle(self.archive,self.view,[2026],c,{2026:self.parsed})
        self.assertEqual(feedback_signature(b),feedback_signature(fresh));self.assertEqual(before,self.repo.write_count)

    def test_preflight_stale_mapping_blocks(self):
        c,_=self.linked();self.links.save_username('Unlinked, Bailey','bb',expected=c)
        with self.assertRaisesRegex(OPDArchiveError,'links changed'):load_feedback_bundle(self.archive,self.view,[2026],c,{2026:self.parsed})

    def test_preflight_updated_csv_blocks_until_refresh(self):
        c,_=self.linked()
        summary=copy.deepcopy(self.summary);summary['rows'][0]['strengths_comments']+='\n\n3. Changed saved feedback.'
        self.summaries.save(self.scope,summary_csv(summary),prepared=self.prepared,username_catalog=self.names)
        with self.assertRaisesRegex(OPDArchiveError,'summary was updated'):load_feedback_bundle(self.archive,self.view,[2026],c,{2026:self.parsed})

    def test_bundle_cannot_be_reused_for_new_dates(self):
        _,b=self.linked();view=teaching_filter_date_range(self.scan,ReportingPeriod('Changed',date(2026,9,8),date(2026,9,13)))
        with self.assertRaisesRegex(OPDArchiveError,'different dates'):feedback_for_preceptor(b,view,'Example, Avery',2026)

    def test_signature_changes_for_username_but_not_unrelated_commit(self):
        _,b=self.linked();other=copy.deepcopy(b);other['periods']['2026']['snapshot']='different'
        self.assertEqual(feedback_signature(b),feedback_signature(other))
        other['mapping_sha']='new';self.assertNotEqual(feedback_signature(b),feedback_signature(other))


class WordAndUITests(Fixture):
    def test_linked_doc_contains_question_text_scores_count_comments(self):
        _,bundle=self.linked()
        raw=teaching_make_docx('Example, Avery',self.view['monthly'],self.view,oasis_feedback=bundle)
        text=all_doc_text(raw)
        for word in ['Built on my knowledge and skill base.','4.50','Submitted evaluations:',
                     'Please indicate this educator\'s strengths','Areas for Improvement',
                     'Clear feedback and helpful bedside explanations.', 'More time to discuss the differential.']:
            self.assertIn(word,text)
        self.assertNotRegex(text,r'q\d+_(?:mean|n|question)')
        self.assertNotIn('Private B',text)

    def test_duration_is_separate_not_a_quality_score(self):
        _,b=self.linked();text=all_doc_text(teaching_make_docx('Example, Avery',self.view['monthly'],self.view,oasis_feedback=b))
        self.assertIn('Time with the preceptor',text);self.assertIn('not weeks, days, or a teaching-quality rating',text)

    def test_unmatched_report_body_identical(self):
        _,b=self.linked()
        before=teaching_make_docx('Unlinked, Bailey',self.view['monthly'],self.view)
        after=teaching_make_docx('Unlinked, Bailey',self.view['monthly'],self.view,oasis_feedback=b)
        before_text, after_text = all_doc_text(before), all_doc_text(after)
        self.assertNotIn('Learner feedback on teaching', before_text)
        self.assertIn('Not attached: No username assigned.', after_text)
        # Only the explanatory missing-feedback disclosure is new.
        self.assertEqual([t.text for t in __import__('docx').Document(BytesIO(before)).paragraphs],
                         [t.text for t in __import__('docx').Document(BytesIO(after)).paragraphs][:-3])

    def test_zip_never_adds_oasis_only_documents_or_extra_feedback_csv(self):
        _,b=self.linked();before,_=teaching_build_zip(self.view,[2026]);after,_=teaching_build_zip(self.view,[2026],oasis_feedback=b)
        with ZipFile(BytesIO(before)) as x, ZipFile(BytesIO(after)) as y:
            self.assertEqual(set(x.namelist()),set(y.namelist()))
            for f in x.namelist():
                if f.endswith('.csv') or f==TEACHING_CHAIR_SUMMARY_FILENAME:
                    self.assertEqual(canonical(x.read(f)),canonical(y.read(f)))
            self.assertNotIn('resident',str(y.namelist()))

    def test_disabled_panel_makes_no_network_calls(self):
        st.reset(secrets=self.secrets, values={P+'include':False})
        with patch.object(self.archive,'_head',side_effect=AssertionError('network called')):
            self.assertEqual(render_teaching_oasis_links(self.archive,self.view,[2026]),(None,True))

    def test_ui_saved_links_match_preview_and_generate(self):
        self.linked()
        result=run_app(ui_values(**{P+'include':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        zips=[data for f,data in result['downloads'].items() if f.endswith('.zip')]
        self.assertTrue(zips,msg=result['messages'])
        with ZipFile(BytesIO(zips[0])) as z:
            text=all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx'))
            self.assertIn('Learner feedback on teaching',text)
        self.assertTrue(any('Learner feedback ready:' in msg for _,msg in result['messages']))

    def test_ui_can_save_username_and_report_selection(self):
        # No links initially: values select the actual OASIS username explicitly.
        suffix=__import__('hashlib').sha256(b'example, avery\0').hexdigest()[:16]
        values=ui_values(**{P+'include':True,P+'save_report_'+period_key(self.period.start,self.period.end):True,
                           P+'username_'+suffix:'abc12',P+'confirm_username_'+suffix:True,P+'save_username':True})
        result=run_app(dict(values,schedule_app_mode='PTS Matching',pts_matching_task='Preceptor usernames'),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        self.assertEqual(self.links.load()['entries']['example, avery']['record_id'],'abc12')
        # Verified saves now rerun immediately to remove the completed name from
        # the entry queue. Select/save the summary and build on the following run.
        result=run_app(ui_values(**{P+'include':True,'teaching_load_archives':False,
                       P+'save_report_'+period_key(self.period.start,self.period.end):True}),
                       secrets=self.secrets,repo=self.repo,state=result['state'], evaluation_login=True)
        self.assertTrue(any(f.endswith('.zip') for f in result['downloads']),msg=result['messages'])

    def test_ui_missing_period_blocks_linked_reports(self):
        result=run_app(ui_values(**{P+'include':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        self.assertFalse(any(f.endswith('.zip') for f in result['downloads']))
        self.assertTrue(any('Save this summary selection' in msg for _,msg in result['messages']))

    def test_ui_date_mismatch_has_clear_instruction(self):
        self.linked()
        result=run_app(ui_values(**{P+'include':True,'teaching_period_end':date(2026,9,14)}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        # Missing exact-date feedback is now a warning, not a teaching-report block.
        self.assertTrue(any(f.endswith('.zip') for f in result['downloads']))
        self.assertTrue(any('No saved OASIS summary has these exact dates' in msg for _,msg in result['messages']))

    def test_ui_toggle_off_removes_evaluation_sections_and_keeps_teaching(self):
        self.linked();first=run_app(ui_values(**{P+'include':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        second=run_app(ui_values(**{P+'include':False,'teaching_load_archives':False}),secrets=self.secrets,repo=self.repo,state=first['state'], evaluation_login=True)
        raw=next(v for k,v in second['downloads'].items() if k.endswith('.zip'))
        with ZipFile(BytesIO(raw)) as z:
            self.assertNotIn('Learner feedback on teaching',all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx')))

    def test_ui_no_changes_keeps_generated_zip_on_rerun(self):
        self.linked();first=run_app(ui_values(**{P+'include':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        second=run_app(ui_values(**{P+'include':True,'teaching_load_archives':False,'teaching_build_zip':False}),
                       secrets=self.secrets,repo=self.repo,state=first['state'], evaluation_login=True)
        self.assertIn('teaching_zip',second['state']);self.assertEqual(first['state']['teaching_zip'],second['state']['teaching_zip'])

    def test_ui_pending_username_edit_invalidates_previous_download(self):
        self.linked();first=run_app(ui_values(**{P+'include':True}),secrets=self.secrets,repo=self.repo, evaluation_login=True)
        suffix=__import__('hashlib').sha256(b'example, avery\0abc12').hexdigest()[:16]
        second=run_app(ui_values(**{P+'include':True,'teaching_load_archives':False,'teaching_build_zip':False,'schedule_app_mode':'PTS Matching','pts_matching_task':'Preceptor usernames',P+'edit_saved':True,P+'username_'+suffix:'different'}),
                       secrets=self.secrets,repo=self.repo,state=first['state'], evaluation_login=True)
        self.assertNotIn('teaching_zip',second['state'])


if __name__=='__main__':unittest.main()
