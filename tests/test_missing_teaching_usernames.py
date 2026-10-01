"""Missing-only username entry, synthetic teaching/OASIS data, no live GitHub."""
from copy import deepcopy
from datetime import date
import hashlib
import unittest
from unittest.mock import patch

from helpers import st, StopRun, run_app
from test_teaching_oasis_links import Fixture
from test_strict_reach_charts import ui_values
from schedule_app.sections import preceptor_oasis_links as ui
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_workflow import GitHubOASISSummaries

P = ui.P


def fields(name, value='', existing=''):
    suffix=hashlib.sha256((name.lower()+'\0'+existing).encode()).hexdigest()[:16]
    return {P+'username_'+suffix:value, P+'confirm_username_'+suffix:True}


class QueueTests(unittest.TestCase):
    def setUp(self):
        self.scan={'monthly':[
            {'preceptor_name':'Alpha, Avery','academic_start_year':2026,'no_of_shifts':2},
            {'preceptor_name':'Beta, Bailey','academic_start_year':2026,'no_of_shifts':4},
            {'preceptor_name':'Beta, Bailey','academic_start_year':2026,'no_of_shifts':1},
            {'preceptor_name':'No Student, Casey','academic_start_year':2026,'no_of_shifts':0},
            {'preceptor_name':'Outside, Dana','academic_start_year':2025,'no_of_shifts':2},
            {'preceptor_name':'SJR_1','academic_start_year':2026,'no_of_shifts':2},
        ],'unresolved_preceptor_labels':['SJR_1']}
        self.catalog={'entries':{'alpha, avery':{'record_id':'aa'},'Old, Person':{'record_id':'old'}}}

    def test_saved_preceptor_is_excluded(self):
        active,missing=ui._username_candidates(self.scan,[2026],self.catalog)
        self.assertEqual(list(active),['alpha, avery','beta, bailey'])
        self.assertEqual(list(missing),['beta, bailey'])

    def test_same_name_in_many_months_only_once(self):
        self.assertEqual(len(ui._username_candidates(self.scan,[2026],{'entries':{}})[1]),2)

    def test_no_student_assignments_omitted(self):
        self.assertNotIn('no student, casey',ui._username_candidates(self.scan,[2026],self.catalog)[0])

    def test_outside_dates_omitted_even_without_mapping(self):
        self.assertNotIn('outside, dana',ui._username_candidates(self.scan,[2026],self.catalog)[1])

    def test_multiple_selected_years_union(self):
        self.assertEqual(set(ui._username_candidates(self.scan,[2025,2026],self.catalog)[1]),
                         {'beta, bailey','outside, dana'})

    def test_generic_labels_omitted(self):
        self.assertNotIn('sjr_1',ui._username_candidates(self.scan,[2026],self.catalog)[0])

    def test_saved_name_not_in_opd_never_enters_missing_list(self):
        self.assertNotIn('Old, Person',ui._username_candidates(self.scan,[2026],self.catalog)[1])

    def test_case_whitespace_and_comma_normalization(self):
        self.scan['monthly'][0]['preceptor_name']='  ALPHA  ,   Avery '
        self.assertNotIn('alpha, avery',ui._username_candidates(self.scan,[2026],self.catalog)[1])

    def test_all_mapped_queue_empty(self):
        self.catalog['entries']['beta, bailey']={'record_id':'bb'}
        self.assertEqual(ui._username_candidates(self.scan,[2026],self.catalog)[1],{})

    def test_no_selected_year_empty_queue(self):
        self.assertEqual(ui._username_candidates(self.scan,[],self.catalog),({},{}))

    def test_does_not_mutate_scan_or_catalog(self):
        original=deepcopy((self.scan,self.catalog))
        ui._username_candidates(self.scan,[2026],self.catalog)
        self.assertEqual((self.scan,self.catalog),original)


class MissingUsernameUITests(Fixture):
    def show(self,values=None,state=None,view=None):
        st.reset(values={P+'include':True,**(values or {})},secrets=self.secrets,state=state)
        choices={}
        original=st.selectbox
        def capture(label,options,*args,**kwargs):
            self.assertTrue(options, 'Never render an empty selectbox')
            choices[kwargs.get('key',label)]=list(options)
            return original(label,options,*args,**kwargs)
        stopped=False
        with patch.object(st,'selectbox',side_effect=capture):
            try:
                plan,ready=ui.render_teaching_oasis_links(self.archive,view or self.view,[2026])
            except StopRun:
                plan,ready=None,False
                stopped=True
        return {'state':dict(st.session_state),'messages':list(st.messages),'choices':choices,
                'plan':plan,'ready':ready,'rerun':stopped,'widgets':list(st.widget_keys)}

    def assert_message(self,result,kind,part):
        self.assertTrue(any(k==kind and part in text for k,text in result['messages']),part)

    def test_warning_names_and_dropdown_only_missing(self):
        self.linked()
        result=self.show()
        self.assertEqual(result['choices'][P+'preceptor_choice'],['unlinked, bailey'])
        self.assert_message(result,'warning','Username needed: 1 of 2')
        self.assert_message(result,'dataframe',"'preceptor_name': 'Unlinked, Bailey'")
        self.assertNotIn(P+'saved_preceptor_choice',result['choices'])

    def test_saved_name_without_oasis_row_still_excluded(self):
        c,_=self.linked(); self.links.save_username('Unlinked, Bailey','noeval',expected=c)
        result=self.show()
        self.assertNotIn(P+'preceptor_choice',result['choices'])
        self.assert_message(result,'success','All 2 teaching preceptors')
        self.assertTrue(result['ready'])
        self.assertTrue(any('not in selected' in row['status'] for row in result['plan']['bundle']['status']))

    def test_empty_catalog_displays_all_current_named_teachers(self):
        result=self.show()
        self.assertEqual(result['choices'][P+'preceptor_choice'],['example, avery','unlinked, bailey'])
        self.assert_message(result,'warning','Username needed: 2 of 2')

    def test_save_reduces_count_and_moves_to_next_name_on_rerun(self):
        first=self.show()
        before=self.repo.write_count
        saved=self.show({**fields('Example, Avery','abc12'),P+'save_username':True},first['state'])
        self.assertTrue(saved['rerun'])
        self.assertEqual(self.repo.write_count,before+1)
        again=self.show(state=saved['state'])
        self.assertEqual(again['choices'][P+'preceptor_choice'],['unlinked, bailey'])
        self.assert_message(again,'warning','Username needed: 1 of 2')
        self.assert_message(again,'success','Username for Example, Avery encrypted, saved, and verified')
        self.assertEqual(self.repo.write_count,before+1)

    def test_last_save_hides_entry_dropdown(self):
        self.linked(); first=self.show()
        saved=self.show({**fields('Unlinked, Bailey','bb'),P+'save_username':True},first['state'])
        again=self.show(state=saved['state'])
        self.assertNotIn(P+'preceptor_choice',again['choices'])
        self.assertFalse(any(key==P+'save_username' for label,key in again['widgets']))
        self.assert_message(again,'success','All 2 teaching preceptors')

    def test_successful_save_keeps_catalog_encrypted(self):
        before=dict(self.repo.tree)
        saved=self.show({**fields('Example, Avery','abc12'),P+'save_username':True})
        self.assertTrue(saved['rerun'])
        token=self.repo.tree[self.links.path]
        self.assertNotIn(b'Example, Avery',token)
        self.assertIn(b'abc12',self.archive.config.cipher().decrypt(token))
        self.assertTrue(all(self.repo.tree[path]==raw for path,raw in before.items()))

    def test_duplicate_username_refused_and_name_stays_in_queue(self):
        self.linked();first=self.show();before=dict(self.repo.tree)
        result=self.show({**fields('Unlinked, Bailey','abc12'),P+'save_username':True},first['state'])
        self.assertFalse(result['rerun'])
        self.assertFalse(result['ready'])
        self.assertEqual(result['choices'][P+'preceptor_choice'],['unlinked, bailey'])
        self.assertEqual(self.repo.tree,before)
        self.assert_message(result,'error','already linked to another preceptor')

    def test_save_requires_confirmation(self):
        values=fields('Example, Avery','abc12')
        values={k:(False if 'confirm' in k else v) for k,v in values.items()}
        before=self.repo.write_count
        result=self.show({**values,P+'save_username':True})
        self.assertFalse(result['rerun'])
        self.assertEqual(self.repo.write_count,before)

    def test_stale_write_keeps_current_catalog(self):
        first=self.show()
        remote=self.links.save_username('Outside, Dana','dd',expected=self.links.load())
        before=self.repo.write_count
        result=self.show({**fields('Example, Avery','abc12'),P+'save_username':True},first['state'])
        self.assertFalse(result['rerun']);self.assertFalse(result['ready'])
        self.assert_message(result,'error','changed in another session')
        self.assertEqual(self.repo.write_count,before)
        self.assertEqual(self.links.load()['sha'],remote['sha'])

    def test_refresh_obtains_links_saved_elsewhere(self):
        first=self.show()
        self.links.save_username('Example, Avery','abc12',expected=self.links.load())
        result=self.show({P+'refresh':True},first['state'])
        self.assertEqual(result['choices'][P+'preceptor_choice'],['unlinked, bailey'])

    def test_failed_catalog_load_is_not_reported_as_empty(self):
        with patch.object(ui.GitHubPreceptorOASISLinks,'load',side_effect=OPDArchiveError('Cannot decrypt catalog')):
            result=self.show()
        self.assert_message(result,'error','Username check unavailable')
        self.assertNotIn(P+'preceptor_choice',result['choices'])
        self.assertFalse(any(k=='warning' and 'Username needed:' in msg for k,msg in result['messages']))

    def test_warning_still_shown_when_summary_listing_fails(self):
        with patch.object(GitHubOASISSummaries,'list_outputs',side_effect=OPDArchiveError('Summary list failed')):
            result=self.show()
        self.assert_message(result,'warning','Username needed: 2 of 2')
        self.assertIn(P+'preceptor_choice',result['choices'])
        self.assertFalse(result['ready'])

    def test_username_save_does_not_require_summary_listing(self):
        with patch.object(GitHubOASISSummaries,'list_outputs',side_effect=AssertionError('Should save first')):
            result=self.show({**fields('Example, Avery','abc12'),P+'save_username':True})
        self.assertTrue(result['rerun'])
        self.assertEqual(self.links.load()['entries']['example, avery']['record_id'],'abc12')

    def test_removing_saved_link_returns_name_to_queue(self):
        self.linked(); first=self.show()
        suffix=hashlib.sha256(b'example, avery\0abc12').hexdigest()[:16]
        result=self.show({P+'edit_saved':True,P+'confirm_remove_'+suffix:True,P+'remove_username':True},first['state'])
        self.assertTrue(result['rerun'])
        again=self.show(state=result['state'])
        self.assertEqual(again['choices'][P+'preceptor_choice'],['example, avery','unlinked, bailey'])
        self.assert_message(again,'warning','Username needed: 2 of 2')
        self.assertTrue(all(v=='' for k,v in again['state'].items() if k.startswith(P+'username_')))

    def test_updating_saved_link_requires_opt_in_and_preserves_queue(self):
        self.linked();first=self.show()
        result=self.show({P+'edit_saved':True,**fields('Example, Avery','changed','abc12'),
                          P+'update_username':True},first['state'])
        self.assertTrue(result['rerun'])
        self.assertEqual(self.links.load()['entries']['example, avery']['record_id'],'changed')
        again=self.show(state=result['state'])
        self.assertEqual(again['choices'][P+'preceptor_choice'],['unlinked, bailey'])

    def test_saved_links_outside_period_only_in_opt_in_maintenance(self):
        c,_=self.linked();self.links.save_username('Outside, Dana','dd',expected=c)
        first=self.show()
        self.assertNotIn(P+'saved_preceptor_choice',first['choices'])
        result=self.show({P+'edit_saved':True},first['state'])
        self.assertIn('outside, dana',result['choices'][P+'saved_preceptor_choice'])
        self.assertNotIn('outside, dana',result['choices'][P+'preceptor_choice'])

    def test_generic_site_labels_not_in_queue(self):
        scan=deepcopy(self.view);scan['unresolved_preceptor_labels']=['Example, Avery']
        result=self.show(view=scan)
        self.assertEqual(result['choices'][P+'preceptor_choice'],['unlinked, bailey'])
        self.assert_message(result,'warning','Username needed: 1 of 1')

    def test_old_selection_is_reset_before_render(self):
        self.linked();first=self.show()
        state={**first['state'],P+'preceptor_choice':'example, avery'}
        result=self.show(state=state)
        self.assertEqual(result['state'][P+'preceptor_choice'],'unlinked, bailey')

    def test_save_clears_reports_but_preserves_opd_scan_and_date_preset(self):
        first=self.show()
        sentinel={'loaded':'unchanged'}
        state={**first['state'],'teaching_scan':sentinel,'teaching_zip':b'old',
               'teaching_zip_signature':'old','teaching_period_label':'26-27'}
        result=self.show({**fields('Example, Avery','abc12'),P+'save_username':True},state)
        self.assertTrue(result['rerun']);self.assertNotIn('teaching_zip',result['state'])
        self.assertEqual(result['state']['teaching_scan'],sentinel)
        self.assertEqual(result['state']['teaching_period_label'],'26-27')

    def test_username_edit_resets_confirmation_and_old_download(self):
        key=P+'confirm_username_demo'
        st.reset(state={key:True,'teaching_zip':b'old','teaching_scan':{'retained':True}})
        ui._username_edited(key)
        self.assertFalse(st.session_state[key]);self.assertNotIn('teaching_zip',st.session_state)
        self.assertEqual(st.session_state['teaching_scan'],{'retained':True})

    def test_mapped_missing_oasis_is_not_a_missing_username_alert(self):
        c,_=self.linked();self.links.save_username('Unlinked, Bailey','bb',expected=c)
        result=self.show()
        self.assert_message(result,'success','All 2 teaching preceptors')
        self.assertFalse(any(k=='warning' and 'Username needed:' in m for k,m in result['messages']))

    def test_disabled_links_still_makes_no_requests(self):
        with patch.object(self.archive,'_head',side_effect=AssertionError('no network')):
            result=self.show({P+'include':False})
        self.assertEqual((result['plan'],result['ready']),(None,True))

    def test_staged_partial_linkage_still_allows_teaching_only_for_unmapped(self):
        self.linked();result=self.show()
        self.assertTrue(result['ready'])
        self.assert_message(result,'warning','Username needed: 1 of 2')
        row=next(r for r in result['plan']['bundle']['status'] if r['preceptor_name']=='Unlinked, Bailey')
        self.assertEqual(row['record_id'],'')


if __name__=='__main__': unittest.main()
