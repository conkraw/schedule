"""Simplified PTS, protected matching, and one-build report reuse. Offline only."""
import ast
from copy import deepcopy
from datetime import date
from io import BytesIO
from unittest.mock import patch
from zipfile import ZipFile

import pytest
from helpers import ROOT, st, run_app, login_for_test, canonical
from test_learner_reach import scan_cells
from test_strict_reach_charts import ui_values
from test_assessment_completion import record, raw_csv, cells, NAME
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.evaluation_access import lock_evaluation_records, P as ACCESS_P
from schedule_app.sections import preceptor_teaching_summary as page
from schedule_app.sections import preceptor_oasis_links as links
from schedule_app.sections.pts_navigation import preserve_pts_preferences, open_pts_matching, return_to_pts
from schedule_app.sections.reporting_date_controls import clear_teaching_downloads
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.teaching_batch import TeachingReportBatch
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.chair_summary import teaching_make_chair_summary
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME


@pytest.fixture
def data():
    scan, repo, archive, secrets, raw = scan_cells({
        ('WARD A','B6'):'Example, Avery ~ Learner A; Learner B',
        ('WARD A','C6'):'Example, Avery ~ Learner A',
        ('WARD A','D8'):'Example, Avery ~ ',
        ('NYES','B6'):'Example, Blair ~ Learner C',
    })
    return dict(scan=scan, repo=repo, archive=archive, secrets=secrets)


def app(data, values=None, state=None, login=True):
    return run_app(ui_values(**(values or {})), secrets=data['secrets'], repo=data['repo'],
                   state=state, evaluation_login=login)


def test_menu_retains_oer_pts_at_end():
    tree=ast.parse((ROOT/'app_sch_2026.py').read_text())
    menu=next(ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign)
              and any(isinstance(t,ast.Name) and t.id=='SECTIONS' for t in n.targets))
    assert list(menu)[-3:]==['PTS Matching','OER','PTS']
    assert menu['PTS Matching']=='pts_matching'


@pytest.mark.parametrize('task',['Preceptor usernames','Student names','Ignored student entries'])
def test_matching_page_requires_password_before_source_controls(data,task):
    with patch.object(data['repo'],'request',wraps=data['repo'].request) as request:
        result=app(data,dict(schedule_app_mode='PTS Matching',pts_matching_task=task),login=False)
    assert not request.called
    assert not result['downloads']
    assert not any(key=='teaching_load_archives' for _,key in result['widgets'])
    assert any('PTS Matching — locked' in text for _,text in result['messages'])


def test_pts_has_no_routine_tables_or_matching_controls(data):
    result=app(data)
    assert 'teaching_zip' in result['state']
    assert not any(kind in ('dataframe','table','image') for kind,_ in result['messages'])
    keys={k for _,k in result['widgets']}
    assert links.P+'preceptor_choice' not in keys
    assert not any('student_choice' in k or 'oasis_choice_' in k for k in keys)
    assert not any(k.startswith('pts_ignored_students_choose_') for k in keys)
    assert 'pts_open_matching' in keys


def test_reports_still_contain_all_word_csv_and_chart_outputs(data):
    result=app(data)
    with ZipFile(BytesIO(result['state']['teaching_zip'])) as z:
        assert z.testzip() is None
        assert TEACHING_CHAIR_SUMMARY_FILENAME in z.namelist()
        assert len([f for f in z.namelist() if f.startswith('Preceptor_Reports/')])==2
        assert any(f.startswith('Learner_Reach_Charts/') and f.endswith('.png') for f in z.namelist())
        assert 'preceptor_teaching_summary.csv' in z.namelist()
        assert 'preceptor_student_assessment_completion.csv' in z.namelist()


def test_tables_are_available_only_when_requested(data):
    result=app(data,dict(pts_show_diagnostics=True))
    assert any(kind=='dataframe' for kind,_ in result['messages'])


def test_chart_previews_are_lazy(data):
    first=app(data)
    assert not any(kind=='image' for kind,_ in first['messages'])
    second=app(data,dict(teaching_load_archives=False,teaching_build_zip=False,pts_preview_charts=True),first['state'])
    assert any(kind=='image' for kind,_ in second['messages'])
    assert first['state']['teaching_zip']==second['state']['teaching_zip']


def test_preceptor_manager_shows_only_unmapped_names(data):
    service=GitHubPreceptorOASISLinks(data['archive'])
    service.save_username('Example, Avery','avery',expected=service.load())
    seen={}; original=st.selectbox
    def capture(label,options,*args,**kwargs):
        seen[kwargs.get('key',label)]=list(options)
        return original(label,options,*args,**kwargs)
    with patch.object(st,'selectbox',side_effect=capture):
        result=app(data,dict(schedule_app_mode='PTS Matching',pts_matching_task='Preceptor usernames'))
    assert seen[links.P+'preceptor_choice']==['example, blair']
    assert not any(kind=='dataframe' for kind,_ in result['messages'])
    assert 'teaching_build_zip' not in {k for _,k in result['widgets']}


def test_student_matching_runs_without_educator_feedback_enabled():
    scan,repo,archive,secrets,_=scan_cells(cells(two=True))
    GitHubOASISStudentEvaluations(archive).save(raw_csv([record(),record('two',sid='id-beta',student='Learner, Beth; MD2028')]))
    result=run_app(ui_values(schedule_app_mode='PTS Matching',pts_matching_task='Student names',
        assessment_completion_refresh=True),repo=repo,secrets=secrets,evaluation_login=True)
    assert 'assessment_completion_inputs' in result['state']
    assert any('OPD students needing an OASIS match'==label for label,_ in result['widgets'])
    assert not any(kind=='dataframe' for kind,_ in result['messages'])
    assert not any(k==links.P+'include' for _,k in result['widgets'])


def test_preceptor_matching_does_not_need_assessment_sources(data):
    result=app(data,dict(schedule_app_mode='PTS Matching',pts_matching_task='Preceptor usernames'))
    assert any(label=='Teaching preceptors needing a username' for label,_ in result['widgets'])
    assert not any(label=='Load / refresh evaluation completeness' for label,_ in result['widgets'])


def test_ignore_tools_only_on_matching_page(data):
    result=app(data,dict(schedule_app_mode='PTS Matching',pts_matching_task='Ignored student entries'))
    assert any('Ignore selected entries' in label for label,_ in result['widgets'])
    assert not any(k=='teaching_build_zip' for _,k in result['widgets'])


def test_navigating_between_pages_reuses_loaded_opds(data):
    first=app(data,dict(teaching_build_zip=False))
    import schedule_app.sections.pts_workspace as workspace
    with patch.object(workspace,'teaching_scan_archives',side_effect=AssertionError('unexpected rescan')):
        second=app(data,dict(schedule_app_mode='PTS Matching',pts_matching_task='Preceptor usernames',
                             teaching_load_archives=False),first['state'])
        third=app(data,dict(teaching_load_archives=False,teaching_build_zip=False),second['state'])
    assert first['state']['teaching_scan']==third['state']['teaching_scan']
    assert third['state']['teaching_period_start']==date(2026,9,7)


@pytest.mark.parametrize('key,value',[
    ('teaching_oasis_include',True),('teaching_name_order','Student ~ Preceptor'),
    ('teaching_selected_years',[2025,2026]),('assessment_completion_enabled',False),
    ('assessment_completion_course_choice',['PED']),('pts_matching_task','Student names')])
def test_widget_preferences_survive_page_navigation(key,value):
    st.reset(state={key:value})
    preserve_pts_preferences()
    assert st.session_state[key]==value


def test_locked_navigation_callbacks_cannot_open_matching_or_reports():
    st.reset(state={'schedule_app_mode':'Instructions'})
    open_pts_matching();return_to_pts()
    assert st.session_state['schedule_app_mode']=='Instructions'


def test_authorized_navigation_callbacks_switch_menu(data):
    st.reset(secrets=data['secrets']);login_for_test()
    open_pts_matching();assert st.session_state['schedule_app_mode']=='PTS Matching'
    return_to_pts();assert st.session_state['schedule_app_mode']=='PTS'


def test_lock_clears_new_download_preview_and_matching_state(data):
    st.reset(secrets=data['secrets'],state={'pts_matching_task':'Student names','teaching_zip_payload':{'private':'x'},
        'teaching_build_seconds':3,'opd_generated_master':b'keep'})
    login_for_test();lock_evaluation_records()
    assert 'pts_matching_task' not in st.session_state
    assert 'teaching_zip_payload' not in st.session_state
    assert st.session_state['opd_generated_master']==b'keep'


def test_date_change_clears_zip_and_extracted_chair_bytes():
    st.reset(state={'teaching_zip':b'z','teaching_zip_signature':'old','teaching_zip_payload':{'private':'x'},
                    'teaching_build_seconds':2})
    clear_teaching_downloads()
    assert not st.session_state


def test_identical_explicit_build_reuses_zip_after_freshness_checks(data):
    with patch.object(page,'teaching_build_zip',wraps=teaching_build_zip) as build:
        first=app(data)
        second=app(data,dict(teaching_load_archives=False),first['state'])
    assert build.call_count==1
    assert first['state']['teaching_zip'] is second['state']['teaching_zip']


def test_changed_dates_force_new_zip(data):
    with patch.object(page,'teaching_build_zip',wraps=teaching_build_zip) as build:
        first=app(data)
        second=app(data,dict(teaching_load_archives=False,teaching_period_end=date(2026,9,14)),first['state'])
    assert build.call_count==2
    assert first['state']['teaching_zip_signature']!=second['state']['teaching_zip_signature']


def test_clear_download_after_build_failure_never_retains_partial(data):
    with patch.object(page,'teaching_build_zip',side_effect=RuntimeError('test failure')):
        result=app(data)
    assert 'teaching_zip' not in result['state']
    assert not any(name.endswith(('.docx','.zip')) for name in result['downloads'])


def test_zip_download_rerun_does_not_rebuild_reports(data):
    first=app(data)
    with patch.object(page,'teaching_build_zip',side_effect=AssertionError('unexpected build')):
        second=app(data,dict(teaching_load_archives=False,teaching_build_zip=False),first['state'])
    assert 'teaching_zip' in second['state']


def test_shared_metrics_computed_three_times_total_not_per_preceptor(data):
    import schedule_app.reports.teaching_batch as bat
    with patch.object(bat,'teaching_time_rows',wraps=bat.teaching_time_rows) as compute:
        teaching_build_zip(data['scan'],[2026])
    assert compute.call_count==3


def test_progress_monotonic_and_complete_without_names(data):
    messages=[]
    teaching_build_zip(data['scan'],[2026],progress=lambda done,total,text:messages.append((done/total,text)))
    assert [r for r,_ in messages]==sorted(r for r,_ in messages)
    assert messages[0][0]==0 and messages[-1][0]==1
    assert sum('Writing preceptor report' in text for _,text in messages)==2
    assert all('Avery' not in text for _,text in messages)


def test_batched_individuals_equal_standalone_document_xml(data):
    blob,_=teaching_build_zip(data['scan'],[2026])
    with ZipFile(BytesIO(blob)) as z:
        for name in ('Example, Avery','Example, Blair'):
            expected=teaching_make_docx(name,data['scan']['monthly'],data['scan'])
            actual=z.read('Preceptor_Reports/'+name.replace(', ','_').replace(' ','_')+'_Teaching_Report.docx')
            assert canonical(actual)==canonical(expected)


def test_chair_xml_equal_to_standalone(data):
    blob,_=teaching_build_zip(data['scan'],[2026])
    with ZipFile(BytesIO(blob)) as z:
        assert canonical(z.read(TEACHING_CHAIR_SUMMARY_FILENAME))==canonical(teaching_make_chair_summary(data['scan'],[2026]))


def test_batch_rejects_other_scan_identity_or_unprepared_year(data):
    b=TeachingReportBatch(data['scan'],[2026])
    with pytest.raises(OPDArchiveError): b.rows(deepcopy(data['scan']),[2026])
    with pytest.raises(OPDArchiveError): b.rows(data['scan'],[2025])


def test_batch_output_never_mutates_scan(data):
    before=deepcopy(data['scan'])
    teaching_build_zip(data['scan'],[2026])
    assert data['scan']==before


def test_zip_build_never_calls_network(data):
    with patch('requests.request',side_effect=AssertionError('Unexpected network during document build')):
        teaching_build_zip(data['scan'],[2026])


def test_unused_tables_and_charts_are_not_constructed_for_display(data):
    first=app(data)
    with patch.object(page,'_optional_details',side_effect=AssertionError('unused preview')):
        result=app(data,dict(teaching_load_archives=False,teaching_build_zip=False),first['state'])
    assert not any(kind=='dataframe' for kind,_ in result['messages'])
