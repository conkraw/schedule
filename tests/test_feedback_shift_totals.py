"""Regression coverage for feedback restoration, shifts-only output and unique assessed totals.
Only invented records are used. No live GitHub or real evaluation data.
"""
from copy import deepcopy
from datetime import date
from io import BytesIO
from zipfile import ZipFile
from unittest.mock import patch
import csv
import pytest
from helpers import st, run_app
from test_documented_completion import make_inputs, run
from test_assessment_completion import NAME, record, raw_csv
from test_teaching_oasis_links import Fixture
from test_strict_reach_charts import ui_values
from test_reporting_dates import all_doc_text
from schedule_app.services.assessment_completion import (
    HP, COLUMNS, ASSESSMENT_VERSION, prepare_student_assessments, build_completion_bundle,
    completion_rows, unverified_bundle, assessed_students_total_display,
    validate_assessed_student_totals,
)
from schedule_app.services.educational_time import TIME_CSV_COLUMNS, teaching_time_rows
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.evaluation_access import lock_evaluation_records
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.chair_summary import teaching_make_chair_summary
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.services.teaching_evaluations import join_feedback
from schedule_app.sections.preceptor_oasis_links import (
    P, FEEDBACK_PREFERENCE_KEY, render_teaching_oasis_links, _feedback_inclusion_control,
)
from schedule_app.sections.pts_navigation import preserve_pts_preferences


def source_rows():
    return [record(), record(**{'Question ID':'another-question'}),
            record('repeat-clinical'), record('alpha-hp',form=HP),
            record('outside-clinical',sid='sid-outside',student='Outside, Casey'),
            record('other-hp',sid='sid-other',student='Other, Alex',form=HP)]


@pytest.fixture
def scenario():
    return make_inputs(records=source_rows())


def test_all_totals_are_unique_not_questions_forms_or_eligible_only(scenario):
    bundle, _, row=run(scenario)
    assert row['eligible_students']==2
    assert row['clinical_forms_submitted']==3
    assert row['hp_forms_submitted']==2
    assert row['clinical_unique_students_assessed_total']==2
    assert row['hp_unique_students_assessed_total']==2
    assert row['either_unique_students_assessed_total']==3
    assert row['clinical_students_evaluated']==row['hp_students_evaluated']==row['either_students_evaluated']==1
    assert row['clinical_completion_pct']==row['hp_completion_pct']==row['either_completion_pct']==50


def test_duplicate_snapshots_and_repeated_forms_do_not_raise_total(scenario):
    raw=raw_csv(source_rows())
    scenario['inputs']['prepared']=prepare_student_assessments([('one',raw),('two',raw)])
    row=run(scenario)[2]
    assert row['either_unique_students_assessed_total']==3


@pytest.mark.parametrize('minimum',[1,3,4,5,20])
def test_all_student_total_independent_of_shift_threshold(scenario,minimum):
    row=run(scenario,minimum_shifts=minimum)[2]
    assert row['either_unique_students_assessed_total']==3
    if minimum>3:
        assert row['eligible_students']==0
        assert row['either_completion_pct'] is None
        assert assessed_students_total_display(row)=='3'


def test_cutoff_filters_all_assessed_totals_using_submit_date(scenario):
    raw=raw_csv([record(submitted='2026-09-10 23:59:59'),
                 record('late',sid='new-id',student='Late, Casey',submitted='2026-09-11 00:00:00')])
    scenario['inputs']['prepared']=prepare_student_assessments([('one',raw)])
    assert run(scenario,as_of=date(2026,9,10))[2]['either_unique_students_assessed_total']==1
    assert run(scenario,as_of=date(2026,9,11))[2]['either_unique_students_assessed_total']==2


def test_handoff_unsubmitted_other_evaluator_and_outside_dates_not_counted(scenario):
    rows=source_rows()+[
        record('handoff',form='*PEDS Handoff',sid='skip1',student='Skip, One'),
        record('draft',sid='skip2',student='Skip, Two',submitted=''),
        record('other',sid='skip3',student='Skip, Three',email='different@example.edu'),
        record('early',sid='skip4',student='Skip, Four',submitted='2026-09-06'),
    ]
    scenario['inputs']['prepared']=prepare_student_assessments([('one',raw_csv(rows))])
    assert run(scenario)[2]['either_unique_students_assessed_total']==3


@pytest.mark.parametrize('problem',['no_sources','no_username','bad_metadata'])
def test_unchecked_or_bad_data_never_fabricate_zero(scenario,problem):
    if problem=='no_sources': scenario['inputs']['prepared']['source_count']=0
    elif problem=='no_username': scenario['inputs']['catalog']['entries']={}
    else:
        scenario['inputs']['prepared']=prepare_student_assessments([('one',raw_csv(source_rows()+[record('bad',sid='')]))])
    row=run(scenario)[2]
    assert row['either_unique_students_assessed_total'] is None
    assert assessed_students_total_display(row) in ('Not checked','Not verified')


def test_genuine_zero_is_numeric_when_sources_verified(scenario):
    scenario['inputs']['prepared']['forms']=[]
    row=run(scenario)[2]
    assert row['total_assessments_status']=='Calculated'
    assert assessed_students_total_display(row)=='0'


@pytest.mark.parametrize('field,value',[
    ('either_unique_students_assessed_total',1),('either_unique_students_assessed_total',5),
    ('clinical_unique_students_assessed_total',-1),('hp_unique_students_assessed_total',None),
    ('total_assessments_status','Not checked'),])
def test_tampered_totals_rejected(scenario,field,value):
    row=deepcopy(run(scenario)[2]);row[field]=value
    with pytest.raises(OPDArchiveError):validate_assessed_student_totals(row)


def test_public_reports_csv_and_preview_remove_three_day_measure(scenario):
    b,_,_=run(scenario);view=dict(scenario['view'],assessment_completion=b)
    rows=teaching_time_rows(view,[2026])
    assert 'unique_students_3plus_days' not in rows[0]
    payload,_=teaching_build_zip(view,[2026])
    with ZipFile(BytesIO(payload)) as z:
        for filename in z.namelist():
            if filename.endswith('.docx'):
                text=all_doc_text(z.read(filename))
            elif filename.endswith(('.csv','.txt')):
                text=z.read(filename).decode('utf-8-sig')
            else: continue
            assert '3+ days' not in text and 'unique_students_3plus_days' not in text
            assert 'sid-outside' not in text and 'Outside, Casey' not in text
        records=list(csv.DictReader(z.read('preceptor_student_assessment_completion.csv').decode('utf-8-sig').splitlines()))
        assert records[0]['either_unique_students_assessed_total']=='3'
        assert records[0]['eligible_students']=='2'
        individual=all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx'))
        assert 'Total students assessed (one per student): 3' in individual
        assert '1 / 2 (50.0%)' in individual
        assert 'Students assigned for 3+ shifts: 2' in individual


def test_totals_count_only_selected_courses(scenario):
    other=record('othercourse',sid='othercourse-student',student='Other, Course',**{'Course ID':'OTHER'})
    scenario['inputs']['prepared']=prepare_student_assessments([('one',raw_csv(source_rows()+[other]))])
    assert run(scenario,courses=['DEMO-101'])[2]['either_unique_students_assessed_total']==3


def test_setting_and_source_originals_not_mutated(scenario):
    before=deepcopy(scenario['inputs']);tree=dict(scenario['repo'].tree)
    b,_,_=run(scenario);teaching_build_zip(dict(scenario['view'],assessment_completion=b),[2026])
    assert scenario['inputs']==before and scenario['repo'].tree==tree


@pytest.fixture
def feedback_fixture():
    f=Fixture();f.setUp();f.linked();return f


def test_feedback_included_by_default_with_saved_links(feedback_fixture):
    f=feedback_fixture
    result=run_app(ui_values(),secrets=f.secrets,repo=f.repo,evaluation_login=True)
    raw=result['state']['teaching_zip']
    with ZipFile(BytesIO(raw)) as z:
        text=all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx'))
    assert 'Learner feedback on teaching' in text
    assert 'Built on my knowledge and skill base.' in text
    assert 'Clear feedback and helpful bedside explanations.' in text
    assert 'q588_mean' not in text
    assert any('Learner feedback ready:' in message for _,message in result['messages'])


def test_stale_legacy_false_widget_no_longer_silently_omits_feedback(feedback_fixture):
    f=feedback_fixture
    result=run_app(ui_values(),secrets=f.secrets,repo=f.repo,state={P+'include':False},evaluation_login=True)
    with ZipFile(BytesIO(result['state']['teaching_zip'])) as z:
        assert 'Learner feedback on teaching' in all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx'))


def test_explicit_feedback_optout_shows_alert_and_makes_no_feedback_request(feedback_fixture):
    f=feedback_fixture
    st.reset(secrets=f.secrets,values={P+'include':False})
    with patch.object(f.archive,'_head',side_effect=AssertionError('unexpected request')):
        assert render_teaching_oasis_links(f.archive,f.view,[2026])==(None,True)
    assert st.session_state[FEEDBACK_PREFERENCE_KEY] is False
    assert any('Learner feedback is OFF' in text for _,text in st.messages)


def test_preference_survives_widget_cleanup_but_lock_clears_it():
    st.reset(state={FEEDBACK_PREFERENCE_KEY:False})
    assert _feedback_inclusion_control() is False
    st.session_state.pop(P+'include',None)
    preserve_pts_preferences()
    assert _feedback_inclusion_control() is False
    st.reset(state={FEEDBACK_PREFERENCE_KEY:True})
    assert _feedback_inclusion_control() is True
    lock_evaluation_records()
    assert FEEDBACK_PREFERENCE_KEY not in st.session_state


def test_unmatched_feedback_is_explained_in_word(feedback_fixture):
    f=feedback_fixture;_,bundle=f.linked()
    word=teaching_make_docx('Unlinked, Bailey',f.view['monthly'],f.view,oasis_feedback=bundle)
    text=all_doc_text(word)
    assert 'Not attached: No username assigned.' in text
    assert 'not a zero evaluation score' in text


def test_missing_summary_is_explained_without_using_another_period(feedback_fixture):
    f=feedback_fixture;catalog=f.links.load()
    bundle=join_feedback(f.view,[2026],catalog,{},allow_missing_summaries=True)
    text=all_doc_text(teaching_make_docx('Example, Avery',f.view['monthly'],f.view,oasis_feedback=bundle))
    assert 'No exact-date OASIS summary linked' in text
    assert 'Clear feedback and helpful bedside explanations.' not in text


def test_feedback_restored_after_navigation_between_pts_matching_and_pts(feedback_fixture):
    f=feedback_fixture
    first=run_app(ui_values(),secrets=f.secrets,repo=f.repo,evaluation_login=True)
    middle=run_app(ui_values(schedule_app_mode='PTS Matching',pts_matching_task='Preceptor usernames',teaching_load_archives=False),
                   secrets=f.secrets,repo=f.repo,state=first['state'],evaluation_login=True)
    state=middle['state'];state.pop(P+'include',None) # simulate widget cleanup
    last=run_app(ui_values(teaching_load_archives=False),secrets=f.secrets,repo=f.repo,state=state,evaluation_login=True)
    with ZipFile(BytesIO(last['state']['teaching_zip'])) as z:
        assert 'Learner feedback on teaching' in all_doc_text(z.read('Preceptor_Reports/Example_Avery_Teaching_Report.docx'))


def test_no_retired_column_in_export_schemas():
    assert 'unique_students_3plus_days' not in COLUMNS
    assert 'unique_students_3plus_days' not in TIME_CSV_COLUMNS
    assert 'either_unique_students_assessed_total' in COLUMNS
