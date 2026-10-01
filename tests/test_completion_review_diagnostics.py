"""Completion review, navigation and verified-name correction. Invented records only."""
from copy import deepcopy
from datetime import date
from io import BytesIO
from unittest.mock import patch
from zipfile import ZipFile

import pytest
from helpers import st, login_for_test, StopRun
from test_learner_reach import scan_cells
from test_assessment_completion import record, raw_csv, NAME
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.assessment_completion import (
    load_completion_inputs, build_completion_bundle, completion_context,
)
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.student_name_review import student_name_review, oasis_student_choices, name_only_student_choices
from schedule_app.services.assessment_diagnostics import (
    FOCUS_KEY, unavailable_completion_choices, unresolved_students_for_row,
    completion_review_detail, make_student_review_focus, focused_missing_names,
)
from schedule_app.sections.assessment_diagnostics import render_completion_diagnostics, open_targeted_student_review, P
from schedule_app.sections.student_name_matches import render_student_name_matches, P as MATCH_P
from schedule_app.services.evaluation_access import lock_evaluation_records
from schedule_app.reports.teaching_export import teaching_build_zip
from test_reporting_dates import all_doc_text


@pytest.fixture
def scenario():
    # 8 eligible students, 6 match. Both two mismatches have correct records with
    # different spelling; below-threshold and another preceptor's names stay separate.
    students = [f'Learner, Match{i}' for i in range(6)] + ['Learner, Alphaa', 'Learner, Betaa']
    text='; '.join(students)
    cells = {('NYES','B6'):f'{NAME} ~ {text}; Learner, Below',
             ('NYES','B8'):f'{NAME} ~ {text}', ('NYES','C6'):f'{NAME} ~ {text}',
             ('NYES','D6'):'Example, Blair ~ Learner, Other',
             ('NYES','D8'):'Example, Blair ~ Learner, Other',
             ('NYES','E6'):'Example, Blair ~ Learner, Other'}
    scan, repo, archive, secrets, _=scan_cells(cells)
    scan=teaching_filter_date_range(scan,ReportingPeriod('26-27', date(2026,9,7),date(2026,9,13)))
    links=GitHubPreceptorOASISLinks(archive)
    links.save_username(NAME,'avery',expected=links.load())
    links.save_username('Example, Blair','blair',expected=links.load())
    rows=[record(str(i+1),sid=f'id-{i}',student=f'Learner, Match{i}') for i in range(6)]
    rows += [record('alpha',sid='id-alpha',student='Learner, Alpha'),
             record('beta',sid='id-beta',student='Learner, Beta')]
    GitHubOASISStudentEvaluations(archive).save(raw_csv(rows))
    inputs=load_completion_inputs(archive,scan,[2026])
    bundle, unmatched=build_completion_bundle(inputs,scan,[2026])
    row=next(r for r in bundle['rows'] if r['preceptor_name']==NAME)
    return dict(archive=archive,repo=repo,secrets=secrets,scan=scan,inputs=inputs,
                bundle=bundle, unmatched=unmatched,row=row)


def show(s, values=None, state=None, auth=True):
    st.reset(values=values,secrets=s['secrets'],state=state)
    if auth:login_for_test()
    render_completion_diagnostics(s['archive'],s['inputs'],s['bundle'],s['unmatched'])


def test_eight_eligible_two_missing_keeps_denominator_and_blocks_only_percentage(scenario):
    s=scenario
    assert s['row']['eligible_students']==8
    assert s['row']['clinical_forms_submitted']==8
    assert s['row']['clinical_completion_pct'] == 75
    assert s['row']['assessment_status'].startswith('Provisional:')
    d=completion_review_detail(s['row'],s['unmatched'])
    assert d['eligible_students']==8
    assert len(d['unresolved_students'])==2
    assert all(r['assigned_shifts']==3 for r in d['unresolved_students'])


def test_diagnostic_names_exclude_below_threshold_and_other_teachers(scenario):
    d=completion_review_detail(scenario['row'],scenario['unmatched'])
    assert {r['student_name'] for r in d['unresolved_students']}=={'Learner, Alphaa','Learner, Betaa'}


def test_unavailable_choices_exclude_calculated_and_no_eligible(scenario):
    b=deepcopy(scenario['bundle'])
    for status in ('Calculated','No eligible students (3+ shifts)'):
        b['rows'].append(dict(scenario['row'],preceptor_name=status,assessment_status=status))
    assert len(unavailable_completion_choices(b))==3  # Partial, calculated results remain inspectable.


def test_display_choices_sorted_and_stable(scenario):
    b=scenario['bundle']; reverse=dict(b,rows=list(reversed(b['rows'])))
    assert unavailable_completion_choices(b)==unavailable_completion_choices(reverse)
    assert list(unavailable_completion_choices(b))==list(unavailable_completion_choices(reverse))


def test_same_preceptor_different_year_excluded(scenario):
    u=deepcopy(scenario['unmatched'])
    u.append(dict(u[0],academic_year='27-28',name_key='other',student_name='Learner, Else'))
    assert unresolved_students_for_row(scenario['row'],u)==unresolved_students_for_row(scenario['row'],scenario['unmatched'])


def test_diagnostics_do_not_mutate_inputs(scenario):
    b,u=deepcopy(scenario['bundle']),deepcopy(scenario['unmatched'])
    completion_review_detail(scenario['row'],u); unavailable_completion_choices(b)
    assert b==scenario['bundle'] and u==scenario['unmatched']


def test_no_internal_ids_or_comments_in_diagnostic_detail(scenario):
    d=completion_review_detail(scenario['row'],scenario['unmatched'])
    assert 'external_id' not in str(d)
    assert 'CONFIDENTIAL_COMMENT' not in str(d)


def test_focus_dynamic_and_does_not_change_eligibility(scenario):
    s=scenario; f=make_student_review_focus(s['bundle'],s['row'])
    review=student_name_review(s['inputs'],s['scan'],[2026],s['unmatched'])
    queue=focused_missing_names(review['missing'],s['unmatched'],f,completion_context(s['scan'],[2026]))
    assert len(queue)==2 and len(review['missing'])==2 and len(review['absent'])==2
    rebuilt,_=build_completion_bundle(s['inputs'],s['scan'],[2026])
    assert rebuilt==s['bundle']


@pytest.mark.parametrize('changed',['opd_commit','periods','student_exclusions_signature'])
def test_focus_does_not_follow_changed_reporting_context(scenario,changed):
    s=scenario; f=make_student_review_focus(s['bundle'],s['row']); c=deepcopy(f['context']);c[changed]='changed'
    assert focused_missing_names({},s['unmatched'],f,c) is None


def test_invalid_focus_is_ignored(scenario):
    assert focused_missing_names({},[],None,completion_context(scenario['scan'],[2026])) is None


def test_default_screen_no_extra_tables(scenario):
    show(scenario)
    assert any('Optional assessment review' in text for _,text in st.messages)
    assert not any(kind=='dataframe' for kind,_ in st.messages)
    assert not st.downloads


def test_opt_in_shows_only_selected_two_student_names(scenario):
    show(scenario,{P+'show':True})
    texts='\n'.join(text for kind,text in st.messages if kind=='dataframe')
    assert 'Learner, Alphaa' in texts and 'Learner, Betaa' in texts
    assert 'Learner, Below' not in texts and 'Learner, Other' not in texts
    assert 'id-alpha' not in texts and 'id-beta' not in texts


def test_locked_diagnostic_exposes_no_names_or_reads(scenario):
    with patch.object(GitHubStudentAssessmentLinks,'load',side_effect=AssertionError('network before login')):
        show(scenario,{P+'show':True,P+'recheck':True},auth=False)
    assert not st.messages and not st.widget_keys


def test_authorized_navigation_uses_student_page_and_selected_filter(scenario):
    s=scenario;f=make_student_review_focus(s['bundle'],s['row'])
    st.reset(secrets=s['secrets']);login_for_test();open_targeted_student_review(f)
    assert st.session_state['schedule_app_mode']=='PTS Matching'
    assert st.session_state['pts_matching_task']=='Student names'
    assert st.session_state[FOCUS_KEY]==f


def test_locked_navigation_cannot_route_sensitive_data(scenario):
    st.reset(state={'schedule_app_mode':'Instructions'})
    open_targeted_student_review(make_student_review_focus(scenario['bundle'],scenario['row']))
    assert st.session_state['schedule_app_mode']=='Instructions' and FOCUS_KEY not in st.session_state


def test_recheck_reads_verified_catalog_and_recalculates_not_cached_old_map(scenario):
    s=scenario;store=GitHubStudentAssessmentLinks(s['archive'])
    store.save_link('Learner, Alphaa','id-alpha',expected=store.load())
    store.save_link('Learner, Betaa','id-beta',expected=store.load())
    old_count=s['repo'].write_count
    with pytest.raises(StopRun):
        show(s,{P+'show':True,P+'recheck':True},state={'teaching_zip':b'old','teaching_zip_payload':{'old':1}})
    inputs=st.session_state['assessment_completion_inputs']
    b,u=build_completion_bundle(inputs,s['scan'],[2026]);r=b['rows'][0]
    assert r['preceptor_name']==NAME and r['eligible_students']==8
    assert r['clinical_completion_pct']==100 and r['hp_completion_pct']==0
    assert len(unresolved_students_for_row(r,u))==0
    assert 'teaching_zip' not in st.session_state and 'teaching_zip_payload' not in st.session_state
    assert s['repo'].write_count==old_count  # explicit recheck is read-only


def test_recheck_failure_keeps_existing_map_and_does_not_invent_success(scenario):
    s=scenario;original=deepcopy(s['inputs']);state={'assessment_completion_inputs':original}
    with patch.object(GitHubStudentAssessmentLinks,'load',side_effect=OPDArchiveError('Could not decrypt saved catalog')):
        show(s,{P+'show':True,P+'recheck':True},state=state)
    assert st.session_state['assessment_completion_inputs']==original
    assert any(kind=='error' for kind,_ in st.messages)
    assert P+'notice' not in st.session_state


def test_verified_name_saves_remove_focused_names_not_other_students(scenario):
    s=scenario;focus=make_student_review_focus(s['bundle'],s['row'])
    store=GitHubStudentAssessmentLinks(s['archive'])
    saved=store.save_link('Learner, Alphaa','id-alpha',expected=store.load())
    inputs=dict(s['inputs'],student_links=saved);b,u=build_completion_bundle(inputs,s['scan'],[2026])
    review=student_name_review(inputs,s['scan'],[2026],u)
    queue=focused_missing_names(review['missing'],u,focus,completion_context(s['scan'],[2026]))
    assert len(queue)==1 and next(iter(queue.values()))['student_name']=='Learner, Betaa'
    assert b['rows'][0]['eligible_students']==8 and b['rows'][0]['clinical_completion_pct']==87.5


def test_matching_dropdown_honors_focus(scenario):
    s=scenario;f=make_student_review_focus(s['bundle'],s['row'])
    st.reset(secrets=s['secrets'],state={FOCUS_KEY:f});login_for_test()
    options={}; orig=st.selectbox
    def capture(label,opts,*args,**kwargs):
        options[label]=list(opts);return orig(label,opts,*args,**kwargs)
    with patch.object(st,'selectbox',side_effect=capture):
        render_student_name_matches(s['archive'],s['inputs'],s['scan'],[2026],s['unmatched'],show_tables=False)
    assert options['OPD students needing an OASIS match']==['learner, alphaa','learner, betaa']
    assert not any(kind=='dataframe' for kind,_ in st.messages)


def test_clear_focus_restores_normal_missing_queue(scenario):
    s=scenario;f=make_student_review_focus(s['bundle'],s['row'])
    st.reset(secrets=s['secrets'],state={FOCUS_KEY:f},values={MATCH_P+'clear_diagnostic_focus':True});login_for_test()
    with pytest.raises(StopRun):
        render_student_name_matches(s['archive'],s['inputs'],s['scan'],[2026],s['unmatched'],show_tables=False)
    assert FOCUS_KEY not in st.session_state


def test_password_lock_clears_diagnostic_state(scenario):
    st.reset(secrets=scenario['secrets'],state={FOCUS_KEY:make_student_review_focus(scenario['bundle'],scenario['row']),P+'preceptor':'private'})
    login_for_test();lock_evaluation_records()
    assert FOCUS_KEY not in st.session_state and P+'preceptor' not in st.session_state


def test_unchecked_sources_are_not_shown_as_zero(scenario):
    s=scenario;s['row']['clinical_forms_submitted']=None;s['row']['hp_forms_submitted']=None
    show(s,{P+'show':True})
    assert any('Not checked / not verified' in text for _,text in st.messages)


def test_diagnostic_names_not_added_to_report_zip(scenario):
    s=scenario
    raw,_=teaching_build_zip(dict(s['scan'],assessment_completion=s['bundle']),[2026])
    with ZipFile(BytesIO(raw)) as z:
        for name in z.namelist():
            if name.endswith('.docx'):content=all_doc_text(z.read(name))
            elif name.endswith(('.csv','.txt','.json')):content=z.read(name).decode('utf-8-sig')
            else:continue
            assert 'Learner, Alphaa' not in content and 'Learner, Betaa' not in content
            assert 'id-alpha' not in content and 'id-beta' not in content
            assert 'CONFIDENTIAL_COMMENT' not in content
