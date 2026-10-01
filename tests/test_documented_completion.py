"""Regression coverage for documented completion and the independent cutoff.
All names, identifiers, OPDs and GitHub responses are invented test fixtures.
"""
from copy import deepcopy
from datetime import date
from io import BytesIO
from unittest.mock import patch
from zipfile import ZipFile
import pytest
from helpers import st, login_for_test, StopRun
from test_learner_reach import scan_cells
from test_assessment_completion import NAME, record, raw_csv
from test_reporting_dates import all_doc_text
from schedule_app.services.assessment_completion import (
    build_completion_bundle, load_completion_inputs, completion_display, completion_signature,
    unverified_bundle, completion_rows, ASSESSMENT_VERSION,
)
from schedule_app.services.assessment_progress import (
    assessment_as_of, AssessmentIdentityResolver, MATCHED, ABSENT, REVIEW, plausible_name_difference,
)
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.educational_time import teaching_time_rows
from schedule_app.services.student_continuity import student_continuity_counts
from schedule_app.services.student_name_review import student_name_review, oasis_student_choices, name_only_student_choices
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.sections.student_name_matches import render_student_name_matches, P as MATCH_P
from schedule_app.sections.assessment_completion import render_assessment_completion, P
from schedule_app.sections.pts_navigation import preserve_pts_preferences
from schedule_app.services.evaluation_access import lock_evaluation_records

TODAY = date(2026, 10, 1)
ASOF = date(2026, 9, 13)
NEW = 'New, Jordan'

@pytest.fixture(autouse=True)
def clock():
    # Deterministic test date only; production uses America/New_York local date.
    with patch('schedule_app.services.assessment_progress.teaching_local_today', return_value=TODAY), \
         patch('schedule_app.sections.assessment_completion.teaching_local_today', return_value=TODAY):
        yield


def make_inputs(students=('Learner, Alpha', NEW), records=None, dates=None):
    text='; '.join(students)
    cells={('NYES','B6'):f'{NAME} ~ {text}', ('NYES','B8'):f'{NAME} ~ {text}',
           ('NYES','C6'):f'{NAME} ~ {text}', ('NYES','D6'):f'{NAME} ~ '}
    scan, repo, archive, secrets, raw=scan_cells(cells)
    view=teaching_filter_date_range(scan, ReportingPeriod('Demo', date(2026,9,7), date(2026,9,13)))
    links=GitHubPreceptorOASISLinks(archive)
    links.save_username(NAME,'avery',expected=links.load())
    if records is not False:
        GitHubOASISStudentEvaluations(archive).save(raw_csv(records or [record()]))
    inputs=load_completion_inputs(archive,view,[2026])
    return dict(view=view, inputs=inputs, archive=archive, repo=repo, secrets=secrets)


def run(s, **kw):
    b,u=build_completion_bundle(s['inputs'],s['view'],[2026],as_of=kw.pop('as_of',ASOF),**kw)
    return b,u,b['rows'][0]


def test_never_evaluated_student_stays_in_denominator_and_percentage_calculates():
    s=make_inputs(); b,u,r=run(s)
    assert r['eligible_students']==2 and r['clinical_students_evaluated']==1
    assert r['clinical_completion_pct']==50 and r['hp_completion_pct']==0
    assert r['assessment_status']=='Calculated'
    assert r['students_without_oasis_name_record']==1
    assert r['either_students_without_assessment']==1
    assert u[0]['match_category']==ABSENT


def test_missing_record_not_in_mandatory_queue():
    s=make_inputs(); b,u,r=run(s)
    review=student_name_review(s['inputs'],s['view'],[2026],u,as_of=ASOF)
    assert not review['missing'] and 'new, jordan' in review['absent']
    st.reset(secrets=s['secrets']);login_for_test()
    render_student_name_matches(s['archive'],s['inputs'],s['view'],[2026],u,show_tables=False,as_of=ASOF)
    assert not any(k==MATCH_P+'student_choice' for _,k in st.widget_keys)
    assert not any(kind=='dataframe' for kind,_ in st.messages)
    assert any('No confirmation is needed' in t for _,t in st.messages)


def test_8_eligible_6_documented_is_75_not_an_error():
    names=tuple(f'Learner, Matched{i}' for i in range(6))+('New, Jordan','New, Cameron')
    rows=[record(str(i),sid=f'id-{i}',student=names[i]) for i in range(6)]
    _,u,r=run(make_inputs(names, rows))
    assert r['eligible_students']==8 and r['clinical_completion_pct']==75
    assert r['assessment_status']=='Calculated' and len(u)==2


def test_no_forms_for_preceptor_is_documented_zero_not_not_verified():
    _,_,r=run(make_inputs(records=[record(email='another@example.edu')]))
    assert r['clinical_completion_pct']==0 and r['eligible_students']==2
    assert r['either_students_without_assessment']==2


def test_no_archive_is_not_checked_not_zero():
    _,_,r=run(make_inputs(records=False))
    assert r['clinical_completion_pct'] is None
    assert completion_display(r,'clinical')=='Not checked'


def test_missing_username_is_not_checked():
    s=make_inputs();s['inputs']['catalog']['entries']={}
    _,_,r=run(s)
    assert r['clinical_completion_pct'] is None and 'username missing' in r['assessment_status']


def test_invalid_provider_form_metadata_still_not_verified():
    _,_,r=run(make_inputs(records=[record(sid='')]))
    assert r['clinical_completion_pct'] is None and 'source metadata' in r['assessment_status']


def test_failed_load_is_not_silently_empty_source():
    s=make_inputs();st.reset(secrets=s['secrets'], values={P+'refresh':True});login_for_test()
    with patch('schedule_app.sections.assessment_completion.load_completion_inputs',side_effect=OPDArchiveError('Unreadable archive')):
        b=render_assessment_completion(s['archive'],s['view'],[2026],manage_students=False,show_tables=False)
    assert b['rows'][0]['clinical_completion_pct'] is None
    assert 'Not checked' in b['rows'][0]['assessment_status']


def test_typo_only_suggests_never_credits_and_is_numeric_provisional():
    s=make_inputs(('Learner, Alphaa', NEW))
    b,u,r=run(s)
    assert r['clinical_students_evaluated']==0 and r['clinical_completion_pct']==0
    assert r['assessment_status'].startswith('Provisional:')
    assert completion_display(r,'clinical')=='0 / 2 (0.0%)*'
    review=student_name_review(s['inputs'],s['view'],[2026],u,as_of=ASOF)
    assert set(review['missing'])=={'learner, alphaa'}
    assert set(review['absent'])=={'new, jordan'}


def test_confirmed_typo_recalculates_while_absent_student_needs_no_link():
    s=make_inputs(('Learner, Alphaa', NEW))
    store=GitHubStudentAssessmentLinks(s['archive'])
    s['inputs']['student_links']=store.save_link('Learner, Alphaa','id-alpha',expected=store.load())
    b,u,r=run(s)
    assert r['clinical_completion_pct']==50 and r['assessment_status']=='Calculated'
    assert not student_name_review(s['inputs'],s['view'],[2026],u,as_of=ASOF)['missing']
    assert all(b'Learner, Alphaa' not in content for content in s['repo'].tree.values())


def test_absent_student_later_evaluated_updates_without_manual_link():
    s=make_inputs();assert run(s)[2]['clinical_completion_pct']==50
    GitHubOASISStudentEvaluations(s['archive']).save(raw_csv([record('new',sid='id-new',student=NEW)]))
    s['inputs']=load_completion_inputs(s['archive'],s['view'],[2026])
    _,u,r=run(s)
    assert r['clinical_completion_pct']==100 and not u
    assert not s['inputs']['student_links']['entries']


def test_exact_ambiguous_name_is_provisional_not_guessed():
    s=make_inputs(('Learner, Alpha',),[record(),record('2',sid='other-id')])
    _,_,r=run(s)
    assert r['eligible_students']==1 and r['clinical_completion_pct']==0
    assert r['assessment_status'].startswith('Provisional:')


def test_no_automatic_false_given_name_match():
    assert not plausible_name_difference('Student, Eva','Student, Ava')
    assert not plausible_name_difference('Student, Alpha','Student, Beta')
    assert plausible_name_difference('Smith, Jordann','Smith, Jordan')
    assert plausible_name_difference('Jordan Smith','Smith, Jordan')


def test_absent_designation_variants_are_one_denominator_member():
    s=make_inputs((NEW,NEW+' (MD)'))
    _,_,r=run(s)
    assert r['eligible_students']==1 and r['clinical_completion_pct']==0


def test_cutoff_excludes_future_shifts_without_changing_time_report():
    s=make_inputs()
    for r in s['inputs']['assignments']:
        if r['student']==NEW:
            r['date']='2026-09-11' if r['shift']=='AM' else '2026-09-12'
    before=teaching_time_rows(s['view'],[2026])
    b,_,r=run(s,as_of=date(2026,9,10))
    assert r['eligible_students']==1 and r['clinical_completion_pct']==100
    after=teaching_time_rows(dict(s['view'],assessment_completion=b),[2026])
    for field in ('educational_hours','total_scheduled_availability_hours','learner_reach_pct','scheduled_shifts','teaching_shifts'):
        assert before[0][field]==after[0][field]
    assert after[0]['unique_students']==1 and after[0]['student_counts_end_date']=='2026-09-10'
    assert r['report_end_date']=='2026-09-13' and r['assessment_end_date']=='2026-09-10'


def test_last_cutoff_day_submission_included_next_day_excluded():
    s=make_inputs(records=[record(submitted='2026-09-10 23:59:59'),
                          record('n',sid='new',student=NEW,submitted='2026-09-11 00:00:00')])
    assert run(s,as_of=date(2026,9,10))[2]['clinical_completion_pct']==50
    assert run(s,as_of=date(2026,9,11))[2]['clinical_completion_pct']==100


def test_cutoff_before_reporting_period_is_no_eligible_not_zero():
    _,_,r=run(make_inputs(),as_of=date(2026,9,1))
    assert r['eligible_students']==0 and r['clinical_completion_pct'] is None
    assert completion_display(r,'clinical')=='No eligible students'


def test_cutoff_after_period_uses_period_end():
    _,_,r=run(make_inputs(),as_of=TODAY)
    assert r['assessment_end_date']=='2026-09-13'


def test_default_cutoff_today_but_future_explicit_rejected():
    assert assessment_as_of()==TODAY
    with pytest.raises(OPDArchiveError):assessment_as_of(date(2027,1,1))
    with pytest.raises(OPDArchiveError):assessment_as_of('not-date')


def test_threshold_and_three_day_count_still_independent():
    s=make_inputs();b,_,r=run(s,minimum_shifts=4)
    assert r['eligible_students']==0
    assert student_continuity_counts(s['view'],NAME,2026)['unique_students_3plus_days']==0


def test_repeated_forms_do_not_inflate_percentage_and_no_assessment_counts():
    _,_,r=run(make_inputs(records=[record(),record('2'),record('3')]))
    assert r['clinical_forms_submitted']==3 and r['clinical_students_evaluated']==1
    assert r['clinical_students_without_assessment']==1 and r['clinical_completion_pct']==50


def test_both_forms_union_cannot_exceed_100():
    from schedule_app.services.assessment_completion import HP
    _,_,r=run(make_inputs(records=[record(),record('2',form=HP)]))
    assert r['either_students_evaluated']==1 and r['either_completion_pct']==50


def test_period_wide_feedback_summary_not_sliced_by_as_of():
    s=make_inputs();s['inputs']['summaries']['2026']={'rows_by_id':{'avery':{'evaluation_count':9}}}
    b,_,r=run(s,as_of=date(2026,9,8))
    assert r['student_feedback_evaluations']==9 and r['clinical_students_evaluated']==0


def test_cutoff_changes_bundle_signature_and_rejects_tampered_bundle():
    s=make_inputs();b,_,_=run(s);b2,_,_=run(s,as_of=date(2026,9,10))
    assert completion_signature(b)!=completion_signature(b2)
    bad=deepcopy(b);bad['rows'][0]['assessments_as_of']='2026-09-09'
    with pytest.raises(OPDArchiveError):completion_rows(bad,s['view'],2026)


def test_changing_cutoff_reuses_loaded_sources():
    s=make_inputs();st.reset(secrets=s['secrets'],values={P+'refresh':True});login_for_test()
    with patch('schedule_app.sections.assessment_completion.load_completion_inputs',return_value=s['inputs']) as load:
        first=render_assessment_completion(s['archive'],s['view'],[2026],manage_students=False,show_tables=False)
        st.values={P+'as_of':date(2026,9,8)}
        second=render_assessment_completion(s['archive'],s['view'],[2026],manage_students=False,show_tables=False)
    assert load.call_count==1
    assert first['rows'][0]['clinical_completion_pct']==50 and second['rows'][0]['clinical_completion_pct']==0


def test_cutoff_survives_page_preference_transfer_and_lock_clears_it():
    s=make_inputs();st.reset(secrets=s['secrets'],state={P+'as_of':ASOF});login_for_test()
    preserve_pts_preferences();assert st.session_state[P+'as_of']==ASOF
    lock_evaluation_records();assert P+'as_of' not in st.session_state


def test_missing_record_optional_correction_not_required_or_preselected():
    s=make_inputs();b,u,r=run(s)
    st.reset(secrets=s['secrets'],values={MATCH_P+'correct_absent':True});login_for_test()
    render_student_name_matches(s['archive'],s['inputs'],s['view'],[2026],u,as_of=ASOF,show_tables=False)
    assert any(label=='OPD name to correct (optional)' for label,_ in st.widget_keys)
    assert not s['inputs']['student_links']['entries']


def test_reports_csv_and_notes_show_progress_no_student_identifiers():
    s=make_inputs();b,u,r=run(s)
    before=s['repo'].write_count
    raw,_=teaching_build_zip(dict(s['view'],assessment_completion=b),[2026])
    assert s['repo'].write_count==before
    with ZipFile(BytesIO(raw)) as z:
        assert z.testzip() is None
        for path in z.namelist():
            if path.endswith('.docx'):
                text=all_doc_text(z.read(path))
                assert 'Documented assessment completion' in text and '1 / 2 (50.0%)' in text
                assert 'Assessments as of: 2026-09-13' in text
                assert 'not overdue' in text
            elif path.endswith('.csv') or path=='Report_Notes.txt':
                text=z.read(path).decode('utf-8-sig')
            else:continue
            assert NEW not in text and 'Learner, Alpha' not in text and 'id-alpha' not in text
        csv_text=z.read('preceptor_student_assessment_completion.csv').decode('utf-8-sig')
        assert 'assessments_as_of' in csv_text and 'clinical_students_without_assessment' in csv_text


def test_provisional_marker_appears_in_individual_report():
    s=make_inputs(('Learner, Alphaa',NEW));b,_,_=run(s)
    raw=teaching_make_docx(NAME,s['view']['monthly'],dict(s['view'],assessment_completion=b))
    text=all_doc_text(raw)
    assert '0 / 2 (0.0%)*' in text and 'Provisional:' in text


def test_handoff_only_archive_can_confirm_no_target_assessments_on_file():
    _,_,r=run(make_inputs(records=[record(form='*PEDS Handoff')]))
    assert r['eligible_students']==2 and r['clinical_completion_pct']==0
    assert r['hp_completion_pct']==0 and r['assessment_status']=='Calculated'


def test_weekend_shift_can_reach_threshold_by_as_of_date():
    s=make_inputs((NEW,))
    s['inputs']['assignments']=[{**s['inputs']['assignments'][0], 'date':d, 'shift':shift}
        for d,shift in [('2026-09-11','PM'),('2026-09-12','AM'),('2026-09-13','AM')]]
    assert run(s,as_of=date(2026,9,12))[2]['eligible_students']==0
    assert run(s,as_of=date(2026,9,13))[2]['eligible_students']==1


def test_cutoff_does_not_persist_or_invent_student_ids():
    s=make_inputs();before=deepcopy(s['repo'].tree)
    for cutoff in (date(2026,9,8),date(2026,9,10),ASOF):
        run(s,as_of=cutoff)
    assert s['repo'].tree==before and not s['inputs']['student_links']['entries']
