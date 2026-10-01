"""Student identity/date consistency regression tests. All data are invented."""
from copy import deepcopy
from datetime import date
from io import BytesIO
import csv
import json
from zipfile import ZipFile
from unittest.mock import patch
import pytest

from helpers import st, login_for_test
from test_learner_reach import scan_cells
from test_assessment_completion import record, raw_csv, NAME
from test_reporting_dates import all_doc_text
from test_documented_completion import make_inputs, run
from schedule_app.services.assessment_completion import (
    build_completion_bundle, load_completion_inputs, completion_rows, completion_signature,
    unverified_bundle, ASSESSMENT_VERSION, COLUMNS,
)
from schedule_app.services.assessment_progress import AssessmentIdentityResolver
from schedule_app.services.student_cohort import (
    group_student_assignments, student_cohort_counts, validate_student_cohort_counts,
)
from schedule_app.services.student_continuity import (
    student_continuity_counts, student_continuity_summary, continuity_period_note,
    require_student_continuity_data,
)
from schedule_app.services.student_name_matching import StudentNameMatcher
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks, student_name_key
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.educational_time import teaching_time_rows, TIME_CSV_COLUMNS
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.reports.chair_summary import teaching_chair_summary_data
from schedule_app.reports.teaching_batch import TeachingReportBatch

TODAY = date(2026,10,1)
@pytest.fixture(autouse=True)
def deterministic_clock():
    with patch('schedule_app.services.assessment_progress.teaching_local_today', return_value=TODAY), \
         patch('schedule_app.sections.assessment_completion.teaching_local_today', return_value=TODAY):
        yield


def shared(s, **kwargs):
    b,u,r=run(s,**kwargs)
    return dict(s['view'],assessment_completion=b),b,r


def five_then_three():
    # Five learners in full schedule; only three have worked by Oct 1.
    # Alpha's program-only variant is a duplicate; Betaa is a confirmed spelling alias.
    past='Learner, Alpha; Learner, Alpha (MD); Learner, Betaa; Learner, Gamma'
    future='Learner, Delta; Learner, Epsilon'
    cells={('NYES',f'{col}6'):f'{NAME} ~ {past}' for col in 'BCD'}
    cells.update({('NYES',f'{col}6'):f'{NAME} ~ {future}' for col in 'FGH'})
    cells[('NYES','E6')]=f'{NAME} ~ '
    scan,repo,archive,secrets,_=scan_cells(cells,start=date(2026,9,28))
    view=teaching_filter_date_range(scan,ReportingPeriod('Example period',date(2026,9,28),date(2026,10,4)))
    links=GitHubPreceptorOASISLinks(archive);links.save_username(NAME,'avery',expected=links.load())
    ids=GitHubStudentAssessmentLinks(archive)
    ids.save_link('Learner, Betaa','id-beta',expected=ids.load())
    GitHubOASISStudentEvaluations(archive).save(raw_csv([
        record(student='Learner, Alpha',submitted='2026-09-30'),
        record('beta',sid='id-beta',student='Learner, Beta',submitted='2026-09-30'),
    ]))
    inputs=load_completion_inputs(archive,view,[2026])
    b,_=build_completion_bundle(inputs,view,[2026],as_of=TODAY)
    return dict(view=view,inputs=inputs,repo=repo,archive=archive,secrets=secrets),dict(view,assessment_completion=b),b


def test_designation_duplicate_one_person_in_both_metrics():
    s=make_inputs(('Learner, Alpha','Learner, Alpha (MD)'))
    view,b,r=shared(s)
    assert student_continuity_counts(view,NAME,2026)['unique_students']==r['eligible_students']==1
    assert student_continuity_counts(s['view'],NAME,2026)['unique_students']==1
    assert r['clinical_completion_pct']==100


def test_confirmed_typo_and_exact_oasis_name_union_dates_not_students():
    s=make_inputs(('Learner, Alpha','Learner, Alphaa'))
    store=GitHubStudentAssessmentLinks(s['archive'])
    s['inputs']['student_links']=store.save_link('Learner, Alphaa','id-alpha',expected=store.load())
    view,b,r=shared(s)
    assert student_continuity_counts(view,NAME,2026)['unique_students']==r['eligible_students']==1


def test_confirmed_spelling_alias_collects_different_days_and_shifts():
    s=make_inputs(('Learner, Alpha',))
    s['inputs']['assignments']=[{**s['inputs']['assignments'][0], 'student':name,'date':day,'shift':'AM'}
        for name,day in [('Learner, Alpha','2026-09-07'),('Learner, Alphaa (MD)','2026-09-08'),('Learner, Alphaa','2026-09-09')]]
    store=GitHubStudentAssessmentLinks(s['archive'])
    s['inputs']['student_links']=store.save_link('Learner, Alphaa (MD)','id-alpha',expected=store.load())
    view,b,r=shared(s)
    assert r['unique_students']==r['unique_students_3plus_days']==r['eligible_students']==1
    assert student_continuity_counts(view,NAME,2026)=={'unique_students':1,'unique_students_3plus_days':1}


def test_unambiguous_saved_program_sibling_reused_without_catalog_mutation():
    entries={student_name_key('Learner, Alphaa (MD)'):{'external_id':'s1'}}
    original=deepcopy(entries)
    m=StudentNameMatcher({},entries)
    assert m.resolve('Learner, Alphaa').candidate_ids==('s1',)
    assert entries==original


def test_conflicting_program_specific_saved_ids_not_guessed_for_bare_name():
    entries={student_name_key('Learner, Alpha (MD)'):{'external_id':'s1'},
             student_name_key('Learner, Alpha (PA)'):{'external_id':'s2'}}
    m=StudentNameMatcher({},entries)
    assert m.resolve('Learner, Alpha (MD)').candidate_ids==('s1',)
    assert m.resolve('Learner, Alpha (PA)').candidate_ids==('s2',)
    assert len(m.resolve('Learner, Alpha').candidate_ids)==2


def test_existing_different_source_id_does_not_inherit_unrelated_saved_id():
    m=StudentNameMatcher({'Learner, Alpha':['s2']},{student_name_key('Learner, Alpha (MD)'):{'external_id':'s1'}})
    assert len(m.resolve('Learner, Alpha').candidate_ids)==2
    assert m.resolve('Learner, Alpha (MD)').candidate_ids==('s1',)


def test_two_distinct_learners_with_identical_dates_remain_two():
    view,b,r=shared(make_inputs(('Learner, Alpha','Learner, Beta')))
    assert r['unique_students']==r['eligible_students']==2


def test_weekends_and_multiple_work_types_combine_in_same_cohort():
    s=make_inputs(('Learner, Alpha',))
    a=s['inputs']['assignments'][0]
    s['inputs']['assignments']=[{**a,'date':day,'shift':shift,'work_type':kind}
        for day,shift,kind in [('2026-09-11','PM','Academic Pediatrics'),('2026-09-12','AM','Ward A'),('2026-09-13','PM','Ward A')]]
    view,b,r=shared(s)
    assert r['unique_students_3plus_days']==r['eligible_students']==1


def test_three_shifts_over_two_days_not_three_days():
    view,b,r=shared(make_inputs(('Learner, Alpha',)))
    assert r['eligible_students']==1 and r['unique_students_3plus_days']==0


def test_five_in_full_schedule_but_three_by_cutoff_match_everywhere():
    s,view,b=five_then_three();r=b['rows'][0]
    assert student_continuity_counts(s['view'],NAME,2026)['unique_students']==5
    assert r['unique_students']==r['unique_students_3plus_days']==r['eligible_students']==3
    assert student_continuity_counts(view,NAME,2026)=={'unique_students':3,'unique_students_3plus_days':3}
    assert r['clinical_completion_pct']==66.7
    summary=teaching_chair_summary_data(view,[2026])[0]['named_preceptors'][0]
    assert summary['unique_students_3plus_days']==3


def test_identity_cutoff_correction_does_not_change_hours_or_reach():
    s,view,_=five_then_three()
    before=teaching_time_rows(s['view'],[2026])[0];after=teaching_time_rows(view,[2026])[0]
    for key in ('total_scheduled_availability_hours','educational_hours','learner_reach_pct','scheduled_shifts','teaching_shifts'):
        assert before[key]==after[key]
    assert after['total_scheduled_availability_hours']==28 and after['educational_hours']==24


@pytest.mark.parametrize('minimum',[1,2,3,4,5])
def test_threshold_inequality_is_only_applied_for_three_or_fewer(minimum):
    s,view,b=five_then_three()
    changed,_=build_completion_bundle(s['inputs'],s['view'],[2026],minimum_shifts=minimum,as_of=TODAY)
    r=changed['rows'][0]
    assert r['unique_students_3plus_days']==3
    assert r['eligible_students']==(3 if minimum<=3 else 0)
    completion_rows(changed,s['view'],2026)


@pytest.mark.parametrize('cutoff,days3,eligible',[(date(2026,9,28),0,0),(date(2026,9,29),0,0),(date(2026,9,30),3,3),(TODAY,3,3)])
def test_cutoff_change_recalculates_both_metrics_inclusive(cutoff,days3,eligible):
    s,_,_=five_then_three()
    b,_=build_completion_bundle(s['inputs'],s['view'],[2026],as_of=cutoff)
    r=b['rows'][0]
    assert r['unique_students_3plus_days']==days3 and r['eligible_students']==eligible
    assert student_continuity_counts(dict(s['view'],assessment_completion=b),NAME,2026)['unique_students_3plus_days']==days3


def test_future_only_preceptor_has_zero_student_cohort_not_zero_hours():
    s,_,_=five_then_three()
    b,_=build_completion_bundle(s['inputs'],s['view'],[2026],as_of=date(2026,9,27))
    v=dict(s['view'],assessment_completion=b)
    assert b['rows'][0]['unique_students']==b['rows'][0]['eligible_students']==0
    assert student_continuity_counts(v,NAME,2026)['unique_students']==0
    assert teaching_time_rows(v,[2026])[0]['educational_hours']==24
    assert 'No dates' in continuity_period_note(v,2026)


def test_absent_student_stays_counted_without_external_id():
    view,b,r=shared(make_inputs(('New, Jordan','New, Jordan (MD)')))
    assert r['unique_students']==r['eligible_students']==1 and r['clinical_completion_pct']==0
    assert not r['student_counts_status'].startswith('Not checked')


def test_ambiguous_oasis_name_provisional_in_both_student_sections():
    s=make_inputs(('Learner, Alpha',),[record(),record('two',sid='different')])
    view,b,r=shared(s)
    assert r['unique_students']==1 and r['clinical_completion_pct']==0
    assert r['student_counts_status'].startswith('Provisional:')
    assert r['assessment_status'].startswith('Provisional:')


def test_missing_preceptor_username_does_not_prevent_identity_counts():
    s=make_inputs();s['inputs']['catalog']['entries']={}
    view,b,r=shared(s)
    assert r['unique_students']==r['eligible_students']==2 and r['clinical_completion_pct'] is None
    assert student_continuity_counts(view,NAME,2026)['unique_students']==2


def test_unloaded_inputs_show_not_checked_not_old_five_student_count():
    s,_,_=five_then_three()
    b=unverified_bundle(s['view'],[2026],as_of=TODAY)
    v=dict(s['view'],assessment_completion=b)
    assert student_continuity_counts(v,NAME,2026)['unique_students'] is None
    raw=teaching_make_docx(NAME,s['view']['monthly'],v)
    assert 'Unique students assigned: Not checked' in all_doc_text(raw)


@pytest.mark.parametrize('field,value', [('unique_students',2),('unique_students_3plus_days',4),('eligible_students',2),('unique_students',None),('unique_students_3plus_days',-1)])
def test_tampered_inconsistent_cohort_refused_before_document(field,value):
    s,view,b=five_then_three();bad=deepcopy(b);bad['rows'][0][field]=value
    with pytest.raises(OPDArchiveError):
        teaching_build_zip(dict(s['view'],assessment_completion=bad),[2026])


def test_same_cohort_window_verified_even_when_asof_exceeds_report_end():
    s=make_inputs();view,b,r=shared(s,as_of=TODAY)
    assert student_continuity_summary(view,NAME,2026)['student_counts_end_date']=='2026-09-13'
    bad=deepcopy(b);bad['rows'][0]['assessment_end_date']='2026-09-12'
    with pytest.raises(OPDArchiveError):student_continuity_counts(dict(s['view'],assessment_completion=bad),NAME,2026)


def test_changed_exclusions_old_bundle_rejected():
    s,view,b=five_then_three();v=dict(view,student_exclusions_signature='changed')
    with pytest.raises(OPDArchiveError):student_continuity_counts(v,NAME,2026)


def test_old_schema_requires_refresh():
    s,view,b=five_then_three();old=deepcopy(s['view']);old['student_continuity_version']=1
    with pytest.raises(OPDArchiveError):require_student_continuity_data(old)
    oldb=deepcopy(b);oldb['version']=4
    with pytest.raises(OPDArchiveError):completion_rows(oldb,s['view'],2026)


def test_raw_groups_never_mutated_and_no_github_writes_from_calculation():
    s,view,b=five_then_three();old=deepcopy(s['view']);tree=deepcopy(s['repo'].tree)
    teaching_build_zip(view,[2026]);assert old==s['view'] and tree==s['repo'].tree


def test_batch_and_standalone_rows_agree():
    s,view,b=five_then_three();batch=TeachingReportBatch(view,[2026])
    assert batch.preceptor_rows(view,NAME,[2026])[0][0]['unique_students']==3
    assert batch.chair_summaries(view,[2026])[0]['named_preceptors'][0]['unique_students']==3


def test_csv_and_both_documents_publish_same_identity_cutoff_without_identifiers():
    s,view,b=five_then_three();raw,_=teaching_build_zip(view,[2026])
    with ZipFile(BytesIO(raw)) as z:
        overall=list(csv.DictReader(z.read('preceptor_teaching_summary.csv').decode('utf-8-sig').splitlines()))[0]
        complete=list(csv.DictReader(z.read('preceptor_student_assessment_completion.csv').decode('utf-8-sig').splitlines()))[0]
        assert overall['unique_students']==complete['unique_students']=='3'
        assert 'unique_students_3plus_days' not in overall and 'unique_students_3plus_days' not in complete
        assert overall['eligible_students']==complete['eligible_students']=='3'
        assert overall['student_counts_end_date']==complete['assessment_end_date']=='2026-10-01'
        for name in z.namelist():
            if name.endswith('.docx'):
                text=all_doc_text(z.read(name))
                assert 'Student-count dates: 2026-09-28 through 2026-10-01' in text
                assert '2 / 3 (66.7%)' in text
                assert 'Learner, Alpha' not in text and 'id-alpha' not in text
            elif name.endswith('.csv'):
                text=z.read(name).decode('utf-8-sig')
                assert 'Learner, Alpha' not in text and 'id-alpha' not in text


def test_export_headers_preserve_old_fields_then_append_context():
    assert TIME_CSV_COLUMNS[:10]==('preceptor_name','academic_year','total_scheduled_availability_hours','educational_hours','learner_reach_pct','unique_students','months_with_students','months_scheduled','scheduled_shifts','teaching_shifts')
    assert COLUMNS[-6:]==('unique_students','student_counts_status','clinical_unique_students_assessed_total','hp_unique_students_assessed_total','either_unique_students_assessed_total','total_assessments_status')


def test_split_dates_across_standard_years_never_sum_continuity_counts():
    s,view,b=five_then_three()
    # A separate year can't use a bundle built for only this custom period.
    mismatched=deepcopy(view);mismatched['reporting_period']['start_date']='2025-07-01'
    with pytest.raises(OPDArchiveError):student_continuity_counts(mismatched,NAME,2026)


def test_unknown_slot_rejected_not_counted():
    resolver=AssessmentIdentityResolver({'name_ids':{},'student_names':{}},{})
    a={'preceptor_name':NAME,'student':'Student','date':'2026-09-30','shift':'EVENING'}
    with pytest.raises(OPDArchiveError):group_student_assignments([a],resolver,[NAME],'2026-09-01','2026-10-01')
