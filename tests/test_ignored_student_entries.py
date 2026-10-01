"""Explicit OPD exclusions: encrypted persistence, complete metrics, protected UI."""
from copy import deepcopy
from datetime import date
from io import BytesIO
from zipfile import ZipFile
from unittest.mock import patch
import json
import pytest

from helpers import st, login_for_test, StopRun, FakeGitHub, secret_settings
from test_learner_reach import scan_cells
from test_assessment_completion import record, raw_csv, NAME
from schedule_app.services.opd_archive import OPDArchiveError, GitHubOPDArchive, get_opd_archive_config
from schedule_app.services.ignored_student_entries import (
    GitHubIgnoredStudentEntries, ignored_entry_key, exclusion_signature,
    EMPTY_EXCLUSION_SIGNATURE, require_matching_exclusions,
)
from schedule_app.services.teaching_analysis import (
    teaching_scan_archives, teaching_filter_date_range, teaching_annual_rows,
)
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.learner_reach import learner_reach_rows
from schedule_app.services.student_continuity import student_continuity_counts
from schedule_app.services.assessment_completion import load_completion_inputs, build_completion_bundle
from schedule_app.services.student_name_review import student_name_review
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.evaluation_access import lock_evaluation_records
from schedule_app.services.teaching_validation import validate_teaching_report, TeachingConflictError
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.sections import ignored_student_entries as ui

@pytest.fixture
def data():
    cells = {('NYES','B6'): f'{NAME} ~ Learner, Alpha; midcycle feedback',
             ('NYES','B8'): f'{NAME} ~ Learner, Alpha',
             ('NYES','C6'): f'{NAME} ~ Learner, Alpha',
             ('NYES','C8'): f'{NAME} ~ midcycle feedback',
             ('NYES','D6'): f'{NAME} ~ midcycle feedback',
             ('NYES','G6'): f'{NAME} ~ midcycle feedback',
             ('NYES','H6'): f'{NAME} ~ '}
    scan, repo, archive, secrets, raw = scan_cells(cells)
    service = GitHubIgnoredStudentEntries(archive)
    return dict(scan=scan,repo=repo,archive=archive,secrets=secrets,raw=raw,service=service)


def exclude(data, names=('midcycle feedback',)):
    svc = data['service']
    cat = svc.ignore_entries(list(names), expected=svc.load())
    inventory=[]
    scan=teaching_scan_archives(data['archive'], ignored_student_entries=cat['entries'], student_entry_collector=inventory.extend)
    return scan, cat, inventory


def view(scan):
    return teaching_filter_date_range(scan, ReportingPeriod('26-27',date(2026,9,7),date(2026,9,13)))


def test_catalog_is_empty_initially_does_not_write(data):
    before=data['repo'].write_count
    c=data['service'].load()
    assert c['entries']=={} and c['sha'] is None
    assert data['repo'].write_count==before


def test_encrypted_roundtrip_and_source_unchanged(data):
    before=data['archive'].load(date(2026,9,7))['raw']
    _,cat,_=exclude(data)
    token=data['repo'].tree[data['service'].path]
    assert b'midcycle feedback' not in token
    payload=json.loads(data['archive'].config.cipher().decrypt(token))
    assert payload['entries']['midcycle feedback']['student_entry']=='midcycle feedback'
    assert GitHubIgnoredStudentEntries(data['archive']).load()['entries']==cat['entries']
    assert data['archive'].load(date(2026,9,7))['raw']==before


def test_identical_ignore_does_not_rewrite(data):
    _,c,_=exclude(data); n=data['repo'].write_count
    result=data['service'].ignore_entries(['MIDCYCLE   FEEDBACK'],expected=c)
    assert result['sha']==c['sha'] and data['repo'].write_count==n


def test_restore_does_not_delete_name_links(data):
    links=GitHubStudentAssessmentLinks(data['archive'])
    link=links.save_link('midcycle feedback','external-test',expected=links.load())
    _,c,_=exclude(data)
    restored=data['service'].restore_entries(['midcycle feedback'],expected=c)
    assert restored['entries']=={}
    assert links.load()['entries']==link['entries']


def test_stale_save_no_overwrite(data):
    old=data['service'].load();exclude(data)
    with pytest.raises(OPDArchiveError,match='another session'):
        data['service'].ignore_entries(['Other label'],expected=old)
    assert set(data['service'].load()['entries'])=={'midcycle feedback'}


def test_stale_restore_no_overwrite(data):
    _,old,_=exclude(data)
    data['service'].ignore_entries(['Other label'],expected=old)
    with pytest.raises(OPDArchiveError):
        data['service'].restore_entries(['midcycle feedback'],expected=old)
    assert len(data['service'].load()['entries'])==2


@pytest.mark.parametrize('value',[None,'',123,'x\nname','x\x00name','x'*257])
def test_invalid_manual_entries_not_written(data,value):
    n=data['repo'].write_count
    with pytest.raises(OPDArchiveError):
        data['service'].ignore_entries([value],expected=data['service'].load())
    assert data['repo'].write_count==n


def test_bad_ciphertext_does_not_become_empty(data):
    exclude(data)
    data['repo'].tree[data['service'].path]=b'bad ciphertext';data['repo']._commit()
    n=data['repo'].write_count
    with pytest.raises(OPDArchiveError): data['service'].load()
    assert data['repo'].write_count==n


def test_exclusion_is_exact_no_fuzzy_or_substring_matching():
    assert ignored_entry_key('  MIDCYCLE  feedback ')=='midcycle feedback'
    assert ignored_entry_key('Example, Jordan (MD)')!=ignored_entry_key('Example, Jordan')
    assert ignored_entry_key('Midcycle feedback today')!='midcycle feedback'
    assert exclusion_signature(['a','A'])==exclusion_signature(['a'])


def test_clinical_hours_retained_student_hours_removed(data):
    scan,_,_=exclude(data)
    row=teaching_annual_rows(scan,[2026])[0]
    assert row['recorded_clinical_hours']==28
    assert row['hours_with_students']==12
    assert row['educational_hours']==12
    assert row['learner_reach_pct']==42.9
    assert row['hours_without_students']==16


def test_two_names_same_shift_ignore_only_selected_student(data):
    scan,_,_=exclude(data)
    # Monday AM still has Alpha even though feedback was removed.
    monday=view(scan)
    day=teaching_filter_date_range(scan,ReportingPeriod('Monday',date(2026,9,7),date(2026,9,7)))
    assert teaching_annual_rows(day,[2026])[0]['hours_with_students']==8
    assert sum(r['no_of_shifts'] for r in scan['monthly'])==3


def test_inventory_preserves_excluded_label_for_ui_only(data):
    scan,_,inventory=exclude(data)
    assert 'midcycle feedback' in {r['student_entry'] for r in inventory}
    encoded=json.dumps(scan)
    assert 'midcycle feedback' not in encoded and 'Learner, Alpha' not in encoded
    assert sum(scan['excluded_student_entry_listings_by_date'].values())==4


def test_no_exclusions_preserves_metrics(data):
    a=data['scan'];b=teaching_scan_archives(data['archive'],ignored_student_entries=())
    assert a['monthly']==b['monthly'] and a['clinical_daily']==b['clinical_daily']


def test_restore_reproduces_original_metrics(data):
    scan,c,_=exclude(data)
    c=data['service'].restore_entries(['midcycle feedback'],expected=c)
    restored=teaching_scan_archives(data['archive'],ignored_student_entries=c['entries'])
    assert restored['clinical_daily']==data['scan']['clinical_daily']
    assert restored['monthly']==data['scan']['monthly']


def test_explicit_ignore_all_students_keeps_capacity_not_teaching(data):
    scan,_,_=exclude(data,('midcycle feedback','Learner, Alpha'))
    assert teaching_annual_rows(scan,[2026])==[]
    row=learner_reach_rows(scan,[2026])[0]
    assert row['recorded_clinical_hours']==28 and row['hours_with_students']==0
    assert row['learner_reach_pct']==0


def test_weekends_and_selected_dates(data):
    scan,_,_=exclude(data)
    weekend=teaching_filter_date_range(scan,ReportingPeriod('Weekend',date(2026,9,12),date(2026,9,13)))
    row=learner_reach_rows(weekend,[2026])[0]
    assert row['recorded_clinical_hours']==8 and row['hours_with_students']==0


def test_name_queue_and_eligibility_use_filtered_assignments(data):
    scan,_,_=exclude(data)
    usernames=GitHubPreceptorOASISLinks(data['archive'])
    usernames.save_username(NAME,'avery',expected=usernames.load())
    GitHubOASISStudentEvaluations(data['archive']).save(raw_csv([record()]))
    report=view(scan)
    inputs=load_completion_inputs(data['archive'],report,[2026])
    assert {r['student'] for r in inputs['assignments']}=={'Learner, Alpha'}
    bundle,unmatched=build_completion_bundle(inputs,report,[2026])
    assert not unmatched and bundle['rows'][0]['eligible_students']==1
    assert bundle['rows'][0]['clinical_completion_pct']==100
    names=student_name_review(inputs,report,[2026],unmatched)
    assert not names['missing']
    bundle4,_=build_completion_bundle(inputs,report,[2026],minimum_shifts=4)
    assert bundle4['rows'][0]['eligible_students']==0


def test_changed_list_blocks_old_teaching_scan_replay(data):
    exclude(data)
    with pytest.raises(OPDArchiveError,match='ignored-student list changed'):
        load_completion_inputs(data['archive'],view(data['scan']),[2026])


def test_changed_list_blocks_old_completion_inputs(data):
    inputs=load_completion_inputs(data['archive'],view(data['scan']),[2026])
    scan,_,_=exclude(data)
    with pytest.raises(OPDArchiveError,match='Ignored student entries changed'):
        build_completion_bundle(inputs,view(scan),[2026])


def test_opd_conflicts_not_hidden_by_exclusions():
    scan,repo,archive,*_=scan_cells({('WARD A','B6'):f'{NAME} ~ midcycle feedback',
                                   ('NYES','B6'):f'{NAME} ~ Learner, Alpha'})
    new=teaching_scan_archives(archive,ignored_student_entries=['midcycle feedback'])
    with pytest.raises(TeachingConflictError): validate_teaching_report(new,[2026])


def test_clinic_priority_retains_blank_clinic_not_nursery_student():
    scan,repo,archive,*_=scan_cells({('PSHCH_NURSERY','B6'):f'{NAME} ~ Learner, Alpha',
                                   ('NYES','B6'):f'{NAME} ~ midcycle feedback'})
    new=teaching_scan_archives(archive,ignored_student_entries=['midcycle feedback'])
    assert teaching_annual_rows(new,[2026])==[]
    row=learner_reach_rows(new,[2026])[0]
    assert row['recorded_clinical_hours']==4 and row['hours_with_students']==0


def test_reverse_order_filter():
    scan,repo,archive,*_=scan_cells({('NYES','B6'):'midcycle feedback ~ Teacher, Avery',
                                   ('NYES','C6'):'Learner A ~ Teacher, Avery'},order='Student ~ Preceptor')
    new=teaching_scan_archives(archive,'Student ~ Preceptor',ignored_student_entries=['midcycle feedback'])
    assert teaching_annual_rows(new,[2026])[0]['learner_reach_pct']==50


def test_reports_counts_and_notes_no_student_names(data):
    scan,_,_=exclude(data)
    raw,_=teaching_build_zip(view(scan),[2026])
    with ZipFile(BytesIO(raw)) as z:
        notes=z.read('Report_Notes.txt').decode()
        assert '4 source student-entry listing(s)' in notes
        assert 'midcycle feedback' not in notes
        assert 'Learner, Alpha' not in notes
        # CSV provides corrected hours rather than a hidden identity-only change.
        csv=z.read('preceptor_teaching_summary.csv').decode('utf-8-sig')
        assert 'midcycle feedback' not in csv
        # Keep internal IDs and names out of all exported notes/tables.
        for name in z.namelist():
            if name.endswith('.docx'):
                with ZipFile(BytesIO(z.read(name))) as doc:
                    xml=doc.read('word/document.xml').decode()
                assert 'midcycle feedback' not in xml and 'Learner, Alpha' not in xml


def setup_ui(data,values=None):
    st.reset(secrets=data['secrets'],values=values)
    login_for_test()
    inv=[]
    scan=teaching_scan_archives(data['archive'],student_entry_collector=inv.extend)
    st.session_state['teaching_scan']=scan
    st.session_state[ui.P+'scope']=(data['archive'].config.signature(),'Preceptor ~ Student')
    st.session_state[ui.P+'inventory']=inv
    st.session_state[ui.P+'catalog']=data['service'].load()


def test_locked_ui_no_archive_calls(data):
    st.reset(secrets=data['secrets'])
    with patch.object(GitHubIgnoredStudentEntries,'load',side_effect=AssertionError('must not read')):
        assert ui.render_ignored_student_entries(data['archive'],'Preceptor ~ Student') is None


def test_ui_exposes_names_only_and_restore(data):
    setup_ui(data)
    ui.render_ignored_student_entries(data['archive'],'Preceptor ~ Student')
    labels=[label for label,_ in st.widget_keys]
    assert 'OPD student entries to ignore' in labels
    assert not any('External ID' in s for s in labels)


def test_ui_ignore_requires_confirmation(data):
    setup_ui(data)
    c=st.session_state[ui.P+'catalog'];rev=ui._control_token([],c['sha'])
    st.values={ui.P+'choose_'+rev:['midcycle feedback'],ui.P+'save':True}
    before=data['repo'].write_count
    ui.render_ignored_student_entries(data['archive'],'Preceptor ~ Student')
    assert data['repo'].write_count==before


def test_ui_success_clears_stale_results_schedules_recount(data):
    setup_ui(data)
    c=st.session_state[ui.P+'catalog'];rev=ui._control_token([],c['sha'])
    st.session_state.update(teaching_zip=b'old',assessment_completion_inputs={'old':'old'})
    st.values={ui.P+'choose_'+rev:['midcycle feedback'],ui.P+'save':True,
               ui.P+'confirm_'+ui._control_token(['midcycle feedback'],c['sha']):True}
    with pytest.raises(StopRun): ui.render_ignored_student_entries(data['archive'],'Preceptor ~ Student')
    assert 'teaching_scan' not in st.session_state and 'teaching_zip' not in st.session_state
    assert 'assessment_completion_inputs' not in st.session_state
    assert st.session_state[ui.P+'rescan']['commit']==data['scan']['commit']
    assert 'midcycle feedback' in st.session_state[ui.P+'catalog']['entries']


def test_ui_failed_save_keeps_flags_and_scan(data):
    setup_ui(data)
    c=st.session_state[ui.P+'catalog'];rev=ui._control_token([],c['sha'])
    st.values={ui.P+'choose_'+rev:['midcycle feedback'],ui.P+'save':True,
               ui.P+'confirm_'+ui._control_token(['midcycle feedback'],c['sha']):True}
    with patch.object(GitHubIgnoredStudentEntries,'ignore_entries',side_effect=OPDArchiveError('temporary failure')):
        ui.render_ignored_student_entries(data['archive'],'Preceptor ~ Student')
    assert st.session_state[ui.P+'catalog']['entries']=={}
    assert 'teaching_scan' in st.session_state


def test_lock_clears_inventory_but_not_git(data):
    setup_ui(data)
    exclude(data)
    lock_evaluation_records()
    assert not any(str(k).startswith(ui.P) for k in st.session_state)
    assert 'midcycle feedback' in data['service'].load()['entries']


def test_main_page_save_rescan_and_restore_end_to_end(data):
    from helpers import run_app
    from test_strict_reach_charts import ui_values
    first=run_app(ui_values(teaching_build_zip=False),secrets=data['secrets'],repo=data['repo'],evaluation_login=True)
    assert 'teaching_scan' in first['state']
    cat=first['state'][ui.P+'catalog'];rev=ui._control_token([],cat['sha'])
    save=ui_values(teaching_load_archives=False,teaching_build_zip=False,**{
        ui.P+'choose_'+rev:['midcycle feedback'],ui.P+'save':True,
        ui.P+'confirm_'+ui._control_token(['midcycle feedback'],cat['sha']):True})
    second=run_app(dict(save,schedule_app_mode='PTS Matching',pts_matching_task='Ignored student entries'),secrets=data['secrets'],repo=data['repo'],state=first['state'],evaluation_login=True)
    assert 'teaching_scan' not in second['state'] and ui.P+'rescan' in second['state']
    third=run_app(ui_values(teaching_load_archives=False),secrets=data['secrets'],repo=data['repo'],state=second['state'],evaluation_login=True)
    row=teaching_annual_rows(third['state']['teaching_scan'],[2026])[0]
    assert row['educational_hours']==12 and row['recorded_clinical_hours']==28
    assert third['state']['teaching_period_start']==date(2026,9,7)
    assert any(name.endswith('.zip') for name in third['downloads'])
    cat=third['state'][ui.P+'catalog'];rev=ui._control_token([],cat['sha'])
    restore=ui_values(teaching_load_archives=False,teaching_build_zip=False,**{
        ui.P+'restore_choose_'+rev:['midcycle feedback'],ui.P+'restore':True,
        ui.P+'restore_confirm_'+ui._control_token(['midcycle feedback'],cat['sha']):True})
    fourth=run_app(dict(restore,schedule_app_mode='PTS Matching',pts_matching_task='Ignored student entries'),secrets=data['secrets'],repo=data['repo'],state=third['state'],evaluation_login=True)
    fifth=run_app(ui_values(teaching_load_archives=False),secrets=data['secrets'],repo=data['repo'],state=fourth['state'],evaluation_login=True)
    assert fifth['state']['teaching_scan']['monthly']==data['scan']['monthly']
    assert any(name.endswith('.zip') for name in fifth['downloads'])


def test_cross_session_change_blocks_cached_report(data):
    from helpers import run_app
    from test_strict_reach_charts import ui_values
    first=run_app(ui_values(),secrets=data['secrets'],repo=data['repo'],evaluation_login=True)
    assert 'teaching_zip' in first['state']
    exclude(data)
    second=run_app(ui_values(teaching_load_archives=False),secrets=data['secrets'],repo=data['repo'],state=first['state'],evaluation_login=True)
    assert not any(name.endswith('.zip') for name in second['downloads'])
    assert any('ignored-student list changed' in msg for _,msg in second['messages'])


def test_ignore_list_does_not_reset_standard_year_selection(data):
    from helpers import run_app
    from test_strict_reach_charts import ui_values
    vals=ui_values(teaching_build_zip=False,teaching_reporting_mode='Standard July-June academic years')
    first=run_app(vals,secrets=data['secrets'],repo=data['repo'],evaluation_login=True)
    state=first['state'];state['teaching_selected_years']=[2026]
    cat=state[ui.P+'catalog'];rev=ui._control_token([],cat['sha'])
    vals.update(teaching_load_archives=False,**{
        ui.P+'choose_'+rev:['midcycle feedback'],ui.P+'save':True,
        ui.P+'confirm_'+ui._control_token(['midcycle feedback'],cat['sha']):True})
    second=run_app(vals,secrets=data['secrets'],repo=data['repo'],state=state,evaluation_login=True)
    third=run_app(ui_values(teaching_load_archives=False,teaching_build_zip=False,teaching_reporting_mode='Standard July-June academic years'),
                  secrets=data['secrets'],repo=data['repo'],state=second['state'],evaluation_login=True)
    assert third['state']['teaching_selected_years']==[2026]


def test_saved_list_filters_new_source_next_session(data):
    exclude(data)
    from helpers import run_app
    from test_strict_reach_charts import ui_values
    result=run_app(ui_values(),secrets=data['secrets'],repo=data['repo'],evaluation_login=True)
    row=teaching_annual_rows(result['state']['teaching_scan'],[2026])[0]
    assert row['educational_hours']==12


def test_percentages_recheck_with_current_catalog_and_ignored_still_stored(data):
    scan,c,_=exclude(data)
    inputs=load_completion_inputs(data['archive'],view(scan),[2026])
    assert inputs['student_exclusions_signature']==exclusion_signature(c['entries'])
    # Storing an original source never discards that ignored entry from GitHub.
    raw=data['archive'].load(date(2026,9,7))['raw']
    from schedule_app.services.teaching_analysis import teaching_extract_assignments
    from schedule_app.services.opd_archive import inspect_opd_rotation
    records,_=teaching_extract_assignments(raw,inspect_opd_rotation(raw),'Preceptor ~ Student')
    assert 'midcycle feedback' in {r['student'] for r in records}
