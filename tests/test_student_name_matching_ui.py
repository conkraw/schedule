"""PTS name-match queue and encrypted, explicit corrections. All data invented."""
from copy import deepcopy
from datetime import date
import hashlib
from io import BytesIO, StringIO
import csv
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import pytest

from helpers import st, StopRun, login_for_test
from test_assessment_completion import NAME, record, raw_csv, cells
from test_learner_reach import scan_cells
from schedule_app.services.student_name_review import (
    oasis_student_choices, selected_student_id, student_name_review, review_table_rows,
    names_for_external_id, name_only_student_choices, selected_name_student_id,
)
from schedule_app.services.assessment_completion import (
    load_completion_inputs, build_completion_bundle,
)
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.teaching_analysis import teaching_filter_date_range
from schedule_app.services.evaluation_access import lock_evaluation_records
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.sections import student_name_matches as ui
from schedule_app.sections.assessment_completion import render_assessment_completion

P = ui.P


@pytest.fixture
def data():
    scan, repo, archive, secrets, raw = scan_cells(cells(two=True))
    view = teaching_filter_date_range(scan, ReportingPeriod('26-27', date(2026, 9, 7), date(2026, 9, 13)))
    names = GitHubPreceptorOASISLinks(archive)
    names.save_username(NAME, 'avery', expected=names.load())
    service = GitHubOASISStudentEvaluations(archive)
    # OPD calls the second student Beta; OASIS uses Beth. Only user confirmation
    # can connect the two. The first student already has an exact name match.
    service.save(raw_csv([record(), record('2', sid='id-beta', student='Learner, Beth; MD2028')]))
    inputs = load_completion_inputs(archive, view, [2026])
    return dict(scan=scan, view=view, archive=archive, repo=repo, secrets=secrets, inputs=inputs)


def review(data, inputs=None):
    inputs = inputs or data['inputs']
    bundle, unmatched = build_completion_bundle(inputs, data['view'], [2026])
    return student_name_review(inputs, data['view'], [2026], unmatched), bundle, unmatched


def show(data, values=None, state=None, authenticated=True):
    state = dict(state or {})
    inputs = state.get(ui.INPUTS_KEY, data['inputs'])
    st.reset(values=values, secrets=data['secrets'], state={**state, ui.INPUTS_KEY: inputs})
    if authenticated:
        login_for_test()
    options_seen = {}
    formatted_options = {}
    base_select = st.selectbox
    def capture(label, options, *args, **kwargs):
        assert options, 'No empty dropdown should be created'
        options_seen[kwargs.get('key', label)] = list(options)
        fmt = kwargs.get('format_func', str)
        formatted_options[kwargs.get('key', label)] = [fmt(value) for value in options]
        return base_select(label, options, *args, **kwargs)
    bundle, unmatched = build_completion_bundle(inputs, data['view'], [2026])
    rerun = False
    with patch.object(st, 'selectbox', side_effect=capture):
        try:
            ui.render_student_name_matches(data['archive'], inputs, data['view'], [2026], unmatched)
        except StopRun:
            rerun = True
    return dict(state=dict(st.session_state), options=options_seen, formatted_options=formatted_options, messages=list(st.messages),
                widgets=list(st.widget_keys), rerun=rerun)


def selected_fields(data, inputs=None, *, editing=False, sid='id-beta', key='learner, beta'):
    inputs = inputs or data['inputs']
    choices = name_only_student_choices(oasis_student_choices(inputs['prepared']))
    suffix = ui._editor_suffix(key, inputs['student_links'], editing)
    chosen = next((token for token, row in choices.items() if row['external_id'] == sid), '')
    chosen_name = choices[chosen]['oasis_student_name'] if chosen else ''
    actual_id = choices[chosen]['external_id'] if chosen else ''
    fields = {ui._confirmation_key(suffix, chosen, actual_id, chosen_name): True,
              P + 'oasis_choice_' + suffix: chosen,
              P + ('saved_choice' if editing else 'student_choice'): key}
    if editing:
        fields[P + 'edit_saved'] = True
    return fields


def message(result, kind, snippet):
    assert any(k == kind and snippet in text for k, text in result['messages']), (kind, snippet)


def test_queue_only_unmatched_names_and_not_oasis_only(data):
    result, _, _ = review(data)
    assert list(result['active']) == ['learner, alpha', 'learner, beta']
    assert list(result['missing']) == ['learner, beta']
    assert 'learner, beth' not in result['active']
    assert result['eligible_missing_count'] == 1


def test_table_one_name_despite_repeated_shifts(data):
    result, _, _ = review(data)
    table = review_table_rows(result['missing'])
    assert len(table) == 1
    assert table[0]['OPD student name'] == 'Learner, Beta'
    assert table[0]['Affects completion percentage'] == 'YES'


def test_same_unresolved_student_many_preceptors_only_one_queue_item(data):
    inputs = deepcopy(data['inputs']); view = deepcopy(data['view'])
    view['monthly'].append({**view['monthly'][0], 'preceptor_name': 'Other, Teacher'})
    inputs['assignments'] += [{**r, 'preceptor_name': 'Other, Teacher'} for r in inputs['assignments']]
    _, unmatched = build_completion_bundle(inputs, view, [2026])
    result = student_name_review(inputs, view, [2026], unmatched)
    assert len(result['missing']) == 1
    assert len(result['missing']['learner, beta']['preceptors']) == 2


def test_below_three_shift_name_is_flagged_without_changing_completion(data):
    inputs = deepcopy(data['inputs'])
    beta = [r for r in inputs['assignments'] if r['student'] == 'Learner, Beta']
    inputs['assignments'] = [r for r in inputs['assignments'] if r['student'] != 'Learner, Beta'] + beta[:1]
    result, bundle, _ = review(data, inputs)
    assert 'learner, beta' in result['missing']
    assert not result['missing']['learner, beta']['affects_completion']
    assert bundle['rows'][0]['eligible_students'] == 1
    assert bundle['rows'][0]['clinical_completion_pct'] == 100


def test_outside_date_range_does_not_add_name_to_queue(data):
    inputs = deepcopy(data['inputs'])
    inputs['assignments'].append({**inputs['assignments'][0], 'student': 'Outside, Student', 'date': '2026-09-01'})
    assert 'outside, student' not in review(data, inputs)[0]['active']


def test_unresolved_provider_labels_not_in_student_queue(data):
    inputs = deepcopy(data['inputs']); view = deepcopy(data['view'])
    view['monthly'].append({**view['monthly'][0], 'preceptor_name': 'SJR_1'})
    view['unresolved_preceptor_labels'].append('SJR_1')
    inputs['assignments'].append({**inputs['assignments'][0], 'student': 'Not Named, Student', 'preceptor_name': 'SJR_1'})
    result = student_name_review(inputs, view, [2026], [])
    assert 'not named, student' not in result['active']


def test_saved_match_excluded_even_when_id_has_no_evaluation(data):
    inputs = deepcopy(data['inputs'])
    inputs['student_links']['entries']['learner, beta'] = {'external_id': 'verified-outside-id', 'student_name': 'Learner, Beta'}
    result, bundle, unmatched = review(data, inputs)
    assert not result['missing'] and not unmatched
    assert bundle['rows'][0]['clinical_completion_pct'] == 50


def test_case_spacing_and_md_suffix_still_match_without_manual_fix(data):
    inputs = deepcopy(data['inputs'])
    for r in inputs['assignments']:
        if r['student'] == 'Learner, Alpha':
            r['student'] = '  LEARNER ,   Alpha; MD2028  '
    result, _, _ = review(data, inputs)
    assert 'learner, alpha' not in result['missing']


def test_same_name_multiple_ids_flagged_not_guessed(data):
    inputs = deepcopy(data['inputs'])
    inputs['prepared']['name_ids']['learner, beta'] = ['id-beta', 'other-person']
    inputs['prepared']['student_names']['learner, beta'] = 'Learner, Beta'
    result, _, _ = review(data, inputs)
    assert 'More than one OASIS student record' in result['missing']['learner, beta']['issue']
    candidates = oasis_student_choices(inputs['prepared'])
    assert sum(row['ambiguous_name'] for row in candidates.values()) == 2


def test_choice_keys_stable_under_input_order_and_leading_zero_retained():
    prep = {'name_ids': {'b': ['001234', '000006'], 'a': ['id-9']}, 'student_names': {'b': 'Learner, B', 'a': 'Learner, A'}}
    reverse = {'name_ids': {'a': ['id-9'], 'b': ['000006', '001234']}, 'student_names': prep['student_names']}
    assert oasis_student_choices(prep) == oasis_student_choices(reverse)
    choices = oasis_student_choices(prep)
    token = next(k for k, r in choices.items() if r['external_id'] == '001234')
    assert selected_student_id(choices, token) == '001234'


@pytest.mark.parametrize('selection', ['', None, 'stale-token'])
def test_missing_or_stale_candidate_rejected(data, selection):
    with pytest.raises(OPDArchiveError, match='Choose the matching'):
        selected_student_id(oasis_student_choices(data['inputs']['prepared']), selection)


def test_empty_and_placeholder_ids_not_selectable():
    assert oasis_student_choices({'name_ids': {'a': ['', 'none', 'N/A', 'null', '-']}}) == {}


def test_alias_names_for_saved_external_id_shown_not_merged_automatically():
    choices = oasis_student_choices({'name_ids': {'a': ['1'], 'b': ['1']}, 'student_names': {'a': 'Alpha', 'b': 'Beta'}})
    assert names_for_external_id(choices, '1') == ['Alpha', 'Beta']
    assert names_for_external_id(choices, '2') == []


def test_helpers_do_not_mutate_inputs(data):
    before = deepcopy(data)
    review(data)
    oasis_student_choices(data['inputs']['prepared'])
    assert data['inputs'] == before['inputs'] and data['view'] == before['view']


def test_ui_warning_and_missing_only_dropdown(data):
    result = show(data)
    message(result, 'warning', 'Student-name match needed: 1 of 2')
    message(result, 'dataframe', 'Learner, Beta')
    assert result['options'][P + 'student_choice'] == ['learner, beta']
    assert P + 'saved_choice' not in result['options']


def test_ui_candidate_starts_unselected_and_no_silent_save(data):
    before = data['repo'].write_count
    result = show(data, {P + 'save': True})
    suffix = ui._editor_suffix('learner, beta', data['inputs']['student_links'])
    assert result['state'][P + 'oasis_choice_' + suffix] == ''
    assert data['repo'].write_count == before and not result['rerun']


def test_ui_saving_requires_confirmation_of_selected_person(data):
    values = {k: v for k, v in selected_fields(data).items() if not k.startswith(P + 'confirm_')}
    before = data['repo'].write_count
    result = show(data, {**values, P + 'save': True})
    assert not result['rerun'] and data['repo'].write_count == before


def test_save_removes_flag_rebuilds_percentage_and_invalidates_old_report(data):
    first = show(data)
    first['state']['teaching_zip'] = b'old report'
    first['state']['teaching_zip_signature'] = 'old'
    result = show(data, {**selected_fields(data), P + 'save': True}, first['state'])
    assert result['rerun'] and 'teaching_zip' not in result['state']
    store = GitHubStudentAssessmentLinks(data['archive'])
    assert store.load()['entries']['learner, beta']['external_id'] == 'id-beta'
    token = data['repo'].tree[store.path]
    assert b'Learner, Beta' not in token and b'id-beta' not in token
    again = show(data, state=result['state'])
    assert P + 'student_choice' not in again['options']
    message(again, 'success', 'No student-name discrepancies')
    bundle, unmatched = build_completion_bundle(again['state'][ui.INPUTS_KEY], data['view'], [2026])
    assert not unmatched
    assert bundle['rows'][0]['eligible_students'] == 2
    assert bundle['rows'][0]['clinical_completion_pct'] == 100


def test_next_unresolved_student_fields_clear_after_save(data):
    inputs = deepcopy(data['inputs'])
    inputs['assignments'] += [{**r, 'student': 'Learner, Gamma'} for r in inputs['assignments'][:3]]
    data['inputs'] = inputs
    result = show(data, {**selected_fields(data), P + 'save': True})
    again = show(data, state=result['state'])
    assert P + 'student_choice' not in again['options']  # Gamma has no OASIS record; no confirmation required.
    current_keys = [key for _, key in again['widgets'] if key.startswith(P + 'oasis_choice_')]
    assert not current_keys


def test_new_session_reuses_encrypted_match_without_input(data):
    result = show(data, {**selected_fields(data), P + 'save': True})
    assert result['rerun']
    data['inputs'] = load_completion_inputs(data['archive'], data['view'], [2026])
    again = show(data)
    assert P + 'student_choice' not in again['options']
    message(again, 'success', 'No student-name discrepancies')


def test_save_error_leaves_flag_and_catalog_unchanged(data):
    first = show(data)
    before = data['repo'].write_count
    with patch.object(GitHubStudentAssessmentLinks, 'save_link', side_effect=OPDArchiveError('Test save failure')):
        result = show(data, {**selected_fields(data), P + 'save': True}, first['state'])
    message(result, 'error', 'Test save failure')
    assert data['repo'].write_count == before and not result['rerun']
    assert 'learner, beta' in review(data, result['state'][ui.INPUTS_KEY])[0]['missing']


def test_stale_edit_is_reported_then_refresh_loads_current(data):
    first = show(data)
    service = GitHubStudentAssessmentLinks(data['archive'])
    service.save_link('Learner, Beta', 'id-beta', expected=service.load())
    blocked = show(data, {**selected_fields(data), P + 'save': True}, first['state'])
    message(blocked, 'error', 'changed in another session')
    assert not blocked['rerun']
    refreshed = show(data, {P + 'refresh': True}, blocked['state'])
    assert refreshed['rerun']
    again = show(data, state=refreshed['state'])
    assert P + 'student_choice' not in again['options']


def test_refresh_failure_never_treats_catalog_as_empty(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    with patch.object(GitHubStudentAssessmentLinks, 'load', side_effect=OPDArchiveError('Cannot read catalog')):
        result = show(data, {P + 'refresh': True}, saved['state'])
    message(result, 'error', 'Cannot read catalog')
    assert result['state'][ui.INPUTS_KEY]['student_links']['entries']['learner, beta']['external_id'] == 'id-beta'


def test_existing_matches_editable_only_in_optional_area(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    result = show(data, {P + 'edit_saved': True}, saved['state'])
    assert result['options'][P + 'saved_choice'] == ['learner, beta']
    assert P + 'student_choice' not in result['options']


def test_remove_saved_match_returns_name_to_flagged_queue(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    inputs = saved['state'][ui.INPUTS_KEY]
    suffix = ui._editor_suffix('learner, beta', inputs['student_links'], True)
    removed = show(data, {P + 'edit_saved': True, P + 'remove_confirm_' + suffix: True, P + 'remove': True}, saved['state'])
    assert removed['rerun']
    again = show(data, state=removed['state'])
    assert again['options'][P + 'student_choice'] == ['learner, beta']
    assert GitHubStudentAssessmentLinks(data['archive']).load()['entries'] == {}


def test_saved_link_update_recalculates_again_without_source_changes(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    inputs = saved['state'][ui.INPUTS_KEY]
    # Correct a saved match by selecting a different available OASIS name.
    inputs['prepared']['name_ids']['learner, delta'] = ['verified-other-id']
    inputs['prepared']['student_names']['learner, delta'] = 'Learner, Delta'
    updated = show(data, {**selected_fields(data, inputs, editing=True, sid='verified-other-id'), P + 'update': True}, saved['state'])
    assert updated['rerun']
    bundle, unmatched = build_completion_bundle(updated['state'][ui.INPUTS_KEY], data['view'], [2026])
    assert not unmatched and bundle['rows'][0]['clinical_completion_pct'] == 50
    assert bundle['rows'][0]['eligible_students'] == 2


def test_same_saved_match_can_be_saved_again_without_extra_commit(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    before = data['repo'].write_count
    inputs = saved['state'][ui.INPUTS_KEY]
    again = show(data, {**selected_fields(data, inputs, editing=True), P + 'update': True}, saved['state'])
    assert again['rerun'] and data['repo'].write_count == before


def test_name_choice_preserves_leading_zero_internal_id(data):
    data['inputs']['prepared']['name_ids']['learner, gamma'] = ['000045']
    data['inputs']['prepared']['student_names']['learner, gamma'] = 'Learner, Gamma'
    saved = show(data, {**selected_fields(data, sid='000045'), P + 'save': True})
    assert saved['rerun']
    assert GitHubStudentAssessmentLinks(data['archive']).load()['entries']['learner, beta']['external_id'] == '000045'


def test_old_manual_widget_state_cannot_save_an_arbitrary_id(data):
    suffix = ui._editor_suffix('learner, beta', data['inputs']['student_links'])
    fields = {P + 'external_id_' + suffix: 'n/a', P + 'method_' + suffix: 'Enter a verified Student External ID',
              P + 'save': True}
    result = show(data, fields)
    assert not result['rerun']
    assert not GitHubStudentAssessmentLinks(data['archive']).load()['entries']
    assert not any('External ID' in label for label, _ in result['widgets'])


def test_changed_candidate_does_not_reuse_other_identity_confirmation(data):
    fields = selected_fields(data)
    choices = oasis_student_choices(data['inputs']['prepared'])
    token = next(k for k, r in choices.items() if r['external_id'] == 'id-alpha')
    suffix = ui._editor_suffix('learner, beta', data['inputs']['student_links'])
    fields[P + 'oasis_choice_' + suffix] = token
    result = show(data, {**fields, P + 'save': True})
    assert not result['rerun']
    assert not GitHubStudentAssessmentLinks(data['archive']).load()['entries']


def test_missing_sources_does_not_invent_oasis_names(data):
    inputs = deepcopy(data['inputs'])
    inputs['prepared'] = {**inputs['prepared'], 'name_ids': {}, 'student_names': {}, 'forms': [], 'source_count': 0}
    data['inputs'] = inputs
    result = show(data)
    message(result, 'info', 'have no record in the loaded OASIS data')
    assert not any(k.startswith(P + 'oasis_choice_') for k in result['options'])
    assert P + 'student_choice' not in result['options']  # No OASIS candidates is not a confirmation task.


def test_nothing_shown_or_saved_before_password_unlock(data):
    before = data['repo'].write_count
    result = show(data, {**selected_fields(data), P + 'save': True}, authenticated=False)
    assert not result['options'] and not result['widgets']
    assert data['repo'].write_count == before
    assert ui.INPUTS_KEY not in result['state']


def test_lock_clears_new_controls_and_matching_data(data):
    result = show(data)
    assert any(k.startswith(P) for k in result['state'])
    lock_evaluation_records()
    assert not any(k.startswith(P) or k == ui.INPUTS_KEY for k in st.session_state)


def test_corrections_do_not_modify_opd_oasis_files_or_other_catalogs(data):
    prior = dict(data['repo'].tree)
    saved = show(data, {**selected_fields(data), P + 'save': True})
    assert saved['rerun']
    assert all(data['repo'].tree[key] == val for key, val in prior.items())
    added = set(data['repo'].tree) - set(prior)
    assert added == {GitHubStudentAssessmentLinks(data['archive']).path}


def test_name_match_cannot_repair_form_missing_external_id(data):
    # A mapping connects identities. The form's own missing ID must still be
    # corrected in OASIS; never remove that separate source-data warning.
    data['inputs']['prepared'] = __import__('schedule_app.services.assessment_completion', fromlist=['prepare_student_assessments']).prepare_student_assessments([
        ('synthetic', raw_csv([record(), record('3', sid='', student='Learner, Beta'),
                               record('4', sid='id-beta', student='Learner, Beth')]))])
    saved = show(data, {**selected_fields(data, sid='id-beta'), P + 'save': True})
    bundle, unmatched = build_completion_bundle(saved['state'][ui.INPUTS_KEY], data['view'], [2026])
    assert not unmatched
    assert bundle['rows'][0]['assessment_status'].startswith('Not verified: assessment source metadata')


def test_corrected_reports_still_omit_student_names_and_ids(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    bundle, _ = build_completion_bundle(saved['state'][ui.INPUTS_KEY], data['view'], [2026])
    view = dict(data['view'], assessment_completion=bundle)
    output, _ = teaching_build_zip(view, [2026])
    from docx import Document
    from test_reporting_dates import all_doc_text
    with ZipFile(BytesIO(output)) as z:
        for filename in z.namelist():
            if filename.endswith('.docx'):
                text = all_doc_text(z.read(filename))
            elif filename.endswith(('.csv', '.txt', '.json')):
                text = z.read(filename).decode('utf-8-sig')
            else:
                continue
            for value in ('Learner, Alpha', 'Learner, Beta', 'Learner, Beth', 'id-alpha', 'id-beta'):
                assert value not in text, filename


def test_new_queue_is_called_from_completion_section(data):
    st.reset(values={'assessment_completion_refresh': False}, secrets=data['secrets'])
    login_for_test()
    with patch('schedule_app.sections.assessment_completion.load_completion_inputs', return_value=data['inputs']):
        st.values['assessment_completion_refresh'] = True
        bundle = render_assessment_completion(data['archive'], data['view'], [2026])
    assert bundle['rows'][0]['assessment_status'].startswith('Provisional:')
    assert any(label == 'OPD students needing an OASIS match' for label, _ in st.widget_keys)


def test_saving_is_not_a_blanket_dismiss_for_other_students(data):
    inputs = deepcopy(data['inputs'])
    inputs['assignments'] += [{**r, 'student': 'Learner, Gamma'} for r in inputs['assignments']]
    data['inputs'] = inputs
    saved = show(data, {**selected_fields(data), P + 'save': True})
    result, bundle, unmatched = review(data, saved['state'][ui.INPUTS_KEY])
    assert not result['missing'] and list(result['absent']) == ['learner, gamma']
    assert bundle['rows'][0]['eligible_students'] == 3
    assert bundle['rows'][0]['clinical_completion_pct'] == 66.7


def test_update_changes_only_three_runtime_files():
    root = Path(__file__).resolve().parents[1]
    baseline = root.parent / 'original'
    if not baseline.exists():
        pytest.skip('Packaging comparison baseline is not bundled')
    changed = {p.relative_to(root).as_posix() for p in [root / 'app_sch_2026.py', *(root / 'schedule_app').rglob('*.py')]
               if not (baseline / p.relative_to(root)).exists() or p.read_bytes() != (baseline / p.relative_to(root)).read_bytes()}
    assert changed == {'schedule_app/sections/assessment_completion.py',
                       'schedule_app/sections/student_name_matches.py',
                       'schedule_app/services/student_name_review.py'}
