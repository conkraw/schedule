"""Name-only UI regression coverage; all students/IDs below are invented."""
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch
import ast
import pytest

from test_student_name_matching_ui import data, show, selected_fields, review, message, P
from helpers import st
from schedule_app.sections import student_name_matches as ui
from schedule_app.services.student_name_review import (
    oasis_student_choices, name_only_student_choices, selected_name_student_id,
    student_name_review, review_table_rows,
)
from schedule_app.services.assessment_completion import build_completion_bundle
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.opd_archive import OPDArchiveError


def test_names_only_in_dropdown_and_proposed_match(data):
    result = show(data, selected_fields(data))
    suffix = ui._editor_suffix('learner, beta', data['inputs']['student_links'])
    labels = result['formatted_options'][P + 'oasis_choice_' + suffix]
    assert labels == ['— Select the correct OASIS student —', 'Learner, Alpha', 'Learner, Beth']
    text = '\n'.join(text for _, text in result['messages'])
    assert 'Name match: Learner, Beta → Learner, Beth' in text
    assert 'id-alpha' not in text and 'id-beta' not in text
    assert any(label == 'I confirm these two names refer to the same student' for label, _ in result['widgets'])
    assert not any('External ID' in label or 'How to confirm' in label for label, _ in result['widgets'])
    assert not any(key.startswith(P + 'external_id_') or key.startswith(P + 'method_') for _, key in result['widgets'])


def test_flags_and_status_ask_for_names_not_external_ids(data):
    result, bundle, _ = review(data)
    table_text = str(review_table_rows(result['missing']))
    assert 'External ID' not in table_text
    assert bundle['rows'][0]['assessment_status'].endswith('need review')
    identity_alerts = [r for r in bundle['warnings'] if r['direction'] == 'Student identity']
    assert 'Student External ID' not in str(identity_alerts)
    shown = show(data)
    for kind, text in shown['messages']:
        if kind == 'warning':
            assert 'Student External ID' not in text


def test_names_only_in_saved_matches_table(data):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    shown = show(data, {P + 'edit_saved': True}, saved['state'])
    text = '\n'.join(t for _, t in shown['messages'])
    assert 'Matching OASIS student name' in text
    assert 'Learner, Beth' in text
    assert 'id-beta' not in text
    assert 'Student External ID' not in text


def test_duplicate_oasis_name_shown_once_and_not_selectable_for_save(data):
    data['inputs']['prepared']['name_ids']['learner, beth'] = ['id-beta', 'other-student']
    options = name_only_student_choices(oasis_student_choices(data['inputs']['prepared']))
    ambiguous = [(t, r) for t, r in options.items() if r['ambiguous_name']]
    assert len(ambiguous) == 1
    token, row = ambiguous[0]
    assert row['external_id'] == ''
    with pytest.raises(OPDArchiveError, match='More than one OASIS'):
        selected_name_student_id(options, token)
    suffix = ui._editor_suffix('learner, beta', data['inputs']['student_links'])
    values = {P + 'oasis_choice_' + suffix: token,
              ui._confirmation_key(suffix, token, '', row['oasis_student_name']): True, P + 'save': True}
    before = data['repo'].write_count
    shown = show(data, values)
    message(shown, 'warning', 'A name alone cannot distinguish them')
    assert not shown['rerun'] and data['repo'].write_count == before
    labels = shown['formatted_options'][P + 'oasis_choice_' + suffix]
    assert len([s for s in labels if 'Learner, Beth' in s]) == 1
    assert 'id-beta' not in str(labels) and 'other-student' not in str(labels)


def test_multiple_aliases_for_one_record_do_not_create_ambiguity():
    choices = oasis_student_choices({'name_ids': {'a': ['00017'], 'b': ['00017']},
                                     'student_names': {'a': 'Learner, Alias', 'b': 'Learner, Original'}})
    names = name_only_student_choices(choices)
    assert len(names) == 2 and not any(r['ambiguous_name'] for r in names.values())
    assert {selected_name_student_id(names, k) for k in names} == {'00017'}


def test_name_options_stable_and_do_not_mutate_pair_level_data():
    pairs = oasis_student_choices({'name_ids': {'b': ['008', '009'], 'a': ['007']},
                                   'student_names': {'b': 'Beta', 'a': 'Alpha'}})
    before = deepcopy(pairs)
    assert name_only_student_choices(pairs) == name_only_student_choices(dict(reversed(list(pairs.items()))))
    assert pairs == before


@pytest.mark.parametrize('token', ['', None, 'stale'])
def test_stale_name_selection_is_not_resolved(data, token):
    options = name_only_student_choices(oasis_student_choices(data['inputs']['prepared']))
    with pytest.raises(OPDArchiveError, match='Choose the matching'):
        selected_name_student_id(options, token)


def test_legacy_match_is_retained_even_without_current_oasis_name(data):
    store = GitHubStudentAssessmentLinks(data['archive'])
    saved = store.save_link('Learner, Beta', 'legacy-known-student', expected=store.load())
    data['inputs']['student_links'] = saved
    shown = show(data, {P + 'edit_saved': True})
    assert P + 'student_choice' not in shown['options']
    assert 'legacy-known-student' not in str(shown['messages'])
    assert store.load()['entries']['learner, beta']['external_id'] == 'legacy-known-student'
    assert not review(data)[0]['missing']


def test_parenthesized_course_suffix_matches_automatically_without_confirmation(data):
    for r in data['inputs']['assignments']:
        if r['student'] == 'Learner, Beta':
            r['student'] = 'Learner, Beth (MD)'
    result, bundle, unmatched = review(data)
    assert not result['missing'] and not unmatched
    assert bundle['rows'][0]['eligible_students'] == 2
    assert bundle['rows'][0]['clinical_completion_pct'] == 100
    before = data['repo'].write_count
    shown = show(data)
    assert P + 'student_choice' not in shown['options']
    assert not any(label == 'I confirm these two names refer to the same student'
                   for label, _ in shown['widgets'])
    assert not shown['rerun'] and data['repo'].write_count == before


def test_version_change_discards_old_entry_fields_without_clearing_mappings(data):
    state = {P + 'scope': ('old version',), P + 'external_id_old': 'old-input',
             P + 'method_old': 'Enter a verified Student External ID', 'teaching_zip': b'old'}
    shown = show(data, state=state)
    assert P + 'external_id_old' not in shown['state'] and P + 'method_old' not in shown['state']
    assert 'teaching_zip' not in shown['state']
    assert shown['state'][ui.INPUTS_KEY] == data['inputs']


@pytest.mark.parametrize('minimum', [3, 4, 5])
def test_name_confirmation_preserves_saved_minimum_shift_calculation(data, minimum):
    saved = show(data, {**selected_fields(data), P + 'save': True})
    bundle, _ = build_completion_bundle(saved['state'][ui.INPUTS_KEY], data['view'], [2026], minimum_shifts=minimum)
    reference = deepcopy(data['inputs'])
    reference['student_links'] = GitHubStudentAssessmentLinks(data['archive']).load()
    expected, _ = build_completion_bundle(reference, data['view'], [2026], minimum_shifts=minimum)
    assert bundle == expected and bundle['minimum_shifts'] == minimum


def test_name_selection_does_not_request_id_via_text_input(data):
    with patch.object(st, 'text_input', side_effect=AssertionError('No manual ID field expected')):
        shown = show(data, selected_fields(data))
        assert not shown['rerun']


def test_only_expected_runtime_changes_and_no_calculation_change():
    new = Path(__file__).resolve().parents[1]
    old = new.parent / 'previous_release'
    if not old.exists():
        pytest.skip('Prior source used only for local packaging verification')
    changed = {p.relative_to(new).as_posix() for p in (new/'schedule_app').rglob('*.py')
               if p.read_bytes() != (old/p.relative_to(new)).read_bytes()}
    assert changed == {'schedule_app/sections/student_name_matches.py',
                       'schedule_app/services/student_name_review.py',
                       'schedule_app/services/assessment_completion.py'}
    # In the completion engine only user-facing string constants changed.
    def without_strings(src):
        node=ast.parse(src)
        for child in ast.walk(node):
            if isinstance(child, ast.Constant) and isinstance(child.value, str):
                child.value='<text>'
        return ast.dump(node)
    rel='schedule_app/services/assessment_completion.py'
    assert without_strings((new/rel).read_text()) == without_strings((old/rel).read_text())
