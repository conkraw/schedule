"""Exact designation-aware name matching; all IDs and assessment data are invented."""
from copy import deepcopy
from unittest.mock import patch
import pytest
from test_student_name_matching_ui import data, show, review, P
from test_assessment_completion import record, raw_csv
from helpers import st
from schedule_app.services.student_name_matching import (
    StudentNameMatcher, student_matching_key, student_name_without_designations,
)
from schedule_app.services.student_assessment_links import student_name_key, GitHubStudentAssessmentLinks
from schedule_app.services.assessment_completion import (
    prepare_student_assessments, build_completion_bundle, ASSESSMENT_VERSION,
)
from schedule_app.services.student_name_review import (
    oasis_student_choices, name_only_student_choices, selected_name_student_id,
    student_name_review, STUDENT_REVIEW_UI_VERSION,
)
from schedule_app.services.opd_archive import OPDArchiveError


@pytest.mark.parametrize('name', [
    'Adlassnig, Sarah (MD)', 'Adlassnig, Sarah (PA)', 'Adlassnig, Sarah (DO)',
    '  ADLASSNIG ,   Sarah   (md)  ', 'Adlassnig, Sarah(MD)',
    'Adlassnig, Sarah (M.D.)', 'Adlassnig, Sarah (D.O.)', 'Adlassnig, Sarah (P.A.)',
    'Adlassnig, Sarah; MD2028', 'Adlassnig, Sarah; PA2028', 'Adlassnig, Sarah; DO 2028',
    'Adlassnig, Sarah (MD2028)', 'Adlassnig, Sarah (PA 2028)',
    'Adlassnig, Sarah (MD); MD2028', 'Adlassnig, Sarah; MD2028 (MD)',
    'Adlassnig, Sarah\u00a0(MD)',
])
def test_known_trailing_designations_match_exact_base_name(name):
    matcher = StudentNameMatcher({'adlassnig, sarah': ['example-id']}, {})
    assert matcher.resolve(name).candidate_ids == ('example-id',)
    assert matcher.resolve(name).match_status != 'Needs review'
    assert student_matching_key(name) == 'adlassnig, sarah'


@pytest.mark.parametrize('name', [
    'Adlassnig, Sara (MD)', 'Adlassnigg, Sarah (MD)',
    'Sarah Adlassnig (MD)', 'Adlassnig, Sarah Jane (MD)',
    'Adlassnig, Sarah J. (MD)', 'Adlassnig, Sarah (visiting)',
    'Adlassnig, Sarah; something', 'Adlassnig, Sarah (MD applicant)',
    'Adlassnig, Sarah MD', 'Adlassnig, Sarah (2028)',
])
def test_real_differences_and_unknown_labels_still_need_review(name):
    result = StudentNameMatcher({'adlassnig, sarah': ['example-id']}, {}).resolve(name)
    assert result.match_status == 'Needs review' and result.candidate_ids == ()


def test_full_family_name_suffix_and_punctuation_are_preserved():
    assert student_matching_key("O'Brien Jr., Alex (MD)") == "o'brien jr., alex"
    assert student_matching_key('Lee-Smith, Casey III (PA)') == 'lee-smith, casey iii'
    assert student_matching_key('Jordan (MD), Chris') == 'jordan (md), chris'
    assert student_matching_key('Learner, Alpha (preferred name) (MD)') == 'learner, alpha (preferred name)'


def test_matching_key_is_separate_from_persistent_catalog_key():
    assert student_name_key('Learner, Alpha (MD)') == 'learner, alpha (md)'
    assert student_matching_key('Learner, Alpha (MD)') == 'learner, alpha'


@pytest.mark.parametrize('opd', ['Learner, Alpha', 'Learner, Alpha (MD)', 'Learner, Alpha (PA)'])
def test_same_normalized_name_multiple_distinct_ids_is_never_guessed(opd):
    matcher = StudentNameMatcher({'learner, alpha (md)': ['one'], 'learner, alpha (pa)': ['two']}, {})
    result = matcher.resolve(opd)
    assert result.match_status == 'Needs review'
    assert result.candidate_ids == ('one', 'two')


def test_repeated_same_record_under_designation_variants_is_not_ambiguous():
    matcher = StudentNameMatcher({'learner, alpha': ['00123'], 'learner, alpha (md)': ['00123']}, {})
    assert matcher.resolve('Learner, Alpha; MD2028').candidate_ids == ('00123',)


def test_blank_and_invalid_id_candidates_do_not_become_matches():
    matcher = StudentNameMatcher({'learner, alpha': ['', 'none', 'N/A', 'null', '-']}, {})
    assert matcher.resolve('Learner, Alpha (MD)').match_status == 'Needs review'
    assert matcher.resolve('(MD)').candidate_ids == ()


def test_no_inputs_are_mutated_by_normalization_or_matching():
    ids = {'learner, alpha (md)': ['00123']}
    entries = {'learner, other (md)': {'student_name': 'Learner, Other (MD)', 'external_id': '00999'}}
    before = deepcopy((ids, entries))
    matcher = StudentNameMatcher(ids, entries)
    matcher.resolve('Learner, Alpha')
    assert (ids, entries) == before


def test_crosswalk_preparation_retains_identity_and_cleans_display_only():
    original = raw_csv([record(student='Learner, Alpha (MD)'), record('2', sid='id-pa', student='Learner, Beta (PA)')])
    prepared = prepare_student_assessments([('sample.csv', original)])
    assert prepared['student_names']['learner, alpha (md)'] == 'Learner, Alpha'
    assert prepared['student_names']['learner, beta (pa)'] == 'Learner, Beta'
    assert prepared['name_ids']['learner, alpha (md)'] == ['id-alpha']
    assert len(prepared['forms']) == 2
    assert b'(MD)' in original
    matcher = StudentNameMatcher(prepared['name_ids'], {})
    assert matcher.resolve('Learner, Alpha').candidate_ids == ('id-alpha',)


def test_auto_matched_names_do_not_appear_in_ui_or_cause_github_writes(data):
    for row in data['inputs']['assignments']:
        row['student'] = 'Learner, Alpha (MD)' if row['student'] == 'Learner, Alpha' else 'Learner, Beth (PA)'
    result, bundle, missing = review(data)
    assert not missing and not result['missing']
    assert bundle['rows'][0]['eligible_students'] == 2
    assert bundle['rows'][0]['clinical_completion_pct'] == 100
    before = data['repo'].write_count
    shown = show(data)
    assert P + 'student_choice' not in shown['options']
    assert not any('confirm these two names' in label for label, _ in shown['widgets'])
    assert data['repo'].write_count == before
    assert not shown['rerun']


def test_only_genuine_name_discrepancy_remains_in_queue(data):
    for row in data['inputs']['assignments']:
        if row['student'] == 'Learner, Alpha':
            row['student'] = 'Learner, Alpha (MD)'
    result, bundle, _ = review(data)
    assert list(result['missing']) == ['learner, beta']
    assert bundle['rows'][0]['eligible_students'] == 2
    assert 'Provisional:' in bundle['rows'][0]['assessment_status']


@pytest.mark.parametrize('threshold,eligible', [(3, 2), (4, 0), (5, 0)])
def test_threshold_uses_unified_student_shifts_not_designation_spelling(data, threshold, eligible):
    for row in data['inputs']['assignments']:
        row['student'] = 'Learner, Alpha (MD)' if row['student'] == 'Learner, Alpha' else 'Learner, Beth (PA)'
    bundle, missing = build_completion_bundle(data['inputs'], data['view'], [2026], minimum_shifts=threshold)
    assert not missing
    assert bundle['rows'][0]['eligible_students'] == eligible
    assert bundle['minimum_shifts'] == threshold
    assert bundle['rows'][0]['clinical_completion_pct'] == (100 if eligible else None)


def test_designation_variants_same_date_shift_do_not_double_eligibility(data):
    original = [r for r in data['inputs']['assignments'] if r['student'] == 'Learner, Alpha']
    data['inputs']['assignments'] = original + [{**r, 'student': 'Learner, Alpha (MD)'} for r in original]
    bundle, missing = build_completion_bundle(data['inputs'], data['view'], [2026], minimum_shifts=4)
    assert not missing
    assert bundle['rows'][0]['eligible_students'] == 0


def test_different_dates_combine_across_plain_and_designated_names(data):
    original = [r for r in data['inputs']['assignments'] if r['student'] == 'Learner, Alpha']
    data['inputs']['assignments'] = original + [{**original[0], 'student': 'Learner, Alpha (MD)', 'date': '2026-09-10'}]
    bundle, missing = build_completion_bundle(data['inputs'], data['view'], [2026], minimum_shifts=4)
    assert not missing and bundle['rows'][0]['eligible_students'] == 1
    assert bundle['rows'][0]['clinical_completion_pct'] == 100


def test_date_filter_still_excludes_out_of_period_designation_variant(data):
    original = [r for r in data['inputs']['assignments'] if r['student'] == 'Learner, Alpha']
    data['inputs']['assignments'] = original + [{**original[0], 'student': 'Learner, Alpha (MD)', 'date': '2026-10-01'}]
    bundle, _ = build_completion_bundle(data['inputs'], data['view'], [2026], minimum_shifts=4)
    assert bundle['rows'][0]['eligible_students'] == 0


def test_duplicate_oasis_designation_variants_disabled_in_name_only_selector():
    prepared = {'name_ids': {'learner, alpha (md)': ['one'], 'learner, alpha (pa)': ['two']},
                'student_names': {'learner, alpha (md)': 'Learner, Alpha (MD)', 'learner, alpha (pa)': 'Learner, Alpha (PA)'}}
    options = name_only_student_choices(oasis_student_choices(prepared))
    assert len(options) == 1
    token, entry = next(iter(options.items()))
    assert entry['ambiguous_name'] and entry['external_id'] == ''
    with pytest.raises(OPDArchiveError, match='More than one OASIS'):
        selected_name_student_id(options, token)


def test_original_id_link_catalog_and_decorated_keys_still_round_trip(data):
    service = GitHubStudentAssessmentLinks(data['archive'])
    saved = service.save_link('Learner, Alpha (MD)', 'verified-id', expected=service.load())
    assert 'learner, alpha (md)' in saved['entries']
    assert service.load()['entries'] == saved['entries']
    # Explicit earlier corrections remain authoritative, even if the source now
    # offers a different record. No automatic rewrite or removal of a saved link.
    matcher = StudentNameMatcher({'learner, alpha': ['id-alpha']}, saved['entries'])
    assert matcher.resolve('Learner, Alpha (MD)').candidate_ids == ('verified-id',)
    updated = service.save_link('Learner, Alpha (MD)', 'id-alpha', expected=saved)
    assert service.load()['entries'] == updated['entries']
    removed = service.remove_link('Learner, Alpha (MD)', expected=updated)
    assert not removed['entries']
    assert StudentNameMatcher({'learner, alpha': ['id-alpha']}, removed['entries']).resolve('Learner, Alpha (MD)').candidate_ids == ('id-alpha',)


def test_conflicting_legacy_saved_corrections_are_not_silently_rekeyed(data):
    store = GitHubStudentAssessmentLinks(data['archive'])
    saved = store.save_link('Learner, Alpha (MD)', 'one', expected=store.load())
    saved = store.save_link('Learner, Alpha (PA)', 'two', expected=saved)
    loaded = store.load()
    assert len(loaded['entries']) == 2
    matcher = StudentNameMatcher({}, loaded['entries'])
    assert matcher.resolve('Learner, Alpha (MD)').candidate_ids == ('one',)
    assert matcher.resolve('Learner, Alpha (PA)').candidate_ids == ('two',)
    # The base name is not granted either of these distinct saved identities.
    assert matcher.resolve('Learner, Alpha').match_status == 'Needs review'


def test_versions_invalidate_older_assessment_results_and_matching_widgets():
    assert ASSESSMENT_VERSION == 6
    assert STUDENT_REVIEW_UI_VERSION == 4


def test_normalization_does_not_change_reporting_denominator_for_true_typos(data):
    for row in data['inputs']['assignments']:
        if row['student'] == 'Learner, Beta':
            row['student'] = 'Learner, Betaa (MD)'
    result, bundle, missing = review(data)
    # Betaa/Beth is not an exact identity: no unconfirmed credit. A distant spelling
    # remains optionally correctable even when the conservative cue misses it.
    assert 'learner, betaa (md)' in result['absent']
    assert bundle['rows'][0]['eligible_students'] == 2
    assert bundle['rows'][0]['clinical_completion_pct'] == 50
    assert len(missing) == 1


def test_all_automatic_names_leave_no_identifiers_in_report_bundle(data):
    for row in data['inputs']['assignments']:
        row['student'] = 'Learner, Alpha (MD)' if row['student'] == 'Learner, Alpha' else 'Learner, Beth (PA)'
    bundle, missing = build_completion_bundle(data['inputs'], data['view'], [2026])
    assert not missing
    for text in ('Learner, Alpha', 'Learner, Beth', 'id-alpha', 'id-beta'):
        assert text not in str(bundle)
