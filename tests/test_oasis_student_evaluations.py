"""Synthetic, offline regressions for the additional student-assessment archive."""
import csv
from dataclasses import replace
from datetime import date
import hashlib
import io
import unittest
from unittest.mock import patch

from schedule_app.services.oasis_privacy import minimize_oasis_csv

def minimal(raw):
    return minimize_oasis_csv(raw, "student")['raw']

from cryptography.fernet import Fernet
from helpers import st, FakeResponse, Upload, secret_settings, run_app
from test_oasis_evaluations import RecordingGitHub, make_csv as educator_archive_csv
from test_oasis_educator_reports import make_csv as educator_report_csv
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive
from schedule_app.services.oasis_evaluations import (
    OASISArchiveError, GitHubOASISEvaluations, OASIS_MAX_BYTES, OASIS_MAX_FIELD_CHARS,
)
from schedule_app.services.oasis_student_evaluations import (
    GitHubOASISStudentEvaluations, inspect_student_oasis_csv,
    validate_educator_upload_kind, student_export_label, STUDENT_FORM_TITLES,
)
from schedule_app.services.oasis_workflow import load_cumulative_evaluations, GitHubOASISSummaries
from schedule_app.sections import oasis_student_evaluations as ui, oasis_workflow as combined

COLUMNS = ["Course ID", "Start Date", "End Date", "Student", "Student Username",
           "Evaluator", "Evaluator Email", "Evaluation", "Form Record", "Question ID",
           "Question", "Answer text", "Multiple Choice Value", "Submit Date"]


def make_student_csv(*, types=STUDENT_FORM_TITLES, encoding="utf-8-sig", answer='Example comment, "verbatim"\nSecond line',
                     start="2026-03-16", end="2026-04-10", submitted="2026-04-11 08:00:00",
                     form_record="100", email="teacher@example.edu"):
    stream = io.StringIO(newline="")
    w = csv.writer(stream, lineterminator="\r\n")
    w.writerow(COLUMNS)
    for form in types:
        for question in ("11", "12"):
            w.writerow(["DEMO-101", start, end, "Private Learner", "privatelearner",
                        "Private Teacher", email, form, form_record, question,
                        "Synthetic student behavior", answer, "4", submitted])
    return stream.getvalue().encode(encoding)


class ValidationTests(unittest.TestCase):
    def test_all_three_student_form_types(self):
        d = inspect_student_oasis_csv(make_student_csv())
        self.assertEqual(d["row_count"], 6)
        self.assertEqual(d["form_count"], 3)
        self.assertEqual(d["form_types"], list(STUDENT_FORM_TITLES))
        self.assertEqual([x["forms_in_file"] for x in d["form_type_counts"]], [1, 1, 1])

    def test_any_one_of_the_three_types_can_be_uploaded(self):
        for form in STUDENT_FORM_TITLES:
            with self.subTest(form=form):
                self.assertEqual(inspect_student_oasis_csv(make_student_csv(types=[form]))["form_count"], 1)

    def test_only_label_case_and_star_whitespace_normalized(self):
        raw = make_student_csv(types=[" *  CLINICAL   ASSESSMENT of STUDENT "])
        self.assertEqual(inspect_student_oasis_csv(raw)["form_types"], [STUDENT_FORM_TITLES[0]])

    def test_output_metadata_contains_no_student_teacher_or_comment(self):
        d = repr(inspect_student_oasis_csv(make_student_csv()))
        for value in ("Private Learner", "privatelearner", "Private Teacher", "verbatim", "teacher@example.edu"):
            self.assertNotIn(value, d)

    def test_unknown_form_rejected_without_printing_source_label(self):
        raw = make_student_csv(types=["PRIVATE_UNKNOWN_FORM"])
        with self.assertRaisesRegex(OASISArchiveError, "not supported") as e:
            inspect_student_oasis_csv(raw)
        self.assertNotIn("PRIVATE_UNKNOWN_FORM", str(e.exception))

    def test_educator_feedback_rejected(self):
        with self.assertRaises(OASISArchiveError):
            inspect_student_oasis_csv(educator_report_csv())

    def test_mixed_student_and_educator_forms_rejected_whole(self):
        raw = make_student_csv(types=[STUDENT_FORM_TITLES[0], "*Clinical Teaching Eval"])
        with self.assertRaisesRegex(OASISArchiveError, "No rows have been saved or dropped"):
            inspect_student_oasis_csv(raw)

    def test_student_file_rejected_by_educator_upload_guard(self):
        with self.assertRaisesRegex(OASISArchiveError, "Evaluations of students"):
            validate_educator_upload_kind(make_student_csv())

    def test_existing_educator_export_passes_upload_guard(self):
        self.assertIsNone(validate_educator_upload_kind(educator_report_csv()))
        self.assertIsNone(validate_educator_upload_kind(educator_archive_csv()))

    def test_missing_student_column_rejected(self):
        raw = make_student_csv().replace(b",Student,", b",Other Column,", 1)
        with self.assertRaisesRegex(OASISArchiveError, "Missing columns: Student"):
            inspect_student_oasis_csv(raw)

    def test_missing_submit_date_column_rejected(self):
        raw = make_student_csv().replace(b"Submit Date", b"Different Column", 1)
        with self.assertRaisesRegex(OASISArchiveError, "Submit Date"):
            inspect_student_oasis_csv(raw)

    def test_blank_username_and_email_do_not_block_original_storage(self):
        raw = make_student_csv(email="").replace(b"privatelearner", b"")
        self.assertEqual(inspect_student_oasis_csv(raw)["form_count"], 3)

    def test_blank_submit_dates_preserved_not_filtered(self):
        d = inspect_student_oasis_csv(make_student_csv(submitted=""))
        self.assertEqual(d["blank_submit_date_rows"], 6)
        self.assertEqual(d["row_count"], 6)

    def test_missing_form_records_preserved_and_not_counted_as_forms(self):
        d = inspect_student_oasis_csv(make_student_csv(form_record=""))
        self.assertEqual(d["missing_form_record_rows"], 6)
        self.assertEqual(d["form_count"], 0)

    def test_invalid_dates_preserved_as_undated(self):
        for start in ("", "nonsense", "2026-06-01"):
            with self.subTest(start=start):
                self.assertEqual(inspect_student_oasis_csv(make_student_csv(start=start))["coverage"], "undated")

    def test_decodes_supported_csv_encodings(self):
        for codec in ("utf-8", "utf-8-sig", "utf-16", "cp1252"):
            with self.subTest(codec=codec):
                self.assertEqual(inspect_student_oasis_csv(make_student_csv(encoding=codec, answer="Café — test"))["row_count"], 6)

    def test_multiline_quotes_formulas_and_html_are_not_interpreted(self):
        d = inspect_student_oasis_csv(make_student_csv(answer='=1+1\r\n<script>Example</script>, "quoted"'))
        self.assertEqual(d["row_count"], 6)

    def test_oversized_field_restores_global_csv_limit(self):
        before = csv.field_size_limit()
        with self.assertRaises(OASISArchiveError):
            inspect_student_oasis_csv(make_student_csv(types=[STUDENT_FORM_TITLES[0]], answer="x"*(OASIS_MAX_FIELD_CHARS+1)))
        self.assertEqual(csv.field_size_limit(), before)

    def test_size_and_non_csv_validation(self):
        for raw in (b"", b"bad\n", b"PK\x03\x04bad", b"%PDFbad", b"x"*(OASIS_MAX_BYTES+1)):
            with self.subTest(size=len(raw)), self.assertRaises(OASISArchiveError):
                inspect_student_oasis_csv(raw)


class Fixture(unittest.TestCase):
    def setUp(self):
        self.secrets = secret_settings()
        st.reset(secrets=self.secrets)
        self.repo = RecordingGitHub()
        self.config = OPDArchiveConfig(**self.secrets["opd_archive"])
        self.archive = GitHubOPDArchive(self.config, transport=self.repo)
        self.service = GitHubOASISStudentEvaluations(self.archive)
        self.raw = make_student_csv()

    def page(self, values=None, state=None):
        return run_app({"schedule_app_mode": "OER", "oasis_evaluation_kind": "Evaluations of students",
                        **(values or {})}, secrets=self.secrets, state=state, repo=self.repo, evaluation_login=True)


class StorageTests(Fixture):
    def test_original_bytes_exactly_recovered_after_encryption(self):
        saved = self.service.save(self.raw)
        self.assertEqual(self.service.load(saved["filename"])["raw"], minimal(self.raw))
        self.assertEqual(self.config.cipher().decrypt(self.repo.tree[saved["path"]]), minimal(self.raw))
        self.assertEqual(saved["details"]["form_count"], 3)

    def test_student_folder_separate_from_educator_and_opd(self):
        saved = self.service.save(self.raw)
        self.assertTrue(saved["path"].startswith("opd_archive/oasis_student_evaluations/"))
        self.assertEqual(GitHubOASISEvaluations(self.archive).list_exports()["filenames"], [])
        self.assertEqual(self.archive.list_rotations(), [])

    def test_namespace_does_not_reuse_educator_content_identifier(self):
        saved = self.service.save(self.raw)
        names = GitHubOASISEvaluations(self.archive)._candidate_names(self.raw, saved["details"])
        self.assertNotIn(saved["filename"], names)

    def test_no_plaintext_personal_values_or_plaintext_hash_sent(self):
        saved = self.service.save(self.raw)
        public = saved["path"].encode() + self.repo.tree[saved["path"]]
        for value in (b"Private Learner", b"Private Teacher", b"verbatim", hashlib.sha256(self.raw).hexdigest().encode()):
            self.assertNotIn(value, public)
        self.assertNotIn("raw", saved)

    def test_identical_upload_new_session_not_duplicated(self):
        saved = self.service.save(self.raw)
        other = GitHubOASISStudentEvaluations(self.archive)
        second = other.save(self.raw)
        self.assertEqual(second["filename"], saved["filename"])
        self.assertEqual(second["action"], "unchanged")
        self.assertEqual(self.repo.write_count, 1)

    def test_changed_file_same_dates_retains_both(self):
        a = self.service.save(self.raw)
        b = self.service.save(make_student_csv(form_record="9001"))
        self.assertNotEqual(a["filename"], b["filename"])
        self.assertEqual(len(self.service.list_exports()["filenames"]), 2)
        self.assertEqual(self.service.load(a["filename"])["raw"], minimal(self.raw))

    def test_custom_archive_folder_preserved(self):
        config = replace(self.config, folder="archives/pediatrics")
        service = GitHubOASISStudentEvaluations(GitHubOPDArchive(config, transport=self.repo))
        self.assertTrue(service.save(self.raw)["path"].startswith("archives/pediatrics/oasis_student_evaluations/"))

    def test_unknown_or_wrong_form_contacts_no_github(self):
        for raw in (educator_report_csv(), make_student_csv(types=["Unexpected Student Form"])):
            with self.assertRaises(OASISArchiveError):
                self.service.save(raw)
        self.assertEqual(self.repo.calls, [])

    def test_public_list_has_only_filenames_and_snapshot(self):
        self.service.save(self.raw)
        listing = self.service.list_exports()
        self.assertEqual(set(listing), {"filenames", "commit"})
        self.assertNotIn("Private", repr(listing))

    def test_label_shows_coverage_not_personal_data(self):
        saved = self.service.save(self.raw)
        label = student_export_label(saved["filename"])
        self.assertIn("2026-03-16 to 2026-04-10", label)
        self.assertNotIn("Private", label)

    def test_course_coverage_sort_and_undated_file(self):
        a = self.service.save(self.raw)
        b = self.service.save(make_student_csv(start="2026-09-01", end="2026-09-30"))
        c = self.service.save(make_student_csv(start=""))
        self.assertEqual(self.service.list_exports()["filenames"], [b["filename"], a["filename"], c["filename"]])

    def test_raw_download_fallback_recovers_original(self):
        self.repo.raw_fallback = True
        saved = self.service.save(self.raw)
        self.assertEqual(self.service.load(saved["filename"])["raw"], minimal(self.raw))

    def test_recovery_with_previous_key_avoids_resaving(self):
        first = self.service.save(self.raw)
        config = replace(self.config, encryption_key=Fernet.generate_key().decode(),
                         previous_encryption_keys=(self.config.encryption_key,))
        service = GitHubOASISStudentEvaluations(GitHubOPDArchive(config, transport=self.repo))
        self.assertEqual(service.load(first["filename"])["raw"], minimal(self.raw))
        self.assertEqual(service.save(self.raw)["action"], "unchanged")

    def test_wrong_key_no_download(self):
        saved = self.service.save(self.raw)
        config = replace(self.config, encryption_key=Fernet.generate_key().decode())
        with self.assertRaisesRegex(OASISArchiveError, "cannot be decrypted"):
            GitHubOASISStudentEvaluations(GitHubOPDArchive(config, transport=self.repo)).load(saved["filename"])

    def test_tampered_ciphertext_not_overwritten(self):
        saved = self.service.save(self.raw)
        self.repo.tree[saved["path"]] = b"tampered"
        self.repo._commit()
        with self.assertRaises(OASISArchiveError):
            self.service.save(self.raw)
        self.assertEqual(self.repo.write_count, 1)

    def test_renamed_ciphertext_rejected(self):
        saved = self.service.save(self.raw)
        renamed = saved["filename"].replace("2026-04-10", "2026-04-11")
        self.repo.tree[self.service.path_for(renamed)] = self.repo.tree[saved["path"]]
        self.repo._commit()
        with self.assertRaisesRegex(OASISArchiveError, "identifier"):
            self.service.load(renamed)

    def test_path_traversal_rejected_before_network(self):
        for name in ("../OPD_2026-03-16.xlsx.enc", "../oasis_evaluations/abc.csv.enc", "presets.json.enc"):
            with self.assertRaises(OASISArchiveError):
                self.service.load(name)
        self.assertEqual(self.repo.calls, [])

    def test_failed_save_and_explicit_retry(self):
        self.repo.put_status = 403
        with self.assertRaises(OASISArchiveError):
            self.service.save(self.raw)
        self.assertEqual(self.repo.write_count, 0)
        self.repo.put_status = None
        self.assertEqual(self.service.save(self.raw)["action"], "created")

    def test_verification_failure_not_reported_as_success(self):
        self.repo.break_verification = True
        with self.assertRaises(OASISArchiveError):
            self.service.save(self.raw)

    def test_lost_save_response_retry_recovers_without_duplicate(self):
        original = self.service._call
        def call(method, route, **kwargs):
            result = original(method, route, **kwargs)
            return {} if method == "PUT" else result
        with patch.object(self.service, "_call", side_effect=call):
            with self.assertRaises(OASISArchiveError):
                self.service.save(self.raw)
        self.assertEqual(self.service.save(self.raw)["action"], "unchanged")
        self.assertEqual(self.repo.write_count, 1)

    def test_list_limit_cannot_silently_return_partial_list(self):
        with patch.object(self.service, "_call", return_value={"entries": [{"name": "x"}]*1000}):
            with self.assertRaisesRegex(OASISArchiveError, "complete list"):
                self.service.list_exports()

    def test_recovery_at_snapshot_and_current_head_are_distinct(self):
        saved = self.service.save(self.raw)
        self.repo.tree.pop(saved["path"])
        self.repo._commit()
        self.assertEqual(self.service.load(saved["filename"], commit=saved["commit"])["raw"], minimal(self.raw))
        with self.assertRaisesRegex(OASISArchiveError, "not found"):
            self.service.load(saved["filename"])

    def test_student_archives_do_not_contribute_to_educator_cumulative_source_data(self):
        educators = GitHubOASISEvaluations(self.archive)
        educators.save(educator_report_csv())
        before = load_cumulative_evaluations(educators)
        self.service.save(self.raw)
        after = load_cumulative_evaluations(educators)
        self.assertEqual(len(before["forms"]), len(after["forms"]))
        self.assertEqual(before["sources"], after["sources"])
        self.assertEqual(len(after["sources"]), 1)

    def test_other_archive_bytes_not_mutated(self):
        keep = {"opd_archive/OPD_2026-03-16.xlsx.enc": b"KEEP_OPD",
                "opd_archive/preceptor_oasis_links.json.enc": b"KEEP_LINKS",
                "opd_archive/reporting_date_presets.json.enc": b"KEEP_DATES"}
        self.repo.tree.update(keep)
        self.repo._commit()
        self.service.save(self.raw)
        self.assertTrue(all(self.repo.tree[k] == v for k, v in keep.items()))


class InterfaceTests(Fixture):
    def test_new_option_available_under_existing_main_menu(self):
        result = self.page()
        self.assertIn(("Evaluation type", "oasis_evaluation_kind"), result["widgets"])
        self.assertTrue(any("Preceptor evaluations of students" in m for _, m in result["messages"]))
        self.assertEqual(self.repo.write_count, 0)

    def test_upload_automatically_saves_only_encrypted_original(self):
        result = self.page({ui.P + "upload": Upload(self.raw, "private_student_filename.csv")})
        self.assertEqual(self.repo.write_count, 1)
        self.assertEqual(len(self.repo.tree), 1)
        self.assertEqual(result["downloads"], {})
        self.assertTrue(any("encrypted, archived, and verified" in m for _, m in result["messages"]))
        self.assertEqual(result["state"][ui.P + "save"]["receipt"]["details"]["form_count"], 3)

    def test_rerun_does_not_resave(self):
        upload = Upload(self.raw, "original.csv")
        first = self.page({ui.P + "upload": upload})
        self.page({ui.P + "upload": upload}, state=first["state"])
        self.assertEqual(self.repo.write_count, 1)

    def test_new_filename_identical_file_not_saved_twice(self):
        self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        next_result = self.page({ui.P + "upload": Upload(self.raw, "renamed.csv")})
        self.assertEqual(next_result["state"][ui.P + "save"]["receipt"]["action"], "unchanged")
        self.assertEqual(self.repo.write_count, 1)

    def test_changed_file_saved_as_separate_snapshot(self):
        first = self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        self.page({ui.P + "upload": Upload(make_student_csv(form_record="9002"), "original.csv")}, first["state"])
        self.assertEqual(len(self.service.list_exports()["filenames"]), 2)

    def test_reload_and_optional_download_byte_for_byte(self):
        saved = self.service.save(self.raw)
        result = self.page({ui.P + "choice": saved["filename"], ui.P + "load": True})
        self.assertEqual(result["downloads"][saved["filename"][:-4]], minimal(self.raw))
        self.assertEqual(self.repo.write_count, 1)

    def test_failed_upload_does_not_claim_success(self):
        self.repo.put_status = 403
        result = self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        self.assertTrue(any("NOT confirmed" in m for _, m in result["messages"]))
        self.assertFalse(result["downloads"])
        self.assertIn(("Retry student-evaluation encrypted save", ui.P + "retry_save"), result["widgets"])

    def test_error_waits_for_explicit_retry_then_continues(self):
        self.repo.put_status = 403
        upload = Upload(self.raw, "original.csv")
        failed = self.page({ui.P + "upload": upload})
        calls = len(self.repo.calls)
        again = self.page({ui.P + "upload": upload}, failed["state"])
        self.assertEqual(len(self.repo.calls), calls)
        self.repo.put_status = None
        retry = self.page({ui.P + "upload": upload, ui.P + "retry_save": True}, again["state"])
        self.assertNotIn(ui.P + "save", retry["state"])
        recovered = self.page({ui.P + "upload": upload}, retry["state"])
        self.assertEqual(recovered["state"][ui.P + "save"]["receipt"]["action"], "created")

    def test_change_selection_removes_previous_download(self):
        a = self.service.save(self.raw)
        b = self.service.save(make_student_csv(form_record="9002"))
        first = self.page({ui.P + "choice": a["filename"], ui.P + "load": True})
        second = self.page({ui.P + "choice": b["filename"]}, first["state"])
        self.assertEqual(second["downloads"], {})

    def test_removed_remote_file_cannot_be_loaded_from_stale_list(self):
        saved = self.service.save(self.raw)
        first = self.page()
        self.repo.tree.pop(saved["path"])
        self.repo._commit()
        result = self.page({ui.P + "choice": saved["filename"], ui.P + "load": True}, first["state"])
        self.assertEqual(result["downloads"], {})
        self.assertTrue(any("not found" in text for _, text in result["messages"]))

    def test_no_student_names_comments_or_user_upload_filename_on_page(self):
        result = self.page({ui.P + "upload": Upload(self.raw, "private_student_filename.csv")})
        for text in ("Private Learner", "Private Teacher", "verbatim", "private_student_filename"):
            self.assertNotIn(text, repr(result["messages"]))

    def test_no_username_or_dates_required_to_archive_students(self):
        result = self.page({ui.P + "upload": Upload(make_student_csv(email=""), "original.csv")})
        self.assertEqual(self.repo.write_count, 1)
        self.assertFalse(any(k.startswith("oasis_dates_") or k.startswith("oasis_combined_") for _, k in result["widgets"]))

    def test_student_option_does_not_execute_educator_reporting(self):
        with patch.object(combined, "_process", side_effect=AssertionError("Should not run")), \
             patch.object(combined, "_review_saved", side_effect=AssertionError("Should not run")), \
             patch.object(combined.oasis_date_controls, "render", side_effect=AssertionError("Should not run")):
            self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        self.assertEqual(self.repo.write_count, 1)

    def test_student_assessments_rejected_in_educator_upload_without_write(self):
        result = self.page({"oasis_evaluation_kind": "Evaluations of educators",
                            combined.P + "upload": Upload(self.raw, "original.csv")})
        self.assertEqual(self.repo.write_count, 0)
        self.assertTrue(any("Evaluations of students" in m for _, m in result["messages"]))
        self.assertEqual(result["downloads"], {})

    def test_educator_feedback_rejected_in_student_upload_without_write(self):
        result = self.page({ui.P + "upload": Upload(educator_report_csv(), "original.csv")})
        self.assertEqual(self.repo.write_count, 0)
        self.assertTrue(any("not supported" in m for _, m in result["messages"]))

    def test_scope_reset_does_not_clear_teaching_or_educator_data(self):
        state = {ui.P + "scope": "old", ui.P + "loaded": {"bad": True},
                 "teaching_zip": b"KEEP_TEACHING", combined.P + "prepared": {"keep": True}}
        result = self.page(state=state)
        self.assertNotIn(ui.P + "loaded", result["state"])
        self.assertEqual(result["state"]["teaching_zip"], b"KEEP_TEACHING")
        self.assertEqual(result["state"][combined.P + "prepared"], {"keep": True})

    def test_no_source_bytes_retained_in_success_receipt(self):
        result = self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        self.assertNotIn("raw", result["state"][ui.P + "save"]["receipt"])
        self.assertNotIn(ui.P + "loaded", result["state"])

    def test_no_summary_saved_in_student_branch(self):
        self.page({ui.P + "upload": Upload(self.raw, "original.csv")})
        self.assertEqual(GitHubOASISSummaries(self.archive).list_outputs()["filenames"], [])


if __name__ == "__main__":
    unittest.main()
