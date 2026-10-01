"""New privacy/auth regressions. Invented data only; GitHub is simulated."""
import ast
import base64
import csv
from datetime import date
import hashlib
import io
from pathlib import Path
import time
from unittest.mock import patch
import pytest

from helpers import st, FakeGitHub, FakeResponse, Upload, secret_settings, run_app, ROOT
from test_oasis_evaluations import RecordingGitHub
from schedule_app.services.opd_archive import OPDArchiveConfig, GitHubOPDArchive
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations, OASISArchiveError
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.oasis_privacy import minimize_oasis_csv, EDUCATOR_COLUMNS, STUDENT_COLUMNS
from schedule_app.services.oasis_educator_reports import prepare_reports, educator_summary
from schedule_app.services.assessment_completion import prepare_student_assessments
from schedule_app.services import evaluation_access as access

PASSWORD = "test-only-long-evaluation-password"
PRIVATE = "DO_NOT_RETAIN_THIS_STUDENT_DETAIL"


def source(kind="educator", *, extra=None, codec="utf-8-sig"):
    base = {
        "Course ID": "DEMO-101", "Department": "Dept", "Course": "Course title", "Location": "Site",
        "Start Date": "2026-03-02", "End Date": "2026-03-27", "Course Type": "Rotation",
        "Student": "Learner, One; MD2028", "Student Username": PRIVATE, "Student External ID": "student001",
        "Student Email": PRIVATE, "Student AAMC ID": PRIVATE, "Student USMLE ID": PRIVATE,
        "Student Gender": PRIVATE, "Student Level": "3", "Student Default Classification": PRIVATE,
        "Student Designation": "MD2028", "Evaluator": "Example, Avery", "Evaluator Username": "avery",
        "Evaluator External ID": "staff001", "Evaluator Email": "avery@example.edu", "Evaluator Gender": PRIVATE,
        "Who Completed": PRIVATE, "Evaluation": "*Clinical Teaching Eval", "Form Record": "1001",
        "Question Number": "2", "Question ID": "588", "Question": "Built on my knowledge and skill base.",
        "Answer text": "", "Multiple Choice Order": "5", "Multiple Choice Value": "5",
        "Multiple Choice Label": "Strongly Agree", "Submit Date": "2026-03-28 23:59:59",
    }
    if kind == "student":
        base.update({"Evaluation": "*Clinical Assessment of Student", "Question ID": "1803",
                     "Question": "<b>Student behavior</b>", "Answer text": PRIVATE})
    base.update(extra or {})
    stream = io.StringIO(newline="")
    w = csv.DictWriter(stream, fieldnames=list(base), lineterminator="\r\n")
    w.writeheader(); w.writerow(base)
    return stream.getvalue().encode(codec)


def read(raw):
    return list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"), newline="")))


@pytest.fixture
def env():
    secrets = secret_settings()
    config = OPDArchiveConfig(**secrets["opd_archive"])
    repo = RecordingGitHub()
    archive = GitHubOPDArchive(config, transport=repo)
    st.reset(secrets=secrets)
    return secrets, config, repo, archive


@pytest.mark.parametrize("kind,columns", [("educator", EDUCATOR_COLUMNS), ("student", STUDENT_COLUMNS)])
def test_exact_allowed_headers_only(kind, columns):
    out = minimize_oasis_csv(source(kind), kind)
    rows = read(out["raw"])
    assert tuple(rows[0]) == columns
    assert out["privacy"]["removed_column_count"] == 33 - len(columns)
    assert out["details"]["row_count"] == 1
    assert PRIVATE.encode() not in out["raw"]


@pytest.mark.parametrize("kind", ["educator", "student"])
@pytest.mark.parametrize("codec", ["utf-8", "utf-8-sig", "utf-16", "cp1252"])
def test_supported_encodings_and_canonical_idempotence(kind, codec):
    raw = source(kind, extra={"Evaluator": "Exámple, Avery"}, codec=codec)
    one = minimize_oasis_csv(raw, kind)
    two = minimize_oasis_csv(one["raw"], kind)
    assert two["raw"] == one["raw"]
    assert two["privacy"]["needs_minimization"] is False
    assert read(one["raw"])[0]["Evaluator"] == "Exámple, Avery"


@pytest.mark.parametrize("kind", ["educator", "student"])
def test_future_unknown_columns_never_persist(kind):
    raw = source(kind, extra={"New sensitive field": PRIVATE})
    out = minimize_oasis_csv(raw, kind)
    assert "New sensitive field" not in read(out["raw"])[0]
    assert PRIVATE.encode() not in out["raw"]


def test_educator_comment_wording_quotes_newlines_preserved():
    text = '  Helpful, "specific" feedback\nSecond line  '
    raw = source(extra={"Question ID": "172", "Question": "Please indicate this educator's strengths",
                        "Multiple Choice Value": "", "Multiple Choice Label": "", "Answer text": text})
    out = minimize_oasis_csv(raw, "educator")["raw"]
    assert read(out)[0]["Answer text"] == text
    assert read(out)[0]["Question"] == "Please indicate this educator's strengths"


def test_educator_summary_unchanged_by_column_selection():
    raw = source()
    before = educator_summary(prepare_reports([("source", raw)]))
    after = educator_summary(prepare_reports([("source", minimize_oasis_csv(raw,"educator")["raw"])]))
    assert before["rows"] == after["rows"]
    assert before["issues"] == after["issues"]
    assert after["rows"][0]["q588_mean"] == before["rows"][0]["q588_mean"]


def test_student_completion_identical_without_question_and_answer_columns():
    raw = source("student")
    before = prepare_student_assessments([("source", raw)])
    after = prepare_student_assessments([("source", minimize_oasis_csv(raw,"student")["raw"])])
    assert before == after
    assert len(after["forms"]) == 1
    assert after["forms"][0]["external_id"] == "student001"


@pytest.mark.parametrize("kind,cls", [("educator", GitHubOASISEvaluations), ("student", GitHubOASISStudentEvaluations)])
def test_all_service_writes_are_minimized_and_encrypt_then_verify(env, kind, cls):
    _, config, repo, archive = env
    service = cls(archive); raw = source(kind)
    saved = service.save(raw)
    stored = config.cipher().decrypt(repo.tree[saved["path"]])
    assert stored == minimize_oasis_csv(raw, kind)["raw"]
    assert stored != raw
    assert service.load(saved["filename"])["raw"] == stored
    assert PRIVATE.encode() not in stored
    assert not service.load(saved["filename"])["privacy"]["needs_minimization"]
    for method, url, args in repo.calls:
        if method == "PUT":
            assert PRIVATE not in repr(args)
            assert source(kind).decode("utf-8-sig") not in repr(args)


@pytest.mark.parametrize("kind,cls", [("educator", GitHubOASISEvaluations), ("student", GitHubOASISStudentEvaluations)])
def test_changes_only_to_discarded_columns_do_not_create_new_snapshot(env, kind, cls):
    _, _, repo, archive = env; service = cls(archive)
    first = service.save(source(kind))
    second = service.save(source(kind, extra={"Student Email": "different-secret@example.edu", "Who Completed": "Someone else"}))
    assert second["filename"] == first["filename"]
    assert second["action"] == "unchanged"
    assert repo.write_count == 1


def test_student_score_and_comment_changes_are_not_saved(env):
    _, _, repo, archive = env; service = GitHubOASISStudentEvaluations(archive)
    first = service.save(source("student"))
    second = service.save(source("student", extra={"Answer text": "Revised private narrative", "Multiple Choice Value": "1"}))
    assert second["filename"] == first["filename"]
    assert repo.write_count == 1


@pytest.mark.parametrize("kind,cls", [("educator", GitHubOASISEvaluations), ("student", GitHubOASISStudentEvaluations)])
def test_changed_retained_form_identity_adds_a_snapshot(env, kind, cls):
    _, _, repo, archive = env; service=cls(archive)
    service.save(source(kind))
    service.save(source(kind, extra={"Form Record": "1002"}))
    assert len(service.list_exports()["filenames"]) == 2


@pytest.mark.parametrize("kind,cls", [("student", GitHubOASISEvaluations), ("educator", GitHubOASISStudentEvaluations)])
def test_wrong_direction_rejected_before_any_network(env, kind, cls):
    _, _, repo, archive=env
    with pytest.raises(OASISArchiveError):cls(archive).save(source(kind))
    assert not repo.calls


def legacy(env, kind):
    _, config, repo, archive = env
    cls = GitHubOASISStudentEvaluations if kind=="student" else GitHubOASISEvaluations
    service=cls(archive);raw=source(kind)
    name=service._candidate_names(raw,service._inspect(raw))[0]
    path=service.path_for(name);repo.tree[path]=config.cipher().encrypt(raw);repo._commit()
    return service, raw, name, path


@pytest.mark.parametrize("kind", ["educator", "student"])
def test_legacy_load_returns_minimal_fields_without_mutating_github(env, kind):
    _, config, repo, _ = env
    service, raw, name, path=legacy(env,kind)
    old=repo.tree[path]
    out=service.load(name)
    assert out["privacy"]["needs_minimization"]
    assert out["raw"] == minimize_oasis_csv(raw,kind)["raw"]
    assert repo.tree[path] == old and repo.write_count==0
    assert config.cipher().decrypt(old)==raw


@pytest.mark.parametrize("kind", ["educator", "student"])
def test_explicit_cleanup_verifies_replacement_before_removal_but_history_remains(env, kind):
    _, config, repo, _=env
    service, raw, name, path=legacy(env,kind)
    old_head=repo.head;old=service.load(name)
    result=service.minimize_saved_export(name,expected_sha=old["sha"])
    assert path not in repo.tree
    assert len(service.list_exports()["filenames"])==1
    assert service.load(result["filename"])["raw"]==minimize_oasis_csv(raw,kind)["raw"]
    assert config.cipher().decrypt(repo.snapshots[old_head][path])==raw
    calls=[m for m,_,_ in repo.calls]
    assert calls.index("PUT") < calls.index("DELETE")


def test_cleanup_rejects_stale_review(env):
    _,_,repo,_=env;service,raw,name,path=legacy(env,"student")
    with pytest.raises(OASISArchiveError,match="changed after"):
        service.minimize_saved_export(name,expected_sha="stale")
    assert path in repo.tree and repo.write_count==0


def test_failed_replacement_never_removes_old_file(env):
    _,_,repo,_=env;service,raw,name,path=legacy(env,"student")
    expected=service.load(name)["sha"];repo.put_status=403
    with pytest.raises(OASISArchiveError):service.minimize_saved_export(name,expected_sha=expected)
    assert path in repo.tree
    assert not any(m=="DELETE" for m,_,_ in repo.calls)


def test_failed_delete_preserves_both_copies_and_can_retry(env):
    _,_,repo,_=env;service,raw,name,path=legacy(env,"educator")
    expected=service.load(name)["sha"];original=repo.request
    with patch.object(repo,"request",side_effect=lambda m,u,**kw:FakeResponse(403) if m=="DELETE" else original(m,u,**kw)):
        with pytest.raises(OASISArchiveError):service.minimize_saved_export(name,expected_sha=expected)
    assert path in repo.tree and len(service.list_exports()["filenames"])==2
    service.minimize_saved_export(name,expected_sha=expected)
    assert len(service.list_exports()["filenames"])==1


def test_already_minimal_cleanup_is_noop(env):
    _,_,repo,archive=env;service=GitHubOASISEvaluations(archive)
    a=service.save(source());n=repo.write_count
    assert service.minimize_saved_export(a["filename"],expected_sha=a["sha"])["action"]=="unchanged"
    assert repo.write_count==n


def login():
    st.secrets["evaluation_access"]={"password": PASSWORD}
    st.session_state[access.P+"password"]=PASSWORD
    access._login_callback()
    assert access.evaluation_access_is_valid()


@pytest.mark.parametrize("password", [None,"", "short", "REPLACE_WITH_A_LONG_PASSWORD"])
def test_section_fails_closed_without_configured_password(env,password):
    secrets,_,repo,_=env
    if password is not None:secrets["evaluation_access"]={"password":password}
    r=run_app({"schedule_app_mode":"OER"},secrets=secrets,repo=repo)
    assert not repo.calls and not r["downloads"]
    assert not any("upload" in k for _,k in r["widgets"])
    assert any("Setup required" in msg for _,msg in r["messages"])


def test_password_cannot_equal_encryption_key_or_token(env):
    secrets,_,_,_=env
    for value in (secrets["opd_archive"]["encryption_key"],"a-very-long-github-token"):
        st.secrets["opd_archive"]["github_token"]="a-very-long-github-token"
        st.secrets["evaluation_access"]={"password":value}
        assert not access._configured_password()


def test_wrong_password_sets_delay_no_cached_data_and_no_password_retained(env):
    st.secrets["evaluation_access"]={"password":PASSWORD}
    st.session_state.update({access.P+"password":"wrong", "oasis_workflow_prepared":{"private":1}, "teaching_zip":b"KEEP"})
    access._login_callback()
    assert not access.evaluation_access_is_valid()
    assert access.P+"password" not in st.session_state
    assert "oasis_workflow_prepared" not in st.session_state
    assert "teaching_zip" not in st.session_state
    assert st.session_state[access.P+"retry_at"]>time.time()


def test_successful_login_no_plain_password_and_lock_clears_oer_and_pts(env):
    login()
    assert PASSWORD not in repr(st.session_state)
    st.session_state.update({"_oasis_archive_loaded":b"old", "oer_report":b"old", "oasis_student_archive_loaded":b"old", "teaching_scan":{"keep":1}})
    access.lock_evaluation_records()
    assert not access.evaluation_access_is_valid()
    assert "teaching_scan" not in st.session_state
    assert all(k not in st.session_state for k in ("_oasis_archive_loaded","oer_report","oasis_student_archive_loaded"))


@pytest.mark.parametrize("offset,field", [(1801,"last_seen"),(8*3600+1,"issued")])
def test_auth_expires_on_idle_or_absolute_time(env,offset,field):
    login();st.session_state[access.P+field]=time.time()-offset
    assert not access.evaluation_access_is_valid()


def test_password_rotation_invalidates_old_authorization(env):
    login();st.secrets["evaluation_access"]["password"]=PASSWORD+"-rotated"
    assert not access.evaluation_access_is_valid()


def test_new_browser_session_does_not_inherit_login(env):
    secrets,_,_,_=env;login();st.reset(secrets=secrets)
    assert not access.evaluation_access_is_valid()


@pytest.mark.parametrize("module", ["oasis_workflow","oasis_student_evaluations","oasis_evaluation_archive","oasis_educator_reports"])
def test_legacy_and_nested_section_entrypoints_are_guarded(env,module):
    from importlib import import_module
    _,_,repo,_=env
    with patch("requests.request", repo.request):
        import_module("schedule_app.sections."+module).render()
    assert not repo.calls and not st.downloads
    assert not any("upload" in key for _,key in st.widget_keys)


def test_other_sections_do_not_require_eval_password(env):
    secrets,_,repo,_=env
    result=run_app({"schedule_app_mode":"Instructions"},secrets=secrets,repo=repo)
    assert not any("locked" in msg or "evaluation_access" in msg for _,msg in result["messages"])


def test_old_menu_state_migrates_to_locked_new_section(env):
    secrets,_,repo,_=env
    r=run_app({},secrets=secrets,state={"schedule_app_mode":"OASIS Evaluations"},repo=repo)
    assert r["state"]["schedule_app_mode"]=="OER"
    assert not repo.calls


def test_only_intended_runtime_files_changed():
    baseline=ROOT.parent/"baseline_package"
    if not baseline.exists():pytest.skip("Local release-comparison check only")
    allowed={"schedule_app/services/teaching_evaluations.py","schedule_app/sections/preceptor_oasis_links.py","app_sch_2026.py","schedule_app/services/oasis_evaluations.py",
             "schedule_app/services/oasis_student_evaluations.py","schedule_app/services/oasis_workflow.py",
             "schedule_app/sections/oasis_workflow.py","schedule_app/sections/oasis_student_evaluations.py",
             "schedule_app/sections/oasis_evaluation_archive.py","schedule_app/sections/oasis_educator_reports.py"}
    for old in baseline.rglob("*.py"):
        rel=old.relative_to(baseline)
        if str(rel).startswith("tests/") or str(rel) in allowed:continue
        assert (ROOT/rel).read_bytes()==old.read_bytes(),str(rel)
    assert (ROOT/"requirements.txt").read_bytes()==(baseline/"requirements.txt").read_bytes()


def test_teaching_summary_now_requires_protected_access(env):
    secrets,_,repo,_=env
    r=run_app({"schedule_app_mode":"PTS"},secrets=secrets,repo=repo)
    assert any("PTS — locked" in m for _,m in r["messages"])
    assert not repo.calls and not r["downloads"]


def test_privacy_review_does_nothing_while_locked(env):
    from schedule_app.sections import evaluation_privacy as ui
    _,_,repo,archive=env
    ui.render(archive)
    assert not repo.calls and not st.widget_keys


def test_privacy_review_scan_only_lists_metadata_and_does_not_write(env):
    from schedule_app.sections import evaluation_privacy as ui
    _,_,repo,archive=env;_,raw,name,path=legacy(env,"student")
    login();st.values[ui.P+"scan"]=True
    ui.render(archive)
    assert repo.write_count==0
    rows=st.session_state[ui.P+"rows"]
    assert len(rows)==1 and rows[0]["columns_to_remove"]==22
    assert PRIVATE not in repr(rows)


def test_privacy_cleanup_requires_its_exact_confirmation(env):
    from schedule_app.sections import evaluation_privacy as ui
    _,_,repo,archive=env;_,raw,name,path=legacy(env,"student")
    login();st.values[ui.P+"scan"]=True
    ui.render(archive)
    st.values={ui.P+"apply":True}
    ui.render(archive)
    assert path in repo.tree and repo.write_count==0
