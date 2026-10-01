"""Navigation and password-boundary regression tests. No live GitHub calls."""
import ast
import hashlib
import time
from importlib import import_module
from unittest.mock import MagicMock, patch
import pytest

from helpers import ROOT, st, run_app, secret_settings, login_for_test
from test_oasis_evaluations import RecordingGitHub
from schedule_app.services import evaluation_access as access
from schedule_app.sections import reporting_date_controls as pts_dates
from schedule_app.sections import oasis_date_controls as oer_dates

PUBLIC = [
    "Instructions", "Format OPD + Summary", "Create Student Schedule", "OPD Check",
    "Create Individual Schedules", "OPD Archive", "OPD MD PA Conflict Detector",
    "Shift Availability Tracker",
]
PASSWORD = "long-test-only-passphrase-for-OER-PTS"


@pytest.fixture
def env():
    config = secret_settings()
    config["evaluation_access"] = {"password": PASSWORD}
    st.reset(secrets=config)
    return config, RecordingGitHub()


def authenticate():
    st.session_state[access.P + "password"] = PASSWORD
    access._login_callback()
    assert access.evaluation_access_is_valid()


def menu():
    tree = ast.parse((ROOT / "app_sch_2026.py").read_text())
    return next(ast.literal_eval(n.value) for n in tree.body if isinstance(n, ast.Assign)
                and any(isinstance(t, ast.Name) and t.id == "SECTIONS" for t in n.targets))


def test_sidebar_order_and_module_targets():
    sections = menu()
    assert list(sections) == PUBLIC + ["PTS Matching", "OER", "PTS"]
    assert sections["OER"] == "oasis_workflow"
    assert sections["PTS"] == "preceptor_teaching_summary"


@pytest.mark.parametrize("old,new", [
    ("Evaluation Records", "OER"), ("OASIS Evaluations", "OER"),
    ("OASIS Evaluation Archive", "OER"), ("OASIS Educator Reports", "OER"),
    ("Preceptor Teaching Summary", "PTS"),
])
def test_old_session_label_migrates_before_widget(env, old, new):
    settings, repo = env
    result = run_app({}, secrets=settings, repo=repo, state={"schedule_app_mode": old})
    assert result["state"]["schedule_app_mode"] == new
    assert any(text == new + " — locked" for _, text in result["messages"])
    assert not repo.calls and not result["downloads"]


@pytest.mark.parametrize("mode", ["OER", "PTS"])
@pytest.mark.parametrize("password", [None, "", "short", "REPLACE_WITH_YOUR_OWN_LONG_PASSWORD", PASSWORD])
def test_locked_entry_never_reads_writes_or_exposes_data(env, mode, password):
    settings, repo = env
    if password is None:
        settings.pop("evaluation_access")
    else:
        settings["evaluation_access"]["password"] = password
    prior = {"teaching_scan": {"private": True}, "teaching_zip": b"private report",
             "assessment_completion_inputs": {"private": True}, "oasis_combined_receipt": {"private": True},
             "teaching_oasis_catalog": {"private": True}, "opd_generated_master": b"keep scheduling"}
    result = run_app({"schedule_app_mode": mode, "teaching_load_archives": True,
                      "teaching_build_zip": True}, secrets=settings, state=prior, repo=repo)
    assert not repo.calls and not result["downloads"]
    assert all(not str(key).startswith(access.PROTECTED_STATE_PREFIXES) for key in result["state"])
    assert result["state"]["opd_generated_master"] == b"keep scheduling"
    assert not any("upload" in key or key == "teaching_load_archives" for _, key in result["widgets"])
    assert any(text == mode + " — locked" for _, text in result["messages"])


@pytest.mark.parametrize("mode", PUBLIC)
def test_other_menu_sections_do_not_require_section_password(env, mode):
    settings, repo = env
    settings.pop("evaluation_access")
    result = run_app({"schedule_app_mode": mode}, secrets=settings, repo=repo)
    assert not any(key.startswith(access.P) for _, key in result["widgets"])
    assert not any(text.endswith(" — locked") for _, text in result["messages"])


def test_one_login_unlocks_both_sections_and_does_not_retain_password(env):
    authenticate()
    for mode in ("OER", "PTS", "OER"):
        assert access.require_evaluation_access(section_name=mode, lock_key="lock_" + mode)
    assert PASSWORD not in repr(st.session_state)
    assert not any(key == access.P + "password" for _, key in st.widget_keys)
    assert all(label == "Lock OER / PTS" for label, _ in st.widget_keys)


def test_lock_both_sections_clears_protected_data_not_schedules_or_repository(env):
    settings, repo = env
    authenticate()
    for prefix in access.PROTECTED_STATE_PREFIXES:
        st.session_state[prefix + "test"] = b"private data"
    st.session_state.update({"individual_schedule_zip": b"keep", "opd_archive_loaded": b"keep",
                             "schedule_app_mode": "PTS"})
    access.lock_evaluation_records()
    assert not access.evaluation_access_is_valid()
    assert not any(str(key).startswith(access.PROTECTED_STATE_PREFIXES) for key in st.session_state)
    assert st.session_state["individual_schedule_zip"] == b"keep"
    assert st.session_state["opd_archive_loaded"] == b"keep"
    assert not repo.calls


@pytest.mark.parametrize("field,offset", [("last_seen", 1801), ("issued", 28801)])
@pytest.mark.parametrize("mode", ["OER", "PTS"])
def test_timeout_clears_both_sections_before_render(env, field, offset, mode):
    authenticate()
    st.session_state.update({"teaching_zip": b"private", "oasis_combined_receipt": {"private": True},
                             access.P + field: time.time() - offset})
    assert not access.require_evaluation_access(section_name=mode)
    assert "teaching_zip" not in st.session_state
    assert "oasis_combined_receipt" not in st.session_state


def test_password_rotation_locks_both_and_clears_old_reports(env):
    authenticate()
    st.session_state["teaching_zip"] = b"private"
    st.secrets["evaluation_access"]["password"] += "changed"
    assert not access.require_evaluation_access(section_name="PTS")
    assert not access.evaluation_access_is_valid()
    assert "teaching_zip" not in st.session_state


def test_older_oer_only_login_does_not_survive_scope_expansion(env):
    st.session_state.update({
        access.P + "auth": hashlib.sha256(("evaluation-section-v1\0" + PASSWORD).encode()).hexdigest(),
        access.P + "issued": time.time(), access.P + "last_seen": time.time(),
        "teaching_zip": b"old unprotected report", "oasis_combined_prepared": {"old": True},
    })
    assert not access.require_evaluation_access(section_name="PTS")
    assert "teaching_zip" not in st.session_state
    assert "oasis_combined_prepared" not in st.session_state


CALLBACKS = [
    (pts_dates.remember_period_inputs, ()), (pts_dates.refresh_saved_presets, ()),
    (pts_dates.load_selected_preset, ("preset-id",)), (pts_dates.save_current_preset, ()),
    (pts_dates.delete_selected_preset, ("preset-id", "confirm")),
    (oer_dates._apply_period, ()), (oer_dates._load_preset, ()),
    (oer_dates._refresh_presets, (MagicMock(),)),
    (oer_dates._save_preset, (MagicMock(), None, "confirm")),
    (oer_dates._delete_preset, (MagicMock(), "preset-id", "confirm")),
]


@pytest.mark.parametrize("callback,args", CALLBACKS, ids=[c.__name__ for c,_ in CALLBACKS])
@pytest.mark.parametrize("expired", [False, True])
def test_callbacks_cannot_operate_before_locked_or_expired_render(env, callback, args, expired):
    if expired:
        authenticate()
        st.session_state[access.P + "last_seen"] = time.time() - 1801
    st.session_state["teaching_scan"] = {"private": True}
    st.session_state["oasis_combined_dates_snapshot"] = {"private": True}
    for arg in args:
        if isinstance(arg, MagicMock): arg.reset_mock()
    with patch("requests.request", side_effect=AssertionError("Unexpected network request")) as request:
        assert callback(*args) is None
        request.assert_not_called()
    for arg in args:
        if isinstance(arg, MagicMock): assert not arg.mock_calls
    assert "teaching_scan" not in st.session_state
    assert "oasis_combined_dates_snapshot" not in st.session_state


def test_authorized_callback_runs_once_with_arguments(env):
    authenticate()
    work = MagicMock(return_value="done")
    guarded = access.protected_evaluation_callback(work)
    assert guarded(1, scope="example") == "done"
    work.assert_called_once_with(1, scope="example")


@pytest.mark.parametrize("module,mode", [("oasis_workflow", "OER"), ("preceptor_teaching_summary", "PTS")])
def test_direct_section_entrypoint_still_checks_password(env, module, mode):
    _, repo = env
    with patch("requests.request", repo.request):
        import_module("schedule_app.sections." + module).render()
    assert not repo.calls and not st.downloads
    assert any(text == mode + " — locked" for _, text in st.messages)


def test_authenticated_pts_reaches_controls_without_another_password(env):
    settings, repo = env
    result = run_app({"schedule_app_mode": "PTS"}, secrets=settings, repo=repo, evaluation_login=True)
    assert any(text == "PTS" for _, text in result["messages"])
    assert any(key == "teaching_period_start" for _, key in result["widgets"])
    assert not any(label.endswith("password") for label, _ in result["widgets"])


def test_new_sessions_do_not_inherit_shared_authorization(env):
    settings, _ = env
    authenticate()
    st.reset(secrets=settings)
    assert not access.evaluation_access_is_valid()
