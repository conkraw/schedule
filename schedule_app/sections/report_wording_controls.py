"""Load presentation preferences once per session/minute or explicitly per build."""
from functools import wraps
from time import monotonic
import streamlit as st
from schedule_app.services.opd_archive import GitHubOPDArchive, get_opd_archive_config, OPDArchiveError
from schedule_app.services.report_wording import GitHubReportWording, report_wording_context, wording_signature

KEY = '_report_wording_cache'


def clear_report_outputs():
    from schedule_app.sections.reporting_date_controls import clear_teaching_downloads
    clear_teaching_downloads()
    # Outputs using editable schedule instructions must also be regenerated.
    for key in ('opd_generated_master', 'individual_schedule_zip', 'individual_preceptor_report',
                'individual_preceptor_preview', 'individual_missing_emails'):
        st.session_state.pop(key, None)


def load_saved_report_wording(client=None, *, refresh=False):
    if client is None:
        try:
            configured = dict(st.secrets.get('opd_archive', {}))
        except Exception:
            configured = {}
        if not configured:
            # General scheduling still works before archive setup, as before.
            # Admin never takes this path: it explicitly requires a configured client.
            st.caption('Report wording: built-in defaults (GitHub archive is not configured).')
            return None
        client = GitHubOPDArchive(get_opd_archive_config())
    scope = client.config.signature()
    cached = st.session_state.get(KEY)
    now = monotonic()
    if not refresh and cached and cached.get('scope') == scope and 0 <= now - cached.get('loaded_at', 0) < 60:
        return cached['snapshot']
    try:
        snapshot = GitHubReportWording(client).load()
    except OPDArchiveError:
        st.session_state.pop(KEY, None)
        clear_report_outputs()
        raise
    if cached and (cached.get('scope') != scope or wording_signature(cached['snapshot']) != wording_signature(snapshot)):
        clear_report_outputs()
    st.session_state[KEY] = {'snapshot': snapshot, 'scope': scope, 'loaded_at': now}
    return snapshot


def accept_saved_report_wording(snapshot):
    clear_report_outputs()
    st.session_state[KEY] = {'snapshot': snapshot, 'scope': snapshot['scope'], 'loaded_at': monotonic()}


def with_saved_report_wording(function):
    """General document/schedule screens: one fresh snapshot for one render/build."""
    @wraps(function)
    def render(*args, **kwargs):
        try:
            snapshot = load_saved_report_wording(refresh=True)
        except OPDArchiveError as exc:
            st.error('Report wording could not be loaded. ' + str(exc))
            return
        with report_wording_context(snapshot):
            return function(*args, **kwargs)
    return render
