"""Editable report presentation, separate from calculations and source records.

Bundled wording lives in resources/report_wording_defaults.json. Only explicit
Admin overrides are saved, encrypted, to GitHub. A ContextVar binds one validated
snapshot to one synchronous report build; no cross-user data cache is used.
"""
from __future__ import annotations
import base64
from contextlib import contextmanager
from contextvars import ContextVar
from copy import deepcopy
from functools import lru_cache, wraps
import hashlib
import json
from pathlib import Path
import re
from urllib.parse import quote
from cryptography.fernet import InvalidToken
from schedule_app.services.opd_archive import OPDArchiveError

KIND = 'report_wording'
VERSION = 1
FILENAME = 'report_wording.json.enc'
MAX_RAW_BYTES = 256 * 1024
MAX_ENCRYPTED_BYTES = 400 * 1024
MAX_FIELD_CHARS = 16000
TOKEN = re.compile(r'\{([a-z][a-z0-9_]*)\}')
ACTIVE = ContextVar('report_wording_snapshot', default=None)
STYLE_GROUPS = ('individual', 'chair', 'instructions', 'opd_changes', 'assignment_summary')
FONT_CHOICES = ('Calibri', 'Aptos', 'Arial', 'Cambria', 'Times New Roman')
STYLE_LIMITS = {'body_size': (9, 13), 'table_size': (8, 11), 'title_size': (16, 28), 'note_size': (8, 11)}


@lru_cache(maxsize=1)
def _catalog_json():
    # Cache only this public, immutable JSON string, never a user's overrides.
    path = Path(__file__).resolve().parents[1] / 'resources' / 'report_wording_defaults.json'
    try:
        return path.read_text(encoding='utf-8')
    except (OSError, UnicodeError):
        raise OPDArchiveError('Report wording defaults are missing or unreadable. Upload schedule_app/resources/report_wording_defaults.json from the update.') from None


@lru_cache(maxsize=1)
def _public_catalog():
    return json.loads(_catalog_json())


def report_catalog():
    return deepcopy(_public_catalog())


def empty_wording():
    return {'kind': KIND, 'version': VERSION, 'texts': {}, 'styles': {}}


def _text_is_safe(text):
    return (isinstance(text, str) and len(text) <= MAX_FIELD_CHARS
            and not any((ord(c) < 32 and c not in '\t\r\n') or 0xD800 <= ord(c) <= 0xDFFF for c in text))


def validate_field(key, text):
    field = _public_catalog()['fields'].get(key)
    if field is None or not _text_is_safe(text):
        raise OPDArchiveError('Unknown report field or unsupported text. No wording was saved.')
    if field['required'] and not text.strip():
        raise OPDArchiveError('This report heading or instruction cannot be blank. Restore its default instead.')
    tokens = set(TOKEN.findall(text))
    residue = TOKEN.sub('', text)
    if '{' in residue or '}' in residue or tokens != set(field['tokens']):
        allowed = ', '.join('{' + t + '}' for t in field['tokens']) or 'none'
        raise OPDArchiveError('Keep the automatic placeholders exactly as shown. Required placeholders: ' + allowed + '.')
    return text


def validate_wording(data):
    if (not isinstance(data, dict) or set(data) != {'kind', 'version', 'texts', 'styles'}
            or data.get('kind') != KIND or type(data.get('version')) is not int or data['version'] != VERSION
            or not isinstance(data['texts'], dict) or not isinstance(data['styles'], dict)):
        raise OPDArchiveError('Report wording has an unsupported format. No defaults were substituted.')
    for key, value in data['texts'].items():
        validate_field(key, value)
    for group, style in data['styles'].items():
        if group not in STYLE_GROUPS or not isinstance(style, dict):
            raise OPDArchiveError('Unknown report appearance settings.')
        if set(style) - (set(STYLE_LIMITS) | {'font'}):
            raise OPDArchiveError('Unknown report appearance field.')
        for key, value in style.items():
            if key == 'font':
                if value not in FONT_CHOICES:
                    raise OPDArchiveError('Select a supported report font.')
            elif type(value) not in (int, float) or not STYLE_LIMITS[key][0] <= value <= STYLE_LIMITS[key][1]:
                raise OPDArchiveError('Report font size is outside its supported range.')
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True).encode('utf-8')
    if len(raw) > MAX_RAW_BYTES:
        raise OPDArchiveError('Saved report wording is too large. Shorten the instructions before saving.')
    return deepcopy(data)


def wording_data(snapshot):
    if snapshot is None:
        return empty_wording()
    return validate_wording({k: snapshot[k] for k in ('kind', 'version', 'texts', 'styles')})


def wording_signature(snapshot):
    return hashlib.sha256(json.dumps(wording_data(snapshot), sort_keys=True, ensure_ascii=False,
                                     separators=(',', ':')).encode('utf-8')).hexdigest()


def effective_text(snapshot, key):
    fields = _public_catalog()['fields']
    if key not in fields:
        raise OPDArchiveError('A report wording field is missing from the installed defaults. Install the complete update.')
    return (snapshot or {}).get('texts', {}).get(key, fields[key]['default'])


def render_text(text, values):
    needed = set(TOKEN.findall(text))
    if not needed.issubset(values):
        raise OPDArchiveError('An automatic report value is unavailable; no incomplete report was created.')
    # No eval, format-expression evaluation, attribute access, imports, or HTML.
    return TOKEN.sub(lambda match: str(values[match.group(1)]), text)


def report_text(key, **values):
    return render_text(effective_text(ACTIVE.get(), key), values)


def active_appearance(group):
    return dict((ACTIVE.get() or {}).get('styles', {}).get(group, {}))


@contextmanager
def report_wording_context(snapshot):
    token = ACTIVE.set(wording_data(snapshot))
    try:
        yield
    finally:
        ACTIVE.reset(token)


def with_report_wording(function):
    """Optional report_wording= argument; inherit an enclosing ZIP's snapshot."""
    @wraps(function)
    def wrapped(*args, **kwargs):
        if 'report_wording' not in kwargs:
            return function(*args, **kwargs)
        snapshot = kwargs.pop('report_wording')
        with report_wording_context(snapshot):
            return function(*args, **kwargs)
    return wrapped


def _unique_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError('Duplicate setting key')
        result[key] = value
    return result


class GitHubReportWording:
    """Ciphertext-only storage with optimistic concurrency and verified readback."""
    def __init__(self, archive):
        self.archive, self.config = archive, archive.config
        self.path = f'{self.config.folder}/{FILENAME}'
        self.route = '/contents/' + quote(self.path, safe='/')
        self.cipher = self.config.cipher()

    def load(self, *, commit=None):
        commit = commit or self.archive._head()
        meta = self.archive._request('GET', self.route, params={'ref': commit}, missing_ok=True)
        if meta is None:
            return {**empty_wording(), 'sha': None, 'commit': commit, 'scope': self.config.signature()}
        if (not isinstance(meta, dict) or meta.get('type') != 'file' or meta.get('target')
                or meta.get('submodule_git_url') or meta.get('path', self.path) != self.path
                or meta.get('encoding') != 'base64' or type(meta.get('size')) is not int
                or not 0 < meta['size'] <= MAX_ENCRYPTED_BYTES
                or not isinstance(meta.get('content'), str) or len(meta['content']) > 2 * MAX_ENCRYPTED_BYTES):
            raise OPDArchiveError('The saved report wording is not a supported encrypted file. No setting was changed.')
        try:
            token = base64.b64decode(''.join(meta['content'].split()), validate=True)
            sha = hashlib.sha1(b'blob ' + str(len(token)).encode() + b'\0' + token).hexdigest()
            if len(token) != meta['size'] or sha != meta.get('sha'):
                raise ValueError('Content mismatch')
            raw = self.cipher.decrypt(token)
            if len(raw) > MAX_RAW_BYTES:
                raise ValueError('Too large')
            data = validate_wording(json.loads(raw.decode('utf-8'), object_pairs_hook=_unique_keys))
        except (InvalidToken, ValueError, TypeError, UnicodeError, RecursionError):
            raise OPDArchiveError('Report wording could not be verified/decrypted. Check the existing key or file; '
                                  'nothing was overwritten and no fallback report was generated.') from None
        return {**data, 'sha': sha, 'commit': commit, 'scope': self.config.signature()}

    def save(self, data, *, expected):
        data = validate_wording(data)
        if (not isinstance(expected, dict) or expected.get('scope') != self.config.signature()
                or 'sha' not in expected):
            raise OPDArchiveError('Load report wording before editing it.')
        current = self.load()
        if current['sha'] != expected['sha']:
            raise OPDArchiveError('Report wording changed in another session. Reload saved wording, review it, then save again.')
        # Store overrides only. Resetting restores the bundled current defaults.
        data['texts'] = {k: v for k, v in data['texts'].items()
                         if v != _public_catalog()['fields'][k]['default']}
        data['styles'] = {k: v for k, v in data['styles'].items() if v}
        if wording_data(current) == data:
            return current
        raw = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')
        body = {'message': 'Update encrypted report presentation settings', 'branch': self.config.branch,
                'content': base64.b64encode(self.cipher.encrypt(raw)).decode('ascii')}
        if current['sha'] is not None:
            body['sha'] = current['sha']
        result = self.archive._request('PUT', self.route, body=body)
        commit = result.get('commit', {}).get('sha') if isinstance(result, dict) else None
        if not commit:
            raise OPDArchiveError('The wording save was not confirmed. Reload to verify its state.')
        verified = self.load(commit=commit)
        if wording_data(verified) != data:
            raise OPDArchiveError('The wording save could not be verified. Reload before using it.')
        return verified
