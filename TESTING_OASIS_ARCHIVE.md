# OASIS archive test results

Local run: September 29, 2026.

**360 unittest tests passed**: the existing 297 tests plus 63 OASIS-specific tests.
The tests use simulated Streamlit widgets and an in-memory GitHub Contents API
transport. No live repository, remote file, or Streamlit deployment was changed.

Run the full suite from the application directory:

```bash
python -m unittest discover -s tests -v
```

Run only the new synthetic OASIS tests:

```bash
python -m unittest discover -s tests -p 'test_oasis_evaluations.py' -v
```

New coverage includes original-byte encryption/decryption, duplicate uploads,
changed exports with the same dates, separate-session recovery, large-file raw
responses, old/wrong keys, ciphertext corruption and filename binding, malformed
CSV handling, encoding/BOM preservation, quoted multiline comments, response
limits, no plaintext in public storage paths or UI messages, save failures,
retry behavior, stale-download clearing, and isolation from OPD/preset data.

Additional local check with the user's uploaded `oasis_eval_export (1).csv`:

- Original size: 2,147,013 bytes.
- 4,162 question-response rows; all 33 columns retained.
- Course-date coverage: March 16 through August 28, 2026.
- Original CSV and decrypted download were byte-for-byte identical.
- Re-uploading the same bytes did not create a second commit.
- Simulated dropdown reload/download returned exactly the same original bytes.
- The raw Contents API fallback was exercised for the encrypted file.

The real sample is intentionally not included in the deliverable ZIPs. All
new packaged test fixtures use invented data. No original report, teaching,
OPD, date-preset, settings, or dependency module was modified.

Only these three runtime files differ from the prior full report-check package:

1. app_sch_2026.py — one new sidebar entry.
2. schedule_app/sections/oasis_evaluation_archive.py — new.
3. schedule_app/services/oasis_evaluations.py — new.

Python compilation also passed. Real Streamlit Cloud and real GitHub network
behavior remain untested in this conversation; deploy and verify using the
application's verified-save status.
