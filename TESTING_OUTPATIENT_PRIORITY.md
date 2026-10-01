# Outpatient priority validation

163 automated tests passed using `python -m unittest discover -s tests -v`.

32 new tests are in `tests/test_outpatient_priority.py`. Existing strict conflict tests still block unrelated settings; the previous NYES/nursery blocking example was changed to NYES/Complex Care because nursery/clinic is now explicitly resolved.

Both supplied sample OPDs were tested separately against independent sets of provider/student/date/shift and clinical provider/date/shift records after applying the new exception. Source-byte and encrypted-archive preservation were verified.

The synthetic chair report rendered to two clean pages, including both category pies; the individual report rendered to one clean page. These internal QA documents are not production archive reports and are not included in the release.

Streamlit/GitHub calls were simulated. No live deployment, repository, or archive was modified.
