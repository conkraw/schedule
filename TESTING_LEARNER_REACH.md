# Historical Learner Reach test record

The current suite has 108 passing tests; see UPDATE_TEACHING_CONTRIBUTORS.md. Zero-only report expectations described below have been replaced with exclusion tests, while raw availability checks remain.

# Learner Reach validation

Run from the project root: `python -m unittest discover -s tests -v`.

90 tests passed locally. New tests cover 80% Learner Reach, multiple simultaneous students, denominator deduplication, blank-only provider listings, zero denominators, exact custom boundaries, July mode, monthly availability, CSV field preservation, no-tilde/nonclinical exclusions, ambiguous work types, no student identifiers in reports, stale scans, corruption checks and simulated interface generation. Prior date tests were updated only for the intentional additional zero-teaching providers and appended CSV columns.

Two additional local checks compared student-weighted overall/monthly/work-type/daily counts on the two supplied OPD examples to the preceding custom-dates implementation; counts were unchanged. Raw inputs are not bundled in this package.

Word previews and representative longer outputs were rendered for visual review. Simulations do not establish that the live deployment works with its current configuration. No live GitHub repository was read or modified by these tests.
