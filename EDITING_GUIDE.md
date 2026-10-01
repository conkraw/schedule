# Which file should I edit?

## Current interface

| Change | File |
|---|---|
| Sidebar names/order and entrypoint | `app_sch_2026.py` |
| Default PTS reporting controls, warnings, progress, downloads | `schedule_app/sections/preceptor_teaching_summary.py` |
| PTS Matching task dropdown | `schedule_app/sections/pts_matching.py` |
| Shared OPD loading and reporting-date context | `schedule_app/sections/pts_workspace.py` |
| Navigation callbacks / retained PTS widget preferences | `schedule_app/sections/pts_navigation.py` |
| Missing preceptor username entry / saved links | `schedule_app/sections/preceptor_oasis_links.py` |
| Missing student name entry / saved matches | `schedule_app/sections/student_name_matches.py` |
| Ignore / restore OPD student labels | `schedule_app/sections/ignored_student_entries.py` |
| Minimum shifts control | `schedule_app/sections/assessment_settings.py` |
| Assessment checks / warning counts | `schedule_app/sections/assessment_completion.py` |
| Custom dates and saved-preset controls | `schedule_app/sections/reporting_date_controls.py` |
| Password gate shared by OER / PTS / PTS Matching | `schedule_app/services/evaluation_access.py` |
| OER uploads and encrypted summary workflow | `schedule_app/sections/oasis_workflow.py` |

## Calculations, reports and storage

| Change | File |
|---|---|
| Email mappings, work-type grouping, explicit name aliases | `schedule_app/settings.py` |
| OPD extraction / teaching aggregates | `schedule_app/services/teaching_analysis.py` |
| Simplified hours / report CSV fields | `schedule_app/services/educational_time.py` |
| Learner Reach denominator and inclusion rules | `schedule_app/services/learner_reach.py` |
| Outpatient-over-nursery priority | `schedule_app/services/teaching_priority.py` |
| Name designation normalization | `schedule_app/services/student_name_matching.py` |
| Assessment eligibility, numerator and form counting | `schedule_app/services/assessment_completion.py` |
| Chair Word report wording/layout | `schedule_app/reports/chair_summary.py` |
| Individual teaching Word report wording/layout | `schedule_app/reports/individual_teaching.py` |
| Linked OASIS question/comment Word section | `schedule_app/reports/preceptor_evaluations.py` |
| Teaching ZIP contents / progress stages | `schedule_app/reports/teaching_export.py` |
| Once-per-build shared teaching tables | `schedule_app/reports/teaching_batch.py` |
| GitHub OPD encryption/read/write | `schedule_app/services/opd_archive.py` |
| Encrypted preceptor links | `schedule_app/services/preceptor_oasis_links.py` |
| Encrypted student matches | `schedule_app/services/student_assessment_links.py` |
| Encrypted ignored entries | `schedule_app/services/ignored_student_entries.py` |
| Encrypted minimum-shifts setting | `schedule_app/services/assessment_settings.py` |
| Saved date presets | `schedule_app/services/reporting_presets.py` |
| Allowed evaluation-source columns | `schedule_app/services/oasis_privacy.py` |

Other sidebar screens remain one module each in `schedule_app/sections/`.
Student schedule templates are in `services/student_schedules.py`; the primary
preceptor/Power Automate workbook is in `services/primary_preceptors.py`.

Keep calculation changes in the relevant service rather than only changing a
Word label. Preserve validation, exact-date boundaries, encrypted save read-back,
source revision checks, and password guards. Shared report data in teaching_batch
is build-local, not an application-wide cache.

Start with UPDATE_PTS_SIMPLIFIED.md for the current update. Earlier release guides
remain historical references, not instructions to replace newer code with older
modules. Never change your encryption key merely to deploy a code update.
