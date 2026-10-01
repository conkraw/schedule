# Learner Reach reference

**Latest inclusion rule:** see UPDATE_TEACHING_CONTRIBUTORS.md. The latest patch hides zero-assignment contributors while retaining the availability needed for the percentages. The original installation section below describes the earlier addition of Learner Reach; use the latest five-file patch when Learner Reach already works.

## Install into your working modular app

Use `Schedule_App_Learner_Reach_Update.zip` for the modular app that already has custom reporting dates.

1. Extract the ZIP.
2. Merge its `schedule_app` directory into the directory of the same name in your existing app repository. Replace the five matching Python files and add the new `services/learner_reach.py` file. Do not delete the rest of your existing folder.
3. Restart the Streamlit app.
4. Open **Preceptor Teaching Summary**, click **Load / refresh archived OPDs**, choose your dates and report label, and click **Create teaching reports ZIP**.

The first refresh is required because earlier scans discarded provider listings without a student. Existing encrypted OPDs already contain that information; they do not need re-uploading or re-encryption.

Leave `app_sch_2026.py`, `schedule_app/settings.py`, your email/name/work-type mappings, `requirements.txt`, Streamlit Secrets and your encryption key unchanged. No app password has been added. Your student scheduling, Power Automate workbook and encrypted archive functions are unchanged.

The alternative `Schedule_App_Modular_Learner_Reach.zip` contains the complete app. Preserve any mappings you have changed in your deployed `settings.py` before replacing the complete folder. Never put keys, tokens, decrypted OPDs or generated staff reports into the public repository.

## The name: Learner Reach

**Learner Reach is the percentage of a preceptor's recorded OPD clinical shifts that include at least one student.** It is a descriptive, schedule-based measure, not a teaching-quality score or a target.

```
Learner Reach (%) = shifts with at least one student / all recorded clinical shifts x 100
```

The denominator includes both assigned and unassigned shifts. It is NOT just the shifts with a blank student field.

Each distinct preceptor + actual date + AM/PM counts as one clinical shift. The same preceptor appearing in multiple rows does not create extra clock time. An AM and a PM on the same date count as two shifts. All hours in this report are four-hour equivalents; the OPD does not verify actual duration or attendance.

### Example

| Recorded activity | Shifts | Four-hour equivalents |
|---|---:|---:|
| All recorded clinical shifts | 10 | 40 |
| Shifts with at least one student | 8 | 32 |
| Shifts without a student | 2 | 8 |

Learner Reach is **80%**. If one of those eight teaching shifts had two different students, the original educational-effort measure is nine student-shifts / 36 student-weighted educational hours. Learner Reach stays 80%, because only eight of the ten clinical shifts included students.

## How an OPD cell is interpreted

Use the existing **Names around '~'** selector for the actual name order in your archive. A blank student side is recognized in either supported order.

| Example in Preceptor ~ Student format | Interpretation |
|---|---|
| `Example, Avery ~ Learner One` | One recorded clinical shift with a student |
| `Example, Avery ~ ` | One recorded clinical shift without a student |
| Two rows for Avery in the same date/AM with different students | One clinical shift with students; two student-shift assignments |
| A duplicate blank row plus an assigned row for Avery in the same date/AM | One clinical shift with students, not two shifts |
| Empty cell | No recorded shift |
| `CLOSED ~ `, `OFF ~ ` or another recognized nonclinical label | Excluded |
| Nonempty cell without `~` | Excluded and flagged by coordinate; the app does not guess its meaning |

Leading/trailing spaces are ignored. The app uses the existing explicit name aliases; it does not infer that two differently spelled names are the same person. Student identifiers are used temporarily for the existing student-assignment deduplication but are not included in the returned scan, CSVs, Word reports or source notes.

Providers with no student assignments in the selected period are omitted from the reports. Included providers keep their non-teaching shifts in the denominator. A missing or unresolved denominator is displayed as **N/A**, not 0%.

## Report changes

The existing ZIP keeps the same filenames and adds a monthly availability CSV.

- `Pediatric_Clerkship_Educational_Effort_Summary.docx`: the chair summary includes recorded OPD hours, hours with students and Learner Reach in the overview and preceptor tables. Student-weighted educational effort and work-type sections remain.
- `Preceptor_Reports/*.docx`: one Word report per preceptor/provider label with student assignments in the selected period. It shows overall and work-type Learner Reach, recorded hours, hours with students and hours without students, followed by the original student-weighted effort detail.
- `preceptor_teaching_summary.csv` and `preceptor_teaching_by_work_type.csv`: their original columns and meanings are preserved at the beginning, with the fields below appended. Entirely zero-teaching entries are excluded.
- `preceptor_learner_reach_monthly.csv`: monthly clinical-shift metrics for participating provider/work-type entries, retaining their months without student assignments. Boundary months contain only the selected dates, not an assumed full calendar month.
- `Clinical_Shift_Review.csv`: created only when a provider's clinical half-day appears in more than one work type.
- `Report_Notes.txt` and `Archive_Sources.csv`: include the definitions, coverage limits and availability parsing diagnostics. The existing saved reporting-date files remain in custom-date ZIPs.

### New CSV columns

| Column | Definition |
|---|---|
| `recorded_clinical_shifts` | Distinct recorded provider/date/AM-or-PM sessions, assigned or unassigned |
| `recorded_clinical_hours` | Recorded clinical shifts x 4 |
| `shifts_with_students` | Distinct sessions with at least one student |
| `hours_with_students` | Shifts with students x 4; not multiplied by simultaneous students |
| `shifts_without_students` | Recorded clinical shifts minus shifts with students |
| `hours_without_students` | Shifts without students x 4 |
| `learner_reach_pct` | Percentage on a 0–100 scale: `80` means 80%, not 0.8. Blank means unavailable/review required |
| `months_scheduled` | Months with recorded clinical sessions, with or without students |
| `availability_review_shifts` | Concurrent clinical sessions whose work type cannot be uniquely assigned |
| `learner_reach_note` | Explanation when a percentage is unavailable or affected by a review issue |

`months_worked` retains its original meaning: months with **student assignments**. Only participating providers are reported; wholly unassigned providers are omitted. `no_of_shifts` and `educational_hours` still count student-weighted effort, not unique clinical shifts.

## Dates, categories and totals

Custom start/end dates still apply exactly and inclusively to BOTH numerator and denominator. The `academic_year` field still uses your report label. A custom period is not split at July 1; the optional standard July–June mode remains available.

HOPE_DRIVE, ETOWN and NYES are combined as **Academic Pediatrics**. Ward A, PSHCH Nursery, Complex Care and other work types remain separate. No setting gets a higher hour multiplier.

Overall Learner Reach uses summed shifts, not an average of each preceptor's percentage. The same rule applies to group subtotals. Generic provider/slot labels remain separate from named preceptors. As in prior reports, those literal labels contribute to overall totals; they are not verified individual clinicians.

A provider/date/half-day recorded in different work types is counted ONCE overall. Since the OPD cannot establish how much of that half-day belonged to each concurrent service, its clinical hours are placed in **Work type needs review**. The affected work-type hours exclude those unresolved sessions and the affected category percentages are withheld as N/A. Overall Learner Reach still counts the shift once. The existing student-weighted effort retains its original assignment-based category rules, so educational hours and clinical hours must not be substituted for one another.

## Important interpretation limit

This is the proportion of **clinical sessions recorded in the archived OPDs**, not necessarily the proportion of a preceptor's entire clinical job. Blank student fields do not prove a preceptor was available to accept another learner, declined teaching, or had unused capacity. Other learners, unlisted work, missing rotations, actual attendance and actual shift lengths are not measured. Some OPD templates repeat standard provider/slot names; confirm that those listings reflect actual schedules before interpreting their percentages.

Missing or excluded cells are flagged. A wrong decryption key, failed download or unreadable workbook still stops the scan rather than producing an apparently complete report from partial files. Source-file counts and warnings are archive-wide; the report totals and monthly metrics are filtered to your selected period. Future scheduled sessions within the selected dates remain included.

## Where to edit later

- `schedule_app/services/learner_reach.py`: definition, clinical-shift deduplication, clinical metrics and work-type ambiguity handling.
- `schedule_app/services/teaching_analysis.py`: OPD parser, archive scan, actual-date filtering and enriched teaching rows.
- `schedule_app/reports/chair_summary.py`: chair report layout and text.
- `schedule_app/reports/individual_teaching.py`: individual report layout and text.
- `schedule_app/reports/teaching_export.py`: ZIP/CSV contents and notes.
- `schedule_app/sections/preceptor_teaching_summary.py`: page controls, previews and download buttons.

## Validation

90 local automated tests passed using invented data and simulated Streamlit/GitHub calls. The existing student-weighted overall, daily and work-type totals were also compared against the preceding custom-dates version using both provided OPD samples and were unchanged. Word previews were rendered and visually inspected. The complete package includes the tests. Live GitHub/Streamlit services were not accessed or changed.

For hidden-service hours and the distinction between overall and displayed work-type denominators, see UPDATE_TEACHING_CONTRIBUTORS.md.
