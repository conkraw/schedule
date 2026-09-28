"""Editable teaching/report settings and preceptor email mappings.

Extracted from the supplied app; this module performs no page rendering on import.
"""


# Edit the mappings in this file rather than searching through the app.
# Credentials stay in Streamlit Secrets, never here.
# After changing Python settings in a running deployment, restart/reboot the app.
# PRECEPTOR_EMAIL_MAP: names -> email for Power Automate reports.
# TEACHING_PRECEPTOR_NAME_MAP: optional explicit name aliases, never guessed identities.
# TEACHING_OPD_NAME_ORDER_OVERRIDES: rotation Monday -> name-order exception.
# TEACHING_WORK_TYPE_MAP: OPD site names -> work types in both Word reports.
# Keep changing settings separate from changing the assignment/counting algorithms.

TEACHING_REPORT_VERSION = 3


TEACHING_HOURS_PER_STUDENT_SHIFT = 4


TEACHING_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "no_of_shifts",
    "months_worked", "educational_hours",
)


TEACHING_MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


TEACHING_NAME_ORDERS = ("Preceptor ~ Student", "Student ~ Preceptor")


TEACHING_PRECEPTOR_NAME_MAP = {
    # "Smith, J.": "Smith, Jane",
}


TEACHING_OPD_NAME_ORDER_OVERRIDES = {
    # "2026-08-03": "Preceptor ~ Student",
}


TEACHING_EMPTY_STUDENT_LABELS = {
    "", "nan", "none", "n/a", "na", "tbd", "unassigned", "no student",
    "no students", "open", "available", "off", "vacation", "holiday",
}


TEACHING_WORK_TYPE_MAP = {
    "HOPE_DRIVE": "Academic Pediatrics",
    "ETOWN": "Academic Pediatrics",
    "NYES": "Academic Pediatrics",
    "WARD_A": "Ward A",
    "PSHCH_NURSERY": "PSHCH Nursery",
    "COMPLEX": "Complex Care",
}


TEACHING_WORK_TYPE_ORDER = (
    "Academic Pediatrics", "Ward A", "PSHCH Nursery", "Complex Care",
)


TEACHING_WORK_TYPE_REVIEW = "Work type needs review"


TEACHING_WORK_TYPE_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "work_type", "no_of_shifts",
    "months_worked", "educational_hours", "source_sites",
)


TEACHING_CHAIR_SUMMARY_FILENAME = "Pediatric_Clerkship_Educational_Effort_Summary.docx"


PRECEPTOR_EMAIL_MAP = {
    # "Preceptor Name": "preceptor_email@pennstatehealth.psu.edu",
}


FOCUS_SITES = {"HOPE_DRIVE", "NYES", "ETOWN"}


REPORT_COLUMNS = [
    "preceptor_name",
    "student_name",
    "no_of_sessions",
    "monday_date",
    "primary_preceptor",
    "fragmented_preceptor",
    "primary_preceptor_flag",
    "primary_preceptor_flag_reason",
    "email",
]


INDIVIDUAL_REPORT_SCHEMA_VERSION = 3


INDIVIDUAL_OUTPUT_STATE_KEYS = (
    "individual_schedule_zip",
    "individual_preceptor_report",
    "individual_preceptor_preview",
    "individual_missing_emails",
)
