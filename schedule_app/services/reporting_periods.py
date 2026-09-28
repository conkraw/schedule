"""User-chosen teaching-report dates. No archive writes or academic-year assumptions.

One custom period is one report group, even when it crosses July or spans more
than twelve months. Legacy integer-year reports retain their original behavior.
"""
from dataclasses import dataclass
from datetime import date, datetime
import json
import re

from schedule_app.services.opd_archive import OPDArchiveError

DATE_RANGE_SCHEMA_VERSION = 1


@dataclass(frozen=True)
class ReportingPeriod:
    label: str
    start: date
    end: date

    def __post_init__(self):
        label = self.label.strip() if isinstance(self.label, str) else ""
        if not label or len(label) > 60 or any(ord(c) < 32 or ord(c) == 127 or c in "\ufffe\uffff" or 0xD800 <= ord(c) <= 0xDFFF for c in label):
            raise OPDArchiveError("Enter a report label of 1-60 characters, such as 26-27.")
        if any(not isinstance(d, date) or isinstance(d, datetime) for d in (self.start, self.end)):
            raise OPDArchiveError("Select both a start date and an end date.")
        if self.start > self.end:
            raise OPDArchiveError("The end date must be on or after the start date.")
        if self.start < date(1970, 1, 1) or self.end > date(2100, 12, 31):
            raise OPDArchiveError("Choose reporting dates between 1970 and 2100.")
        object.__setattr__(self, "label", label)

    def as_dict(self):
        return {"label": self.label, "start_date": self.start.isoformat(), "end_date": self.end.isoformat()}

    def signature(self):
        return (self.label, self.start.isoformat(), self.end.isoformat())

    def filename_part(self):
        label = re.sub(r"[^A-Za-z0-9._-]+", "_", self.label).strip("._")[:60] or "Custom"
        return f"{label}_{self.start.isoformat()}_to_{self.end.isoformat()}"


def period_from_dict(values):
    try:
        return ReportingPeriod(values["label"], date.fromisoformat(values["start_date"]),
                               date.fromisoformat(values["end_date"]))
    except (KeyError, TypeError, ValueError):
        raise OPDArchiveError("The saved reporting dates are invalid. Select a reporting-date JSON file created by this app.") from None


def reporting_period_json(period):
    return (json.dumps({"reporting_period_version": 1, **period.as_dict()}, indent=2) + "\n").encode("utf-8")


def read_reporting_period_json(raw):
    if not isinstance(raw, bytes) or len(raw) > 8192:
        raise OPDArchiveError("The reporting-date settings file must be a JSON file smaller than 8 KB.")
    try:
        data = json.loads(raw.decode("utf-8-sig"))
    except (ValueError, UnicodeError, RecursionError):
        raise OPDArchiveError("The reporting-date settings file is not valid JSON.") from None
    if not isinstance(data, dict) or data.get("reporting_period_version") != 1:
        raise OPDArchiveError("Select a reporting-date JSON file created by this app.")
    return period_from_dict(data)


def teaching_period(scan):
    values = scan.get("reporting_period")
    return period_from_dict(values) if values is not None else None


def teaching_report_label(scan, group_year):
    period = teaching_period(scan)
    return period.label if period else f"{group_year % 100:02d}-{(group_year + 1) % 100:02d}"


def teaching_report_bounds(scan, group_year):
    period = teaching_period(scan)
    return (period.start, period.end) if period else (date(group_year, 7, 1), date(group_year + 1, 6, 30))


def teaching_report_heading(scan, group_year):
    prefix = "Reporting period" if scan.get("reporting_period") else "Academic year"
    return f"{prefix} {teaching_report_label(scan, group_year)}"


def teaching_report_date_text(scan, group_year):
    start, end = teaching_report_bounds(scan, group_year)
    return f"{start:%B} {start.day}, {start.year} - {end:%B} {end.day}, {end.year}"
