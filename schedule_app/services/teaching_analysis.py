"""Read-only archive analysis: student-shifts, academic years, deduplication and work types.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from collections import Counter
from collections import defaultdict
from datetime import date as CalendarDate
from datetime import datetime
from datetime import timedelta
from datetime import timezone as _teaching_timezone
from io import BytesIO
from openpyxl import load_workbook
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import _opd_date
from schedule_app.settings import TEACHING_EMPTY_STUDENT_LABELS
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT
from schedule_app.settings import TEACHING_MONTH_NAMES
from schedule_app.settings import TEACHING_NAME_ORDERS
from schedule_app.settings import TEACHING_OPD_NAME_ORDER_OVERRIDES
from schedule_app.settings import TEACHING_PRECEPTOR_NAME_MAP
from schedule_app.settings import TEACHING_REPORT_VERSION
from schedule_app.settings import TEACHING_WORK_TYPE_MAP
from schedule_app.settings import TEACHING_WORK_TYPE_ORDER
from schedule_app.settings import TEACHING_WORK_TYPE_REVIEW
from zoneinfo import ZoneInfo as _TeachingZoneInfo
import hashlib
import hmac
import json
import re
import secrets as _teaching_secrets


def teaching_site_key(site):
    return re.sub(r"[^A-Z0-9]+", "_", str(site or "").upper()).strip("_")


def teaching_work_type(site):
    """Classify the assignment by its OPD site, not the preceptor's name."""
    key = teaching_site_key(site)
    if not key:
        raise OPDArchiveError("An assigned shift has no OPD site; its work type cannot be determined.")
    if key in TEACHING_WORK_TYPE_MAP:
        label = str(TEACHING_WORK_TYPE_MAP[key]).strip()
        if not label or label == TEACHING_WORK_TYPE_REVIEW:
            raise OPDArchiveError("Check TEACHING_WORK_TYPE_MAP: use a nonempty work-type label, not the reserved review label.")
        return label
    # Preserve unmapped services, rather than silently assuming inpatient or outpatient.
    return key.replace("_", " ")


def teaching_work_type_sort(label):
    if label in TEACHING_WORK_TYPE_ORDER:
        return (0, TEACHING_WORK_TYPE_ORDER.index(label), "")
    if label == TEACHING_WORK_TYPE_REVIEW:
        return (2, 0, "")
    return (1, 0, str(label).casefold())


def teaching_require_work_type_data(scan):
    """Reject cached scans without site detail; totals alone cannot reconstruct it."""
    if scan.get("version") != TEACHING_REPORT_VERSION or "monthly_by_work_type" not in scan:
        raise OPDArchiveError("This teaching scan predates work-type reporting. Click Load / refresh archived OPDs, then generate the reports again.")
    expected, actual = Counter(), Counter()
    for row in scan["monthly"]:
        expected[(row["preceptor_name"], row["academic_start_year"], row["month"])] += row["no_of_shifts"]
    for row in scan["monthly_by_work_type"]:
        if not row.get("work_type") or int(row["no_of_shifts"]) <= 0:
            raise OPDArchiveError("The work-type breakdown is incomplete. Refresh the archived OPDs; no partial report was generated.")
        actual[(row["preceptor_name"], row["academic_start_year"], row["month"])] += row["no_of_shifts"]
    if dict(expected) != dict(actual):
        raise OPDArchiveError("Work-type subtotals do not match the overall teaching totals. Refresh the archived OPDs before generating a report.")


def teaching_work_type_rows(scan, selected_years):
    """One row per preceptor, academic year and work type; no student identifiers."""
    teaching_require_work_type_data(scan)
    years = {int(year) for year in selected_years}
    grouped = defaultdict(list)
    for item in scan["monthly_by_work_type"]:
        if item["academic_start_year"] in years:
            grouped[(item["preceptor_name"], item["academic_start_year"], item["work_type"])].append(item)
    rows = []
    for (name, year, work_type), items in sorted(
        grouped.items(), key=lambda pair: (teaching_name_key(pair[0][0]), pair[0][1], teaching_work_type_sort(pair[0][2]))
    ):
        months = sorted({item["month"] for item in items})
        total = sum(item["no_of_shifts"] for item in items)
        sites = sorted({site for item in items for site in item["source_sites"]})
        rows.append({
            "preceptor_name": name, "academic_year": teaching_academic_label(year),
            "work_type": work_type, "no_of_shifts": total,
            "months_worked": "; ".join(teaching_month_label(CalendarDate.fromisoformat(month)) for month in months),
            "educational_hours": total * TEACHING_HOURS_PER_STUDENT_SHIFT,
            "source_sites": "; ".join(sites),
        })
    return rows


def teaching_name_key(value):
    cleaned = re.sub(r"\s+", " ", str(value or "").strip())
    return re.sub(r"\s*,\s*", ", ", cleaned).casefold()


def teaching_display_name(value):
    return re.sub(r"\s*,\s*", ", ", re.sub(r"\s+", " ", str(value or "").strip()))


def teaching_local_today():
    # Streamlit Cloud may run on UTC. Keep the academic-year boundary local.
    try:
        return datetime.now(_TeachingZoneInfo("America/New_York")).date()
    except Exception:
        return datetime.now(_teaching_timezone.utc).date()


def teaching_academic_start(day):
    return day.year if day.month >= 7 else day.year - 1


def teaching_academic_label(start_year):
    return f"{start_year % 100:02d}-{(start_year + 1) % 100:02d}"


def teaching_month_label(month):
    return f"{TEACHING_MONTH_NAMES[month.month - 1]} {month.year}"


def teaching_label_needs_review(name, site_names=()):
    key = re.sub(r"[^A-Z0-9]+", "_", str(name).upper()).strip("_")
    site_keys = {re.sub(r"[^A-Z0-9]+", "_", str(s).upper()).strip("_") for s in site_names}
    return (
        key in site_keys or key in {"HAMPDEN_NURSERY", "PSHCH_NURSERY", "SJR_HOSPITALIST", "WARD_A"}
        or bool(re.fullmatch(r"(?:SJR|AAC|HOPE_DRIVE|NYES|ETOWN|LANCASTER|HAMPDEN_NURSERY|PSHCH_NURSERY)_?\d+", key))
        or key in {"TBD", "UNKNOWN", "PRECEPTOR", "PROVIDER", "TO_BE_ASSIGNED"}
        or bool(re.search(r"[;&|]|\s/\s|\s+and\s+", str(name), re.I))
    )


def teaching_split_assignment(value, order):
    """Return (provider, [students]) groups without splitting 'Last, First'.

    Names are used only temporarily to count separate students and recognize
    exact duplicates. They never enter the exported CSV, Word files or notes.
    Multiple students: separate OPD rows, or ; / newline / | / ' & ' / ' and '
    on the student side. Repeated full Provider ~ Student pairs may be separated
    with a newline, semicolon or |. A bare repeated ~ is ambiguous and rejected.
    """
    if order not in TEACHING_NAME_ORDERS:
        raise OPDArchiveError("Choose Preceptor ~ Student or Student ~ Preceptor for the teaching summary.")
    if not isinstance(value, str) or "~" not in value:
        return []
    text = value.strip()
    if text.count("~") == 1:
        segments = [text]
    else:
        segments = [part.strip() for part in re.split(r"[\r\n;|]+", text) if part.strip()]
        if any(part.count("~") != 1 for part in segments):
            raise OPDArchiveError("An assignment contains an ambiguous repeated '~'. Use separate OPD rows "
                                  "or a semicolon-separated list of students on the student side.")
    parsed = []
    for segment in segments:
        left, right = (part.strip() for part in segment.split("~", 1))
        provider, student_text = (left, right) if order == "Preceptor ~ Student" else (right, left)
        students = [teaching_display_name(part) for part in
                    re.split(r"[\r\n;|]+|\s+&\s+|\s+and\s+", student_text, flags=re.I)]
        students = [name for name in students if teaching_name_key(name) not in TEACHING_EMPTY_STUDENT_LABELS]
        parsed.append((teaching_display_name(provider), students))
    return parsed


def teaching_extract_assignments(raw, details, order):
    """Read AM/PM cells under each *actual* date, including hidden OPD rows.

    Returns temporary records containing student names. The caller aggregates
    and discards these records before storing anything in Streamlit session_state.
    """
    wb = load_workbook(BytesIO(raw), read_only=True, data_only=False)
    records, missing_provider_cells = [], []
    expected_days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    try:
        for sheet_name in details["site_names"]:
            ws = wb[sheet_name]
            if (ws.max_row or 0) > 1200:
                raise OPDArchiveError(f"Worksheet '{sheet_name}' exceeds the supported 1,200-row OPD layout; "
                                      "no partial teaching report was generated.")
            rows = list(ws.iter_rows(max_col=8))
            current_dates = None
            for row_index, row in enumerate(rows):
                values = [cell.value for cell in row]
                if [str(v or "").strip().casefold() for v in values[1:8]] == expected_days:
                    if row_index + 1 >= len(rows):
                        raise OPDArchiveError(f"Missing dates on worksheet '{sheet_name}'.")
                    current_dates = [_opd_date(cell.value, wb.epoch) for cell in rows[row_index + 1][1:8]]
                    if not all(current_dates) or any(
                        day != current_dates[0] + timedelta(days=i) for i, day in enumerate(current_dates)
                    ):
                        raise OPDArchiveError(f"Invalid date row on worksheet '{sheet_name}'.")
                    continue
                match = re.match(r"^\s*(AM|PM)\b", str(values[0] or ""), re.I)
                if not match:
                    continue
                for col_index, cell in enumerate(row[1:8]):
                    if cell.data_type == "f":
                        raise OPDArchiveError(f"{sheet_name}!{cell.coordinate} contains a formula in a session cell. "
                                              "Use assignment values rather than formulas for this summary.")
                    try:
                        groups = teaching_split_assignment(cell.value, order)
                    except OPDArchiveError as exc:
                        raise OPDArchiveError(f"{sheet_name}!{cell.coordinate}: {exc}") from None
                    for provider, students in groups:
                        if not students:
                            continue  # provider availability alone earns no teaching hours
                        if teaching_name_key(provider) in TEACHING_EMPTY_STUDENT_LABELS:
                            missing_provider_cells.append(f"{sheet_name}!{cell.coordinate}")
                            continue
                        if current_dates is None:
                            raise OPDArchiveError(f"{sheet_name}!{cell.coordinate} has an assignment before a date header.")
                        for student in students:
                            records.append({
                                "preceptor_name": provider, "student": student,
                                "day": current_dates[col_index], "shift": match.group(1).upper(),
                                "site": sheet_name, "cell": cell.coordinate,
                            })
        return records, missing_provider_cells
    finally:
        wb.close()


def teaching_scan_archives(client, default_order=TEACHING_NAME_ORDERS[0], progress=None):
    """Read/decrypt each current OPD at one repository snapshot, retaining work type.

    Student names are temporary. Run-local keyed digests deduplicate assignments;
    neither the digests nor learner names appear in returned aggregates or reports.
    Work-type subtotals always reconcile to the existing overall counting rules.
    """
    if default_order not in TEACHING_NAME_ORDERS:
        raise OPDArchiveError("Select a supported OPD name order.")
    commit = client._head()
    rotations = client.list_rotations(commit=commit)
    counts = Counter()
    names, name_variants, review_labels, manifests, warnings = {}, defaultdict(set), set(), [], []
    # One anonymous item per provider/student/date/AM-or-PM assignment.
    seen_assignments = {}
    salt = _teaching_secrets.token_bytes(32)
    aliases = {teaching_name_key(k): teaching_display_name(v)
               for k, v in TEACHING_PRECEPTOR_NAME_MAP.items() if str(k).strip() and str(v).strip()}
    duplicates = future_assignments = 0
    today = teaching_local_today()
    observed_site_groups = {}
    for number, rotation in enumerate(rotations, start=1):
        order = TEACHING_OPD_NAME_ORDER_OVERRIDES.get(rotation.isoformat(), default_order)
        if order not in TEACHING_NAME_ORDERS:
            raise OPDArchiveError(f"Invalid teaching-summary name-order override for rotation {rotation.isoformat()}.")
        try:
            loaded = client.load(rotation, commit=commit)
            records, missing_cells = teaching_extract_assignments(loaded["raw"], loaded["details"], order)
        except OPDArchiveError as exc:
            raise OPDArchiveError(f"Rotation {rotation.isoformat()}: {exc} No ZIP was generated from partial data.") from None
        counted = removed = 0
        for item in records:
            original_name = item["preceptor_name"]
            name = aliases.get(teaching_name_key(original_name), original_name)
            key = teaching_name_key(name)
            work_type = teaching_work_type(item["site"])
            site = teaching_site_key(item["site"])
            observed_site_groups[site] = work_type
            identity = json.dumps([key, teaching_name_key(item["student"]),
                                   item["day"].isoformat(), item["shift"]], ensure_ascii=False)
            digest = hmac.new(salt, identity.encode("utf-8"), hashlib.sha256).digest()
            # A duplicate within Academic Pediatrics is still counted once even
            # if it is copied onto both HOPE_DRIVE and NYES worksheets. Distinct
            # students in the same shift have different identities and count twice.
            if digest in seen_assignments:
                duplicates += 1
                removed += 1
                seen_assignments[digest]["work_types"].add(work_type)
                seen_assignments[digest]["sites"].add(site)
                continue
            month = item["day"].replace(day=1)
            seen_assignments[digest] = {
                "provider_key": key, "month": month, "day": item["day"], "shift": item["shift"],
                "work_types": {work_type}, "sites": {site},
            }
            names.setdefault(key, name)
            name_variants[key].add(original_name)
            counts[(key, month)] += 1
            counted += 1
            future_assignments += int(item["day"] > today)
            if teaching_label_needs_review(name, loaded["details"]["site_names"]):
                review_labels.add(key)
        if missing_cells:
            warnings.append({"rotation_start": rotation.isoformat(), "issue": "Missing provider; not attributed",
                             "details": ", ".join(sorted(set(missing_cells)))})
        manifests.append({
            "rotation_start": rotation.isoformat(),
            "last_scheduled_date": (loaded["details"]["week_mondays"][-1] + timedelta(days=6)).isoformat(),
            "archive_file": loaded["path"].rsplit("/", 1)[-1],
            "github_blob_sha": loaded["sha"], "name_order": order,
            "assigned_student_shifts_read": len(records),
            "assigned_student_shifts_counted": counted,
            "duplicate_student_shifts_removed": removed,
            "missing_provider_cells": len(set(missing_cells)),
        })
        del records, loaded
        if progress:
            progress(number, len(rotations))
    monthly = [{"preceptor_name": names[key], "academic_year": teaching_academic_label(teaching_academic_start(month)),
                "academic_start_year": teaching_academic_start(month), "month": month.isoformat(),
                "no_of_shifts": int(count), "educational_hours": int(count * TEACHING_HOURS_PER_STUDENT_SHIFT)}
               for (key, month), count in sorted(counts.items(), key=lambda item: (item[0][0], item[0][1]))]

    type_counts, conflict_counts = Counter(), Counter()
    type_sites = defaultdict(set)
    for item in seen_assignments.values():
        key, month = item["provider_key"], item["month"]
        work_type = next(iter(item["work_types"])) if len(item["work_types"]) == 1 else TEACHING_WORK_TYPE_REVIEW
        type_counts[(key, month, work_type)] += 1
        type_sites[(key, month, work_type)].update(item["sites"])
        if len(item["work_types"]) > 1:
            # Do not silently credit the same assignment to the first site read.
            # It stays in the overall total once, in an explicit review category.
            conflict_counts[(key, item["day"], item["shift"],
                             "; ".join(sorted(item["work_types"], key=teaching_work_type_sort)),
                             "; ".join(sorted(item["sites"])))] += 1
    monthly_by_type = [{
        "preceptor_name": names[key], "academic_year": teaching_academic_label(teaching_academic_start(month)),
        "academic_start_year": teaching_academic_start(month), "month": month.isoformat(),
        "work_type": work_type, "no_of_shifts": int(count),
        "educational_hours": int(count * TEACHING_HOURS_PER_STUDENT_SHIFT),
        "source_sites": sorted(type_sites[(key, month, work_type)]),
    } for (key, month, work_type), count in sorted(
        type_counts.items(), key=lambda pair: (pair[0][0], pair[0][1], teaching_work_type_sort(pair[0][2]))) ]
    conflicts = [{
        "preceptor_name": names[key], "date": day.isoformat(), "shift": shift,
        "academic_year": teaching_academic_label(teaching_academic_start(day)),
        "conflicting_work_types": types, "source_sites": sites,
        "no_of_student_shifts": count,
    } for (key, day, shift, types, sites), count in sorted(conflict_counts.items())]
    result = {
        "version": TEACHING_REPORT_VERSION, "commit": commit,
        "generated_at": datetime.now(_teaching_timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "default_name_order": default_order, "monthly": monthly,
        "monthly_by_work_type": monthly_by_type, "work_type_conflicts": conflicts,
        "site_work_type_mapping": dict(sorted(observed_site_groups.items())),
        "sources": manifests, "warnings": warnings, "duplicate_assignments_removed": duplicates,
        "future_assignments_in_archive": future_assignments,
        "unresolved_preceptor_labels": sorted((names[k] for k in review_labels), key=teaching_name_key),
        "name_variants": {names[k]: sorted(v) for k, v in name_variants.items() if len(v) > 1},
    }
    teaching_require_work_type_data(result)
    return result


def teaching_annual_rows(scan, selected_years):
    selected = {int(year) for year in selected_years}
    grouped = defaultdict(list)
    for item in scan["monthly"]:
        if item["academic_start_year"] in selected:
            grouped[(item["preceptor_name"], item["academic_start_year"])].append(item)
    rows = []
    for (name, start_year), items in sorted(grouped.items(), key=lambda item: (teaching_name_key(item[0][0]), item[0][1])):
        ordered = sorted(items, key=lambda item: item["month"])
        total = sum(item["no_of_shifts"] for item in ordered)
        rows.append({"preceptor_name": name, "academic_year": teaching_academic_label(start_year),
                     "no_of_shifts": total,
                     "months_worked": "; ".join(teaching_month_label(CalendarDate.fromisoformat(item["month"])) for item in ordered),
                     "educational_hours": total * TEACHING_HOURS_PER_STUDENT_SHIFT})
    return rows


def teaching_brief_months(month_values):
    """Compact month labels without implying teaching occurred in missing months."""
    months = sorted({CalendarDate.fromisoformat(value).replace(day=1) for value in month_values})
    if not months:
        return "Not recorded"
    groups, first, last = [], months[0], months[0]
    for month in months[1:]:
        if month.year == last.year and month.month == last.month + 1:
            last = month
        else:
            groups.append((first, last))
            first = last = month
    groups.append((first, last))
    labels = []
    for first, last in groups:
        label = TEACHING_MONTH_NAMES[first.month - 1][:3]
        if first != last:
            label += "-" + TEACHING_MONTH_NAMES[last.month - 1][:3]
        labels.append(f"{label} {first.year}")
    return "; ".join(labels)
