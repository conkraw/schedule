"""One Learner Reach pie per clinical experience, never per individual preceptor.

Uses the same area totals as the chair table. Images stay in memory until added
to the staff report ZIP; this module never writes reports to GitHub.
"""
from io import BytesIO
import hashlib
import re
from textwrap import fill
from threading import RLock

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.learner_reach import reach_percent, reach_totals
from schedule_app.services.reporting_periods import teaching_report_label, teaching_report_date_text

_CHART_LOCK = RLock()
CHART_DATA_COLUMNS = (
    "academic_year", "work_type", "recorded_clinical_hours", "hours_with_students",
    "hours_without_students", "learner_reach_pct", "source_sites",
)


def learner_reach_pie(work_type, metrics, period_label="", date_text=""):
    """Return a PNG with two slices: clinical hours with and without students.

    Object-oriented Figure + Agg avoids global pyplot figures in Streamlit.
    The lock also keeps rendering separate across simultaneous app sessions.
    """
    checked = reach_totals([metrics])
    percent = reach_percent(checked["learner_reach_pct"])
    for field in ("recorded_clinical_hours", "hours_with_students", "hours_without_students"):
        if metrics.get(field) != checked[field]:
            raise OPDArchiveError("Clinical chart hours do not match the report counts. Reports are blocked.")
    if metrics.get("learner_reach_pct") != checked["learner_reach_pct"]:
        raise OPDArchiveError("The clinical chart percentage does not match the table. Reports are blocked.")
    try:
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_agg import FigureCanvasAgg
    except ImportError:
        raise OPDArchiveError("The chart dependency is missing. Add matplotlib>=3.9,<4 to requirements.txt "
                              "and restart the app before generating teaching reports.") from None
    with_students, without_students = checked["hours_with_students"], checked["hours_without_students"]
    total = checked["recorded_clinical_hours"]
    with _CHART_LOCK:
        figure = Figure(figsize=(7.0, 2.65), dpi=180)
        FigureCanvasAgg(figure)
        try:
            # A single independent pie, not a multi-area subplot or a percentage
            # of all teaching effort. Keep the same default color order in every chart.
            ax = figure.add_axes((0.02, 0.06, 0.35, 0.87))
            wedges, _ = ax.pie([with_students, without_students], startangle=90,
                                counterclock=False, normalize=True,
                                wedgeprops={"linewidth": 0.8})
            ax.set_aspect("equal")
            figure.text(0.40, 0.89, fill(str(work_type), width=35),
                        fontsize=12, fontweight="bold", va="top")
            figure.text(0.40, 0.59, f"{percent} Learner Reach", fontsize=16, fontweight="bold")
            figure.legend(wedges,
                [f"With students: {with_students:,} h ({percent})",
                 f"Without students: {without_students:,} h ({100 * without_students / total:.1f}%)"],
                loc="center left", bbox_to_anchor=(0.39, 0.40), fontsize=10,
                frameon=False, handlelength=1.1)
            figure.text(0.40, 0.18, f"Total recorded OPD hours: {total:,}", fontsize=10)
            # The label and date range are also retained in image metadata/alt text.
            output = BytesIO()
            figure.savefig(output, format="png", dpi=180,
                           metadata={"Title": f"{work_type} - Learner Reach",
                                     "Description": f"{period_label}; {date_text}; "
                                                    f"{with_students} of {total} recorded hours with students ({percent})."})
            return output.getvalue()
        finally:
            figure.clear()


def teaching_clinical_charts(scan, summaries):
    """List of chart bytes, filenames and table-backed alt text for included areas."""
    result = []
    for item in summaries:
        year = item["academic_start_year"]
        period_label = teaching_report_label(scan, year)
        date_text = teaching_report_date_text(scan, year)
        for group in item["work_types"]:
            if int(group["no_of_shifts"]) <= 0:
                continue
            work_type = group["work_type"]
            safe = re.sub(r"[^A-Za-z0-9._-]+", "_", work_type).strip("._")[:75] or "Experience"
            suffix = hashlib.sha256(work_type.encode()).hexdigest()[:8]
            filename = f"Learner_Reach_Charts/{year}_{safe}_{suffix}.png"
            png = learner_reach_pie(work_type, group, period_label, date_text)
            description = (
                f"{work_type}. Reporting period {period_label}, {date_text}. "
                f"Learner Reach {reach_percent(group['learner_reach_pct'])}: "
                f"{group['hours_with_students']:,} of {group['recorded_clinical_hours']:,} "
                f"recorded OPD hours included students; {group['hours_without_students']:,} hours did not. "
                "Only preceptors with student assignments in this experience are represented."
            )
            result.append({"key": (year, work_type), "filename": filename, "png": png,
                           "alt_text": description, "caption": f"{work_type} | {period_label} | {date_text}",
                           "data": {"academic_year": period_label, "work_type": work_type,
                                    **{key: group[key] for key in ("recorded_clinical_hours", "hours_with_students",
                                                                 "hours_without_students", "learner_reach_pct")},
                                    "source_sites": "; ".join(group["source_sites"])}})
    return result
