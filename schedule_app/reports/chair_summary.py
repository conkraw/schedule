"""Combined chair-facing Word summary, by academic year and type of work.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.teaching_priority import outpatient_priority_report_note
from schedule_app.services.teaching_validation import validate_teaching_report
from schedule_app.services.report_diagnostics import checked_report_reach
from schedule_app.services.student_continuity import (
    require_student_continuity_data, student_continuity_counts,
)
from schedule_app.reports.teaching_tables import teaching_add_work_table
from collections import defaultdict
from datetime import date as CalendarDate
from docx import Document
from docx.shared import Pt
from io import BytesIO
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import (
    teaching_report_bounds, teaching_report_label, teaching_report_heading, teaching_report_date_text,
)
from schedule_app.services.teaching_analysis import teaching_annual_rows
from schedule_app.services.teaching_analysis import teaching_brief_months
from schedule_app.services.teaching_analysis import teaching_name_key
from schedule_app.services.teaching_analysis import teaching_work_type_rows
from schedule_app.services.teaching_analysis import teaching_work_type_sort
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT
from schedule_app.settings import TEACHING_WORK_TYPE_REVIEW
from schedule_app.services.learner_reach import (
    reach_totals, reach_percent, REACH_DEFINITION, REACH_SCOPE_NOTE,
    require_learner_reach_data, PARTICIPATION_SCOPE_NOTE, REACH_DETAIL_TOTAL_NOTE,
)


# Presentation-only version: invalidate old report downloads without discarding
# an otherwise current OPD scan or the user's selected GitHub date preset.
CHAIR_STUDENT_CONTINUITY_REPORT_VERSION = 2


def teaching_chair_summary_data(scan, selected_years):
    """Year-by-year summaries using the exact rows used for the existing CSV.

    Generic site/slot/combined labels remain in overall assignment totals, but
    are separated from named preceptors so they are not presented as people.
    """
    years = sorted({int(year) for year in selected_years})
    validate_teaching_report(scan, years)
    require_student_continuity_data(scan)
    annual = teaching_annual_rows(scan, years)
    review_keys = {teaching_name_key(name) for name in scan.get("unresolved_preceptor_labels", [])}
    typed = teaching_work_type_rows(scan, years)
    summaries = []
    for year in years:
        label = teaching_report_label(scan, year)
        monthly = [row for row in scan["monthly"] if row["academic_start_year"] == year]
        by_name = defaultdict(list)
        for row in monthly:
            by_name[row["preceptor_name"]].append(row["month"])
        named, unresolved = [], []
        for row in annual:
            if row["academic_year"] != label:
                continue
            entry = checked_report_reach(row, report="Chair summary",
                                         section="Overall preceptor totals", academic_year=label)
            # Overall preceptor counts, identical to the individual Word report.
            # Do not sum monthly or work-type unique counts: one student may
            # appear in several months and settings with the same preceptor.
            entry.update(student_continuity_counts(scan, row["preceptor_name"], year))
            entry["months_brief"] = teaching_brief_months(by_name[row["preceptor_name"]])
            target = unresolved if teaching_name_key(row["preceptor_name"]) in review_keys else named
            target.append(entry)
        if not named and not unresolved:
            continue  # Do not add empty academic-year sections to a combined report.
        work_types = []
        year_typed = [row for row in typed if row["academic_year"] == label]
        year_type_months = [row for row in scan["monthly_by_work_type"] if row["academic_start_year"] == year]
        for work_type in sorted({row["work_type"] for row in year_typed}, key=teaching_work_type_sort):
            entries = [checked_report_reach(row, report="Chair summary",
                        section="Preceptor detail", academic_year=label, work_type=work_type)
                       for row in year_typed if row["work_type"] == work_type]
            type_months = [row for row in year_type_months if row["work_type"] == work_type]
            for entry in entries:
                entry["months_brief"] = teaching_brief_months(row["month"] for row in type_months
                                                             if row["preceptor_name"] == entry["preceptor_name"])
            sites = sorted({site for row in entries for site in row.get("source_sites", "").split("; ") if site})
            work_types.append({
                "work_type": work_type,
                "named_preceptors": [row for row in entries if teaching_name_key(row["preceptor_name"]) not in review_keys],
                "unresolved_labels": [row for row in entries if teaching_name_key(row["preceptor_name"]) in review_keys],
                "no_of_shifts": sum(row["no_of_shifts"] for row in entries),
                "educational_hours": sum(row["educational_hours"] for row in entries),
                "months_brief": teaching_brief_months(row["month"] for row in type_months),
                "source_sites": sites,
                **reach_totals(entries, report="Chair summary", section="Clinical experience total",
                               academic_year=label, work_type=work_type),
            })
        start, end = teaching_report_bounds(scan, year)
        relevant_sources = [source for source in scan.get("sources", [])
                            if CalendarDate.fromisoformat(source["rotation_start"]) <= end
                            and CalendarDate.fromisoformat(source["last_scheduled_date"]) >= start]
        rows = named + unresolved
        summaries.append({
            "academic_start_year": year,
            "academic_year": label,
            "work_types": work_types,
            "named_preceptors": sorted(named, key=lambda row: teaching_name_key(row["preceptor_name"])),
            "unresolved_labels": sorted(unresolved, key=lambda row: teaching_name_key(row["preceptor_name"])),
            "named_preceptor_count": len(named),
            "named_preceptors_with_students": sum(row["no_of_shifts"] > 0 for row in named),
            **reach_totals(rows, report="Chair summary", section="Overall total", academic_year=label),
            "has_unlisted_clinical_hours": (
                sum(row["recorded_clinical_hours"] for row in rows)
                != sum(group["recorded_clinical_hours"] for group in work_types)
            ),
            "no_of_shifts": sum(row["no_of_shifts"] for row in rows),
            "educational_hours": sum(row["educational_hours"] for row in rows),
            "months_brief": teaching_brief_months(row["month"] for row in monthly),
            "source_count": len(relevant_sources),
            # Cells may contain more than one learner; this is not a shift count.
            "has_missing_provider": any(int(source.get("missing_provider_cells", 0)) > 0
                                        for source in relevant_sources),
        })
    return summaries


def teaching_make_chair_summary(scan, selected_years, *, charts=None):
    """One editable, chair-friendly Word report covering all selected years.

    Educational hours still use scheduled student-shifts, not elapsed hours or
    verified attendance. A separate all-work-types table gives each preceptor's
    unique students and students assigned on three or more distinct dates.
    """
    from docx.shared import Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    require_learner_reach_data(scan)
    summaries = teaching_chair_summary_data(scan, selected_years)
    if not summaries or not any(item["no_of_shifts"] for item in summaries):
        raise OPDArchiveError("No student assignments were found for the selected reporting period(s); no empty chair summary was generated.")

    from schedule_app.reports.learner_reach_charts import teaching_clinical_charts
    chart_map = {chart["key"]: chart for chart in (charts if charts is not None else teaching_clinical_charts(scan, summaries))}
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin = section.bottom_margin = Inches(0.7)
    section.left_margin = section.right_margin = Inches(0.8)
    section.header_distance = section.footer_distance = Inches(0.3)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.05
    for name, size in (("Title", 23), ("Heading 1", 16), ("Heading 2", 12)):
        style = doc.styles[name]
        style.font.name, style.font.size = "Calibri", Pt(size)
        style.font.color.rgb = RGBColor.from_string("24466B")
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8 if name == "Heading 2" else 0)
        style.paragraph_format.space_after = Pt(4)
        properties = style.element.find(qn("w:pPr"))
        if properties is not None:
            borders = properties.find(qn("w:pBdr"))
            if borders is not None:
                properties.remove(borders)
    subtitle = doc.styles["Subtitle"]
    subtitle.font.name, subtitle.font.size = "Calibri", Pt(12)
    subtitle.font.italic = False
    subtitle.font.color.rgb = RGBColor.from_string("526475")
    subtitle.paragraph_format.space_before = Pt(0)
    subtitle.paragraph_format.space_after = Pt(12)
    subtitle.paragraph_format.keep_with_next = True
    header = section.header.paragraphs[0]
    header.text = "PENN STATE  |  PEDIATRIC CLERKSHIP"
    header.runs[0].font.size = Pt(9)
    header.runs[0].font.color.rgb = RGBColor.from_string("526475")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("Educational effort summary  |  Page ").font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.core_properties.title = "Pediatric clerkship educational effort summary"
    doc.core_properties.author = "Pediatric Clerkship"
    doc.core_properties.subject = "Scheduled third-year student teaching by reporting period and type of work"

    def note(text, *, warning=False):
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string("8C3B25" if warning else "526475")
        return p

    def add_effort_table(entries, *, pending=False, first_title=None, total_title=None, total_metrics=None,
                         work_type="All work types", academic_year=""):
        entries = list(entries)
        if not entries:
            return None  # No table or percentage for an absent subgroup.
        context = {"report": "Chair summary", "section": first_title or "Preceptor detail",
                   "work_type": work_type, "academic_year": academic_year}
        entries = [checked_report_reach(row, **context) for row in entries]
        metrics = checked_report_reach(
            total_metrics if total_metrics is not None else reach_totals(entries, **context),
            **{**context, "section": total_title or "Preceptor subtotal"})
        table = doc.add_table(rows=1, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        # 6.9 inches = the printable width; also set the table grid, not only cells.
        widths = tuple(Inches(value) for value in (2.35, 0.9, 0.95, 0.95, 0.8, 0.95))
        for col, width in zip(table.columns, widths):
            col.width = width
        titles = (first_title or ("Provider label" if pending else "Preceptor / teaching months"),
                  "OPD hours", "Hours with students", "Learner Reach", "Student-shifts", "Education hours*")
        repeat_header = OxmlElement("w:tblHeader")
        table.rows[0]._tr.get_or_add_trPr().append(repeat_header)

        def fill_row(row, values, *, heading=False, total=False, band=False):
            row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            for index, (cell, value, width) in enumerate(zip(row.cells, values, widths)):
                cell.width = width
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.text = str(value)
                properties = cell._tc.get_or_add_tcPr()
                margin = OxmlElement("w:tcMar")
                for edge, amount in (("top", 40), ("bottom", 40), ("left", 90), ("right", 90)):
                    item = OxmlElement("w:" + edge)
                    item.set(qn("w:w"), str(amount))
                    item.set(qn("w:type"), "dxa")
                    margin.append(item)
                properties.append(margin)
                if heading or total or band:
                    shade = OxmlElement("w:shd")
                    shade.set(qn("w:fill"), "24466B" if heading else "E8EEF5" if total else "F5F7FA")
                    properties.append(shade)
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.keep_with_next = heading
                if index >= 1:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for run in p.runs:
                    run.font.name, run.font.size = "Calibri", Pt(9.5)
                    run.bold = heading or total
                    if heading:
                        run.font.color.rgb = RGBColor(255, 255, 255)

        fill_row(table.rows[0], titles, heading=True)
        for index, row in enumerate(entries):
            months = row.get("months_brief", "Not recorded")
            label = row["preceptor_name"] + "\n" + (months if row["no_of_shifts"] else "No student assignments")
            fill_row(table.add_row(), (label, f"{row['recorded_clinical_hours']:,}",
                                      f"{row['hours_with_students']:,}", reach_percent(row["learner_reach_pct"]),
                                      f"{row['no_of_shifts']:,}", f"{row['educational_hours']:,}"),
                     band=index % 2 == 1)
        subtotal = total_title or ("Awaiting attribution" if pending else "Named preceptors total")
        fill_row(table.add_row(), (subtotal, f"{metrics['recorded_clinical_hours']:,}",
                                   f"{metrics['hours_with_students']:,}", reach_percent(metrics["learner_reach_pct"]),
                                   f"{sum(row['no_of_shifts'] for row in entries):,}",
                                   f"{sum(row['educational_hours'] for row in entries):,}"), total=True)
        # Avoid leaving the subtotal by itself at the top of a new page.
        if len(table.rows) > 2:
            for cell in table.rows[-2].cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
        return table

    for index, item in enumerate(summaries):
        year = item["academic_start_year"]
        item = checked_report_reach(item, report="Chair summary", section="Overall total")
        if index:
            doc.add_page_break()
        doc.add_paragraph("Preceptor educational effort", style="Title")
        doc.add_paragraph("Third-year medical student teaching", style="Subtitle")
        doc.add_heading(teaching_report_heading(scan, year), level=1)
        p = doc.add_paragraph(teaching_report_date_text(scan, year))
        p.paragraph_format.space_after = Pt(9)
        if not item["recorded_clinical_shifts"]:
            doc.add_paragraph("No clinical shifts were recorded in the archived OPDs for this reporting period. "
                              "This does not establish that no clinical work or teaching occurred.")
            continue

        p = doc.add_paragraph()
        p.add_run("Archived OPD schedules record ")
        p.add_run(f"{item['no_of_shifts']:,} " + ("student-shift" if item["no_of_shifts"] == 1 else "student-shifts")).bold = True
        p.add_run(", representing ")
        p.add_run(f"{item['educational_hours']:,} educational hours").bold = True
        p.add_run(" of scheduled teaching.")
        p = doc.add_paragraph()
        p.add_run("Named preceptors with students: ").bold = True
        p.add_run(str(item["named_preceptor_count"]))
        p.add_run("   |   Months with assignments: ").bold = True
        p.add_run(item["months_brief"])
        if item["unresolved_labels"]:
            pending_shifts = sum(row["no_of_shifts"] for row in item["unresolved_labels"])
            pending_hours = sum(row["educational_hours"] for row in item["unresolved_labels"])
            note(f"The totals include {pending_shifts:,} student-shifts ({pending_hours:,} hours) recorded under "
                 "site, slot, or combined provider labels. Their recorded clinical hours are also included in the overall totals and Learner Reach. "
                 "These are listed separately below and are not credited to an individual.",
                 warning=True)

        p = doc.add_paragraph()
        p.add_run("Learner Reach: " + reach_percent(item["learner_reach_pct"])).bold = True
        p.add_run(f" — students were assigned during {item['hours_with_students']:,} of "
                  f"{item['recorded_clinical_hours']:,} recorded OPD hours. "
                  f"{item['hours_without_students']:,} hours had no student recorded.")
        note("OPD hours include shifts with and without students. Hours with students count each clinical half-day once; "
             "Learner Reach is their ratio, not an average of individual percentages.")
        doc.add_heading("Overview by type of work", level=2)
        add_effort_table([
            {"preceptor_name": group["work_type"], "months_brief": group["months_brief"],
             "no_of_shifts": group["no_of_shifts"], "educational_hours": group["educational_hours"],
             **{key: group[key] for key in ("recorded_clinical_hours", "recorded_clinical_shifts",
                 "shifts_with_students", "shifts_without_students", "hours_with_students",
                 "learner_reach_pct", "availability_review_shifts")}}
            for group in item["work_types"]
        ], first_title="Type of work / teaching months", total_title="Overall", total_metrics=item,
           academic_year=item["academic_year"])
        note(PARTICIPATION_SCOPE_NOTE)
        if item["has_unlisted_clinical_hours"]:
            note(REACH_DETAIL_TOTAL_NOTE)
        if scan.get("reporting_period"):
            note("Both reporting dates are included. Boundary months contain only the selected dates; the period is not split at July 1.")
        priority_note = outpatient_priority_report_note(scan, [year])
        if priority_note:
            note(priority_note)
        note("Academic Pediatrics combines HOPE_DRIVE, ETOWN and NYES. Ward A, PSHCH Nursery, Complex Care and other services remain separate. No additional weighting is applied by setting.")
        note(f"*One student assigned to one AM or PM shift = one student-shift and "
             f"{TEACHING_HOURS_PER_STUDENT_SHIFT} educational hours. Two students in the same shift count twice. "
             "These are student-weighted hours, not distinct clock hours or verified attendance.")
        note(REACH_SCOPE_NOTE)
        note(f"Coverage: {item['source_count']} saved rotation schedule(s) overlap this reporting period. "
             "Only archived assignments are represented; missing rotations are not assumed to have no teaching. "
             "Future scheduled assignments are included.")
        if item["has_missing_provider"]:
            note("Data review: at least one source rotation overlapping this year contains assignments without "
                 "an identifiable preceptor. Those assignments are excluded from provider totals; review Report_Notes.txt.",
                 warning=True)

        doc.add_heading("Student continuity by preceptor", level=2)
        scope = note("All work types combined for each preceptor; these counts match the individual reports.")
        scope.paragraph_format.keep_with_next = True
        if item["named_preceptors"]:
            teaching_add_work_table(
                doc,
                ("Preceptor", "Unique students", "Students assigned on 3+ days"),
                [(row["preceptor_name"], f"{row['unique_students']:,}",
                  f"{row['unique_students_3plus_days']:,}")
                 for row in item["named_preceptors"]],
                widths=(3.25, 1.5, 2.15), number_columns=(1, 2),
            )
        if item["unresolved_labels"]:
            pending = note("Provider labels awaiting an individual name (not credited to a named preceptor)", warning=True)
            pending.paragraph_format.keep_with_next = True
            teaching_add_work_table(
                doc,
                ("Provider label", "Unique students", "Students assigned on 3+ days"),
                [(row["preceptor_name"], f"{row['unique_students']:,}",
                  f"{row['unique_students_3plus_days']:,}")
                 for row in item["unresolved_labels"]],
                widths=(3.25, 1.5, 2.15), number_columns=(1, 2),
            )
        note("3+ days means at least three distinct dates in this period, not three shifts. "
             "AM and PM on the same date count as one day; days need not be consecutive. "
             "Do not add these columns for a clerkship-wide unique-student total: "
             "a student may appear under multiple preceptors.")

        doc.add_heading("Preceptor detail by type of work", level=2)
        for group in item["work_types"]:
            heading = doc.add_heading(group["work_type"], level=2)
            heading.paragraph_format.space_before = Pt(12)
            source_note = note("OPD site(s): " + ", ".join(group["source_sites"]))
            source_note.paragraph_format.keep_with_next = True
            chart = chart_map.get((year, group["work_type"]))
            if chart is None:
                raise OPDArchiveError("A clinical experience chart is missing. No incomplete chair summary was generated.")
            picture_paragraph = doc.add_paragraph()
            picture_paragraph.paragraph_format.keep_with_next = True
            picture_paragraph.paragraph_format.space_after = Pt(3)
            picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture = picture_paragraph.add_run().add_picture(BytesIO(chart["png"]), width=Inches(6.4))
            picture._inline.docPr.set("descr", chart["alt_text"])
            if group["named_preceptors"]:
                add_effort_table(group["named_preceptors"], total_title="Named preceptors subtotal",
                                 work_type=group["work_type"], academic_year=item["academic_year"])
            if group["unresolved_labels"]:
                pending_note = note("Assignments awaiting an individual preceptor name", warning=True)
                pending_note.paragraph_format.keep_with_next = True
                add_effort_table(group["unresolved_labels"], pending=True,
                                 work_type=group["work_type"], academic_year=item["academic_year"])
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(5)
                p.add_run("Work-type total: ").bold = True
                p.add_run(f"{group['no_of_shifts']:,} student-shifts | {group['educational_hours']:,} educational hours")
        source = note("Each named preceptor is counted once overall; work-type student-shifts and educational hours sum to overall totals. "
                      f"Source: current saved OPDs, retrieved {scan['generated_at']}. "
                      "Student names omitted; file-level source details are in the ZIP.")
        source.paragraph_format.space_before = Pt(2)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()
