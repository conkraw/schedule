"""Individual preceptor Word report, by academic year and type of work.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from collections import defaultdict
from datetime import date as CalendarDate
from docx import Document
from docx.shared import Pt
from io import BytesIO
from schedule_app.reports.teaching_tables import teaching_add_work_table
from schedule_app.services.teaching_analysis import teaching_academic_label
from schedule_app.services.teaching_analysis import teaching_brief_months
from schedule_app.services.teaching_analysis import teaching_month_label
from schedule_app.services.teaching_analysis import teaching_require_work_type_data
from schedule_app.services.teaching_analysis import teaching_work_type_sort
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT
from schedule_app.settings import TEACHING_WORK_TYPE_REVIEW


def teaching_make_docx(name, monthly, scan):
    """One document per preceptor; academic years and work types stay separate."""
    from docx.shared import Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    teaching_require_work_type_data(scan)
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
    for style_name, size in (("Title", 23), ("Heading 1", 16), ("Heading 2", 12)):
        style = doc.styles[style_name]
        style.font.name, style.font.size = "Calibri", Pt(size)
        style.font.color.rgb = RGBColor.from_string("24466B")
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8 if style_name == "Heading 2" else 0)
        style.paragraph_format.space_after = Pt(4)
        properties = style.element.find(qn("w:pPr"))
        if properties is not None:
            border = properties.find(qn("w:pBdr"))
            if border is not None:
                properties.remove(border)
    subtitle = doc.styles["Subtitle"]
    subtitle.font.name, subtitle.font.size = "Calibri", Pt(12)
    subtitle.font.italic = False
    subtitle.font.color.rgb = RGBColor.from_string("526475")
    subtitle.paragraph_format.space_after = Pt(7)
    subtitle.paragraph_format.keep_with_next = True
    header = section.header.paragraphs[0]
    header.text = "PENN STATE  |  PEDIATRIC CLERKSHIP"
    header.runs[0].font.size = Pt(9)
    header.runs[0].font.color.rgb = RGBColor.from_string("526475")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("Third-year student teaching  |  Page ").font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.core_properties.title = f"Preceptor teaching report - {name}"
    doc.core_properties.author = "Pediatric Clerkship"
    doc.core_properties.subject = "Scheduled student-shifts by academic year and type of work"

    def note(text, warning=False):
        paragraph = doc.add_paragraph(text)
        for run in paragraph.runs:
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string("8C3B25" if warning else "526475")
        return paragraph

    grouped = defaultdict(list)
    for row in monthly:
        grouped[row["academic_start_year"]].append(row)
    type_rows = [row for row in scan["monthly_by_work_type"] if row["preceptor_name"] == name]
    for index, (year, rows) in enumerate(sorted(grouped.items())):
        if index:
            doc.add_page_break()
        doc.add_paragraph("Preceptor teaching report", style="Subtitle")
        doc.add_paragraph(name, style="Title")
        doc.add_heading(f"Academic year {teaching_academic_label(year)}", level=1)
        doc.add_paragraph(f"July 1, {year} - June 30, {year + 1}")
        if name in scan["unresolved_preceptor_labels"]:
            note("Review required: this is a site, slot, or combined provider label, not a verified individual preceptor.", warning=True)
        total = sum(row["no_of_shifts"] for row in rows)
        p = doc.add_paragraph()
        p.add_run("All work types: ").bold = True
        p.add_run(f"{total:,} assigned student-shifts  |  {total * TEACHING_HOURS_PER_STUDENT_SHIFT:,} educational hours")
        by_type = defaultdict(list)
        for row in type_rows:
            if row["academic_start_year"] == year:
                by_type[row["work_type"]].append(row)
        doc.add_heading("Educational effort by type of work", level=2)
        overview = []
        for work_type, items in sorted(by_type.items(), key=lambda pair: teaching_work_type_sort(pair[0])):
            count = sum(item["no_of_shifts"] for item in items)
            overview.append((work_type, teaching_brief_months(item["month"] for item in items),
                             f"{count:,}", f"{count * TEACHING_HOURS_PER_STUDENT_SHIFT:,}"))
        teaching_add_work_table(doc,
            ("Type of work", "Months with assignments", "Student-shifts", "Educational hours*"), overview,
            widths=(2.45, 2.15, 1.05, 1.25), number_columns=(2, 3),
            total=("All work types", "", f"{total:,}", f"{total * TEACHING_HOURS_PER_STUDENT_SHIFT:,}"))
        note("Academic Pediatrics combines HOPE_DRIVE, ETOWN and NYES. Ward A, PSHCH Nursery, Complex Care and other services remain separate. No additional weighting is applied by setting.")
        note(f"*One student assigned to one AM or PM shift = one student-shift and {TEACHING_HOURS_PER_STUDENT_SHIFT} educational hours. Two students in the same shift count twice. These are student-weighted scheduled hours, not distinct clock hours or verified attendance.")
        if TEACHING_WORK_TYPE_REVIEW in by_type:
            note("Work type needs review: the same assignment appears under different work types. It is counted once in this review category, not credited twice or assigned to a guessed setting.", warning=True)
        doc.add_heading("Monthly detail by type of work", level=2)
        month_values = []
        for work_type, items in sorted(by_type.items(), key=lambda pair: teaching_work_type_sort(pair[0])):
            for item in sorted(items, key=lambda item: item["month"]):
                month_values.append((work_type, teaching_month_label(CalendarDate.fromisoformat(item["month"])),
                                     f"{item['no_of_shifts']:,}",
                                     f"{item['no_of_shifts'] * TEACHING_HOURS_PER_STUDENT_SHIFT:,}"))
        teaching_add_work_table(doc,
            ("Type of work", "Month", "Student-shifts", "Educational hours*"), month_values,
            widths=(2.45, 2.15, 1.05, 1.25), number_columns=(2, 3),
            total=("All work types", "", f"{total:,}", f"{total * TEACHING_HOURS_PER_STUDENT_SHIFT:,}"))
        note("Source: current encrypted OPD archive files, decrypted for this report. "
             f"Archive snapshot: {scan['commit'][:12]}. Retrieved: {scan['generated_at']}. "
             "Only months with assignments are shown. Future scheduled assignments are included. Student names are omitted.")
    output = BytesIO()
    doc.save(output)
    return output.getvalue()
