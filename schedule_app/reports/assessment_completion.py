"""Compact provider-level completion tables. Never export student identifiers."""
from schedule_app.reports.teaching_tables import teaching_add_work_table
from schedule_app.services.assessment_completion import (
    completion_rows, completion_display, assessment_method_note, completion_threshold, SCOPE_NOTE,
)


def append_individual_completion(doc, scan, name, year):
    bundle = scan.get("assessment_completion")
    if bundle is None:
        return
    rows = [r for r in completion_rows(bundle, scan, year) if r["preceptor_name"] == name]
    if not rows:
        return
    row = rows[0]
    heading = doc.add_heading("Documented assessment completion", level=2)
    # Keep teaching time on its own readable page; feedback follows separately.
    heading.paragraph_format.page_break_before = True
    denominator = row.get("eligible_students")
    minimum = completion_threshold(bundle)
    title = f"Students assigned for {minimum}+ shifts" if minimum is not None else "Eligible students"
    doc.add_paragraph("Assessments as of: " + bundle["assessments_as_of"])
    doc.add_paragraph(title + " by this cutoff: " + (str(denominator) if denominator is not None else "Not checked"))
    teaching_add_work_table(doc, ("Assessment form", "Eligible students assessed"), [
        ("Clinical Assessment of Student", completion_display(row, "clinical")),
        ("History Taking & Physical Exam", completion_display(row, "hp")),
        ("At least one of these forms", completion_display(row, "either")),
    ], widths=(4.2, 2.7), number_columns=(1,))
    doc.add_paragraph(assessment_method_note(completion_threshold(bundle), as_of=bundle["assessments_as_of"]))
    if str(row["assessment_status"]).startswith("Provisional:"):
        doc.add_paragraph("* " + row["assessment_status"] + ". The displayed percentage uses confirmed matches only; "
                          "every eligible OPD student remains in the denominator. Review names in PTS Matching when needed.")
    elif row["assessment_status"] != "Calculated":
        doc.add_paragraph("Review: " + row["assessment_status"] + ". The teaching-hours report is unaffected.")
    else:
        doc.add_paragraph("Each cell shows evaluated eligible students / all eligible students (percentage). "
                          "A student assessed twice still counts once. The combined row counts either form, not the sum.")
    if row.get("either_students_without_assessment"):
        doc.add_paragraph(f"No assessment on file for {row['either_students_without_assessment']} of {row['eligible_students']} eligible students. "
                          "No name confirmation is needed solely because a student is absent from OASIS.")
    doc.add_paragraph(SCOPE_NOTE)
    if any(r["preceptor_name"] == "Unattributed / other evaluator" and r["academic_year"] == row["academic_year"]
           for r in bundle["warnings"]):
        doc.add_paragraph("Some source forms could not be attributed to an evaluator and are not credited. "
                          "These source issues are listed in the app and chair summary; percentages describe identifiable matched records.")
    issues = [r for r in bundle["warnings"] if r["preceptor_name"] == name and r["academic_year"] == row["academic_year"]
              and r["direction"] != "Assessment on file"]
    if issues:
        doc.add_heading("Evaluation records to review", level=2)
        teaching_add_work_table(doc, ("Direction", "Review note"),
            [(r["direction"], r["issue"]) for r in issues], widths=(1.65, 5.25))
    doc.add_paragraph("Source: archived OASIS student assessments; Submit Date uses the report period through the assessment cutoff. "
                      f"Check retrieved: {bundle['retrieved_at']}. Student names, IDs, scores and assessment comments are omitted.")


def append_chair_completion(doc, scan, year):
    bundle = scan.get("assessment_completion")
    if bundle is None:
        return
    rows = completion_rows(bundle, scan, year)
    if not rows:
        return
    heading = doc.add_heading("Documented assessment completion by preceptor", level=2)
    heading.paragraph_format.page_break_before = True
    minimum = completion_threshold(bundle)
    doc.add_paragraph("Assessments as of: " + bundle["assessments_as_of"] + ". This cutoff does not change teaching-hour totals.")
    students_title = f"Students\n{minimum}+ shifts" if minimum is not None else "Eligible\nstudents"
    teaching_add_work_table(doc,
        ("Preceptor", students_title, "Clinical\nAssessment", "History &\nPhysical", "Either form"),
        [(r["preceptor_name"], str(r["eligible_students"]) if r["eligible_students"] is not None else "Not checked",
          completion_display(r, "clinical"), completion_display(r, "hp"), completion_display(r, "either")) for r in rows],
        widths=(1.85, 0.8, 1.45, 1.4, 1.4), number_columns=(1, 2, 3, 4))
    doc.add_paragraph("Cells show confirmed assessments / all eligible students (percentage). "
                      "* Provisional: possible name discrepancies remain; no similar-name record was credited automatically. "
                      "Each student counts once per preceptor/form type. 'Either form' counts the union, not the sum.")
    doc.add_paragraph(assessment_method_note(completion_threshold(bundle), as_of=bundle["assessments_as_of"]))
    doc.add_paragraph("These are per-preceptor measures across all work types. Do not sum unique students across preceptors. "
                      "The existing 3+ days continuity measure remains separate.")
    doc.add_paragraph(SCOPE_NOTE)
    wanted = {r["preceptor_name"] for r in rows}
    labels = {r["academic_year"] for r in rows}
    issues = [r for r in bundle["warnings"] if (r["preceptor_name"] in wanted or r["preceptor_name"] == "Unattributed / other evaluator") and r["academic_year"] in labels
              and r["direction"] != "Assessment on file"]
    no_record_checks = sum(bool(row.get("either_students_without_assessment")) for row in rows)
    if no_record_checks:
        doc.add_paragraph(f"{no_record_checks} preceptor result(s) include eligible students with no assessment on file. "
                          "Those students stay in the denominator. This does not mean their assessments are overdue.")
    if issues:
        doc.add_heading("Evaluation records to review", level=2)
        teaching_add_work_table(doc, ("Preceptor", "Direction", "Review note"),
            [(r["preceptor_name"], r["direction"], r["issue"]) for r in issues], widths=(1.75, 1.3, 3.85))
    doc.add_paragraph(f"Source: {bundle['source_count']} archived student-assessment CSV(s). Check retrieved: {bundle['retrieved_at']}. "
                      "Missing-data alerts do not stop teaching reports; unverified counts are not represented as zero.")
