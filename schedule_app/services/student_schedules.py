"""OPD assignment parsing, four-week student template and schedule population.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from io import BytesIO
from openpyxl import load_workbook
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import OPD_NAME_ORDERS
import io
import re
import xlsxwriter


def _opd_name_key(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def split_opd_assignment(value, roster, order=OPD_NAME_ORDERS[0]):
    """Return (preceptor, canonical_student), ignoring spaces around '~'.

    Auto mode identifies the student using the rotation list, not an assumption
    about which side contains the preceptor. Unknown students are reported by the
    caller; a provider-only cell with an empty opposite side is not assigned.
    """
    if not isinstance(value, str) or "~" not in value:
        return None, None
    left, right = (part.strip() for part in value.split("~", 1))
    left_student, right_student = roster.get(_opd_name_key(left)), roster.get(_opd_name_key(right))
    if order == "Student ~ Preceptor":
        return (right, left_student) if left_student else (None, None)
    if order == "Preceptor ~ Student":
        return (left, right_student) if right_student else (None, None)
    if left_student and right_student:
        raise OPDArchiveError("Both sides of an assignment match the student list. Choose an explicit '~' name order.")
    if left_student:
        return right, left_student
    if right_student:
        return left, right_student
    return None, None


def collect_opd_assignments(raw, students, order=OPD_NAME_ORDERS[0]):
    """Read the existing four-week AM/PM layout without changing the original OPD."""
    roster = {_opd_name_key(name): name for name in students}
    wb = load_workbook(BytesIO(raw), data_only=True)
    assignments, unmatched = [], []
    try:
        for ws in wb.worksheets:
            for shift in ("AM", "PM"):
                rows = [cell.row for cell in ws["A"] if re.match(rf"^\s*{shift}\b", str(cell.value or ""), re.I)]
                blocks, block = [], []
                for r in rows:
                    if block and r != block[-1] + 1:
                        blocks.append(block)
                        block = []
                    block.append(r)
                if block:
                    blocks.append(block)
                for week, block in enumerate(blocks[:4]):
                    for col in range(2, 9):
                        for row in block:
                            cell = ws.cell(row=row, column=col)
                            preceptor, student = split_opd_assignment(cell.value, roster, order)
                            if student is not None:
                                assignments.append({"student": student, "preceptor": preceptor,
                                                    "site": ws.title, "week": week, "shift": shift,
                                                    "column": col, "coordinate": cell.coordinate})
                            elif isinstance(cell.value, str) and "~" in cell.value:
                                if all(part.strip() for part in cell.value.split("~", 1)):
                                    unmatched.append(f"{ws.title}!{cell.coordinate}")
        return assignments, unmatched
    finally:
        wb.close()


def populate_ms_schedule(blank_bytes, assignments):
    wb = load_workbook(BytesIO(blank_bytes))
    try:
        sheets = {_opd_name_key(ws["B1"].value): ws for ws in wb.worksheets}
        for item in assignments:
            ws = sheets.get(_opd_name_key(item["student"]))
            if ws is None:
                raise OPDArchiveError("A student could not be matched to the generated schedule tabs.")
            row = (6 if item["shift"] == "AM" else 7) + 8 * item["week"]
            ws.cell(row=row, column=item["column"], value=f"{item['preceptor']} - [{item['site']}]")
        output = BytesIO()
        wb.save(output)
        return output.getvalue()
    finally:
        wb.close()


def create_ms_schedule_template(students, dates):
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})

    # — Formats —
    f1 = wb.add_format({'font_size':14,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'black','text_wrap':True,'bg_color':'#FEFFCC','border':1})
    f2 = wb.add_format({'font_size':10,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'yellow','bg_color':'black','border':1,'text_wrap':True})
    f3 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'black','bg_color':'#FFC7CE','border':1})
    f4 = wb.add_format({'num_format':'mm/dd/yyyy','font_size':12,'bold':1,'align':'center',
                        'valign':'vcenter','font_color':'black','bg_color':'#F4F6F7','border':1})
    f5 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'black','bg_color':'#F4F6F7','border':1})
    f6 = wb.add_format({'bg_color':'black','border':1})
    f7 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'black','bg_color':'#90EE90','border':1})
    f8 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                        'font_color':'black','bg_color':'#89CFF0','border':1})

    days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    start_rows = [2, 10, 18, 26]
    weeks = ['Week 1','Week 2','Week 3','Week 4']
    due_texts = [
        '',
        '',
        '',
        'All Clinical Encounter Logs are Due, Solicitation of Clinical Assessments, Observed H&Ps and Observed Handoff Due']

    used_titles = set()
    for name in students:
        safe = re.sub(r"[\[\]:*?/\\]", "-", str(name)).strip().strip("'") or "Student"
        title, suffix_number = safe[:31], 1
        while title.casefold() in used_titles:
            suffix_number += 1
            suffix = f"_{suffix_number}"
            title = safe[:31-len(suffix)] + suffix
        used_titles.add(title.casefold())
        ws = wb.add_worksheet(title)
        ws.set_zoom(70)

        # Header
        ws.merge_range('A1:A2','Student Name:', f1)
        ws.merge_range('B1:B2',      str(name),      f1)
        #note = ("*Note* Protected Self-Study Time is for coursework only. During this time period, "
        #        "we expect students to do coursework, be available for any additional educational "
        #        "activities, and any extra clinical time that may be available. If the student is not "
        #        "available during this time period and has not made an absence request, the student "
        #        "will be cited for unprofessionalism and will risk failing the course.")

        note = ("*Note* Protected Self-Study Time is reserved for independent learning and completion of required "
                "coursework. Students are encouraged to use this time to complete assignments, review course materials, "
                "prepare for patient care, and reinforce concepts encountered during the clerkship.")

        ws.merge_range('C1:H2', note, f2)

        # Column widths & row height
        ws.set_column('A:A', 20)
        ws.set_column('B:B', 30)
        ws.set_column('C:G', 40)
        ws.set_column('H:H',155)
        ws.set_row(0, 37.25)

        # Days headers and dates
        date_idx = 0
        for block, row in enumerate(start_rows):
            # 1) write the days on row `row`, cols B–H
            for col_offset, day in enumerate(days, start=1):
                ws.write(row, col_offset, day, f3)

        # 2) write the dates directly beneath in B–H (row+1)
            for col_offset in range(7):
                if date_idx < len(dates):
                    # note the +1 here instead of +2
                    ws.write(row+1, col_offset+1, dates[date_idx], f4)
                    date_idx += 1

        # Week labels
        for i, week in enumerate(weeks):
            row = 4 + (i * 8)
            ws.write(f'A{row}', week, f3)

        # AM / PM labels
        for i in range(4):
            ws.write(f'A{6 + i*8}', 'AM', f3)
            ws.write(f'A{7 + i*8}', 'PM', f3)

        # Fill AM/PM blocks with Asynchronous Time (cols C–J)
        for block in range(4):
            am_row = 5 + block*8
            pm_row = 6 + block*8
            for col in range(1, 8):
                ws.write(am_row, col, "Protected Self-Study Time", f5)
                ws.write(pm_row, col, "Protected Self-Study Time", f5)

        # Separators
        for sep in [10, 18, 26, 34]:
            ws.merge_range(f'A{sep}:H{sep}', '', f6)

        # Green filler rows
        for filler in [8, 16, 24, 32]:
            for col in range(8):
                ws.write(filler, col, ' ', f7)

        # Assignment‑due rows
        for i, base in enumerate([8, 16, 24, 32]):
            ws.write(f'A{base}', 'ASSIGNMENT DUE:', f8)
            for col in range(1, 8):
                if col == 5:
                    ws.write(base-1, col, 'Ask for Feedback!', f8)
                elif col == 7:
                    ws.write(base-1, col, due_texts[i], f8)
                else:
                    ws.write(base-1, col, ' ', f8)

    wb.close()
    buf.seek(0)
    return buf
