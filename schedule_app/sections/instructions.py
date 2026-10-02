"""Schedule download instructions; wording is managed in protected Admin."""
from datetime import datetime, timedelta
import streamlit as st
from schedule_app.services.report_wording import report_text as rt
from schedule_app.reports.qgenda_instructions import make_qgenda_instructions
from schedule_app.sections.report_wording_controls import with_saved_report_wording


@with_saved_report_wording
def render():
    d = st.text_input('Start date (m/d/yyyy)')
    if not d:
        return
    try:
        s = datetime.strptime(d, '%m/%d/%Y')
        e = s + timedelta(days=34)
    except ValueError:
        st.error('Invalid format – use m/d/yyyy (e.g. 7/6/2021)')
        return
    st.write(f'{s:%B %d, %Y} → {e:%B %d, %Y}')
    st.write(rt('instructions.screen_intro'))
    st.markdown(rt('instructions.screen_summary', start_long=f'{s:%B %d, %Y}', end_long=f'{e:%B %d, %Y}'))
    st.write(rt('instructions.screen_download'))
    st.download_button(label='📄 Download Instructions (Word)', data=make_qgenda_instructions(s, e),
                       file_name='Qgenda_Report_Instructions.docx',
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
