import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from io import BytesIO

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="PEDATI Master Planner", layout="wide")
st.title("🎓 PEDATI SMART LESSON PLANNER")

# --- MAIN PAGE CONFIGURATION & USER API KEY BAR (AT THE VERY TOP) ---
user_api_key = st.text_input(
    "🔑 ENTER YOUR GEMINI API KEY:", 
    type="password", 
    help="Get your API key from Google AI Studio using your Gmail account."
)

# Helper function to dynamically check and load models based on the user's key
def get_working_model(api_key):
    try:
        genai.configure(api_key=api_key)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except Exception as e:
        st.error(f"INVALID API KEY OR CONNECTION ERROR: {str(e)}")
        return None
    return "models/gemini-1.5-flash"  # Default fallback


# Process model assignment if the key is provided
selected_model_name = None
if user_api_key:
    selected_model_name = get_working_model(user_api_key)
    if selected_model_name:
        st.info(f"SYSTEM CONNECTED VIA YOUR API KEY. ACTIVE MODEL: {selected_model_name.upper()}")
else:
    st.warning("⚠️ PLEASE ENTER YOUR PERSONAL GEMINI API KEY ABOVE TO START.")


# --- 2. AI GENERATION ENGINE ---
def generate_pedati_plan(topic, syllabus, extra_context, api_key, model_name):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    
    prompt = f"""
    Topic: {topic}. Syllabus Code: {syllabus}. Context: {extra_context}.
    Generate a lesson plan in English. 
    NO Malay terms.
    
    CRITICAL TEXT FORMATTING RULES:
    1. DO NOT use double asterisks (**) anywhere in the output.
    2. Ensure every single section title and stage marker is in full CAPITAL LETTERS.
    3. Use numbers for lists where appropriate.
    
    Use these exact stage names:
    P [Prior Knowledge], E [Engage], D [Develop], A [Apply], T [Test], I [Improve].

    Structure with these exact markers for boxing:
    SECTION: LESSON OBJECTIVES
    [4 points]
    SECTION: LESSON OUTCOMES
    [4 points]
    SECTION: SUCCESS CRITERIA
    [4 points]
    SECTION: PREREQUISITE
    [1 point]
    SECTION: KEYWORDS
    [6 items]
    SECTION: HOTS
    [any 4 main domains in the Bloom's taxonomy]
    SECTION: DIGITAL CITIZENSHIP
    [4 points on the use of online resources like youtube channel or canva application or use of chromebooks or use of digital devices]

    SECTION: PEDATI STAGES
    STAGE: P [PRIOR KNOWLEDGE] | CB: [Activity] | SB: [Activity]
    STAGE: E [ENGAGE] | CB: [Activity] | SB: [Activity]
    STAGE: D [DEVELOP] | CB: [Activity] | SB: [Activity]
    STAGE: A [APPLY] | CB: [Activity] | SB: [Activity]
    STAGE: T [TEST] | CB: [Activity] | SB: [Activity]
    STAGE: I [IMPROVE] | CB: [Activity] | SB: [Activity]
    """
    try:
        response = model.generate_content(prompt)
        # Explicit clean-up replacement step to wipe out double asterisks completely
        return response.text.replace("**", "")
    except Exception as e:
        return f"SYSTEM ERROR: {str(e)}"


# --- 3. WORD DOCUMENT EXPORT ENGINE (14PT FONT & SINGLE SPACING) ---
def create_word_export(topic, syllabus, text):
    doc = Document()
    
    # Global document layout formatting override configurations
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(14)  # Set 14 points font size
    
    p_format = style.paragraph_format
    p_format.line_spacing = 1.0  # Set 1 paragraph / single spacing
    p_format.space_after = Pt(12)  # Maintain clean paragraph block dynamic gap

    # Document Header Title - FULL CAPITAL LETTERS
    title_p = doc.add_paragraph()
    run_title = title_p.add_run(f'LESSON PLAN: {topic.upper()} ({syllabus.upper()})')
    run_title.bold = True
    run_title.font.size = Pt(18)

    # 1. Admin Header Table (6-field layout)
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [["WEEK NO:", "DATE:"], ["NO. OF STUDENTS:", "DAY:"], ["VENUE / LAB NO:", "DURATION (MINS):"]]
    for r in range(3):
        admin_table.cell(r, 0).paragraphs[0].add_run(labels[r][0]).bold = True
        admin_table.cell(r, 2).paragraphs[0].add_run(labels[r][1]).bold = True
    doc.add_paragraph()

    # 2. Resources Table
    p_res = doc.add_paragraph()
    p_res.add_run("RESOURCES & MATERIALS").bold = True
    res_table = doc.add_table(rows=1, cols=1)
    res_table.style = 'Table Grid'
    res_cell_p = res_table.cell(0, 0).paragraphs[0]
    res_cell_p.paragraph_format.line_spacing = 1.0
    res_cell_p.add_run("Smart board, Chromebook, Writing table, Projector, Screen share with laptop")
    doc.add_paragraph()

    # 3. Content Parsing & Table Boxing with Asterisk Filtering
    sections = text.split('SECTION:')
    for section in sections:
        if not section.strip(): 
            continue
            
        lines = section.strip().split('\n')
        title = lines[0].strip().upper().replace("**", "")  # Enforce FULL CAPITAL LETTERS for titles
        content_lines = lines[1:]

        p_sec = doc.add_paragraph()
        p_sec.add_run(title).bold = True

        if "|" in section and "PEDATI" in title:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            
            hdr[0].paragraphs[0].add_run('STAGE (PEDATI)').bold = True
            hdr[1].paragraphs[0].add_run(' ACTIVITY 1 ').bold = True
            hdr[2].paragraphs[0].add_run(' ACTIVITY 2 ').bold = True

            for line in content_lines:
                cleaned_line = line.replace("**", "")
                if "|" in cleaned_line:
                    p = cleaned_line.split("|")
                    row_cells = table.add_row().cells
                    
                    # Layout formatting rule injections to cell text elements
                    for cell in row_cells:
                        cell.paragraphs[0].paragraph_format.line_spacing = 1.0
                    
                    row_cells[0].paragraphs[0].add_run(p[0].split(":")[-1].strip().upper()) # Keep stages inside table clean and capitalized
                    row_cells[1].paragraphs[0].add_run(p[1].split(":")[-1].strip())
                    row_cells[2].paragraphs[0].add_run(p[2].split(":")[-1].strip())
            doc.add_paragraph()
        else:
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            cell_p = table.cell(0, 0).paragraphs[0]
            cell_p.paragraph_format.line_spacing = 1.0
            
            cleaned_body = "\n".join([l.strip() for l in content_lines if l.strip()]).replace("**", "")
            cell_p.add_run(cleaned_body)
            doc.add_paragraph()

    # 4. HOD Approval Page
    doc.add_page_break()
    p_hod = doc.add_paragraph()
    p_hod.add_run("HOD APPROVAL & REMARKS").bold = True
    
    hod_table = doc.add_table(rows=3, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).paragraphs[0].add_run("REMARK").bold = True
    hod_table.cell(0, 1).paragraphs[0].add_run("SIGNATURE / STAMP").bold = True
    hod_table.rows[1].height = Pt(40)
    hod_table.cell(2, 0).paragraphs[0].add_run("DATE:").bold = True
    hod_table.cell(2, 1).paragraphs[0].add_run("NAME:").bold = True

    # Adjust spacing across all generated grid cells
    for row in admin_table.rows:
        for cell in row.cells: cell.paragraphs[0].paragraph_format.line_spacing = 1.0
    for row in hod_table.rows:
        for cell in row.cells: cell.paragraphs[0].paragraph_format.line_spacing = 1.0

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- 4. MAIN GUI INTERFACE ---
st.write("---")
u_topic = st.text_input("LESSON TOPIC:")
u_syllabus = st.text_input("SYLLABUS CODE:")
u_extra = st.text_area("SPECIFIC CONTEXT / KEYWORDS (OPTIONAL):")

if st.button("🚀 GENERATE PEDATI LESSON PLAN", type="primary"):
    if not user_api_key:
        st.error("❌ KEY CONFIGURATION ERROR! PLEASE INPUT YOUR GOOGLE GEMINI API KEY AT THE TOP OF THE PAGE FIRST.")
    elif not u_topic or not u_syllabus:
        st.error("❌ PLEASE PROVIDE BOTH A LESSON TOPIC AND A SYLLABUS CODE.")
    else:
        with st.spinner("AI IS BUILDING YOUR PEDATI PLAN..."):
            result = generate_pedati_plan(u_topic, u_syllabus, u_extra, user_api_key, selected_model_name)
            st.session_state['pedati_out'] = result

if 'pedati_out' in st.session_state:
    st.divider()
    st.subheader("👁️ AI PREVIEW")
    st.text_area("GENERATED CONTENT CONTENT PREVIEW", st.session_state['pedati_out'], height=350)
    
    doc_file = create_word_export(u_topic, u_syllabus, st.session_state['pedati_out'])
    st.download_button(
        label="📥 DOWNLOAD WORD (.DOCX)", 
        data=doc_file, 
        file_name=f"PEDATI_LP_{u_topic.upper().replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# --- FOOTER SECTION ---
st.markdown("---") 
st.markdown(
    """
    <div style='text-align: center; color: grey; font-size: 0.8em;'>
        <p><b>SMART PEDATI LESSON PLAN AI-GENERATOR V1.0</b></p>
        <p>CONCEPTUALIZED BY: <b>HAJAH NURUL HAZIQAH @ HJH HARTINI HJ NORDIN</b></p>
        <p>© 2026 BSC H.M IN COMPUTER SCIENCE, UNIVERSITY OF STRATHCLYDE</p>
    </div>
    """,
    unsafe_allow_html=True
)
