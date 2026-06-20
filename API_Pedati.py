import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from io import BytesIO

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="PEDATI Master Planner", layout="wide")
st.title("🎓 PEDATI Lesson Plan Generator")

# --- SIDEBAR FOR USER API KEY ---
st.sidebar.header("🔑 Authentication")
user_api_key = st.sidebar.text_input(
    "Enter your Gemini API Key:", 
    type="password", 
    help="Get your API key from Google AI Studio using your Gmail account."
)

# Helper function to dynamically check and load models based on the user's key
def get_working_model(api_key):
    try:
        genai.configure(api_key=api_key)
        # Attempt to list models to verify if the key works
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except Exception as e:
        st.sidebar.error(f"Invalid API Key or Connection Error: {str(e)}")
        return None
    return "models/gemini-1.5-flash"  # Default fallback


# Process model assignment if the key is provided
selected_model_name = None
if user_api_key:
    selected_model_name = get_working_model(user_api_key)
    if selected_model_name:
        st.info(f"System connected via your API key. Active Model: {selected_model_name}")
else:
    st.warning("⚠️ Please enter your personal Gemini API Key in the sidebar to start.")


def generate_pedati_plan(topic, syllabus, extra_context, api_key, model_name):
    # Enforce key presence
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    
    prompt = f"""
    Topic: {topic}. Syllabus Code: {syllabus}. Context: {extra_context}.
    Generate a lesson plan in English. 
    NO Malay terms. Use these exact stage names:
    P [Prior Knowledge], E [Engage], D [Develop], A [Apply], T [Test], I [Improve].

    Structure with these markers for boxing:
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
    STAGE: P [Prior Knowledge] | CB: [Activity] | SB: [Activity]
    STAGE: E [Engage] | CB: [Activity] | SB: [Activity]
    STAGE: D [Develop] | CB: [Activity] | SB: [Activity]
    STAGE: A [Apply] | CB: [Activity] | SB: [Activity]
    STAGE: T [Test] | CB: [Activity] | SB: [Activity]
    STAGE: I [Improve] | CB: [Activity] | SB: [Activity]
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"System Error: {str(e)}"


def create_word_export(topic, syllabus, text):
    doc = Document()
    doc.add_heading(f'Lesson Plan: {topic} ({syllabus})', 0)

    # 1. Admin Header Table (6-field layout)
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [["Week No :", "Date:"], ["No. of Students:", "Day:"], ["Venue / Lab No:", "Duration (mins):"]]
    for r in range(3):
        admin_table.cell(r, 0).text = labels[r][0]
        admin_table.cell(r, 2).text = labels[r][1]
    doc.add_paragraph()

    # 2. Resources Table
    doc.add_heading("Resources & Materials", level=1)
    res_table = doc.add_table(rows=1, cols=1)
    res_table.style = 'Table Grid'
    res_table.cell(0, 0).text = "Smart board, Chromebook, Writing table, Projector, Screen share with laptop"

    # 3. Content Parsing & Table Boxing
    sections = text.split('SECTION:')
    for section in sections:
        if not section.strip(): continue
        lines = section.strip().split('\n')
        title = lines[0].strip()
        content_lines = lines[1:]
        doc.add_heading(title.title(), level=1)

        if "|" in section and "PEDATI" in title.upper():
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Stage (PEDATI)', 'Facilitator (CB)', 'Student (SB)'

            for line in content_lines:
                if "|" in line:
                    p = line.split("|")
                    row = table.add_row().cells
                    row[0].text = p[0].split(":")[-1].strip()
                    row[1].text = p[1].split(":")[-1].strip()
                    row[2].text = p[2].split(":")[-1].strip()
        else:
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "\n".join([l.strip() for l in content_lines if l.strip()])

    # 4. HOD Approval Page
    doc.add_page_break()
    doc.add_heading("HOD Approval & Remarks", level=1)
    hod_table = doc.add_table(rows=3, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).text = "Remark"
    hod_table.cell(0, 1).text = "Signature / Stamp"
    hod_table.rows[1].height = Pt(60)
    hod_table.cell(2, 0).text = "Date:"
    hod_table.cell(2, 1).text = "Name:"

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- 2. MAIN GUI ---
c1, c2 = st.columns(2)
with c1: u_topic = st.text_input("Lesson Topic:")
with c2: u_syllabus = st.text_input("Syllabus Code:")
u_extra = st.text_area("Specific Context/Keywords (Optional):")

if st.button("🚀 GENERATE PEDATI LESSON PLAN"):
    if not user_api_key:
        st.error("❌ Please input your Google Gemini API key in the sidebar before clicking generate.")
    elif not u_topic or not u_syllabus:
        st.error("❌ Please provide both a Lesson Topic and a Syllabus Code.")
    else:
        with st.spinner("AI is building your PEDATI plan..."):
            result = generate_pedati_plan(u_topic, u_syllabus, u_extra, user_api_key, selected_model_name)
            st.session_state['pedati_out'] = result

if 'pedati_out' in st.session_state:
    st.divider()
    st.text_area("AI Preview", st.session_state['pedati_out'], height=300)
    doc_file = create_word_export(u_topic, u_syllabus, st.session_state['pedati_out'])

    st.download_button("📥 Download Word (.docx)", doc_file, f"PEDATI_{u_topic}.docx")

# --- FOOTER SECTION ---
st.markdown("---") 
st.markdown(
    """
    <div style='text-align: center; color: grey; font-size: 0.8em;'>
        <p><b>Smart PEDATI Lesson Plan AI-Generator v1.0</b></p>
        <p>Developed & Conceptualized by: <b>[Hajah Nurul Haziqah @ Hjh Hartini Hj Nordin]</b></p>
        <p>© 2026 PTES Academic Innovation Computer Science</p>
    </div>
    """,
    unsafe_allow_html=True
)
