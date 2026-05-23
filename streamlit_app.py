import streamlit as st
import json
import os
from pypdf import PdfReader
from google import genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import io

# ==============================================================================
# HIGH-END ENTERPRISE INTERFACE CUSTOM GLASSMOPHISM STYLING (CSS INJECTION)
# ==============================================================================
st.set_page_config(page_title="Enterprise ESG Engine", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* Global Background & Body Fonts */
    .stApp {
        background-color: #F8FAFC !important;
        font-family: 'Inter', 'Segoe UI', system-ui, sans-serif !important;
    }
    
    /* Executive Main Title Banner */
    .main-header-card {
        background: linear-gradient(135deg, #0B192C 0%, #1E3E62 100%);
        padding: 2.5rem 2rem;
        border-radius: 12px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(11, 25, 44, 0.15);
        margin-bottom: 2rem;
    }
    .main-header-card h1 {
        color: #FFFFFF !important;
        font-size: 2.2rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.05em;
        margin: 0 0 0.5rem 0 !important;
    }
    .main-header-card p {
        color: #94A3B8 !important;
        font-size: 1rem !important;
        margin: 0 !important;
    }
    
    /* Clean Cards for Drag & Drop Workspace Zones */
    .workspace-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
    }
    
    /* Custom Styling for Streamlit Native Elements via Selectors */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #CBD5E1 !important;
        border-radius: 8px !important;
        background: #F1F5F9 !important;
        padding: 1rem !important;
    }
    
    /* Custom High-Status Corporate Tabs */
    button[data-baseweb="tab"] {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #64748B !important;
        border-bottom: 2px solid transparent !important;
        padding: 0.75rem 1.5rem !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #0B192C !important;
        border-bottom: 3px solid #1E3E62 !important;
    }
    
    /* Premium Data Grid Metric Badges */
    .omitted-badge {
        background-color: #FFE4E6 !important;
        color: #991B1B !important;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        border: 1px solid #FCA5A5;
        display: inline-block;
    }
    
    /* Download Call-To-Action Button overrides */
    div.stDownloadButton > button {
        background: #10B981 !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.75rem 2rem !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.2) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stDownloadButton > button:hover {
        background: #059669 !important;
        transform: translateY(-2px) !important;
    }
    </style>
""", unsafe_allow_html=True)

# Render Header Card View Component
st.markdown("""
    <div class="main-header-card">
        <h1>🏛️ GovGuard: Enterprise ESG Core Engine</h1>
        <p>SEC Automated Compliance Pipeline • Mapped to IFRS S1/S2, SOE Act 2023 & ISO 20400 Protocols</p>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# STATE & CACHE CONTROLLERS
# ==============================================================================
MAX_FILE_SIZE_MB = 100
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

if "uploaded_report_text" not in st.session_state:
    st.session_state.uploaded_report_text = ""
if "gazette_override_text" not in st.session_state:
    st.session_state.gazette_override_text = ""
if "pipeline_completed" not in st.session_state:
    st.session_state.pipeline_completed = False
if "audit_data_cache" not in st.session_state:
    st.session_state.audit_data_cache = {}

# ==============================================================================
# MS-WORD PRODUCTION EXPORT STRUCTURING FUNCTIONS
# ==============================================================================
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def apply_executive_formatting(doc, title_text):
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1), Inches(1)
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(11)
    
    p = doc.add_paragraph()
    run = p.add_run(title_text.upper())
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor(11, 25, 44)
    
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '18')
    bottom.set(qn('w:space'), '12')
    bottom.set(qn('w:color'), '0B192C')
    pBdr.append(bottom)
    pPr.append(pBdr)
    doc.add_paragraph("").paragraph_format.space_after = Pt(18)

def add_executive_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(11, 25, 44)
    else:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(30, 62, 98)

def add_justified_paragraph(doc, text, space_after=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)

def add_executive_data_table(doc, headers, rows, hex_header_color="0B192C"):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = True
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_header_color}"/>')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)
        for run in hdr_cells[i].paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(10.5)
            
    for r_idx, row_data in enumerate(rows):
        row_cells = table.add_row().cells
        bg_color = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, cell_value in enumerate(row_data):
            row_cells[c_idx].text = str(cell_value)
            set_cell_margins(row_cells[c_idx], top=100, bottom=100, left=150, right=150)
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            row_cells[c_idx]._tc.get_or_add_tcPr().append(shading)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.15
            for run in p.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(10)
                if "[OMITTED]" in str(cell_value).upper() or "[TO BE COMPLETED]" in str(cell_value).upper():
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(153, 27, 27)
    doc.add_paragraph("").paragraph_format.space_after = Pt(12)

# ==============================================================================
# UI NAVIGATION PANELS
# ==============================================================================
tab1, tab2, tab3 = st.tabs([
    "📥 Ingestion Workspace", 
    "📊 Gap Audit Summary", 
    "📜 Comprehensive ESG Report Portfolio"
])

with tab1:
    st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
    st.subheader("🛠️ Data Asset Dropzone")
    col1, col2 = st.columns(2)
    with col1:
        uploaded_report = st.file_uploader("Corporate Operating Profile (PDF):", type=["pdf"])
    with col2:
        uploaded_gazette = st.file_uploader("SECP Circular / Gazette Patches (Optional PDF):", type=["pdf"])
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_report and not st.session_state.uploaded_report_text:
        if uploaded_report.size > MAX_FILE_SIZE_BYTES:
            st.error("❌ Transmission error: Ingestion payload limited to 100MB.")
        else:
            with st.spinner("Compiling cross-border text systems..."):
                reader = PdfReader(uploaded_report)
                st.session_state.uploaded_report_text = "".join([page.extract_text() or "" for page in reader.pages])
                st.success("Master organizational data uploaded cleanly to session memory cache.")
                
    if uploaded_gazette and not st.session_state.gazette_override_text:
        with st.spinner("Syncing secondary legal parameters..."):
            reader = PdfReader(uploaded_gazette)
            st.session_state.gazette_override_text = "".join([page.extract_text() or "" for page in reader.pages])
            st.info("Live legislative system override file patched successfully.")

# ==============================================================================
# CORE AI COMPLIANCE EXTRACTION MATRIX
# ==============================================================================
if st.session_state.uploaded_report_text and not st.session_state.pipeline_completed:
    override_rules = st.session_state.gazette_override_text if st.session_state.gazette_override_text else "No modifications filed."
    
    orchestration_prompt = f"""
    You are an enterprise regulatory audit system. Process the text payload against the SECP ESG Disclosure Guidelines 2023, Companies Act 2017 (Sec 238), and the SOE Act 2023.
    Generate an intense, highly-detailed disclosure suite containing strict itemized data rows.
    
    If any SECP-required variable listed below is missing from the source text corpus, you must explicitly value map it as "[OMITTED / DATA NOT AVAILABLE]" inside the report structures, and call it out as a legal risk in the gap logs.

    SOURCE MATERIAL PAYLOAD:
    {st.session_state.uploaded_report_text[:50000]}

    USER GAZETTE OVERRIDES:
    {override_rules}

    Your response must be a single, raw JSON block matching this layout exactly without backticks:
    {{
        "gap_audit": {{
            "executive_summary": "Extensive corporate compliance synthesis.",
            "environmental_gaps": [["Metric Component", "Current Status", "Statutory Breach / Legal Action Justification"]],
            "social_gaps": [["Metric Component", "Current Status", "Statutory Breach / Legal Action Justification"]],
            "governance_gaps": [["Metric Component", "Current Status", "Statutory Breach / Legal Action Justification"]],
            "red_flags": ["Explicit enforcement risk warnings"],
            "recommendations": ["Direct compliance actions required"]
        }},
        "esg_report": {{
            "company_vision": "Strategic positioning text.",
            "env_table_data": [["SECP Technical Indicator Component", "Extracted Disclosed Corporate Value Matrix", "Global Framework Reference Mapping (IFRS S2)"]],
            "soc_table_data": [["SECP Technical Indicator Component", "Extracted Disclosed Corporate Value Matrix", "Global Framework Reference Mapping (IFRS S1 / ISO 20400)"]],
            "gov_table_data": [["SECP Technical Indicator Component", "Extracted Disclosed Corporate Value Matrix", "Global Framework Reference Mapping (SOE Act 2023)"]],
            "env_narrative": "Detailed text analysis of carbon footprint, transition steps, and physical asset protections.",
            "soc_narrative": "Detailed text analysis of pay equity ratios, human rights protection, and vendor oversight protocols.",
            "gov_narrative": "Detailed text analysis of committee structural splits, internal audit verification matrices, and public disclosure loops."
        }}
    }}

    REQUIRED COMPONENT ROW ITEMS:
    Ensure you write out custom row items in the tables matching all of these components:
    - Environment: GHG Emissions, Emissions Intensity, Energy Usage, Energy Intensity, Energy Mix, Water Ingestion, Environmental Operations, Environmental Board Oversight, Climate Risk Mitigation Plans.
    - Social: CEO Pay Ratio (CEO to median employee), Gender Pay Ratio, Turnover Rates, Gender Diversity Splits, Temporary Worker Ratio, Non-Discrimination Protections, Workplace Injury Rates, Child & Forced Labor, Sustainable Procurement (ISO 20400).
    - Governance: Board Diversity, Board Independence, Incentivized Performance Pay, Collective Bargaining Controls, Supplier Code of Conduct, Ethics & Anti-Corruption, Sustainability Public Reporting, External Assurance Third-Party Audits.
    """
    
    with st.spinner("Processing complex multi-standard regulatory data streams..."):
        try:
            client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=orchestration_prompt,
                config={'response_mime_type': 'application/json'}
            )
            st.session_state.audit_data_cache = json.loads(response.text)
            st.session_state.pipeline_completed = True
        except Exception as err:
            st.error(f"Pipeline Execution Failure: {err}")

# ==============================================================================
# VIEW RENDERING TAB 2: EXECS GAP MATRIX
# ==============================================================================
with tab2:
    if not st.session_state.pipeline_completed:
        st.info("Awaiting file uploads in the Ingestion Workspace to launch the data pipeline.")
    else:
        audit_payload = st.session_state.audit_data_cache.get("gap_audit", {})
        
        # Word File Constructor
        audit_doc = Document()
        apply_executive_formatting(audit_doc, "Statutory ESG Gap Audit Advisory Report")
        
        add_executive_heading(audit_doc, "Table of Contents Index Outline", level=1)
        add_justified_paragraph(audit_doc, "1. Executive Compliance Summary & Statutory Justification Matrix.......................................... Page 1")
        add_justified_paragraph(audit_doc, "2. Environmental Performance Deficiencies Evaluation Table......................................................... Page 2")
        add_justified_paragraph(audit_doc, "3. Social Responsibility & Procurement Metric Mapping Table....................................................... Page 3")
        add_justified_paragraph(audit_doc, "4. Corporate Governance Control Gaps (SOE Act 2023 / Companies Act)........................................... Page 4")
        add_justified_paragraph(audit_doc, "5. Operational Red Flags, Action Items & Long-Term Recommendations............................................. Page 5")
        
        add_executive_heading(audit_doc, "1. Executive Compliance Summary & Statutory Justification Matrix", level=1)
        add_justified_paragraph(audit_doc, audit_payload.get("executive_summary", ""))
        
        add_executive_heading(audit_doc, "2. Environmental Performance Deficiencies Evaluation Table", level=1)
        add_executive_data_table(audit_doc, ["Pillar Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("environmental_gaps", []))
        
        add_executive_heading(audit_doc, "3. Social Responsibility & Procurement Metric Mapping Table", level=1)
        add_executive_data_table(audit_doc, ["Pillar Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("social_gaps", []))
        
        add_executive_heading(audit_doc, "4. Corporate Governance Control Gaps (SOE Act 2023 / Companies Act)", level=1)
        add_executive_data_table(audit_doc, ["Pillar Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("governance_gaps", []))
        
        add_executive_heading(audit_doc, "5. Operational Red Flags, Action Items & Long-Term Recommendations", level=1)
        for flag in audit_payload.get("red_flags", []):
            add_justified_paragraph(audit_doc, f"⚠️ WARNING: {flag}")
        for rec in audit_payload.get("recommendations", []):
            add_justified_paragraph(audit_doc, f"🎯 ACTION DIRECTIVE: {rec}")
            
        audit_buf = io.BytesIO()
        audit_doc.save(audit_buf)
        audit_buf.seek(0)
        
        # UI Presentation Components
        col_dl, _ = st.columns([1, 3])
        with col_dl:
            st.download_button("📥 Export Gap Audit (.DOCX)", data=audit_buf, file_name="Regulatory_ESG_Gap_Audit.docx")
            
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.subheader("Executive System Evaluation")
        st.write(audit_payload.get("executive_summary"))
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# VIEW RENDERING TAB 3: LONG FORMAL ESG DISCLOSURE SUITE
# ==============================================================================
with tab3:
    if not st.session_state.pipeline_completed:
        st.info("Awaiting file uploads in the Ingestion Workspace to launch the data pipeline.")
    else:
        esg_payload = st.session_state.audit_data_cache.get("esg_report", {})
        
        esg_doc = Document()
        apply_executive_formatting(esg_doc, "Corporate Sustainability & ESG Disclosure Statement")
        
        add_executive_heading(esg_doc, "Document Layout Index Directory", level=1)
        add_justified_paragraph(esg_doc, "1. Strategic Corporate Stance & Technological Vision Statements................................................ Page 1")
        add_justified_paragraph(esg_doc, "2. Environmental (E) Performance Metrics & Asset Data Grids (IFRS S2 / SECP).................................. Page 2")
        add_justified_paragraph(esg_doc, "3. Social (S) Responsibility Frameworks & Sustainable Sourcing Performance (IFRS S1 / ISO 20400).............. Page 3")
        add_justified_paragraph(esg_doc, "4. Corporate Governance (G) Control Metrics & Board Integrity Grids (SOE Act 2023)............................ Page 4")
        
        add_executive_heading(esg_doc, "1. Strategic Corporate Stance & Technological Vision Statements", level=1)
        add_justified_paragraph(esg_doc, esg_payload.get("company_vision", ""))
        
        add_executive_heading(esg_doc, "2. Environmental (E) Performance Metrics & Asset Data Grids (IFRS S2 / SECP)", level=1)
        add_executive_data_table(esg_doc, ["SECP Environmental Indicator", "Disclosed Value / Asset Standing", "Framework Reference"], esg_payload.get("env_table_data", []), hex_header_color="1E3E62")
        add_executive_heading(esg_doc, "Exhaustive Environmental & Decarbonization Narrative Analysis", level=2)
        add_justified_paragraph(esg_doc, esg_payload.get("env_narrative", ""))
        
        add_executive_heading(esg_doc, "3. Social (S) Responsibility Frameworks & Sustainable Sourcing Performance (IFRS S1 / ISO 20400)", level=1)
        add_executive_data_table(esg_doc, ["SECP Social Indicator", "Disclosed Value / Asset Standing", "Framework Reference"], esg_payload.get("soc_table_data", []), hex_header_color="1E3E62")
        add_executive_heading(esg_doc, "Exhaustive Social Impact & Supply Chain Sustainability Narrative Analysis", level=2)
        add_justified_paragraph(esg_doc, esg_payload.get("soc_narrative", ""))
        
        add_executive_heading(esg_doc, "4. Corporate Governance (G) Control Metrics & Board Integrity Grids (SOE Act 2023)", level=1)
        add_executive_data_table(esg_doc, ["SECP Governance Indicator", "Disclosed Value / Asset Standing", "Framework Reference"], esg_payload.get("gov_table_data", []), hex_header_color="1E3E62")
        add_executive_heading(esg_doc, "Exhaustive Governance Control, Internal Risk Management, and Assurance Analysis", level=2)
        add_justified_paragraph(esg_doc, esg_payload.get("gov_narrative", ""))
        
        esg_buf = io.BytesIO()
        esg_doc.save(esg_buf)
        esg_buf.seek(0)
        
        col_dl2, _ = st.columns([1, 3])
        with col_dl2:
            st.download_button("📥 Export Formal ESG Disclosures (.DOCX)", data=esg_buf, file_name="Formal_Corporate_ESG_Disclosures.docx")
            
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.subheader("Strategic Organizational Vision & National Alignment Stance")
        st.write(esg_payload.get("company_vision"))
        st.markdown('</div>', unsafe_allow_html=True)

        # On-Screen Interactive UX Grids for Review Team
        st.markdown("### Interactive Dashboard Review Panel")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.info("☘️ Environmental Disclosures Mapped")
            for row in esg_payload.get("env_table_data", [])[1:]:
                if "[OMITTED]" in str(row[1]).upper():
                    st.markdown(f"**{row[0]}:** <span class='omitted-badge'>Omitted Data Gap</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"**{row[0]}:** `{row[1]}`", unsafe_allow_html=True)
        with c2:
            st.info("🤝 Social Disclosures Mapped")
            for row in esg_payload.get("soc_table_data", [])[1:]:
                if "[OMITTED]" in str(row[1]).upper():
                    st.markdown(f"**{row[0]}:** <span class='omitted-badge'>Omitted Data Gap</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"**{row[0]}:** `{row[1]}`", unsafe_allow_html=True)
        with c3:
            st.info("🏛️ Governance Indicators Mapped")
            for row in esg_payload.get("gov_table_data", [])[1:]:
                if "[OMITTED]" in str(row[1]).upper():
                    st.markdown(f"**{row[0]}:** <span class='omitted-badge'>Omitted Data Gap</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"**{row[0]}:** `{row[1]}`", unsafe_allow_html=True)
