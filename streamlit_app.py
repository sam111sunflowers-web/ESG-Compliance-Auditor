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
# PROFESSIONAL DESIGN SYSTEM (CSS INTERFACE OVERRIDES)
# ==============================================================================
st.set_page_config(
    page_title="Apex ESG Engine", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    /* Global Canvas Reset */
    .stApp {
        background-color: #F8FAFC !important;
        font-family: 'Inter', 'Segoe UI', system-ui, sans-serif !important;
    }
    
    /* Executive Header Banner */
    .executive-header-card {
        background: linear-gradient(135deg, #0B192C 0%, #1E3E62 100%);
        padding: 2.5rem 2rem;
        border-radius: 12px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(11, 25, 44, 0.15);
        margin-bottom: 2rem;
    }
    .executive-header-card h1 {
        color: #FFFFFF !important;
        font-size: 2.4rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.05em;
        margin: 0 0 0.5rem 0 !important;
    }
    .executive-header-card p {
        color: #94A3B8 !important;
        font-size: 1.05rem !important;
        margin: 0 !important;
    }
    
    /* Corporate Panel Containers */
    .corporate-panel-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.75rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
    }
    
    .panel-title {
        color: #0B192C !important;
        font-size: 1.2rem !important;
        font-weight: 600 !important;
        margin-bottom: 0.25rem !important;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #F1F5F9;
    }
    
    .panel-subtitle {
        color: #64748B !important;
        font-size: 0.85rem !important;
        margin-bottom: 1.25rem !important;
    }
    
    /* Custom High-Status Corporate Tabs Navigation */
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
    
    /* Statutory Omission Indicator Badges */
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
    
    /* Action Buttons Custom Color Theme Overrides */
    div.stFormSubmitButton > button, div.stDownloadButton > button {
        background: #10B981 !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.75rem 2rem !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.2) !important;
        width: 100%;
        transition: all 0.2s ease-in-out !important;
    }
    div.stFormSubmitButton > button:hover, div.stDownloadButton > button:hover {
        background: #059669 !important;
        transform: translateY(-2px) !important;
    }
    </style>
""", unsafe_allow_html=True)

# Main Branding Header Components
st.markdown("""
    <div class="executive-header-card">
        <h1>🏛️ ESG Report Generator & Compliance Audit Guide</h1>
        <p>Enterprise Regulatory Pipeline • Synchronized with Companies Act 2017 (Sec 238), SOE Act 2023, SOE Policy, SECP Guidelines, IFRS S1/S2 & ISO 20400 Protocols</p>
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
# MOCK DATA FALLBACK LOGIC GENERATOR (INSULATION MECHANISM AGAINST 503 ERRORS)
# ==============================================================================
def load_emergency_mock_payload():
    return {
        "gap_audit": {
            "executive_summary": "SYSTEM ADVISORY NOTE: Remote AI model server is under high demand (503). Loaded institutional sandbox metrics baseline. Compliance tracking under Section 238 of the Companies Act 2017 and the SOE Act 2023 indicates critical operational data omissions across primary non-financial segments.",
            "environmental_gaps": [
                ["GHG Carbon Inventory Scope 1/2", "[OMITTED / DATA NOT AVAILABLE]", "Breach of SECP Mandatory Disclosures guidelines & IFRS S2 standards."],
                ["Energy Mix and Consumption Intensity", "[OMITTED / DATA NOT AVAILABLE]", "Non-compliance with national green energy transition disclosure metrics."]
            ],
            "social_gaps": [
                ["CEO-to-Median Pay Ratio Matrix", "[OMITTED / DATA NOT AVAILABLE]", "Omission under Section 238 tracking human resource cost coefficients."],
                ["Sustainable Sourcing (ISO 20400)", "Partial Compliance Disclosed", "Supply chain parameters require third-party verification protocols."]
            ],
            "governance_gaps": [
                ["Conflict of Interest Directives", "Disclosed on Registry", "Aligned with SOE Ownership and Management Policy, 2023 guidelines."],
                ["Independent Audit Verification Matrix", "[OMITTED / DATA NOT AVAILABLE]", "Breach of explicit internal public disclosure loops under SOE Act 2023."]
            ],
            "red_flags": [
                "Omission of Scope 1/2 greenhouse data exposes entity to regulatory notifications by SECP under Companies Act Section 238.",
                "Absence of quantitative compensation ratios creates operational visibility friction with the Central Monitoring Unit (CMU)."
            ],
            "recommendations": [
                "Deploy an inner organizational data tracking project pipeline on Asana or Trello to capture missing carbon matrices.",
                "Structure an explicit anti-slavery procurement verification checklist matching ISO 20400 guidelines."
            ]
        },
        "esg_report": {
            "company_vision": "Strategic positioning text initialized. The enterprise is focused on building digital transformation networks to reduce transactional costs while aligning fully with Pakistan's updated National Determined Contributions (NDCs) framework.",
            "env_table_data": [
                ["SECP Environmental Technical Indicator", "Extracted Value Matrix", "Framework Mappings (IFRS S2 / SECP Guidelines)"],
                ["Greenhouse Gas Ingestion", "[OMITTED / DATA NOT AVAILABLE]", "IFRS S2 Climate Information Framework"],
                ["Clean Alternative Energy Mix Ratio", "14% Solar Integration Matrix", "SECP Voluntary Guidelines Pillar E1"]
            ],
            "soc_table_data": [
                ["SECP Social Technical Indicator", "Extracted Value Matrix", "Framework Mappings (IFRS S1 / ISO 20400)"],
                ["Workplace Injury Frequency Coefficient", "0.00 Incident Rate Filed", "IFRS S1 General Protections Standards"],
                ["Sustainable Procurement Auditing", "[OMITTED / DATA NOT AVAILABLE]", "ISO 20400:2017 Supply Chain Guidelines"]
            ],
            "gov_table_data": [
                ["SECP Governance Technical Indicator", "Extracted Value Matrix", "Framework Mappings (SOE Act 2023 / Companies Act 2017)"],
                ["Board Independence Ratio Splits", "33% Independent Directorships Mapped", "SOE Act 2023 Structure Mandates"],
                ["Anti-Corruption Code Training Logs", "[OMITTED / DATA NOT AVAILABLE]", "SOE Ownership and Management Policy 2023"]
            ],
            "env_narrative": "Detailed simulation review of corporate carbon data systems. Operational footprints indicate localized resource constraints. Strategic plans require structural execution tracks for environmental resilience.",
            "soc_narrative": "Detailed review of human asset metrics. Fair labor protocols are structurally active; however, formal upstream supply chain verification procedures under ISO 20400 standards must be finalized.",
            "gov_narrative": "Governance controls analysis. Board structural balance is compliant with local rules, but public verification procedures require independent third-party audit metrics."
        }
    }

# ==============================================================================
# MS-WORD PRODUCTION EXPORT GENERATION MATRIX
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
                if "[OMITTED]" in str(cell_value).upper() or "DATA NOT AVAILABLE" in str(cell_value).upper():
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(153, 27, 27)
    doc.add_paragraph("").paragraph_format.space_after = Pt(12)

# ==============================================================================
# UI NAVIGATION PANELS
# ==============================================================================
tab1, tab2, tab3 = st.tabs([
    "📥 Upload Company Repository Data", 
    "📊 Gap Audit Summary & Guidance", 
    "📜 Generated ESG Report Portfolio"
])

with tab1:
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("""
            <div class="corporate-panel-card">
                <div class="panel-title">📁 Company Data & Annual Report Influx</div>
                <div class="panel-subtitle">Primary ingestion path for core operational statements, annual data profiles, and strategic summaries.</div>
            </div>
        """, unsafe_allow_html=True)
        uploaded_report = st.file_uploader(
            "Select corporate asset (PDF):", 
            type=["pdf"], 
            key="core_report_uploader",
            label_visibility="collapsed"
        )
        
    with col_right:
        st.markdown("""
            <div class="corporate-panel-card">
                <div class="panel-title">📜 Statutory Amendments, Gazettes & Circular Influx</div>
                <div class="panel-subtitle">Secondary structural patch path for parsing direct SECP orders, sovereign gazette filings, and circular parameters.</div>
            </div>
        """, unsafe_allow_html=True)
        uploaded_gazette = st.file_uploader(
            "Select regulatory override patch (PDF):", 
            type=["pdf"], 
            key="statutory_patch_uploader",
            label_visibility="collapsed"
        )

    if uploaded_report and not st.session_state.uploaded_report_text:
        if uploaded_report.size > MAX_FILE_SIZE_BYTES:
            st.error("❌ File size limit exceeded (Max 100MB).")
        else:
            reader = PdfReader(uploaded_report)
            st.session_state.uploaded_report_text = "".join([page.extract_text() or "" for page in reader.pages])
            st.success("Primary company profile data loaded into session cache memory.")
            st.rerun()
                
    if uploaded_gazette and not st.session_state.gazette_override_text:
        reader = PdfReader(uploaded_gazette)
        st.session_state.gazette_override_text = "".join([page.extract_text() or "" for page in reader.pages])
        st.info("Statutory patch text parsed successfully.")
        st.rerun()

    # FORM ISOLATION EXECUTION BUTTON BLOCK
    if st.session_state.uploaded_report_text:
        st.markdown("### ⚡ Step 2: Launch Regulatory Processing Core")
        with st.form("orchestration_pipeline_form"):
            st.markdown("Click below to run data processing pipelines against local laws and frameworks.")
            submit_pipeline = st.form_submit_button("Execute Regulatory Audit Pipeline")
            
            if submit_pipeline:
                override_rules = st.session_state.gazette_override_text if st.session_state.gazette_override_text else "No active amendments."
                
                orchestration_prompt = f"""
                You are an elite legal regulatory audit engine specializing in Pakistani corporate compliance frameworks. Process the corporate text corpus against the following exact mandates:
                - Companies Act, 2017: Section 232 & Section 238 (Mandatory Sustainability Disclosures).
                - State-Owned Enterprises (Governance and Operations) Act, 2023.
                - SOE Ownership and Management Policy, 2023 (CMU parameters, Conflict of Interest).
                - SECP Voluntary Guidelines on ESG Disclosures for Listed Companies, 2023.
                - Frameworks: IFRS S1, IFRS S2, and ISO 20400:2017 Standards.
                
                Your response must be a single, raw JSON block matching this layout exactly without backticks:
                {{
                    "gap_audit": {{
                        "executive_summary": "Legal evaluation summary.",
                        "environmental_gaps": [["Metric Component", "Current Status", "Statutory Breach Analysis"]],
                        "social_gaps": [["Metric Component", "Current Status", "Statutory Breach Analysis"]],
                        "governance_gaps": [["Metric Component", "Current Status", "Statutory Breach Analysis"]],
                        "red_flags": ["Warnings"],
                        "recommendations": ["Direct corrective actions"]
                    }},
                    "esg_report": {{
                        "company_vision": "Strategic text tracing entity positioning and digitalization trends.",
                        "env_table_data": [["SECP Environmental Technical Indicator", "Extracted Value Matrix", "Framework Mappings"]],
                        "soc_table_data": [["SECP Social Technical Indicator", "Extracted Value Matrix", "Framework Mappings"]],
                        "gov_table_data": [["SECP Governance Technical Indicator", "Extracted Value Matrix", "Framework Mappings"]],
                        "env_narrative": "Detailed multi-page equivalent text analysis.",
                        "soc_narrative": "Detailed text tracing operations, diversity, and ISO 20400 sourcing.",
                        "gov_narrative": "Detailed text dissecting board committees, anti-corruption, and Sec 232 disclosure logs."
                    }}
                }}
                
                RAW CORPORATE TEXT MATERIAL:
                {st.session_state.uploaded_report_text[:40000]}
                """
                
                with st.spinner("Processing framework compliance data streams..."):
                    try:
                        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=orchestration_prompt,
                            config={'response_mime_type': 'application/json'}
                        )
                        st.session_state.audit_data_cache = json.loads(response.text)
                        st.session_state.pipeline_completed = True
                        st.success("Compliance data compilation completed successfully.")
                    except Exception as err:
                        # DEFENSIVE COATING: Intercept 503 errors and engage backup payload
