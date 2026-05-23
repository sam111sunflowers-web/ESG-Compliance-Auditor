import streamlit as st
import json
import os
from pypdf import PdfReader
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import io

# ==============================================================================
# SYSTEM INITIALIZATION & GUARDRAILS
# ==============================================================================
st.set_page_config(page_title="Enterprise ESG Data Pipeline", layout="wide")
st.title("🏛️ Corporate ESG Report & Statutory Gap Audit Engine")
st.caption("Aligned with SECP ESG Voluntary Guidelines, SOE Act 2023, SOE Policy, and Companies Act 2017")

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
# PROFESSIONAL OFFICE STYLING ENGINE (.DOCX)
# ==============================================================================
def apply_professional_formatting(doc, title_text):
    """Enforces strict corporate typography rules: 1.5 spacing, full justification, Aptos/Calibri styling."""
    # Set standard margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Style definitions
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'  # Universally compliant fallback for Aptos
    font.size = Pt(12)
    
    # Configure Header/Cover Section
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_run = title_p.add_run(title_text.upper())
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.name = 'Calibri'
    
    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run("Generated for Executive Review & Design Team Optimization")
    sub_run.font.size = Pt(10)
    sub_run.font.italic = True
    
    doc.add_paragraph("").paragraph_format.space_after = Pt(24)

def add_executive_heading(doc, text, level=1):
    """Creates clear, justified headings with structured padding."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(14)
        # Add bottom border styling element to Heading 1 via XML
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '4')
        bottom.set(qn('w:color'), '1A365D')
        pBdr.append(bottom)
        pPr.append(pBdr)
    else:
        run.font.size = Pt(12)

def add_justified_paragraph(doc, text, is_bold=False, is_italic=False, space_after=12):
    """Applies strict 1.5 line spacing, fully justified text blocks, and paragraph spacing."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(space_after)
    
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = is_bold
    run.font.italic = is_italic
    return p

def add_styled_table(doc, headers, rows):
    """Generates clean corporate tables with colored headers and alternating rows."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False
    
    # Header styling
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        # Apply dark background via XML shading
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1A365D"/>')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = None
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            
    # Data row styling
    for r_idx, row_data in enumerate(rows):
        row_cells = table.add_row().cells
        bg_color = "F7FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, cell_value in enumerate(row_data):
            row_cells[c_idx].text = str(cell_value)
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            row_cells[c_idx]._tc.get_or_add_tcPr().append(shading)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.15
            for run in p.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(10.5)
    
    doc.add_paragraph("").paragraph_format.space_after = Pt(12)

# ==============================================================================
# TAB DESIGN & INTERFACE LAYOUT
# ==============================================================================
tab1, tab2, tab3 = st.tabs([
    "📥 Ingestion Workspace", 
    "📊 Tab 2: ESG Statutory Gap Audit Report", 
    "📜 Tab 3: Formal Corporate ESG Report"
])

with tab1:
    st.subheader("📋 Document Ingestion & Storage Control Panel")
    st.markdown("Upload current company performance records along with dynamic legislative overrides.")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_report = st.file_uploader("Corporate Annual Report or Operating Data (PDF):", type=["pdf"])
    with col2:
        uploaded_gazette = st.file_uploader("Recent Gazette Amendments or SECP Circulars (Optional PDF):", type=["pdf"])

    if uploaded_report and not st.session_state.uploaded_report_text:
        if uploaded_report.size > MAX_FILE_SIZE_BYTES:
            st.error("❌ Corporate file size exceeds strict system 100MB guardrails.")
        else:
            with st.spinner("Extracting master document coordinates..."):
                reader = PdfReader(uploaded_report)
                extracted_text = "".join([page.extract_text() or "" for page in reader.pages])
                st.session_state.uploaded_report_text = extracted_text
                st.success("✅ Main corporate corpus compiled safely in storage.")
                
    if uploaded_gazette and not st.session_state.gazette_override_text:
        with st.spinner("Extracting legislative patch text..."):
            reader = PdfReader(uploaded_gazette)
            st.session_state.gazette_override_text = "".join([page.extract_text() or "" for page in reader.pages])
            st.info("⚡ Live dynamic regulatory adjustments integrated.")

# ==============================================================================
# SINGLE-PASS PIPELINE EXECUTION ENGINE
# ==============================================================================
if st.session_state.uploaded_report_text and not st.session_state.pipeline_completed:
    override_rules = st.session_state.gazette_override_text if st.session_state.gazette_override_text else "No user-specified live amendments provided."
    
    orchestration_prompt = f"""
    You are an expert enterprise auditor specialized in Pakistani corporate governance frameworks and SECP ESG Guidelines.
    Analyze the company data and return a JSON object with two specific keys: "gap_audit" and "esg_report".

    CRITICAL STATUTORY LAWS YOU MUST MAP:
    - SOE Policy 2023 & SOE Act 2023 (Oversight, competitive neutrality, board composition)
    - Companies Act, 2017 (Section 232 financial loops and Section 238 sustainability enforcement matrices)
    
    USER OVERRIDES / GAZETTE AMENDMENTS:
    {override_rules}

    SOURCE ENTERPRISE CORPUS MATERIAL:
    {st.session_state.uploaded_report_text[:40000]}

    EVALUATION INSTRUCTIONS:
    Evaluate across all 3 pillars:
    - Environment: GHG Emissions, Emissions Intensity, Energy Usage, Energy Intensity, Energy Mix, Water Usage, Environmental Operations, Environmental Oversight, Climate Risk Mitigation.
    - Social: CEO Pay Ratio, Gender Pay Ratio, Employee Turnover, Gender Diversity, Temporary Worker Ratio, Non-Discrimination, Injury Rate, Global Health & Safety, Child & Forced Labor, Human Rights, Marketing.
    - Governance: Board Diversity, Board Independence, Incentivized Pay, Collective Bargaining, Supplier Code of Conduct, Ethics & Anti-Corruption, Injury Rate, Sustainability Reporting, Disclosure Practices, External Assurance.

    If data points required by SECP or international frameworks (IFRS S1/S2) are not found in the text, mark them as "[Omitted / Data Not Available]" in the report section and list them as critical gaps in the audit framework.

    Your response must be a single, valid JSON object matching this structure exactly without markdown wrapping strings:
    {{
        "gap_audit": {{
            "executive_summary": "string text context",
            "environmental_gaps": [["Pillar", "Status / What is Done", "Missing Justification & Gap Analysis"]],
            "social_gaps": [["Pillar", "Status / What is Done", "Missing Justification & Gap Analysis"]],
            "governance_gaps": [["Pillar", "Status / What is Done", "Missing Justification & Gap Analysis"]],
            "red_flags": ["string alert"],
            "recommendations": ["string strategic action item"]
        }},
        "esg_report": {{
            "company_vision": "string",
            "environmental_disclosures": {{
                "ghg_metrics": "string",
                "energy_mix_intensity": "string",
                "water_and_operations": "string",
                "climate_risk_mitigation_s2": "string"
            }},
            "social_disclosures": {{
                "human_capital_and_pay_ratios": "string",
                "diversity_and_inclusion": "string",
                "labor_and_human_rights": "string"
            }},
            "governance_disclosures": {{
                "board_composition_and_independence": "string",
                "ethics_anti_corruption": "string",
                "assurance_practices": "string"
            }}
        }}
    }}
    """
    
    with st.spinner("Processing document data matrices via single-pass architecture..."):
        try:
            client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=orchestration_prompt,
                config={'response_mime_type': 'application/json'}
            )
            
            st.session_state.audit_data_cache = json.loads(response.text)
            st.session_state.pipeline_completed = True
        except Exception as e:
            st.error(f"Data Pipeline Interruption: {e}")

# ==============================================================================
# REPORT 1 OUT-STREAM: COMPLIANCE AUDIT & GAP SHEET
# ==============================================================================
with tab2:
    if not st.session_state.pipeline_completed:
        st.warning("Awaiting document transmission workspace ingestion pools inside Tab 1.")
    else:
        audit_payload = st.session_state.audit_data_cache.get("gap_audit", {})
        st.subheader("📊 Structured Compliance Framework Audit & Gap Analysis Sheet")
        
        # Build the physical Word Document natively in memory
        audit_doc = Document()
        apply_professional_formatting(audit_doc, "Statutory ESG Gap Audit Advisory Report")
        
        # Table of Contents Outline Block
        add_executive_heading(audit_doc, "Table of Contents", level=1)
        add_justified_paragraph(audit_doc, "1. Executive Compliance Summary & Statutory Justification Matrix.......................................... Page 1", is_italic=True, space_after=4)
        add_justified_paragraph(audit_doc, "2. Environmental Performance Deficiencies Evaluation Table......................................................... Page 2", is_italic=True, space_after=4)
        add_justified_paragraph(audit_doc, "3. Social Responsibility & Procurement Metric Mapping Table....................................................... Page 3", is_italic=True, space_after=4)
        add_justified_paragraph(audit_doc, "4. Corporate Governance Control Gaps (SOE Act 2023 / Companies Act)........................................... Page 4", is_italic=True, space_after=4)
        add_justified_paragraph(audit_doc, "5. Operational Red Flags, Action Items & Long-Term Recommendations............................................. Page 5", is_italic=True, space_after=18)
        
        # Section 1
        add_executive_heading(audit_doc, "1. Executive Compliance Summary & Statutory Justification Matrix", level=1)
        add_justified_paragraph(audit_doc, audit_payload.get("executive_summary", "Evaluation missing."))
        
        # Section 2
        add_executive_heading(audit_doc, "2. Environmental Performance Deficiencies Evaluation Table", level=1)
        add_styled_table(audit_doc, ["Pillar Indicator Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("environmental_gaps", []))
        
        # Section 3
        add_executive_heading(audit_doc, "3. Social Responsibility & Procurement Metric Mapping Table", level=1)
        add_styled_table(audit_doc, ["Pillar Indicator Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("social_gaps", []))
        
        # Section 4
        add_executive_heading(audit_doc, "4. Corporate Governance Control Gaps (SOE Act 2023 / Companies Act)", level=1)
        add_styled_table(audit_doc, ["Pillar Indicator Component", "Identified Active Status", "Missing Metric Justification & Gap Analysis"], audit_payload.get("governance_gaps", []))
        
        # Section 5
        add_executive_heading(audit_doc, "5. Operational Red Flags, Action Items & Long-Term Recommendations", level=1)
        add_executive_heading(audit_doc, "Critical Operational Red Flags", level=2)
        for flag in audit_payload.get("red_flags", []):
            add_justified_paragraph(audit_doc, f"⚠️ {flag}", space_after=6)
            
        add_executive_heading(audit_doc, "Targeted Remediation Action Steps", level=2)
        for rec in audit_payload.get("recommendations", []):
            add_justified_paragraph(audit_doc, f"🎯 {rec}", space_after=6)
            
        # Compile document buffer stream for live client download tracking
        audit_buffer = io.BytesIO()
        audit_doc.save(audit_buffer)
        audit_buffer.seek(0)
        
        st.success("🎉 Word Document Structure Built Successfully with 1.5 spacing and justified text blocks.")
        st.download_button(
            label="📥 Download Audit Gap Report (.DOCX)",
            data=audit_buffer,
            file_name="Statutory_ESG_Gap_Audit_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # On-screen preview panels
        st.markdown("### Executive Summary Preview")
        st.write(audit_payload.get("executive_summary"))

# ==============================================================================
# REPORT 2 OUT-STREAM: CORPORATE ESG REPORT DISCLOSURE ASSET
# ==============================================================================
with tab3:
    if not st.session_state.pipeline_completed:
        st.warning("Awaiting document transmission workspace ingestion pools inside Tab 1.")
    else:
        esg_payload = st.session_state.audit_data_cache.get("esg_report", {})
        st.subheader("📜 Formal Corporate Sustainability Disclosure Statement")
        
        report_doc = Document()
        apply_professional_formatting(report_doc, "Corporate Sustainability & ESG Disclosure Statement")
        
        # Table of Contents
        add_executive_heading(report_doc, "Table of Contents", level=1)
        add_justified_paragraph(report_doc, "1. Strategic Organizational Stance & Vision Statements......................................................... Page 1", is_italic=True, space_after=4)
        add_justified_paragraph(report_doc, "2. Environmental (E) Performance Metrics & Asset Architecture (IFRS S2).................................. Page 2", is_italic=True, space_after=4)
        add_justified_paragraph(report_doc, "3. Social (S) Responsibility Frameworks & Supply Chains (IFRS S1 / ISO 20400).............................. Page 3", is_italic=True, space_after=4)
        add_justified_paragraph(report_doc, "4. Corporate Governance (G) Control Metrics (SOE Act 2023 / Companies Act).............................. Page 4", is_italic=True, space_after=18)
        
        # Section 1
        add_executive_heading(report_doc, "1. Strategic Organizational Stance & Vision Statements", level=1)
        add_justified_paragraph(report_doc, esg_payload.get("company_vision", "Vision missing."))
        
        # Section 2
        add_executive_heading(report_doc, "2. Environmental (E) Performance Metrics & Asset Architecture (IFRS S2)", level=1)
        env_data = esg_payload.get("environmental_disclosures", {})
        add_executive_heading(report_doc, "Greenhouse Gas Tracking & Emissions Intensity Profiles", level=2)
        add_justified_paragraph(report_doc, env_data.get("ghg_metrics", ""))
        add_executive_heading(report_doc, "Energy Ingestion Mix, Resource Efficiency and Intensities", level=2)
        add_justified_paragraph(report_doc, env_data.get("energy_mix_intensity", ""))
        add_executive_heading(report_doc, "Water Resource Allocations & Environmental Infrastructure Assets", level=2)
        add_justified_paragraph(report_doc, env_data.get("water_and_operations", ""))
        add_executive_heading(report_doc, "IFRS S2 Climate Change Transition Risks & Mitigation Strategy", level=2)
        add_justified_paragraph(report_doc, env_data.get("climate_risk_mitigation_s2", ""))
        
        # Section 3
        add_executive_heading(report_doc, "3. Social (S) Responsibility Frameworks & Supply Chains (IFRS S1 / ISO 20400)", level=1)
        soc_data = esg_payload.get("social_disclosures", {})
        add_executive_heading(report_doc, "Human Capital Allocations, Executive Ratios, and Compensation Ratios", level=2)
        add_justified_paragraph(report_doc, soc_data.get("human_capital_and_pay_ratios", ""))
        add_executive_heading(report_doc, "Workplace Inclusion Parameters, Gender Splits, and Non-Discrimination Protections", level=2)
        add_justified_paragraph(report_doc, soc_data.get("diversity_and_inclusion", ""))
        add_executive_heading(report_doc, "Global Health & Safety Metrics, Vendor Mapping, and Anti-Slavery Controls (ISO 20400)", level=2)
        add_justified_paragraph(report_doc, soc_data.get("labor_and_human_rights", ""))
        
        # Section 4
        add_executive_heading(report_doc, "4. Corporate Governance (G) Control Metrics (SOE Act 2023 / Companies Act)", level=1)
        gov_data = esg_payload.get("governance_disclosures", {})
        add_executive_heading(report_doc, "Board Independence Limits, Committee Splitting Ratios, and Structural Diversity", level=2)
        add_justified_paragraph(report_doc, gov_data.get("board_composition_and_independence", ""))
        add_executive_heading(report_doc, "Code of Conduct Mandates, Whistleblowing, and Institutional Anti-Corruption Control Systems", level=2)
        add_justified_paragraph(report_doc, gov_data.get("ethics_anti_corruption", ""))
        add_executive_heading(report_doc, "Public Reporting Assurance Matrices, Disclosure Frequency, and Third-Party External Auditing", level=2)
        add_justified_paragraph(report_doc, gov_data.get("assurance_practices", ""))
        
        report_buffer = io.BytesIO()
        report_doc.save(report_buffer)
        report_buffer.seek(0)
        
        st.success("🎉 Word Document Core Framework Disclosures generated successfully.")
        st.download_button(
            label="📥 Download Formal Corporate ESG Report (.DOCX)",
            data=report_buffer,
            file_name="Formal_Corporate_ESG_Disclosure_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.markdown("### Organizational Vision Strategy Preview")
        st.write(esg_payload.get("company_vision"))
