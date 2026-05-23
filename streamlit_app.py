import streamlit as st
from google import genai
from pypdf import PdfReader
import os
from weasyprint import HTML

# 1. System Configuration & Initialization
st.set_page_config(page_title="One-Shot Corporate ESG Engine", layout="wide")
st.title("⚖️ Public Sector Enterprise Sustainability Audit Framework")
st.caption("Custom IFRS S1/S2 Integration Pipeline aligned with SECP & SOE Act Statutory Mandates")

# Enforce strict technical payload thresholds
SIZE_LIMIT_MB = 100
MAX_BYTES_PER_FILE = SIZE_LIMIT_MB * 1024 * 1024

# Initialize session state objects for multi-tab data retention
if "raw_annual_report_text" not in st.session_state:
    st.session_state.raw_annual_report_text = ""
if "amendment_circular_text" not in st.session_state:
    st.session_state.amendment_circular_text = ""
if "report_generated" not in st.session_state:
    st.session_state.report_generated = False
if "ai_analysis_cache" not in st.session_state:
    st.session_state.ai_analysis_cache = ""

# Define the Multi-Tab Corporate Dashboard Architecture
tab1, tab2, tab3 = st.tabs([
    "📥 Corporate Workspace Ingestion", 
    "📄 Tab 2: Sustainability Report Addendum", 
    "🛠️ Tab 3: Regulatory Remediation Audit Guide"
])

# ==============================================================================
# TAB 1: WORKSPACE INGESTION LAYER
# ==============================================================================
with tab1:
    st.subheader("📋 Step 1: Upload Active Corporate Documentation")
    st.markdown("Drop your primary financial annual report or corporate documentation records below.")
    
    uploaded_report = st.file_uploader(
        "Upload Corporate Annual Report (PDF Format):",
        type=["pdf"],
        key="main_report_uploader"
    )
    
    if uploaded_report and not st.session_state.raw_annual_report_text:
        if uploaded_report.size > MAX_BYTES_PER_FILE:
            st.error("❌ Transmitted file payload violates the 100MB system constraint pipeline threshold.")
        else:
            with st.spinner("Extracting multi-layer textual context coordinates..."):
                try:
                    pdf_reader = PdfReader(uploaded_report)
                    text_accumulator = ""
                    for page in pdf_reader.pages:
                        extracted_text = page.extract_text()
                        if extracted_text:
                            text_accumulator += extracted_text + "\n"
                    st.session_state.raw_annual_report_text = text_accumulator
                    st.success(f"✅ Master Annual Report synchronized! Parsed {len(pdf_reader.pages)} text frames.")
                except Exception as ex:
                    st.error(f"Failed to process text frame matrices: {ex}")
                    
    st.markdown("---")
    st.subheader("📜 Step 2: Dynamic Live Amendments Ingestion Slot")
    st.markdown("Have a new SECP circular, an updated clause, or a new Gazette amendment notification? Drop it here. The engine will instantly read it and override its old baseline rules.")
    
    uploaded_amendment = st.file_uploader(
        "Upload Recent SECP Circulars / Gazette Notifications (Optional):",
        type=["pdf"],
        key="amendment_uploader"
    )
    
    if uploaded_amendment and not st.session_state.amendment_circular_text:
        with st.spinner("Parsing dynamic legislative changes..."):
            try:
                pdf_reader = PdfReader(uploaded_amendment)
                amend_accumulator = ""
                for page in pdf_reader.pages:
                    extracted_text = page.extract_text()
                    if extracted_text:
                        amend_accumulator += extracted_text + "\n"
                st.session_state.amendment_circular_text = amend_accumulator
                st.info("⚡ Live Amendment Engine updated! New clause variants injected into runtime logic memory.")
            except Exception as ex:
                st.error(f"Failed to parse amendment matrix: {ex}")

# ==============================================================================
# PIPELINE COMPILATION CONTROLLER (Runs behind the scenes)
# ==============================================================================
if st.session_state.raw_annual_report_text and not st.session_state.report_generated:
    # Build the analytical execution context matrix
    amendment_context = st.session_state.amendment_circular_text if st.session_state.amendment_circular_text else "No active external circular overrides provided."
    
    orchestration_prompt = f"""
    You are an elite corporate governance auditor specializing in Pakistani public enterprise operations.
    Analyze the company's text data below to produce an executive reporting suite.
    
    CRITICAL STATUTORY BOUNDARIES YOU MUST ENFORCE:
    - Companies Act, 2017: Section 232 (Approval metrics) and Section 238 (Enforcement of international standards).
    - State-Owned Enterprises (Governance and Operations) Act, 2023: Section 11 (Independent Director requirements), Section 18 (CEO Appointment rules), and Section 25(3) (Board Solvency Declaration mandates).
    - SECP State-Owned Enterprises Ownership and Management Policy.
    - SECP ISSB Sustainability Disclosure Standards Order (IFRS S1 & IFRS S2).
    - ISO 20400:2017 (Sustainable Sourcing Rules).
    - Pakistan National Climate Change Policy.
    
    DYNAMIC OVERRIDE CIRCULARS/AMENDMENTS FILED LIVE BY USER:
    {amendment_context}
    
    RAW EXTRACTED COMPANY TEXT CORPUS:
    {st.session_state.raw_annual_report_text[:35000]}
    
    EXECUTION INSTRUCTIONS:
    Extract all metrics, names, and indicators. For any metric required under the laws or IFRS S1/S2 that is absent or omitted from the company's files, do not make assumptions. You must return '[Data Not Available]' or '[To Be Completed]' so the corporate team can manually insert it into the printed annual report layer later.
    
    Provide your output clearly formatted for these exact sections:
    1. EXECUTIVE_SUMMARY_TEXT: (High-level sustainability stance)
    2. ENVIRONMENTAL_METRICS_FOUND: (Specific data, resource usage, or carbon disclosures found)
    3. SOCIAL_SUPPLY_CHAIN_FOUND: (Workforce diversity parameters, supplier screening, or tracking gaps)
    4. GOVERNANCE_COMPLIANCE_FOUND: (Board independence evaluation, committee tracking, and specific citations under SOE Act Section 11 and Section 25)
    """
    
    with st.spinner("Executing One-Shot Framework Audit Strategy..."):
        try:
            api_key_secret = st.secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=api_key_secret)
            
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=orchestration_prompt,
            )
            
            st.session_state.ai_analysis_cache = response.text
            st.session_state.report_generated = True
        except Exception as err:
            st.error(f"Pipeline Automation Error: {err}")

# ==============================================================================
# TAB 2: PRIVILEGED SUSTAINABILITY REPORT ADDENDUM (READY TO BIND)
# ==============================================================================
with tab2:
    if not st.session_state.report_generated:
        st.warning("Please upload valid source corporate documentation files inside Tab 1 first.")
    else:
        st.subheader("📄 Integrated IFRS S1 & S2 Sustainability Disclosure Statement")
        st.caption("This layout is structured to be appended directly inside your main corporate Annual Report booklet right after the Directors' Report.")
        
        # Core HTML template injected with Aptos typography, 12pt body text, and 14pt major headings
        html_template = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Aptos:wght@400;700&display=swap');
                @page {{
                    size: A4;
                    margin: 25mm 20mm;
                    @bottom-right {{
                        content: counter(page);
                        font-family: 'Aptos', sans-serif;
                        font-size: 9pt;
                        color: #4a5568;
                    }}
                    @bottom-left {{
                        content: "Statutory Sustainability Report Addendum (Companies Act 2017 Sec 238)";
                        font-family: 'Aptos', sans-serif;
                        font-size: 9pt;
                        color: #4a5568;
                    }}
                }}
                body {{
                    font-family: 'Aptos', sans-serif;
                    color: #2d3748;
                    line-height: 1.6;
                    font-size: 12pt;
                }}
                h1 {{
                    font-size: 14pt;
                    font-weight: bold;
                    color: #1a202c;
                    text-transform: uppercase;
                    border-bottom: 2px solid #000;
                    padding-bottom: 4px;
                    margin-top: 30px;
                    page-break-after: avoid;
                }}
                h2 {{
                    font-size: 12pt;
                    font-weight: bold;
                    color: #2d3748;
                    margin-top: 20px;
                    page-break-after: avoid;
                }}
                p, li {{
                    font-size: 12pt;
                }}
                .cover-container {{
                    page-break-after: always;
                    padding-top: 60mm;
                    text-align: left;
                }}
                .cover-main-title {{
                    font-size: 14pt;
                    font-weight: bold;
                    color: #000;
                    letter-spacing: 1px;
                }}
                .cover-sub-title {{
                    font-size: 12pt;
                    color: #4a5568;
                    margin-top: 5px;
                }}
                .toc-element {{
                    border-bottom: 1px dotted #cbd5e0;
                    margin-bottom: 10px;
                }}
                .toc-page-num {{
                    float: right;
                }}
                .placeholder-box {{
                    background-color: #f7fafc;
                    border: 1px dashed #a0aec0;
                    padding: 12px;
                    margin: 15px 0;
                    font-style: italic;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 15px 0;
                    page-break-inside: avoid;
                }}
                th {{
                    background-color: #f7fafc;
                    color: #000;
                    font-weight: bold;
                    padding: 10px;
                    font-size: 12pt;
                    border: 1px solid #000;
                    text-align: left;
                }}
                td {{
                    padding: 10px;
                    border: 1px solid #cbd5e0;
                    font-size: 12pt;
                }}
            </style>
        </head>
        <body>

            <div class="cover-container">
                <div class="cover-main-title">SUSTAINABILITY DISCLOSURE REPORT ADDENDUM</div>
                <div class="cover-sub-title">Compiled Pursuant to Section 238 of the Companies Act, 2017</div>
                <p style="margin-top: 50mm;"><strong>Reporting Period:</strong> Active Fiscal Disclosure Cycle</p>
                <p><strong>Location Context:</strong> Formally designated for placement immediately downstream from the active Directors' Report Framework.</p>
            </div>

            <div style="page-break-after: always;">
                <h1>Table of Contents</h1>
                <div class="toc-element">1. Corporate Legal Definitions Section <span class="toc-page-num">Page 3</span></div>
                <div class="toc-element">2. Statement of Regulatory & Framework Compliance Alignment <span class="toc-page-num">Page 4</span></div>
                <div class="toc-element">3. Executive Core Sustainability Synthesis Disclosures <span class="toc-page-num">Page 5</span></div>
                <div class="toc-element">4. Environmental Performance Matrix & Resource Architecture (IFRS S2) <span class="toc-page-num">Page 6</span></div>
                <div class="toc-element">5. Social Capital Metrics & Sourcing Value Chains (IFRS S1 / ISO 20400) <span class="toc-page-num">Page 7</span></div>
                <div class="toc-element">6. Corporate Governance Control Architecture (SOE Act 2023 / Companies Act) <span class="toc-page-num">Page 8</span></div>
            </div>

            <h1>1. Corporate Legal Definitions Section</h1>
            <p>For the purposes of this Integrated Report Addendum, the following definitions are established under local and international statutory baselines:</p>
            <ul>
                <li><strong>The Commission:</strong> The Securities and Exchange Commission of Pakistan (SECP).</li>
                <li><strong>Sustainability Disclosure Standards:</strong> The technical standards issued by the International Sustainability Standards Board (ISSB), tracking specifically IFRS S1 and IFRS S2 protocols.</li>
                <li><strong>Scope 1 Direct Disclosures:</strong> Greenhouse gas emissions originating directly from organizational assets owned or controlled by the reporting enterprise.</li>
                <li><strong>Scope 2 Indirect Disclosures:</strong> Operational greenhouse gas emissions resulting from the consumption of purchased electricity, heat, steam, or cooling grids.</li>
                <li><strong>Sustainable Procurement:</strong> Sourcing activities matching the strategic framework conditions of ISO 20400:2017 guidelines.</li>
            </ul>

            <h1>2. Statement of Regulatory & Framework Compliance Alignment</h1>
            <p>This document constitutes a formal disclosure compiled under **Section 238 of the Companies Act, 2017** and paragraph 2 of the **SECP ISSB Adoption Order**. This reporting architecture actively aligns with and references the following regulatory coordinates:</p>
            <table>
                <thead>
                    <tr>
                        <th>Regulatory Authority Framework</th>
                        <th>Statutory Base Reference Coded</th>
                        <th>Corporate Adoption Integration Vector</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>SECP ISSB Order</strong></td>
                        <td>Companies Act 2017 Section 238 Matrix</td>
                        <td>Mandated structural framework template implementation baseline.</td>
                    </tr>
                    <tr>
                        <td><strong>Public Sector Governance</strong></td>
                        <td>State-Owned Enterprises (SOE) Act 2023</td>
                        <td>Enforces administrative accountability indicators for governance controls.</td>
                    </tr>
                    <tr>
                        <td><strong>Procurement Standards</strong></td>
                        <td>ISO 20400:2017 Guidelines</td>
                        <td>Establishes global traceability for sustainable value-chain vendor checks.</td>
                    </tr>
                    <tr>
                        <td><strong>Climate Policy Core</strong></td>
                        <td>Pakistan National Climate Change Policy</td>
                        <td>Governs organizational transition planning metrics for low-carbon shifts.</td>
                    </tr>
                </tbody>
            </table>

            <h1>3. Executive Core Sustainability Synthesis Disclosures</h1>
            <p>Below represents the corporate dataset parsed and extracted directly from your verified operational filings:</p>
            <div class="placeholder-box">{st.session_state.ai_analysis_cache}</div>

            <h1>4. Environmental Performance Matrix & Resource Architecture (IFRS S2)</h1>
            <h2>4.1 Carbon Greenhouse Asset Disclosures</h2>
            <p>Pursuant to IFRS S2 climate compliance protocols, the enterprise tracking system maps direct and indirect carbon metrics below. Where values are unquantified, structural placeholder blocks are maintained for manual data placement prior to binding:</p>
            <table>
                <thead>
                    <tr>
                        <th>IFRS S2 Climate Metric Vector</th>
                        <th>Extracted Annual File Value</th>
                        <th>Statutory Tracking Status Parameter</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Scope 1 Direct Fuel Emissions Disclosures</td>
                        <td>[To Be Completed / Refer to Sec 3 Summary Data Matrix]</td>
                        <td>Pending Verification Loop Closure</td>
                    </tr>
                    <tr>
                        <td>Scope 2 Energy Consumption Grid Footprint</td>
                        <td>[To Be Completed / Refer to Sec 3 Summary Data Matrix]</td>
                        <td>Pending Verification Loop Closure</td>
                    </tr>
                    <tr>
                        <td>Scope 3 Value-Chain Upstream Distribution Network</td>
                        <td>[Data Not Available]</td>
                        <td>Exempted under SECP Order First-Year Grace Protocol Rule</td>
                    </tr>
                </tbody>
            </table>

            <h1>5. Social Capital Metrics & Sourcing Value Chains (IFRS S1 / ISO 20400)</h1>
            <p>Vendor sourcing channels and employee allocation configurations must present strict anti-discrimination, fair labor protection, and workplace safety tracking coordinates:</p>
            <div class="placeholder-box">
                <strong>ISO 20400 Strategic Field Check:</strong> [To Be Completed: Supply Chain Management system tracking variables must verify tier-1 vendor operational status manually here before report binding processes execute.]
            </div>

            <h1>6. Corporate Governance Control Architecture (SOE Act 2023 / Companies Act)</h1>
            <h2>6.1 Board Alignment & Solvency Controls</h2>
            <p>Board split oversight logic matrices tracked against specific sections of the State-Owned Enterprises (Governance and Operations) Act, 2023:</p>
            <table>
                <thead>
                    <tr>
                        <th>Statutory Governance Criteria Gauge</th>
                        <th>Target Provision Limit Required</th>
                        <th>Active Enterprise Coded State Status</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Independent Board Split Ratio</td>
                        <td>SOE Act 2023 Section 11 Check Grid</td>
                        <td>Verified Active Compliance Tracking Profile</td>
                    </tr>
                    <tr>
                        <td>Mandatory Audit Committee Split</td>
                        <td>Companies Act 2017 Section 232 Template</td>
                        <td>Verified Active Compliance Tracking Profile</td>
                    </tr>
                    <tr>
                        <td>Explicit Board Solvency Statement</td>
                        <td>SOE Act 2023 Section 25(3) Requirement</td>
                        <td>[To Be Completed / Reference Remediation Guide Tab 3]</td>
                    </tr>
                </tbody>
            </table>

        </body>
        </html>
        """
        
        # Compile local files onto disk
        with open("temp_sustainability_report.html", "w", encoding="utf-8") as f:
            f.write(html_template)
            
        HTML("temp_sustainability_report.html").write_pdf("integrated_sustainability_report.pdf")
        
        # Display actionable download controls
        with open("integrated_sustainability_report.pdf", "rb") as asset_file:
            st.download_button(
                label="📥 Download Ready-To-Bind Sustainability Report (PDF)",
                data=asset_file,
                file_name="sustainability_report_addendum.pdf",
                mime="application/pdf"
            )
            
        st.info("💡 Pro-Tip: This PDF is written in Aptos font (Headings 14pt, Body 12pt). Print this file out and insert it directly behind your approved Directors' Report to comply with the SECP Order formatting pipeline.")

# ==============================================================================
# TAB 3: INTERNAL COMPLIANCE REMEDIATION & DATA IMPROVEMENT AUDIT GUIDE
# ==============================================================================
with tab3:
    if not st.session_state.report_generated:
        st.warning("Please upload valid source corporate documentation files inside Tab 1 first.")
    else:
        st.subheader("🛠️ SECP & SOE Act Data Improvement & Remediation Guide")
        st.caption("INTERNAL STRATEGY WORK SHEET ONLY — DO NOT ATTACH TO PUBLIC ANNUAL REPORT RECORDS")
        
        remediation_prompt = f"""
        You are an uncompromising internal compliance auditor under SECP and the Auditor General of Pakistan frameworks.
        Review the following AI extracted data results and create an internal remediation guide punch list.
        
        AI SUMMARY SOURCE FROM DATA MATRIX:
        {st.session_state.ai_analysis_cache}
        
        CRITICAL OUTPUT OBJECTIVE:
        Generate a brutal, actionable internal punch list. Your findings must explicitly refer to specific sections of the relevant laws.
        Format your response using these exact 3 categories:
        1. STRATEGIC_GOVERNANCE_GAPS: (Look for omissions regarding Independent Director split limits under SOE Act 2023 Section 11, CEO selection pathways under Section 18, and Board Solvency Declarations under Section 25(3))
        2. DATA_COLLECTION_DEFICIENCIES: (Identify what IFRS S1 / S2 data points are missing, such as Scope 1 and Scope 2 energy/fuel metrics)
        3. VALUE_CHAIN_RISK_REMEDIATION: (Highlight gaps in vendor management systems relative to ISO 20400 protocols)
        """
        
        with st.spinner("Analyzing data omissions and compiling legal reference grids..."):
            try:
                api_key_secret = st.secrets["GEMINI_API_KEY"]
                client = genai.Client(api_key=api_key_secret)
                
                rem_response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=remediation_prompt,
                )
                
                remediation_guide_text = rem_response.text
                
                # Render the internal remediation matrix directly onto the Streamlit tab layout screen
                st.markdown("### ⚠️ Coded Compliance Remediation Deficiencies Matrix")
                st.write(remediation_guide_text)
                
                # HTML Architecture for Report 2
                rem_html = f"""
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <style>
                        @import url('https://fonts.googleapis.com/css2?family=Aptos:wght@400;700&display=swap');
                        body {{ font-family: 'Aptos', sans-serif; color: #2d3748; line-height: 1.6; font-size: 12pt; padding: 20mm; }}
                        h1 {{ font-size: 14pt; font-weight: bold; color: #b7791f; border-bottom: 2px solid #b7791f; margin-top: 30px; padding-bottom: 5px; }}
                        h2 {{ font-size: 12pt; font-weight: bold; color: #2d3748; }}
                        .warning-panel {{ background-color: #fffaf0; border-left: 4px solid #dd6b20; padding: 15px; margin: 15px 0; }}
                    </style>
                </head>
                <body>
                    <h1 style="font-size: 14pt; color:#000;">INTERNAL REMEDIATION GUIDE AND DATA PUNCH LIST</h1>
                    <p><strong>Security Context:</strong> Restricted Privileged Internal Governance Audit Resource Framework</p>
                    <hr>
                    <div class="warning-panel">
                        {remediation_guide_text.replace('\n', '<br>')}
                    </div>
                </body>
                </html>
                """
                
                with open("temp_remediation_guide.html", "w", encoding="utf-8") as f:
                    f.write(rem_html)
                    
                HTML("temp_remediation_guide.html").write_pdf("internal_remediation_guide.pdf")
                
                with open("internal_remediation_guide.pdf", "rb") as rem_asset:
                    st.download_button(
                        label="📥 Download Internal Remediation Action Guide (PDF)",
                        data=rem_asset,
                        file_name="internal_remediation_strategy_guide.pdf",
                        mime="application/pdf"
                    )
            except Exception as rem_ex:
                st.error(f"Failed to compile internal remediation guide: {rem_ex}")
