import streamlit as st
from google import genai
from pypdf import PdfReader
import os
from weasyprint import HTML

# 1. Page Configuration Architecture
st.set_page_config(page_title="Batch ESG Auditor", layout="wide")
st.title("⚖️ Enterprise Multi-Report ESG Compliance Auditor")
st.caption("Automated Extraction & High-Speed Policy Risk Analysis Pipeline")

st.markdown("---")

# Strict operational boundaries (100MB per file, 1GB total batch)
SIZE_LIMIT_MB = 100
TOTAL_BATCH_LIMIT_MB = 1024

MAX_BYTES_PER_FILE = SIZE_LIMIT_MB * 1024 * 1024
MAX_TOTAL_BYTES = TOTAL_BATCH_LIMIT_MB * 1024 * 1024

# 2. Folder Upload Interface Simulation (The "reports" Input Block)
st.subheader("📂 Step 1: Input Corporate Documentation Drop zone")
st.markdown("Drop multiple annual files or sustainability records here. System constraint: **Max 100MB per PDF**.")

uploaded_files = st.file_uploader(
    "Upload corporate documents into the active processing queue:",
    type=["pdf"],
    accept_multiple_files=True
)

# 3. Dynamic Text Layer Extraction Matrix
all_extracted_corporate_text = ""
valid_files_processed = []
running_total_bytes = 0

if uploaded_files:
    # Calculate total size of the batch
    for uploaded_file in uploaded_files:
        running_total_bytes += uploaded_file.size
        
    # Check if total batch exceeds 1GB
    if running_total_bytes > MAX_TOTAL_BYTES:
        st.error(f"❌ Batch Rejected: Total size ({running_total_bytes / 1024 / 1024:.2f} MB) exceeds the 1GB system limit.")
    else:
        for uploaded_file in uploaded_files:
            # Enforce individual 100MB gate constraint
            if uploaded_file.size > MAX_BYTES_PER_FILE:
                st.error(f"❌ File Rejected: {uploaded_file.name} exceeds the 100MB limit.")
                continue
                
            st.write(f"⏳ Processing layer: `{uploaded_file.name}` ({uploaded_file.size / 1024 / 1024:.2f} MB)")
            
            try:
                pdf_reader = PdfReader(uploaded_file)
                file_text_accumulator = ""
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        file_text_accumulator += page_text + "\n"
                
                all_extracted_corporate_text += f"\n--- DOCUMENT SOURCE: {uploaded_file.name} ---\n" + file_text_accumulator
                valid_files_processed.append(uploaded_file.name)
                
            except Exception as file_error:
                st.error(f"⚠️ Could not parse structural layers for {uploaded_file.name}: {file_error}")

        if valid_files_processed:
            st.success(f"✅ Input Pipeline Synchronized! Successfully stored text layers from {len(valid_files_processed)} source documents.")

st.markdown("---")

# 4. Deep Operational AI Analysis Loop
st.subheader("⚙️ Step 2: System Evaluation & Audit Execution")

if st.button("Run Comprehensive ESG Cross-Audit"):
    if not all_extracted_corporate_text:
        st.warning("Execution halted. Please drop valid corporate records into the input zone first.")
    else:
        # Constructing a rigid tracking prompt structure for the AI model mapping
        audit_prompt = f"""
        You are an expert Corporate ESG Consultant and Compliance Officer.
        Analyze the raw text extracted from the company's files below and extract data for an executive report.
        
        CRITICAL DIRECTIVE: You must align with ISO 20400 (Sustainable Procurement Guidelines) and the Institute of Corporate Responsibility and Sustainability (ICRS) reporting frameworks.
        Extract any relevant numbers, company names, or metrics you can find. 
        If specific data points are missing or unavailable from the text, you must return '[Data not available]' or '[To be completed]' for that section—do not skip or omit the section completely.
        
        CORPUS OF EXTRACTED COMPANY TEXT DATA:
        {all_extracted_corporate_text[:40000]} 
        
        Provide your audit findings clearly written for these exact four areas:
        1. EXECUTIVE_SUMMARY: (Summarize the general sustainability stance of the company)
        2. ENVIRONMENTAL_FINDINGS: (Highlight direct/indirect emissions data or resource usage found)
        3. SOCIAL_FINDINGS: (Highlight workplace metrics, safety data, and supply chain procurement checks)
        4. GOVERNANCE_FINDINGS: (Highlight board independence, ethics policies, or whistleblower paths)
        """
        
        with st.spinner("AI engine analyzing records and compiling framework matrices..."):
            try:
                # Initialize direct API communication link with the Gemini engine using manual token placement
                api_key_secret = st.secrets["GEMINI_API_KEY"]
                client = genai.Client(api_key=api_key_secret)
                
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=audit_prompt,
                )
                
                ai_analysis_text = response.text
                
                # PREVIOUS WORK DISPLAYED: Render report preview visually on-screen inside Streamlit
                st.markdown("### 📄 Real-Time ESG Assessment Analysis Preview")
                st.info("The screen layout below displays your raw parsed AI data stream. Scroll down to download the formatted PDF.")
                st.write(ai_analysis_text)
                
                # ADDITIONAL WORK EXECUTED: Core HTML-to-PDF Professional Layout Generator Engine
                html_template = f"""
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <style>
                        @page {{
                            size: A4;
                            margin: 20mm 15mm;
                            @bottom-right {{
                                content: counter(page);
                                font-family: sans-serif;
                                font-size: 8pt;
                                color: #718096;
                            }}
                            @bottom-left {{
                                content: "ISO 20400 & ICRS Integrated ESG Disclosure Framework Matrix";
                                font-family: sans-serif;
                                font-size: 8pt;
                                color: #718096;
                            }}
                        }}
                        body {{
                            font-family: sans-serif;
                            color: #2d3748;
                            line-height: 1.6;
                            font-size: 10pt;
                        }}
                        .cover-page {{ page-break-after: always; padding-top: 50mm; }}
                        .cover-title {{ font-size: 26pt; font-weight: bold; color: #1a202c; border-left: 6px solid #2f855a; padding-left: 15px; }}
                        .cover-subtitle {{ font-size: 13pt; color: #4a5568; margin-top: 10px; text-transform: uppercase; }}
                        h1 {{ font-size: 18pt; color: #1a202c; border-bottom: 2px solid #2f855a; padding-bottom: 4px; margin-top: 30px; page-break-after: avoid; }}
                        h2 {{ font-size: 13pt; color: #2f855a; margin-top: 20px; page-break-after: avoid; }}
                        .placeholder {{ background-color: #f7fafc; border-left: 3px dashed #cbd5e0; padding: 10px; margin: 10px 0; font-style: italic; color: #718096; }}
                        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; page-break-inside: avoid; }}
                        th {{ background-color: #2f855a; color: white; padding: 8px; font-size: 9pt; border: 1px solid #2f855a; text-align: left; }}
                        td {{ padding: 8px; border: 1px solid #e2e8f0; font-size: 9pt; }}
                        .matrix-table th {{ background-color: #1a202c; border: 1px solid #1a202c; }}
                        .toc-item {{ border-bottom: 1px dotted #cbd5e0; margin-bottom: 8px; font-weight: bold; }}
                        .toc-page {{ float: right; color: #2f855a; }}
                    </style>
                </head>
                <body>

                    <div class="cover-page">
                        <div class="cover-title">Integrated Corporate ESG<br>Compliance Template Report</div>
                        <div class="cover-subtitle">Standardized Regulatory Analysis Protocol Baseline</div>
                        <p style="margin-top: 40mm; font-size: 11pt;"><strong>Frameworks:</strong> ISO 20400:2017 Sustainable Procurement & ICRS Benchmarks</p>
                        <p><strong>Status:</strong> System Verified Complete Template</p>
                    </div>

                    <div style="page-break-after: always;">
                        <h1>Table of Contents</h1>
                        <div class="toc-item">1. Executive Summary <span class="toc-page">Page 3</span></div>
                        <div class="toc-item">2. Environmental (E) Performance Architecture <span class="toc-page">Page 4</span></div>
                        <div class="toc-item">3. Social (S) Impact & Sustainable Procurement <span class="toc-page">Page 5</span></div>
                        <div class="toc-item">4. Corporate Governance (G) Control Matrix <span class="toc-page">Page 6</span></div>
                        <div class="toc-item">5. Risk & Compliance Registry <span class="toc-page">Page 7</span></div>
                        <div class="toc-item">6. Sustainability Strategy & Forward-Looking Goals <span class="toc-page">Page 8</span></div>
                        <div class="toc-item">7. Core Metrics & KPIs Grid <span class="toc-page">Page 9</span></div>
                        <div class="toc-item">8. Global Reporting Standards Mapping Matrix <span class="toc-page">Page 10</span></div>
                    </div>

                    <h1>1. Executive Summary</h1>
                    <h2>1.1 Core Synthesis Findings</h2>
                    <p>The following parameters represent the extracted organizational data matching international compliance structures:</p>
                    <div class="placeholder">{ai_analysis_text}</div>

                    <h1>2. Environmental (E) Performance Architecture</h1>
                    <h2>2.1 Carbon Emissions & Resource Balances</h2>
                    <table>
                        <thead>
                            <tr><th>Environmental Asset Track</th><th>Assessed Document Value Baseline</th><th>Compliance Status Tracking</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>Scope 1 Direct Carbon Disclosures</td><td>[To be completed / Refer to Sec 1.1]</td><td>Active Verification Pending</td></tr>
                            <tr><td>Scope 2 Indirect Energy Footprint</td><td>[To be completed / Refer to Sec 1.1]</td><td>Active Verification Pending</td></tr>
                            <tr><td>Scope 3 Upstream Sourcing Matrix</td><td>[Data not available]</td><td>ISO 20400 Screening Gate Required</td></tr>
                        </tbody>
                    </table>

                    <h1>3. Social (S) Impact & Sustainable Procurement</h1>
                    <h2>3.1 ISO 20400 Sustainable Sourcing Framework Alignment</h2>
                    <p>Pursuant to ISO 20400 rules and 'The Complete Guide to Sustainability Reporting (2026)', tier-1 vendor supply loops must present zero material human rights exceptions.</p>
                    <div class="placeholder"><strong>Operational Rule Verification:</strong> Sourcing mechanisms must actively trace labor matrices down to local distribution segments. If data missing, leave section fillable for the next cycle.</div>

                    <h1>4. Corporate Governance (G) Control Matrix</h1>
                    <h2>4.1 Fiduciary Safeguards & Whistleblower Mechanisms</h2>
                    <table>
                        <thead>
                            <tr><th>Governance Criteria Metric</th><th>Evaluated System State</th><th>Target Regulatory Constraint</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>Board Independence Profile Split</td><td>[To be completed]</td><td>Must exceed 50% split verification</td></tr>
                            <tr><td>Anonymous Whistleblower Hotline Link</td><td>[Data not available]</td><td>Mandatory compliance safety rail requirement</td></tr>
                        </tbody>
                    </table>

                    <h1>5. Risk & Compliance Registry</h1>
                    <h2>5.1 Regulatory and Framework Compliance Exposure</h2>
                    <p>Maps institutional vulnerabilities and legislative transition tracking vectors under ICRS standards.</p>
                    <div class="placeholder">[To be completed: Manually enter enterprise transition financial risk coefficients here]</div>

                    <h1>6. Sustainability Strategy & Forward-Looking Goals</h1>
                    <h2>6.1 Decarbonization Targets & Strategy Track</h2>
                    <p>Clear, multi-year milestones translating corporate metrics into technical operational milestones.</p>
                    <div class="placeholder">[Data not available: Insert specific corporate net-zero target year data tables here]</div>

                    <h1>7. Core Metrics & KPIs Grid</h1>
                    <p>Standardized placeholders for quantitative data streams utilized during multi-cycle internal governance audits.</p>
                    <div class="placeholder">[To be completed: Insert verified historical baseline metric data streams here]</div>

                    <h1>8. Global Reporting Standards Mapping Matrix</h1>
                    <h2>8.1 Cross-Standard Compatibility Reference</h2>
                    <table class="matrix-table">
                        <thead>
                            <tr>
                                <th>Report Core Section</th>
                                <th>ISO 20400 Guideline</th>
                                <th>ICRS ESG Parameter</th>
                                <th>GRI Reference</th>
                                <th>SASB / TCFD Metric Grid</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>1. Executive Summary</strong></td>
                                <td>Clause 4.3: Context</td>
                                <td>Sec 1: Leadership</td>
                                <td>GRI 2 (2-22)</td>
                                <td>TCFD Governance (A)</td>
                            </tr>
                            <tr>
                                <td><strong>2. Environmental</strong></td>
                                <td>Clause 6.4: Environment</td>
                                <td>Sec 4: Climate Core</td>
                                <td>GRI 302 / 305</td>
                                <td>SASB EM-CM-110a.1</td>
                            </tr>
                            <tr>
                                <td><strong>3. Social Impact</strong></td>
                                <td>Clause 6.3: Human Rights</td>
                                <td>Sec 2: Workplace</td>
                                <td>GRI 401 / 403</td>
                                <td>SASB HC-DY-320a.1</td>
                            </tr>
                            <tr>
                                <td><strong>4. Governance Matrix</strong></td>
                                <td>Clause 5.2: Procurement</td>
                                <td>Sec 3: Ethics Controls</td>
                                <td>GRI 2 Governance</td>
                                <td>TCFD Risk Framework</td>
                            </tr>
                        </tbody>
                    </table>

                </body>
                </html>
                """
                
                # Save the compilation HTML template text to disk
                with open("temp_report.html", "w", encoding="utf-8") as f:
                    f.write(html_template)
                
                # Process the local HTML asset directly into the final layout PDF
                HTML("temp_report.html").write_pdf("esg_report_template.pdf")
                
                st.markdown("---")
                st.subheader("💾 Step 3: Export Framework Document Asset")
                
                # Supply the downloadable PDF button interface component
                with open("esg_report_template.pdf", "rb") as pdf_asset:
                    st.download_button(
                        label="📥 Download Generated ESG Report Template (PDF)",
                        data=pdf_asset,
                        file_name="esg_report_template.pdf",
                        mime="application/pdf"
                    )
                st.balloons()
                
            except Exception as api_err:
                st.error(f"API Executing Error: {api_err}")
