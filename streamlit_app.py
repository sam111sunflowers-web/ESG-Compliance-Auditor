import streamlit as st
from google import genai
from pypdf import PdfReader
import os

# 1. Page Configuration Architecture
st.set_page_config(page_title="Batch ESG Auditor", layout="wide")
st.title("⚖️ Enterprise Multi-Report ESG Compliance Auditor")
st.caption("Automated Extraction & High-Speed Policy Risk Analysis Pipeline")

st.markdown("---")

# Define strict operational boundaries
# Updated strict operational boundaries
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
    # First, calculate total size of the batch
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
            st.success(f"✅ Input Pipeline Synchronized! Successfully processed {len(valid_files_processed)} source documents.")
            
            # Combine the individual text chunks into our master processing corpus
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
        You are a Principal Executive Auditing Lead specializing in Corporate Governance and ESG Frameworks.
        Analyze the raw extracted text segments pulled from multiple company files to check for regulatory adherence.
        
        CORPUS OF EXTRACTED COMPANY TEXT DATA:
        {all_extracted_corporate_text[:50000]} 
        
        REQUIRED AUDIT ANALYSIS SHEET STRUCTURE TO GENERATE:
        # REGULATORY ESG AUDIT ANALYSIS REPORT
        
        ## 1. AGGREGATED EXECUTIVE SUMMARY
        [Provide a formal evaluation summarizing the general data footprint of the analyzed documents.]
        
        ## 2. COMPLIANCE TARGET TRACKER & VERIFIED NUMBERS
        [List hard data points, percentages, or metrics explicitly mentioned. Note where figures are missing.]
        
        ## 3. IDENTIFIED GREENWASHING RISKS & RED FLAGS
        [Flag instances of vague corporate language, weak tracking structures, or gaps in disclosure policies.]
        
        ## 4. IMMEDIATE ACTION PLANS
        [Detail 3 clear operational adjustments recommended for the board of directors.]
        """
        
        with st.spinner("AI engine cross-referencing multi-file text data against auditing standards..."):
            try:
                # Initialize direct API communication link with the Gemini engine
                client = genai.Client()
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=audit_prompt,
                )
                
                final_report_output = response.text
                
                # Render report visually on-screen
                st.markdown("### 📄 Compiled ESG Assessment Output View")
                st.markdown(final_report_output)
                
                st.markdown("---")
                
                # 5. Output File Creation Engine (Download Matrix)
                st.subheader("💾 Step 3: Export Compliance Asset")
                st.download_button(
                    label="📥 Download Generated esg_report.md",
                    data=final_report_output,
                    file_name="esg_report.md",
                    mime="text/markdown"
                )
                st.balloons()
                
            except Exception as api_err:
                st.error("API Token validation signature missing. Ensure your GEMINI_API_KEY is configured in your Streamlit Cloud Settings.")
