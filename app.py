import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter
import os
import shutil
import subprocess
import zipfile
import io

# ================= CONFIGURATION =================
MIN_MINUTES = 10
COL_ZOOM_TIME = 'Time in Session (minutes)'

# ================= HELPER FUNCTIONS =================
def scrape_event_details(url):
    """Scrapes the HKIE website for event details."""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        
        title_tag = soup.find(id='ctl00_ContentPlaceHolder1_ContentName')
        event_title = title_tag.get_text(strip=True) if title_tag else "Unknown Event"
        
        dtv_tag = soup.find(id='ctl00_ContentPlaceHolder1_dtv')
        event_details = dtv_tag.get_text(strip=True) if dtv_tag else "Unknown Details"
        
        return event_title, event_details
    except Exception as e:
        return None, str(e)

def convert_docx_to_pdf(input_path, output_folder):
    """Converts DOCX to PDF using LibreOffice (Linux/Server compatible)."""
    # Note: This requires LibreOffice to be installed on the server
    cmd = [
        'libreoffice', '--headless', '--convert-to', 'pdf', 
        '--outdir', output_folder, input_path
    ]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    
    # LibreOffice saves with same filename but .pdf extension
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(output_folder, f"{base_name}.pdf")

def encrypt_pdf(input_pdf_path, password):
    """Encrypts the PDF with a password."""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(password)
    
    # Overwrite with encrypted version
    with open(input_pdf_path, "wb") as f:
        writer.write(f)

# ================= STREAMLIT APP =================
st.set_page_config(page_title="CPD Cert Generator", layout="wide")
st.title("🎓 HKIE CPD Certificate Generator")
st.markdown("Upload your files below to generate encrypted PDF certificates.")

# --- 1. FILE UPLOAD SECTION ---
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        reg_file = st.file_uploader("1. Registration Excel (CSV)", type=['csv'])
        zoom_file = st.file_uploader("2. Zoom Report (CSV)", type=['csv'])
    with col2:
        template_file = st.file_uploader("3. Word Template (.docx)", type=['docx'])
        event_url = st.text_input("4. Event URL", placeholder="http://it.hkie.org.hk/...")

# --- 2. PROCESSING SECTION ---
if st.button("Generate Certificates", type="primary"):
    if not all([reg_file, zoom_file, template_file, event_url]):
        st.error("Please upload all 3 files and provide the Event URL.")
        st.stop()
        
    status = st.empty()
    progress = st.progress(0)
    
    # Create temp workspace
    temp_dir = "temp_workspace"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    try:
        # Step A: Scrape Data
        status.info("⏳ Scraping event details...")
        event_title, event_details = scrape_event_details(event_url)
        if not event_title:
            st.error("Failed to scrape URL.")
            st.stop()
        
        # Step B: Process Dataframes
        status.info("⏳ Processing data matches...")
        
        # Save template to disk for docxtpl
        tpl_path = os.path.join(temp_dir, "template.docx")
        with open(tpl_path, "wb") as f:
            f.write(template_file.getbuffer())
            
        df_reg = pd.read_csv(reg_file)
        
        # Handle Zoom metadata rows
        zoom_content = zoom_file.getvalue().decode('utf-8', errors='replace')
        header_idx = 0
        for i, line in enumerate(zoom_content.splitlines()):
            if COL_ZOOM_TIME in line:
                header_idx = i
                break
        
        zoom_file.seek(0)
        df_zoom = pd.read_csv(zoom_file, skiprows=header_idx)
        
        # Normalize columns
        reg_email_col = next(col for col in df_reg.columns if "Email" in col)
        zoom_email_col = next(col for col in df_zoom.columns if "Email" in col)
        
        df_reg['clean_email'] = df_reg[reg_email_col].astype(str).str.lower().str.strip()
        df_zoom['clean_email'] = df_zoom[zoom_email_col].astype(str).str.lower().str.strip()
        
        # Filter Zoom
        df_zoom[COL_ZOOM_TIME] = pd.to_numeric(df_zoom[COL_ZOOM_TIME], errors='coerce').fillna(0)
        valid_attendees = df_zoom[df_zoom[COL_ZOOM_TIME] >= MIN_MINUTES]
        
        # Merge
        merged = pd.merge(df_reg, valid_attendees, on='clean_email', how='inner').drop_duplicates(subset=['clean_email'])
        
        if len(merged) == 0:
            st.warning("No matching attendees found meeting the time requirements.")
            st.stop()
            
        # Step C: Generate Certs
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, false_compress=True) as zf:
            
            for idx, row in merged.iterrows():
                # Update progress
                pct = int((idx + 1) / len(merged) * 100)
                progress.progress(pct)
                status.info(f"Generating certificate {idx+1}/{len(merged)}...")

                # Prepare Data
                salutation = str(row.get('Salutation 稱呼', '')).strip()
                first = str(row.get('First Name 名字', '')).strip().title()
                last = str(row.get('Last Name 姓氏', '')).strip().upper()
                full_name = f"{salutation} {last} {first}".strip()
                
                # Determine Password
                mem_no = str(row.get('Membership No. 會員編號 (If Any, 如有)', ''))
                password = mem_no if mem_no and mem_no.lower() != 'nan' else row['clean_email']
                
                # Render Docx
                doc = DocxTemplate(tpl_path)
                context = {'name': full_name, 'event_title': event_title, 'event_details': event_details}
                doc.render(context)
                
                temp_docx = os.path.join(temp_dir, f"temp_{idx}.docx")
                doc.save(temp_docx)
                
                # Convert to PDF (LibreOffice)
                pdf_path = convert_docx_to_pdf(temp_docx, temp_dir)
                
                # Encrypt
                if os.path.exists(pdf_path):
                    encrypt_pdf(pdf_path, password)
                    
                    # Add to Zip
                    clean_filename = f"{full_name}_CPD_Cert.pdf".replace("/", "-")
                    zf.write(pdf_path, clean_filename)

        status.success(f"✅ Generated {len(merged)} certificates!")
        
        # Download Button
        st.download_button(
            label="Download All Certificates (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="cpd_certificates.zip",
            mime="application/zip"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        # Cleanup
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)