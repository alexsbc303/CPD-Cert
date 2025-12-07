import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
from docx2pdf import convert
from pypdf import PdfReader, PdfWriter
import os
import sys
import threading
if sys.platform == 'win32':
    import pythoncom  # Required for docx2pdf in threads on Windows

# ================= CONFIGURATION =================
# Zoom columns
COL_ZOOM_EMAIL = 'Email'
COL_ZOOM_TIME = 'Time in Session (minutes)'
MIN_MINUTES = 10

# Registration columns (Based on your provided CSV)
COL_REG_EMAIL = 'Email Address 電郵地址'
COL_REG_FIRST = 'First Name 名字'
COL_REG_LAST = 'Last Name 姓氏'
COL_REG_SALUTATION = 'Salutation 稱呼'
COL_REG_MEMBERSHIP = 'Membership No. 會員編號 (If Any, 如有)'

# ================= LOGIC =================

def scrape_event_details(url):
    """Scrapes the HKIE website using the specific IDs provided."""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # 1) Find Title
        title_tag = soup.find(id='ctl00_ContentPlaceHolder1_ContentName')
        event_title = title_tag.get_text(strip=True) if title_tag else "Unknown Event"

        # 2) Find Date, Time & Venue
        dtv_tag = soup.find(id='ctl00_ContentPlaceHolder1_dtv')
        event_details = dtv_tag.get_text(strip=True) if dtv_tag else "Unknown Details"

        return event_title, event_details
    except Exception as e:
        return None, str(e)

def get_header_row_index(csv_path, target_col_name):
    """Finds the row number where the actual header starts (handling metadata rows)."""
    with open(csv_path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
        for i, line in enumerate(lines):
            if target_col_name in line:
                return i
    return 0

def generate_certificates(reg_path, zoom_path, template_path, url, output_folder, update_status):
    if sys.platform == 'win32':
        pythoncom.CoInitialize() # Initialize COM for Word automation
    
    try:
        update_status("Step 1/5: Scraping Event Details...")
        event_title, event_details = scrape_event_details(url)
        if not event_title:
            raise Exception(f"Failed to scrape URL: {event_details}")
        
        update_status(f"Event Found: {event_title}")

        # --- Load Data ---
        update_status("Step 2/5: Processing Excel Files...")
        
        # Load Reg (Excel 1)
        df_reg = pd.read_csv(reg_path)
        
        # Load Zoom (Excel 2) - Auto-detect header
        zoom_header_idx = get_header_row_index(zoom_path, COL_ZOOM_TIME)
        df_zoom = pd.read_csv(zoom_path, skiprows=zoom_header_idx)

        # Normalize Emails
        df_reg['clean_email'] = df_reg[COL_REG_EMAIL].astype(str).str.lower().str.strip()
        df_zoom['clean_email'] = df_zoom[COL_ZOOM_EMAIL].astype(str).str.lower().str.strip()

        # Filter Zoom Attendance
        # Ensure time column is numeric
        df_zoom[COL_ZOOM_TIME] = pd.to_numeric(df_zoom[COL_ZOOM_TIME], errors='coerce').fillna(0)
        df_zoom_valid = df_zoom[df_zoom[COL_ZOOM_TIME] >= MIN_MINUTES]

        # Merge Data
        merged_df = pd.merge(df_reg, df_zoom_valid, on='clean_email', how='inner')
        merged_df = merged_df.drop_duplicates(subset=['clean_email'])
        
        total_count = len(merged_df)
        if total_count == 0:
            raise Exception("No matching attendees found who met the time requirement.")

        update_status(f"Step 3/5: Generating {total_count} Certificates...")

        # --- Generate ---
        temp_docx_folder = os.path.join(output_folder, "temp_docs")
        if not os.path.exists(temp_docx_folder):
            os.makedirs(temp_docx_folder)

        for index, row in merged_df.iterrows():
            # 1. Prepare Data
            salutation = str(row.get(COL_REG_SALUTATION, '')).strip()
            first_name = str(row.get(COL_REG_FIRST, '')).strip().title()
            last_name = str(row.get(COL_REG_LAST, '')).strip().upper()
            full_name = f"{salutation} {last_name} {first_name}".strip()

            # 2. Determine Password
            # Priority 1: Membership No.
            mem_no = str(row.get(COL_REG_MEMBERSHIP, '')).strip()
            # Check if membership no is valid (not 'nan', not empty)
            if mem_no and mem_no.lower() != 'nan':
                password = mem_no
            else:
                # Priority 2: Email
                password = str(row.get(COL_REG_EMAIL, '')).strip()

            # 3. Fill Template
            doc = DocxTemplate(template_path)
            context = {
                'name': full_name,
                'event_title': event_title,
                'event_details': event_details
            }
            doc.render(context)
            
            temp_docx_path = os.path.join(temp_docx_folder, f"temp_{index}.docx")
            doc.save(temp_docx_path)

            # 4. Convert to PDF
            pdf_name = f"{full_name}_CPD_Cert.pdf".replace("/", "-") # Sanitize filename
            final_pdf_path = os.path.join(output_folder, pdf_name)
            
            # Note: docx2pdf converts to same folder, we handle move/rename
            convert(temp_docx_path, final_pdf_path)

            # 5. Encrypt PDF
            reader = PdfReader(final_pdf_path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            
            writer.encrypt(password)
            
            # Overwrite with encrypted version
            with open(final_pdf_path, "wb") as f:
                writer.write(f)

            update_status(f"Generated ({index + 1}/{total_count}): {full_name}")

        # Cleanup Temp Files
        update_status("Cleaning up...")
        for f in os.listdir(temp_docx_folder):
            os.remove(os.path.join(temp_docx_folder, f))
        os.rmdir(temp_docx_folder)

        update_status("Success! All certificates generated.")
        messagebox.showinfo("Complete", f"Successfully generated {total_count} certificates in:\n{output_folder}")

    except Exception as e:
        update_status(f"Error: {str(e)}")
        messagebox.showerror("Error", str(e))

# ================= GUI =================
class CPDApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HKIE CPD Certificate Generator")
        self.root.geometry("600x450")

        # Variables
        self.reg_path = tk.StringVar()
        self.zoom_path = tk.StringVar()
        self.tpl_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.url_var = tk.StringVar()

        # Layout
        pad_opts = {'padx': 10, 'pady': 5, 'sticky': 'w'}
        
        # 1. Reg File
        tk.Label(root, text="1. Registration Excel (csv):").grid(row=0, column=0, **pad_opts)
        tk.Entry(root, textvariable=self.reg_path, width=40).grid(row=0, column=1, **pad_opts)
        tk.Button(root, text="Browse", command=lambda: self.browse_file(self.reg_path, "csv")).grid(row=0, column=2, **pad_opts)

        # 2. Zoom File
        tk.Label(root, text="2. Zoom Report (csv):").grid(row=1, column=0, **pad_opts)
        tk.Entry(root, textvariable=self.zoom_path, width=40).grid(row=1, column=1, **pad_opts)
        tk.Button(root, text="Browse", command=lambda: self.browse_file(self.zoom_path, "csv")).grid(row=1, column=2, **pad_opts)

        # 3. Template File
        tk.Label(root, text="3. Word Template (docx):").grid(row=2, column=0, **pad_opts)
        tk.Entry(root, textvariable=self.tpl_path, width=40).grid(row=2, column=1, **pad_opts)
        tk.Button(root, text="Browse", command=lambda: self.browse_file(self.tpl_path, "docx")).grid(row=2, column=2, **pad_opts)

        # 4. Output Folder
        tk.Label(root, text="4. Output Folder:").grid(row=3, column=0, **pad_opts)
        tk.Entry(root, textvariable=self.out_path, width=40).grid(row=3, column=1, **pad_opts)
        tk.Button(root, text="Browse", command=self.browse_folder).grid(row=3, column=2, **pad_opts)

        # 5. URL
        tk.Label(root, text="5. Event URL:").grid(row=4, column=0, **pad_opts)
        tk.Entry(root, textvariable=self.url_var, width=50).grid(row=4, column=1, columnspan=2, **pad_opts)

        # Separator
        ttk.Separator(root, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

        # Generate Button
        self.btn_run = tk.Button(root, text="GENERATE CERTIFICATES", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=self.start_process)
        self.btn_run.grid(row=6, column=0, columnspan=3, pady=10)

        # Status Label
        self.lbl_status = tk.Label(root, text="Ready", fg="blue")
        self.lbl_status.grid(row=7, column=0, columnspan=3)

    def browse_file(self, var, ftype):
        if ftype == "csv":
            path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        else:
            path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if path: var.set(path)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path: self.out_path.set(path)

    def update_status(self, msg):
        self.lbl_status.config(text=msg)
        self.root.update_idletasks()

    def start_process(self):
        # Validation
        if not all([self.reg_path.get(), self.zoom_path.get(), self.tpl_path.get(), self.out_path.get(), self.url_var.get()]):
            messagebox.showwarning("Missing Info", "Please select all files, folder and enter the URL.")
            return

        self.btn_run.config(state="disabled")
        
        # Run in thread to keep GUI responsive
        t = threading.Thread(target=lambda: generate_certificates(
            self.reg_path.get(),
            self.zoom_path.get(),
            self.tpl_path.get(),
            self.url_var.get(),
            self.out_path.get(),
            self.update_status
        ))
        t.start()
        
        # Check thread to re-enable button
        self.root.after(100, self.check_thread, t)

    def check_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, self.check_thread, thread)
        else:
            self.btn_run.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = CPDApp(root)
    root.mainloop()