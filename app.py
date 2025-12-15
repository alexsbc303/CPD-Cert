import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
import pikepdf
import os
import re
import zipfile
import tempfile
import sys
import platform
import time

# --- 自動偵測作業系統 ---
current_os = platform.system()
is_windows = current_os == 'Windows'
is_mac = current_os == 'Darwin'

# 根據 OS 載入對應的轉檔工具
if is_windows:
    import pythoncom
    import win32com.client
elif is_mac:
    # Mac 需要 docx2pdf
    try:
        from docx2pdf import convert as mac_convert
    except ImportError:
        pass

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator", layout="wide")

st.title("⚡ HKIE CPD 證書生成器")

# 顯示當前系統狀態
if is_mac:
    st.info("🍎 偵測到 macOS 環境：使用 docx2pdf 進行轉檔。")
elif is_windows:
    st.info("🪟 偵測到 Windows 環境：使用 Win32 COM 極速引擎進行轉檔。")
else:
    st.warning("⚠️ 偵測到 Linux/其他環境：僅支援生成 Word 檔，無法生成 PDF。")

# --- 1. 獲取活動資訊 ---
st.header("1. 獲取活動資訊")
url = st.text_input("輸入 HKIE 活動網址", "http://it.hkie.org.hk/en_it_events_inside_Past.aspx?EventID=600&&TypeName=Events+%2f+Activities")

if 'event_title' not in st.session_state:
    st.session_state['event_title'] = ""
if 'event_details' not in st.session_state:
    st.session_state['event_details'] = ""

if st.button("抓取活動資訊"):
    try:
        response = requests.get(url)
        response.encoding = 'utf-8' 
        soup = BeautifulSoup(response.content, 'html.parser')
        
        title_tag = soup.find(id="ctl00_ContentPlaceHolder1_ContentName")
        if title_tag:
            st.session_state['event_title'] = title_tag.get_text(strip=True)
        else:
            st.warning("找不到標題，請手動輸入。")

        dtv_tag = soup.find(id="ctl00_ContentPlaceHolder1_dtv")
        if dtv_tag:
            raw_dtv = dtv_tag.get_text(strip=True).replace(';', ' ')
            st.session_state['event_details'] = raw_dtv
        else:
            st.warning("找不到日期時間，請手動輸入。")
            
        st.success("資訊抓取成功！")
    except Exception as e:
        st.error(f"抓取失敗: {e}")

col1, col2 = st.columns(2)
with col1:
    event_title = st.text_input("活動標題", value=st.session_state['event_title'])
with col2:
    event_details = st.text_input("日期與時間", value=st.session_state['event_details'])

# --- 2. 上傳檔案 ---
st.header("2. 上傳資料檔")

reg_file = st.file_uploader("上傳報名表 (Registration Excel) [必填]", type=['xlsx', 'csv'])
template_file = st.file_uploader("上傳證書範本 (Word .docx) [必填]", type=['docx'])

use_zoom = st.checkbox("需要核對 Zoom 出席紀錄？", value=True)
zoom_file = None
if use_zoom:
    zoom_file = st.file_uploader("上傳 Zoom 報告 (Attendee Excel) [選填]", type=['xlsx', 'csv'])

# --- 輔助函式 ---
def normalize_name(name):
    if pd.isna(name): return ""
    name = str(name).lower()
    name = re.sub(r'\b(ir|mr|ms|miss|dr|prof)\b\.?', '', name)
    name = re.sub(r'[^a-z\s]', '', name)
    return " ".join(name.split())

def find_header_row(df_preview, keywords=["User Name", "Email"]):
    for i, row in df_preview.iterrows():
        row_str_list = [str(val) for val in row.values]
        if all(any(kw in cell for cell in row_str_list) for kw in keywords):
            return i
    return 0

def load_data(uploaded_file):
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    except:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)

# --- 3. 數據處理 ---
df_final = pd.DataFrame()

if reg_file and template_file:
    if use_zoom and not zoom_file:
        st.warning("請上傳 Zoom 檔案或取消勾選核對選項。")
    else:
        st.header("3. 處理名單")
        try:
            # A. 讀取報名表
            df_reg = load_data(reg_file)
            
            # 欄位對應
            col_map = {}
            for c in df_reg.columns:
                c_lower = str(c).lower()
                if 'first name' in c_lower or '名字' in c_lower:
                    col_map[c] = 'First Name'
                elif 'last name' in c_lower or '姓氏' in c_lower:
                    col_map[c] = 'Last Name'
                elif 'email' in c_lower or '電郵' in c_lower:
                    col_map[c] = 'Email'
                elif 'membership' in c_lower or '會員編號' in c_lower:
                    col_map[c] = 'Membership No'
                elif 'salutation' in c_lower or '稱呼' in c_lower:
                    col_map[c] = 'Salutation'
            
            df_reg.rename(columns=col_map, inplace=True)
            
            if 'Membership No' not in df_reg.columns:
                st.warning("⚠️ 警告：無法自動識別 'Membership No' 欄位。")
                df_reg['Membership No'] = ""
            
            required_cols = ['First Name', 'Last Name', 'Email']
            if not all(col in df_reg.columns for col in required_cols):
                st.error(f"報名表缺少必要欄位: {required_cols}")
                st.write("目前偵測到的欄位:", df_reg.columns.tolist())
                st.stop()

            # --- [新增功能 1] 顯示報名表 Email 清單 ---
            st.info(f"📄 已成功讀取報名表，共 {len(df_reg)} 筆資料。")
            with st.expander("🔍 點擊查看原始報名名單 (Email List)"):
                st.dataframe(df_reg[['First Name', 'Last Name', 'Email', 'Membership No']])

            # B. 核對 Zoom
            if not use_zoom:
                df_final = df_reg.copy()
                df_final['Full Name'] = df_final['First Name'].astype(str) + " " + df_final['Last Name'].astype(str)
                df_final['Match Method'] = "Registration Only"
            else:
                # Zoom 處理
                try:
                    zoom_file.seek(0)
                    df_preview = pd.read_csv(zoom_file, header=None, nrows=20)
                except:
                    zoom_file.seek(0)
                    df_preview = pd.read_excel(zoom_file, header=None, nrows=20)
                
                header_row = find_header_row(df_preview)
                
                zoom_file.seek(0)
                try:
                    df_zoom = pd.read_csv(zoom_file, header=header_row)
                except:
                    df_zoom = pd.read_excel(zoom_file, header=header_row)
                
                z_user_col = next((c for c in df_zoom.columns if "User Name" in str(c)), None)
                z_email_col = next((c for c in df_zoom.columns if "Email" in str(c)), None)
                
                if not z_user_col or not z_email_col:
                    st.error("Zoom 檔案無法識別 User Name 或 Email 欄位。")
                    st.stop()
                
                if 'Attended' in df_zoom.columns:
                    df_zoom = df_zoom[df_zoom['Attended'] == 'Yes']

                st.write("正在核對 Zoom 資料...")
                df_reg['Name_Norm'] = (df_reg['First Name'].astype(str) + " " + df_reg['Last Name'].astype(str)).apply(normalize_name)
                df_reg['Email_Norm'] = df_reg['Email'].astype(str).str.lower().str.strip()
                
                df_zoom['Name_Norm'] = df_zoom[z_user_col].apply(normalize_name)
                df_zoom['Email_Norm'] = df_zoom[z_email_col].astype(str).str.lower().str.strip()
                
                zoom_email_map = df_zoom.set_index('Email_Norm')[z_user_col].to_dict()
                zoom_name_map = df_zoom.set_index('Name_Norm')[z_user_col].to_dict()
                
                matched_list = []
                for _, row in df_reg.iterrows():
                    status = "Unmatched"
                    if row['Email_Norm'] in zoom_email_map:
                        status = "Matched (Email)"
                    elif row['Name_Norm'] in zoom_name_map:
                        status = "Matched (Name)"
                    
                    if "Matched" in status:
                        matched_list.append({
                            "Salutation": row.get('Salutation', ''),
                            "Full Name": f"{row.get('First Name', '')} {row.get('Last Name', '')}",
                            "Membership No": row.get('Membership No', ''),
                            "Email": row.get('Email', ''),
                            "Match Method": status
                        })
                df_final = pd.DataFrame(matched_list)

            if not df_final.empty:
                st.success(f"✅ 核對完成！共產生 {len(df_final)} 筆證書名單。")
                # --- [新增功能 2] 顯示最終 Email 清單 ---
                with st.expander("🔍 點擊查看將獲發證書的 Email 清單 (Final List)"):
                    st.dataframe(df_final[['Salutation', 'Full Name', 'Email', 'Membership No', 'Match Method']])
            else:
                st.warning("沒有符合的名單。")

        except Exception as e:
            st.error(f"資料處理發生錯誤: {e}")

    # --- 4. 生成與下載 ---
    st.header("4. 生成證書")
    
    output_format = st.radio(
        "選擇輸出格式：",
        ('Word 文件 (.docx) - 不加密', 'PDF 文件 (.pdf) - 加密 (密碼: Email)')
    )
    
    if st.button("開始生成"):
        if df_final.empty:
            st.error("名單為空。")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            with tempfile.TemporaryDirectory() as tmpdirname:
                zip_filename = "certs_output.zip"
                zip_path = os.path.join(tmpdirname, zip_filename)
                template_path = os.path.join(tmpdirname, "template.docx")
                
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                
                generated_files = []
                total = len(df_final)
                success_count = 0
                
                # Windows 批次初始化
                word_app = None
                if output_format.startswith('PDF') and is_windows:
                    try:
                        pythoncom.CoInitialize()
                        word_app = win32com.client.DispatchEx("Word.Application")
                        word_app.Visible = False
                        word_app.DisplayAlerts = False
                    except Exception as e:
                        st.error(f"無法啟動 Word (Windows): {e}")
                        st.stop()
                
                try:
                    for i, person in df_final.iterrows():
                        person_name = str(person['Full Name']).strip()
                        status_text.text(f"處理中 ({i+1}/{total}): {person_name}")
                        
                        try:
                            # 1. 產生 DOCX
                            doc_tpl = DocxTemplate(template_path)
                            
                            mem_no = str(person['Membership No'])
                            if mem_no.lower() in ['nan', 'none', '']: mem_no = ""
                            if mem_no.endswith('.0'): mem_no = mem_no[:-2]

                            context = {
                                'name': f"{person['Salutation']} {person_name}",
                                'membership_no': mem_no, 
                                'event_title': event_title,
                                'event_details': event_details
                            }
                            doc_tpl.render(context)
                            
                            safe_name = re.sub(r'[\\/*?:"<>|]', "", person_name)
                            docx_filename = f"{safe_name}.docx"
                            docx_path = os.path.join(tmpdirname, docx_filename)
                            doc_tpl.save(docx_path)
                            
                            final_file_path = docx_path
                            
                            # 2. 轉 PDF
                            if output_format.startswith('PDF'):
                                try:
                                    pdf_filename = f"{safe_name}.pdf"
                                    # Mac 需要絕對路徑
                                    pdf_path = os.path.join(tmpdirname, pdf_filename)
                                    abs_docx = os.path.abspath(docx_path)
                                    abs_pdf = os.path.abspath(pdf_path)

                                    # A. Windows: COM
                                    if is_windows and word_app:
                                        wb_doc = word_app.Documents.Open(abs_docx)
                                        wb_doc.SaveAs(abs_pdf, FileFormat=17)
                                        wb_doc.Close(SaveChanges=False)
                                    
                                    # B. Mac: docx2pdf (修復這裡的邏輯)
                                    elif is_mac:
                                        mac_convert(abs_docx, abs_pdf)
                                        time.sleep(0.5) # 緩衝

                                    # C. 加密
                                    password = str(person['Email']).strip()
                                    if not password or password == 'nan':
                                        password = "hkie"
                                        
                                    encrypted_path = os.path.join(tmpdirname, f"Encrypted_{safe_name}.pdf")
                                    
                                    # 確保檔案已生成
                                    if not os.path.exists(abs_pdf):
                                         raise FileNotFoundError("PDF file was not created.")

                                    with pikepdf.Pdf.open(abs_pdf) as pdf:
                                        pdf.save(encrypted_path, encryption=pikepdf.Encryption(owner=password, user=password, R=6))
                                    
                                    final_file_path = encrypted_path
                                    
                                except Exception as e:
                                    # st.warning(f"{person_name} 轉檔失敗 (保留 Word): {e}")
                                    final_file_path = docx_path
                            
                            generated_files.append(final_file_path)
                            success_count += 1
                            
                        except Exception as e:
                            st.error(f"生成 {person_name} 時錯誤: {e}")
                            if "expected token" in str(e): st.stop()

                        progress_bar.progress((i + 1) / total)
                        
                finally:
                    if is_windows and word_app:
                        try:
                            word_app.Quit()
                        except:
                            pass
                        pythoncom.CoUninitialize()
                
                if generated_files:
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for file in generated_files:
                            zipf.write(file, os.path.basename(file))
                            
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label=f"📥 下載完成 ({success_count} 個檔案)",
                            data=f,
                            file_name=zip_filename,
                            mime="application/zip"
                        )
                    st.success("任務完成！")