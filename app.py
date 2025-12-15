import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
from docx2pdf import convert
import pikepdf
import os
import re
import zipfile
import tempfile
import sys
import platform

# 檢查是否在 Windows 環境
is_windows = platform.system() == 'Windows'
if is_windows:
    import pythoncom

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator (Debug Mode)", layout="wide")

st.title("🎓 HKIE CPD 證書生成器")
st.markdown("""
**功能說明：**
1. 抓取活動資訊
2. 核對 Zoom 出席名單 (可選)
3. 生成 PDF (加密密碼為 Email) 或 Word 檔
**注意：PDF 生成功能需要伺服器/本機已安裝 Microsoft Word。**
""")

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
        
        # 抓取標題
        title_tag = soup.find(id="ctl00_ContentPlaceHolder1_ContentName")
        if title_tag:
            st.session_state['event_title'] = title_tag.get_text(strip=True)
        else:
            st.warning("找不到標題，請手動輸入。")

        # 抓取日期時間
        dtv_tag = soup.find(id="ctl00_ContentPlaceHolder1_dtv")
        if dtv_tag:
            # 將分號替換為空格
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

reg_file = st.file_uploader("上傳報名表 (Registration Excel) [必填]", type=['csv', 'xlsx'])
template_file = st.file_uploader("上傳證書範本 (Word .docx) [必填]", type=['docx'])

use_zoom = st.checkbox("需要核對 Zoom 出席紀錄？", value=True)
zoom_file = None
if use_zoom:
    zoom_file = st.file_uploader("上傳 Zoom 報告 (Attendee Excel) [選填]", type=['csv', 'xlsx'])

# --- 輔助函式 ---
def normalize_name(name):
    if pd.isna(name): return ""
    name = str(name).lower()
    name = re.sub(r'\b(ir|mr|ms|miss|dr|prof)\b\.?', '', name)
    name = re.sub(r'[^a-z\s]', '', name)
    return " ".join(name.split())

def find_header_row(df_preview, keywords=["User Name", "Email"]):
    """自動尋找 Zoom 報告的標題列"""
    for i, row in df_preview.iterrows():
        row_str_list = [str(val) for val in row.values]
        if all(any(kw in cell for cell in row_str_list) for kw in keywords):
            return i
    return 0

# --- 3. 數據處理 ---
df_final = pd.DataFrame()

if reg_file and template_file:
    if use_zoom and not zoom_file:
        st.warning("請上傳 Zoom 檔案或取消勾選核對選項。")
    else:
        st.header("3. 處理名單")
        try:
            # A. 讀取報名表
            if reg_file.name.endswith('.csv'):
                df_reg = pd.read_csv(reg_file)
            else:
                df_reg = pd.read_excel(reg_file)
            
            # 欄位映射
            col_map = {}
            for c in df_reg.columns:
                if 'First Name' in c: col_map[c] = 'First Name'
                elif 'Last Name' in c: col_map[c] = 'Last Name'
                elif 'Email Address' in c: col_map[c] = 'Email'
                elif 'Membership No' in c: col_map[c] = 'Membership No'
                elif 'Salutation' in c: col_map[c] = 'Salutation'
            
            df_reg.rename(columns=col_map, inplace=True)
            
            # 檢查必要欄位
            required_cols = ['First Name', 'Last Name', 'Email']
            if not all(col in df_reg.columns for col in required_cols):
                st.error(f"報名表缺少必要欄位，請檢查: {required_cols}")
                st.stop()

            # B. 核對 Zoom
            if not use_zoom:
                df_final = df_reg.copy()
                df_final['Full Name'] = df_final['First Name'].astype(str) + " " + df_final['Last Name'].astype(str)
                df_final['Match Method'] = "Registration Only"
            else:
                # 預讀
                if zoom_file.name.endswith('.csv'):
                    df_preview = pd.read_csv(zoom_file, header=None, nrows=20)
                else:
                    df_preview = pd.read_excel(zoom_file, header=None, nrows=20)
                
                header_row = find_header_row(df_preview)
                
                # 重新讀取
                zoom_file.seek(0)
                if zoom_file.name.endswith('.csv'):
                    df_zoom = pd.read_csv(zoom_file, header=header_row)
                else:
                    df_zoom = pd.read_excel(zoom_file, header=header_row)
                
                z_user_col = next((c for c in df_zoom.columns if "User Name" in str(c)), None)
                z_email_col = next((c for c in df_zoom.columns if "Email" in str(c)), None)
                
                if not z_user_col or not z_email_col:
                    st.error("Zoom 檔案無法識別 User Name 或 Email 欄位。")
                    st.stop()
                
                if 'Attended' in df_zoom.columns:
                    df_zoom = df_zoom[df_zoom['Attended'] == 'Yes']

                # 配對
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
                st.success(f"共產生 {len(df_final)} 筆證書名單。")
                st.dataframe(df_final.head())
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
                
                # 儲存範本
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                
                generated_files = []
                total = len(df_final)
                success_count = 0
                error_shown = False # 避免重複顯示相同的 PDF 錯誤
                
                for i, person in df_final.iterrows():
                    person_name = str(person['Full Name']).strip()
                    status_text.text(f"處理中 ({i+1}/{total}): {person_name}")
                    
                    # 1. 產生 Word
                    try:
                        doc = DocxTemplate(template_path)
                        mem_no = str(person['Membership No'])
                        if mem_no == 'nan' or mem_no == 'None': mem_no = ""
                        
                        context = {
                            'name': f"{person['Salutation']} {person_name}",
                            'membership_number': mem_no, 
                            'event_title': event_title,
                            'event_details': event_details
                        }
                        doc.render(context)
                        
                        safe_name = re.sub(r'[\\/*?:"<>|]', "", person_name)
                        docx_filename = f"{safe_name}.docx"
                        docx_path = os.path.join(tmpdirname, docx_filename)
                        doc.save(docx_path)
                        
                        final_file_path = docx_path
                        
                        # 2. 轉 PDF
                        if output_format.startswith('PDF'):
                            try:
                                pdf_filename = f"{safe_name}.pdf"
                                pdf_path = os.path.join(tmpdirname, pdf_filename)
                                
                                # Windows COM 初始化
                                if is_windows:
                                    pythoncom.CoInitialize()
                                
                                # 嘗試轉換 (如果沒有安裝 Word，這裡會報錯)
                                convert(docx_path, pdf_path)
                                
                                # 加密
                                password = str(person['Email']).strip()
                                if not password or password == 'nan':
                                    password = "hkie"
                                    
                                encrypted_path = os.path.join(tmpdirname, f"Encrypted_{safe_name}.pdf")
                                with pikepdf.Pdf.open(pdf_path) as pdf:
                                    pdf.save(encrypted_path, encryption=pikepdf.Encryption(owner=password, user=password, R=6))
                                
                                final_file_path = encrypted_path
                                
                            except Exception as e_pdf:
                                # PDF 失敗時，顯示錯誤但不中斷，回退到 Word
                                if not error_shown:
                                    st.error(f"⚠️ PDF 轉換失敗 (僅顯示一次，後續將自動轉為 Word): {e_pdf}")
                                    st.warning("可能原因：伺服器未安裝 Microsoft Word，或 COM 元件呼叫失敗。")
                                    error_shown = True
                                final_file_path = docx_path
                        
                        generated_files.append(final_file_path)
                        success_count += 1
                        
                    except Exception as e:
                        st.error(f"生成 {person_name} 時發生嚴重錯誤: {e}")
                        if "expected token" in str(e):
                            st.error("❌ 請檢查 Word 範本變數名稱 (不能有空格)。")
                            st.stop()

                    progress_bar.progress((i + 1) / total)
                
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