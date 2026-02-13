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
from io import StringIO

# --- Windows COM 設定 ---
if os.name == 'nt':
    import pythoncom
    import win32com.client

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator", layout="wide")

st.title("⚡ HKIE CPD 證書生成器")

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

def parse_zoom_report(file_obj):
    """Parse Zoom attendee report (CSV or Excel).
    Handles multi-section format, trailing commas, and multiple join/leave entries.
    Returns (DataFrame, error_msg). error_msg is None on success."""
    is_csv = hasattr(file_obj, 'name') and file_obj.name.endswith('.csv')

    if hasattr(file_obj, 'seek'):
        file_obj.seek(0)

    if is_csv:
        # Read raw content for manual section detection
        if hasattr(file_obj, 'read'):
            content = file_obj.read()
            if isinstance(content, bytes):
                content = content.decode('utf-8-sig', errors='ignore')
        else:
            with open(file_obj, 'r', encoding='utf-8-sig', errors='ignore') as f:
                content = f.read()

        lines = content.split('\n')

        # Locate the "Attendee Details" section header row
        header_idx = -1
        for i, line in enumerate(lines):
            if 'Attendee Details' in line:
                for j in range(i + 1, min(i + 5, len(lines))):
                    if 'User Name' in lines[j] and 'Email' in lines[j]:
                        header_idx = j
                        break
                if header_idx == -1 and i + 1 < len(lines):
                    header_idx = i + 1
                break

        if header_idx == -1:
            # Fallback: find any row with expected columns
            for i, line in enumerate(lines):
                if 'User Name' in line and 'Email' in line and 'Join Time' in line:
                    header_idx = i
                    break

        if header_idx == -1:
            return None, "Cannot find Attendee Details section in Zoom report"

        # Extract header + data lines; strip trailing commas that create ghost columns
        attendee_lines = []
        for i in range(header_idx, len(lines)):
            stripped = lines[i].strip()
            if not stripped:
                continue
            if stripped.endswith(','):
                stripped = stripped[:-1]
            attendee_lines.append(stripped)

        if len(attendee_lines) < 2:
            return None, "No attendee data found after header"

        csv_text = '\n'.join(attendee_lines)
        df = pd.read_csv(StringIO(csv_text), skipinitialspace=True)
    else:
        # Excel: scan for the header row containing expected columns
        df_raw = pd.read_excel(file_obj, header=None)
        header_idx = 0
        for i, row in df_raw.iterrows():
            row_vals = [str(v) for v in row.values]
            if any('User Name' in v for v in row_vals) and any('Email' in v for v in row_vals):
                header_idx = i
                break
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_idx)

    # Drop completely empty columns (artefact of trailing commas)
    df = df.dropna(axis=1, how='all')
    return df, None

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
            
            # --- 強化的欄位對應邏輯 ---
            col_map = {}
            has_full_name = False
            
            for c in df_reg.columns:
                c_lower = str(c).lower().strip()
                if 'full name' in c_lower:
                    col_map[c] = 'Full Name'
                    has_full_name = True
                elif 'first name' in c_lower or '名字' in c_lower:
                    col_map[c] = 'First Name'
                elif 'last name' in c_lower or '姓氏' in c_lower:
                    col_map[c] = 'Last Name'
                elif 'contact email' in c_lower:
                    col_map[c] = 'Email'
                elif 'email' in c_lower or '電郵' in c_lower:
                    col_map[c] = 'Email'
                elif 'hkie membership' in c_lower or 'membership no' in c_lower or '會員編號' in c_lower:
                    col_map[c] = 'Membership No'
                elif 'salutation' in c_lower or '稱呼' in c_lower:
                    col_map[c] = 'Salutation'
            
            df_reg.rename(columns=col_map, inplace=True)
            
            # 如果有 Full Name 但沒有 First Name / Last Name，需要拆分
            if has_full_name and 'Full Name' in df_reg.columns:
                if 'First Name' not in df_reg.columns or 'Last Name' not in df_reg.columns:
                    # 拆分 Full Name 為 First Name 和 Last Name
                    name_split = df_reg['Full Name'].astype(str).str.strip().str.split(n=1, expand=True)
                    if name_split.shape[1] == 2:
                        df_reg['First Name'] = name_split[0]
                        df_reg['Last Name'] = name_split[1]
                    else:
                        # 如果只有一個詞，全部當作 First Name
                        df_reg['First Name'] = name_split[0]
                        df_reg['Last Name'] = ""
            
            # 檢查是否成功抓到 Membership No
            if 'Membership No' not in df_reg.columns:
                st.warning("⚠️ 警告：無法自動識別 'Membership No' 欄位。這可能導致證書上的會員編號為空白。請檢查 Excel 標題是否包含 'Membership' 或 '會員編號'。")
                # 嘗試建立一個空的欄位以防報錯
                df_reg['Membership No'] = ""
            
            # 如果沒有 Salutation 欄位，建立一個空的
            if 'Salutation' not in df_reg.columns:
                df_reg['Salutation'] = ""
            
            required_cols = ['First Name', 'Last Name', 'Email']
            if not all(col in df_reg.columns for col in required_cols):
                st.error(f"報名表缺少必要欄位: {required_cols}")
                st.write("目前偵測到的欄位:", df_reg.columns.tolist())
                st.stop()

            # B. 核對 Zoom
            if not use_zoom:
                df_final = df_reg.copy()
                df_final['Full Name'] = df_final['First Name'].astype(str) + " " + df_final['Last Name'].astype(str)
                df_final['Match Method'] = "Registration Only"
            else:
                # Parse Zoom attendee report
                df_zoom, zoom_err = parse_zoom_report(zoom_file)
                if zoom_err:
                    st.error(f"Zoom 檔案解析失敗: {zoom_err}")
                    st.stop()

                st.write(f"📊 Zoom 檔案欄位: {df_zoom.columns.tolist()}")
                st.write(f"📈 Zoom 原始資料筆數: {len(df_zoom)}")

                z_user_col = next((c for c in df_zoom.columns if "User Name" in str(c)), None)
                z_email_col = next((c for c in df_zoom.columns if "Email" in str(c)), None)

                if not z_user_col or not z_email_col:
                    st.error("Zoom 檔案無法識別 User Name 或 Email 欄位。")
                    st.write("偵測到的欄位:", df_zoom.columns.tolist())
                    st.stop()

                # Aggregate total time in session per attendee (handles re-joins)
                z_time_col = next((c for c in df_zoom.columns if "Time in Session" in str(c)), None)
                if z_time_col:
                    df_zoom[z_time_col] = pd.to_numeric(df_zoom[z_time_col], errors='coerce').fillna(0)
                    time_agg = df_zoom.groupby(z_email_col, as_index=False)[z_time_col].sum()
                    time_agg.rename(columns={z_time_col: '_total_time'}, inplace=True)
                    df_zoom = df_zoom.drop_duplicates(subset=[z_email_col], keep='first')
                    df_zoom = df_zoom.merge(time_agg, on=z_email_col, how='left')
                    df_zoom[z_time_col] = df_zoom['_total_time']
                    df_zoom.drop(columns=['_total_time'], inplace=True)
                else:
                    df_zoom = df_zoom.drop_duplicates(subset=[z_email_col], keep='first')

                st.write(f"✅ 去除重複後 (已合計出席時間): {len(df_zoom)} 位出席者")

                st.write("正在核對 Zoom 資料...")
                df_reg['Name_Norm'] = (df_reg['First Name'].astype(str) + " " + df_reg['Last Name'].astype(str)).apply(normalize_name)
                df_reg['Email_Norm'] = df_reg['Email'].astype(str).str.lower().str.strip()
                
                df_zoom['Name_Norm'] = df_zoom[z_user_col].apply(normalize_name)
                df_zoom['Email_Norm'] = df_zoom[z_email_col].astype(str).str.lower().str.strip()
                
                # 建立 Zoom 對應字典
                zoom_email_map = df_zoom.set_index('Email_Norm')[z_user_col].to_dict()
                zoom_name_map = df_zoom.set_index('Name_Norm')[z_user_col].to_dict()
                
                st.write(f"📧 Zoom Email 數量: {len(zoom_email_map)}")
                st.write(f"👤 Zoom Name 數量: {len(zoom_name_map)}")
                st.write(f"📝 報名表數量: {len(df_reg)}")
                
                # 顯示前幾筆 Zoom 資料供檢查
                st.write("Zoom 資料預覽 (前5筆):")
                st.dataframe(df_zoom[[z_user_col, z_email_col, 'Email_Norm']].head())
                
                matched_list = []
                unmatched_list = []
                
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
                    else:
                        unmatched_list.append({
                            "Name": f"{row.get('First Name', '')} {row.get('Last Name', '')}",
                            "Email": row.get('Email', ''),
                            "Email_Norm": row['Email_Norm'],
                            "Name_Norm": row['Name_Norm']
                        })
                
                df_final = pd.DataFrame(matched_list)
                
                # 顯示未匹配的記錄
                if unmatched_list:
                    st.warning(f"⚠️ {len(unmatched_list)} 筆報名記錄未在 Zoom 中找到")
                    with st.expander("查看未匹配的記錄"):
                        st.dataframe(pd.DataFrame(unmatched_list))

            if not df_final.empty:
                st.success(f"共產生 {len(df_final)} 筆證書名單。")
                # 顯示前幾筆資料供檢查
                st.write("預覽將生成的資料 (請確認 Membership No 是否有值):")
                st.dataframe(df_final[['Salutation', 'Full Name', 'Membership No', 'Email']].head())
            else:
                st.warning("沒有符合的名單。")

        except Exception as e:
            st.error(f"資料處理發生錯誤: {e}")

    # --- 4. 生成與下載 ---
    st.header("4. 生成證書")
    
    output_format = st.radio(
        "選擇輸出格式：",
        ('Word 文件 (.docx) - 不加密', 'PDF 文件 (.pdf) - 加密 (密碼: Membership No.)')
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
                
                # PDF 批次處理初始化
                word = None
                if output_format.startswith('PDF') and os.name == 'nt':
                    try:
                        pythoncom.CoInitialize()
                        word = win32com.client.DispatchEx("Word.Application")
                        word.Visible = False
                        word.DisplayAlerts = False
                    except Exception as e:
                        st.error(f"無法啟動 Word: {e}")
                        st.stop()
                
                try:
                    for i, person in df_final.iterrows():
                        person_name = str(person['Full Name']).strip()
                        status_text.text(f"處理中 ({i+1}/{total}): {person_name}")
                        
                        try:
                            # 1. 產生 DOCX
                            doc_tpl = DocxTemplate(template_path)
                            
                            # 處理 Membership No (避免 NaN 或 .0)
                            mem_no = str(person['Membership No'])
                            if mem_no.lower() in ['nan', 'none', '']: 
                                mem_no = ""
                            if mem_no.endswith('.0'): # 去除 Excel 數字轉字串可能出現的 .0
                                mem_no = mem_no[:-2]

                            # 建立變數對應 (Context)
                            # 注意：這裡使用 membership_no 對應新範本
                            context = {
                                'name': f"{person['Salutation']} {person_name}",
                                'membership_no': mem_no,  # 對應 Word 中的 {{ membership_no }}
                                'event_title': event_title,
                                'event_details': event_details
                            }
                            doc_tpl.render(context)
                            
                            safe_name = re.sub(r'[\\/*?:"<>|]', "", person_name)
                            docx_filename = f"{safe_name}.docx"
                            docx_path = os.path.join(tmpdirname, docx_filename)
                            doc_tpl.save(docx_path)
                            
                            final_file_path = docx_path
                            
                            # 2. 轉 PDF (若需要)
                            if word:
                                try:
                                    pdf_filename = f"{safe_name}.pdf"
                                    pdf_path = os.path.join(tmpdirname, pdf_filename)
                                    
                                    wb_doc = word.Documents.Open(os.path.abspath(docx_path))
                                    wb_doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
                                    wb_doc.Close(SaveChanges=False)
                                    
                                    password = str(person['Membership No']).strip()
                                    if not password or password == 'nan':
                                        password = "hkie"
                                        
                                    encrypted_path = os.path.join(tmpdirname, f"Encrypted_{safe_name}.pdf")
                                    with pikepdf.Pdf.open(pdf_path) as pdf:
                                        pdf.save(encrypted_path, encryption=pikepdf.Encryption(owner=password, user=password, R=6))
                                    
                                    final_file_path = encrypted_path
                                except Exception as e:
                                    # st.warning(f"{person_name} 轉檔失敗: {e}")
                                    final_file_path = docx_path
                            
                            generated_files.append(final_file_path)
                            success_count += 1
                            
                        except Exception as e:
                            st.error(f"生成 {person_name} 時錯誤: {e}")
                            if "expected token" in str(e):
                                st.stop()

                        progress_bar.progress((i + 1) / total)
                        
                finally:
                    if word:
                        try:
                            word.Quit()
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