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

# --- Windows COM 設定 ---
if os.name == 'nt':
    import pythoncom
    import win32com.client

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator (Fixed)", layout="wide")

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

def find_header_row(df_preview, keywords=["User Name", "Email"]):
    """找到包含所有關鍵字的標題行"""
    for i, row in df_preview.iterrows():
        row_str_list = [str(val) for val in row.values]
        if all(any(kw in cell for cell in row_str_list) for kw in keywords):
            return i
    return 0

def find_attendee_section_in_zoom(file_path_or_obj):
    """專門處理 Zoom 報告，找到 Attendee Details 區域的標題行"""
    # 讀取整個文件內容
    if hasattr(file_path_or_obj, 'seek'):
        file_path_or_obj.seek(0)
    
    if hasattr(file_path_or_obj, 'read'):
        content = file_path_or_obj.read()
        if isinstance(content, bytes):
            content = content.decode('utf-8-sig', errors='ignore')
        else:
            content = str(content)
    else:
        with open(file_path_or_obj, 'r', encoding='utf-8-sig', errors='ignore') as f:
            content = f.read()
    
    lines = content.split('\n')
    
    # 找到 "Attendee Details" 行
    attendee_section_idx = -1
    for i, line in enumerate(lines):
        if 'Attendee Details' in line:
            attendee_section_idx = i
            break
    
    if attendee_section_idx == -1:
        # 如果找不到 Attendee Details，嘗試找包含 "User Name" 和 "Email" 的行
        for i, line in enumerate(lines):
            if 'User Name' in line and 'Email' in line and 'Join Time' in line:
                return i
        return 0  # 找不到，回傳預設值
    
    # Attendee Details 的下一行應該是標題行
    # 找到包含 "User Name" 和 "Email" 的那一行
    for i in range(attendee_section_idx + 1, min(attendee_section_idx + 5, len(lines))):
        if 'User Name' in lines[i] and 'Email' in lines[i]:
            return i
    
    return attendee_section_idx + 1  # 預設為 Attendee Details 的下一行

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
                # 使用新的 Zoom 解析方法
                header_row = find_attendee_section_in_zoom(zoom_file)
                st.write(f"🔍 Zoom 檔案標題行位置: {header_row}")
                
                zoom_file.seek(0)
                if zoom_file.name.endswith('.csv'):
                    # skip_blank_lines=False 和 skipinitialspace=True 處理格式問題
                    # 先讀取看看有多少列
                    temp_df = pd.read_csv(zoom_file, header=header_row, encoding='utf-8-sig', 
                                         on_bad_lines='skip', nrows=5)
                    st.write(f"🔍 臨時讀取前5行檢查欄位: {temp_df.columns.tolist()}")
                    st.write(f"🔍 第一行資料樣本:")
                    st.dataframe(temp_df.head(1))
                    
                    # 重新讀取完整資料，使用 skipinitialspace 去除多餘空格
                    zoom_file.seek(0)
                    df_zoom = pd.read_csv(zoom_file, header=header_row, encoding='utf-8-sig', 
                                         on_bad_lines='skip', skipinitialspace=True)
                else:
                    df_zoom = pd.read_excel(zoom_file, header=header_row)
                
                st.write(f"📊 Zoom 檔案欄位: {df_zoom.columns.tolist()}")
                st.write(f"📈 Zoom 原始資料筆數: {len(df_zoom)}")
                
                # 檢查是否有欄位錯位問題 - 如果 Attended 欄位包含名字而不是 Yes/No
                if len(df_zoom) > 0:
                    first_attended = str(df_zoom['Attended'].iloc[0]) if 'Attended' in df_zoom.columns else ""
                    st.write(f"🔍 第一筆 Attended 值: '{first_attended}'")
                    # 如果 Attended 不是 Yes/No，可能有欄位錯位
                    if first_attended and first_attended.lower() not in ['yes', 'no', 'nan', '']:
                        st.warning("⚠️ 偵測到欄位可能錯位，嘗試修正...")
                        # 檢查是否有未命名的第一欄
                        if df_zoom.columns[0].startswith('Unnamed'):
                            st.write("發現未命名的第一欄，移除它")
                            df_zoom = df_zoom.iloc[:, 1:]  # 移除第一欄
                        # 或者檢查第二欄是否才是真正的 Attended
                        elif 'Attended' not in df_zoom.columns and len(df_zoom.columns) > 1:
                            # 嘗試使用第一列資料作為欄位名
                            st.write("嘗試重新解析欄位...")
                
                st.write(f"📊 修正後欄位: {df_zoom.columns.tolist()}")
                
                z_user_col = next((c for c in df_zoom.columns if "User Name" in str(c)), None)
                z_email_col = next((c for c in df_zoom.columns if "Email" in str(c)), None)
                
                if not z_user_col or not z_email_col:
                    st.error("Zoom 檔案無法識別 User Name 或 Email 欄位。")
                    st.write("偵測到的欄位:", df_zoom.columns.tolist())
                    st.stop()
                
                st.write(f"✅ User Name 欄位: {z_user_col}")
                st.write(f"✅ Email 欄位: {z_email_col}")
                
                # 不過濾 Attended，因為這個檔案本身就是 Attendee Report
                st.write(f"ℹ️ 使用所有記錄 (此檔案為出席者報告)")
                
                # 去除重複的 Email (保留第一筆)
                df_zoom = df_zoom.drop_duplicates(subset=[z_email_col], keep='first')
                st.write(f"✓ 去除重複後: {len(df_zoom)} 筆")

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