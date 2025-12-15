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

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator", layout="wide")

st.title("🎓 HKIE CPD 證書生成小幫手 (增強版)")
st.markdown("此工具協助你從網站抓取活動資訊，(選擇性)核對出席者，並生成加密 PDF 或原始 Word 證書。")

# --- 1. 爬蟲功能：獲取活動資訊 ---
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
            st.warning("找不到標題 ID，請手動輸入。")

        # 抓取日期時間
        dtv_tag = soup.find(id="ctl00_ContentPlaceHolder1_dtv")
        if dtv_tag:
            raw_dtv = dtv_tag.get_text(strip=True)
            formatted_dtv = raw_dtv.replace(';', ' ') 
            st.session_state['event_details'] = formatted_dtv
        else:
            st.warning("找不到日期時間 ID，請手動輸入。")
            
        st.success("網頁資訊抓取完成！")
        
    except Exception as e:
        st.error(f"抓取失敗: {e}")

col1, col2 = st.columns(2)
with col1:
    event_title = st.text_input("活動標題 (Event Title)", value=st.session_state['event_title'])
with col2:
    event_details = st.text_input("日期與時間 (Date & Time)", value=st.session_state['event_details'])

# --- 2. 上傳檔案 ---
st.header("2. 上傳資料檔")

# 必填
reg_file = st.file_uploader("上傳報名表 (Excel 1 - Registration) [必填]", type=['csv', 'xlsx'])
template_file = st.file_uploader("上傳證書範本 (Word - .docx) [必填]", type=['docx'])

# 選填：Zoom 核對
st.subheader("Zoom 核對選項")
use_zoom = st.checkbox("需要核對 Zoom 出席紀錄？", value=True)

zoom_file = None
if use_zoom:
    zoom_file = st.file_uploader("上傳 Zoom 報告 (Excel 2 - Attendee) [選填]", type=['csv', 'xlsx'])
else:
    st.info("ℹ️ 已跳過 Zoom 核對，將直接使用報名表所有名單生成證書。")

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

# --- 3. 數據處理與核對 ---
df_final = pd.DataFrame() # 用來存放最終要生成證書的名單

if reg_file and template_file:
    # 如果勾選了 Zoom 但還沒上傳，先不執行
    if use_zoom and not zoom_file:
        st.warning("請上傳 Zoom 報告以進行核對，或取消勾選「需要核對 Zoom...」選項。")
    else:
        st.header("3. 處理名單")
        try:
            # --- A. 讀取報名表 ---
            if reg_file.name.endswith('.csv'):
                df_reg = pd.read_csv(reg_file)
            else:
                df_reg = pd.read_excel(reg_file)
                
            col_map = {
                'First Name 名字': 'First Name', 
                'Last Name 姓氏': 'Last Name', 
                'Email Address 電郵地址': 'Email',
                'Membership No. 會員編號 (If Any, 如有)': 'Membership No',
                'Salutation 稱呼': 'Salutation'
            }
            df_reg.rename(columns=col_map, inplace=True)

            # 確保有 Email 欄位 (加密用)
            if 'Email' not in df_reg.columns:
                st.error("報名表中找不到 'Email' 欄位，無法進行後續加密。請檢查欄位名稱。")
                st.stop()

            # --- B. 邏輯分支：是否核對 Zoom ---
            if not use_zoom:
                # 不核對，直接用報名表
                st.info("使用全數報名者名單...")
                df_final = df_reg.copy()
                df_final['Full Name'] = df_final['First Name'].astype(str) + " " + df_final['Last Name'].astype(str)
                df_final['Match Method'] = "Registration Only"
                
            else:
                # 需要核對 Zoom
                # 1. 讀取 Zoom
                if zoom_file.name.endswith('.csv'):
                    df_preview = pd.read_csv(zoom_file, header=None, nrows=20)
                else:
                    df_preview = pd.read_excel(zoom_file, header=None, nrows=20)
                    
                header_row = find_header_row(df_preview, keywords=["User Name", "Email"])
                
                zoom_file.seek(0)
                if zoom_file.name.endswith('.csv'):
                    df_zoom = pd.read_csv(zoom_file, header=header_row)
                else:
                    df_zoom = pd.read_excel(zoom_file, header=header_row)

                user_name_candidates = [c for c in df_zoom.columns if "User Name" in str(c)]
                email_candidates = [c for c in df_zoom.columns if "Email" in str(c)]
                
                if not user_name_candidates or not email_candidates:
                    st.error("Zoom 檔案欄位錯誤。")
                    st.stop()
                    
                zoom_name_col = user_name_candidates[0]
                zoom_email_col = email_candidates[0]
                
                if 'Attended' in df_zoom.columns:
                    df_zoom = df_zoom[df_zoom['Attended'] == 'Yes']

                # 2. 配對邏輯
                st.write("正在核對 Zoom 出席紀錄...")
                
                df_reg['Name_Norm'] = (df_reg['First Name'].astype(str) + " " + df_reg['Last Name'].astype(str)).apply(normalize_name)
                df_reg['Email_Norm'] = df_reg['Email'].astype(str).str.lower().str.strip()
                
                df_zoom['Name_Norm'] = df_zoom[zoom_name_col].apply(normalize_name)
                df_zoom['Email_Norm'] = df_zoom[zoom_email_col].astype(str).str.lower().str.strip()
                
                zoom_email_map = df_zoom.set_index('Email_Norm')[zoom_name_col].to_dict()
                zoom_name_map = df_zoom.set_index('Name_Norm')[zoom_name_col].to_dict()
                
                matched_list = []
                for idx, row in df_reg.iterrows():
                    status = "Unmatched"
                    if row['Email_Norm'] in zoom_email_map:
                        status = "Matched (Email)"
                    elif row['Name_Norm'] in zoom_name_map:
                        status = "Matched (Name)"
                        
                    if "Matched" in status:
                        matched_list.append({
                            "Salutation": row.get('Salutation', ''),
                            "First Name": row.get('First Name', ''),
                            "Last Name": row.get('Last Name', ''),
                            "Full Name": f"{row.get('First Name', '')} {row.get('Last Name', '')}",
                            "Membership No": row.get('Membership No', 'N/A'),
                            "Email": row.get('Email', ''),
                            "Match Method": status
                        })
                df_final = pd.DataFrame(matched_list)

            # 顯示結果
            if len(df_final) > 0:
                st.success(f"準備生成 {len(df_final)} 份證書。")
                st.dataframe(df_final[['Salutation', 'Full Name', 'Email', 'Membership No', 'Match Method']].head())
            else:
                st.warning("名單為空，無法進行下一步。")

        except Exception as e:
            st.error(f"資料處理錯誤: {e}")
            import traceback
            st.text(traceback.format_exc())

    # --- 4. 生成證書選項 ---
    st.header("4. 生成與下載")
    
    # 新增：輸出格式選擇
    output_format = st.radio(
        "選擇輸出格式：",
        ('Word 文件 (.docx) - 不加密', 'PDF 文件 (.pdf) - 加密 (使用 Email)')
    )
    
    if st.button("開始生成"):
        if df_final.empty:
            st.error("沒有可用的名單。")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            with tempfile.TemporaryDirectory() as tmpdirname:
                zip_filename = "certs_word.zip" if output_format.startswith('Word') else "certs_pdf_encrypted.zip"
                zip_path = os.path.join(tmpdirname, zip_filename)
                template_path = os.path.join(tmpdirname, "template.docx")
                
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                
                generated_files = []
                total = len(df_final)
                
                for i, person in df_final.iterrows():
                    status_text.text(f"處理中 ({i+1}/{total}): {person['Full Name']}...")
                    
                    # 1. 填寫 Word
                    doc = DocxTemplate(template_path)
                    mem_no = str(person['Membership No'])
                    if mem_no == 'nan' or mem_no == 'None': mem_no = ""
                    
                    # 變數名稱修正
                    context = {
                        'name': f"{person['Salutation']} {person['Full Name']}",
                        'membership_number': mem_no, 
                        'event_title': event_title,
                        'event_details': event_details
                    }
                    doc.render(context)
                    
                    safe_name = re.sub(r'[\\/*?:"<>|]', "", str(person['Full Name']))
                    docx_filename = f"{safe_name}.docx"
                    docx_path = os.path.join(tmpdirname, docx_filename)
                    doc.save(docx_path)
                    
                    if output_format.startswith('Word'):
                        # 如果選 Word，直接加入列表
                        generated_files.append(docx_path)
                    else:
                        # 如果選 PDF，進行轉換與加密
                        try:
                            pdf_filename = f"{safe_name}.pdf"
                            pdf_path = os.path.join(tmpdirname, pdf_filename)
                            convert(docx_path, pdf_path)
                            
                            # 加密邏輯：使用 Email
                            password = str(person['Email']).strip()
                            if not password or password == 'nan':
                                password = "hkie" # Fallback password
                                
                            encrypted_pdf_path = os.path.join(tmpdirname, f"Encrypted_{pdf_filename}")
                            
                            with pikepdf.Pdf.open(pdf_path) as pdf:
                                pdf.save(encrypted_pdf_path, encryption=pikepdf.Encryption(owner=password, user=password, R=6))
                            
                            generated_files.append(encrypted_pdf_path)
                        except Exception as e:
                            # 轉換失敗 (通常是伺服器沒 Word)，回退為 Word
                            generated_files.append(docx_path)

                    progress_bar.progress((i + 1) / total)

                # 打包下載
                if generated_files:
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for file in generated_files:
                            zipf.write(file, os.path.basename(file))
                    
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="📥 下載證書壓縮檔 (ZIP)",
                            data=f,
                            file_name=zip_filename,
                            mime="application/zip"
                        )
                    st.success(f"完成！已生成 {len(generated_files)} 個檔案。")
                else:
                    st.error("未能生成檔案。")