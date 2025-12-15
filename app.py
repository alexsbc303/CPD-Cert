# pip install streamlit pandas beautifulsoup4 requests docxtpl docx2pdf pikepdf openpyxl

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
from datetime import datetime

# --- 設定頁面 ---
st.set_page_config(page_title="CPD Cert Generator", layout="wide")

st.title("🎓 HKIE CPD 證書生成小幫手")
st.markdown("此工具協助你從網站抓取活動資訊，核對出席者，並自動生成加密的 PDF 證書。")

# --- 1. 爬蟲功能：獲取活動資訊 ---
st.header("1. 獲取活動資訊")
url = st.text_input("輸入 HKIE 活動網址", "http://it.hkie.org.hk/en_it_events_inside_Past.aspx?EventID=600&&TypeName=Events+%2f+Activities")

if 'event_info' not in st.session_state:
    st.session_state['event_info'] = {}

if st.button("抓取活動資訊"):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 根據網頁結構尋找特定欄位 (需根據實際網頁調整選擇器)
        # 這裡使用簡單的關鍵字搜尋作為範例
        text_content = soup.get_text()
        
        # 簡易提取邏輯 (實際專案可針對 HTML 結構優化)
        title = "未能自動抓取，請手動輸入"
        date_str = ""
        time_str = ""
        
        # 嘗試尋找標題 (假設在特定的 header 或透過關鍵字截取)
        # 這裡為了示範，我們先讓使用者確認抓到的原始文字，或手動填寫
        st.session_state['event_info']['raw_text'] = text_content[:500] # 預覽
        
        st.success("網頁讀取成功！請在下方確認資訊。")
        
    except Exception as e:
        st.error(f"抓取失敗: {e}")

# 讓使用者確認或手動修改抓取到的資訊
col1, col2 = st.columns(2)
with col1:
    event_title = st.text_input("活動標題 (Event Title)", "Technical Seminar: Embodied Intelligence")
with col2:
    event_details = st.text_input("日期與時間 (Date & Time)", "4 Dec (Thu) 17:00-18:00")

# --- 2. 上傳檔案 ---
st.header("2. 上傳資料檔")
reg_file = st.file_uploader("上傳報名表 (Excel 1 - Registration)", type=['csv', 'xlsx'])
zoom_file = st.file_uploader("上傳 Zoom 報告 (Excel 2 - Attendee)", type=['csv', 'xlsx'])
template_file = st.file_uploader("上傳證書範本 (Word - .docx)", type=['docx'])

# --- 輔助函式：姓名標準化 ---
def normalize_name(name):
    if pd.isna(name): return ""
    name = str(name).lower()
    # 去除稱謂
    name = re.sub(r'\b(ir|mr|ms|miss|dr|prof)\b\.?', '', name)
    # 只保留英文字母和空格
    name = re.sub(r'[^a-z\s]', '', name)
    return " ".join(name.split())

# --- 3. 數據處理與核對 ---
if reg_file and zoom_file and template_file:
    st.header("3. 核對出席者")
    
    # 讀取報名表
    try:
        if reg_file.name.endswith('.csv'):
            df_reg = pd.read_csv(reg_file)
        else:
            df_reg = pd.read_excel(reg_file)
            
        # 欄位對應 (根據你的檔案)
        # 假設欄位名稱如下，若不同需調整
        col_map = {
            'First Name 名字': 'First Name', 
            'Last Name 姓氏': 'Last Name', 
            'Email Address 電郵地址': 'Email',
            'Membership No. 會員編號 (If Any, 如有)': 'Membership No',
            'Salutation 稱呼': 'Salutation'
        }
        df_reg.rename(columns=col_map, inplace=True)
        
        # 讀取 Zoom 報告 (處理 Header 在第 4-5 行的情況)
        # 這裡用一種比較聰明的方式找 Header
        if zoom_file.name.endswith('.csv'):
            # 先讀取前 10 行來判斷
            zoom_preview = pd.read_csv(zoom_file, header=None, nrows=10)
            header_row = 0
            for i, row in zoom_preview.iterrows():
                row_str = str(row.values)
                if "User Name" in row_str and "Email" in row_str:
                    header_row = i
                    break
            zoom_file.seek(0)
            df_zoom = pd.read_csv(zoom_file, header=header_row)
        else:
            df_zoom = pd.read_excel(zoom_file) # Excel 類似處理略
            
        # 篩選有出席的人
        if 'Attended' in df_zoom.columns:
            df_zoom = df_zoom[df_zoom['Attended'] == 'Yes']

        # --- 開始配對邏輯 ---
        st.write("正在進行配對 (優先比對 Email，其次比對標準化姓名)...")
        
        # 準備欄位
        df_reg['Name_Norm'] = (df_reg['First Name'].astype(str) + " " + df_reg['Last Name'].astype(str)).apply(normalize_name)
        df_reg['Email_Norm'] = df_reg['Email'].astype(str).str.lower().str.strip()
        
        # Zoom 欄位名稱可能不同，這裡做些容錯
        zoom_name_col = [c for c in df_zoom.columns if "User Name" in c][0]
        zoom_email_col = [c for c in df_zoom.columns if "Email" in c][0]
        
        df_zoom['Name_Norm'] = df_zoom[zoom_name_col].apply(normalize_name)
        df_zoom['Email_Norm'] = df_zoom[zoom_email_col].astype(str).str.lower().str.strip()
        
        # 建立 Zoom 查找字典
        zoom_email_map = df_zoom.set_index('Email_Norm')[zoom_name_col].to_dict()
        zoom_name_map = df_zoom.set_index('Name_Norm')[zoom_name_col].to_dict()
        
        matched_list = []
        
        for idx, row in df_reg.iterrows():
            status = "Unmatched"
            
            # 1. Email 配對
            if row['Email_Norm'] in zoom_email_map:
                status = "Matched (Email)"
            # 2. 姓名配對
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
                
        df_matched = pd.DataFrame(matched_list)
        st.success(f"配對完成！共找到 {len(df_matched)} 位出席者。")
        st.dataframe(df_matched)
        
    except Exception as e:
        st.error(f"檔案讀取或處理錯誤: {e}")

    # --- 4. 生成證書 ---
    st.header("4. 生成證書 (PDF + 加密)")
    
    if st.button("開始生成證書"):
        if len(df_matched) == 0:
            st.warning("沒有配對到的出席者，無法生成。")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 建立暫存資料夾
            with tempfile.TemporaryDirectory() as tmpdirname:
                zip_path = os.path.join(tmpdirname, "certs.zip")
                template_path = os.path.join(tmpdirname, "template.docx")
                
                # 儲存範本
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                
                generated_files = []
                
                for i, person in df_matched.iterrows():
                    status_text.text(f"正在處理: {person['Full Name']}...")
                    
                    # 1. 填寫 Word 範本
                    doc = DocxTemplate(template_path)
                    context = {
                        'name': f"{person['Salutation']} {person['Full Name']}",
                        'Membership number': str(person['Membership No']),
                        'event_title': event_title,
                        'event_details': event_details
                    }
                    doc.render(context)
                    
                    docx_filename = f"{person['Full Name']}_cert.docx"
                    docx_path = os.path.join(tmpdirname, docx_filename)
                    doc.save(docx_path)
                    
                    # 2. 轉換為 PDF (需安裝 Word)
                    # 注意：在沒有 Word 的 Server 上這步會失敗，建議在本機執行
                    try:
                        pdf_filename = f"{person['Full Name']}_cert.pdf"
                        pdf_path = os.path.join(tmpdirname, pdf_filename)
                        convert(docx_path, pdf_path) # docx2pdf
                        
                        # 3. 加密 PDF (使用會員編號)
                        password = str(person['Membership No']).strip()
                        if not password or password == 'nan':
                            password = "hkie" # 預設密碼
                            
                        encrypted_pdf_path = os.path.join(tmpdirname, f"Encrypted_{pdf_filename}")
                        
                        with pikepdf.Pdf.open(pdf_path) as pdf:
                            pdf.save(encrypted_pdf_path, encryption=pikepdf.Encryption(owner=password, user=password, R=6))
                        
                        generated_files.append(encrypted_pdf_path)
                        
                    except Exception as e:
                        # 如果 PDF 轉換失敗 (例如無 Word 環境)，我們只提供 DOCX
                        generated_files.append(docx_path)
                        # print(f"PDF Conversion failed for {person['Full Name']}: {e}")

                    progress_bar.progress((i + 1) / len(df_matched))

                # 打包成 ZIP
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in generated_files:
                        zipf.write(file, os.path.basename(file))
                
                # 下載按鈕
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="下載所有證書 (ZIP)",
                        data=f,
                        file_name="cpd_certificates.zip",
                        mime="application/zip"
                    )
            
            st.success("所有證書生成完畢！")