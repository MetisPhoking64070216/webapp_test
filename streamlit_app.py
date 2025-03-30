import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
import time

def process_excel(before_file_path, template_file_path):
    # อ่านข้อมูลจากไฟล์ Excel
    df = pd.read_excel(before_file_path, skiprows=6)

    if not os.path.exists(template_file_path) or os.path.getsize(template_file_path) == 0:
        st.error("Template file not found or empty!")
        return None

    try:
        # เปิดไฟล์ template โดยใช้ openpyxl
        wb = openpyxl.load_workbook(template_file_path)
        template_sheet = wb.active  # ใช้ active แทน
    except Exception as e:
        st.error(f"Error opening template file: {e}")
        return None

    for _, row in df.iterrows():
        store_code = int(row.iloc[1]) if not pd.isna(row.iloc[1]) else 0
        sheet_name = str(store_code)[:31]  # จำกัดชื่อแผ่นงานไม่เกิน 31 ตัวอักษร

        # ตรวจสอบว่าแผ่นงานมีอยู่หรือไม่ ถ้าไม่ก็ทำการคัดลอก
        if sheet_name not in wb.sheetnames:
            new_sheet = wb.copy_worksheet(template_sheet)
            new_sheet.title = sheet_name
        else:
            new_sheet = wb[sheet_name]
        
        # ทำการกรอกข้อมูลหรือปรับแต่งข้อมูลที่จำเป็นในแผ่นงานนี้
        # ในกรณีนี้สามารถเพิ่มข้อมูลจาก df ได้ตามต้องการ
        new_sheet["A1"] = store_code  # ตัวอย่างการเพิ่มข้อมูลในเซลล์

    # บันทึกไฟล์ใหม่
    output_path = os.path.join(tempfile.gettempdir(), "all_stores.xlsx")
    try:
        wb.save(output_path)
    except Exception as e:
        st.error(f"Error saving the output file: {e}")
        return None
    
    return output_path

st.title("📊 Excel Processing Web App")

before_file = st.file_uploader("Upload before.xlsx", type=["xlsx"])
template_file = st.file_uploader("Upload template_withoutbc.xlsx", type=["xlsx"])

if st.button("Generate Excel File"):
    if before_file and template_file:
        temp_dir = tempfile.gettempdir()
        
        before_file_path = os.path.join(temp_dir, "before.xlsx")
        with open(before_file_path, "wb") as f:
            f.write(before_file.getbuffer())
        
        template_file_path = os.path.join(temp_dir, "template.xlsx")
        with open(template_file_path, "wb") as f:
            f.write(template_file.getbuffer())
        
        time.sleep(1)  # รอให้ไฟล์ถูกเขียนก่อนเปิด
        
        if os.path.exists(template_file_path) and os.path.getsize(template_file_path) > 0:
            output_file = process_excel(before_file_path, template_file_path)
            if output_file:
                with open(output_file, "rb") as file:
                    st.download_button("Download Processed Excel", file, file_name="all_stores.xlsx")
        else:
            st.error("Template file was not saved correctly!")
    else:
        st.error("Please upload both Excel files!")
