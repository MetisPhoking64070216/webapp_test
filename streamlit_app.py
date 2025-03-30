import streamlit as st
import pandas as pd
import xlwings as xw
import os
import tempfile
import time

def process_excel(before_file_path, template_file_path):
    df = pd.read_excel(before_file_path, skiprows=6)
    
    if not os.path.exists(template_file_path) or os.path.getsize(template_file_path) == 0:
        st.error("Template file not found or empty!")
        return None
    
    app = xw.App(visible=False)
    try:
        wb = app.books.open(template_file_path, update_links=False, read_only=False)
        template_sheet = wb.sheets[0]
        
        for _, row in df.iterrows():
            store_code = int(row.iloc[1]) if not pd.isna(row.iloc[1]) else 0
            sheet_name = str(store_code)[:31]
            
            if sheet_name not in [s.name for s in wb.sheets]:
                new_sheet = template_sheet.copy(after=wb.sheets[wb.sheets.count - 1])
                new_sheet.name = sheet_name
            else:
                new_sheet = wb.sheets[sheet_name]
            
            # เอาส่วน barcode ออก
            # barcode_path_jpg = os.path.join(barcode_folder, f"{store_code}.jpg")
            # barcode_path_png = os.path.join(barcode_folder, f"{store_code}.png")
            # barcode_path = barcode_path_png if os.path.exists(barcode_path_png) else barcode_path_jpg
            
            # print(f"Checking barcode file: {barcode_path}")
            # if os.path.exists(barcode_path) and os.path.getsize(barcode_path) > 0:
            #     target_cell = new_sheet.range("C2")
            #     left = target_cell.left + (target_cell.width - barcode_width) / 2
            #     top = target_cell.top + (target_cell.height - barcode_height) / 2
            #     new_sheet.pictures.add(barcode_path, left=left, top=top, width=barcode_width, height=barcode_height)
            # else:
            #     print(f"❌ Barcode not found for {store_code}")
        
        output_path = os.path.join(tempfile.gettempdir(), "all_stores.xlsx")
        wb.save(output_path)
        return output_path
    finally:
        wb.close()
        app.quit()

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
