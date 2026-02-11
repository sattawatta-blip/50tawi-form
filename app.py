import streamlit as st
import pandas as pd
from fillpdf import fillpdfs
import os
import zipfile
import io
import shutil
from datetime import datetime

# ────────────────────────────────────────────────
# ฟังก์ชันแปลงเลขเป็นตัวอักษรไทย (จากของคุณ ไม่มี "ถ้วน")
# ────────────────────────────────────────────────
def number_to_thai_text(num):
    if not num or pd.isna(num):
        return ""
    
    try:
        num = float(num)
    except:
        return ""

    if num == 0:
        return "ศูนย์บาท"

    units = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"]
    teens = ["สิบ", "สิบเอ็ด", "สิบสอง", "สิบสาม", "สิบสี่", "สิบห้า", "สิบหก", "สิบเจ็ด", "สิบแปด", "สิบเก้า"]
    ones = ["", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]

    def convert_integer(n):
        if n == 0:
            return ""
        s = str(int(n))
        result = []
        for i, digit in enumerate(reversed(s)):
            d = int(digit)
            if d == 0:
                continue
            if i == 0 and d == 1 and len(s) > 1:
                result.append("เอ็ด")
            elif i == 1 and d == 1:
                result.append(teens[0])
            elif i == 1 and d > 1:
                result.append(ones[d] + teens[0])
            else:
                result.append(ones[d] + units[i])
        return "".join(reversed(result))

    integer_part = int(num)
    decimal_part = round((num - integer_part) * 100)

    text = convert_integer(integer_part) + "บาท"

    if decimal_part > 0:
        text += convert_integer(decimal_part) + "สตางค์"

    text = text.replace("หนึ่งสิบ", "สิบ")

    return text

# ────────────────────────────────────────────────
# Streamlit App
# ────────────────────────────────────────────────
st.set_page_config(page_title="สร้าง ภ.ง.ด.50 ทวิ ปี 2568", layout="wide")

st.title("เครื่องมือสร้าง ภ.ง.ด.50 ทวิ ปี 2568")
st.markdown("อัปโหลดไฟล์ Excel → เลือกชีท → กดสร้าง → ดาวน์โหลด ZIP")

# ── อัปโหลดไฟล์ ───────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("เลือกไฟล์ Excel (.xlsx)", type=["xlsx"])
with col2:
    template_file = st.file_uploader("เลือก Template PDF (optional)", type=["pdf"])

# จัดการ template path
if template_file is not None:
    template_path = "temp_template.pdf"
    with open(template_path, "wb") as f:
        f.write(template_file.getbuffer())
else:
    template_path = "50ทวิ68(1).pdf"  # default ถ้ามีใน folder

# ── เลือก sheet ────────────────────────────────────────
selected_sheet = None
df = None

if excel_file is not None:
    try:
        excel = pd.ExcelFile(excel_file)
        sheet_names = excel.sheet_names
        
        st.success(f"ไฟล์ Excel พร้อมใช้งาน: {excel_file.name} ({excel_file.size / 1024:.1f} KB)")
        st.info(f"พบชีททั้งหมด: {', '.join(sheet_names)}")
        
        default_index = 0
        if "MasterSheet" in sheet_names:
            default_index = sheet_names.index("MasterSheet")
        
        selected_sheet = st.selectbox(
            "เลือกชีทข้อมูลหลัก",
            options=sheet_names,
            index=default_index
        )
        
        if st.button("เริ่มสร้าง PDF ทั้งหมด", type="primary", use_container_width=True):
            with st.spinner(f"กำลังอ่านชีท '{selected_sheet}'..."):
                df = pd.read_excel(excel_file, sheet_name=selected_sheet)
            
            st.success(f"อ่านข้อมูลสำเร็จ: {len(df)} แถว จากชีท '{selected_sheet}'")
            
            # ── สร้าง output folder ชั่วคราว ───────────────────────────
            output_folder = "temp_output_50tawi"
            os.makedirs(output_folder, exist_ok=True)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            success_count = 0
            error_count = 0
            created_files = []
            
            # ข้อมูลผู้จ่าย (คงเดิม)
            PAYER_TIN   = "0-9940-09392-50-8"
            PAYER_NAME  = "เทศบาลเมืองบ้านไผ่"
            PAYER_ADDR  = "905 หมู่ 3, ถนนเจนจบทิศ, ตำบลในเมือง อำเภอบ้านไผ่ จังหวัดขอนแก่น 40110"
            PAYER_ADDR2 = "สำนักงานเทศบาลเมืองบ้านไผ่"
            
            year = "2568"
            
            for index, row in df.iterrows():
                try:
                    recipient_name = str(row.get('ชื่อ-สกุล', '')).strip()
                    recipient_tin_raw = str(row.get('เลขบัตรประจำตัวประชาชน', '')).strip()
                    
                    if recipient_tin_raw.isdigit():
                        tin = recipient_tin_raw.zfill(13)
                        recipient_tin = f"{tin[0]}-{tin[1:5]}-{tin[5:10]}-{tin[10:12]}-{tin[12]}"
                    else:
                        recipient_tin = "0-0000-00000-00-0"
                    
                    value = row.get('รวม', '')
                    recipient_pay = f"{float(value):,.2f}" if value != '' and pd.notna(value) else ''
                    
                    recipient_tax = str(row.get('ภาษี', '')).strip()
                    tax_thai = number_to_thai_text(recipient_tax)
                    
                    chk_values = {'chk1': 'Yes'}
                    
                    data_dict = {
                        'id1': PAYER_TIN,
                        'name1': PAYER_NAME,
                        'add1': PAYER_ADDR,
                        'add2': PAYER_ADDR2,
                        
                        'id1_2': recipient_tin,
                        'name2': recipient_name,
                        
                        'book_no': "1",
                        'run_no': f"{index+1:03d}",
                        'item': str(index+1),
                        
                        **chk_values,
                        
                        'date2': year,
                        'pay1.1': recipient_pay,
                        'tax1.1': recipient_tax,
                        'pay1.14': recipient_pay,
                        'tax1.14': recipient_tax,
                        'total': tax_thai,
                    }
                    
                    output_filename = f"50ทวิ_{recipient_name}_{year}.pdf"
                    output_path = os.path.join(output_folder, output_filename)
                    
                    fillpdfs.write_fillable_pdf(
                        template_path,
                        output_path,
                        data_dict,
                        flatten=True
                    )
                    
                    created_files.append(output_path)
                    success_count += 1
                
                except Exception as e:
                    error_count += 1
                    st.warning(f"แถว {index+1} ({recipient_name}): {str(e)}")
                
                progress_bar.progress((index + 1) / len(df))
                status_text.text(f"ประมวลผล {index+1}/{len(df)} | สำเร็จ {success_count} | ผิดพลาด {error_count}")
            
            # ── สรุปผล + ดาวน์โหลด ZIP ───────────────────────────────
            if success_count > 0:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for file_path in created_files:
                        zipf.write(file_path, os.path.basename(file_path))
                
                zip_buffer.seek(0)
                
                st.success(f"สร้างสำเร็จ {success_count} ไฟล์ (ผิดพลาด {error_count} รายการ)")
                
                st.download_button(
                    label="ดาวน์โหลด ZIP ไฟล์ PDF ทั้งหมด",
                    data=zip_buffer,
                    file_name=f"50ทวิทั้งหมด_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.error("ไม่สามารถสร้าง PDF ได้เลย กรุณาตรวจสอบข้อมูลในชีท")
            
            # Cleanup
            shutil.rmtree(output_folder, ignore_errors=True)
            if template_file is not None and os.path.exists(template_path):
                os.remove(template_path)
    
    except Exception as e:
        st.error(f"ปัญหาการอ่านไฟล์ Excel: {str(e)}")
        st.info("แนะนำ: ลองบันทึกไฟล์ Excel เป็น .xlsx ใหม่ แล้วอัปโหลดซ้ำ")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel (.xlsx) ก่อน")

st.markdown("---")
st.caption("พัฒนาโดยใช้ Streamlit + fillpdf | ปรับปรุงล่าสุดเพื่อรองรับการเลือกชีทและ ZIP download")