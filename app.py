import streamlit as st
import pandas as pd
from fillpdf import fillpdfs
import os
import zipfile
import io
import shutil
from datetime import datetime
from pypdf import PdfWriter

# ฟังก์ชันแปลงเลขเป็นตัวอักษรไทย (เหมือนเดิม)
def number_to_thai_text(num):
    # กรณี NaN / None / pd.NA → ว่าง
    if pd.isna(num):
        return ""
    
    # แปลงเป็น float ก่อน (ถ้าเป็น string เช่น '0' หรือ '0.00')
    try:
        num = float(num)
    except (ValueError, TypeError):
        return ""  # ถ้าแปลงไม่ได้ → ว่าง

    # ตอนนี้ num เป็น float แน่นอน
    if num == 0:
        return "ศูนย์บาทถ้วน"   # หรือ "ศูนย์บาทถ้วน" ถ้าต้องการ

    # ส่วนที่เหลือเหมือนเดิม (สำหรับค่ามากกว่า 0)
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
    else:         text += "ถ้วน"

    text = text.replace("หนึ่งสิบ", "สิบ")

    return text

# ────────────────────────────────────────────────
# Streamlit App
# ────────────────────────────────────────────────
st.set_page_config(page_title="สร้าง ภ.ง.ด.50 ทวิ ปี 2568", layout="wide")

st.title("เครื่องมือสร้าง ภ.ง.ด.50 ทวิ ปี 2568")
st.markdown("อัปโหลดไฟล์ Excel → เลือกชีท → กดสร้าง → ดาวน์โหลด PDF รวมทุกหน้า")

# อัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("เลือกไฟล์ Excel (.xlsx)", type=["xlsx"])
with col2:
    template_file = st.file_uploader("เลือก Template PDF (optional)", type=["pdf"])

# จัดการ template
if template_file is not None:
    template_path = "temp_template.pdf"
    with open(template_path, "wb") as f:
        f.write(template_file.getbuffer())
else:
    template_path = "50ทวิ68(1).pdf"  # ต้องมีไฟล์นี้ใน directory หรือจะ error

if excel_file is not None:
    try:
        excel = pd.ExcelFile(excel_file)
        sheet_names = excel.sheet_names
        
        st.success(f"ไฟล์ Excel พร้อมใช้งาน: {excel_file.name}")
        st.info(f"ชีทที่มี: {', '.join(sheet_names)}")
        
        default_index = sheet_names.index("MasterSheet") if "MasterSheet" in sheet_names else 0
        selected_sheet = st.selectbox("เลือกชีทข้อมูลหลัก", options=sheet_names, index=default_index)
        
        if st.button("เริ่มสร้าง PDF รวมทุกหน้า", type="primary", use_container_width=True):
            with st.spinner(f"กำลังอ่านชีท '{selected_sheet}'..."):
                df = pd.read_excel(excel_file, sheet_name=selected_sheet)
            
            st.success(f"อ่านข้อมูลสำเร็จ: {len(df)} แถว")
            
            output_folder = "temp_output_50tawi"
            os.makedirs(output_folder, exist_ok=True)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            success_count = 0
            error_count = 0
            created_files = []  # สำคัญ! อยู่ตรงนี้ นอก loop
            
            PAYER_TIN   = "0-9940-00392-50-8"
            PAYER_NAME  = "เทศบาลเมืองบ้านไผ่"
            PAYER_ADDR  = "905 หมู่ 3, ถนนเจนจบทิศ, ตำบลในเมือง อำเภอบ้านไผ่ จังหวัดขอนแก่น 40110"
            PAYER_ADDR2 = "สำนักงานเทศบาลเมืองบ้านไผ่"
            year = "2568"
            
            for index, row in df.iterrows():
                try:
                    recipient_number = str(row.get('ลำดับที่', '')).strip()
                    recipient_name = str(row.get('ชื่อ-สกุล', '')).strip()
                    recipient_tin_raw = str(row.get('เลขบัตรประจำตัวประชาชน', '')).strip()
                    
                    if recipient_tin_raw.isdigit():
                        tin = recipient_tin_raw.zfill(13)
                        recipient_tin = f"{tin[0]}-{tin[1:5]}-{tin[5:10]}-{tin[10:12]}-{tin[12]}"
                    else:
                        recipient_tin = "0-0000-00000-00-0"
                    
                    value = row.get('รวม', '')
                    recipient_pay = f"{float(value):,.2f}" if value != '' and pd.notna(value) else ''
                    recipient_pay2 = row.get('ประกันสังคม', '')
                    
                    # ดึงค่า raw
                    tax_raw = row.get('ภาษี', pd.NA)

                    recipient_tax = '0.00'
                    tax_thai = ""

                    if pd.notna(tax_raw):
                        try:
                            tax_float = float(tax_raw)          # แปลงก่อน
                            recipient_tax = f"{tax_float:,.2f}"
                            tax_thai = number_to_thai_text(tax_float)  # ส่ง float ไปเลย
                        except (ValueError, TypeError):
                            # ถ้าแปลงไม่ได้ (เช่น string แปลก ๆ) → ปล่อยว่างหรือ '0.00'
                            pass
                    
                    chk_values = {'chk1': 'Yes'}
                    
                    data_dict = {
                        'id1': PAYER_TIN,
                        'name1': PAYER_NAME,
                        'add1': PAYER_ADDR,
                        'add2': PAYER_ADDR2,
                        'id1_2': recipient_tin,
                        'name2': recipient_name,
                        'book_no': "1",
                        'run_no': recipient_number,
                        'item': recipient_number,
                        **chk_values,
                        'date1': year,
                        'pay1.0': recipient_pay,
                        'tax1.0': recipient_tax,
                        'pay1.14': recipient_pay,
                        'tax1.14': recipient_tax,
                        'pragan' : recipient_pay2,
                        'total': tax_thai,
                    }
                    
                    temp_filename = f"temp_page_{index+1:03d}.pdf"
                    output_path = os.path.join(output_folder, temp_filename)

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
                    continue

                progress_bar.progress((index + 1) / len(df))
                status_text.text(f"ประมวลผล {index+1}/{len(df)} | สำเร็จ {success_count} | ผิดพลาด {error_count}")
            
            # ── Merge เป็น PDF เดียวหลัง loop เสร็จ ───────────────────────────
            if success_count > 0 and created_files:
                final_pdf_path = os.path.join(output_folder, f"50ทวิ_ทั้งหมด_{year}.pdf")
                writer = PdfWriter()
                
                for pdf_path in created_files:
                    writer.append(pdf_path)
                
                writer.write(final_pdf_path)
                writer.close()
                
                # อ่านไฟล์เพื่อ download
                with open(final_pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                
                st.success(f"สร้าง PDF รวมสำเร็จ {success_count} หน้า (ผิดพลาด {error_count} รายการ)")
                
                st.download_button(
                    label="ดาวน์โหลด PDF รวมทุกหน้า",
                    data=pdf_bytes,
                    file_name=f"50ทวิ_รวมทั้งหมด_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.error("ไม่สามารถสร้าง PDF ได้เลย กรุณาตรวจสอบข้อมูล")
            
            # Cleanup
            shutil.rmtree(output_folder, ignore_errors=True)
            if template_file is not None and os.path.exists(template_path):
                os.remove(template_path)
    
    except Exception as e:
        st.error(f"ปัญหาการอ่านไฟล์ Excel: {str(e)}")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel (.xlsx) ก่อน")

st.markdown("---")
st.caption("พัฒนาโดยใช้ Streamlit + fillpdf + pypdf | รองรับ PDF รวมหลายหน้า")