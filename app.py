import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
import requests
import re

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบรายงานผู้ป่วยจิตเวชพ้นโทษ", layout="wide")
st.title("ระบบออกรายงานการติดตามผู้ต้องขังจิตเวชพ้นโทษ (พ.ร.บ.สุขภาพจิต 2551)")

# ฟังก์ชันแปลงลิงก์ Google Drive เป็นลิงก์ดาวน์โหลดตรง และโหลดรูปชั่วคราว
def download_gdrive_image(url, save_path):
    try:
        if pd.isna(url) or url.strip() == "":
            return False
            
        # ถ้ามีหลายลิงก์ (ผู้ใช้อัปโหลดหลายรูป) ให้ดึงมารูปแรก
        first_url = url.split(',')[0].strip()
        
        file_id = None
        if 'id=' in first_url:
            file_id = first_url.split('id=')[1].split('&')[0]
        elif '/d/' in first_url:
            file_id = first_url.split('/d/')[1].split('/')[0]
            
        if file_id:
            direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            response = requests.get(direct_url, timeout=10)
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    f.write(response.content)
                return True
        return False
    except Exception as e:
        return False

# คลาสสร้างแบบฟอร์ม PDF
class PDFReport(FPDF):
    def header(self):
        # โหลดฟอนต์ (ตรวจสอบว่ามีไฟล์ฟอนต์ในโฟลเดอร์เดียวกัน)
        try:
            self.add_font("THSarabun", "", "THSarabunNew.ttf", uni=True)
            self.add_font("THSarabun-Bold", "B", "THSarabunNew_Bold.ttf", uni=True)
            font_b = "THSarabun-Bold"
        except:
            # ใช้ฟอนต์เริ่มต้นหากไม่มีไฟล์ (ภาษาไทยจะกลายเป็นสี่เหลี่ยม)
            font_b = "Arial"
            
        self.set_font("THSarabun-Bold", "B", 18)
        self.cell(0, 10, "รายงานการติดตามผู้ต้องขังจิตเวชพ้นโทษ", 0, 1, "C")
        self.set_font("THSarabun-Bold", "B", 16)
        self.cell(0, 8, "ตาม พ.ร.บ. สุขภาพจิต พ.ศ. 2551", 0, 1, "C")
        self.ln(5)

# ฟังก์ชันพิมพ์หัวข้อและรายละเอียด
def print_field(pdf, title, value):
    # ป้องกันค่าที่เป็น NaN (ค่าว่างจาก Excel)
    val_str = str(value) if pd.notna(value) else "-"
    
    pdf.set_font("THSarabun-Bold", "B", 15)
    pdf.cell(0, 8, f"{title}:", 0, 1)
    
    pdf.set_font("THSarabun", "", 15)
    # ใช้ multi_cell เพื่อให้ข้อความขึ้นบรรทัดใหม่เองอัตโนมัติหากยาวเกินขอบกระดาษ
    pdf.multi_cell(0, 7, f"    {val_str}")
    pdf.ln(2)

def generate_pdf(dataframe):
    pdf = PDFReport()
    
    # สร้างโฟลเดอร์เก็บรูปชั่วคราว
    if not os.path.exists("temp_images"):
        os.makedirs("temp_images")
        
    # วนลูปอ่านข้อมูลผู้ป่วยทีละคน (1 คน / 1 หน้า)
    for index, row in dataframe.iterrows():
        pdf.add_page()
        
        # พิมพ์ข้อมูลทั่วไป
        pdf.set_font("THSarabun-Bold", "B", 16)
        pdf.cell(0, 8, "ส่วนที่ 1: ข้อมูลผู้ป่วยและผู้ติดตาม", 0, 1, fill=False)
        
        # พิมพ์แบบบรรทัดเดียวกัน
        pdf.set_font("THSarabun-Bold", "B", 15)
        pdf.cell(20, 8, "ชื่อ-สกุล:", 0, 0)
        pdf.set_font("THSarabun", "", 15)
        pdf.cell(70, 8, str(row.get('ชื่อ - สกุล ผู้ป่วย', '-')), 0, 0)
        
        pdf.set_font("THSarabun-Bold", "B", 15)
        pdf.cell(10, 8, "อายุ:", 0, 0)
        pdf.set_font("THSarabun", "", 15)
        pdf.cell(20, 8, f"{row.get('อายุ', '-')} ปี", 0, 0)
        
        pdf.set_font("THSarabun-Bold", "B", 15)
        pdf.cell(15, 8, "อำเภอ:", 0, 0)
        pdf.set_font("THSarabun", "", 15)
        pdf.cell(30, 8, str(row.get('อำเภอ', '-')), 0, 1)
        pdf.ln(3)
        
        print_field(pdf, "วันที่ติดตาม", row.get('วัน/เดือน/ปี ที่ติดตาม', '-'))
        print_field(pdf, "ผู้แจ้งผลการติดตาม", row.get('ชื่อ- สกุล  (ผู้แจ้งผลการติดตาม)', '-'))
        pdf.ln(5)
        
        # ส่วนที่ 2
        pdf.set_font("THSarabun-Bold", "B", 16)
        pdf.cell(0, 8, "ส่วนที่ 2: ผลการติดตาม", 0, 1)
        
        print_field(pdf, "สภาพความเป็นอยู่และสภาพแวดล้อม", row.get('สภาพความเป็นอยู่และสภาพแวดล้อมของผู้ป่วย', '-'))
        print_field(pdf, "การรับประทานยาและการรักษา", row.get('ผู้ป่วยรับประทานยาและการรักษาทางจิตเวชต่อเนื่อง อย่างไร', '-'))
        print_field(pdf, "อาการทางจิต", row.get('3. มีอาการทางจิต อย่างไร', '-'))
        print_field(pdf, "การใช้สารเสพติด", row.get('ผู้ป่วยใช้สารเสพติดหรือไม่', '-'))
        print_field(pdf, "ความเสี่ยงที่พบ", row.get('ผู้ป่วยมีความเสี่ยง อย่างไร', '-'))
        print_field(pdf, "ข้อเสนอแนะ", row.get('ท่านมีความเห็นหรือข้อเสนอแนะอย่างไร', '-'))
        
        # จัดการรูปภาพ (ถ้ามี)
        img_url = str(row.get('ภาพการดำเนินงาน', ''))
        temp_img_path = f"temp_images/img_{index}.jpg"
        
        if download_gdrive_image(img_url, temp_img_path):
            # หาพิกัด Y ปัจจุบัน เพื่อวางรูปภาพ
            current_y = pdf.get_y()
            # ถ้ารูปภาพยาวเกินหน้า ให้ขึ้นหน้าใหม่
            if current_y > 200:
                pdf.add_page()
                current_y = 30
            
            pdf.set_font("THSarabun-Bold", "B", 15)
            pdf.cell(0, 10, "ภาพการดำเนินงาน:", 0, 1)
            # แปะรูปลงใน PDF โดยกำหนดความกว้าง 80mm
            pdf.image(temp_img_path, x=20, y=pdf.get_y(), w=80)
            
    return pdf

# ส่วนการอัปโหลดไฟล์
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel ที่ดาวน์โหลดจาก Google Sheet", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success(f"โหลดข้อมูลสำเร็จ จำนวน {len(df)} รายการ")
    
    # แสดงตัวอย่างข้อมูล
    st.dataframe(df[['ชื่อ - สกุล ผู้ป่วย', 'วัน/เดือน/ปี ที่ติดตาม', 'อำเภอ']].head())
    
    if st.button("🖨️ สร้างไฟล์รายงาน PDF"):
        with st.spinner('กำลังโหลดรูปภาพและสร้าง PDF... อาจใช้เวลาสักครู่'):
            pdf = generate_pdf(df)
            pdf_output = "patient_tracking_report.pdf"
            pdf.output(pdf_output)
            
            with open(pdf_output, "rb") as pdf_file:
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ PDF",
                    data=pdf_file,
                    file_name="Report_Psychiatric_Tracking.pdf",
                    mime="application/pdf"
                )
