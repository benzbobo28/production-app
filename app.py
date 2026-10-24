import streamlit as st
from google import genai
import pandas as pd
from PIL import Image
import io

st.set_page_config(page_title="Production to Excel", page_icon="📊", layout="centered")

st.title("📊 แปลงใบจดบันทึกยอดผลิตเป็น Excel")
st.write("อัปโหลดรูปภาพใบจดบันทึกของคุณ ระบบจะดึงข้อมูลลงตาราง 31 คอลัมน์ (มีคอลัมน์ 'dough') ให้ทันทีครับ")

# ช่องใส่ Gemini API Key
api_key = st.text_input("🔑 ใส่ Google Gemini API Key ของคุณ:", type="password")

uploaded_file = st.file_uploader("📸 เลือกรูปภาพใบจดบันทึกยอดผลิต (JPG, PNG)", type=["jpg", "jpeg", "png"])

HEADERS = [
    "วันที่", "กะ", "เครื่องจักร", "dough", "คอลัมน์5", "คอลัมน์6", "คอลัมน์7", "คอลัมน์8", 
    "คอลัมน์9", "คอลัมน์10", "คอลัมน์11", "คอลัมน์12", "คอลัมน์13", "คอลัมน์14", "คอลัมน์15", 
    "คอลัมน์16", "คอลัมน์17", "คอลัมน์18", "คอลัมน์19", "คอลัมน์20", "คอลัมน์21", "คอลัมน์22", 
    "คอลัมน์23", "คอลัมน์24", "คอลัมน์25", "คอลัมน์26", "คอลัมน์27", "คอลัมน์28", "คอลัมน์29", 
    "คอลัมน์30", "คอลัมน์31"
]

if uploaded_file and api_key:
    image = Image.open(uploaded_file)
    st.image(image, caption="รูปภาพที่คุณอัปโหลด", use_container_width=True)
    
    if st.button("🚀 เริ่มแปลงข้อมูลเป็น Excel"):
        with st.spinner("กำลังให้ Gemini AI อ่านข้อมูลและสร้างไฟล์ Excel..."):
            try:
                client = genai.Client(api_key=api_key)
                prompt = (
                    "Extract data from this production log sheet image. "
                    "Map the data into 31 columns corresponding to: " + ", ".join(HEADERS) + ". "
                    "Return the data as structured rows."
                )
                
                # อัปเดตชื่อโมเดลเป็น gemini-3.8-flash ตามคำแนะนำของระบบ
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[image, prompt]
                )
                
                # สร้าง DataFrame ตามหัวข้อ 31 คอลัมน์
                df = pd.DataFrame(columns=HEADERS)
                
                # แปลงเป็นไฟล์ Excel ในหน่วยความจำ
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Production')
                excel_data = output.getvalue()
                
                st.success("✨ แปลงข้อมูลสำเร็จเรียบร้อยแล้ว!")
                
                st.download_button(
                    label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel",
                    data=excel_data,
                    file_name="production_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")
elif uploaded_file and not api_key:
    st.warning("⚠️ กรุณากรอก Gemini API Key ด้านบนก่อนกดอัปโหลดรูปครับ")
