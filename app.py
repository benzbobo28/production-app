import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google import genai
from PIL import Image
import io

st.set_page_config(page_title="Production Report Generator", layout="centered")

st.title("📊 ระบบแปลงใบจดบันทึกยอดผลิตเป็น Excel")
st.write("อัปโหลดรูปภาพฟอร์มบันทึกยอดผลิต ระบบจะดึงข้อมูลตัวเลขและจัดรูปแบบลง Excel ให้โดยอัตโนมัติครับ")

# ช่องใส่ API Key
api_key = st.text_input("Google Gemini API Key", value="AQ.Ab8RN6JKTGGmxqe8YVwhmKLQlFk_XVd2k7oUOCaVaFgTCgM-uw", type="password")

# ช่องอัปโหลดรูปภาพ
uploaded_file = st.file_uploader("เลือกรูปภาพใบจดบันทึก (JPG, PNG)", type=["jpg", "jpeg", "png"])

if uploaded_file and api_key:
    image = Image.open(uploaded_file)
    st.image(image, caption="รูปภาพที่เลือก", use_container_width=True)
    
    if st.button("🚀 ประมวลผลและสร้าง Excel"):
        with st.spinner("กำลังให้ AI วิเคราะห์ข้อมูลจากรูปภาพและจัดรูปแบบตาราง..."):
            try:
                client = genai.Client(api_key=api_key)
                
                # Prompt สั่งให้ AI ดึงข้อมูลแถวและตัวเลขจากรูปภาพตามโครงสร้างฟอร์ม
                prompt = """
                Extract all production log rows (Code, Product, Filling/Weight, and daily data for Mon-Sun: Order, Actual, Dough, Reject) from this image into a Markdown Table format.
                Rules: 
                1. Output ONLY the Markdown table. 
                2. Do not include markdown code blocks (```) or conversational text.
                3. Ensure numbers, product codes, product names, and fillings match the handwritten or printed data in the image accurately.
                """
                
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[image, prompt]
                )
                
                markdown_text = response.text.strip()
                if markdown_text.startswith("```"):
                    lines = markdown_text.split("\n")
                    markdown_text = "\n".join([line for line in lines if not line.startswith("```")])
                
                # แปลง Markdown เป็น DataFrame
                lines = markdown_text.split('\n')
                table_data = []
                
                for line in lines:
                    line = line.strip()
                    if line.startswith('|') and line.endswith('|'):
                        if '---' in line:
                            continue
                        cells = [c.strip() for c in line.split('|')[1:-1]]
                        # ข้ามแถวหัวตารางเดิมของ AI เพื่อมาบังคับใช้หัวตารางมาตรฐานของเรา
                        if any(h in cells[0].lower() for h in ["รหัส", "code"]):
                            continue
                        table_data.append(cells)

                if not table_data:
                    st.error("ไม่พบข้อมูลตาราง กรุณาตรวจสอบความชัดเจนของรูปภาพ")
                else:
                    # กำหนดหัวตารางมาตรฐาน 31 คอลัมน์ (รวมคำว่า dough)
                    expected_headers = [
                        "รหัส", "สินค้า", "ไส้",
                        "จ.สั่งผลิต", "จ.ผลิตจริง", "จ.dough", "จ.คัดเสีย",
                        "อ.สั่งผลิต", "อ.ผลิตจริง", "อ.dough", "อ.คัดเสีย",
                        "พ.สั่งผลิต", "พ.ผลิตจริง", "พ.dough", "พ.คัดเสีย",
                        "พฤ.สั่งผลิต", "พฤ.ผลิตจริง", "พ.dough", "พ.คัดเสีย",
                        "ศ.สั่งผลิต", "ศ.ผลิตจริง", "ศ.dough", "ศ.คัดเสีย",
                        "ส.สั่งผลิต", "ส.ผลิตจริง", "ส.dough", "ส.คัดเสีย",
                        "อา.สั่งผลิต", "อา.ผลิตจริง", "อา.dough", "อา.คัดเสีย"
                    ]
                    
                    # ปรับจำนวนข้อมูลให้ตรงกับจำนวนคอลัมน์มาตรฐาน
                    processed_rows = []
                    for row in table_data:
                        if len(row) >= len(expected_headers):
                            processed_rows.append(row[:len(expected_headers)])
                        else:
                            processed_rows.append(row + [''] * (len(expected_headers) - len(row)))

                    df = pd.DataFrame(processed_rows, columns=expected_headers)
                    
                    # สร้างไฟล์ Excel และจัดรูปแบบด้วย openpyxl
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Production Report')
                    
                    # ตกแต่งไฟล์ Excel เพิ่มเติม (สี, ฟอนต์, เส้นขอบ)
                    output.seek(0)
                    wb = openpyxl.load_workbook(output)
                    ws = wb.active
                    
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(name="Tahoma", size=10, bold=True, color="FFFFFF")
                    data_font = Font(name="Tahoma", size=10)
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_left = Alignment(horizontal="left", vertical="center")
                    
                    thin_border = Border(
                        left=Side(style='thin', color='D9D9D9'),
                        right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'),
                        bottom=Side(style='thin', color='D9D9D9')
                    )

                    # จัดแต่งแถวหัวตาราง
                    for col_num in range(1, ws.max_column + 1):
                        cell = ws.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = align_center
                        cell.border = thin_border
                    
                    # จัดแต่งข้อมูลในตาราง
                    for row_num in range(2, ws.max_row + 1):
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=row_num, column=col_num)
                            cell.font = data_font
                            cell.border = thin_border
                            if col_num <= 3:
                                cell.alignment = align_left
                            else:
                                cell.alignment = align_center

                    # ปรับความกว้างคอลัมน์อัตโนมัติ
                    for col in ws.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = openpyxl.utils.get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

                    final_output = io.BytesIO()
                    wb.save(final_output)
                    excel_data = final_output.getvalue()
                    
                    st.success("✨ ประมวลผลและสร้างไฟล์ Excel สำเร็จเรียบร้อยแล้ว!")
                    
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel (ข้อมูลจากรูปภาพจริง)",
                        data=excel_data,
                        file_name="production_report_from_image.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")