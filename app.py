import os
import uuid
from flask import Flask, request, abort, send_from_directory
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, ImageMessage, TextSendMessage
from google import genai
import pandas as pd
from PIL import Image
import io

app = Flask(__name__)

# โฟลเดอร์สำหรับเก็บไฟล์ Excel เพื่อให้ผู้ใช้กดดาวน์โหลด
UPLOAD_FOLDER = 'static'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

HEADERS = [
    "วันที่", "กะ", "เครื่องจักร", "dough", "คอลัมน์5", "คอลัมน์6", "คอลัมน์7", "คอลัมน์8", 
    "คอลัมน์9", "คอลัมน์10", "คอลัมน์11", "คอลัมน์12", "คอลัมน์13", "คอลัมน์14", "คอลัมน์15", 
    "คอลัมน์16", "คอลัมน์17", "คอลัมน์18", "คอลัมน์19", "คอลัมน์20", "คอลัมน์21", "คอลัมน์22", 
    "คอลัมน์23", "คอลัมน์24", "คอลัมน์25", "คอลัมน์26", "คอลัมน์27", "คอลัมน์28", "คอลัมน์29", 
    "คอลัมน์30", "คอลัมน์31"
]

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# Route สำหรับดาวน์โหลดไฟล์ Excel
@app.route('/static/<filename>')
def download_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

@handler.add(MessageEvent, message=ImageMessage)
def handle_image(event):
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text="📸 ได้รับรูปภาพแล้วครับ! ระบบกำลังประมวลผลแปลงเป็น Excel รอสักครู่...")
    )
    
    try:
        # 1. ดาวน์โหลดรูปภาพจาก LINE
        message_content = line_bot_api.get_message_content(event.message.id)
        image_bytes = message_content.content
        image = Image.open(io.BytesIO(image_bytes))
        
        # 2. ใช้ Gemini AI แปลงรูปเป็นข้อมูล
        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt = (
            "Extract data from this production log sheet image. "
            "Map the data into 31 columns corresponding to: " + ", ".join(HEADERS) + ". "
            "Return the data as structured JSON rows."
        )
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[image, prompt]
        )
        
        # 3. แปลงลง Excel และเซฟลงโฟลเดอร์ static
        filename = f"report_{uuid.uuid4().hex[:8]}.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        
        df = pd.DataFrame(columns=HEADERS)
        df.to_excel(filepath, index=False)
        
        # 4. สร้าง ลิงก์ดาวน์โหลด และส่งกลับหาผู้ใช้ใน LINE
        # (Render จะใช้ Request Host URL อัตโนมัติ)
        download_url = f"{request.host_url}static/{filename}"
        
        line_bot_api.push_message(
            event.source.user_id,
            TextSendMessage(text=f"✅ แปลงข้อมูลสำเร็จแล้วครับ!\n\n📊 กดดาวน์โหลดไฟล์ Excel ได้ที่นี่:\n{download_url}")
        )
        
    except Exception as e:
        print(f"Error: {e}")
        line_bot_api.push_message(
            event.source.user_id,
            TextSendMessage(text=f"❌ เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
        )

if __name__ == "__main__":
    app.run(port=5000)
