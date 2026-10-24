import os
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, ImageMessage, TextSendMessage
from google import genai
import pandas as pd
from PIL import Image
import io

app = Flask(__name__)

# ดึงค่า Config จาก Environment Variables บน Cloud Server
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# กำหนดหัวตาราง 31 คอลัมน์ (ใช้ 'dough')
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

@handler.add(MessageEvent, message=ImageMessage)
def handle_image(event):
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text="📸 ได้รับรูปภาพแล้วครับ! ระบบกำลังประมวลผลแปลงเป็น Excel รอสักครู่...")
    )
    try:
        message_content = line_bot_api.get_message_content(event.message.id)
        image_bytes = message_content.content
        image = Image.open(io.BytesIO(image_bytes))
        
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
        
        df = pd.DataFrame(columns=HEADERS)
        output_file = "production_report.xlsx"
        df.to_excel(output_file, index=False)
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    app.run(port=5000)
