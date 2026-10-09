import os
import re
from flask import Flask, abort, request
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import ApiClient, Configuration, MessagingApi, ReplyMessageRequest, TextMessage
from linebot.v3.webhooks import FileMessageContent, MessageEvent
import pandas as pd

app = Flask(__name__)

# ดึงค่า Configuration จาก Environment Variables ของ Render
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")

if not CHANNEL_ACCESS_TOKEN or not CHANNEL_SECRET:
  print(
      "Warning: LINE_CHANNEL_ACCESS_TOKEN or LINE_CHANNEL_SECRET not set in"
      " environment variables!"
  )

configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)


@app.route("/")
def home():
  return "FlowCheckOil Ass. is running successfully!"


@app.route("/callback", methods=["POST"])
def callback():
  signature = request.headers.get("X-Line-Signature", "")
  body = request.get_data(as_text=True)
  app.logger.info("Request body: " + body)

  try:
    handler.handle(body, signature)
  except InvalidSignatureError:
    abort(400)
  return "OK"


@handler.add(MessageEvent, message=FileMessageContent)
def handle_file_message(event):
  with ApiClient(configuration) as api_client:
    line_bot_api = MessagingApi(api_client)
    message_id = event.message.id

    # ดาวน์โหลดไฟล์ที่ผู้ใช้อส่งเข้ามาในแชท
    from linebot.v3.messaging import MessagingApiBlob

    blob_api = MessagingApiBlob(api_client)
    message_content = blob_api.get_message_content(message_id)

    original_filename = getattr(event.message, "file_name", "uploaded_file.xlsx")
    local_path = f"/tmp/{message_id}_{original_filename}"

    with open(local_path, "wb") as fd:
      fd.write(message_content)

    reply_text = process_excel_file(local_path, original_filename)

    line_bot_api.reply_message(
        ReplyMessageRequest(
            reply_token=event.reply_token,
            messages=[TextMessage(text=reply_text)],
        )
    )


def process_excel_file(file_path, filename):
  try:
    # โหลดไฟล์ Excel รองรับหลาย Sheet
    xl = pd.ExcelFile(file_path)
    extracted_data = []

    for sheet_name in xl.sheet_names:
      df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

      # ค้นหาแถวข้อมูลที่มีการขนส่งน้ำมันโดยสแกนหาคำว่า HSD หรือตัวเลขน้ำมัน
      for r_idx, row in df.iterrows():
        row_values = [str(val) for val in row.values if pd.notna(val)]
        row_str = " ".join(row_values)

        # ตรวจสอบว่าในแถวมีคำว่า HSD หรือไม่
        if "HSD" in row_str.upper() or "DIESEL" in row_str.upper():
          # ค้นหาเลข ภส. (เลขตั๋วที่ขึ้นต้นด้วย 70...)
          phs_number = "-"
          volume = "-"

          for val in row.values:
            val_str = str(val).strip()
            # เลข ภส. / เลขตั๋ว มักจะขึ้นต้นด้วย 70 และมีความยาวประมาณ 10-14 หลัก
            if re.match(r"^70\d{8,12}$", val_str):
              phs_number = val_str
            # ค้นหาปริมาณน้ำมันที่เป็นตัวเลขหลักหมื่นขึ้นไป (เช่น 40000, 42000)
            elif pd.notna(val) and isinstance(val, (int, float)):
              if val >= 1000 and val < 100000:
                volume = f"{val:,.0f}"
            elif val_str.isdigit() and int(val_str) >= 1000:
              if int(val_str) < 100000:
                volume = f"{val_str} ลิตร"

          # พยายามหาทะเบียนรถหรือสถานที่ส่งในแถวเดียวกัน
          truck_no = "-"
          destination = "-"
          for val in row.values:
            val_str = str(val).strip()
            if "-" in val_str and len(val_str) <= 12 and any(char.isdigit() for char in val_str):
              if truck_no == "-":
                truck_no = val_str

          extracted_data.append(
              f"✅ อ่านไฟล์ '{filename}' สำเร็จ\n"
              f"• เลข ภส. (ตั๋ว): {phs_number}\n"
              f"• ปริมาณ: {volume}\n"
              f"• ทะเบียนรถ: {truck_no}"
          )
          break

    if extracted_data:
      return "\n\n".join(extracted_data)
    else:
      return (
          f"⚠️ อ่านไฟล์ '{filename}' สำเร็จ แต่ไม่พบรายการข้อมูลเลข ภส.1"
          " หรือปริมาณน้ำมันในแถวข้อมูล กรุณาตรวจสอบรูปแบบไฟล์อีกครั้งครับ"
      )

  except Exception as e:
    return (
        f"❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์ '{filename}': {str(e)}"
    )


@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
  with ApiClient(configuration) as api_client:
    line_bot_api = MessagingApi(api_client)
    line_bot_api.reply_message(
        ReplyMessageRequest(
            reply_token=event.reply_token,
            messages=[
                TextMessage(
                    text=(
                        "สวัสดีครับ! ส่งไฟล์ Excel รายงานการขนส่งน้ำมันเข้ามาได้เลยครับ"
                        " เดี๋ยวผมช่วยตรวจสอบข้อมูลและเลข ภส. ให้ครับ 🛢️"
                    )
                )
            ],
        )
    )


if __name__ == "__main__":
  port = int(os.environ.get("PORT", 5000))
  app.run(host="0.0.0.0", port=port)
