import os
import re
from flask import Flask, abort, request
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import ApiClient, Configuration, MessagingApi, ReplyMessageRequest, TextMessage
from linebot.v3.webhooks import FileMessageContent, MessageEvent
import pandas as pd

app = Flask(__name__)

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
    xl = pd.ExcelFile(file_path)
    all_records = []

    for sheet_name in xl.sheet_names:
      df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

      for r_idx, row in df.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        row_str = " ".join(row_values)

        # ตรวจสอบแถวที่เป็นรายการขนส่งน้ำมัน (มี HSD หรือ DIESEL)
        if "HSD" in row_str.upper() or "DIESEL" in row_str.upper():
          phs_number = "-"
          volume = "-"
          truck_no = "-"
          driver_name = "-"
          phone_no = "-"

          # ค้นหาข้อมูลเชิงลึกในแถว
          for val in row.values:
            val_str = str(val).strip()
            # 1. เลข ภส. (ขึ้นต้นด้วย 70 และมีความยาว 10-14 หลัก)
            if re.match(r"^70\d{8,12}$", val_str):
              phs_number = val_str
            # 2. ปริมาณน้ำมัน (ตัวเลขหลักหมื่น เช่น 40000, 42000, 43000)
            elif pd.notna(val) and isinstance(val, (int, float)):
              if 1000 <= val < 100000:
                volume = f"{val:,.0f} ลิตร"
            elif val_str.isdigit() and 1000 <= int(val_str) < 100000:
              volume = f"{val_str} ลิตร"
            # 3. เบอร์โทรศัพท์ (รูปแบบ 0xx-xxx-xxxx หรือ 0xxxxxxxx)
            elif re.match(r"^0\d{1,2}[-\s]?\d{3}[-\s]?\d{4}$", val_str):
              phone_no = val_str
            # 4. ทะเบียนรถ (มีขีดและตัวเลขผสม เช่น 70-4329 หรือ กท.700)
            elif "-" in val_str and any(char.isdigit() for char in val_str) and len(val_str) <= 18:
              if truck_no == "-":
                truck_no = val_str
              else:
                truck_no += f" / {val_str}"

          # ค้นหาชื่อคนขับ (มักจะเป็นข้อความที่มีตัวอักษรไทยยาว 2-4 คำติดกันในแถว)
          for val in row.values:
            val_str = str(val).strip()
            if re.match(r"^[ก-ฮ\s]{4,30}$", val_str) and "HSD" not in val_str and "ระนอง" not in val_str:
              if driver_name == "-":
                driver_name = val_str

          all_records.append(
              f"🚚 รายการที่ {len(all_records)+1}\n"
              f"• เลข ภส. (ตั๋ว): {phs_number}\n"
              f"• ปริมาณ: {volume}\n"
              f"• ทะเบียนรถ: {truck_no}\n"
              f"• พนักงานขับรถ: {driver_name}\n"
              f"• เบอร์โทร: {phone_no}"
          )

    if all_records:
      header_msg = f"✅ อ่านไฟล์ '{filename}' สำเร็จ (พบ {len(all_records)} รายการ):\n\n"
      return header_msg + "\n\n".join(all_records)
    else:
      return (
          f"⚠️ อ่านไฟล์ '{filename}' สำเร็จ แต่ไม่พบรายการข้อมูลเลข ภส.1"
          " หรือปริมาณน้ำมันในแถวข้อมูล กรุณาตรวจสอบรูปแบบไฟล์อีกครั้งครับ"
      )

  except Exception as e:
    return f"❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์ '{filename}': {str(e)}"


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
                        " เดี๋ยวผมช่วยตรวจเช็คข้อมูลทุกคันให้ครับ 🛢️"
                    )
                )
            ],
        )
    )


if __name__ == "__main__":
  port = int(os.environ.get("PORT", 5000))
  app.run(host="0.0.0.0", port=port)
