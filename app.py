from flask import Flask, request, abort
import os
import pandas as pd
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, FileMessage, TextSendMessage

app = Flask(__name__)

# ตั้งค่า Token และ Secret จาก LINE Developers
LINE_CHANNEL_ACCESS_TOKEN = 'b3THdRpQxz9QmE1QzysGnt7U8ng6rsKMq+e56CEjP12XItIQcPJtkzavYeQ50ep6HQpIXc8up87wmHIxmyWOs99sP6P7MND8RP/H8ePNCmAYMuV70GIsjMfa0p2dR2Q6UFQOLqF4JKTa2hEwcf9biAdB04t89/1O/w1cDnyilFU='
LINE_CHANNEL_SECRET = 'b431c328c07682a159b4c45314a61147'

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

UPLOAD_FOLDER = './downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ฐานข้อมูลจำลองในหน่วยความจำ สำหรับเก็บประวัติเลข ภส.1 (ในการใช้งานจริงสามารถเปลี่ยนเป็น Database เช่น SQLite)
excise_database = {}

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=FileMessage)
def handle_file_message(event):
    message_id = event.message.id
    file_name = event.message.file_name
    
    # ตรวจสอบว่าเป็นไฟล์ Excel (.xls หรือ .xlsx) หรือไม่
    if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="❌ กรุณาส่งเฉพาะไฟล์ Excel (.xls หรือ .xlsx) สำหรับตรวจสอบข้อมูลใบขนและเลข ภส.1 เท่านั้นครับ")
        )
        return

    # ดาวน์โหลดไฟล์เก็บไว้ในระบบ
    message_content = line_bot_api.get_message_content(message_id)
    file_path = os.path.join(UPLOAD_FOLDER, file_name)
    with open(file_path, 'wb') as fd:
        for chunk in message_content.iter_content():
            fd.write(chunk)

    # ประมวลผลและตรวจสอบไฟล์ Excel
    result_message = process_excel_file(file_path, file_name)
    
    # ส่งข้อความสรุปผลกลับเข้าแชท LINE
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=result_message)
    )

def process_excel_file(file_path, file_name):
    try:
        # อ่านไฟล์ดิบเพื่อค้นหาแถวหัวตารางอัตโนมัติ (Smart Parser รองรับฟอร์มที่หลากหลาย)
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    except Exception as e:
        return f"❌ ไม่สามารถเปิดไฟล์ Excel ได้: {str(e)}"

    header_row_index = -1
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
        if 'ภส' in row_str or 'จำนวน' in row_str or 'Transport' in row_str or 'Order No' in row_str:
            header_row_index = idx
            break

    if header_row_index == -1:
        # หากหาแถวหัวตารางไม่เจอ ลองโหลดแบบปกติโดยใช้แถวที่ 0 หรือ 1 เป็นหลัก
        header_row_index = 0

    try:
        df = pd.read_excel(file_path, sheet_name=0, header=header_row_index)
    except Exception as e:
        return f"❌ เกิดข้อผิดพลาดในการอ่านโครงสร้างตาราง: {str(e)}"

    # แปลงชื่อคอลัมน์ให้เป็นมาตรฐานกลาง (Mapping)
    column_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()
        if 'ภส' in col_str or 'กระดุมรถ' in col_str:
            column_mapping[col] = 'Excise_No'
        elif 'จำนวน' in col_str or 'ลิตร' in col_str or 'ความจุ' in col_str:
            if 'ลิตร' in col_str or 'ความจุ' in col_str or 'จำนวน(ลิตร)' in col_str:
                column_mapping[col] = 'Quantity'
        elif 'ขนส่ง' in col_str or 'Transport' in col_str:
            column_mapping[col] = 'Transporter'
        elif 'ชนิด' in col_str or 'Product' in col_str:
            column_mapping[col] = 'Product_Type'

    df = df.rename(columns=column_mapping)

    # ตรวจสอบคอลัมน์บังคับ
    if 'Excise_No' not in df.columns or 'Quantity' not in df.columns:
        return f"⚠️ ไฟล์ '{file_name}' มีรูปแบบที่ระบบยังไม่รองรับอัตโนมัติ (ไม่พบคอลัมน์เลข ภส. หรือ ปริมาณลิตรที่ชัดเจน)"

    report_lines = [f"📁 **รายงานผลการตรวจสอบไฟล์:**\n`{file_name}`\n" + "—" * 24]
    valid_count = 0

    for index, row in df.iterrows():
        excise_raw = row.get('Excise_No')
        if pd.isna(excise_raw):
            continue
            
        excise_no = str(excise_raw).strip()
        # ข้ามแถวที่เป็นค่าว่างหรือแถวสรุปยอดรวม
        if not excise_no or excise_no.lower() == 'nan' or 'total' in excise_no.lower() or 'รวม' in excise_no:
            continue

        # ดึงข้อมูลอื่น ๆ (ถ้ามี)
        qty_val = row.get('Quantity', 0)
        try:
            new_qty = float(str(qty_val).replace(',', ''))
        except ValueError:
            new_qty = 0.0

        if new_qty <= 0:
            continue

        transporter = str(row.get('Transporter', 'ไม่ระบุ')).strip()
        if transporter == 'nan':
            transporter = 'ไม่ระบุ'
            
        product = str(row.get('Product_Type', 'HSD/น้ำมัน')).strip()
        if product == 'nan':
            product = 'HSD'

        valid_count += 1
        status_text = ""

        # ตรวจสอบประวัติการส่งไฟล์ซ้ำ (Version Control / Reconciliation)
        if excise_no in excise_database:
            old_data = excise_database[excise_no]
            old_qty = old_data['Quantity']
            diff = new_qty - old_qty

            if diff != 0:
                change_type = "เพิ่มขึ้น 📈" if diff > 0 else "ลดลง 📉"
                status_text = (
                    f"\n   ⚠️ **[พบเลข ภส.1 ซ้ำในระบบ]**\n"
                    f"   🔄 *มีการแก้ไขปริมาณ:*\n"
                    f"   • ยอดเดิม: {old_qty:,.2f} ลิตร\n"
                    f"   • ยอดใหม่: {new_qty:,.2f} ลิตร\n"
                    f"   • **เปลี่ยนแปลง:** {change_type} {abs(diff):,.2f} ลิตร"
                )
            else:
                status_text = "\n   📌 *[ข้อมูลเลข ภส.1 ซ้ำ แต่ปริมาณคงเดิม]*"
        else:
            status_text = "\n   ✨ *[บันทึกข้อมูลใหม่ครั้งแรก]*"

        # บันทึกสถานะล่าสุดลงฐานข้อมูล
        excise_database[excise_no] = {
            'Quantity': new_qty,
            'Product': product,
            'Transporter': transporter,
            'File': file_name
        }

        item_summary = (
            f"\n🔹 **เลข ภส.1:** `{excise_no}`\n"
            f"   - สินค้า: {product}\n"
            f"   - ปริมาณ: **{new_qty:,.2f} ลิตร**\n"
            f"   - ขนส่งโดย: {transporter}\n"
            f"   - สถานะ: {status_text}\n"
        )
        report_lines.append(item_summary)

    if valid_count == 0:
        return f"⚠️ อ่านไฟล์ '{file_name}' สำเร็จ แต่ไม่พบรายการข้อมูลตัวเลข ภส.1 หรือปริมาณน้ำมันในแถวข้อมูล กรุณาตรวจสอบรูปแบบไฟล์อีกครั้งครับ"

    report_lines.append("—" * 24)
    report_lines.append(f"✅ ประมวลผลสำเร็จ {valid_count} รายการ | พร้อมตรวจสอบความถูกต้องครับ")

    return "\n".join(report_lines)

if __name__ == "__main__":
    app.run(port=5000)