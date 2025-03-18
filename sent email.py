import os
import base64
import pandas as pd
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build



# Scope สำหรับ Gmail API
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# โหลด Credentials
creds = None
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

# ถ้าไม่มี Credentials หรือหมดอายุ ให้ทำการ Login ใหม่
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
        token.write(creds.to_json())

# สร้าง Service ของ Gmail API
service = build("gmail", "v1", credentials=creds)

def send_email(to, subject, message, image_path):
    """ ส่งอีเมลพร้อมแนบรูป """
    msg = MIMEMultipart()
    msg["to"] = to
    msg["subject"] = subject
    msg["From"] = "ablelink.thailand99@gmail.com"

    # เพิ่ม HTML Body
    html_content = f"""
    <html>
    <body>
        <p>{message.replace('\n', '<br>')}</p>
        <img src="cid:image1" width="600">
    </body>
    </html>
    """
    msg.attach(MIMEText(html_content, "html"))

    # แนบรูป
    with open(image_path, "rb") as img_file:
        img = MIMEBase("image", "png", filename=os.path.basename(image_path))
        img.set_payload(img_file.read())
        encoders.encode_base64(img)
        img.add_header("Content-ID", "<image1>")
        img.add_header("Content-Disposition", "inline", filename=os.path.basename(image_path))
        msg.attach(img)

    raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    message_body = {"raw": raw_msg}

    service.users().messages().send(userId="me", body=message_body).execute()
    print(f"Email sent to {to}")

# อ่านไฟล์ Excel และส่งอีเมลให้ทุกคน
excel_file = "email_list.xlsx"  # เปลี่ยนเป็นชื่อไฟล์จริงของคุณ
email_column = "email"  # เปลี่ยนเป็นชื่อคอลัมน์ที่มีอีเมล
image_path = "image.jpg"  # ใส่ชื่อไฟล์รูปภาพที่ต้องการแนบ

df = pd.read_excel(excel_file)

for index, row in df.iterrows():
    recipient_email = row[email_column]
    send_email(
        to=recipient_email,
        subject="subject",
        message="text",
        image_path=image_path
    )
    time.sleep(1.1)
