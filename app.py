from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from docxtpl import DocxTemplate
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import tempfile
import shutil
from datetime import datetime
import logging
import re

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))

def send_email(sender_email, sender_password, recipient_email, subject, body, attachment_path, smtp_server, smtp_port, encryption):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
    msg.attach(part)
    
    try:
        if encryption == "SSL":
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True, None
    except Exception as e:
        return False, str(e)

def generate_letter(template_file, context, temp_dir, output_dir, recipient_name):
    doc = DocxTemplate(template_file)
    doc.render(context)
    output_doc_path = os.path.join(temp_dir, f"appointment_{recipient_name}.docx")
    doc.save(output_doc_path)
    shutil.copy(output_doc_path, os.path.join(output_dir, f"appointment_{recipient_name}.docx"))
    return output_doc_path

@app.post("/generate-letters")
async def generate_letters(
    template_file: UploadFile = File(...),
    recipient_mode: str = Form(...),
    sender_email: str = Form(...),
    sender_password: str = Form(...),
    smtp_server: str = Form(...),
    smtp_port: int = Form(...),
    encryption: str = Form(...),
    email_body: str = Form(...),
    recipient_name: str = Form(None),
    date_of_joining: str = Form(None),
    recipient_email: str = Form(None),
    excel_file: UploadFile = File(None)
):
    if not is_valid_email(sender_email):
        raise HTTPException(status_code=400, detail="Invalid sender email format")
    
    temp_dir = tempfile.mkdtemp()
    output_dir = f"generated_letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        template_path = os.path.join(temp_dir, "template.docx")
        with open(template_path, "wb") as f:
            f.write(await template_file.read())
        
        if recipient_mode == "Single Recipient":
            if not recipient_name or not date_of_joining or not recipient_email:
                raise HTTPException(status_code=400, detail="Missing recipient details")
            if not is_valid_email(recipient_email):
                raise HTTPException(status_code=400, detail="Invalid recipient email format")
            context = {
                "name": recipient_name,
                "date_of_joining": date_of_joining,
                "date_of_sending": datetime.now().strftime("%Y-%m-%d"),
                "email": recipient_email
            }
            output_path = generate_letter(template_path, context, temp_dir, output_dir, recipient_name)
            success, error = send_email(sender_email, sender_password, recipient_email, "Job Appointment Letter", 
                                     email_body.replace("{{name}}", recipient_name), output_path, smtp_server, smtp_port, encryption)
            if success:
                return {"message": f"Letter generated and email sent to {recipient_email}", "output_dir": output_dir}
            else:
                raise HTTPException(status_code=500, detail=f"Email sending failed: {error}")
        else:
            if not excel_file:
                raise HTTPException(status_code=400, detail="Excel file required for multiple recipients")
            excel_path = os.path.join(temp_dir, "candidates.xlsx")
            with open(excel_path, "wb") as f:
                f.write(await excel_file.read())
            df = pd.read_excel(excel_path)
            if not all(col in df.columns for col in ["name", "date_of_joining", "email"]):
                raise HTTPException(status_code=400, detail="Excel file must contain name, date_of_joining, email columns")
            
            results = []
            total = len(df)
            success_count = 0
            for index, row in df.iterrows():
                if not is_valid_email(row["email"]):
                    results.append({"name": row["name"], "email": row["email"], "status": "Failed: Invalid email format"})
                    continue
                context = {
                    "name": row["name"],
                    "date_of_joining": pd.to_datetime(row["date_of_joining"]).strftime("%Y-%m-%d"),
                    "date_of_sending": datetime.now().strftime("%Y-%m-%d"),
                    "email": row["email"]
                }
                output_path = generate_letter(template_path, context, temp_dir, output_dir, row["name"])
                success, error = send_email(sender_email, sender_password, row["email"], "Job Appointment Letter", 
                                         email_body.replace("{{name}}", row["name"]), output_path, smtp_server, smtp_port, encryption)
                if success:
                    success_count += 1
                    results.append({"name": row["name"], "email": row["email"], "status": "Success"})
                else:
                    results.append({"name": row["name"], "email": row["email"], "status": f"Failed: {error}"})
            return {
                "message": f"Processed {success_count} out of {total} letters successfully",
                "results": results,
                "output_dir": output_dir,
                "failed_emails": [r for r in results if "Failed" in r["status"]]
            }
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

@app.post("/test-email")
async def test_email(
    sender_email: str = Form(...),
    sender_password: str = Form(...),
    smtp_server: str = Form(...),
    smtp_port: int = Form(...),
    encryption: str = Form(...)
):
    if not is_valid_email(sender_email):
        raise HTTPException(status_code=400, detail="Invalid sender email format")
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = sender_email
    msg['Subject'] = "Test Email from Oxford Education Consultancy"
    msg.attach(MIMEText("This is a test email to verify SMTP settings.", 'plain'))
    try:
        if encryption == "SSL":
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return {"message": "Test email sent successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to send test email: {str(e)}")
