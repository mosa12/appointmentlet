import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import shutil
import re
import time
from datetime import datetime, date
import logging
import tempfile

# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Streamlit app configuration
st.set_page_config(
    page_title="Oxford Education Consultancy - Letter Generator",
    page_icon="📧",
    layout="wide"
)

# Custom CSS for neon aesthetic
st.markdown("""
    <style>
    .main { background-color: #1a1a1a; color: #ffffff; padding: 20px; }
    .stButton>button { 
        background-color: #00ffaa; 
        color: #000000; 
        border: 2px solid #00ffaa; 
        border-radius: 5px; 
        box-shadow: 0 0 10px #00ffaa; 
        padding: 10px 20px; 
    }
    .stButton>button:hover { 
        background-color: #00cc88; 
        box-shadow: 0 0 15px #00ffaa; 
    }
    .stTextInput>label, .stSelectbox>label, .stFileUploader>label, .stDateInput>label, .stRadio>label { 
        color: #00ffaa; 
        font-weight: bold; 
    }
    .stTextInput>div>input, .stTextArea>div>textarea, .stSelectbox>div>select { 
        background-color: #2a2a2a; 
        color: #ffffff; 
        border: 1px solid #00ffaa; 
        border-radius: 5px; 
    }
    .error { color: #ff4444; font-weight: bold; }
    .success { color: #00ffaa; font-weight: bold; }
    .sidebar .sidebar-content { background-color: #2a2a2a; }
    .stProgress .st-bo { background-color: #00ffaa; }
    h1, h2, h3 { color: #00ffaa; }
    </style>
""", unsafe_allow_html=True)

# Header with branding
st.title("📄 Oxford Education Consultancy")
st.subheader("Appointment Letter Generator")

# Sidebar for email configuration
with st.sidebar:
    st.header("Email Configuration")
    sender_email = st.text_input("Sender Email (GoDaddy)", placeholder="user@yourdomain.com", help="Enter your GoDaddy Professional Email.")
    sender_password = st.text_input("Sender Password", type="password", help="Enter your GoDaddy email password.")
    smtp_server = st.text_input("SMTP Server", value="smtpout.secureserver.net", help="Default: smtpout.secureserver.net")
    smtp_port = st.selectbox("SMTP Port", options=[587, 465, 80], index=0, help="587 for TLS, 465 for SSL, 80 for alternate TLS.")
    encryption = st.selectbox("Encryption", options=["TLS", "SSL"], index=0, help="Choose TLS or SSL encryption.")
    email_body_template = st.text_area(
        "Email Body Template",
        value="Dear {{name}},\n\nI hope this email finds you well. On behalf of Oxford Education Consultancy, I am pleased to extend our formal offer for the position of admission counsellor to you. We are excited about the prospect of welcoming you to our team and are confident that your skills and experience will be invaluable in furthering our mission of providing exceptional educational services.\n\nPlease find attached your job appointment letter, which outlines the terms and conditions of your employment, including your start date, compensation package, and other relevant details. We kindly request that you review the document carefully and signify your acceptance by signing and returning a scanned copy to us at your earliest convenience.\n\nShould you have any questions or require clarification on any aspect of the job appointment letter, please do not hesitate to reach out to me directly. We are committed to ensuring a smooth transition for you as you prepare to join Oxford Education Consultancy.\n\nAs part of the onboarding process, we kindly request that you provide the following documents:\n\n1. Photo\n2. Proof of Identification\n3. Address proof\n4. Last qualification certificate\n5. Last pay slip\n6. Experience certificate\n7. Bank passbook\n8. Appointment letter by signing it.\n(Please send everything in pdf format)\n\nThese documents are essential for us to complete your onboarding smoothly and ensure compliance with company policies and regulations.\n\nPlease submit the required documents at your earliest convenience. If you have any questions or need assistance, feel free to reach out to me directly.\n\nOnce again, congratulations on your appointment to this role. We look forward to the opportunity to work together and achieve great things.\n\nWarm regards,\n\nMonirul Islam\nManager\nOxford Education Consultancy",
        height=400,
        help="Use {{name}} for recipient's name. Date is autofetched."
    )

# Main content
st.subheader("Upload Files")
recipient_mode = st.radio("Recipient Mode", ["Single Recipient", "Multiple Recipients"], help="Choose whether to send to one person or multiple people via Excel.")

# Validate email format
def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))

# Send test email
if st.button("Send Test Email", help="Send a test email to verify SMTP settings"):
    if not sender_email or not sender_password:
        st.error("Please provide sender email and password.", icon="❌")
    elif not is_valid_email(sender_email):
        st.error("Invalid sender email format.", icon="❌")
    else:
        test_msg = MIMEMultipart()
        test_msg['From'] = sender_email
        test_msg['To'] = sender_email
        test_msg['Subject'] = "Test Email from Oxford Education Consultancy"
        test_msg.attach(MIMEText("This is a test email to verify SMTP settings.", 'plain'))
        try:
            if encryption == "SSL":
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(test_msg)
            server.quit()
            st.success("Test email sent successfully!", icon="✅")
            logging.info("Test email sent successfully.")
        except Exception as e:
            st.error(f"Failed to send test email: {str(e)}", icon="❌")
            logging.error(f"Test email failed: {str(e)}")

def send_email(sender_email, sender_password, recipient_email, subject, body, attachment_path, smtp_server, smtp_port, encryption):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach PDF
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(attachment_path)}"
    )
    msg.attach(part)
    
    # Connect to SMTP server
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
    # Load template
    doc = DocxTemplate(template_file)
    
    # Render document
    doc.render(context)
    output_doc_path = os.path.join(temp_dir, f"appointment_{recipient_name}.docx")
    doc.save(output_doc_path)
    
    # Convert to PDF
    output_pdf_path = os.path.join(temp_dir, f"appointment_{recipient_name}.pdf")
    try:
        convert(output_doc_path, output_pdf_path)
        if not os.path.exists(output_pdf_path):
            raise Exception("PDF file was not created")
        # Save files
        shutil.copy(output_doc_path, os.path.join(output_dir, f"appointment_{recipient_name}.docx"))
        shutil.copy(output_pdf_path, os.path.join(output_dir, f"appointment_{recipient_name}.pdf"))
        return output_pdf_path
    except Exception as e:
        st.error(f"PDF conversion failed for {recipient_name}: {str(e)}", icon="❌")
        logging.error(f"PDF conversion failed for {recipient_name}: {str(e)}")
        return None

def process_single_recipient(template_file, sender_email, sender_password, smtp_server, smtp_port, encryption, email_body_template, name, date_of_joining, email):
    if not template_file or not name or not date_of_joining or not email:
        st.error("Please provide all required fields: template file, name, date of joining, and email.", icon="❌")
        return
    if not is_valid_email(email):
        st.error(f"Invalid email address: {email}", icon="❌")
        return

    # Create output directory
    output_dir = f"generated_letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(output_dir, exist_ok=True)
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Save uploaded template
        template_path = os.path.join(temp_dir, "template.docx")
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())
        
        # Prepare context for template
        context = {
            'name': name,
            'date_of_joining': date_of_joining.strftime('%Y-%m-%d'),
            'date_of_sending': datetime.now().strftime('%Y-%m-%d'),
            'email': email
        }
        
        # Generate letter
        output_pdf_path = generate_letter(template_path, context, temp_dir, output_dir, name)
        
        if output_pdf_path:
            # Prepare email
            subject = "Job Appointment Letter"
            body = email_body_template.replace("{{name}}", name).replace("Candidate", name)
            success, error = send_email(sender_email, sender_password, email, subject, body, output_pdf_path, smtp_server, smtp_port, encryption)
            
            if success:
                st.success(f"Email sent to {email}", icon="✅")
                logging.info(f"Email sent to {email}")
            else:
                st.error(f"Failed to send email to {email}: {error}", icon="❌")
                logging.error(f"Failed to send email to {email}: {error}")
            
            st.info(f"Generated files saved in: {output_dir}", icon="ℹ️")
        else:
            st.error(f"Skipping email for {name} due to PDF conversion failure.", icon="❌")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}", icon="❌")
        logging.error(f"Processing error: {str(e)}")
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def process_multiple_recipients(template_file, excel_file, sender_email, sender_password, smtp_server, smtp_port, encryption, email_body_template):
    if not template_file or not excel_file or not sender_email or not sender_password:
        st.error("Please upload both files and provide email credentials.", icon="❌")
        return
    
    if not is_valid_email(sender_email):
        st.error("Invalid sender email format.", icon="❌")
        return

    # Create output directory
    output_dir = f"generated_letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(output_dir, exist_ok=True)
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    failed_emails = []
    
    try:
        # Save uploaded template
        template_path = os.path.join(temp_dir, "template.docx")
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())
        
        # Read Excel file
        df = pd.read_excel(excel_file)
        
        # Validate required columns
        required_columns = ['name', 'date_of_joining', 'email']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Excel file must contain columns: {', '.join(required_columns)}", icon="❌")
            return
        
        # Validate email addresses
        for email in df['email']:
            if not is_valid_email(email):
                st.error(f"Invalid email address: {email}", icon="❌")
                return
        
        # Progress bar
        progress_bar = st.progress(0)
        total = len(df)
        success_count = 0
        
        # Process each candidate
        for index, row in df.iterrows():
            # Prepare context for template
            context = {
                'name': row['name'],
                'date_of_joining': pd.to_datetime(row['date_of_joining']).strftime('%Y-%m-%d'),
                'date_of_sending': datetime.now().strftime('%Y-%m-%d'),
                'email': row['email']
            }
            
            # Generate letter
            output_pdf_path = generate_letter(template_path, context, temp_dir, output_dir, row['name'])
            
            if output_pdf_path:
                # Prepare email
                subject = "Job Appointment Letter"
                body = email_body_template.replace("{{name}}", row['name']).replace("Candidate", row['name'])
                success, error = send_email(sender_email, sender_password, row['email'], subject, body, output_pdf_path, smtp_server, smtp_port, encryption)
                
                if success:
                    st.success(f"Email sent to {row['email']}", icon="✅")
                    success_count += 1
                    logging.info(f"Email sent to {row['email']}")
                else:
                    st.error(f"Failed to send email to {row['email']}: {error}", icon="❌")
                    failed_emails.append({'name': row['name'], 'email': row['email'], 'error': error})
                    logging.error(f"Failed to send email to {row['email']}: {error}")
            else:
                st.error(f"Skipping email for {row['name']} due to PDF conversion failure.", icon="❌")
                failed_emails.append({'name': row['name'], 'email': row['email'], 'error': 'PDF conversion failed'})
                logging.error(f"Skipping email for {row['name']} due to PDF conversion failure.")
            
            # Update progress
            progress_bar.progress((index + 1) / total)
        
        # Summary
        st.success(f"Processed {success_count} out of {total} letters successfully.", icon="✅")
        if failed_emails:
            st.warning(f"Failed to send {len(failed_emails)} emails.", icon="⚠️")
            failed_df = pd.DataFrame(failed_emails)
            csv = failed_df.to_csv(index=False)
            st.download_button(
                label="Download Failed Emails",
                data=csv,
                file_name=f"failed_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            logging.info(f"Generated failed emails CSV with {len(failed_emails)} entries.")
        
        st.info(f"Generated files saved in: {output_dir}", icon="ℹ️")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}", icon="❌")
        logging.error(f"Processing error: {str(e)}")
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

# File upload and input fields based on recipient mode
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"], help="Upload a .docx file with placeholders like {{name}}, {{date_of_joining}}, {{date_of_sending}}.")
with col2:
    if recipient_mode == "Multiple Recipients":
        excel_file = st.file_uploader("Upload Candidate Data (.xlsx)", type=["xlsx", "xls"], help="Excel file must have columns: name, date_of_joining, email.")
    else:
        excel_file = None
        single_name = st.text_input("Recipient Name", help="Enter the recipient's full name.")
        single_date_of_joining = st.date_input("Date of Joining", help="Select the recipient's joining date.")
        single_email = st.text_input("Recipient Email", help="Enter the recipient's email address.")

# Button to trigger processing
if st.button("Generate and Send Letters", help="Generate and email appointment letters"):
    with st.spinner("Processing and sending letters..."):
        if recipient_mode == "Single Recipient":
            process_single_recipient(template_file, sender_email, sender_password, smtp_server, smtp_port, encryption, email_body_template, single_name, single_date_of_joining, single_email)
        else:
            process_multiple_recipients(template_file, excel_file, sender_email, sender_password, smtp_server, smtp_port, encryption, email_body_template)

# Reset button
if st.button("Reset Inputs", help="Clear all inputs"):
    st.rerun()

# Instructions
with st.expander("Instructions"):
    st.markdown("""
    1. **Prepare the Template**: Create a Word document with placeholders like `{{name}}`, `{{date_of_joining}}`, `{{date_of_sending}}`, and `{{email}}`.
    2. **Choose Recipient Mode**:
       - **Single Recipient**: Enter the name, date of joining, and email directly.
       - **Multiple Recipients**: Ensure the Excel file has columns: `name`, `date_of_joining`, `email`. The `date_of_sending` column is ignored; the current date is used.
    3. **Email Setup**: Enter your GoDaddy Professional Email and password. Ensure SMTP authentication is enabled in your GoDaddy account.
    4. **Upload Files**: Upload the Word template, and for multiple recipients, the Excel file.
    5. **Send Test Email**: Use the test button to verify SMTP settings.
    6. **Generate and Send**: Click "Generate and Send Letters" to process and email letters.
    7. **Check Outputs**: Generated files are saved locally, and failed emails can be downloaded as a CSV (for multiple recipients).
    8. **Note**: If PDF conversion fails, no email will be sent, and the failure will be logged.
    """)
