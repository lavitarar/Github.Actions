import pandas as pd
import smtplib
import time
import sys
import os
from email.message import EmailMessage

# ================= CONFIG =================
EXCEL_FILE = "data.xlsx"
SHEET_NAME = "Sheet1"
RESUME_FILE = "Lavi_Tarar_1_Year_EXP.pdf"   # ✅ YOUR FILE NAME

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ================= REGION INPUT =================
if len(sys.argv) < 2:
    print("❌ Please provide region (singapore/dubai)")
    sys.exit(1)

region = sys.argv[1].lower()

# ================= LOAD DATA =================
try:
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    sys.exit(1)

# ================= CHECK COLUMNS =================
required_columns = ["Singapore Mail", "Dubai Mail"]

for col in required_columns:
    if col not in df.columns:
        print(f"❌ Missing column: {col}")
        sys.exit(1)

# ================= SELECT EMAIL COLUMN =================
if region == "singapore":
    email_column = "Singapore Mail"
elif region == "dubai":
    email_column = "Dubai Mail"
else:
    print("❌ Invalid region")
    sys.exit(1)

# ================= FILTER EMAILS =================
df_filtered = df[df[email_column].notna()]

print(f"\n📍 Region: {region.upper()}")
print(f"📧 Total Emails: {len(df_filtered)}\n")

# ================= EMAIL FUNCTION =================
def send_email(to_email, name):
    try:
        msg = EmailMessage()
        msg["From"] = EMAIL_USER
        msg["To"] = to_email
        msg["Subject"] = "Application for Data Analyst / Data Engineer Role"

        msg.set_content(f"""
Hi {name},

I hope you're doing well.

I am writing to express my interest in Data Analyst / Data Engineer opportunities at your organization.

I have experience in SQL, Power BI, and ETL pipelines.

Please find my resume attached.

I would really appreciate any opportunity or referral.

Best regards,  
Your Name
""")

        # 📎 ATTACH RESUME
        with open(RESUME_FILE, "rb") as f:
            file_data = f.read()
            file_name = RESUME_FILE

        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="pdf",
            filename=file_name
        )

        # 📤 SEND EMAIL
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)

        print(f"✅ Sent with Resume: {to_email}")

    except Exception as e:
        print(f"❌ Failed: {to_email} | Error: {e}")

# ================= SEND EMAILS =================
for index, row in df_filtered.iterrows():
    email = row[email_column]
    name = row.get("Name", "Sir/Madam")

    send_email(email, name)

    time.sleep(10)  # ⏱ avoid spam block

print("\n🎉 All emails sent successfully!")
