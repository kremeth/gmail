import gspread
import time
from google.oauth2.service_account import Credentials
import smtplib
import json
from email.message import EmailMessage

# -----------------------------
# CONFIGURATION
# -----------------------------
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1yhShQay-yFhqolaa3CuMoUQ1705uiqefaBubkGGZvrA/edit?gid=0#gid=0'
EMAIL_ADDRESS = 'mathieu.kremeth@gmail.com'
EMAIL_PASSWORD = 'cqis dzyt ttqz yica'  # Use Gmail App Password

# -----------------------------
# AUTHENTICATE WITH GOOGLE SHEETS
# -----------------------------
scopes = ['https://www.googleapis.com/auth/spreadsheets']
GOOGLE_CREDENTIALS = json.loads(os.environ['GOOGLE_CREDENTIALS'])
creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS, scopes=scopes)
client = gspread.authorize(creds)

sheet = client.open_by_url(SHEET_URL).sheet1  # Adjust if not first sheet
data = sheet.get_all_records()

# -----------------------------
# EMAIL FUNCTION
# -----------------------------
def send_email(to_email, first_name):
    msg = EmailMessage()
    msg['Subject'] = 'Nutricode Pre-Seed'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    msg.set_content(f"""\
Hi {first_name},

I found your email address on Crunchbase!

I am working on Nutricode (Singapore-based) — personalized supplement subscriptions powered by your wearable (Apple Watch, Garmin, or WHOOP). The app tracks your health metrics and adapts your supplements monthly based on what your body actually needs.

Quick snapshot (USD):
– $5K MRR from 81 active users
– $69 CAC, down from $494
– $230 LTV, with 80% retention
– Fully bootstrapped

We’re raising $150K, aiming to close by end of June, to scale a profitable go-to-market funnel — with a goal of hitting $100K in MRR by December.

Would you like me to send over our deck and a quick product demo?

Best,  
Mathieu
""")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

# -----------------------------
# LOOP THROUGH CONTACTS
# -----------------------------
for i, row in enumerate(data, start=2):  # start=2 to match actual row number in Sheet (accounting for header)
    name = row['Investor']
    email = row['Linkedin/email']
    blurb = row.get('Blurb', '').strip().upper()

    if blurb != 'Y' and email:
        first_name = name.split()[0]
        try:
            send_email(email, first_name)
            print(f"✅ Sent to {first_name} at {email}")
            sheet.update_cell(i, 3, 'Y')  # Column 3 is 'Blurb' assuming it's the third column
        except Exception as e:
            print(f"❌ Failed to send to {email}: {e}")
        time.sleep(180)  # Sleep for 3 minutes


