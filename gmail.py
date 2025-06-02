import gspread
import time
import random
import smtplib
from email.message import EmailMessage
from google.oauth2.service_account import Credentials
import json
import os

# -----------------------------
# CONFIGURATION
# -----------------------------
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1yhShQay-yFhqolaa3CuMoUQ1705uiqefaBubkGGZvrA/edit?gid=0#gid=0'
EMAIL_ADDRESS = 'mathieu.kremeth@gmail.com'
EMAIL_PASSWORD = 'cqis dzyt ttqz yica'  # Gmail App Password
SPREADSHEET_COL_BLURB = 3  # Assumes 'Blurb' is in column C

# -----------------------------
# GOOGLE SHEETS AUTHENTICATION
# -----------------------------
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds_dict = json.loads(os.environ['GOOGLE_CREDENTIALS'])
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
client = gspread.authorize(creds)
sheet = client.open_by_url(SHEET_URL).sheet1
data = sheet.get_all_records()

# -----------------------------
# EMAIL FUNCTION
# -----------------------------
def send_email(to_email, first_name):
    msg = EmailMessage()
    msg['Subject'] = 'Nutricode Pre-Seed'
    msg['From'] = f"Mathieu Kremeth <{EMAIL_ADDRESS}>"
    msg['To'] = to_email
    msg['Reply-To'] = EMAIL_ADDRESS
    msg['List-Unsubscribe'] = f"<mailto:{EMAIL_ADDRESS}>"

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
# SENDING LOOP WITH DELAY AND ERROR HANDLING
# -----------------------------
emails_sent = 0
MAX_DAILY_LIMIT = 300

for i, row in enumerate(data, start=2):
    if emails_sent >= MAX_DAILY_LIMIT:
        print("🚫 Reached daily sending limit. Stopping.")
        break

    name = row.get('Investor', '')
    email = row.get('Linkedin/email', '')
    # email = 'm42611638@gmail.com'
    blurb = row.get('Blurb', '').strip().upper()

    if blurb != 'Y' and email:
        first_name = name.split()[0] if name else 'there'
        try:
            send_email(email, first_name)
            sheet.update_cell(i, SPREADSHEET_COL_BLURB, 'Y')
            print(f"✅ Sent to {first_name} at {email}")
            emails_sent += 1

            if emails_sent % 50 == 0:
                print("⏸️ Batch complete. Taking a longer 5-minute break...")
                time.sleep(300)
            else:
                sleep_time = random.uniform(180, 240)
                print(f"⏱️ Sleeping for {int(sleep_time)} seconds...")
                time.sleep(sleep_time)

        except Exception as e:
            print(f"❌ Failed to send to {email}: {e}")
            continue
