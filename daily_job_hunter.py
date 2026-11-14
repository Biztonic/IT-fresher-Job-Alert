"""
Daily Job Hunter & WhatsApp Notification Automation for Rohan Tingare
Searches for QA Automation / Selenium roles in Pune and sends a curated digest to WhatsApp via Twilio.
Supports multiple recipient WhatsApp numbers simultaneously!
"""

import os
import sys
import datetime
import pandas as pd
from jobspy import scrape_jobs

# Candidate Configuration
CANDIDATE_NAME = "Rohan Tingare"
SEARCH_TERM = "QA Automation Engineer OR Selenium Test Engineer"
LOCATION = "Pune, Maharashtra, India"

# Default Target Phone Numbers (comma-separated)
# Every number must send 'join century-nails' to +1 415 523 8886 once from WhatsApp
DEFAULT_TARGET_PHONES = os.getenv("TARGET_PHONES", os.getenv("TARGET_PHONE", "918999771788, 919021753459"))

# Twilio Credentials (loaded securely from environment variables)
DEFAULT_TWILIO_SID = os.getenv("TWILIO_ACCOUNT_SID", "")
DEFAULT_TWILIO_TOKEN = os.getenv("TWILIO_AUTH_TOKEN", "")
DEFAULT_TWILIO_FROM = os.getenv("TWILIO_FROM", "whatsapp:+14155238886")

def fetch_daily_jobs(limit=10, hours_old=48):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Searching job boards for '{SEARCH_TERM}' in {LOCATION}...")
    try:
        jobs = scrape_jobs(
            site_name=["linkedin", "indeed", "glassdoor", "google"],
            search_term=SEARCH_TERM,
            location=LOCATION,
            results_wanted=limit,
            hours_old=hours_old,
            country_indeed="India",
        )
        print(f"Successfully retrieved {len(jobs)} jobs.")
        return jobs
    except Exception as e:
        print(f"Error fetching jobs: {e}")
        return pd.DataFrame()

def format_whatsapp_message(jobs_df):
    today_str = datetime.date.today().strftime("%d %b %Y")
    if jobs_df.empty:
        return f"Good morning *{CANDIDATE_NAME}*! 🚀\nNo new QA Automation roles found in the last 24 hours. Will check again tomorrow morning!"
    
    msg_lines = [
        f"🎯 *Daily QA Automation Job Alert for {CANDIDATE_NAME}*",
        f"📅 Date: *{today_str}* | Location: *Pune / Remote*",
        f"🔍 Top {min(len(jobs_df), 5)} curated openings matching your resume:\n"
    ]
    
    for i, (_, row) in enumerate(jobs_df.head(5).iterrows(), 1):
        title = str(row.get('title', 'QA Automation Engineer')).strip()
        company = str(row.get('company', 'Company')).strip()
        loc = str(row.get('location', 'Pune')).strip()
        site = str(row.get('site', 'Job Board')).capitalize()
        url = str(row.get('job_url', ''))
        
        msg_lines.append(f"*{i}. {title}*")
        msg_lines.append(f"🏢 Company: {company}")
        msg_lines.append(f"📍 Location: {loc} ({site})")
        if url:
            msg_lines.append(f"🔗 Apply: {url}")
        msg_lines.append("")
        
    msg_lines.append("💡 *Tip:* Highlight your Provilac Selenium POM framework & Postman API testing before applying!")
    return "\n".join(msg_lines)

def send_whatsapp_twilio(message_text, account_sid, auth_token, from_num=None, to_nums=None):
    from twilio.rest import Client
    
    from_whatsapp = from_num or os.getenv("TWILIO_FROM", DEFAULT_TWILIO_FROM)
    if not from_whatsapp.startswith("whatsapp:"):
        from_whatsapp = f"whatsapp:{from_whatsapp}"
        
    raw_targets = to_nums or os.getenv("TARGET_PHONES") or os.getenv("TARGET_PHONE") or DEFAULT_TARGET_PHONES or ""
    phone_list = [p.strip().replace("+", "") for p in raw_targets.split(",") if p.strip().replace("+", "")]
    
    if not phone_list:
        print("❌ No recipient phone numbers specified.")
        return False
        
    client = Client(account_sid, auth_token)
    success_count = 0
    
    for phone in phone_list:
        to_whatsapp = f"whatsapp:+{phone}" if not phone.startswith("whatsapp:") else phone
        print(f"Dispatching WhatsApp alert from {from_whatsapp} to {to_whatsapp} via Twilio...")
        try:
            msg = client.messages.create(
                body=message_text,
                from_=from_whatsapp,
                to=to_whatsapp
            )
            print(f"✅ Delivered to {to_whatsapp} (SID: {msg.sid}, Status: {msg.status})")
            success_count += 1
        except Exception as e:
            print(f"❌ Failed to dispatch to {to_whatsapp}: {e}")
            if "63015" in str(e):
                print(f"ℹ️ Note: {to_whatsapp} must send 'join century-nails' to +1 415 523 8886 on WhatsApp first to receive messages.")
                
    return success_count > 0

def run_job_hunt(twilio_sid=None, twilio_token=None, target_phones=None, save_excel=True):
    sid = twilio_sid or os.getenv("TWILIO_ACCOUNT_SID", DEFAULT_TWILIO_SID)
    token = twilio_token or os.getenv("TWILIO_AUTH_TOKEN", DEFAULT_TWILIO_TOKEN)
    
    # Fetch jobs
    jobs = fetch_daily_jobs(limit=15, hours_old=72)
    
    if save_excel and not jobs.empty:
        today_slug = datetime.date.today().strftime("%Y_%m_%d")
        csv_path = f"qa_jobs_pune_{today_slug}.csv"
        jobs.to_csv(csv_path, index=False)
        print(f"Saved full job details to: {csv_path}")
        
    # Format message
    msg = format_whatsapp_message(jobs)
    print("\n--- PREVIEW OF WHATSAPP MESSAGE ---")
    print(msg)
    print("------------------------------------\n")
    
    if sid and token:
        send_whatsapp_twilio(msg, sid, token, to_nums=target_phones)
    else:
        print("ℹ️ Twilio credentials not provided. Set TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN.")

if __name__ == "__main__":
    sid = None
    token = None
    targets = None
    for i, arg in enumerate(sys.argv):
        if arg == "--sid" and i + 1 < len(sys.argv):
            sid = sys.argv[i + 1]
        elif arg == "--token" and i + 1 < len(sys.argv):
            token = sys.argv[i + 1]
        elif (arg == "--phone" or arg == "--phones") and i + 1 < len(sys.argv):
            targets = sys.argv[i + 1]
            
    run_job_hunt(twilio_sid=sid, twilio_token=token, target_phones=targets)
