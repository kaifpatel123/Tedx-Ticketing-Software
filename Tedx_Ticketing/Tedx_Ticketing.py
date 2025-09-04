# ===============================
# TEDx KREA 2025 Ticket Generator & Emailer
# ===============================

import os
import smtplib
import mimetypes
import pandas as pd
import hashlib
from PIL import Image, ImageDraw, ImageFont
from email.message import EmailMessage
import getpass
import time

# ---------------- Paths ----------------
BASE_PATH = "C:/Users/kaifp/Tedx Ticketing Software/Templates"

CSV_FILE = "C:/Users/kaifp/Tedx Ticketing Software/Tedx.csv"

OUTPUT_DIR = "C:/Users/kaifp/Tedx Ticketing Software/GeneratedTickets/Day6 - Faculty DH"
LOG_FILE = os.path.join(OUTPUT_DIR, "issued_tickets7.csv")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- Codes ----------------
speakers = {
    "Pankaj Rai": "PR",
    "Sarika Singh": "SS1",
    "Meenakshi Anantram": "MA",
    "Neeraj Khanna": "NK",
    "Ananth Padmanabhan": "AP",
    "Ishan Shanavas": "IS",
    "Sandhya Sriram": "SS2",
    "Full-Day": "FD",
    "Diamond": "DM"
}

tiers = {
    "silver": "S",
    "gold": "G",
    "platinum": "P",
    "platinum+": "PP",
    "diamond": "D"
}

ticket_types = {
    "early bird": "E",
    "full-day": "F",
    "regular": "R"
}

# ---------------- Ticket Number Generator ----------------
def generate_ticket_number(speaker, tier, ticket_type, ticket_count):
    speaker_code = speakers.get(str(speaker).strip(), "XX")
    tier_code = tiers.get(str(tier).lower().strip(), "X")
    pass_code = ticket_types.get(str(ticket_type).lower().strip(), "X")

    base = f"{speaker_code}-{tier_code}{pass_code}-{ticket_count:04d}"
    unique_str = f"{speaker}-{tier}-{ticket_type}-{ticket_count}"
    hash_suffix = hashlib.sha1(unique_str.encode()).hexdigest()[:6].upper()

    return f"{base}-{hash_suffix}"

# ---------------- Template Path ----------------
def draw_ticket_number(img, ticket_no):
    d = ImageDraw.Draw(img)
    try:
        fnt = ImageFont.truetype("./fonts/CourierPrime-Regular.ttf", 24)  # smaller + neat
    except IOError:
        print("⚠️ Courier Prime not found, using Arial.")
        fnt = ImageFont.truetype("arial.ttf", 24)

    # --- Render ticket number with padding ---
    bbox = d.textbbox((0, 0), ticket_no, font=fnt)
    padding = 20  # extra space around text
    w = (bbox[2] - bbox[0]) + padding
    h = (bbox[3] - bbox[1]) + padding

    txt_img = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    txt_draw = ImageDraw.Draw(txt_img)
    txt_draw.text((padding//2, padding//2), ticket_no, font=fnt, fill=(0, 0, 0))

    # Rotate vertical
    rotated = txt_img.rotate(270, expand=1)

    # --- ALIGNMENT FIX ---
    x = 20   # horizontal position (left ↔ right)
    y = 230  # vertical position (up ↕ down)

    img.paste(rotated, (x, y), rotated)
    return img


# ---------------- Template Path ----------------
def get_template(speaker, ticket_tier):
    folder = speaker if speaker in speakers else "General"
    speaker_folder = os.path.join(BASE_PATH, folder)
    filename = f"{speaker}_{ticket_tier}.jpeg"

    path = os.path.join(speaker_folder, filename)
    if not os.path.exists(path):
        # fallback to generic template
        path = os.path.join(BASE_PATH, "General", "General.jpeg")
    return path
# ---------------- Generate Ticket ----------------
def generate_ticket(speaker, ticket_tier, ticket_type, ticket_count):
    template_path = get_template(speaker, ticket_tier)
    img = Image.open(template_path).convert("RGB")
    ticket_no = generate_ticket_number(speaker, ticket_tier, ticket_type, ticket_count)
    img = draw_ticket_number(img, ticket_no)

    filename = os.path.join(OUTPUT_DIR, f"ticket_{ticket_no}.jpeg")
    img.save(filename, quality=95)
    return ticket_no, filename

# ---------------- Send Email ----------------
def send_mail(to, first_name, last_name, ticket_no, filename, smtp, sender):
    receiver = str(to)
    msg_body = f"""
Dear {first_name} {last_name},

Thank you for booking your TEDx Krea ticket! 🎟️

Event Details:
TEDx KREA University
Date: 6th Sept 2025
Venue: NAB Auditorium

Your Ticket Number: {ticket_no}

Please find your ticket attached along with the Rules for Attendees.
Kindly carry your ticket (print or digital) and arrive 15 minutes early.
Your phones will be collected at entry, and please avoid bringing any food and drink in the auditorium.
Also please read the rule book before attending the event.
Looking forward to having you!

Regards,
Team TEDx Krea
"""

    msg = EmailMessage()
    msg['subject'] = f'Your TEDx Krea 2025 Ticket - {ticket_no}'
    msg['from'] = sender
    msg['to'] = receiver
    msg.set_content(msg_body)

    # Attach the ticket image
    with open(filename, 'rb') as fp:
        file_data = fp.read()
        maintype, _, subtype = (mimetypes.guess_type(filename)[0] or 'application/octet-stream').partition("/")
        msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=os.path.basename(filename))

    # Attach the rules PDF
    rules_file = r"C:/Users/kaifp/Tedx Ticketing Software/Templates/Rules for TEDx Event Attendees.pdf"
    with open(rules_file, 'rb') as fp:
        file_data = fp.read()
        maintype, _, subtype = (mimetypes.guess_type(rules_file)[0] or 'application/pdf').partition("/")
        msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename="Rules.pdf")

    smtp.send_message(msg)
    print(f"✅ Sent to {first_name} {last_name} ({receiver})")


# ---------------- Process CSV ------------
def process_csv():
    df = pd.read_csv(CSV_FILE)
    df["Ticket Number"] = ""
    ticket_count = 0

    print("CSV Columns:", df.columns.tolist())

    required_cols = ["First Name", "Email ID", "Speaker", "Ticket Type", "Ticket Tier"]
    if not all(col in df.columns for col in required_cols):
        raise ValueError(f"CSV missing required columns. Found: {df.columns.tolist()}")

    # Setup SMTP inside function
    sender = "tedx@krea.edu.in"
    password = getpass.getpass("Enter your Gmail App Password: ")
    smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    smtp.login(sender, password)

    for i, row in df.iterrows():
        if pd.isna(row["First Name"]) or pd.isna(row["Email ID"]) or pd.isna(row["Speaker"]) or pd.isna(row["Ticket Type"]) or pd.isna(row["Ticket Tier"]):
            print(f"⚠️ Skipping row {i+2} due to missing data.")
            continue

        first = str(row["First Name"]).strip()
        last = str(row.get("Last Name", "")).strip()
        email = str(row["Email ID"]).strip()
        speaker = str(row["Speaker"]).strip()
        ticket_tier = str(row["Ticket Tier"]).strip()
        ticket_type = str(row["Ticket Type"]).strip()

        ticket_count += 1
        print(f"Generating for {first} {last} - {speaker} ({ticket_tier}, {ticket_type})")

        try:
            ticket_no, filename = generate_ticket(speaker, ticket_tier, ticket_type, ticket_count)
            send_mail(email, first, last, ticket_no, filename, smtp, sender)
            df.at[i, "Ticket Number"] = ticket_no
        except FileNotFoundError as e:
            print(f"⚠️ Skipping {first} {last}: {e}")

        time.sleep(2)  # Gmail anti-spam delay

    smtp.quit()
    df.to_csv(LOG_FILE, index=False)
    print(f"\n📑 Log saved at: {LOG_FILE}")


# ---------------- Run ----------------
if __name__ == "__main__":
    process_csv()
