from datetime import datetime, timedelta
from flask import Flask, render_template_string, request, jsonify
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import threading
import time
import os

app = Flask(__name__)

TEAM_MEMBERS = [
    {"id": "abhishek",  "name": "Abhishek",  "email": "abhishek.kapur1@maersk.com",  "color": "#4A90D9"},
    {"id": "akshay",    "name": "Akshay",    "email": "akshay.mathur@maersk.com",    "color": "#7B68EE"},
    {"id": "ashesh",    "name": "Ashesh",    "email": "ashesh.garg@maersk.com",      "color": "#20B2AA"},
    {"id": "avisek",    "name": "Avisek",    "email": "avisek.nath@maersk.com",      "color": "#FF6B6B"},
    {"id": "dhiraj",    "name": "Dhiraj",    "email": "dhiraj.singh@maersk.com",     "color": "#FFD700"},
    {"id": "kamakhya",  "name": "Kamakhya",  "email": "kamakhya.kinkar@maersk.com",  "color": "#FF8C00"},
    {"id": "manish",    "name": "Manish",    "email": "manish.sambhar@maersk.com",   "color": "#32CD32"},
    {"id": "mohini",    "name": "Mohini",    "email": "mohini.agarwal@maersk.com",   "color": "#FF69B4"},
    {"id": "nibedita",  "name": "Nibedita",  "email": "nibedita.basak@maersk.com",   "color": "#40E0D0"},
    {"id": "shantanu",  "name": "Shantanu",  "email": "shantanu.wagh@maersk.com",    "color": "#9370DB"},
]

TOTAL_SEATS = 10

def init_db():
    conn = sqlite3.connect("database.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            date TEXT,
            seat INTEGER
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_bookings_for_date(date_str):
    conn = sqlite3.connect("database.db")
    c = conn.cursor()
    c.execute("SELECT name, seat FROM bookings WHERE date=?", (date_str,))
    rows = c.fetchall()
    conn.close()
    return rows

def get_user_booking(name, date_str):
    conn = sqlite3.connect("database.db")
    c = conn.cursor()
    c.execute("SELECT seat FROM bookings WHERE name=? AND date=?", (name, date_str))
    row = c.fetchone()
    conn.close()
    return row[0] if row else None

def send_friday_reminder():
    EMAIL_USER = os.environ.get("EMAIL_USER", "")
    EMAIL_PASS = os.environ.get("EMAIL_PASS", "")

    if not EMAIL_USER or not EMAIL_PASS:
        print("EMAIL_USER or EMAIL_PASS not set. Skipping email.")
        return

    next_monday = datetime.today() + timedelta(days=3)
    next_friday = next_monday + timedelta(days=4)
    recipients = [m["email"] for m in TEAM_MEMBERS]

    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto">
      <div style="background:#00243D;padding:28px 32px">
        <h2 style="color:#42B4E6;margin:0">SeatSync — Weekly Reminder</h2>
      </div>
      <div style="padding:28px 32px;background:#ffffff">
        <p>Hey team! It is Friday — time to plan your office days for next week
        ({next_monday.strftime('%d %b')} to {next_friday.strftime('%d %b %Y')}).</p>
        <p>10 seats available | 14 team members</p>
        <a href="{os.environ.get('APP_URL', 'https://your-app.onrender.com')}"
           style="display:inline-block;padding:13px 28px;background:#0073AB;color:#fff;border-radius:8px;text-decoration:none;font-weight:bold">
          Book My Seat for Next Week
        </a>
        <p style="margin-top:28px;color:#999;font-size:12px">
          This reminder is sent every Friday at 4:00 PM IST.
        </p>
      </div>
    </div>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Book Your Office Seats - Week of {next_monday.strftime('%d %b %Y')}"
    msg["From"] = f"SeatSync <{EMAIL_USER}>"
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.sendmail(EMAIL_USER, recipients, msg.as_string())
        print(f"Friday reminder sent at {datetime.now()}")
    except Exception as e:
        print(f"Email failed: {e}")

def email_scheduler():
    last_sent_date = None
    while True:
        now = datetime.utcnow() + timedelta(hours=5, minutes=30)
        is_friday = now.weekday() == 4
        is_4pm = now.hour == 16 and now.minute == 0
        today_str = now.strftime("%Y-%m-%d")
        if is_friday and is_4pm and last_sent_date != today_str:
            print("It is Friday 4PM IST - sending reminder emails...")
            send_friday_reminder()
            last_sent_date = today_str
        time.sleep(60)

scheduler_thread = threading.Thread(target=email_scheduler, daemon=True)
scheduler_thread.start()

@app.route("/")
def home():
    return app.send_static_file("index.html")

@app.route("/api/bookings/<date_str>")
def api_get_bookings(date_str):
    rows = get_bookings_for_date(date_str)
    return jsonify({"date": date_str, "bookings": [{"name": r[0], "seat": r[1]} for r in rows]})

@app.route("/api/book", methods=["POST"])
def api_book():
    data = request.json
    name = data.get("name")
    date_str = data.get("date")
    seat = data.get("seat")

    if not name or not date_str or not seat:
        return jsonify({"success": False, "message": "Missing fields"}), 400

    conn = sqlite3.connect("database.db")
    c = conn.cursor()

    c.execute("SELECT * FROM bookings WHERE name=? AND date=?", (name, date_str))
    if c.fetchone():
        conn.close()
        return jsonify({"success": False, "message": "You already have a booking on this day"}), 409

    c.execute("SELECT * FROM bookings WHERE date=? AND seat=?", (date_str, seat))
    if c.fetchone():
        conn.close()
        return jsonify({"success": False, "message": "That seat is already taken"}), 409

    c.execute("SELECT COUNT(*) FROM bookings WHERE date=?", (date_str,))
    if c.fetchone()[0] >= TOTAL_SEATS:
        conn.close()
        return jsonify({"success": False, "message": "No seats available on this day"}), 409

    c.execute("INSERT INTO bookings (name, date, seat) VALUES (?,?,?)", (name, date_str, seat))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Seat {seat} booked for {name} on {date_str}"})

@app.route("/api/cancel", methods=["POST"])
def api_cancel():
    data = request.json
    name = data.get("name")
    date_str = data.get("date")

    conn = sqlite3.connect("database.db")
    c = conn.cursor()
    c.execute("DELETE FROM bookings WHERE name=? AND date=?", (name, date_str))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Booking cancelled for {name} on {date_str}"})

@app.route("/send-reminder", methods=["POST"])
def manual_reminder():
    try:
        send_friday_reminder()
        return jsonify({"success": True, "message": "Reminder emails sent!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=False)
