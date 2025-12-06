from flask import Flask, render_template, request, redirect, url_for, flash
from openpyxl import Workbook, load_workbook
from twilio.rest import Client
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = "some_secret_key"  # needed for flash messages

EXCEL_FILE = "booking.xlsx"
SHEET_NAME = "Bookings"


def init_excel():
    """Create Excel file with header row if it doesn't exist."""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(["S.No", "Name", "Game", "Date", "Time Slot", "Phone", "Created At"])
        wb.save(EXCEL_FILE)


def add_booking_to_excel(name, game, date, time_slot, phone):
    """Append a new booking row to the Excel file."""
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]

    # Auto-increment S.No
    next_row_number = ws.max_row  # header is row 1, so first data row is 2
    s_no = next_row_number  # simple numbering

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws.append([s_no, name, game, date, time_slot, phone, created_at])
    wb.save(EXCEL_FILE)




def send_whatsapp_message(to_number, name, game, date, time_slot):
    account_sid = "YOUR_TWILIO_SID"
    auth_token = "YOUR_TWILIO_AUTH_TOKEN"
    client = Client(account_sid, auth_token)

    message_text = f"""
Hello {name}! 🎮

Your slot is confirmed for:
Game: {game}
Date: {date}
Time: {time_slot}

Thank you for booking!
"""

    client.messages.create(
        from_='whatsapp:+918328648371',   # Twilio WhatsApp sandbox number
        body=message_text,
        to=f'whatsapp:+91{to_number}'    # user's WhatsApp number
    )


@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")


@app.route("/book", methods=["POST"])
def book():
    name = request.form.get("name", "").strip()
    game = request.form.get("game", "").strip()
    date = request.form.get("date", "").strip()
    time_slot = request.form.get("time_slot", "").strip()
    phone = request.form.get("phone", "").strip()

    if not name or not game or not date or not time_slot:
        flash("Please fill in all required fields.", "error")
        return redirect(url_for("home"))

    add_booking_to_excel(name, game, date, time_slot, phone)

    flash("Your slot has been booked successfully!", "success")
    return redirect(url_for("home"))

 
   


if __name__ == "__main__":
    init_excel()
    app.run(debug=True)
