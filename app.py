from __future__ import annotations

import os
import re
import sqlite3
from datetime import datetime
from functools import wraps
from pathlib import Path
from typing import Any

from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "hotel_pms.db"
EXPORT_DIR = BASE_DIR / "exports"
EXPORT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key")

DEFAULT_USERS = [
    ("admin", generate_password_hash("admin123"), "admin"),
    ("staff", generate_password_hash("staff123"), "staff"),
]


# -----------------------------
# Database helpers
# -----------------------------
def get_db() -> sqlite3.Connection:
    if "db" not in g:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        g.db = conn
    return g.db


@app.teardown_appcontext
def close_db(_: Any) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


# -----------------------------
# Auth helpers
# -----------------------------
def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if session.get("user_id") is None:
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped_view


def admin_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)

    return wrapped_view


# -----------------------------
# Initialization
# -----------------------------
def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'staff')),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS resorts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            location TEXT,
            total_rooms INTEGER NOT NULL DEFAULT 0,
            phone TEXT,
            email TEXT,
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id TEXT UNIQUE NOT NULL,
            resort_id INTEGER NOT NULL,
            resort_name TEXT NOT NULL,
            guest_name TEXT NOT NULL,
            contact TEXT,
            email TEXT,
            checkin TEXT,
            checkout TEXT,
            nights INTEGER NOT NULL DEFAULT 0,
            room_type TEXT,
            rooms INTEGER NOT NULL DEFAULT 1,
            adults INTEGER NOT NULL DEFAULT 0,
            child_under_6 INTEGER NOT NULL DEFAULT 0,
            child_above_6 INTEGER NOT NULL DEFAULT 0,
            total_guests INTEGER NOT NULL DEFAULT 0,
            net_rate REAL NOT NULL DEFAULT 0,
            advance_paid REAL NOT NULL DEFAULT 0,
            balance_due REAL NOT NULL DEFAULT 0,
            payment_mode TEXT,
            booking_source TEXT,
            special_request TEXT,
            meal_plan TEXT,
            status TEXT NOT NULL DEFAULT 'Confirmed',
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (resort_id) REFERENCES resorts(id)
        )
        """
    )

    # New table for manual expense entry
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            resort_id INTEGER,
            category TEXT NOT NULL,
            amount REAL NOT NULL DEFAULT 0,
            notes TEXT,
            expense_date TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (resort_id) REFERENCES resorts(id)
        )
        """
    )

    for username, pwd_hash, role in DEFAULT_USERS:
        cur.execute(
            "INSERT OR IGNORE INTO users (username, password_hash, role) VALUES (?, ?, ?)",
            (username, pwd_hash, role),
        )

    conn.commit()
    conn.close()


# -----------------------------
# Utility helpers
# -----------------------------
DATE_PATTERNS = [
    "%d-%b-%y",
    "%d-%b-%Y",
    "%d/%m/%Y",
    "%d/%m/%y",
    "%Y-%m-%d",
    "%d %b %Y",
    "%d %b %y",
]


WORD_NUMBERS = {
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
}


def parse_date_value(raw: str | None) -> str:
    if not raw:
        return ""
    raw = raw.strip().replace(".", "-")
    for fmt in DATE_PATTERNS:
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def extract_date(text: str, labels: list[str]) -> str:
    for label in labels:
        pattern = rf"{label}\s*[:\-]?\s*([0-9]{{1,2}}[-/ ][A-Za-z]{{3,9}}[-/ ][0-9]{{2,4}}|[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}|[0-9]{{1,2}}/[0-9]{{1,2}}/[0-9]{{2,4}})"
        match = re.search(pattern, text, flags=re.I)
        if match:
            parsed = parse_date_value(match.group(1))
            if parsed:
                return parsed
    return ""


def word_to_number(token: str) -> int | None:
    token = token.strip().lower()
    if token.isdigit():
        return int(token)
    return WORD_NUMBERS.get(token)


def extract_first_amount(text: str, keywords: list[str]) -> float:
    for key in keywords:
        pattern = rf"{key}[^0-9₹]*₹?\s*([0-9,]+(?:\.\d{{1,2}})?)"
        match = re.search(pattern, text, flags=re.I)
        if match:
            return float(match.group(1).replace(",", ""))
    return 0.0


def next_booking_id(conn: sqlite3.Connection) -> str:
    row = conn.execute("SELECT id FROM bookings ORDER BY id DESC LIMIT 1").fetchone()
    next_id = 1 if row is None else int(row["id"]) + 1
    return f"BKG{next_id:05d}"


def compute_nights(checkin: str, checkout: str) -> int:
    if not checkin or not checkout:
        return 0
    try:
        d1 = datetime.strptime(checkin, "%Y-%m-%d")
        d2 = datetime.strptime(checkout, "%Y-%m-%d")
        return max((d2 - d1).days, 0)
    except ValueError:
        return 0


def normalize_booking_form(form: dict[str, Any], resort_name: str) -> dict[str, Any]:
    adults = int(form.get("adults") or 0)
    child_under = int(form.get("child_under_6") or 0)
    child_above = int(form.get("child_above_6") or 0)
    rooms = int(form.get("rooms") or 1)
    net_rate = float(form.get("net_rate") or 0)
    advance_paid = float(form.get("advance_paid") or 0)
    checkin = parse_date_value(form.get("checkin"))
    checkout = parse_date_value(form.get("checkout"))
    nights = compute_nights(checkin, checkout)
    total_guests = adults + child_under + child_above

    return {
        "resort_name": resort_name,
        "guest_name": (form.get("guest_name") or "").strip(),
        "contact": (form.get("contact") or "").strip(),
        "email": (form.get("email") or "").strip(),
        "checkin": checkin,
        "checkout": checkout,
        "nights": nights,
        "room_type": (form.get("room_type") or "").strip(),
        "rooms": rooms,
        "adults": adults,
        "child_under_6": child_under,
        "child_above_6": child_above,
        "total_guests": total_guests,
        "net_rate": net_rate,
        "advance_paid": advance_paid,
        "balance_due": max(net_rate - advance_paid, 0),
        "payment_mode": (form.get("payment_mode") or "").strip(),
        "booking_source": (form.get("booking_source") or "").strip(),
        "special_request": (form.get("special_request") or "").strip(),
        "meal_plan": (form.get("meal_plan") or "").strip(),
        "status": (form.get("status") or "Confirmed").strip(),
    }


def parse_raw_booking(text: str, resort_name: str) -> dict[str, Any]:
    raw = text.strip()
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    guest_name = parts[0] if parts else ""

    phone_match = re.search(r"\b\d{7,15}\b", raw)
    email_match = re.search(r"[\w.\-+]+@[\w\-]+(?:\.[\w\-]+)+", raw)

    checkin = extract_date(raw, [r"c/in", r"check\s*in", r"checkin"])
    checkout = extract_date(raw, [r"c/out", r"check\s*out", r"checkout"])

    adults_match = re.search(r"(\d+)\s*adults?", raw, flags=re.I)
    adults = int(adults_match.group(1)) if adults_match else 0

    child_under = 0
    child_above = 0
    age_matches = re.findall(r"ages?\s*(\d{1,2})", raw, flags=re.I)
    if age_matches:
        for age in map(int, age_matches):
            if age < 6:
                child_under += 1
            else:
                child_above += 1
    else:
        kids_match = re.search(r"(\w+|\d+)\s*kids?", raw, flags=re.I)
        if kids_match:
            kids_num = word_to_number(kids_match.group(1)) or 0
            child_above = kids_num

    rooms = 1
    rooms_match = re.search(r"(\d+)\s*rooms?", raw, flags=re.I)
    if rooms_match:
        rooms = int(rooms_match.group(1))
    else:
        word_room_match = re.search(r"(one|two|three|four|five|six|seven|eight|nine|ten)\s*rooms?", raw, flags=re.I)
        if word_room_match:
            rooms = word_to_number(word_room_match.group(1)) or 1

    net_rate = extract_first_amount(raw, [r"total\s*(?:is|amount)?", r"payment\s*total", r"net\s*rate", r"amount"])
    advance_paid = extract_first_amount(raw, [r"advance\s*paid", r"advance", r"paid\s*advance"])

    payment_mode = ""
    for mode in ["UPI", "Cash", "Card", "Bank Transfer", "GPay", "PhonePe"]:
        if mode.lower() in raw.lower():
            payment_mode = mode
            break

    booking_source = ""
    for source in ["Direct", "Booking.com", "Airbnb", "Agoda", "MakeMyTrip", "Goibibo"]:
        if source.lower() in raw.lower():
            booking_source = source
            break

    meal_plan = ""
    if re.search(r"breakfast\s*(?:incl|included|inclusive)", raw, flags=re.I):
        meal_plan = "Breakfast Included"

    special_bits = []
    early = re.search(r"early\s*check\s*in[^,]*", raw, flags=re.I)
    if early:
        special_bits.append(early.group(0).strip())
    if meal_plan:
        special_bits.append(meal_plan)

    nights = compute_nights(checkin, checkout)

    return {
        "resort_name": resort_name,
        "guest_name": guest_name,
        "contact": phone_match.group(0) if phone_match else "",
        "email": email_match.group(0) if email_match else "",
        "checkin": checkin,
        "checkout": checkout,
        "nights": nights,
        "room_type": "",
        "rooms": rooms,
        "adults": adults,
        "child_under_6": child_under,
        "child_above_6": child_above,
        "total_guests": adults + child_under + child_above,
        "net_rate": net_rate,
        "advance_paid": advance_paid,
        "balance_due": max(net_rate - advance_paid, 0),
        "payment_mode": payment_mode,
        "booking_source": booking_source,
        "special_request": "; ".join(special_bits) if special_bits else raw,
        "meal_plan": meal_plan,
        "status": "Confirmed",
    }


def insert_booking(conn: sqlite3.Connection, resort_id: int, data: dict[str, Any], created_by: str) -> None:
    booking_id = next_booking_id(conn)
    conn.execute(
        """
        INSERT INTO bookings (
            booking_id, resort_id, resort_name, guest_name, contact, email, checkin, checkout,
            nights, room_type, rooms, adults, child_under_6, child_above_6, total_guests,
            net_rate, advance_paid, balance_due, payment_mode, booking_source, special_request,
            meal_plan, status, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            booking_id,
            resort_id,
            data["resort_name"],
            data["guest_name"],
            data["contact"],
            data["email"],
            data["checkin"],
            data["checkout"],
            data["nights"],
            data["room_type"],
            data["rooms"],
            data["adults"],
            data["child_under_6"],
            data["child_above_6"],
            data["total_guests"],
            data["net_rate"],
            data["advance_paid"],
            data["balance_due"],
            data["payment_mode"],
            data["booking_source"],
            data["special_request"],
            data["meal_plan"],
            data["status"],
            created_by,
        ),
    )
    conn.commit()


# -----------------------------
# Routes
# -----------------------------
@app.route("/")
def index():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_db().execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()

        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("dashboard"))

        flash("Invalid username or password.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    db = get_db()
    resort_count = db.execute("SELECT COUNT(*) AS c FROM resorts").fetchone()["c"]
    booking_count = db.execute("SELECT COUNT(*) AS c FROM bookings").fetchone()["c"]
    revenue = db.execute("SELECT COALESCE(SUM(net_rate), 0) AS s FROM bookings").fetchone()["s"]
    pending = db.execute("SELECT COALESCE(SUM(balance_due), 0) AS s FROM bookings").fetchone()["s"]
    recent_bookings = db.execute(
        "SELECT booking_id, resort_name, guest_name, checkin, net_rate, status FROM bookings ORDER BY id DESC LIMIT 8"
    ).fetchall()
    return render_template(
        "dashboard.html",
        resort_count=resort_count,
        booking_count=booking_count,
        revenue=revenue,
        pending=pending,
        recent_bookings=recent_bookings,
    )


@app.route("/resorts", methods=["GET", "POST"])
@login_required
@admin_required
def resorts():
    db = get_db()
    if request.method == "POST":
        try:
            db.execute(
                "INSERT INTO resorts (name, location, total_rooms, phone, email, notes) VALUES (?, ?, ?, ?, ?, ?)",
                (
                    request.form.get("name", "").strip(),
                    request.form.get("location", "").strip(),
                    int(request.form.get("total_rooms") or 0),
                    request.form.get("phone", "").strip(),
                    request.form.get("email", "").strip(),
                    request.form.get("notes", "").strip(),
                ),
            )
            db.commit()
            flash("Resort added successfully.", "success")
        except sqlite3.IntegrityError:
            flash("Resort name already exists.", "error")
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()
    return render_template("resorts.html", resorts=resorts_rows)


@app.route("/resort/edit/<int:resort_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_resort(resort_id):
    db = get_db()

    resort = db.execute(
        "SELECT * FROM resorts WHERE id = ?", (resort_id,)
    ).fetchone()

    if not resort:
        flash("Resort not found.", "error")
        return redirect(url_for("resorts"))

    if request.method == "POST":
        try:
            db.execute("""
                UPDATE resorts SET
                name = ?,
                location = ?,
                total_rooms = ?,
                phone = ?,
                email = ?,
                notes = ?
                WHERE id = ?
            """, (
                request.form.get("name", "").strip(),
                request.form.get("location", "").strip(),
                int(request.form.get("total_rooms") or 0),
                request.form.get("phone", "").strip(),
                request.form.get("email", "").strip(),
                request.form.get("notes", "").strip(),
                resort_id
            ))

            db.commit()
            flash("Resort updated successfully.", "success")
            return redirect(url_for("resorts"))

        except Exception:
            flash("Error updating resort.", "error")

    return render_template("edit_resort.html", resort=resort)


@app.route("/booking/new", methods=["GET", "POST"])
@login_required
def booking_new():
    db = get_db()
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()
    if request.method == "POST":
        resort_id = int(request.form.get("resort_id") or 0)
        resort = db.execute("SELECT * FROM resorts WHERE id = ?", (resort_id,)).fetchone()
        if not resort:
            flash("Please choose a valid resort.", "error")
            return render_template("booking_form.html", resorts=resorts_rows)

        data = normalize_booking_form(request.form, resort["name"])
        if not data["guest_name"]:
            flash("Guest name is required.", "error")
            return render_template("booking_form.html", resorts=resorts_rows)

        insert_booking(db, resort_id, data, session.get("username", "system"))
        flash("Booking saved successfully.", "success")
        return redirect(url_for("search_bookings"))

    return render_template("booking_form.html", resorts=resorts_rows)


@app.route("/booking/quick", methods=["GET", "POST"])
@login_required
def booking_quick():
    db = get_db()
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()

    if request.method == "POST":
        resort_id = int(request.form.get("resort_id") or 0)
        raw_text = request.form.get("raw_text", "")
        resort = db.execute("SELECT * FROM resorts WHERE id = ?", (resort_id,)).fetchone()
        if not resort:
            flash("Please choose a valid resort.", "error")
            return render_template("booking_quick.html", resorts=resorts_rows)
        if not raw_text.strip():
            flash("Paste the raw booking text.", "error")
            return render_template("booking_quick.html", resorts=resorts_rows)

        data = parse_raw_booking(raw_text, resort["name"])
        return render_template(
            "booking_quick_preview.html",
            resorts=resorts_rows,
            resort_id=resort_id,
            raw_text=raw_text,
            parsed=data,
        )

    return render_template("booking_quick.html", resorts=resorts_rows)


@app.route("/booking/quick/confirm", methods=["POST"])
@login_required
def booking_quick_confirm():
    db = get_db()
    resort_id = int(request.form.get("resort_id") or 0)
    resort = db.execute("SELECT * FROM resorts WHERE id = ?", (resort_id,)).fetchone()
    if not resort:
        flash("Invalid resort selected.", "error")
        return redirect(url_for("booking_quick"))

    data = normalize_booking_form(request.form, resort["name"])
    if not data["guest_name"]:
        flash("Guest name is required before confirming.", "error")
        return redirect(url_for("booking_quick"))

    insert_booking(db, resort_id, data, session.get("username", "system"))
    flash("Quick entry booking saved.", "success")
    return redirect(url_for("search_bookings"))


@app.route("/search")
@login_required
def search_bookings():
    db = get_db()
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()

    resort_id = request.args.get("resort_id", "").strip()
    guest = request.args.get("guest", "").strip()
    contact = request.args.get("contact", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()

    query = "SELECT * FROM bookings WHERE 1=1"
    params: list[Any] = []

    if resort_id:
        query += " AND resort_id = ?"
        params.append(int(resort_id))
    if guest:
        query += " AND guest_name LIKE ?"
        params.append(f"%{guest}%")
    if contact:
        query += " AND contact LIKE ?"
        params.append(f"%{contact}%")
    if date_from:
        query += " AND checkin >= ?"
        params.append(date_from)
    if date_to:
        query += " AND checkin <= ?"
        params.append(date_to)

    query += " ORDER BY checkin DESC, id DESC"
    rows = db.execute(query, params).fetchall()
    return render_template("search.html", resorts=resorts_rows, rows=rows)


@app.route("/booking/edit/<int:booking_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_booking(booking_id):
    db = get_db()

    booking = db.execute(
        "SELECT * FROM bookings WHERE id = ?", (booking_id,)
    ).fetchone()

    if not booking:
        flash("Booking not found.", "error")
        return redirect(url_for("search_bookings"))

    resorts = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()

    if request.method == "POST":
        data = normalize_booking_form(request.form, booking["resort_name"])

        db.execute("""
            UPDATE bookings SET
            guest_name = ?,
            contact = ?,
            email = ?,
            checkin = ?,
            checkout = ?,
            nights = ?,
            room_type = ?,
            rooms = ?,
            adults = ?,
            child_under_6 = ?,
            child_above_6 = ?,
            total_guests = ?,
            net_rate = ?,
            advance_paid = ?,
            balance_due = ?,
            payment_mode = ?,
            booking_source = ?,
            special_request = ?,
            meal_plan = ?,
            status = ?
            WHERE id = ?
        """, (
            data["guest_name"],
            data["contact"],
            data["email"],
            data["checkin"],
            data["checkout"],
            data["nights"],
            data["room_type"],
            data["rooms"],
            data["adults"],
            data["child_under_6"],
            data["child_above_6"],
            data["total_guests"],
            data["net_rate"],
            data["advance_paid"],
            data["balance_due"],
            data["payment_mode"],
            data["booking_source"],
            data["special_request"],
            data["meal_plan"],
            data["status"],
            booking_id
        ))

        db.commit()
        flash("Booking updated successfully.", "success")
        return redirect(url_for("search_bookings"))

    return render_template("edit_booking.html", booking=booking, resorts=resorts)


@app.route("/pnl")
@login_required
@admin_required
def pnl():
    db = get_db()
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()
    resort_id = request.args.get("resort_id", "").strip()
    month = request.args.get("month", "").strip()  # YYYY-MM

    query = "SELECT * FROM bookings WHERE 1=1"
    params: list[Any] = []

    if resort_id:
        query += " AND resort_id = ?"
        params.append(int(resort_id))
    if month:
        query += " AND substr(checkin, 1, 7) = ?"
        params.append(month)

    rows = db.execute(query, params).fetchall()

    total_bookings = len(rows)
    total_revenue = sum(float(r["net_rate"] or 0) for r in rows)
    total_advance = sum(float(r["advance_paid"] or 0) for r in rows)
    total_balance = sum(float(r["balance_due"] or 0) for r in rows)
    total_nights = sum(int(r["nights"] or 0) for r in rows)
    avg_booking = round(total_revenue / total_bookings, 2) if total_bookings else 0

    return render_template(
        "pnl.html",
        resorts=resorts_rows,
        total_bookings=total_bookings,
        total_revenue=total_revenue,
        total_advance=total_advance,
        total_balance=total_balance,
        total_nights=total_nights,
        avg_booking=avg_booking,
        month=month,
        rows=rows,
    )


@app.route("/expenses", methods=["GET", "POST"])
@login_required
@admin_required
def expenses():
    db = get_db()
    resorts_rows = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()

    if request.method == "POST":
        try:
            db.execute(
                """
                INSERT INTO expenses (resort_id, category, amount, notes, expense_date)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    int(request.form.get("resort_id") or 0) or None,
                    request.form.get("category", "").strip(),
                    float(request.form.get("amount") or 0),
                    request.form.get("notes", "").strip(),
                    request.form.get("expense_date", "").strip(),
                ),
            )
            db.commit()
            flash("Expense added successfully.", "success")
            return redirect(url_for("expenses"))
        except Exception:
            flash("Error adding expense.", "error")

    rows = db.execute(
        """
        SELECT e.*, r.name AS resort_name
        FROM expenses e
        LEFT JOIN resorts r ON r.id = e.resort_id
        ORDER BY e.id DESC
        """
    ).fetchall()

    return render_template("expenses.html", resorts=resorts_rows, rows=rows)


# -----------------------------
# Availability Calendar
# -----------------------------
@app.route("/availability")
@login_required
def availability():
    db = get_db()

    resort_id = request.args.get("resort_id", "")
    month = request.args.get("month", "")

    resorts = db.execute("SELECT * FROM resorts ORDER BY name").fetchall()

    availability_data = []

    if resort_id and month:
        resort = db.execute("SELECT * FROM resorts WHERE id=?", (resort_id,)).fetchone()

        if resort:
            total_rooms = resort["total_rooms"]

            bookings = db.execute("""
                SELECT checkin, checkout, rooms FROM bookings
                WHERE resort_id = ?
            """, (resort_id,)).fetchall()

            from datetime import datetime, timedelta

            year, m = map(int, month.split("-"))
            start_date = datetime(year, m, 1)

            if m == 12:
                end_date = datetime(year+1, 1, 1)
            else:
                end_date = datetime(year, m+1, 1)

            current = start_date

            while current < end_date:
                booked_rooms = 0

                for b in bookings:
                    if b["checkin"] and b["checkout"]:
                        cin = datetime.strptime(b["checkin"], "%Y-%m-%d")
                        cout = datetime.strptime(b["checkout"], "%Y-%m-%d")

                        if cin <= current < cout:
                            booked_rooms += b["rooms"]

                availability_data.append({
                    "date": current.strftime("%Y-%m-%d"),
                    "booked": booked_rooms,
                    "available": max(total_rooms - booked_rooms, 0)
                })

                current += timedelta(days=1)

    return render_template(
        "availability.html",
        resorts=resorts,
        availability=availability_data,
        selected_resort=resort_id,
        selected_month=month
    )


@app.route("/export/bookings.xlsx")
@login_required
@admin_required
def export_bookings():
    db = get_db()
    rows = db.execute("SELECT * FROM bookings ORDER BY id DESC").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bookings"
    headers = [
        "Booking ID",
        "Resort Name",
        "Guest Name",
        "Contact",
        "Email",
        "Check-in",
        "Check-out",
        "Nights",
        "Room Type",
        "Rooms",
        "Adults",
        "Children <6",
        "Children >6",
        "Total Guests",
        "Net Rate",
        "Advance Paid",
        "Balance Due",
        "Payment Mode",
        "Booking Source",
        "Special Request",
        "Meal Plan",
        "Status",
        "Created By",
        "Created At",
    ]
    ws.append(headers)

    for row in rows:
        ws.append([
            row["booking_id"],
            row["resort_name"],
            row["guest_name"],
            row["contact"],
            row["email"],
            row["checkin"],
            row["checkout"],
            row["nights"],
            row["room_type"],
            row["rooms"],
            row["adults"],
            row["child_under_6"],
            row["child_above_6"],
            row["total_guests"],
            row["net_rate"],
            row["advance_paid"],
            row["balance_due"],
            row["payment_mode"],
            row["booking_source"],
            row["special_request"],
            row["meal_plan"],
            row["status"],
            row["created_by"],
            row["created_at"],
        ])

    out_path = EXPORT_DIR / "bookings_export.xlsx"
    wb.save(out_path)
    return send_file(out_path, as_attachment=True)


@app.route("/export/pnl.xlsx")
@login_required
@admin_required
def export_pnl():
    db = get_db()

    resort_id = request.args.get("resort_id", "").strip()
    month = request.args.get("month", "").strip()

    booking_query = "SELECT * FROM bookings WHERE 1=1"
    booking_params: list[Any] = []

    if resort_id:
        booking_query += " AND resort_id = ?"
        booking_params.append(int(resort_id))
    if month:
        booking_query += " AND substr(checkin, 1, 7) = ?"
        booking_params.append(month)

    bookings = db.execute(booking_query, booking_params).fetchall()

    expense_query = "SELECT * FROM expenses WHERE 1=1"
    expense_params: list[Any] = []

    if resort_id:
        expense_query += " AND resort_id = ?"
        expense_params.append(int(resort_id))
    if month:
        expense_query += " AND substr(expense_date, 1, 7) = ?"
        expense_params.append(month)

    expenses = db.execute(expense_query, expense_params).fetchall()

    room_revenue = sum(float(r["net_rate"] or 0) for r in bookings)

    category_totals: dict[str, float] = {}
    for row in expenses:
        category = row["category"] or "Other"
        category_totals[category] = category_totals.get(category, 0) + float(row["amount"] or 0)

    food_cost = category_totals.get("Food Cost", 0.0)
    kitchen = category_totals.get("Kitchen", 0.0)
    amenities = category_totals.get("Amenities", 0.0)
    laundry = category_totals.get("Laundry", 0.0)
    salary = category_totals.get("Salary", 0.0)
    staff_welfare = category_totals.get("Staff Welfare", 0.0)
    ota = category_totals.get("OTA", 0.0)
    utilities = category_totals.get("Utilities", 0.0)
    maintenance = category_totals.get("Maintenance", 0.0)
    marketing = category_totals.get("Marketing", 0.0)
    rent = category_totals.get("Rent", 0.0)
    depreciation = category_totals.get("Depreciation", 0.0)
    other_income = category_totals.get("Other Income", 0.0)
    fnb_revenue = category_totals.get("F&B Revenue", 0.0)

    total_revenue = room_revenue + fnb_revenue + other_income
    total_direct = food_cost + kitchen + amenities + laundry
    total_staff = salary + staff_welfare
    total_overheads = ota + utilities + maintenance + marketing
    ebitda = total_revenue - (total_direct + total_staff + total_overheads)
    total_fixed = rent + depreciation
    net_profit = ebitda - total_fixed

    wb = Workbook()
    ws = wb.active
    ws.title = "Monthly P&L"

    ws.append(["Monthly Resort P&L & Performance Report"])
    ws.append(["For a Rooms & Meals focused property in Kerala"])
    ws.append([])
    ws.append(["PARTICULARS", "AMOUNT (INR)", "% OF TOTAL REV"])

    def pct(value: float) -> str:
        if total_revenue <= 0:
            return ""
        return f"{round((value / total_revenue) * 100, 1)}%"

    ws.append(["I. REVENUE (Net of GST)", "", ""])
    ws.append(["Room Revenue", room_revenue, pct(room_revenue)])
    ws.append(["Food & Beverage Revenue", fnb_revenue, pct(fnb_revenue)])
    ws.append(["Other Income", other_income, pct(other_income)])
    ws.append(["TOTAL REVENUE (A)", total_revenue, "100%"])

    ws.append([])
    ws.append(["II. DIRECT OPERATING COSTS", "", ""])
    ws.append(["Cost of Food & Provisions (COGS)", food_cost, ""])
    ws.append(["Kitchen Consumables & Gas", kitchen, ""])
    ws.append(["Guest Amenities & Toiletries", amenities, ""])
    ws.append(["Laundry & Linen Expenses", laundry, ""])
    ws.append(["TOTAL DIRECT COSTS (B)", total_direct, pct(total_direct)])

    ws.append([])
    ws.append(["III. STAFF COSTS", "", ""])
    ws.append(["Salaries & Wages", salary, ""])
    ws.append(["Staff Welfare & Meals", staff_welfare, ""])
    ws.append(["TOTAL STAFF COSTS (C)", total_staff, pct(total_staff)])

    ws.append([])
    ws.append(["IV. OVERHEADS", "", ""])
    ws.append(["OTA Commissions", ota, ""])
    ws.append(["Utilities", utilities, ""])
    ws.append(["Repairs, Maintenance & AMC", maintenance, ""])
    ws.append(["Marketing & Local Admin Licenses", marketing, ""])
    ws.append(["TOTAL OVERHEADS (D)", total_overheads, pct(total_overheads)])

    ws.append([])
    ws.append(["V. OPERATING PROFIT (EBITDA)", ebitda, pct(ebitda)])

    ws.append([])
    ws.append(["VI. FIXED CHARGES", "", ""])
    ws.append(["Rent / Lease / Interest", rent, ""])
    ws.append(["Depreciation & Amortization", depreciation, ""])
    ws.append(["TOTAL FIXED CHARGES (E)", total_fixed, pct(total_fixed)])

    ws.append([])
    ws.append(["VII. NET PROFIT / (LOSS) (V - E)", net_profit, pct(net_profit)])

    file_path = EXPORT_DIR / "monthly_pnl.xlsx"
    wb.save(file_path)
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
