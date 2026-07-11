"""
MHT-CET College Predictor 2026 - Full Website (single-file run)
Run:  python app.py
Then open http://127.0.0.1:5000

Data: place your cutoff sheet at  data/Final MH-CET_Cutoff.xlsx
Default admin login is printed on first run.
"""

import io
import json
import math
import os
import re
import sqlite3
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import (Flask, g, jsonify, redirect, render_template, request,
                   send_file, session, url_for)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(BASE_DIR, "predictor.db")
DEFAULT_XLSX = os.path.join(DATA_DIR, "Final MH-CET_Cutoff.xlsx")
DEFAULT_CSV = os.path.join(DATA_DIR, "cutoff_data.csv.gz")
# Prefer the slim compressed CSV when present - loads with far less memory
# than xlsx (important on 512MB hosts). Admin xlsx uploads still work.
DEFAULT_DATA = DEFAULT_CSV if os.path.exists(DEFAULT_CSV) else DEFAULT_XLSX

DEFAULT_ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "admin@mhtcet.local")
DEFAULT_ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

ROUND_PCT_COLS = ["Round 1Percentile", "Round 2 Percentile",
                  "Round 3 Percentile", "Round 4 Percentile"]
ROUND_RANK_COLS = ["Round 1 Rank", "Round 2 Rank",
                   "Round 3 Rank", "Round 4 Rank"]

DEFAULT_SETTINGS = {
    "pct_band": 2.0,
    "priority_codes": [16006, 3012, 6271, 3215, 3119, 6273, 6276, 6175,
                       6007, 6072],
    "zone_safe": 1.5,     # gap >= safe  -> Safe
    "zone_ambitious": -1.0,  # gap >= this (and < 0) -> Ambitious, else Reach
    "registration_open": 1,
    "active_data_file": DEFAULT_DATA,
    "data_year": "2025 (Latest)",
    "payment_amount": "99",
}

# ----------------------------------------------------------------- database
def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exc):
    d = g.pop("db", None)
    if d is not None:
        d.close()


def init_db():
    con = sqlite3.connect(DB_PATH)
    con.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        disabled INTEGER NOT NULL DEFAULT 0,
        approved INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT NOT NULL,
        txn_ref TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS prediction_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        mode TEXT, value REAL, category TEXT,
        branches TEXT, districts TEXT,
        results INTEGER, created_at TEXT
    );
    """)
    # Migration: add the approved column to older databases.
    # Existing accounts are grandfathered in as approved so nobody
    # already using the site gets locked out.
    cols = [r[1] for r in con.execute("PRAGMA table_info(users)")]
    if "approved" not in cols:
        con.execute("ALTER TABLE users ADD COLUMN approved "
                    "INTEGER NOT NULL DEFAULT 0")
        con.execute("UPDATE users SET approved = 1")

    cur = con.execute("SELECT COUNT(*) c FROM users WHERE role='admin'")
    if cur.fetchone()[0] == 0:
        con.execute(
            "INSERT INTO users (name,email,password_hash,role,approved,"
            "created_at) VALUES (?,?,?,?,?,?)",
            ("Administrator", DEFAULT_ADMIN_EMAIL,
             generate_password_hash(DEFAULT_ADMIN_PASSWORD), "admin", 1,
             datetime.now().isoformat()))
        print("=" * 60)
        print("Default admin account created:")
        print(f"  email:    {DEFAULT_ADMIN_EMAIL}")
        print(f"  password: {DEFAULT_ADMIN_PASSWORD}")
        print("=" * 60)
    for k, v in DEFAULT_SETTINGS.items():
        con.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)",
                    (k, json.dumps(v)))
    con.commit()
    con.close()


def get_setting(key):
    row = db().execute("SELECT value FROM settings WHERE key=?",
                       (key,)).fetchone()
    return json.loads(row["value"]) if row else DEFAULT_SETTINGS.get(key)


def set_setting(key, value):
    db().execute("INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)",
                 (key, json.dumps(value)))
    db().commit()


# ---------------------------------------------------------------- data load
_DATA_CACHE = {"path": None, "mtime": None, "df": None}


def parse_code(code: str):
    """Return (gender, base_category) from a CAP seat code."""
    if code.startswith("PWD") or code.startswith("DEF"):
        return "Any", code
    gender, base = "Any", code
    if code.startswith("G"):
        gender, base = "Gender-Neutral", code[1:]
    elif code.startswith("L"):
        gender, base = "Female (Ladies)", code[1:]
    if base not in ("EWS", "TFWS", "MI", "ORPHAN"):
        base = re.sub(r"[SHO]$", "", base)
    return gender, base


def load_data() -> pd.DataFrame:
    path = get_setting("active_data_file")
    if not os.path.exists(path):
        path = DEFAULT_DATA
    mtime = os.path.getmtime(path)
    if _DATA_CACHE["path"] == path and _DATA_CACHE["mtime"] == mtime:
        return _DATA_CACHE["df"]

    if path.endswith((".csv", ".csv.gz")):
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    if "Round 1Percentile" not in df.columns and \
            "Round 1 Percentile" in df.columns:
        df = df.rename(columns={"Round 1 Percentile": "Round 1Percentile"})

    df = df.dropna(subset=["Category"])
    df["Category"] = df["Category"].astype(str).str.strip()
    parsed = df["Category"].map(parse_code)
    df["Gender"] = parsed.map(lambda t: t[0])
    df["Base Category"] = parsed.map(lambda t: t[1])

    pct_cols = [c for c in ROUND_PCT_COLS if c in df.columns]
    rank_cols = [c for c in ROUND_RANK_COLS if c in df.columns]
    df["Cutoff Percentile"] = df[pct_cols].min(axis=1)
    df["Cutoff Rank"] = df[rank_cols].max(axis=1)

    df = df.dropna(subset=["Cutoff Percentile"])
    _DATA_CACHE.update({"path": path, "mtime": mtime, "df": df})
    return df


# ------------------------------------------------------------- predictions
def estimate_merit_rank(df, p):
    pairs = df[["Cutoff Percentile", "Cutoff Rank"]].dropna()
    if pairs.empty:
        return None
    nearest = pairs.iloc[(pairs["Cutoff Percentile"] - p).abs()
                         .argsort()[:25]]
    return int(nearest["Cutoff Rank"].median())


def estimate_percentile_from_rank(df, r):
    pairs = df[["Cutoff Percentile", "Cutoff Rank"]].dropna()
    if pairs.empty:
        return None
    nearest = pairs.iloc[(pairs["Cutoff Rank"] - r).abs().argsort()[:25]]
    return float(nearest["Cutoff Percentile"].median())


def classify_college(cutoff, percentile):
    if cutoff > percentile:
        return "Dream College"
    if cutoff >= percentile - 1:
        return "Target College"
    return "Safe College"


def run_prediction(form):
    df = load_data()
    band = float(get_setting("pct_band"))
    priority_codes = [int(c) for c in get_setting("priority_codes")]

    mode = form.get("mode", "percentile")
    if mode == "rank":
        rank_in = int(form.get("value") or 0)
        if rank_in <= 0:
            return {"error": "Enter a valid rank."}
        percentile = estimate_percentile_from_rank(df, rank_in)
        entered = f"Rank {rank_in:,}"
        counter_label = "Your Approx. Percentile"
        counter_value = f"~{percentile:.2f}"
    else:
        percentile = float(form.get("value") or -1)
        if not 0 <= percentile <= 100:
            return {"error": "Enter a valid percentile (0-100)."}
        est = estimate_merit_rank(df, percentile)
        entered = f"Percentile {percentile}"
        counter_label = "Your Approx. Merit Rank"
        counter_value = f"~{est:,}" if est else "N/A"

    d = df[df["Base Category"] == form.get("category", "OPEN")]

    gender = form.get("gender", "Gender-Neutral")
    if gender != "Any":
        if gender == "Female (Ladies)":
            d = d[d["Gender"].isin(
                ["Female (Ladies)", "Gender-Neutral", "Any"])]
        else:
            d = d[d["Gender"].isin([gender, "Any"])]

    branches = form.get("branches") or []
    districts = form.get("districts") or []
    quota = form.get("quota", "All Quotas")
    if branches:
        d = d[d["Course Name"].isin(branches)]
    if quota != "All Quotas":
        d = d[d["Level"] == quota]
    if districts:
        d = d[d["District"].isin(districts)]

    lo, hi = percentile - band, percentile + band
    in_band = d[d["Cutoff Percentile"].between(lo, hi)]
    prio = d[d["Institute Code"].isin(priority_codes)]
    combined = pd.concat([prio, in_band]).drop_duplicates().copy()

    if combined.empty:
        return {"entered": entered, "counter_label": counter_label,
                "counter_value": counter_value, "results": []}

    rank_map = {c: i for i, c in enumerate(priority_codes)}
    combined["_prio"] = combined["Institute Code"].map(
        lambda c: rank_map.get(c, len(priority_codes)))
    combined["college_type"] = combined["Cutoff Percentile"].map(
        lambda c: classify_college(c, percentile))
    combined = combined.sort_values(["_prio", "Cutoff Percentile"],
                                    ascending=[True, False])

    results = []
    for _, r in combined.iterrows():
        results.append({
            "priority": bool(r["_prio"] < len(priority_codes)),
            "code": int(r["Institute Code"]),
            "institute": str(r["Institute Name"]),
            "district": str(r["District"]),
            "course": str(r["Course Name"]),
            "seat_code": str(r["Category"]),
            "quota": str(r["Level"]),
            "cutoff_pct": round(float(r["Cutoff Percentile"]), 4),
            "cutoff_rank": (int(r["Cutoff Rank"])
                            if pd.notna(r["Cutoff Rank"]) else None),
            "college_type": r["college_type"],
        })

    tc = combined["college_type"].value_counts().to_dict()
    return {"entered": entered, "counter_label": counter_label,
            "counter_value": counter_value,
            "percentile": round(percentile, 4),
            "priority_count": int((combined["_prio"] <
                                   len(priority_codes)).sum()),
            "type_counts": {"dream": tc.get("Dream College", 0),
                            "target": tc.get("Target College", 0),
                            "safe": tc.get("Safe College", 0)},
            "results": results}


# ----------------------------------------------------------------- auth
def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **k):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*a, **k)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **k):
        if session.get("role") != "admin":
            return redirect(url_for("login"))
        return fn(*a, **k)
    return wrapper


# ----------------------------------------------------------------- routes
@app.route("/")
def root():
    if "user_id" in session:
        return redirect(url_for("features"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        action = request.form.get("action")
        email = request.form.get("email", "").strip().lower()
        pw = request.form.get("password", "")
        if action == "register":
            if not get_setting("registration_open"):
                error = "Registration is currently closed."
            elif not email or not pw or not request.form.get("name"):
                error = "All fields are required."
            else:
                try:
                    db().execute(
                        "INSERT INTO users (name,email,password_hash,role,"
                        "approved,created_at) VALUES (?,?,?,?,?,?)",
                        (request.form["name"].strip(), email,
                         generate_password_hash(pw), "user", 0,
                         datetime.now().isoformat()))
                    db().commit()
                    return redirect(url_for("payment", email=email))
                except sqlite3.IntegrityError:
                    error = "Email already registered."
        else:
            row = db().execute("SELECT * FROM users WHERE email=?",
                               (email,)).fetchone()
            if row and not row["disabled"] and \
                    check_password_hash(row["password_hash"], pw):
                if not row["approved"]:
                    error = ("Your account is not activated yet. Complete "
                             "the payment and wait for admin approval. "
                             "<a href='/payment'>Payment details</a>")
                else:
                    session["user_id"] = row["id"]
                    session["name"] = row["name"]
                    session["role"] = row["role"]
                    return redirect(url_for("features"))
            else:
                error = "Invalid credentials or account disabled."
    return render_template("login.html", error=error)


@app.route("/payment", methods=["GET", "POST"])
def payment():
    msg = None
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        txn = request.form.get("txn_ref", "").strip()
        if email and txn:
            db().execute(
                "INSERT INTO payments (email,txn_ref,created_at) "
                "VALUES (?,?,?)",
                (email, txn, datetime.now().isoformat()))
            db().commit()
            msg = ("Payment reference submitted. The admin will verify it "
                   "and activate your account shortly.")
        else:
            msg = "Please fill in both your email and the transaction ID."
    return render_template(
        "payment.html", msg=msg,
        amount=get_setting("payment_amount"),
        email=request.args.get("email", ""))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/features")
@login_required
def features():
    df = load_data()
    stats = {"rows": len(df),
             "institutes": df["Institute Name"].nunique(),
             "courses": df["Course Name"].nunique(),
             "districts": df["District"].nunique()}
    return render_template("features.html", stats=stats)


@app.route("/predictor")
@login_required
def predictor():
    df = load_data()
    return render_template(
        "predictor.html",
        categories=sorted(df["Base Category"].unique()),
        branches=sorted(df["Course Name"].dropna().unique()),
        quotas=sorted(df["Level"].dropna().unique()),
        districts=sorted(df["District"].dropna().unique()))


@app.route("/api/predict", methods=["POST"])
@login_required
def api_predict():
    payload = request.get_json(force=True)
    out = run_prediction(payload)
    if "error" not in out:
        db().execute(
            "INSERT INTO prediction_log (user_id,mode,value,category,"
            "branches,districts,results,created_at) VALUES (?,?,?,?,?,?,?,?)",
            (session["user_id"], payload.get("mode"),
             float(payload.get("value") or 0), payload.get("category"),
             json.dumps(payload.get("branches") or []),
             json.dumps(payload.get("districts") or []),
             len(out.get("results", [])), datetime.now().isoformat()))
        db().commit()
    return jsonify(out)


@app.route("/api/report", methods=["POST"])
@login_required
def api_report():
    payload = request.get_json(force=True)
    out = run_prediction(payload)
    if out.get("error") or not out.get("results"):
        return jsonify({"error": out.get("error", "No results to export.")}), 400
    pdf = build_pdf(out, payload)
    fname = f"MHCET_Prediction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                     as_attachment=True, download_name=fname)


def build_pdf(out, payload):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm)
    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=7.5,
                          leading=9)
    head = ParagraphStyle("head", parent=styles["Normal"], fontSize=8,
                          leading=10, textColor=colors.white,
                          fontName="Helvetica-Bold")
    story = [
        Paragraph("MHT-CET College Predictor Report", styles["Title"]),
        Paragraph(
            f"Entered: <b>{out['entered']}</b> "
            f"({out['counter_label']}: {out['counter_value']}) | "
            f"Category: <b>{payload.get('category')}</b> | "
            f"Gender: <b>{payload.get('gender')}</b> | "
            f"Options: <b>{len(out['results'])}</b> | "
            f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')} | "
            f"* = priority institute", styles["Normal"]),
        Spacer(1, 6 * mm)]

    headers = ["Sr.No", "Institute Code", "Institute Name", "District",
               "Course Name", "Quota", "Cutoff Percentile", "Cutoff Rank"]
    data = [[Paragraph(h, head) for h in headers]]
    for i, r in enumerate(out["results"], 1):
        star = "* " if r["priority"] else ""
        data.append([
            Paragraph(str(i), cell),
            Paragraph(str(r["code"]), cell),
            Paragraph(star + r["institute"], cell),
            Paragraph(r["district"], cell),
            Paragraph(r["course"], cell),
            Paragraph(r["quota"], cell),
            Paragraph(f"{r['cutoff_pct']:.4f}", cell),
            Paragraph(f"{r['cutoff_rank']:,}" if r["cutoff_rank"] else "-",
                      cell)])
    table = Table(data, colWidths=[11 * mm, 18 * mm, 70 * mm, 24 * mm,
                                   56 * mm, 48 * mm, 21 * mm, 21 * mm],
                  repeatRows=1)
    cmds = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#eef3f8")])]
    for i, r in enumerate(out["results"], 1):
        if r["priority"]:
            cmds.append(("BACKGROUND", (0, i), (-1, i),
                         colors.HexColor("#fff3cd")))
    table.setStyle(TableStyle(cmds))
    story.append(table)
    doc.build(story)
    return buf.getvalue()


# ----------------------------------------------------------------- admin
@app.route("/admin")
@admin_required
def admin():
    df = load_data()
    payments = db().execute(
        "SELECT * FROM payments ORDER BY id DESC LIMIT 50").fetchall()
    users = db().execute(
        "SELECT id,name,email,role,disabled,approved,created_at FROM users "
        "ORDER BY approved ASC, id ASC").fetchall()
    pending = sum(1 for u in users if not u["approved"])
    logs = db().execute(
        "SELECT COUNT(*) c FROM prediction_log").fetchone()["c"]
    top_branches = db().execute(
        "SELECT branches, COUNT(*) c FROM prediction_log "
        "WHERE branches != '[]' GROUP BY branches ORDER BY c DESC LIMIT 5"
    ).fetchall()
    data_info = {"path": os.path.basename(get_setting("active_data_file")),
                 "rows": len(df),
                 "institutes": df["Institute Name"].nunique(),
                 "courses": df["Course Name"].nunique()}
    settings = {"pct_band": get_setting("pct_band"),
                "priority_codes": ", ".join(
                    str(c) for c in get_setting("priority_codes")),
                "zone_safe": get_setting("zone_safe"),
                "zone_ambitious": get_setting("zone_ambitious"),
                "registration_open": get_setting("registration_open"),
                "data_year": get_setting("data_year"),
                "payment_amount": get_setting("payment_amount")}
    return render_template("admin.html", users=users, data_info=data_info,
                           settings=settings, total_predictions=logs,
                           top_branches=top_branches, pending=pending,
                           payments=payments)


@app.route("/admin/upload", methods=["POST"])
@admin_required
def admin_upload():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return redirect(url_for("admin"))
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(DATA_DIR, f"cutoff_{ts}.xlsx")
    f.save(path)
    try:  # validate before activating
        test = pd.read_excel(path)
        test.columns = [c.strip() for c in test.columns]
        required = {"Institute Code", "Institute Name", "District",
                    "Course Name", "Level", "Category"}
        missing = required - set(test.columns)
        if missing:
            os.remove(path)
            return render_template("admin_error.html",
                                   msg=f"Missing columns: {missing}")
        set_setting("active_data_file", path)
    except Exception as e:
        os.remove(path)
        return render_template("admin_error.html", msg=str(e))
    return redirect(url_for("admin"))


@app.route("/admin/settings", methods=["POST"])
@admin_required
def admin_settings():
    set_setting("pct_band", float(request.form.get("pct_band", 2)))
    codes = [int(x) for x in
             re.findall(r"\d+", request.form.get("priority_codes", ""))]
    set_setting("priority_codes", codes)
    set_setting("zone_safe", float(request.form.get("zone_safe", 1.5)))
    set_setting("zone_ambitious",
                float(request.form.get("zone_ambitious", -1)))
    set_setting("registration_open",
                1 if request.form.get("registration_open") else 0)
    set_setting("data_year", request.form.get("data_year", "2025 (Latest)"))
    set_setting("payment_amount",
                request.form.get("payment_amount", "99").strip())
    return redirect(url_for("admin"))


@app.route("/admin/user/<int:uid>/<action>", methods=["POST"])
@admin_required
def admin_user(uid, action):
    if uid == session.get("user_id"):
        return redirect(url_for("admin"))
    if action == "approve":
        db().execute("UPDATE users SET approved=1 WHERE id=?", (uid,))
    elif action == "revoke":
        db().execute("UPDATE users SET approved=0 WHERE id=?", (uid,))
    elif action == "toggle":
        db().execute("UPDATE users SET disabled = 1 - disabled WHERE id=?",
                     (uid,))
    elif action == "promote":
        db().execute("UPDATE users SET role='admin' WHERE id=?", (uid,))
    elif action == "reset":
        db().execute("UPDATE users SET password_hash=? WHERE id=?",
                     (generate_password_hash("changeme123"), uid))
    db().commit()
    return redirect(url_for("admin"))


if __name__ == "__main__":
    os.makedirs(DATA_DIR, exist_ok=True)
    init_db()
    if not os.path.exists(DEFAULT_DATA):
        print(f"WARNING: cutoff data not found at {DEFAULT_DATA}")
    print("Starting MHT-CET College Predictor at http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
