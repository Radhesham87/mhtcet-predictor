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
    "active_data_file": DEFAULT_XLSX,
    "data_year": "2025 (Latest)",
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
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS prediction_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        mode TEXT, value REAL, category TEXT,
        branches TEXT, districts TEXT,
        results INTEGER, created_at TEXT
    );
    """)
    cur = con.execute("SELECT COUNT(*) c FROM users WHERE role='admin'")
    if cur.fetchone()[0] == 0:
        con.execute(
            "INSERT INTO users (name,email,password_hash,role,created_at) "
            "VALUES (?,?,?,?,?)",
            ("Administrator", DEFAULT_ADMIN_EMAIL,
             generate_password_hash(DEFAULT_ADMIN_PASSWORD), "admin",
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
        path = DEFAULT_XLSX
    mtime = os.path.getmtime(path)
    if _DATA_CACHE["path"] == path and _DATA_CACHE["mtime"] == mtime:
        return _DATA_CACHE["df"]

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

    # Volatility: spread between the first and the lowest round percentile
    first = df[pct_cols[0]] if pct_cols else pd.NA
    df["Volatility Value"] = (first - df["Cutoff Percentile"]).abs()

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


def zone_for_gap(gap, safe_th, amb_th):
    if gap >= safe_th:
        return "Safe"
    if gap >= 0:
        return "Moderate"
    if gap >= amb_th:
        return "Ambitious"
    return "Reach"


def admission_probability(gap):
    """Logistic curve on the percentile gap, clamped to 2-98%."""
    p = 100 / (1 + math.exp(-1.8 * gap))
    return round(min(max(p, 2), 98), 1)


def volatility_label(v):
    if pd.isna(v):
        return "N/A"
    if v < 0.3:
        return "Low"
    if v < 1.0:
        return "Medium"
    return "High"


def run_prediction(form):
    df = load_data()
    settings_band = float(get_setting("pct_band"))
    priority_codes = [int(c) for c in get_setting("priority_codes")]
    safe_th = float(get_setting("zone_safe"))
    amb_th = float(get_setting("zone_ambitious"))

    mode = form.get("mode", "rank")
    if mode == "rank":
        rank_in = int(form.get("value") or 0)
        if rank_in <= 0:
            return {"error": "Enter a valid rank."}
        percentile = estimate_percentile_from_rank(df, rank_in)
        entered = f"Rank {rank_in:,}"
        counterpart = f"~{percentile:.2f} percentile"
    else:
        percentile = float(form.get("value") or -1)
        if not 0 <= percentile <= 100:
            return {"error": "Enter a valid percentile (0-100)."}
        est = estimate_merit_rank(df, percentile)
        entered = f"Percentile {percentile}"
        counterpart = f"~{est:,} merit rank" if est else "N/A"

    d = df

    # Quota toggle
    if form.get("quota_scope") == "all_india":
        return {"error": "All India Merit data is not present in the "
                         "current dataset. It contains Maharashtra State "
                         "CAP rounds only.", "results": []}

    # Category + additional flag codes (flags ADD extra seat pools)
    category = form.get("category", "OPEN")
    mask = d["Base Category"] == category
    if form.get("flag_defence"):
        mask |= d["Category"].str.startswith("DEF")
    if form.get("flag_pwd"):
        mask |= d["Category"].str.startswith("PWD")
    if form.get("opt_tfws"):
        mask |= d["Base Category"] == "TFWS"
    if form.get("opt_minority"):
        mask |= d["Base Category"] == "MI"
    d = d[mask]

    # Gender
    gender = form.get("gender", "Any")
    if gender == "Male":
        d = d[d["Gender"].isin(["Gender-Neutral", "Any"])]
    elif gender == "Female":
        d = d[d["Gender"].isin(["Female (Ladies)", "Gender-Neutral", "Any"])]

    # Seat type (Level column)
    seat = form.get("seat_type", "all")
    if seat == "state":
        d = d[d["Level"].str.contains("State", na=False)]
    elif seat == "home":
        d = d[d["Level"].str.startswith("Home University Seats Allotted to Home",
                                        na=False)]
    elif seat == "other":
        d = d[d["Level"].str.startswith("Other Than Home", na=False)]

    # Advanced filters
    colleges = form.get("colleges") or []
    branches = form.get("branches") or []
    districts = form.get("districts") or []
    if colleges:
        d = d[d["Institute Name"].isin(colleges)]
    if branches:
        d = d[d["Course Name"].isin(branches)]
    if districts:
        d = d[d["District"].isin(districts)]

    # Band + priority pinning
    lo, hi = percentile - settings_band, percentile + settings_band
    in_band = d[d["Cutoff Percentile"].between(lo, hi)]
    prio = d[d["Institute Code"].isin(priority_codes)]
    combined = pd.concat([prio, in_band]).drop_duplicates().copy()

    if combined.empty:
        return {"entered": entered, "counterpart": counterpart,
                "results": []}

    combined["gap"] = percentile - combined["Cutoff Percentile"]
    combined["zone"] = combined["gap"].map(
        lambda gp: zone_for_gap(gp, safe_th, amb_th))

    # Gap filter + safety-zone chip filter
    if form.get("gap_filter", "met") == "met":
        keep_prio = combined["Institute Code"].isin(priority_codes)
        combined = combined[(combined["gap"] >= 0) | keep_prio]
    zones = form.get("zones") or []
    if zones:
        keep_prio = combined["Institute Code"].isin(priority_codes)
        combined = combined[combined["zone"].isin(zones) | keep_prio]

    if combined.empty:
        return {"entered": entered, "counterpart": counterpart,
                "results": []}

    combined["probability"] = combined["gap"].map(admission_probability)
    # Smart Score: 60% admission probability + 40% institute demand
    pct_norm = (combined["Cutoff Percentile"] / 100).clip(0, 1)
    combined["smart_score"] = (0.6 * combined["probability"] +
                               0.4 * pct_norm * 100).round(1)
    combined["volatility"] = combined["Volatility Value"].map(volatility_label)

    rank_map = {c: i for i, c in enumerate(priority_codes)}
    combined["_prio"] = combined["Institute Code"].map(
        lambda c: rank_map.get(c, len(priority_codes)))
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
            "quota": str(r["Level"]),
            "seat_code": str(r["Category"]),
            "cutoff_pct": round(float(r["Cutoff Percentile"]), 4),
            "cutoff_rank": (int(r["Cutoff Rank"])
                            if pd.notna(r["Cutoff Rank"]) else None),
            "gap": round(float(r["gap"]), 2),
            "zone": r["zone"],
            "probability": float(r["probability"]),
            "smart_score": float(r["smart_score"]),
            "volatility": r["volatility"],
            "yoy": "N/A",   # needs multi-year data
        })

    zone_counts = combined["zone"].value_counts().to_dict()
    return {"entered": entered, "counterpart": counterpart,
            "percentile": round(percentile, 4),
            "zone_counts": zone_counts, "results": results}


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
                        "created_at) VALUES (?,?,?,?,?)",
                        (request.form["name"].strip(), email,
                         generate_password_hash(pw), "user",
                         datetime.now().isoformat()))
                    db().commit()
                    error = "Account created. Please sign in."
                except sqlite3.IntegrityError:
                    error = "Email already registered."
        else:
            row = db().execute("SELECT * FROM users WHERE email=?",
                               (email,)).fetchone()
            if row and not row["disabled"] and \
                    check_password_hash(row["password_hash"], pw):
                session["user_id"] = row["id"]
                session["name"] = row["name"]
                session["role"] = row["role"]
                return redirect(url_for("features"))
            error = "Invalid credentials or account disabled."
    return render_template("login.html", error=error)


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
        colleges=sorted(df["Institute Name"].dropna().unique()),
        branches=sorted(df["Course Name"].dropna().unique()),
        districts=sorted(df["District"].dropna().unique()),
        data_year=get_setting("data_year"))


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
            f"Entered: <b>{out['entered']}</b> ({out['counterpart']}) | "
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
    users = db().execute(
        "SELECT id,name,email,role,disabled,created_at FROM users "
        "ORDER BY id").fetchall()
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
                "data_year": get_setting("data_year")}
    return render_template("admin.html", users=users, data_info=data_info,
                           settings=settings, total_predictions=logs,
                           top_branches=top_branches)


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
    return redirect(url_for("admin"))


@app.route("/admin/user/<int:uid>/<action>", methods=["POST"])
@admin_required
def admin_user(uid, action):
    if uid == session.get("user_id"):
        return redirect(url_for("admin"))
    if action == "toggle":
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
    if not os.path.exists(DEFAULT_XLSX):
        print(f"WARNING: cutoff file not found at {DEFAULT_XLSX}")
    print("Starting MHT-CET College Predictor at http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
