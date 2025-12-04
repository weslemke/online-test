import os
import re
import csv
import sqlite3
from io import BytesIO
from datetime import datetime, timezone

from flask import (
    Flask, g, render_template, request, abort, redirect, session, send_file
)

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.utils import ImageReader


# -----------------------------
# Config
# -----------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "test.db")

ADMIN_PASSWORD = "Rotamotion1"
ADMIN_BASE = "/controlpanel"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-render-env")


# -----------------------------
# Helpers
# -----------------------------
def now_utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def slugify(text: str) -> str:
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-") or "test"


def is_admin() -> bool:
    return session.get("is_admin") is True


def db() -> sqlite3.Connection:
    if "db" not in g:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        g.db = conn
    return g.db


@app.teardown_appcontext
def _close_db(exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


def init_db():
    # tests table
    db().execute("""
        CREATE TABLE IF NOT EXISTS tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            slug TEXT NOT NULL UNIQUE,
            pass_score INTEGER NOT NULL DEFAULT 70,
            created_at TEXT NOT NULL
        )
    """)

    # questions table
    db().execute("""
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            test_id INTEGER NOT NULL,
            prompt TEXT NOT NULL,
            qtype TEXT NOT NULL DEFAULT 'MCQ',  -- 'MCQ' or 'TF'
            a TEXT NOT NULL,
            b TEXT NOT NULL,
            c TEXT,
            d TEXT,
            correct TEXT NOT NULL, -- 'A','B','C','D'
            FOREIGN KEY(test_id) REFERENCES tests(id)
        )
    """)

    # attempts table (results)
    db().execute("""
        CREATE TABLE IF NOT EXISTS attempts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            test_id INTEGER NOT NULL,
            student_name TEXT NOT NULL,
            score INTEGER NOT NULL,
            passed INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(test_id) REFERENCES tests(id)
        )
    """)

    db().commit()


def ensure_slug_column():
    # If old DB existed without slug column, add it
    cols = [r["name"] for r in db().execute("PRAGMA table_info(tests)").fetchall()]
    if "slug" not in cols:
        db().execute("ALTER TABLE tests ADD COLUMN slug TEXT")
        db().commit()

    # Ensure every test has a slug
    tests = db().execute("SELECT id, title, slug FROM tests").fetchall()
    for t in tests:
        if not t["slug"]:
            s = slugify(t["title"])
            # make unique
            base = s
            i = 2
            while db().execute("SELECT 1 FROM tests WHERE slug=?", (s,)).fetchone():
                s = f"{base}-{i}"
                i += 1
            db().execute("UPDATE tests SET slug=? WHERE id=?", (s, t["id"]))
    db().commit()


def seed_line_breaking_exam():
    # Seed only if not exists
    title = "Line Breaking Final Exam"
    slug = "line-breaking-final-exam"

    row = db().execute("SELECT id FROM tests WHERE slug=?", (slug,)).fetchone()
    if row:
        return

    db().execute(
        "INSERT INTO tests (title, slug, pass_score, created_at) VALUES (?,?,?,?)",
        (title, slug, 100, now_utc_iso())
    )
    test_id = db().execute("SELECT id FROM tests WHERE slug=?", (slug,)).fetchone()["id"]

    questions = [
        # prompt, options A-D, correct, qtype
        ("Intentionally opening a pipe, line or duct for the purpose of cleaning, inspection, maintenance or replacing components within a system is referred to as:",
         "Line Breaking", "Pipe Sealing", "System Flushing", "Duct Isolation", "A", "MCQ"),
        ("Routine line breaking instructions can be found in the:",
         "Emergency action plan (EAP)", "PPE assessment (PPEA)", "Safety Data Sheets (SDS)", "Safe operation procedure (SOP)", "D", "MCQ"),
        ("Workers performing non-routine line-breaking tasks must obtain a(n) [ blank ] prior to starting work to identify the risks and controls that are needed.",
         "Safe Work Permit", "Safe operating Procedure (SOP)", "Emergency action plan (EAP)", "Open-end wrench", "A", "MCQ"),
        ("What type of flange slips over the pipe without needing to be welded to it and can swivel around the pipe to help line up opposing bolt holes?",
         "Orifice plate", "Lap joint flange", "Spectacle blind", "Blind flange", "B", "MCQ"),
        ("The minimum PPE for line-breaking jobs include a hard hat/helmet, gloves, face shield, goggles/safety glasses and:",
         "Hair net", "Chemical-protective clothing", "Disposable paper gown", "Shoe cover", "B", "MCQ"),
        ("To work in a defensive position opening a flange, begin by loosening the bolts [ blank. ]",
         "Farthest away from you.", "That are close to you first, working clockwise around the line.",
         "While crouched as low to the ground as possible.", "With open-end wrenches pointed toward your body.", "A", "MCQ"),
        ("When performing line breaking, it's best to use a cheater bar to gain leverage.",
         "True", "False", None, None, "B", "TF"),
        ("Before beginning line breaking, which of the following steps should be taken?",
         "Ensure lines and equipment are as free from recognized hazards as possible and test to confirm.",
         "Increase the pressure within the lines to check for leaks.",
         "Remove all tags from the equipment to allow for easy access.",
         "Turn on all cathodic protection rectifiers affecting the piping.", "A", "MCQ"),
        ("If the proper steps of bleeding off pressure are not followed, a flammable chemical release can create a potential ignition source, leading to:",
         "Biological hazards", "Fire and explosion.", "Slips, trips, and falls.", "Ultraviolet light exposure.", "B", "MCQ"),
        ("What should you do if you are unsure about which type of flange or procedure to use for line breaking?",
         "Check schematics and diagrams.", "Consult the human resources department.",
         "Perform the line break as best you can.", "Use a \"shop-made\" flange.", "A", "MCQ"),
    ]

    for (prompt, a, b, copt, dopt, correct, qtype) in questions:
        # If TF, store only a/b; c/d should be empty strings (NOT NULL handling)
        c_val = copt if copt is not None else ""
        d_val = dopt if dopt is not None else ""
        db().execute("""
            INSERT INTO questions (test_id, prompt, qtype, a, b, c, d, correct)
            VALUES (?,?,?,?,?,?,?,?)
        """, (test_id, prompt, qtype, a, b, c_val, d_val, correct))

    db().commit()


@app.before_request
def _ensure_db():
    init_db()
    ensure_slug_column()
    seed_line_breaking_exam()


# -----------------------------
# Certificate PDF
# -----------------------------
def make_certificate_pdf(student_name: str, test_title: str, date_str: str) -> BytesIO:
    buf = BytesIO()

    pagesize = landscape(letter)
    c = canvas.Canvas(buf, pagesize=pagesize)
    width, height = pagesize
    c.setTitle("Certificate of Completion")

    # --- WATERMARK LOGO (full-page cover) ---
    logo_path = os.path.join(app.root_path, "static", "logo.jpg")
    if os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            c.saveState()
            try:
                c.setFillAlpha(0.10)
            except Exception:
                pass

            margin = 0
            wm_w = width - (margin * 2)
            wm_h = height - (margin * 2)

            img_w, img_h = img.getSize()
            scale = max(wm_w / img_w, wm_h / img_h)
            draw_w = img_w * scale
            draw_h = img_h * scale

            x = margin + (wm_w - draw_w) / 2
            y = margin + (wm_h - draw_h) / 2

            c.drawImage(img, x, y, width=draw_w, height=draw_h, mask="auto")
            c.restoreState()
        except Exception:
            pass

    # --- TEXT ---
    c.setFont("Helvetica-Bold", 34)
    c.drawCentredString(width / 2, height - 110, "Certificate of Completion")

    c.setFont("Helvetica", 16)
    c.drawCentredString(width / 2, height - 165, "This certifies that")

    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(width / 2, height - 215, student_name)

    c.setFont("Helvetica", 16)
    c.drawCentredString(width / 2, height - 265, "has successfully completed")

    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 305, test_title)

    c.setFont("Helvetica", 14)
    c.drawCentredString(width / 2, 85, f"Date: {date_str}")

    # --- SIGNATURE (image) + name + line ---
    sig_path = os.path.join(app.root_path, "static", "signature.png")
    sig_x = 80
    sig_y = 70
    sig_w = 200
    sig_h = 45

    if os.path.exists(sig_path):
        try:
            sig_img = ImageReader(sig_path)
            c.drawImage(sig_img, sig_x, sig_y, width=sig_w, height=sig_h, mask="auto")
        except Exception:
            c.setLineWidth(1)
            c.line(sig_x, sig_y + 15, sig_x + 200, sig_y + 15)
    else:
        c.setLineWidth(1)
        c.line(sig_x, sig_y + 15, sig_x + 200, sig_y + 15)

    name_y = sig_y - 8
    c.setFont("Helvetica-Bold", 12)
    c.drawString(sig_x, name_y, "Chad Riley")

    line_y = name_y - 6
    c.setLineWidth(1)
    c.line(sig_x, line_y, sig_x + 180, line_y)

    c.setFont("Helvetica", 10)
    c.drawString(sig_x, line_y - 14, "Authorized Signature")

    c.showPage()
    c.save()

    buf.seek(0)
    return buf


# -----------------------------
# Student routes
# -----------------------------
@app.get("/")
def home():
    tests = db().execute("SELECT id, title, slug, pass_score FROM tests ORDER BY id DESC").fetchall()
    return render_template("home.html", tests=tests)


@app.get("/tests/<slug>/take")
def take_test(slug):
    t = db().execute("SELECT * FROM tests WHERE slug=?", (slug,)).fetchone()
    if not t:
        abort(404)

    qs = db().execute("SELECT * FROM questions WHERE test_id=? ORDER BY id ASC", (t["id"],)).fetchall()
    saved_name = session.get("saved_name", "")
    return render_template("take_test.html", test=t, questions=qs, saved_name=saved_name)


@app.post("/tests/<slug>/submit")
def submit_test(slug):
    t = db().execute("SELECT * FROM tests WHERE slug=?", (slug,)).fetchone()
    if not t:
        abort(404)

    name = (request.form.get("student_name") or "").strip()
    if not name:
        abort(400)
    session["saved_name"] = name  # make name stick

    qs = db().execute("SELECT * FROM questions WHERE test_id=? ORDER BY id ASC", (t["id"],)).fetchall()

    correct_count = 0
    review = []

    for q in qs:
        chosen = (request.form.get(f"q_{q['id']}") or "").strip()
        correct = q["correct"]

        chosen_text = ""
        correct_text = ""

        def opt_text(letter):
            if letter == "A":
                return q["a"]
            if letter == "B":
                return q["b"]
            if letter == "C":
                return q["c"]
            if letter == "D":
                return q["d"]
            return ""

        chosen_text = opt_text(chosen)
        correct_text = opt_text(correct)

        is_correct = (chosen == correct)
        if is_correct:
            correct_count += 1

        review.append({
            "prompt": q["prompt"],
            "chosen": chosen or "(no answer)",
            "chosen_text": chosen_text or "",
            "correct": correct,
            "correct_text": correct_text or "",
            "is_correct": is_correct
        })

    total = len(qs) if qs else 1
    score = int(round((correct_count / total) * 100))
    passed = 1 if score >= int(t["pass_score"]) else 0

    cur = db().execute("""
        INSERT INTO attempts (test_id, student_name, score, passed, created_at)
        VALUES (?,?,?,?,?)
    """, (t["id"], name, score, passed, now_utc_iso()))
    db().commit()

    attempt_id = cur.lastrowid

    return render_template(
        "result.html",
        test=t,
        student_name=name,
        score=score,
        passed=bool(passed),
        review=review,
        attempt_id=attempt_id
    )


@app.get("/tests/<slug>/certificate/<int:attempt_id>")
def certificate(slug, attempt_id: int):
    t = db().execute("SELECT * FROM tests WHERE slug=?", (slug,)).fetchone()
    if not t:
        abort(404)

    a = db().execute("""
        SELECT * FROM attempts
        WHERE id=? AND test_id=?
    """, (attempt_id, t["id"])).fetchone()

    if not a:
        abort(404)

    if not a["passed"]:
        abort(403)

    # Generate PDF
    date_str = datetime.now().strftime("%m/%d/%Y")
    pdf = make_certificate_pdf(a["student_name"], t["title"], date_str)

    safe_name = re.sub(r"[^a-zA-Z0-9_-]+", "_", a["student_name"]).strip("_") or "student"
    filename = f"Certificate_{safe_name}_{t['slug']}.pdf"

    return send_file(pdf, mimetype="application/pdf", as_attachment=True, download_name=filename)


# -----------------------------
# Admin routes (hidden)
# -----------------------------
@app.get(ADMIN_BASE)
def controlpanel_login():
    return """
    <html><body style="font-family:system-ui;max-width:420px;margin:40px auto;padding:0 16px;">
      <h1>Admin Login</h1>
      <form method="post" action="/controlpanel/login">
        <input type="password" name="password" placeholder="Admin password"
               style="width:100%;padding:10px;margin:10px 0;" required />
        <button type="submit" style="width:100%;padding:10px;">Login</button>
      </form>
    </body></html>
    """


@app.post("/controlpanel/login")
def controlpanel_login_post():
    pw = (request.form.get("password") or "").strip()
    if pw == ADMIN_PASSWORD:
        session["is_admin"] = True
        return redirect(f"{ADMIN_BASE}/results")
    abort(403)


@app.get(f"{ADMIN_BASE}/logout")
def controlpanel_logout():
    session.pop("is_admin", None)
    return redirect(ADMIN_BASE)


@app.get(f"{ADMIN_BASE}/results")
def controlpanel_results():
    if not is_admin():
        return redirect(ADMIN_BASE)

    rows = db().execute("""
        SELECT a.id, a.created_at, a.student_name, a.score, a.passed,
               t.title AS test_title, t.slug AS test_slug
        FROM attempts a
        JOIN tests t ON t.id = a.test_id
        ORDER BY a.id DESC
        LIMIT 500
    """).fetchall()

    table_rows = ""
    for r in rows:
        status = "PASS" if r["passed"] else "FAIL"
        cert_link = ""
        if r["passed"]:
            cert_link = f'<a href="/tests/{r["test_slug"]}/certificate/{r["id"]}">PDF</a>'

        table_rows += f"""
          <tr>
            <td>{r["created_at"]}</td>
            <td>{r["test_title"]}</td>
            <td>{r["student_name"]}</td>
            <td>{r["score"]}%</td>
            <td>{status}</td>
            <td>{cert_link}</td>
          </tr>
        """

    return f"""
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Admin Results</title>
        <style>
          body {{ font-family: system-ui, Arial; max-width: 1100px; margin: 32px auto; padding: 0 16px; }}
          a.btn {{ display:inline-block; padding:10px 14px; border:1px solid #444; border-radius:10px; text-decoration:none; margin-right: 8px; }}
          table {{ width: 100%; border-collapse: collapse; margin-top: 16px; }}
          th, td {{ border-bottom: 1px solid #eee; padding: 10px; text-align: left; font-size: 14px; }}
          th {{ background: #fafafa; }}
          .muted {{ opacity: .75; }}
        </style>
      </head>
      <body>
        <h1>Admin Results</h1>
        <p class="muted">Private admin area at <b>/controlpanel</b>.</p>

        <p>
          <a class="btn" href="{ADMIN_BASE}/export.csv">Export CSV</a>
          <a class="btn" href="{ADMIN_BASE}/logout">Logout</a>
          <a class="btn" href="/">Student Home</a>
        </p>

        <table>
          <thead>
            <tr>
              <th>Time</th>
              <th>Test</th>
              <th>Student</th>
              <th>Score</th>
              <th>Status</th>
              <th>Certificate</th>
            </tr>
          </thead>
          <tbody>
            {table_rows if table_rows else '<tr><td colspan="6" class="muted">No results yet.</td></tr>'}
          </tbody>
        </table>
      </body>
    </html>
    """
@app.get(f"{ADMIN_BASE}/export.csv")
def controlpanel_export_csv():
    if not is_admin():
        return redirect(ADMIN_BASE)

    rows = db().execute("""
        SELECT a.created_at, t.title AS test_title, t.slug AS test_slug,
               a.student_name, a.score, a.passed, a.id AS attempt_id
        FROM attempts a
        JOIN tests t ON t.id = a.test_id
        ORDER BY a.id DESC
        LIMIT 5000
    """).fetchall()

    # Build the CSV fully as bytes (no streaming / no wrapper closing issues)
    lines = []
    lines.append("created_at,test_title,student_name,score,status,attempt_id,certificate_url")

    def esc(s):
        s = "" if s is None else str(s)
        s = s.replace('"', '""')
        return f'"{s}"'

    for r in rows:
        status = "PASS" if r["passed"] else "FAIL"
        cert_url = ""
        if r["passed"]:
            cert_url = f"/tests/{r['test_slug']}/certificate/{r['attempt_id']}"

        lines.append(",".join([
            esc(r["created_at"]),
            esc(r["test_title"]),
            esc(r["student_name"]),
            esc(r["score"]),
            esc(status),
            esc(r["attempt_id"]),
            esc(cert_url),
        ]))

    csv_bytes = ("\n".join(lines) + "\n").encode("utf-8")
    out = BytesIO(csv_bytes)

    filename = f"results_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return send_file(out, mimetype="text/csv", as_attachment=True, download_name=filename)





if __name__ == "__main__":
    app.run(debug=True)

