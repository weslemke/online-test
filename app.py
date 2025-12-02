import os
import sqlite3
import random
from datetime import datetime
from io import BytesIO

from flask import Flask, g, render_template, request, abort, redirect, session, send_file
from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.secret_key = "change-this-to-a-long-random-secret"  # change if you deploy online

DB = "test.db"

# Admin password (requested)
ADMIN_PASSWORD = "Rotamotion1"

# Certificate storage (admin controlled)
CERT_DIR = "certificates"
CERT_XLSX = "certificates_log.xlsx"


# -------------------------
# Database helpers
# -------------------------
def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    conn = g.pop("db", None)
    if conn:
        conn.close()


def init_db():
    d = db()
    d.executescript("""
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS tests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      slug TEXT,
      pass_score INTEGER NOT NULL DEFAULT 70,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS questions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      test_id INTEGER NOT NULL,
      qtype TEXT NOT NULL DEFAULT 'MCQ' CHECK(qtype IN ('MCQ','TF')),
      prompt TEXT NOT NULL,
      a TEXT NOT NULL,
      b TEXT NOT NULL,
      c TEXT,
      d TEXT,
      correct CHAR(1) NOT NULL CHECK(correct IN ('A','B','C','D')),
      FOREIGN KEY(test_id) REFERENCES tests(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS attempts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      test_id INTEGER NOT NULL,
      student_name TEXT NOT NULL,
      score INTEGER NOT NULL,
      passed INTEGER NOT NULL CHECK(passed IN (0,1)),
      created_at TEXT NOT NULL,
      FOREIGN KEY(test_id) REFERENCES tests(id) ON DELETE CASCADE
    );
    """)
    d.commit()

    # --- MIGRATIONS for older DBs ---
    test_cols = [r["name"] for r in d.execute("PRAGMA table_info(tests)").fetchall()]
    if "slug" not in test_cols:
        d.execute("ALTER TABLE tests ADD COLUMN slug TEXT")
        d.commit()

    q_cols = [r["name"] for r in d.execute("PRAGMA table_info(questions)").fetchall()]
    if "qtype" not in q_cols:
        d.execute("ALTER TABLE questions ADD COLUMN qtype TEXT NOT NULL DEFAULT 'MCQ'")
        d.commit()


def slugify(text: str) -> str:
    text = (text or "").strip().lower()
    out = []
    prev_dash = False
    for ch in text:
        if ch.isalnum():
            out.append(ch)
            prev_dash = False
        else:
            if not prev_dash:
                out.append("-")
                prev_dash = True
    slug = "".join(out).strip("-")
    return slug or "test"


def unique_slug(base: str) -> str:
    s = base
    i = 2
    while db().execute("SELECT 1 FROM tests WHERE slug=?", (s,)).fetchone():
        s = f"{base}-{i}"
        i += 1
    return s


def ensure_slugs():
    tests = db().execute("SELECT id, title, slug FROM tests").fetchall()
    for t in tests:
        if not t["slug"]:
            slug = unique_slug(slugify(t["title"]))
            db().execute("UPDATE tests SET slug=? WHERE id=?", (slug, t["id"]))
    db().commit()


def seed_line_breaking_exam():
    title = "Line Breaking Final Exam"
    pass_score = 100

    # If it already exists, do nothing
    existing = db().execute("SELECT id FROM tests WHERE title=?", (title,)).fetchone()
    if existing:
        return

    slug = unique_slug(slugify(title))
    cur = db().execute(
        "INSERT INTO tests(title, slug, pass_score, created_at) VALUES (?,?,?,?)",
        (title, slug, pass_score, datetime.utcnow().isoformat()),
    )
    test_id = cur.lastrowid

    questions = [
        # (qtype, prompt, a, b, c, d, correct)
        ("MCQ",
         "Intentionally opening a pipe, line or duct for the purpose of cleaning, inspection, maintenance or replacing components within a system is referred to as:",
         "Line Breaking", "Pipe Sealing", "System Flushing", "Duct Isolation", "A"),

        ("MCQ",
         "Routine line breaking instructions can be found in the:",
         "Emergency action plan (EAP)", "PPE assessment (PPEA)", "Safety Data Sheets (SDS)", "Safe operation procedure (SOP)",
         "D"),

        ("MCQ",
         "Workers performing non-routine line-breaking tasks must obtain a(n) [ blank ] prior to starting work to identify the risks and controls that are needed.",
         "Safe Work Permit", "Safe operating Procedure (SOP)", "Emergency action plan (EAP)", "Open-end wrench", "A"),

        ("MCQ",
         "What type of flange slips over the pipe without needing to be welded to it and can swivel around the pipe to help line up opposing bolt holes?",
         "Orifice plate", "Lap joint flange", "Spectacle blind", "Blind flange", "B"),

        ("MCQ",
         "The minimum PPE for line-breaking jobs include a hard hat/helmet, gloves, face shield, goggles/safety glasses and:",
         "Hair net", "Chemical-protective clothing", "Disposable paper gown", "Shoe cover", "B"),

        ("MCQ",
         "To work in a defensive position opening a flange, begin by loosening the bolts [ blank. ]",
         "Farthest away from you.",
         "That are close to you first, working clockwise around the line.",
         "While crouched as low to the ground as possible.",
         "With open-end wrenches pointed toward your body.",
         "A"),

        ("TF",
         "When performing line breaking, it's best to use a cheater bar to gain leverage.",
         "True", "False", None, None, "B"),

        ("MCQ",
         "Before beginning line breaking, which of the following steps should be taken?",
         "Ensure lines and equipment are as free from recognized hazards as possible and test to confirm.",
         "Increase the pressure within the lines to check for leaks.",
         "Remove all tags from the equipment to allow for easy access.",
         "Turn on all cathodic protection rectifiers affecting the piping.",
         "A"),

        ("MCQ",
         "If the proper steps of bleeding off pressure are not followed, a flammable chemical release can create a potential ignition source, leading to:",
         "Biological hazards", "Fire and explosion.", "Slips, trips, and falls.", "Ultraviolet light exposure.",
         "B"),

        ("MCQ",
         "What should you do if you are unsure about which type of flange or procedure to use for line breaking?",
         "Check schematics and diagrams.", "Consult the human resources department.",
         "Perform the line break as best you can.", 'Use a "shop-made" flange.',
         "A"),
    ]

    for qtype, prompt, a, b, copt, dopt, correct in questions:
        db().execute("""
            INSERT INTO questions(test_id, qtype, prompt, a, b, c, d, correct)
            VALUES (?,?,?,?,?,?,?,?)
        """, (test_id, qtype, prompt, a, b, copt, dopt, correct))

    db().commit()


@app.before_request
def _ensure_db():
    init_db()
    ensure_slugs()
    seed_line_breaking_exam()


# -------------------------
# Admin auth
# -------------------------
def is_admin() -> bool:
    return session.get("is_admin") is True


# -------------------------
# Certificate generation + Excel log (admin controlled)
# -------------------------
def make_certificate_pdf(student_name: str, test_title: str, date_str: str) -> BytesIO:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    c.setTitle("Certificate of Completion")

    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(width / 2, height - 140, "Certificate of Completion")

    c.setFont("Helvetica", 14)
    c.drawCentredString(width / 2, height - 190, "This certifies that")

    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 235, student_name)

    c.setFont("Helvetica", 14)
    c.drawCentredString(width / 2, height - 280, "has successfully completed")

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2, height - 315, test_title)

    c.setFont("Helvetica", 12)
    c.drawCentredString(width / 2, height - 365, f"Date: {date_str}")

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width / 2, 80, "Generated by your training portal")

    c.showPage()
    c.save()

    buf.seek(0)
    return buf


def append_certificate_to_excel(test_title: str, student_name: str, score: int, pdf_path: str):
    os.makedirs(CERT_DIR, exist_ok=True)

    if os.path.exists(CERT_XLSX):
        wb = load_workbook(CERT_XLSX)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Certificates"
        ws.append(["timestamp", "test_title", "student_name", "score", "pdf_file"])

    ws.append([
        datetime.now().isoformat(timespec="seconds"),
        test_title,
        student_name,
        score,
        pdf_path
    ])
    wb.save(CERT_XLSX)


def save_certificate_pdf(student_name: str, test_title: str) -> str:
    os.makedirs(CERT_DIR, exist_ok=True)

    date_str = datetime.now().strftime("%B %d, %Y")
    pdf_buf = make_certificate_pdf(student_name, test_title, date_str)

    safe_name = "".join(ch for ch in student_name if ch.isalnum() or ch in (" ", "-", "_")).strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Certificate_{safe_name}_{slugify(test_title)}_{timestamp}.pdf"
    pdf_path = os.path.join(CERT_DIR, filename)

    with open(pdf_path, "wb") as out:
        out.write(pdf_buf.getbuffer())

    return pdf_path


# -------------------------
# Admin: login/logout
# -------------------------
@app.route("/admin/login", methods=["GET"])
def admin_login():
    return """
    <html><body style="font-family:system-ui;max-width:420px;margin:40px auto;padding:0 16px;">
      <h1>Admin Login</h1>
      <form method="post" action="/admin/login">
        <input type="password" name="password" placeholder="Admin password"
               style="width:100%;padding:10px;margin:10px 0;" required />
        <button type="submit" style="width:100%;padding:10px;">Login</button>
      </form>
    </body></html>
    """


@app.route("/admin/login", methods=["POST"])
def admin_login_post():
    pw = (request.form.get("password") or "").strip()
    if pw == ADMIN_PASSWORD:
        session["is_admin"] = True
        return redirect("/admin/certificates")
    abort(403)


@app.route("/admin/logout", methods=["GET"])
def admin_logout():
    session.pop("is_admin", None)
    return redirect("/")


# -------------------------
# Admin: certificates dashboard (admin only) - NO Test ID shown
# -------------------------
@app.route("/admin/certificates", methods=["GET"])
def admin_certificates():
    if not is_admin():
        return redirect("/admin/login")

    rows = []
    if os.path.exists(CERT_XLSX):
        wb = load_workbook(CERT_XLSX)
        ws = wb.active
        data = list(ws.iter_rows(values_only=True))
        if len(data) >= 2:
            headers = data[0]
            for r in data[1:]:
                rows.append(dict(zip(headers, r)))

    html_rows = ""
    for r in reversed(rows):  # newest first
        pdf_file = r.get("pdf_file") or ""
        pdf_name = os.path.basename(str(pdf_file)) if pdf_file else ""
        pdf_link = f'<a href="/admin/certificates/pdf/{pdf_name}">{pdf_name}</a>' if pdf_name else ""
        html_rows += f"""
          <tr>
            <td>{r.get("timestamp","")}</td>
            <td>{r.get("test_title","")}</td>
            <td>{r.get("student_name","")}</td>
            <td>{r.get("score","")}</td>
            <td>{pdf_link}</td>
          </tr>
        """

    return f"""
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Admin Certificates</title>
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
        <h1>Certificates (Admin)</h1>
        <p class="muted">Only admins can view/download certificates and the Excel log.</p>

        <p>
          <a class="btn" href="/admin/certificates/excel">Download Excel Log</a>
          <a class="btn" href="/admin/logout">Logout</a>
          <a class="btn" href="/">Home</a>
        </p>

        <table>
          <thead>
            <tr>
              <th>Timestamp</th>
              <th>Test Title</th>
              <th>Student Name</th>
              <th>Score</th>
              <th>PDF</th>
            </tr>
          </thead>
          <tbody>
            {html_rows if html_rows else '<tr><td colspan="5" class="muted">No certificates logged yet.</td></tr>'}
          </tbody>
        </table>
      </body>
    </html>
    """


@app.route("/admin/certificates/excel", methods=["GET"])
def admin_download_excel():
    if not is_admin():
        return redirect("/admin/login")
    if not os.path.exists(CERT_XLSX):
        abort(404, "No Excel log yet.")
    return send_file(CERT_XLSX, as_attachment=True, download_name="certificates_log.xlsx")


@app.route("/admin/certificates/pdf/<path:filename>", methods=["GET"])
def admin_download_pdf(filename):
    if not is_admin():
        return redirect("/admin/login")

    safe = os.path.basename(filename)
    pdf_path = os.path.join(CERT_DIR, safe)

    if not os.path.exists(pdf_path):
        abort(404)

    return send_file(pdf_path, as_attachment=True, download_name=safe)


# -------------------------
# Home
# -------------------------
@app.route("/", methods=["GET"])
def home():
    tests = db().execute("SELECT * FROM tests ORDER BY id DESC").fetchall()
    return render_template("home.html", tests=tests)


# -------------------------
# Admin: create test + add questions (admin only)
# -------------------------
@app.route("/admin/create", methods=["GET"])
def admin_create():
    if not is_admin():
        return redirect("/admin/login")
    return render_template("admin_create.html")


@app.route("/admin/create", methods=["POST"])
def admin_create_post():
    if not is_admin():
        return redirect("/admin/login")

    title = (request.form.get("title") or "").strip()
    pass_score = int(request.form.get("pass_score", "70") or 70)

    if not title:
        abort(400, "Title required")
    if pass_score < 0 or pass_score > 100:
        abort(400, "Pass score must be 0-100")

    slug = unique_slug(slugify(title))

    cur = db().execute(
        "INSERT INTO tests(title, slug, pass_score, created_at) VALUES (?,?,?,?)",
        (title, slug, pass_score, datetime.utcnow().isoformat()),
    )
    db().commit()
    return redirect(f"/admin/tests/{cur.lastrowid}/questions")


@app.route("/admin/tests/<int:test_id>/questions", methods=["GET"])
def admin_questions(test_id):
    if not is_admin():
        return redirect("/admin/login")

    t = db().execute("SELECT * FROM tests WHERE id=?", (test_id,)).fetchone()
    if not t:
        abort(404)
    qs = db().execute("SELECT * FROM questions WHERE test_id=? ORDER BY id DESC", (test_id,)).fetchall()
    return render_template("admin_create.html", test=t, questions=qs)


@app.route("/admin/tests/<int:test_id>/questions", methods=["POST"])
def admin_add_question(test_id):
    if not is_admin():
        return redirect("/admin/login")

    qtype = ((request.form.get("qtype") or "MCQ").strip().upper())
    if qtype not in ("MCQ", "TF"):
        abort(400, "Invalid question type")

    prompt = (request.form.get("prompt") or "").strip()
    a = (request.form.get("a") or "").strip()
    b = (request.form.get("b") or "").strip()
    copt = (request.form.get("c") or "").strip()
    dopt = (request.form.get("d") or "").strip()
    correct = ((request.form.get("correct") or "A").strip().upper())

    if not (prompt and a and b):
        abort(400, "Prompt, A, and B are required")

    if qtype == "MCQ":
        if not (copt and dopt):
            abort(400, "C and D required for multiple choice")
        if correct not in ("A", "B", "C", "D"):
            abort(400, "Correct must be A/B/C/D")
    else:
        # True/False is ONLY two options
        a = "True"
        b = "False"
        copt = None
        dopt = None
        if correct not in ("A", "B"):
            abort(400, "Correct must be A (True) or B (False) for TF")

    db().execute("""
        INSERT INTO questions(test_id, qtype, prompt, a, b, c, d, correct)
        VALUES (?,?,?,?,?,?,?,?)
    """, (test_id, qtype, prompt, a, b, copt, dopt, correct))
    db().commit()
    return redirect(f"/admin/tests/{test_id}/questions")


# -------------------------
# Admin: edit test + edit/delete questions + delete test (admin only)
# -------------------------
@app.route("/admin/tests/<int:test_id>/edit", methods=["GET"])
def admin_edit(test_id):
    if not is_admin():
        return redirect("/admin/login")

    t = db().execute("SELECT * FROM tests WHERE id=?", (test_id,)).fetchone()
    if not t:
        abort(404)

    qs = db().execute("SELECT * FROM questions WHERE test_id=? ORDER BY id ASC", (test_id,)).fetchall()
    return render_template("admin_edit.html", test=t, questions=qs)


@app.route("/admin/tests/<int:test_id>/edit", methods=["POST"])
def admin_edit_post(test_id):
    if not is_admin():
        return redirect("/admin/login")

    title = (request.form.get("title") or "").strip()
    pass_score = int(request.form.get("pass_score", "70") or 70)

    if not title:
        abort(400, "Title required")
    if pass_score < 0 or pass_score > 100:
        abort(400, "Pass score must be 0-100")

    old = db().execute("SELECT slug, title FROM tests WHERE id=?", (test_id,)).fetchone()
    new_slug = old["slug"]
    if old and title != old["title"]:
        new_slug = unique_slug(slugify(title))

    db().execute("UPDATE tests SET title=?, slug=?, pass_score=? WHERE id=?", (title, new_slug, pass_score, test_id))
    db().commit()
    return redirect(f"/admin/tests/{test_id}/edit")


@app.route("/admin/questions/<int:question_id>/update", methods=["POST"])
def admin_question_update(question_id):
    if not is_admin():
        return redirect("/admin/login")

    qtype = ((request.form.get("qtype") or "MCQ").strip().upper())
    if qtype not in ("MCQ", "TF"):
        abort(400, "Invalid question type")

    prompt = (request.form.get("prompt") or "").strip()
    a = (request.form.get("a") or "").strip()
    b = (request.form.get("b") or "").strip()
    copt = (request.form.get("c") or "").strip()
    dopt = (request.form.get("d") or "").strip()
    correct = ((request.form.get("correct") or "A").strip().upper())

    if not (prompt and a and b):
        abort(400, "Prompt, A, and B are required")

    if qtype == "MCQ":
        if not (copt and dopt):
            abort(400, "C and D required for multiple choice")
        if correct not in ("A", "B", "C", "D"):
            abort(400, "Correct must be A/B/C/D")
    else:
        a = "True"
        b = "False"
        copt = None
        dopt = None
        if correct not in ("A", "B"):
            abort(400, "Correct must be A (True) or B (False) for TF")

    row = db().execute("SELECT test_id FROM questions WHERE id=?", (question_id,)).fetchone()
    if not row:
        abort(404)
    test_id = row["test_id"]

    db().execute("""
        UPDATE questions
        SET qtype=?, prompt=?, a=?, b=?, c=?, d=?, correct=?
        WHERE id=?
    """, (qtype, prompt, a, b, copt, dopt, correct, question_id))
    db().commit()
    return redirect(f"/admin/tests/{test_id}/edit")


@app.route("/admin/questions/<int:question_id>/delete", methods=["POST"])
def admin_question_delete(question_id):
    if not is_admin():
        return redirect("/admin/login")

    row = db().execute("SELECT test_id FROM questions WHERE id=?", (question_id,)).fetchone()
    if not row:
        abort(404)
    test_id = row["test_id"]

    db().execute("DELETE FROM questions WHERE id=?", (question_id,))
    db().commit()
    return redirect(f"/admin/tests/{test_id}/edit")


@app.route("/admin/tests/<int:test_id>/delete", methods=["POST"])
def admin_test_delete(test_id):
    if not is_admin():
        return redirect("/admin/login")

    db().execute("DELETE FROM tests WHERE id=?", (test_id,))
    db().commit()
    return redirect("/")


# -------------------------
# Student: take test (by slug) + submit (by slug)
# -------------------------
@app.route("/tests/<slug>/take", methods=["GET"])
def take_test_by_slug(slug):
    t = db().execute("SELECT * FROM tests WHERE slug=?", (slug,)).fetchone()
    if not t:
        abort(404)

    qs = list(db().execute("SELECT * FROM questions WHERE test_id=?", (t["id"],)).fetchall())
    if not qs:
        abort(400, "This test has no questions yet.")

    random.shuffle(qs)
    saved_name = session.get("student_name", "")

    return render_template("take_test.html", test=t, questions=qs, saved_name=saved_name)


@app.route("/tests/<slug>/submit", methods=["POST"])
def submit_test_by_slug(slug):
    t = db().execute("SELECT * FROM tests WHERE slug=?", (slug,)).fetchone()
    if not t:
        abort(404)

    qs = list(db().execute("SELECT * FROM questions WHERE test_id=?", (t["id"],)).fetchall())
    if not qs:
        abort(400, "No questions yet")

    name = (request.form.get("student_name") or "").strip()
    if not name:
        abort(400, "Name required")

    session["student_name"] = name

    def opt_text(q, letter):
        mapping = {"A": q["a"], "B": q["b"], "C": q["c"], "D": q["d"]}
        val = mapping.get(letter)
        return val if val else ""

    review = []
    correct_count = 0

    for q in qs:
        chosen = (request.form.get(f"q_{q['id']}", "") or "").upper()
        is_correct = (chosen == q["correct"])
        if is_correct:
            correct_count += 1

        review.append({
            "prompt": q["prompt"],
            "qtype": q["qtype"],
            "chosen": chosen if chosen else "(No answer)",
            "chosen_text": opt_text(q, chosen) if chosen in ("A", "B", "C", "D") else "(No answer)",
            "correct": q["correct"],
            "correct_text": opt_text(q, q["correct"]),
            "is_correct": is_correct,
        })

    score = round((correct_count / len(qs)) * 100)
    passed = 1 if score >= t["pass_score"] else 0

    db().execute("""
      INSERT INTO attempts(test_id, student_name, score, passed, created_at)
      VALUES (?,?,?,?,?)
    """, (t["id"], name, score, passed, datetime.utcnow().isoformat()))
    db().commit()

    # If passed: create certificate (server only) + log to Excel for admin
    if passed == 1:
        pdf_path = save_certificate_pdf(name, t["title"])
        append_certificate_to_excel(t["title"], name, score, pdf_path)

    return render_template(
        "result.html",
        test=t,
        student_name=name,
        score=score,
        passed=bool(passed),
        review=review
    )


# -------------------------
# Compatibility: old numeric routes redirect to slug
# -------------------------
@app.route("/tests/<int:test_id>/take", methods=["GET"])
def take_test_numeric(test_id):
    t = db().execute("SELECT slug FROM tests WHERE id=?", (test_id,)).fetchone()
    if not t or not t["slug"]:
        abort(404)
    return redirect(f"/tests/{t['slug']}/take")


@app.route("/tests/<int:test_id>/submit", methods=["POST"])
def submit_test_numeric(test_id):
    t = db().execute("SELECT slug FROM tests WHERE id=?", (test_id,)).fetchone()
    if not t or not t["slug"]:
        abort(404)
    return redirect(f"/tests/{t['slug']}/take")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
