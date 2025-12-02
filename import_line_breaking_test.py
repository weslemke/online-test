import sqlite3
from datetime import datetime

DB = "test.db"

TEST_TITLE = "Line Breaking Final Exam"
PASS_SCORE = 100

# Put the correct letter for each question here (A/B/C/D).
# IMPORTANT: Replace the "A" placeholders with the real correct answers.
ANSWER_KEY = {
    1: "A",
    2: "D",
    3: "A",
    4: "B",
    5: "B",
    6: "A",
    7: "B",  # True/False: use A for True, B for False (we map below)
    8: "A",
    9: "B",
    10: "A",
}

QUESTIONS = [
    (1, "Intentionally opening a pipe, line or duct for the purpose of cleaning, inspection, maintenance or replacing components within a system is referred to as:",
        ["Line Breaking", "Pipe Sealing", "System Flushing", "Duct Isolation"]),
    (2, "Routine line breaking instructions can be found in the:",
        ["Emergency action plan (EAP)", "PPE assessment (PPEA)", "Safety Data Sheets (SDS)", "Safe operation procedure (SOP)"]),
    (3, "Workers performing non-routine line-breaking tasks must obtain a(n) [ blank ] prior to starting work to identify the risks and controls that are needed.",
        ["Safe Work Permit", "Safe operating Procedure (SOP)", "Emergency action plan (EAP)", "Open-end wrench"]),
    (4, "What type of flange slips over the pipe without needing to be welded to it and can swivel around the pipe to help line up opposing bolt holes?",
        ["Orifice plate", "Lap joint flange", "Spectacle blind", "Blind flange"]),
    (5, "The minimum PPE for line-breaking jobs include a hard hat/helmet, gloves, face shield, goggles/safety glasses and:",
        ["Hair net", "Chemical-protective clothing", "Disposable paper gown", "Shoe cover"]),
    (6, 'To work in a defensive position opening a flange, begin by loosening the bolts [ blank. ]',
        ["Farthest away from you.", "That are close to you first, working clockwise around the line.", "While crouched as low to the ground as possible.", "With open-end wrenches pointed toward your body."]),
    (7, "When performing line breaking, it's best to use a cheater bar to gain leverage.",
        ["True", "False", "—", "—"]),
    (8, "Before beginning line breaking, which of the following steps should be taken?",
        ["Ensure lines and equipment are as free from recognized hazards as possible and test to confirm.",
         "Increase the pressure within the lines to check for leaks.",
         "Remove all tags from the equipment to allow for easy access.",
         "Turn on all cathodic protection rectifiers affecting the piping."]),
    (9, "If the proper steps of bleeding off pressure are not followed, a flammable chemical release can create a potential ignition source, leading to:",
        ["Biological hazards", "Fire and explosion.", "Slips, trips, and falls.", "Ultraviolet light exposure."]),
    (10, "What should you do if you are unsure about which type of flange or procedure to use for line breaking?",
        ["Check schematics and diagrams.", "Consult the human resources department.", "Perform the line break as best you can.", 'Use a "shop-made" flange.']),
]

def create_schema(conn):
    conn.executescript("""
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS tests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      pass_score INTEGER NOT NULL DEFAULT 70,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS questions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      test_id INTEGER NOT NULL,
      prompt TEXT NOT NULL,
      a TEXT NOT NULL,
      b TEXT NOT NULL,
      c TEXT NOT NULL,
      d TEXT NOT NULL,
      correct CHAR(1) NOT NULL CHECK(correct IN ('A','B','C','D')),
      FOREIGN KEY(test_id) REFERENCES tests(id) ON DELETE CASCADE
    );
    """)

def main():
    conn = sqlite3.connect(DB)
    try:
        create_schema(conn)
        cur = conn.execute(
            "INSERT INTO tests(title, pass_score, created_at) VALUES (?,?,?)",
            (TEST_TITLE, PASS_SCORE, datetime.utcnow().isoformat())
        )
        test_id = cur.lastrowid

        for qnum, prompt, opts in QUESTIONS:
            # Map True/False to A/B. (A=True, B=False)
            correct = ANSWER_KEY.get(qnum, "A").upper()
            if qnum == 7:
                if correct == "T": correct = "A"
                if correct == "F": correct = "B"
            if correct not in ("A","B","C","D"):
                raise ValueError(f"Question {qnum}: correct must be A/B/C/D (got {correct})")

            a, b, c, d = opts[0], opts[1], opts[2], opts[3]
            conn.execute("""
                INSERT INTO questions(test_id, prompt, a, b, c, d, correct)
                VALUES (?,?,?,?,?,?,?)
            """, (test_id, prompt, a, b, c, d, correct))

        conn.commit()
        print(f"✅ Imported '{TEST_TITLE}' as test_id={test_id}")
        print(f"Student link: http://127.0.0.1:5000/tests/{test_id}/take")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
