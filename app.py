import io, os, re, sqlite3, datetime, secrets
from functools import wraps
from flask import (Flask, request, jsonify, send_file, render_template,
                   redirect, url_for, session, g)
from werkzeug.security import generate_password_hash, check_password_hash
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE, "letters.db"))


# ---------- SECRET KEY (sessions/login के लिए) ----------
def load_secret():
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    # gunicorn के सभी workers एक ही key इस्तेमाल करें, इसलिए file में रखते हैं
    path = os.path.join(BASE, ".secret_key")
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "w") as f:
            f.write(secrets.token_hex(32))
    except FileExistsError:
        pass
    for _ in range(50):
        with open(path) as f:
            key = f.read().strip()
        if key:
            return key
        import time; time.sleep(0.05)
    raise RuntimeError("secret key file empty")


app = Flask(__name__)
app.secret_key = load_secret()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(days=30),
)


# ---------- DATABASE ----------
def db():
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            login TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            bar_no TEXT DEFAULT '', address TEXT DEFAULT '', phone TEXT DEFAULT '',
            created_at TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS letters(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client TEXT, sender TEXT, recipient TEXT,
            letter_date TEXT, subject TEXT, body TEXT, updated_at TEXT)""")
        cols = [r["name"] for r in con.execute("PRAGMA table_info(letters)")]
        if "user_id" not in cols:
            con.execute("ALTER TABLE letters ADD COLUMN user_id INTEGER")
        if "created_at" not in cols:
            con.execute("ALTER TABLE letters ADD COLUMN created_at TEXT")
            con.execute("UPDATE letters SET created_at=updated_at WHERE created_at IS NULL")
        con.execute("CREATE INDEX IF NOT EXISTS idx_letters_user ON letters(user_id, updated_at)")


init_db()


def now_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


# ---------- AUTH HELPERS ----------
def current_user():
    if "user" in g:
        return g.user
    g.user = None
    uid = session.get("uid")
    if uid:
        with db() as con:
            g.user = con.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
        if g.user is None:
            session.clear()
    return g.user


@app.context_processor
def inject_user():
    return {"user": current_user()}


def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        if not current_user():
            if request.path.startswith("/api/"):
                return jsonify(error="login required"), 401
            return redirect(url_for("login", next=request.path))
        return view(*a, **kw)
    return wrapped


def normalize_login(s):
    s = (s or "").strip().lower()
    digits = re.sub(r"\D", "", s)
    if "@" not in s and len(digits) >= 10:
        return digits[-10:]  # मोबाइल नंबर: +91 / 0 / स्पेस हटाकर आखिरी 10 अंक
    return s


def valid_login(s):
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s) or re.fullmatch(r"\d{10}", s))


def safe_next(nxt):
    return nxt if nxt and nxt.startswith("/") and not nxt.startswith("//") else url_for("editor")


# ---------- PAGES ----------
@app.route("/")
def home():
    return render_template("home.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("editor"))
    error, login_val = None, ""
    if request.method == "POST":
        login_val = normalize_login(request.form.get("login"))
        pw = request.form.get("password") or ""
        with db() as con:
            u = con.execute("SELECT * FROM users WHERE login=?", (login_val,)).fetchone()
        if u and check_password_hash(u["password_hash"], pw):
            session.clear()
            session.permanent = True
            session["uid"] = u["id"]
            return redirect(safe_next(request.args.get("next")))
        error = "ईमेल/मोबाइल या पासवर्ड गलत है। दोबारा जाँचकर लिखें।"
    return render_template("auth.html", mode="login", error=error, login_val=request.form.get("login", ""))


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if current_user():
        return redirect(url_for("editor"))
    error = None
    form = request.form
    if request.method == "POST":
        name = (form.get("name") or "").strip()
        login_val = normalize_login(form.get("login"))
        pw = form.get("password") or ""
        if not name:
            error = "अपना नाम लिखें।"
        elif not valid_login(login_val):
            error = "सही ईमेल या 10 अंकों का मोबाइल नंबर लिखें।"
        elif len(pw) < 6:
            error = "पासवर्ड कम से कम 6 अक्षर का रखें।"
        else:
            try:
                with db() as con:
                    first_user = con.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0
                    cur = con.execute(
                        "INSERT INTO users(name,login,password_hash,created_at) VALUES(?,?,?,?)",
                        (name, login_val, generate_password_hash(pw), now_iso()))
                    uid = cur.lastrowid
                    # पुराने (login से पहले के) पत्र पहले खाते को मिल जाएँ
                    if first_user:
                        con.execute("UPDATE letters SET user_id=? WHERE user_id IS NULL", (uid,))
                session.clear()
                session.permanent = True
                session["uid"] = uid
                return redirect(url_for("editor"))
            except sqlite3.IntegrityError:
                error = "इस ईमेल/मोबाइल से खाता पहले से बना है। लॉग इन करें।"
    return render_template("auth.html", mode="signup", error=error,
                           name_val=form.get("name", ""), login_val=form.get("login", ""))


@app.route("/logout", methods=["POST", "GET"])
def logout():
    session.clear()
    return redirect(url_for("home"))


@app.route("/app")
@login_required
def editor():
    return render_template("editor.html")


# ---------- PROFILE API ----------
@app.get("/api/profile")
@login_required
def get_profile():
    u = current_user()
    return jsonify(name=u["name"], bar_no=u["bar_no"] or "", address=u["address"] or "",
                   phone=u["phone"] or "", login=u["login"])


@app.post("/api/profile")
@login_required
def save_profile():
    d = request.get_json(silent=True) or {}
    name = (d.get("name") or "").strip() or current_user()["name"]
    with db() as con:
        con.execute("UPDATE users SET name=?,bar_no=?,address=?,phone=? WHERE id=?",
                    (name, d.get("bar_no", ""), d.get("address", ""), d.get("phone", ""),
                     current_user()["id"]))
    return jsonify(ok=True)


# ---------- LETTERS API (हर user के अपने पत्र) ----------
@app.get("/api/letters")
@login_required
def list_letters():
    uid = current_user()["id"]
    q = request.args.get("q", "").strip()
    with db() as con:
        if q:
            like = f"%{q}%"
            rows = con.execute("""SELECT id,client,subject,letter_date,updated_at FROM letters
                WHERE user_id=? AND (client LIKE ? OR subject LIKE ? OR body LIKE ? OR recipient LIKE ?)
                ORDER BY updated_at DESC""", (uid, like, like, like, like)).fetchall()
        else:
            rows = con.execute("""SELECT id,client,subject,letter_date,updated_at FROM letters
                WHERE user_id=? ORDER BY updated_at DESC LIMIT 500""", (uid,)).fetchall()
    return jsonify([dict(r) for r in rows])


@app.get("/api/letters/<int:lid>")
@login_required
def get_letter(lid):
    with db() as con:
        r = con.execute("SELECT * FROM letters WHERE id=? AND user_id=?",
                        (lid, current_user()["id"])).fetchone()
    return (jsonify(dict(r)), 200) if r else (jsonify(error="not found"), 404)


@app.post("/api/letters")
@login_required
def save_letter():
    d = request.get_json(silent=True)
    if d is None:
        return jsonify(error="bad request"), 400
    uid = current_user()["id"]
    now = now_iso()
    vals = (d.get("client", ""), d.get("sender", ""), d.get("recipient", ""),
            d.get("letter_date", ""), d.get("subject", ""), d.get("body", ""), now)
    with db() as con:
        lid = d.get("id")
        if lid:
            cur = con.execute("""UPDATE letters SET client=?,sender=?,recipient=?,letter_date=?,
                subject=?,body=?,updated_at=? WHERE id=? AND user_id=?""", vals + (lid, uid))
            if cur.rowcount == 0:
                return jsonify(error="not found"), 404
        else:
            cur = con.execute("""INSERT INTO letters(client,sender,recipient,letter_date,
                subject,body,updated_at,created_at,user_id) VALUES(?,?,?,?,?,?,?,?,?)""",
                vals + (now, uid))
            lid = cur.lastrowid
    return jsonify(id=lid, updated_at=now)


@app.delete("/api/letters/<int:lid>")
@login_required
def delete_letter(lid):
    with db() as con:
        con.execute("DELETE FROM letters WHERE id=? AND user_id=?", (lid, current_user()["id"]))
    return jsonify(ok=True)


# ---------- WORD EXPORT ----------
def set_hindi_font(obj, name="Mangal"):
    try:
        obj.font.name = name
        rpr = obj.element.rPr
        if rpr is not None:
            rpr.rFonts.set(qn("w:cs"), name)
            rpr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass


@app.post("/api/export")
@login_required
def export_docx():
    d = request.get_json(silent=True) or {}
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Cm(2.5)
    sec.left_margin = sec.right_margin = Cm(2.5)
    st = doc.styles["Normal"]
    st.font.size = Pt(13)
    set_hindi_font(st)

    def para(text, bold=False, right=False):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.bold = bold
        set_hindi_font(r)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.4
        if right:
            p.alignment = 2
        return p

    if d.get("sender"):
        para(d["sender"], bold=True)
    if d.get("letter_date"):
        try:
            y, m, dd = d["letter_date"].split("-")
            para(f"दिनांक: {dd}/{m}/{y}", right=True)
        except ValueError:
            pass
    if d.get("recipient"):
        para("सेवा में,\n" + d["recipient"])
    if d.get("subject"):
        para("विषय: " + d["subject"], bold=True)
    para(d.get("body", ""))
    if d.get("sender"):
        para("\nभवदीय,\n" + d["sender"].split("\n")[0] + "\nअधिवक्ता")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="letter.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/health")
def health():
    return jsonify(status="ok")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
