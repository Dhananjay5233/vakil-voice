import io, os, sqlite3, tempfile, datetime
from flask import Flask, request, jsonify, send_file, send_from_directory

# Create static folder if not exists
os.makedirs("static", exist_ok=True)

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, "letters.db")

app = Flask(__name__, static_folder="static", static_url_path="/static")

# ===== DATABASE =====
def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    try:
        with db() as con:
            con.execute("""CREATE TABLE IF NOT EXISTS letters(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client TEXT, sender TEXT, recipient TEXT,
                letter_date TEXT, subject TEXT, body TEXT, updated_at TEXT)""")
    except Exception as e:
        print(f"Database error: {e}")

init_db()

# ===== ROUTES =====
@app.route("/")
def home():
    try:
        return send_from_directory("static", "index.html")
    except:
        return jsonify(error="index.html not found"), 404

@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory("static", filename)

@app.get("/api/letters")
def list_letters():
    try:
        q = request.args.get("q", "").strip()
        with db() as con:
            if q:
                like = f"%{q}%"
                rows = con.execute("""SELECT id,client,subject,letter_date,updated_at FROM letters
                    WHERE client LIKE ? OR subject LIKE ? ORDER BY id DESC""", (like, like)).fetchall()
            else:
                rows = con.execute("""SELECT id,client,subject,letter_date,updated_at
                    FROM letters ORDER BY id DESC LIMIT 100""").fetchall()
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.get("/api/letters/<int:lid>")
def get_letter(lid):
    try:
        with db() as con:
            r = con.execute("SELECT * FROM letters WHERE id=?", (lid,)).fetchone()
        return (jsonify(dict(r)), 200) if r else (jsonify(error="not found"), 404)
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.post("/api/letters")
def save_letter():
    try:
        d = request.get_json(force=True)
        now = datetime.datetime.now().isoformat(timespec="seconds")
        vals = (d.get("client", ""), d.get("sender", ""), d.get("recipient", ""),
                d.get("letter_date", ""), d.get("subject", ""), d.get("body", ""), now)
        with db() as con:
            if d.get("id"):
                con.execute("""UPDATE letters SET client=?,sender=?,recipient=?,letter_date=?,
                    subject=?,body=?,updated_at=? WHERE id=?""", vals + (d["id"],))
                lid = d["id"]
            else:
                cur = con.execute("""INSERT INTO letters(client,sender,recipient,letter_date,
                    subject,body,updated_at) VALUES(?,?,?,?,?,?,?)""", vals)
                lid = cur.lastrowid
        return jsonify(id=lid)
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.delete("/api/letters/<int:lid>")
def delete_letter(lid):
    try:
        with db() as con:
            con.execute("DELETE FROM letters WHERE id=?", (lid,))
        return jsonify(ok=True)
    except Exception as e:
        return jsonify(error=str(e)), 500

# ===== WORD EXPORT =====
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

def set_hindi_font(run_or_style, name="Mangal"):
    try:
        run_or_style.font.name = name
        el = run_or_style.element.rPr if hasattr(run_or_style.element, "rPr") else None
        if el is not None:
            el.rFonts.set(qn("w:cs"), name)
            el.rFonts.set(qn("w:eastAsia"), name)
    except:
        pass

@app.post("/api/export")
def export_docx():
    try:
        d = request.get_json(force=True)
        doc = Document()
        sec = doc.sections[0]
        sec.top_margin = sec.bottom_margin = Cm(2.5)
        sec.left_margin = sec.right_margin = Cm(2.5)
        st = doc.styles["Normal"]
        st.font.size = Pt(13)
        set_hindi_font(st)

        def para(text, bold=False, align=None):
            p = doc.add_paragraph()
            r = p.add_run(text)
            r.bold = bold
            set_hindi_font(r)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.line_spacing = 1.4
            if align == "right":
                p.alignment = 2
            return p

        para(d.get("sender", ""), bold=True)
        if d.get("letter_date"):
            y, m, dd = d["letter_date"].split("-")
            para(f"दिनांक: {dd}/{m}/{y}", align="right")
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
    except Exception as e:
        return jsonify(error=str(e)), 500

# ===== HEALTH CHECK =====
@app.get("/health")
def health():
    return jsonify(status="ok")

# ===== ERROR HANDLER =====
@app.errorhandler(404)
def not_found(e):
    return jsonify(error="Not found"), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify(error="Server error"), 500

# ===== MAIN =====
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"\n🚀 Vakil Voice running on port {port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)