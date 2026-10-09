import io, os, sqlite3, datetime
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Pt, Cm

os.makedirs("static", exist_ok=True)
BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, "letters.db")
app = Flask(__name__)

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
    except: pass

init_db()

HTML = """<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>वकील वॉइस</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;600&display=swap" rel="stylesheet">
<style>
:root{--navy:#0f2742;--gold:#b8893b;--bg:#f4f1ea;--card:#fffdf9}
*{box-sizing:border-box}
body{margin:0;display:flex;min-height:100vh;background:var(--bg);font-family:"Noto Sans Devanagari",sans-serif}
#side{width:250px;background:linear-gradient(180deg,var(--navy),#0a1b2f);color:#fff;padding:20px;display:flex;flex-direction:column;height:100vh;position:sticky;top:0}
.brand{margin-bottom:20px}
.brand h1{margin:0;font-size:1.1rem}
.btn-new{width:100%;padding:10px;border:0;border-radius:8px;background:var(--gold);color:var(--navy);font-weight:600;cursor:pointer;margin-bottom:10px}
#list{flex:1;overflow:auto;list-style:none;margin:0;padding:0}
#list li{padding:8px;cursor:pointer;border-radius:6px}
#list li:hover{background:#ffffff20}
main{flex:1;padding:20px 40px}
.top h2{color:var(--navy);margin:0 0 20px}
.card{background:var(--card);border:1px solid #e3ddd0;border-radius:12px;padding:16px;margin-bottom:16px}
.field{margin-bottom:12px}
.field label{display:block;font-size:.85rem;font-weight:600;color:#6b7685;margin-bottom:5px}
input, textarea, select{width:100%;font:inherit;font-size:.95rem;padding:8px;border:1px solid #e3ddd0;border-radius:8px}
textarea{min-height:120px;resize:vertical}
.field-group{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.field-wrapper{display:flex;gap:8px}
.field-wrapper textarea{flex:1}
.mic-btn{width:40px;height:40px;border-radius:50%;border:0;cursor:pointer;background:var(--navy);color:#fff;font-size:1.2rem;transition:.2s;flex-shrink:0}
.mic-btn:hover{transform:scale(1.1)}
.mic-btn.on{background:#c0392b;animation:pulse 1s infinite}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 rgba(192,57,43,.7)}50%{box-shadow:0 0 0 8px rgba(192,57,43,0)}}
.actions{display:flex;gap:8px;margin-top:16px}
.btn{padding:8px 16px;border:1px solid #e3ddd0;border-radius:8px;background:#fff;cursor:pointer}
.btn:hover{background:#fffaf0}
.btn-primary{background:var(--navy);color:#fff;border:0}
.interim{font-size:.85rem;margin-top:8px;min-height:1.5em;color:var(--gold);font-weight:600}
.interim.rec{color:#c0392b}
</style>
</head>
<body>
<aside id="side">
<div class="brand"><h1>⚖️ वकील वॉइस</h1></div>
<button class="btn-new" onclick="newLetter()">+ नया पत्र</button>
<input style="width:100%;padding:6px;margin-bottom:10px;border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.1);color:#fff;border-radius:6px;font-size:.9rem" id="search" placeholder="खोजें..." oninput="searchLetters()">
<ul id="list"></ul>
</aside>

<main>
<div class="top"><h2>आवाज़ से पत्र लिखें</h2></div>

<div class="card">
<div class="field-group">
<div class="field"><label>मुवक्किल</label><input id="client"></div>
<div class="field"><label>दिनांक</label><input id="date" type="date"></div>
</div>
</div>

<div class="card">
<div class="field"><label>अधिवक्ता का नाम व पता</label>
<div class="field-wrapper">
<textarea id="from" placeholder="नाम, पता, फोन"></textarea>
<button class="mic-btn" onclick="startMic('from')">🎤</button>
</div>
</div>

<div class="field"><label>प्राप्तकर्ता</label>
<div class="field-wrapper">
<textarea id="to" placeholder="किसे भेजना है"></textarea>
<button class="mic-btn" onclick="startMic('to')">🎤</button>
</div>
</div>

<div class="field"><label>विषय</label>
<div class="field-wrapper">
<input id="subject" placeholder="पत्र का विषय">
<button class="mic-btn" onclick="startMic('subject')" style="height:37px">🎤</button>
</div>
</div>
</div>

<div class="card">
<label style="display:block;font-weight:600;margin-bottom:8px">पत्र की सामग्री</label>
<div class="field-wrapper" style="align-items:flex-start">
<textarea id="body" placeholder="यहाँ लिखें..."></textarea>
<button class="mic-btn" style="width:50px;height:50px;margin-top:8px" onclick="startMic('body')">🎤</button>
</div>
<div class="interim" id="status">👂 तैयार हूँ</div>
</div>

<div class="actions">
<button class="btn btn-primary" onclick="saveLetter()">💾 सहेजें</button>
<button class="btn" onclick="exportWord()">⬇️ Word</button>
<button class="btn" onclick="window.print()">🖨️ प्रिंट</button>
</div>
</main>

<script>
const $ = id => document.getElementById(id);

let rec = null;
let isRecording = false;
let currentField = null;

// Initialize date
$('date').valueAsDate = new Date();

// Initialize speech recognition
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
if (SpeechRecognition) {
  rec = new SpeechRecognition();
  rec.lang = 'hi-IN';
  rec.continuous = true;
  rec.interimResults = true;

  rec.onstart = () => {
    console.log('🎤 Recording started for:', currentField);
    $('status').textContent = '🎤 सुन रहा हूँ...';
    $('status').classList.add('rec');
  };

  rec.onresult = (event) => {
    console.log('📝 Recognition result:', event.results.length);
    
    let interim = '';
    let final = '';
    
    for (let i = event.resultIndex; i < event.results.length; i++) {
      const transcript = event.results[i][0].transcript;
      
      if (event.results[i].isFinal) {
        final += transcript + ' ';
        console.log('✅ Final:', transcript);
      } else {
        interim += transcript;
        console.log('⏳ Interim:', transcript);
      }
    }
    
    // Add final text to field
    if (final) {
      addToField(final);
    }
    
    // Update interim display
    if (interim) {
      $('status').textContent = '🎤 सुन रहा हूँ: ' + interim;
    }
  };

  rec.onerror = (event) => {
    console.error('❌ Error:', event.error);
    $('status').textContent = '❌ Error: ' + event.error;
  };

  rec.onend = () => {
    console.log('🛑 Recording ended');
    if (isRecording && currentField) {
      try {
        rec.start();
      } catch (e) {
        console.error('Restart failed:', e);
      }
    }
  };
}

function addToField(text) {
  if (!currentField) return;
  
  const field = $(currentField);
  if (!field) {
    console.error('Field not found:', currentField);
    return;
  }

  console.log('📝 Adding to', currentField, ':', text);
  
  // Get cursor position
  const start = field.selectionStart || field.value.length;
  const end = field.selectionEnd || start;
  
  const before = field.value.substring(0, start);
  const after = field.value.substring(end);
  
  // Add space if needed
  const space = (before && before[before.length - 1] !== ' ' && before[before.length - 1] !== '\n') ? ' ' : '';
  
  // Update field
  field.value = before + space + text + after;
  
  // Move cursor
  const newPos = before.length + space.length + text.length;
  field.setSelectionRange(newPos, newPos);
  
  console.log('✅ Field updated');
}

function startMic(fieldId) {
  if (!rec) {
    alert('Chrome खोलें');
    return;
  }

  // If same field already recording, stop it
  if (isRecording && currentField === fieldId) {
    stopMic();
    return;
  }

  // Stop previous recording
  if (isRecording) {
    try { rec.stop(); } catch (e) {}
  }

  // Start new recording
  isRecording = true;
  currentField = fieldId;
  
  // Update button state
  document.querySelectorAll('.mic-btn').forEach(btn => btn.classList.remove('on'));
  event.target.classList.add('on');
  
  // Start recognition
  try {
    rec.start();
    console.log('🎤 Mic started for:', fieldId);
  } catch (e) {
    console.error('Start failed:', e);
    $('status').textContent = '❌ Mic error: ' + e.message;
  }
}

function stopMic() {
  isRecording = false;
  currentField = null;
  
  try { 
    rec.stop(); 
  } catch (e) {}
  
  document.querySelectorAll('.mic-btn').forEach(btn => btn.classList.remove('on'));
  $('status').textContent = '👂 तैयार हूँ';
  $('status').classList.remove('rec');
  
  console.log('🛑 Mic stopped');
}

function newLetter() {
  $('client').value = '';
  $('from').value = '';
  $('to').value = '';
  $('subject').value = '';
  $('body').value = '';
  $('date').valueAsDate = new Date();
}

async function saveLetter() {
  const data = {
    client: $('client').value,
    sender: $('from').value,
    recipient: $('to').value,
    letter_date: $('date').value,
    subject: $('subject').value,
    body: $('body').value
  };
  
  const r = await fetch('/api/letters', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });
  
  const j = await r.json();
  alert('सहेज लिया ✓');
  refreshList();
}

async function refreshList() {
  const r = await fetch('/api/letters');
  const rows = await r.json();
  $('list').innerHTML = rows.map(x => `
    <li onclick="loadLetter(${x.id})">
      <b>${x.client || '(बिना नाम)'}</b>
      <small>${x.subject || ''}</small>
    </li>
  `).join('');
}

async function loadLetter(id) {
  const r = await fetch('/api/letters/' + id);
  const d = await r.json();
  $('client').value = d.client;
  $('from').value = d.sender;
  $('to').value = d.recipient;
  $('date').value = d.letter_date;
  $('subject').value = d.subject;
  $('body').value = d.body;
}

function searchLetters() {
  const q = $('search').value;
  if (q) {
    fetch('/api/letters?q=' + encodeURIComponent(q))
      .then(r => r.json())
      .then(rows => {
        $('list').innerHTML = rows.map(x => `
          <li onclick="loadLetter(${x.id})">
            <b>${x.client}</b>
            <small>${x.subject}</small>
          </li>
        `).join('');
      });
  }
}

async function exportWord() {
  const data = {
    sender: $('from').value,
    recipient: $('to').value,
    letter_date: $('date').value,
    subject: $('subject').value,
    body: $('body').value
  };
  
  const r = await fetch('/api/export', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });
  
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = ($('client').value || 'letter') + '.docx';
  a.click();
}

refreshList();
</script>
</body>
</html>"""

@app.route("/")
def home():
    return HTML

@app.get("/api/letters")
def list_letters():
    q = request.args.get("q", "").strip()
    with db() as con:
        if q:
            like = f"%{q}%"
            rows = con.execute("SELECT id,client,subject FROM letters WHERE client LIKE ? OR subject LIKE ? ORDER BY id DESC", (like, like)).fetchall()
        else:
            rows = con.execute("SELECT id,client,subject FROM letters ORDER BY id DESC LIMIT 100").fetchall()
    return jsonify([dict(r) for r in rows])

@app.get("/api/letters/<int:id>")
def get_letter(id):
    with db() as con:
        r = con.execute("SELECT * FROM letters WHERE id=?", (id,)).fetchone()
    return jsonify(dict(r)) if r else ("", 404)

@app.post("/api/letters")
def save_letter():
    d = request.get_json()
    now = datetime.datetime.now().isoformat()
    with db() as con:
        con.execute("INSERT INTO letters(client,sender,recipient,letter_date,subject,body,updated_at) VALUES(?,?,?,?,?,?,?)",
            (d.get("client"), d.get("sender"), d.get("recipient"), d.get("letter_date"), d.get("subject"), d.get("body"), now))
    return jsonify(ok=True)

@app.post("/api/export")
def export_docx():
    d = request.get_json()
    doc = Document()
    
    p = doc.add_paragraph(d.get("sender", ""))
    p.runs[0].bold = True
    
    if d.get("letter_date"):
        y, m, dd = d["letter_date"].split("-")
        doc.add_paragraph(f"दिनांक: {dd}/{m}/{y}")
    
    if d.get("recipient"):
        doc.add_paragraph("सेवा में,\n" + d["recipient"])
    
    if d.get("subject"):
        p = doc.add_paragraph(d["subject"])
        p.runs[0].bold = True
    
    doc.add_paragraph(d.get("body", ""))
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    return send_file(buf, as_attachment=True, download_name="letter.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"\n🚀 Vakil Voice on http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)