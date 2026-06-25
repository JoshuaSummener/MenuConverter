#!/usr/bin/env python3
"""
app.py — Menu Scanner web app
=============================
Upload a menu (one or more photos) and download two spreadsheets:

  1. "Userve format" — our combined workbook (Menu / Variants / Section sheets)
                       after edit_menu_excel.py runs on it (extra columns,
                       collapsed section names, Optional Modifiers sheet).
  2. "Ulite format"  — the normalized POS workbook from to_pos_format.py,
                       with the import-guide validation rules enforced.

Pipeline behind the button:
    photos -> menu_folder_to_excel.py (uses Claude vision) -> combined .xlsx
           -> edit_menu_excel.py  -> Menu Excel deliverable
           -> to_pos_format.py    -> POS Import deliverable

Run it:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-...      # or paste it into the form
    python app.py
    open http://127.0.0.1:5000

The extraction step calls the Claude API, so an API key is required.
"""

import os
import sys
import uuid
import zipfile
import subprocess

from flask import (Flask, request, render_template_string, send_file, abort)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.join(BASE_DIR, "menu_to_excel_2.py")
ORCH = os.path.join(BASE_DIR, "menu_folder_to_excel.py")
EDIT = os.path.join(BASE_DIR, "edit_menu_excel.py")
POS = os.path.join(BASE_DIR, "to_pos_format.py")
TEMPLATE = os.path.join(BASE_DIR, "pos_template.xlsx")   # optional base for POS
JOBS = os.path.join(BASE_DIR, "jobs")
os.makedirs(JOBS, exist_ok=True)


def _load_local_keys():
    """Load API keys from a gitignored 'secrets.env' next to app.py, if present.
    File format (one per line):  ANTHROPIC_API_KEY=sk-ant-...   /   GEMINI_API_KEY=AIza...
    Existing environment variables win, so `export ...` still overrides the file.
    Keeping the key here (and in .gitignore) means it is NOT baked into tracked code."""
    path = os.path.join(BASE_DIR, "secrets.env")
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            name, _, value = line.partition("=")
            name, value = name.strip(), value.strip().strip('"').strip("'")
            if name and value and not os.environ.get(name):
                os.environ[name] = value


_load_local_keys()

ALLOWED_EXT = {".png", ".jpg", ".jpeg", ".webp", ".gif"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024      # 80 MB total upload


# --------------------------------------------------------------------------- #
# Pipeline
# --------------------------------------------------------------------------- #

def safe_name(filename):
    """Keep the basename (incl. unicode/digits for page ordering), drop paths."""
    base = os.path.basename(filename.replace("\\", "/"))
    return base.replace("\x00", "").strip() or "upload"


def run(cmd, env=None):
    p = subprocess.run(cmd, capture_output=True, text=True, env=env)
    if p.returncode != 0:
        raise RuntimeError((p.stdout or "") + "\n" + (p.stderr or ""))
    return p.stdout


def extract_to_combined(img_dir, combined_path, provider="claude"):
    """photos -> our combined workbook (runs the vision orchestrator)."""
    run([sys.executable, ORCH, "--folder", img_dir,
         "--output", combined_path, "--engine", ENGINE, "--provider", provider],
        env=os.environ.copy())


def build_outputs(combined_path, job_dir, menu_name):
    """combined workbook -> the two deliverables."""
    userve = os.path.join(job_dir, "menu_userve.xlsx")
    run([sys.executable, EDIT, "--input", combined_path, "--output", userve])

    ulite = os.path.join(job_dir, "menu_ulite.xlsx")
    cmd = [sys.executable, POS, "--input", combined_path, "--output", ulite]
    if menu_name:
        cmd += ["--menu-name", menu_name]
    if os.path.exists(TEMPLATE):
        cmd += ["--template", TEMPLATE]
    run(cmd)
    return userve, ulite


# --------------------------------------------------------------------------- #
# Pages
# --------------------------------------------------------------------------- #

PAGE = """
<!doctype html><html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Menu Scanner</title>
<style>
  :root{--ink:#15233b;--muted:#6b7280;--line:#e5e7eb;--brand:#2f5496;--bg:#eef2f8}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Segoe UI",-apple-system,Roboto,Helvetica,Arial,sans-serif;line-height:1.6}
  .wrap{max-width:680px;margin:48px auto;padding:0 18px}
  .card{background:#fff;border:1px solid var(--line);border-radius:16px;
    padding:30px 30px 26px;box-shadow:0 6px 24px rgba(20,35,59,.06)}
  h1{font-size:24px;margin:0 0 4px}
  p.sub{color:var(--muted);margin:0 0 22px;font-size:14.5px}
  label{display:block;font-weight:600;font-size:13.5px;margin:16px 0 6px}
  input[type=text],input[type=password]{width:100%;padding:10px 12px;border:1px solid var(--line);
    border-radius:9px;font-size:14px}
  .drop{border:2px dashed #c4d0e6;border-radius:12px;padding:26px;text-align:center;
    color:var(--muted);background:#fafcff;cursor:pointer}
  .drop b{color:var(--brand)}
  .files{font-size:13px;color:var(--ink);margin-top:10px}
  button{margin-top:22px;width:100%;background:var(--brand);color:#fff;border:0;border-radius:10px;
    padding:13px;font-size:15px;font-weight:600;cursor:pointer}
  button:hover{background:#26426f}
  .key{margin-top:8px;font-size:12.5px}
  .ok{color:#047857}.bad{color:#b45309}
  .err{background:#fee2e2;border:1px solid #fca5a5;color:#991b1b;border-radius:10px;
    padding:12px 14px;font-size:13.5px;margin-bottom:18px;white-space:pre-wrap}
  .hint{color:var(--muted);font-size:12.5px;margin-top:18px}
  #overlay{display:none;position:fixed;inset:0;background:rgba(238,242,248,.92);
    align-items:center;justify-content:center;flex-direction:column;text-align:center;padding:20px}
  .spin{width:38px;height:38px;border:4px solid #c4d0e6;border-top-color:var(--brand);
    border-radius:50%;animation:r 1s linear infinite;margin-bottom:14px}
  @keyframes r{to{transform:rotate(360deg)}}
</style></head><body>
<div class=wrap><div class=card>
  <h1>🍜 Menu Scanner</h1>
  <p class=sub>Upload your menu photos and get back two spreadsheets: the Userve format and the Ulite format.</p>
  {% if error %}<div class=err>{{ error }}</div>{% endif %}
  <form method=post action="/process" enctype="multipart/form-data" onsubmit="document.getElementById('overlay').style.display='flex'">
    <label>Menu photos (one per page)</label>
    <div class=drop onclick="document.getElementById('imgs').click()">
      <b>Click to choose images</b> &nbsp;·&nbsp; PNG / JPG / WEBP
      <div class=files id=filelist></div>
    </div>
    <input id=imgs type=file name=images accept="image/*" multiple style="display:none"
      onchange="document.getElementById('filelist').textContent=this.files.length+' file(s) selected'">

    <label>Menu name <span style="font-weight:400;color:#6b7280">(optional)</span></label>
    <input type=text name=menu_name placeholder="e.g. China Taste">

    <label>Model provider</label>
    <select name=provider id=provider onchange="keyHint()">
      <option value=claude>Claude (Anthropic)</option>
      <option value=gemini>Gemini (Google)</option>
    </select>

    <label>API key <span style="font-weight:400;color:#6b7280">(for the selected provider)</span></label>
    <input type=password name=api_key id=apikey placeholder="sk-ant-...">
    <div class=key id=keystatus>
      <span class="{{ 'ok' if has_claude else 'bad' }}">{{ '✓' if has_claude else '✗' }} Anthropic key {{ 'detected' if has_claude else 'not set' }}</span>
      &nbsp;·&nbsp;
      <span class="{{ 'ok' if has_gemini else 'bad' }}">{{ '✓' if has_gemini else '✗' }} Gemini key {{ 'detected' if has_gemini else 'not set' }}</span>
    </div>

    <button type=submit>Convert menu</button>
  </form>
  <div class=hint id=hint>Extraction runs one page at a time, so a multi-page menu can take a minute or two.</div>
  <script>
    function keyHint(){
      var p=document.getElementById('provider').value;
      document.getElementById('apikey').placeholder = (p=='gemini') ? 'AIza...' : 'sk-ant-...';
    }
    keyHint();
  </script>
</div></div>
<div id=overlay><div class=spin></div><div><b>Reading your menu…</b><br>This can take a minute or two per page.</div></div>
</body></html>
"""

RESULT = """
<!doctype html><html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Done</title>
<style>
  :root{--ink:#15233b;--muted:#6b7280;--line:#e5e7eb;--brand:#2f5496;--bg:#eef2f8}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:"Segoe UI",-apple-system,Roboto,Arial,sans-serif;line-height:1.6}
  .wrap{max-width:680px;margin:48px auto;padding:0 18px}
  .card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:30px;
    box-shadow:0 6px 24px rgba(20,35,59,.06)}
  h1{font-size:23px;margin:0 0 18px}
  a.dl{display:flex;align-items:center;justify-content:space-between;text-decoration:none;
    color:var(--ink);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin:10px 0;
    font-weight:600}
  a.dl:hover{border-color:var(--brand);background:#fafcff}
  a.dl small{display:block;color:var(--muted);font-weight:400;font-size:12.5px;margin-top:2px}
  a.dl span.go{color:var(--brand);font-size:13px}
  a.both{background:var(--brand);color:#fff;justify-content:center}
  a.back{display:inline-block;margin-top:18px;color:var(--muted);font-size:13.5px}
</style></head><body>
<div class=wrap><div class=card>
  <h1>✅ Your menu is ready</h1>
  <a class=dl href="/download/{{job}}/combined"><span>Original (Claude extraction)<small>Raw Menu / Variants / Section, straight from the API</small></span><span class=go>Download ↓</span></a>
  <a class=dl href="/download/{{job}}/userve"><span>Userve format<small>Menu / Variants / Section + Optional Modifiers</small></span><span class=go>Download ↓</span></a>
  <a class=dl href="/download/{{job}}/ulite"><span>Ulite format<small>POS import file, validated against the import guide</small></span><span class=go>Download ↓</span></a>
  <a class="dl both" href="/download/{{job}}/all"><span>Download all (.zip)</span></a>
  <a class=back href="/">← Convert another menu</a>
</div></div></body></html>
"""


def _key_ctx():
    return {
        "has_claude": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "has_gemini": bool(os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")),
    }


@app.route("/")
def index():
    return render_template_string(PAGE, **_key_ctx())


@app.route("/process", methods=["POST"])
def process():
    files = [f for f in request.files.getlist("images") if f and f.filename]
    if not files:
        return render_template_string(PAGE, error="Please choose at least one menu image.", **_key_ctx())

    provider = request.form.get("provider", "claude").strip().lower()
    if provider not in ("claude", "gemini"):
        provider = "claude"
    key_var = "GEMINI_API_KEY" if provider == "gemini" else "ANTHROPIC_API_KEY"
    other_var = "GOOGLE_API_KEY" if provider == "gemini" else None

    api_key = request.form.get("api_key", "").strip()
    if api_key:
        os.environ[key_var] = api_key
    have = os.environ.get(key_var) or (os.environ.get(other_var) if other_var else None)
    if not have:
        nice = "Gemini" if provider == "gemini" else "Anthropic"
        return render_template_string(PAGE, error=f"A {nice} API key is required to read the menu.", **_key_ctx())

    menu_name = request.form.get("menu_name", "").strip() or None

    job = uuid.uuid4().hex[:12]
    job_dir = os.path.join(JOBS, job)
    img_dir = os.path.join(job_dir, "images")
    os.makedirs(img_dir, exist_ok=True)

    saved = 0
    for f in files:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext in ALLOWED_EXT:
            f.save(os.path.join(img_dir, safe_name(f.filename)))
            saved += 1
    if saved == 0:
        return render_template_string(PAGE, error="No supported image files were uploaded (PNG/JPG/WEBP).", **_key_ctx())

    combined = os.path.join(job_dir, "menu_combined.xlsx")
    try:
        extract_to_combined(img_dir, combined, provider)
        build_outputs(combined, job_dir, menu_name)
    except Exception as exc:
        return render_template_string(PAGE, error="Something went wrong while processing:\n\n" + str(exc), **_key_ctx())
    return render_template_string(RESULT, job=job)


@app.route("/download/<job>/<which>")
def download(job, which):
    job_dir = os.path.join(JOBS, os.path.basename(job))
    names = {"userve": "menu_userve.xlsx", "ulite": "menu_ulite.xlsx", "combined": "menu_combined.xlsx"}

    if which in ("all", "both"):
        zpath = os.path.join(job_dir, "menu_outputs.zip")
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for nm in ("menu_combined.xlsx", "menu_userve.xlsx", "menu_ulite.xlsx"):
                fp = os.path.join(job_dir, nm)
                if os.path.exists(fp):
                    z.write(fp, nm)
        return send_file(zpath, as_attachment=True, download_name="menu_outputs.zip")

    fn = names.get(which)
    if not fn:
        abort(404)
    fp = os.path.join(job_dir, fn)
    if not os.path.exists(fp):
        abort(404)
    return send_file(fp, as_attachment=True, download_name=fn)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
