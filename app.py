# app.py
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import StreamingResponse, HTMLResponse, PlainTextResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile, os

from converter import convert_pdf_to_excel_with_meta

app = FastAPI(title="PDF → Excel Converter")

# Optional CORS (safe defaults; keep if you later host a separate frontend)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten later if you have a fixed domain
    allow_methods=["*"],
    allow_headers=["*"],
)

MAX_BYTES = 12 * 1024 * 1024  # 12 MB

# ---------- Security headers (CSP, nosniff, referrer) ----------
@app.middleware("http")
async def security_headers(request: Request, call_next):
    resp = await call_next(request)
    # Minimal but useful CSP; allows our inline styles & scripts
    resp.headers["Content-Security-Policy"] = "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; base-uri 'none'; frame-ancestors 'none';"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "no-referrer"
    return resp

# ---------- UI (English) ----------
HTML_FORM = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>PDF → Excel Converter</title>
  <style>
    :root {
      --bg: #0b1020;
      --panel: #10162a;
      --panel-2: #0e1426;
      --fg: #eef2ff;
      --muted: #a7b0c4;
      --accent: #7c9cff;
      --accent-2: #9fe870;
      --border: #1f2a4a;
      --err: #ff5c7c;
      --ok: #31d0aa;
    }
    * { box-sizing: border-box; }
    html, body { height: 100%; }
    body {
      margin: 0;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial;
      background:
        radial-gradient(1200px 800px at 80% -10%, #203061 0%, rgba(32,48,97,0) 60%),
        radial-gradient(1200px 800px at -20% 110%, #17314b 0%, rgba(23,49,75,0) 60%),
        var(--bg);
      color: var(--fg);
      display: grid;
      place-items: center;
      padding: 24px;
    }
    .wrap { width: 100%; max-width: 820px; }
    .hdr { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }
    .logo {
      width: 40px; height: 40px; border-radius: 12px; display:grid; place-items:center;
      background: linear-gradient(135deg, var(--accent) 0%, #6ee7f9 100%);
      color:#0b1020; font-weight: 800;
      box-shadow: 0 10px 30px rgba(124,156,255,.25);
    }
    h1 { margin: 0; font-size: 26px; letter-spacing: .2px; }
    .sub { margin: 4px 0 22px; color: var(--muted); }
    .card {
      background: linear-gradient(180deg, var(--panel) 0%, var(--panel-2) 100%);
      border: 1px solid var(--border);
      border-radius: 18px;
      padding: 22px;
      box-shadow: 0 20px 50px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.03);
    }
    .dropzone {
      border: 1.5px dashed #2a3a68; border-radius: 16px; padding: 28px; text-align: center; cursor: pointer;
      background: rgba(124,156,255,.06);
      transition: border-color .15s ease, background .15s ease, transform .05s ease;
    }
    .dropzone:hover { border-color: var(--accent); background: rgba(124,156,255,.1); }
    .dropzone:active { transform: scale(.997); }
    .dz-title { font-weight: 700; }
    .dz-hint { color: var(--muted); font-size: 13px; margin-top: 6px; }
    .hidden-input { display:none; }

    .actions { display:flex; flex-wrap:wrap; gap:12px; align-items:center; margin-top:14px; }
    button {
      padding: 12px 18px; border-radius: 12px; border: 0; cursor: pointer;
      background: linear-gradient(135deg, var(--accent) 0%, #6ee7f9 100%);
      color:#0b1020; font-weight: 700; letter-spacing:.2px;
      box-shadow: 0 10px 30px rgba(124,156,255,.35);
      transition: transform .06s ease, filter .2s ease, box-shadow .2s ease;
    }
    button:hover { filter: brightness(1.05); box-shadow: 0 14px 36px rgba(124,156,255,.45); }
    button:active { transform: translateY(1px); }
    button:disabled { opacity: .45; cursor: not-allowed; filter: none; box-shadow: none; }

    .filebadge {
      display:inline-flex; gap:8px; align-items:center;
      background: #0b1328; border:1px solid var(--border); color: var(--fg);
      padding: 8px 12px; border-radius: 10px; font-size: 13px;
    }
    .hint { color: var(--muted); font-size: 13px; }

    .status {
      margin-top:16px; padding: 14px; border-radius: 12px;
      background: #0b1328; border:1px solid var(--border); display:none;
    }
    .status .label { color: var(--muted); font-size: 13px; margin-bottom: 6px; }
    .ok { color: var(--ok); }
    .err { color: var(--err); white-space: pre-wrap; }

    .meta { margin-top: 8px; color:#cfd7ef; font-size: 14px; }
    progress { width: 100%; height: 10px; margin-top: 10px; appearance: none; }
    progress::-webkit-progress-bar { background: #0b1320; border-radius: 999px; }
    progress::-webkit-progress-value {
      background: linear-gradient(90deg, var(--accent) 0%, var(--accent-2) 100%);
      border-radius: 999px;
      box-shadow: 0 0 18px rgba(124,156,255,.35);
    }

    .foot { margin-top: 14px; display:flex; justify-content:space-between; gap:12px; flex-wrap:wrap; }
    .foot .left { color: var(--muted); font-size: 12px; }
    .badge {
      background: #0b1328; border:1px solid var(--border); color:#cfd7ef;
      padding: 6px 10px; border-radius: 999px; font-size: 12px;
    }

    @media (max-width: 560px) {
      h1 { font-size: 22px; }
      .dropzone { padding: 22px; }
      .actions { flex-direction: column; align-items: stretch; }
      .filebadge { justify-content: center; }
    }
  </style>
</head>
<body>

<div class="wrap">
  <div class="hdr">
    <div class="logo" aria-hidden="true">×</div>
    <div>
      <h1>PDF → Excel Converter</h1>
      <div class="sub">Drop or select a PDF and we’ll extract tables into a downloadable .xlsx. Scanned PDFs are OCR’d automatically.</div>
    </div>
  </div>

  <div class="card">
    <div id="drop" class="dropzone" role="button" tabindex="0" aria-label="Upload PDF">
      <div class="dz-title">Drop PDF here <span aria-hidden="true">📄⬇️</span></div>
      <div class="dz-hint">…or click to choose a file (max 12&nbsp;MB)</div>
      <input id="file" type="file" accept="application/pdf" class="hidden-input" />
    </div>

    <div class="actions">
      <button id="convertBtn" disabled>Convert to Excel</button>
      <span id="chosen" class="filebadge">No file selected</span>
    </div>

    <div id="status" class="status" aria-live="polite">
      <div class="label">Status</div>
      <div id="statustext">Ready</div>
      <progress id="prog" max="100" value="0" style="display:none"></progress>
      <div id="meta" class="meta"></div>
      <div id="error" class="err"></div>
    </div>

    <div class="foot">
      <div class="left">Built with FastAPI · Camelot · pdfplumber · OCRmyPDF</div>
      <div class="badge">Beta</div>
    </div>
  </div>
</div>

<script>
// Prevent the browser from opening PDFs if dropped outside the dropzone
window.addEventListener('dragover', e => { e.preventDefault(); }, { passive: false });
window.addEventListener('drop',     e => { e.preventDefault(); e.stopPropagation(); }, { passive: false });

(() => {
  const dz = document.getElementById('drop');
  const fi = document.getElementById('file');
  const btn = document.getElementById('convertBtn');
  const chosen = document.getElementById('chosen');
  const statusBox = document.getElementById('status');
  const statusText = document.getElementById('statustext');
  const meta = document.getElementById('meta');
  const error = document.getElementById('error');
  const prog = document.getElementById('prog');
  let theFile = null;

  const MAX_BYTES = 12 * 1024 * 1024;

  function showStatus(msg) {
    statusBox.style.display = 'block';
    statusText.textContent = msg;
  }

  function fmtMB(n) { return (n/1024/1024).toFixed(2) + " MB"; }

  function setFile(file) {
    if (!file) return;
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      showStatus('Only PDF files are supported.');
      error.textContent = 'File type is not PDF.';
      btn.disabled = true;
      return;
    }
    if (file.size > MAX_BYTES) {
      showStatus('File is too large.');
      error.textContent = 'Max 12 MB.';
      btn.disabled = true;
      return;
    }
    theFile = file;
    chosen.textContent = `Selected: ${file.name} (${fmtMB(file.size)})`;
    btn.disabled = false;
    error.textContent = '';
    meta.textContent = '';
  }

  // Accessibility: click/enter/space on the dropzone
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); }
  });

  // Drag & drop
  dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.style.borderColor = 'var(--accent)'; });
  dz.addEventListener('dragleave', () => { dz.style.borderColor = '#2a3a68'; });
  dz.addEventListener('drop', (e) => {
    e.preventDefault(); dz.style.borderColor = '#2a3a68';
    setFile(e.dataTransfer.files?.[0]);
  });

  fi.addEventListener('change', (e) => setFile(e.target.files?.[0]));

  btn.addEventListener('click', async () => {
    if (!theFile) return;
    btn.disabled = true;
    showStatus('Uploading and converting…');
    error.textContent = '';
    prog.style.display = 'block';
    prog.value = 20;

    try {
      const form = new FormData();
      form.append('file', theFile);

      const resp = await fetch('/convert', { method: 'POST', body: form });
      prog.value = 60;

      const ctype = resp.headers.get('content-type') || '';
      if (!resp.ok || !ctype.includes('spreadsheet')) {
        const text = await resp.text();
        error.textContent = text || `Error: ${resp.status}`;
        statusText.textContent = 'Failed';
        btn.disabled = false;
        prog.style.display = 'none';
        return;
      }

      // Metadata from headers
      const tables = resp.headers.get('X-Tables-Found');
      const pages = resp.headers.get('X-Pages');
      const ocr = resp.headers.get('X-OCR-Used') === '1' ? 'Yes' : 'No';
      meta.textContent = `Tables: ${tables ?? '-'} • Pages: ${pages ?? '-'} • OCR: ${ocr}`;

      // Save file
      const blob = await resp.blob();
      prog.value = 90;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const base = theFile.name.replace(/\\.pdf$/i, '');
      a.download = base + '.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      statusText.innerHTML = '<span class="ok">Done! Excel downloaded.</span>';
      prog.style.display = 'none';
    } catch (e) {
      error.textContent = (e && e.message) ? e.message : String(e);
      statusText.textContent = 'Failed';
      prog.style.display = 'none';
    } finally {
      btn.disabled = false;
    }
  });
})();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def index():
    return HTML_FORM

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    # Basic filename check (UX)
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")

    # Read bytes (enforce size)
    raw = await file.read()
    if len(raw) > MAX_BYTES:
        raise HTTPException(status_code=413, detail=f"File too large (> {MAX_BYTES // (1024*1024)} MB).")

    # Quick content signature check: PDF files should start with %PDF
    if not raw.startswith(b"%PDF"):
        raise HTTPException(status_code=400, detail="File content is not a valid PDF.")

    # Work in temp file
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, file.filename)
        with open(pdf_path, "wb") as f:
            f.write(raw)

        try:
            xlsx_bytes, meta = convert_pdf_to_excel_with_meta(pdf_path)
        except ValueError as e:
            # "No tables found" or similar, treat as unprocessable
            return PlainTextResponse(
                f"No tables found.\n\nTips:\n- Try a PDF with clearer table borders or structured columns.\n- For scans: ensure a readable scan (300+ DPI).",
                status_code=422
            )
        except Exception as e:
            # Generic failure
            return PlainTextResponse(
                f"Conversion failed:\n{str(e)}\n\nTips:\n- Try a different PDF with visible table lines.\n- If it is a scanned document: ensure it is readable.\n- Contact us if the problem persists.",
                status_code=500
            )

    headers = {
        "Content-Disposition": f'attachment; filename="{os.path.splitext(file.filename)[0]}.xlsx"',
        "X-Tables-Found": str(meta.get("tables_found")),
        "X-Pages": str(meta.get("pages")),
        "X-OCR-Used": "1" if meta.get("ocr_used") else "0",
    }

    return StreamingResponse(
        iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )
