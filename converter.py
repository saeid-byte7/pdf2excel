# converter.py
import io, os, re, tempfile, subprocess
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import camelot

# -------- Helpers --------
def _pdf_has_selectable_text(pdf_path: str) -> bool:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:
                if (page.extract_text() or "").strip():
                    return True
        return False
    except Exception:
        return False

def _ocr_to_searchable_pdf(input_pdf: str) -> str:
    tmpdir = tempfile.mkdtemp()
    out_pdf = os.path.join(tmpdir, "ocr_out.pdf")
    cmd = [
        "ocrmypdf", "--skip-text", "--force-ocr", "--deskew",
        "-l", "eng+sv", "--optimize", "1", input_pdf, out_pdf
    ]
    subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    return out_pdf

def _clean_cell(x: str):
    if x is None:
        return x
    if not isinstance(x, str):
        return x
    s = x.replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s{2,}", " ", s)

    # Tal som '27.847p' → '27.847' (behåll pence som tal)
    m = re.fullmatch(r"(-?\d+(?:[\.,]\d+)?)[ ]*p", s, flags=re.IGNORECASE)
    if m:
        num = m.group(1).replace(",", ".")
        return num

    # Standardisera decimalkomma -> punkt
    if re.fullmatch(r"-?\d+(?:[\.,]\d+)?", s):
        return s.replace(",", ".")
    return s

def _maybe_headerize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    # Heuristik: om första raden är textigare (unika kolumnnamn) – gör den till header
    head = df.iloc[0].astype(str).apply(lambda v: v.strip())
    if head.nunique(dropna=True) == len(head):
        df = df[1:].copy()
        df.columns = [str(c).strip()[:60] for c in head]
    return df

def _write_excel(dfs: List[pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for i, df in enumerate(dfs, start=1):
            cleaned = df.applymap(_clean_cell)
            # Ta bort helt tomma rader/kolumner
            cleaned.replace(r"^\s*$", pd.NA, regex=True, inplace=True)
            cleaned.dropna(how="all", axis=0, inplace=True)
            cleaned.dropna(how="all", axis=1, inplace=True)
            cleaned = _maybe_headerize(cleaned)
            cleaned.to_excel(writer, index=False, sheet_name=f"Table_{i}")
    return output.getvalue()

# -------- Extractors --------
def _camelot_configs():
    # Testa flera kombinationer – ofta räddar line_scale=40 E.ON-liknande PDF:er
    return [
        {"flavor": "lattice", "line_scale": 40, "strip_text": "\n"},
        {"flavor": "lattice", "line_scale": 20, "strip_text": "\n"},
        {"flavor": "stream",  "row_tol": 10, "column_tol": 10, "strip_text": "\n"},
        {"flavor": "stream",  "row_tol":  5, "column_tol": 15, "strip_text": "\n"},
    ]

def _try_camelot(pdf_path: str) -> Tuple[List[pd.DataFrame], Dict]:
    tried = []
    for cfg in _camelot_configs():
        try:
            tables = camelot.read_pdf(pdf_path, pages="all", **cfg)
            tried.append((cfg, len(tables)))
            if len(tables) > 0:
                dfs = [t.df for t in tables]
                return dfs, {"extractor": f"camelot:{cfg}"}
        except Exception:
            continue
    return [], {"extractor": f"camelot:none", "tried": tried}

def _try_pdfplumber(pdf_path: str) -> Tuple[List[pd.DataFrame], Dict]:
    frames: List[pd.DataFrame] = []
    with pdfplumber.open(pdf_path) as pdf:
        for p, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            for t, tbl in enumerate(tables, start=1):
                frames.append(pd.DataFrame(tbl))
    meta = {"extractor": "pdfplumber"}
    return frames, meta

def _try_tabula(pdf_path: str) -> Tuple[List[pd.DataFrame], Dict]:
    # Kräver Java (lokalt har du det; i Docker lägg till openjdk-headless om du vill köra detta där)
    try:
        import tabula
    except Exception:
        return [], {"extractor": "tabula:not-available"}

    dfs = []
    try:
        # Försök lattice först, sedan stream
        for lattice in (True, False):
            tdfs = tabula.read_pdf(pdf_path, pages="all", lattice=lattice, stream=not lattice, multiple_tables=True)
            for df in tdfs or []:
                if isinstance(df, pd.DataFrame) and not df.empty:
                    dfs.append(df)
            if dfs:
                return dfs, {"extractor": f"tabula:{'lattice' if lattice else 'stream'}"}
    except Exception:
        pass
    return dfs, {"extractor": "tabula:none"}

# -------- Public API --------
def convert_pdf_to_excel_with_meta(pdf_path: str) -> Tuple[bytes, Dict]:
    """
    - OCR vid behov (om PDF saknar text)
    - Camelot med parameter-svep
    - Fallback: pdfplumber
    - Extra fallback (om Java finns): tabula
    """
    source = pdf_path
    used_ocr = False

    if not _pdf_has_selectable_text(pdf_path):
        try:
            source = _ocr_to_searchable_pdf(pdf_path)
            used_ocr = True
        except Exception as e:
            # Fortsätt utan OCR om det fallerar (kan vara text-PDF ändå)
            source = pdf_path

    # 1) Camelot (flera configs)
    dfs, meta = _try_camelot(source)
    if not dfs:
        # 2) pdfplumber
        dfs, meta = _try_pdfplumber(source)

    if not dfs:
        # 3) tabula (om möjligt)
        dfs, meta = _try_tabula(source)

    if not dfs:
        raise ValueError("No tables found after OCR and all extractors.")

    xlsx = _write_excel(dfs)
    meta.update({
        "tables_found": len(dfs),
        "ocr_used": used_ocr,
    })
    try:
        with pdfplumber.open(source) as pdf:
            meta["pages"] = len(pdf.pages)
    except Exception:
        meta["pages"] = None
    return xlsx, meta

# Backward-compat
def convert_pdf_to_excel(pdf_path: str) -> bytes:
    x, _ = convert_pdf_to_excel_with_meta(pdf_path)
    return x
