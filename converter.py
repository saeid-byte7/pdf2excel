# converter.py
import io
import os
import subprocess
import tempfile
from typing import Dict, Tuple

import pandas as pd
import pdfplumber
import camelot


def _pdf_has_selectable_text(pdf_path: str) -> bool:
    """Kolla om PDF:en verkar innehålla text (inte bara bilder)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:  # snabbkoll på första sidorna
                txt = page.extract_text() or ""
                if txt.strip():
                    return True
        return False
    except Exception:
        # Om något går fel, anta att vi behöver OCR.
        return False


def _ocr_to_searchable_pdf(input_pdf: str) -> str:
    """
    Kör ocrmypdf för att skapa en sökbar PDF om den bara är bilder.
    Returnerar sökbar PDF-sökväg (temporär).
    """
    tmpdir = tempfile.mkdtemp()
    output_pdf = os.path.join(tmpdir, "ocr_out.pdf")
    # --skip-text försöker hoppa över sidor som redan har text.
    # --force-ocr säkerställer att bildsidor OCR:as.
    cmd = [
        "ocrmypdf",
        "--skip-text",
        "--force-ocr",
        "--deskew",
        "--optimize", "1",
        input_pdf,
        output_pdf,
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True)
        return output_pdf
    except subprocess.CalledProcessError as e:
        # Om OCR misslyckas, bubbla upp ett begripligt fel.
        raise RuntimeError(
            f"OCR misslyckades. stderr:\n{e.stderr.decode(errors='ignore')}"
        )


def _camelot_tables(pdf_path: str) -> list:
    """
    Försök med Camelot lattice -> stream.
    Returnerar en lista av TableList-entrys (Camelot Table-objekt).
    """
    tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
    if len(tables) == 0:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
    return list(tables)


def _pdfplumber_tables(pdf_path: str) -> list[pd.DataFrame]:
    frames = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                frames.append(pd.DataFrame(tbl))
    return frames


def _write_excel_from_dfs(dfs: list[pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            # Rensa eventuella helt tomma kolumner/rader lite försiktigt
            cleaned = df.copy()
            # Ta bort rader som är helt NaN/None/empty strings
            cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True)
            cleaned.dropna(how="all", inplace=True)
            # Skriv
            cleaned.to_excel(writer, index=False, sheet_name=f"Table_{i+1}")
    return output.getvalue()


def convert_pdf_to_excel_with_meta(pdf_path: str) -> Tuple[bytes, Dict]:
    """
    Huvudfunktion:
    - Säkerställ sökbar PDF via OCR vid behov
    - Extrahera tabeller (Camelot -> pdfplumber fallback)
    - Returnera (xlsx_bytes, metadata)
    """
    source_pdf = pdf_path
    ocr_used = False

    # 1) OCR om nödvändigt
    if not _pdf_has_selectable_text(pdf_path):
        source_pdf = _ocr_to_searchable_pdf(pdf_path)
        ocr_used = True

    # 2) Camelot först
    camelot_tbls = _camelot_tables(source_pdf)
    dfs: list[pd.DataFrame] = []
    if len(camelot_tbls) > 0:
        for t in camelot_tbls:
            dfs.append(t.df)
    else:
        # 3) Fallback: pdfplumber
        dfs = _pdfplumber_tables(source_pdf)

    if not dfs:
        raise ValueError("Hittade inga tabeller i PDF:en efter OCR/fallback.")

    # 4) Skriv Excel
    xlsx_bytes = _write_excel_from_dfs(dfs)

    # 5) Metadata
    meta = {
        "tables_found": len(dfs),
        "ocr_used": ocr_used,
    }
    try:
        with pdfplumber.open(source_pdf) as pdf:
            meta["pages"] = len(pdf.pages)
    except Exception:
        meta["pages"] = None

    return xlsx_bytes, meta


# Bakåtkompatibel wrapper om appen skulle anropa gamla namnet
def convert_pdf_to_excel(pdf_path: str) -> bytes:
    x, _ = convert_pdf_to_excel_with_meta(pdf_path)
    return x
