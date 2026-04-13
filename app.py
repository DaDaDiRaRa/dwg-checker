"""
app.py
======
Two-way cross-check between a master "PDF Drawing List" and the actual
CAD drawings (DWG / DXF) in a target directory.

Pipeline
--------
1. Prompt the user for the Title Block block name (CLI `input()`).
2. Read the drawing list PDF table into a pandas DataFrame
   (Drawing Number / Drawing Name / Scale) using pdfplumber.
3. Parse every .dwg / .dxf in the target directory with ezdxf
   (no AutoCAD required). For every INSERT that matches the
   user-supplied Title Block name:
       - compute its real bounding box with ezdxf.bbox
       - build the "Drawing Number / Drawing Name" search rectangle
         using the bottom-right corner (Max_X, Min_Y) and the
         tweakable X_RATIO / Y_RATIO constants defined below
       - collect TEXT / MTEXT entities whose insertion / center
         point falls inside that rectangle
       - pair them up into Drawing Number + Drawing Name
4. Outer-merge PDF and DWG data on Drawing Number, write the result
   to report.xlsx and highlight any Drawing Name / Scale mismatch
   cells with a red background.

Usage
-----
    python app.py
    python app.py --pdf drawing_list.pdf --dwg-dir ./dwg --out report.xlsx

Notes
-----
* Reading *.dwg* directly requires ezdxf's `odafc` addon, which shells
  out to the free "ODA File Converter" utility. If you only have DXF
  files that step is skipped entirely - ezdxf reads DXF natively.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import pdfplumber
import ezdxf
from ezdxf import bbox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ============================================================================
# GLOBAL TWEAKABLE RATIOS  (requirement #3 - "extract as global variables")
# ============================================================================
# Search rectangle width  = Title-Block Width  * X_RATIO  (to the LEFT of Max_X)
# Search rectangle height = Title-Block Height * Y_RATIO  (ABOVE Min_Y)
X_RATIO: float = 0.101    # 10.1 %
Y_RATIO: float = 0.2138   # 21.38 %


# ============================================================================
# Defaults - can be overridden with CLI flags
# ============================================================================
DEFAULT_PDF_PATH = "drawing_list.pdf"
DEFAULT_DWG_DIR  = "./dwg"
DEFAULT_REPORT   = "report.xlsx"


# Regex used to separate "Drawing Number" (e.g. A-101, S_002, MEP-12B)
# from the longer descriptive "Drawing Name".
_DRAWING_NUMBER_RE = re.compile(
    r"^[A-Za-z]{0,5}[-_]?\d{1,5}[A-Za-z0-9\-_.]*$"
)


# ============================================================================
# 1. PDF EXTRACTION
# ============================================================================
def _find_col(header: List[str], keys: List[str]) -> Optional[int]:
    """Return the index of the first header cell containing any of *keys*."""
    for i, cell in enumerate(header):
        for k in keys:
            if k in cell:
                return i
    return None


def extract_pdf_table(pdf_path: str) -> pd.DataFrame:
    """
    Read every table on every page of *pdf_path* and return a DataFrame
    with columns ``Drawing Number``, ``Drawing Name``, ``Scale``.
    """
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 2:
                    continue

                header = [(c or "").strip().lower() for c in table[0]]
                idx_no    = _find_col(header, ["drawing number", "dwg no", "dwg. no", "no."])
                idx_name  = _find_col(header, ["drawing name", "drawing title", "title", "name"])
                idx_scale = _find_col(header, ["scale"])

                if idx_no is None or idx_name is None:
                    continue  # not a drawing-list table

                for raw in table[1:]:
                    if not raw or all((c or "").strip() == "" for c in raw):
                        continue
                    number = (raw[idx_no]   or "").strip() if idx_no   < len(raw) else ""
                    name   = (raw[idx_name] or "").strip() if idx_name < len(raw) else ""
                    scale  = ""
                    if idx_scale is not None and idx_scale < len(raw):
                        scale = (raw[idx_scale] or "").strip()
                    if not number:
                        continue
                    rows.append({
                        "Drawing Number": number,
                        "Drawing Name":   name,
                        "Scale":          scale,
                    })

    df = pd.DataFrame(rows, columns=["Drawing Number", "Drawing Name", "Scale"])
    if not df.empty:
        df = df.drop_duplicates(subset=["Drawing Number"]).reset_index(drop=True)
    return df


# ============================================================================
# 2. DWG / DXF EXTRACTION
# ============================================================================
def _load_doc(path: Path):
    """Load a DXF directly, or convert a DWG on-the-fly via ODA File Converter."""
    suffix = path.suffix.lower()
    if suffix == ".dxf":
        return ezdxf.readfile(str(path))
    if suffix == ".dwg":
        # Requires the ODA File Converter to be installed on the system.
        from ezdxf.addons import odafc
        return odafc.readfile(str(path))
    raise ValueError(f"Unsupported CAD file extension: {suffix}")


def _entity_point(ent) -> Optional[Tuple[float, float]]:
    """Return a representative (x, y) point for a TEXT or MTEXT entity."""
    t = ent.dxftype()
    if t == "TEXT":
        # Prefer the align point when the text is non-default aligned.
        try:
            if ent.dxf.halign or ent.dxf.valign:
                p = ent.dxf.align_point
            else:
                p = ent.dxf.insert
        except AttributeError:
            p = ent.dxf.insert
        return (float(p[0]), float(p[1]))
    if t == "MTEXT":
        p = ent.dxf.insert
        return (float(p[0]), float(p[1]))
    return None


def _entity_text(ent) -> str:
    t = ent.dxftype()
    if t == "TEXT":
        return (ent.dxf.text or "").strip()
    if t == "MTEXT":
        try:
            return ent.plain_text().strip()
        except Exception:
            return (ent.text or "").strip()
    return ""


def _pair_number_and_name(hits: List[Tuple[float, float, str]]) -> Tuple[str, str]:
    """
    From the texts that fall inside the title-block search rectangle,
    decide which one is the *Drawing Number* and which one is the
    *Drawing Name*. Strategy:

    1. Use a regex to spot drawing-number-looking strings (short,
       alphanumeric, at least one digit).
    2. Everything else is a candidate for Drawing Name.
    3. Fall back to positional logic (topmost text = number) if the
       regex yields nothing useful.
    """
    if not hits:
        return "", ""

    numbers: List[Tuple[float, float, str]] = []
    others:  List[Tuple[float, float, str]] = []
    for x, y, text in hits:
        squashed = text.replace(" ", "")
        if _DRAWING_NUMBER_RE.match(squashed) and any(ch.isdigit() for ch in squashed):
            numbers.append((x, y, text))
        else:
            others.append((x, y, text))

    if numbers and others:
        # Prefer the topmost candidate in each bucket so we pick the
        # most prominent line within the sub-rectangle.
        numbers.sort(key=lambda h: -h[1])
        others.sort(key=lambda h: -h[1])
        return numbers[0][2], others[0][2]

    # Positional fallback - topmost is drawing number.
    hits_sorted = sorted(hits, key=lambda h: -h[1])
    if len(hits_sorted) >= 2:
        return hits_sorted[0][2], hits_sorted[1][2]
    return hits_sorted[0][2], ""


def extract_dwg_data(dwg_dir: str, title_block_name: str) -> pd.DataFrame:
    """
    Walk *dwg_dir* recursively and extract Drawing Number / Drawing Name
    for every INSERT whose block name matches *title_block_name*.
    """
    if not os.path.isdir(dwg_dir):
        raise FileNotFoundError(f"CAD directory not found: {dwg_dir}")

    patterns = ("*.dxf", "*.DXF", "*.dwg", "*.DWG")
    files: List[str] = []
    for pat in patterns:
        files.extend(glob.glob(os.path.join(dwg_dir, "**", pat), recursive=True))
    files = sorted(set(files))

    if not files:
        print(f"[WARN] no .dwg/.dxf files found in {dwg_dir}")
        return pd.DataFrame(columns=["Source File", "Drawing Number", "Drawing Name"])

    target = title_block_name.strip().lower()
    rows: List[dict] = []

    for f in files:
        path = Path(f)
        try:
            doc = _load_doc(path)
        except Exception as exc:
            print(f"[WARN] cannot open {path.name}: {exc}")
            continue

        msp = doc.modelspace()

        # 2a. Find every INSERT that matches the requested title block.
        tb_inserts = [
            ins for ins in msp.query("INSERT")
            if ins.dxf.name.strip().lower() == target
        ]
        if not tb_inserts:
            print(f"[INFO] '{title_block_name}' not found in {path.name}")
            continue

        # Cache ALL text entities once per file (faster than querying
        # per title-block instance).
        text_entities = list(msp.query("TEXT MTEXT"))

        for tb in tb_inserts:
            # 2b. Real bounding box of *this* INSERT
            box = bbox.extents([tb])
            if not box.has_data:
                print(f"[WARN] bbox unavailable for a title block in {path.name}")
                continue

            min_x, min_y = float(box.extmin.x), float(box.extmin.y)
            max_x, max_y = float(box.extmax.x), float(box.extmax.y)
            width  = max_x - min_x
            height = max_y - min_y
            if width <= 0 or height <= 0:
                continue

            # 2c. Search rectangle anchored at bottom-right corner.
            #     Bottom-right corner = (Max_X, Min_Y)
            #     Expand LEFT  by width  * X_RATIO
            #     Expand UP    by height * Y_RATIO
            sx_max = max_x
            sx_min = max_x - width  * X_RATIO
            sy_min = min_y
            sy_max = min_y + height * Y_RATIO

            # 2d. Filter text entities whose insertion point sits inside
            #     the search rectangle.
            hits: List[Tuple[float, float, str]] = []
            for ent in text_entities:
                pt = _entity_point(ent)
                if pt is None:
                    continue
                x, y = pt
                if sx_min <= x <= sx_max and sy_min <= y <= sy_max:
                    content = _entity_text(ent)
                    if content:
                        hits.append((x, y, content))

            if not hits:
                print(f"[INFO] no texts inside search box of {path.name}")
                continue

            dwg_no, dwg_name = _pair_number_and_name(hits)
            rows.append({
                "Source File":    path.name,
                "Drawing Number": dwg_no,
                "Drawing Name":   dwg_name,
            })

    df = pd.DataFrame(rows, columns=["Source File", "Drawing Number", "Drawing Name"])
    if not df.empty:
        df = (
            df[df["Drawing Number"] != ""]
            .drop_duplicates(subset=["Drawing Number"])
            .reset_index(drop=True)
        )
    return df


# ============================================================================
# 3. COMPARE & WRITE EXCEL REPORT
# ============================================================================
def build_report(pdf_df: pd.DataFrame, dwg_df: pd.DataFrame, out_path: str) -> None:
    """Merge the two DataFrames, write to Excel and highlight mismatches."""
    # Rename so Drawing Name / Scale become *_PDF / *_DWG side by side.
    pdf = pdf_df.rename(columns={
        "Drawing Name": "Drawing Name (PDF)",
        "Scale":        "Scale (PDF)",
    })
    dwg = dwg_df.rename(columns={
        "Drawing Name": "Drawing Name (DWG)",
    })

    merged = pdf.merge(dwg, on="Drawing Number", how="outer", indicator=True)

    # Guarantee column order / existence.
    for col in ["Drawing Name (PDF)", "Drawing Name (DWG)",
                "Scale (PDF)", "Source File"]:
        if col not in merged.columns:
            merged[col] = ""

    merged["Match Status"] = merged["_merge"].map({
        "both":        "MATCHED",
        "left_only":   "DWG MISSING",
        "right_only":  "PDF MISSING",
    })
    merged = merged[[
        "Drawing Number",
        "Drawing Name (PDF)",
        "Drawing Name (DWG)",
        "Scale (PDF)",
        "Source File",
        "Match Status",
    ]].fillna("")

    merged.to_excel(out_path, index=False)

    # ----- highlight mismatches with openpyxl ----------------------------
    red = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
    wb  = load_workbook(out_path)
    ws  = wb.active

    header_row = {cell.value: cell.column for cell in ws[1]}
    col_no        = header_row["Drawing Number"]
    col_name_pdf  = header_row["Drawing Name (PDF)"]
    col_name_dwg  = header_row["Drawing Name (DWG)"]
    col_scale     = header_row["Scale (PDF)"]
    col_status    = header_row["Match Status"]

    for row in range(2, ws.max_row + 1):
        status   = ws.cell(row=row, column=col_status).value or ""
        name_pdf = (ws.cell(row=row, column=col_name_pdf).value or "").strip()
        name_dwg = (ws.cell(row=row, column=col_name_dwg).value or "").strip()
        scale    = (ws.cell(row=row, column=col_scale).value    or "").strip()

        # (a) drawing number missing on one side -> paint the whole row.
        if status in ("DWG MISSING", "PDF MISSING"):
            for c in (col_no, col_name_pdf, col_name_dwg, col_scale, col_status):
                ws.cell(row=row, column=c).fill = red
            continue

        # (b) drawing name mismatch on a matched pair.
        if name_pdf.lower() != name_dwg.lower():
            ws.cell(row=row, column=col_name_pdf).fill = red
            ws.cell(row=row, column=col_name_dwg).fill = red

        # (c) scale value missing / blank on a matched pair.
        if scale == "":
            ws.cell(row=row, column=col_scale).fill = red

    wb.save(out_path)
    print(f"[OK] Report saved to {out_path}")


# ============================================================================
# 4. CLI
# ============================================================================
def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Cross-check a PDF drawing list against CAD drawings."
    )
    p.add_argument("--pdf",     default=DEFAULT_PDF_PATH,
                   help=f"Path to the master PDF (default: {DEFAULT_PDF_PATH})")
    p.add_argument("--dwg-dir", default=DEFAULT_DWG_DIR,
                   help=f"Directory holding .dwg/.dxf files (default: {DEFAULT_DWG_DIR})")
    p.add_argument("--out",     default=DEFAULT_REPORT,
                   help=f"Output Excel path (default: {DEFAULT_REPORT})")
    return p.parse_args()


def main() -> None:
    args = _parse_args()

    # Requirement #1 - runtime prompt
    title_block_name = input("Enter the Name of the Title Block to search for: ").strip()
    if not title_block_name:
        print("[ERROR] Title block name cannot be empty.")
        sys.exit(1)

    print(f"[INFO] Parsing PDF drawing list: {args.pdf}")
    pdf_df = extract_pdf_table(args.pdf)
    print(f"       -> {len(pdf_df)} rows from PDF")

    print(f"[INFO] Scanning CAD files in: {args.dwg_dir}")
    dwg_df = extract_dwg_data(args.dwg_dir, title_block_name)
    print(f"       -> {len(dwg_df)} drawings extracted from CAD")

    if pdf_df.empty and dwg_df.empty:
        print("[ERROR] Both PDF and DWG datasets are empty - aborting.")
        sys.exit(1)

    build_report(pdf_df, dwg_df, args.out)


if __name__ == "__main__":
    main()
