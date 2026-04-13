"""
app.py  —  PDF <-> CAD Drawing Cross-Check  (V6)
================================================
Compares a master "PDF Drawing List" against the actual CAD drawings
(.dwg / .dxf) living in a directory, without moving or copying any files
(so external references / Xref paths remain intact).

Pipeline
--------
1. Interactive prompts:
      TARGET_DIR   - full directory path holding the DWG files
      PDF_PATH     - full path of the master PDF drawing list
      BLOCK_NAME   - name of the Title Block to search for
2. PDF extraction      : pdfplumber -> pandas DataFrame
                         [Drawing Number, Drawing Name, Scale]
3. DWG extraction      : ezdxf, Model Space only
      * os.chdir(TARGET_DIR) so relative xref paths resolve naturally
      * find every INSERT whose block name matches BLOCK_NAME
      * compute its bounding box via ezdxf.bbox.extents
      * build a search window anchored at the bottom-right corner
        (Max_X, Min_Y) using the global X_RATIO / Y_RATIO constants
      * collect TEXT / MTEXT entities whose insertion point lies inside
        that window, then pair them into Drawing Number + Drawing Name
4. Compare & Report   : outer-merge on Drawing Number, write report.xlsx
                        and highlight Drawing Name / Scale mismatches in red

Xref note
---------
ezdxf reads entity data straight from the *host* DWG. If the Title Block
itself is an Xref, the INSERT entity still lives in the host file, so it
is picked up by ``msp.query("INSERT")`` as usual. However, text that
resides *inside* an unresolved xref cannot be read without loading the
xref body — this script intentionally operates only on the host's
coordinate system and the host's own TEXT / MTEXT entities.

DWG note
--------
ezdxf reads .dxf files natively. To read binary .dwg directly the free
ODA File Converter is required (``ezdxf.addons.odafc``). If it is not
installed, batch-convert the .dwg files to .dxf first — the script
happily processes whichever extension it finds in the directory.
"""

from __future__ import annotations

import glob
import os
import re
import sys
import traceback
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import pdfplumber
import ezdxf
from ezdxf import bbox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ============================================================================
# GLOBAL TWEAKABLE RATIOS  (requirement: expose as top-level constants)
# ============================================================================
X_RATIO: float = 0.101    # 10.10 %  search width  = Title-Block Width  * X_RATIO
Y_RATIO: float = 0.2138   # 21.38 %  search height = Title-Block Height * Y_RATIO


# Regex used to separate a short "Drawing Number" (A-101, S_002, MEP-12B ...)
# from longer descriptive "Drawing Name" strings.
_DRAWING_NUMBER_RE = re.compile(
    r"^[A-Za-z]{0,5}[-_]?\d{1,5}[A-Za-z0-9\-_.]*$"
)

# Default name of the Excel report (written to the cwd where app.py is run).
REPORT_NAME = "report.xlsx"


# ============================================================================
# 1.  INTERACTIVE CLI PROMPTS
# ============================================================================
def _prompt_path(label: str, *, must_be_dir: bool = False,
                 must_be_file: bool = False) -> str:
    """Prompt the user for a filesystem path and validate it."""
    while True:
        raw = input(label).strip().strip('"').strip("'")
        if not raw:
            print("    ! Empty input. Please try again.")
            continue
        path = os.path.expanduser(os.path.expandvars(raw))
        if must_be_dir and not os.path.isdir(path):
            print(f"    ! Not a valid directory: {path}")
            continue
        if must_be_file and not os.path.isfile(path):
            print(f"    ! Not a valid file: {path}")
            continue
        return os.path.abspath(path)


def prompt_inputs() -> Tuple[str, str, str]:
    """Collect TARGET_DIR, PDF_PATH and BLOCK_NAME from the user."""
    print("=" * 72)
    print(" PDF <-> CAD Drawing Cross-Check")
    print("=" * 72)

    target_dir = _prompt_path(
        "Enter the full directory path that contains the DWG files: ",
        must_be_dir=True,
    )
    pdf_path = _prompt_path(
        "Enter the full path of the master PDF drawing list: ",
        must_be_file=True,
    )
    block_name = input("Enter the Name of the Title Block to search for: ").strip()
    if not block_name:
        print("[ERROR] Title block name cannot be empty.")
        sys.exit(1)

    return target_dir, pdf_path, block_name


# ============================================================================
# 2.  PDF EXTRACTION
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
    Parse every table on every page of *pdf_path* and return a DataFrame
    with columns ``Drawing Number``, ``Drawing Name``, ``Scale``.
    """
    print(f"[PDF ] Opening: {pdf_path}")
    rows: List[dict] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_no, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables() or []
                page_count_before = len(rows)
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    header = [(c or "").strip().lower() for c in table[0]]
                    idx_no    = _find_col(header, ["drawing number", "dwg no",
                                                   "dwg. no", "no."])
                    idx_name  = _find_col(header, ["drawing name", "drawing title",
                                                   "title", "name"])
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

                added = len(rows) - page_count_before
                print(f"[PDF ] Page {page_no}: +{added} rows (total {len(rows)})")

    except FileNotFoundError:
        print(f"[ERROR] PDF not found: {pdf_path}")
        raise
    except Exception as exc:
        print(f"[ERROR] Failed to read PDF: {exc}")
        raise

    df = pd.DataFrame(rows, columns=["Drawing Number", "Drawing Name", "Scale"])
    if not df.empty:
        df = df.drop_duplicates(subset=["Drawing Number"]).reset_index(drop=True)
    print(f"[PDF ] Extracted {len(df)} unique drawings")
    return df


# ============================================================================
# 3.  DWG / DXF EXTRACTION  (Model Space only)
# ============================================================================
def _load_doc(path: Path):
    """Load a .dxf natively or a .dwg via the ODA File Converter addon."""
    suffix = path.suffix.lower()
    if suffix == ".dxf":
        return ezdxf.readfile(str(path))
    if suffix == ".dwg":
        from ezdxf.addons import odafc  # requires ODA File Converter on PATH
        return odafc.readfile(str(path))
    raise ValueError(f"Unsupported CAD file extension: {suffix}")


def _entity_point(ent) -> Optional[Tuple[float, float]]:
    """Return a representative (x, y) insertion point for TEXT / MTEXT."""
    t = ent.dxftype()
    try:
        if t == "TEXT":
            # Use align_point when the text is non-default aligned.
            if getattr(ent.dxf, "halign", 0) or getattr(ent.dxf, "valign", 0):
                p = ent.dxf.align_point
            else:
                p = ent.dxf.insert
            return (float(p[0]), float(p[1]))
        if t == "MTEXT":
            p = ent.dxf.insert
            return (float(p[0]), float(p[1]))
    except Exception:
        return None
    return None


def _entity_text(ent) -> str:
    """Return the plain text content of a TEXT / MTEXT entity."""
    t = ent.dxftype()
    try:
        if t == "TEXT":
            return (ent.dxf.text or "").strip()
        if t == "MTEXT":
            return ent.plain_text().strip()
    except Exception:
        return ""
    return ""


def _compute_insert_bbox(insert):
    """Return the bounding box of an INSERT or None if it cannot be computed."""
    try:
        box = bbox.extents([insert])
        if box.has_data:
            return box
    except Exception:
        pass
    return None


def _pair_number_and_name(
    hits: List[Tuple[float, float, str]],
) -> Tuple[str, str]:
    """
    Pick Drawing Number / Drawing Name out of the texts that fell inside
    the search rectangle. Strategy:

    1. A short alphanumeric token containing a digit -> Drawing Number.
    2. The remaining (longer descriptive) text -> Drawing Name.
    3. Fallback: if the regex yields nothing, sort by Y descending and
       use the two topmost strings.
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
        numbers.sort(key=lambda h: -h[1])  # topmost first
        others.sort(key=lambda h: -h[1])
        return numbers[0][2], others[0][2]

    hits_sorted = sorted(hits, key=lambda h: -h[1])
    if len(hits_sorted) >= 2:
        return hits_sorted[0][2], hits_sorted[1][2]
    return hits_sorted[0][2], ""


def extract_dwg_data(target_dir: str, block_name: str) -> pd.DataFrame:
    """
    Walk *target_dir* and, for every INSERT whose block name matches
    *block_name* in the host Model Space, derive Drawing Number /
    Drawing Name from TEXT / MTEXT inside the ratio-based search window.
    """
    target_block = block_name.strip().lower()
    rows: List[dict] = []

    # 3a. Discover CAD files (prefer .dwg per spec, also accept .dxf)
    cad_files: List[str] = []
    for pat in ("*.dwg", "*.DWG", "*.dxf", "*.DXF"):
        cad_files.extend(glob.glob(os.path.join(target_dir, pat)))
    cad_files = sorted(set(cad_files))

    empty = pd.DataFrame(columns=["Source File", "Drawing Number", "Drawing Name"])
    if not cad_files:
        print(f"[WARN] No .dwg / .dxf files found in {target_dir}")
        return empty

    # 3b. chdir into the target directory so Xref relative paths resolve.
    prev_cwd = os.getcwd()
    try:
        os.chdir(target_dir)
        print(f"[CAD ] Working directory: {os.getcwd()}")
        print(f"[CAD ] Files discovered : {len(cad_files)}")

        for idx, full_path in enumerate(cad_files, 1):
            fname = os.path.basename(full_path)
            print(f"[CAD ] ({idx}/{len(cad_files)}) Processing {fname} ...",
                  end=" ", flush=True)

            # ---- open ------------------------------------------------------
            try:
                doc = _load_doc(Path(fname))
            except Exception as exc:
                print(f"FAILED to open ({exc})")
                continue

            # ---- scan model space ------------------------------------------
            try:
                msp = doc.modelspace()

                tb_inserts = [
                    ins for ins in msp.query("INSERT")
                    if ins.dxf.name.strip().lower() == target_block
                ]
                if not tb_inserts:
                    print(f"no '{block_name}' in model space")
                    continue

                text_entities = list(msp.query("TEXT MTEXT"))
                file_rows = 0

                for tb in tb_inserts:
                    box = _compute_insert_bbox(tb)
                    if box is None:
                        # Likely an unresolved xref - cannot derive its extents.
                        continue

                    min_x = float(box.extmin.x)
                    min_y = float(box.extmin.y)
                    max_x = float(box.extmax.x)
                    max_y = float(box.extmax.y)
                    width  = max_x - min_x
                    height = max_y - min_y
                    if width <= 0 or height <= 0:
                        continue

                    # Search rectangle anchored at bottom-right (Max_X, Min_Y)
                    sx_max = max_x
                    sx_min = max_x - width  * X_RATIO
                    sy_min = min_y
                    sy_max = min_y + height * Y_RATIO

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
                        continue

                    dwg_no, dwg_name = _pair_number_and_name(hits)
                    if dwg_no or dwg_name:
                        rows.append({
                            "Source File":    fname,
                            "Drawing Number": dwg_no,
                            "Drawing Name":   dwg_name,
                        })
                        file_rows += 1

                print(f"Done ({file_rows} drawings, "
                      f"{len(tb_inserts)} title-block instances)")

            except Exception as exc:
                print(f"FAILED (unexpected error: {exc})")
                traceback.print_exc()

    finally:
        os.chdir(prev_cwd)

    df = pd.DataFrame(rows, columns=["Source File", "Drawing Number", "Drawing Name"])
    if not df.empty:
        df = (
            df[df["Drawing Number"] != ""]
            .drop_duplicates(subset=["Drawing Number"])
            .reset_index(drop=True)
        )
    print(f"[CAD ] Total unique drawings extracted: {len(df)}")
    return df


# ============================================================================
# 4.  COMPARE & WRITE EXCEL REPORT
# ============================================================================
def build_report(pdf_df: pd.DataFrame, dwg_df: pd.DataFrame, out_path: str) -> None:
    """Outer-merge the two datasets and write a highlighted report.xlsx."""
    pdf = pdf_df.rename(columns={
        "Drawing Name": "Drawing Name (PDF)",
        "Scale":        "Scale (PDF)",
    })
    dwg = dwg_df.rename(columns={
        "Drawing Name": "Drawing Name (DWG)",
    })

    merged = pdf.merge(dwg, on="Drawing Number", how="outer", indicator=True)

    # Guarantee columns exist even when one side is empty.
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

    # ---- highlight mismatches with openpyxl --------------------------------
    red = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
    wb  = load_workbook(out_path)
    ws  = wb.active

    header_row   = {cell.value: cell.column for cell in ws[1]}
    col_no       = header_row["Drawing Number"]
    col_name_pdf = header_row["Drawing Name (PDF)"]
    col_name_dwg = header_row["Drawing Name (DWG)"]
    col_scale    = header_row["Scale (PDF)"]
    col_status   = header_row["Match Status"]

    for row in range(2, ws.max_row + 1):
        status   = ws.cell(row=row, column=col_status).value or ""
        name_pdf = (ws.cell(row=row, column=col_name_pdf).value or "").strip()
        name_dwg = (ws.cell(row=row, column=col_name_dwg).value or "").strip()
        scale    = (ws.cell(row=row, column=col_scale).value    or "").strip()

        # (a) drawing missing on one side -> paint whole row
        if status in ("DWG MISSING", "PDF MISSING"):
            for c in (col_no, col_name_pdf, col_name_dwg, col_scale, col_status):
                ws.cell(row=row, column=c).fill = red
            continue

        # (b) drawing name mismatch between PDF and DWG
        if name_pdf.lower() != name_dwg.lower():
            ws.cell(row=row, column=col_name_pdf).fill = red
            ws.cell(row=row, column=col_name_dwg).fill = red

        # (c) scale missing on a matched row
        if scale == "":
            ws.cell(row=row, column=col_scale).fill = red

    wb.save(out_path)
    print(f"[XLSX] Report saved: {out_path}")


# ============================================================================
# 5.  MAIN
# ============================================================================
def main() -> None:
    target_dir, pdf_path, block_name = prompt_inputs()

    print("-" * 72)
    print(f"[INFO] Target dir : {target_dir}")
    print(f"[INFO] PDF path   : {pdf_path}")
    print(f"[INFO] Block name : {block_name}")
    print("-" * 72)

    # Resolve the output path now, before we chdir into target_dir.
    out_path = os.path.abspath(REPORT_NAME)

    try:
        pdf_df = extract_pdf_table(pdf_path)
    except Exception as exc:
        print(f"[FATAL] PDF extraction aborted: {exc}")
        sys.exit(1)

    try:
        dwg_df = extract_dwg_data(target_dir, block_name)
    except Exception as exc:
        print(f"[FATAL] DWG extraction aborted: {exc}")
        traceback.print_exc()
        sys.exit(1)

    if pdf_df.empty and dwg_df.empty:
        print("[ERROR] Both datasets are empty. Nothing to compare.")
        sys.exit(1)

    try:
        build_report(pdf_df, dwg_df, out_path)
    except Exception as exc:
        print(f"[FATAL] Report generation failed: {exc}")
        traceback.print_exc()
        sys.exit(1)

    print("-" * 72)
    print("[DONE] All tasks completed successfully.")


if __name__ == "__main__":
    main()
