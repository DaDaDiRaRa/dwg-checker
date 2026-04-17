"""
app.py  —  DWG 자동 검토기 v_1.4 (LISP 연동형 + 스마트 매칭 + 회전 감지)
========================================================================
[V1.4 주요 업데이트]
1. 회전 감지 레이더(Rotation Matrix) 탑재: 세로로 90도 회전된 도곽이라도
   글자 좌표를 역회전(Un-rotate) 시켜 정확히 ROI 구역을 찾아냅니다.
========================================================================
"""

from __future__ import annotations
import glob, os, re, sys, webbrowser, json, math
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from typing import List, Optional, Tuple
import concurrent.futures

import pandas as pd
import ezdxf
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 리포트 파일 기본 이름
리포트_이름: str = "도면검토리포트_최종.xlsx"
ODA_DOWNLOAD_URL = "https://www.opendesign.com/guestfiles/oda_file_converter"

# ============================================================================
# 0. 설정 로드 및 엔진 세팅
# ============================================================================
def load_roi_config(block_name: str) -> Optional[dict]:
    config_dir = os.path.join(os.environ.get('APPDATA', ''), 'AutoDWG_Checker')
    config_path = os.path.join(config_dir, f"{block_name}.json")
    
    if os.path.exists(config_path):
        for enc in ['cp949', 'utf-8', 'euc-kr']:
            try:
                with open(config_path, 'r', encoding=enc) as f:
                    return json.load(f)
            except Exception:
                continue
    return None

def _oda_환경_설정() -> str:
    found_path = ""
    for 경로 in [r"C:\Program Files\ODA", r"C:\Program Files (x86)\ODA"]:
        실행파일들 = glob.glob(os.path.join(경로, "**", "ODAFileConverter.exe"), recursive=True)
        if 실행파일들:
            found_path = sorted(실행파일들, reverse=True)[0]
            break
    if found_path:
        폴더경로 = os.path.dirname(found_path)
        if 폴더경로 not in os.environ.get("PATH", ""):
            os.environ["PATH"] = 폴더경로 + os.pathsep + os.environ.get("PATH", "")
        try:
            ezdxf.options.odafc_win_exec_path = found_path
        except AttributeError:
            if not ezdxf.options.has_section('odafc'):
                ezdxf.options._config.add_section('odafc')
            ezdxf.options.set('odafc', 'win_exec_path', found_path)
    return found_path

def check_oda_installation():
    if not _oda_환경_설정():
        root = tk.Tk(); root.withdraw()
        msg = (
            "⚠️ CAD 분석 엔진(ODA)이 설치되어 있지 않습니다!\n\n"
            "확인을 누르면 다운로드 페이지가 열립니다. 기본 경로에 설치해 주세요."
        )
        if messagebox.askokcancel("엔진 설치 안내", msg):
            webbrowser.open(ODA_DOWNLOAD_URL)
        sys.exit()

# ============================================================================
# 1. 공통 유틸리티
# ============================================================================
_도면번호_패턴 = re.compile(r"([A-Z\u0391-\u03A9\.가-힣]{1,4})[-_ ]*(\d{1,5}[A-Z]*|TOE)")
_축척_패턴 = re.compile(r"(1\s?[/:,]\s?(\d{1,4})|NONE|N/A)", re.I)

def _도면번호_세척(raw_s: str) -> str:
    if not raw_s: return ""
    s = raw_s.strip().upper().replace("Λ", "A").replace("Δ", "A").replace("TOE", "108")
    if s.startswith("."): s = "AA" + s[1:]
    return re.sub(r"\s+", " ", s)

def _축척_텍스트_정리(txt: str) -> str:
    if not txt: return "X"
    u = txt.upper()
    if "NONE" in u or "N/A" in u: return "NONE"
    m = _축척_패턴.search(u)
    return f"1/{m.group(2)}" if m and m.group(2) else "X"

def _extract_drawing_number(text: str) -> Optional[str]:
    for m in _도면번호_패턴.finditer(text):
        prefix = m.group(1)
        if prefix.endswith("도") or prefix.endswith("표") or prefix.endswith("층") or prefix.endswith("동"): continue
        if any(k in prefix for k in ["상세", "일람", "배치", "전개", "마감", "계획", "조감", "구조", "코어", "지하", "옥상", "옥탑", "지붕", "주동", "단위", "세대"]): continue
        return m.group(0)
    return None

def _정리문자열(txt: str) -> str:
    return re.sub(r"\s+", " ", (txt or "")).strip()

def _is_none_scale(scale_txt: str) -> bool:
    return scale_txt == "NONE"

def _cad_로드(path: Path):
    if path.suffix.lower() == ".dxf": return ezdxf.readfile(str(path))
    _oda_환경_설정()
    from ezdxf.addons import odafc
    return odafc.readfile(str(path))

def _텍스트_데이터_추출(ent) -> List[Tuple[float, float, str, float]]:
    유형 = ent.dxftype()
    결과 = []
    try:
        if 유형 == "TEXT":
            p = ent.dxf.align_point if getattr(ent.dxf, "halign", 0) or getattr(ent.dxf, "valign", 0) else ent.dxf.insert
            h = getattr(ent.dxf, "height", 10.0)
            txt = (ent.dxf.text or "").strip()
            if txt: 결과.append((float(p[0]), float(p[1]), txt, float(h)))
        elif 유형 == "MTEXT":
            h = getattr(ent.dxf, "char_height", 10.0)
            bx, by = float(ent.dxf.insert[0]), float(ent.dxf.insert[1])
            lines = ent.plain_text().split('\n')
            for i, line in enumerate(lines):
                txt = line.strip()
                if txt: 결과.append((bx, by - (i * h * 1.5), txt, float(h)))
        elif 유형 == "ATTRIB":
            p = ent.dxf.insert
            h = getattr(ent.dxf, "height", 10.0)
            txt = (ent.dxf.text or "").strip()
            if txt: 결과.append((float(p[0]), float(p[1]), txt, float(h)))
    except Exception: pass
    return 결과

# ============================================================================
# 2. 목록표 (DWG) 파싱 로직 (회전 감지 적용)
# ============================================================================
def _collect_layout_texts(layout) -> List[Tuple[float, float, str, float]]:
    texts = []
    try:
        for ent in layout.query("TEXT MTEXT"): texts.extend(_텍스트_데이터_추출(ent))
        for ins in layout.query("INSERT"):
            for att in getattr(ins, "attribs", []): texts.extend(_텍스트_데이터_추출(att))
    except Exception: pass
    seen, out = set(), []
    for x, y, txt, h in texts:
        clean = _정리문자열(txt)
        key = (round(x, 3), round(y, 3), clean)
        if key not in seen:
            seen.add(key); out.append((float(x), float(y), clean, float(h)))
    return out

def _split_lines_from_cell_texts(cell_texts: List[Tuple[float, float, str, float]], row_h: float) -> List[str]:
    if not cell_texts: return []
    cell_texts = sorted(cell_texts, key=lambda t: (-t[1], t[0]))
    y_tol = max(row_h * 0.015, 1.0)
    lines, current, current_y = [], [], None
    for x, y, txt, _ in cell_texts:
        if current_y is None: current_y = y; current.append((x, txt)); continue
        if abs(current_y - y) <= y_tol: current.append((x, txt))
        else:
            current.sort(key=lambda v: v[0]); lines.append(current)
            current_y = y; current = [(x, txt)]
    if current: current.sort(key=lambda v: v[0]); lines.append(current)
    return [" ".join([txt for _, txt in line]) for line in lines]

def _extract_list_scales_from_cell(cell_texts: List[Tuple[float, float, str, float]], row_h: float) -> Tuple[str, str]:
    if not cell_texts: return "X", "X"
    label_a1 = [t for t in cell_texts if re.search(r"\bA1\b", t[2].upper())]
    label_a3 = [t for t in cell_texts if re.search(r"\bA3\b", t[2].upper())]
    scale_items = []
    for x, y, txt, _ in cell_texts:
        normalized = _축척_텍스트_정리(txt)
        if normalized != "X": scale_items.append((x, y, normalized))
    if not scale_items: return "X", "X"
    
    numeric_scales = [(x, y, s) for x, y, s in scale_items if not _is_none_scale(s)]
    none_scales = [(x, y, s) for x, y, s in scale_items if _is_none_scale(s)]
    chosen_a1, chosen_a3, used, y_tol = "X", "X", set(), max(row_h * 0.35, 2.0)

    def pick_nearest(label_items, candidates):
        if not label_items or not candidates: return None
        label_y = label_items[0][1]
        ordered = sorted(candidates, key=lambda c: (abs(c[1] - label_y), 1 if _is_none_scale(c[2]) else 0, c[0]))
        return ordered[0] if abs(ordered[0][1] - label_y) <= y_tol else None

    picked = pick_nearest(label_a1, numeric_scales if numeric_scales else scale_items)
    if picked: chosen_a1 = picked[2]; used.add((picked[0], picked[1], picked[2]))
    picked = pick_nearest(label_a3, [c for c in (numeric_scales if numeric_scales else scale_items) if (c[0], c[1], c[2]) not in used])
    if picked: chosen_a3 = picked[2]; used.add((picked[0], picked[1], picked[2]))
    
    ordered_values = [s[2] for s in sorted(scale_items, key=lambda v: (-v[1], v[0])) if (s[0], s[1], s[2]) not in used]
    if chosen_a1 == "X" and ordered_values: chosen_a1 = ordered_values.pop(0)
    if chosen_a3 == "X" and ordered_values: chosen_a3 = ordered_values.pop(0)
    return chosen_a1, chosen_a3

def _extract_number_and_title_from_lines(lines: List[str]) -> Tuple[str, str]:
    번호, 명칭후보 = "", []
    for line in lines:
        clean = line
        raw_no = _extract_drawing_number(clean)
        if raw_no:
            cleaned_no = _도면번호_세척(raw_no)
            if not 번호: 번호 = cleaned_no
            clean = clean.replace(raw_no, " ")
        clean = re.sub(r"\bA1\b|\bA3\b|NONE|N/A|1\s?[/:,]\s?\d{1,4}", " ", clean, flags=re.I)
        clean = re.sub(r"\s+", " ", clean).strip(" ,")
        if len(clean) > 1 and not any(bw in clean.upper() for bw in ["도면명", "SCALE", "SUBJECT", "TITLE", "축 척", "축척"]): 명칭후보.append(clean)
    명칭 = re.sub(r"\s+", " ", " ".join(명칭후보)).strip()
    return 번호, 명칭

def extract_dwg_list_table(dwg_path: str, block_name: str, base_w: float, base_h: float) -> pd.DataFrame:
    print(f"\n[LIST] DWG 도면목록표 분석 시작: {os.path.basename(dwg_path)}")
    데이터, 목표블록 = [], block_name.strip().lower()
    try:
        doc = _cad_로드(Path(dwg_path))
        for layout in doc.layouts:
            도곽들 = [ins for ins in layout.query("INSERT") if 목표블록 in ins.dxf.name.lower()]
            if not 도곽들: continue
            모든텍스트 = _collect_layout_texts(layout)
            for 도곽 in 도곽들:
                ix, iy = float(도곽.dxf.insert.x), float(도곽.dxf.insert.y)
                xscale, yscale = abs(float(도곽.dxf.xscale)), abs(float(도곽.dxf.yscale))
                너비, 높이 = base_w * xscale, base_h * yscale
                
                # [회전 레이더] 회전 각도를 가져와서 역회전용 삼각함수 세팅
                rot_deg = getattr(도곽.dxf, 'rotation', 0.0)
                rad = math.radians(-rot_deg) # -를 붙여 반대로 돌림
                cos_val, sin_val = math.cos(rad), math.sin(rad)

                col_ranges = [(ix + (너비 * 0.05758), ix + (너비 * 0.28946)), (ix + (너비 * 0.47970), ix + (너비 * 0.71159))]
                y_min, y_max = iy + (높이 * 0.05235), iy + (높이 * 0.92600)
                
                for min_x, max_x in col_ranges:
                    구역_텍스트 = []
                    for t in 모든텍스트:
                        tx, ty, txt, th = t
                        # 텍스트 좌표를 삽입점 기준으로 역회전(Un-rotate)시킵니다.
                        dx, dy = tx - ix, ty - iy
                        unrot_x = ix + (dx * cos_val - dy * sin_val)
                        unrot_y = iy + (dx * sin_val + dy * cos_val)
                        
                        if min_x <= unrot_x <= max_x and y_min <= unrot_y <= y_max:
                            # 정렬을 위해 역회전된 좌표를 저장합니다.
                            구역_텍스트.append((unrot_x, unrot_y, txt, th))
                            
                    if not 구역_텍스트: continue
                    구역_텍스트.sort(key=lambda x: -x[1]) # Y축 위에서 아래로
                    
                    줄목록, 현재_줄, 현재_y, y_tol = [], [], None, 높이 * 0.012
                    for t in 구역_텍스트:
                        if 현재_y is None or abs(현재_y - t[1]) <= y_tol: 현재_y = t[1]; 현재_줄.append(t)
                        else: 줄목록.append(현재_줄); 현재_y = t[1]; 현재_줄 = [t]
                    if 현재_줄: 줄목록.append(현재_줄)
                    for row_texts in 줄목록:
                        lines = _split_lines_from_cell_texts(row_texts, y_tol * 2)
                        번호, 명칭 = _extract_number_and_title_from_lines(lines)
                        if not 번호: continue
                        a1, a3 = _extract_list_scales_from_cell(row_texts, y_tol * 2)
                        데이터.append({"도면번호(LIST)": 번호, "도면명(LIST)": 명칭, "축척_A1(LIST)": a1, "축척_A3(LIST)": a3})
    except Exception as e: print(f"[ERROR] 목록표 분석 중 오류: {e}")
    
    df = pd.DataFrame(데이터)
    if df.empty: return pd.DataFrame(columns=["도면번호(LIST)", "도면명(LIST)", "축척_A1(LIST)", "축척_A3(LIST)"])
    return df.drop_duplicates(subset=["도면번호(LIST)"]).reset_index(drop=True)

# ============================================================================
# 3. 개별 도면 (DWG) 파싱 (LISP ROI + 회전 감지 적용)
# ============================================================================
def _process_single_dwg(args: Tuple[str, str, dict, float, float]) -> Tuple[List[dict], str]:
    전체경로, 목표블록, roi_cfg, base_w, base_h = args
    파일명, 데이터, 에러메시지 = os.path.basename(전체경로), [], ""
    try:
        doc = _cad_로드(Path(전체경로))
        도곽_발견됨 = False

        for layout in doc.layouts:
            도곽들 = [ins for ins in layout.query("INSERT") if 목표블록 in ins.dxf.name.lower()]
            if not 도곽들: continue
            
            도곽_발견됨 = True
            모든텍스트_raw = []
            for ent in layout.query("TEXT MTEXT"): 모든텍스트_raw.extend(_텍스트_데이터_추출(ent))
            for ins in layout.query("INSERT"):
                for att in getattr(ins, "attribs", []): 모든텍스트_raw.extend(_텍스트_데이터_추출(att))
            
            seen, 모든텍스트 = set(), []
            for x, y, txt, h in 모든텍스트_raw:
                key = (round(x, 3), round(y, 3), _정리문자열(txt))
                if key not in seen: seen.add(key); 모든텍스트.append((x, y, txt, h))

            for 도곽 in 도곽들:
                ix, iy = float(도곽.dxf.insert.x), float(도곽.dxf.insert.y)
                xscale, yscale = abs(float(도곽.dxf.xscale)), abs(float(도곽.dxf.yscale))
                너비, 높이 = base_w * xscale, base_h * yscale

                # [회전 레이더] 개별 도곽의 회전을 감지하여 역회전 행렬 준비
                rot_deg = getattr(도곽.dxf, 'rotation', 0.0)
                rad = math.radians(-rot_deg)
                cos_val, sin_val = math.cos(rad), math.sin(rad)

                def get_txt_in_roi(roi):
                    x_min, x_max = ix + (너비 * roi[0]), ix + (너비 * roi[1])
                    y_min, y_max = iy + (높이 * roi[2]), iy + (높이 * roi[3])
                    
                    박스내글자 = []
                    for t in 모든텍스트:
                        tx, ty, txt, th = t
                        # 텍스트 좌표를 삽입점 기준으로 역회전(Un-rotate)시킵니다.
                        dx, dy = tx - ix, ty - iy
                        unrot_x = ix + (dx * cos_val - dy * sin_val)
                        unrot_y = iy + (dx * sin_val + dy * cos_val)
                        
                        if x_min <= unrot_x <= x_max and y_min <= unrot_y <= y_max:
                            # 정렬을 위해 역회전된 X, Y 위치를 넣습니다.
                            박스내글자.append((unrot_x, unrot_y, txt))
                            
                    # 역회전된 좌표계에서 Y(위에서 아래), X(좌에서 우)로 정렬
                    박스내글자.sort(key=lambda t: (-t[1], t[0]))
                    return " ".join([t[2] for t in 박스내글자])

                t_str = get_txt_in_roi(roi_cfg['title_roi'])
                n_str = get_txt_in_roi(roi_cfg['num_roi'])
                s_str = get_txt_in_roi(roi_cfg['scale_roi'])

                번호_후보 = _extract_drawing_number(n_str)
                번호 = _도면번호_세척(번호_후보) if 번호_후보 else ""
                
                명칭 = t_str
                if 번호_후보:
                    명칭 = 명칭.replace(번호_후보, "")
                    
                명칭 = re.sub(r"\bA1\b|\bA3\b|NONE|N/A", "", 명칭, flags=re.IGNORECASE)
                명칭 = re.sub(r"1\s?[/:,]\s?\d{1,4}", "", 명칭, flags=re.IGNORECASE).strip(" ,")
                명칭 = re.sub(r"\s+", " ", 명칭)
                
                matches = list(_축척_패턴.finditer(s_str.upper()))
                a1, a3 = "X", "X"
                
                if matches:
                    val1 = matches[0].group(2)
                    a1 = f"1/{val1}" if val1 else "NONE"
                    if len(matches) >= 2:
                        val2 = matches[1].group(2)
                        a3 = f"1/{val2}" if val2 else "NONE"
                
                if a1 == "X" and ("NONE" in s_str.upper() or "N/A" in s_str.upper()):
                    a1, a3 = "NONE", "NONE"

                if 번호: 
                    데이터.append({
                        "파일명": 파일명, 
                        "도면번호(DWG)": 번호, 
                        "도면명(DWG)": 명칭.strip(), 
                        "축척_A1(DWG)": a1, 
                        "축척_A3(DWG)": a3
                    })
        del doc
        if not 도곽_발견됨: return 데이터, "도곽 블록 없음"
    except Exception as e: 에러메시지 = str(e)
    return 데이터, 에러메시지

def extract_dwg_data_multiprocess(target_dirs: List[str], block_name: str, roi_cfg: dict, base_w: float, base_h: float) -> pd.DataFrame:
    모든_캐드파일 = []
    for d in target_dirs:
        폴더 = Path(d)
        if 폴더.exists(): 모든_캐드파일.extend([str(p) for p in 폴더.iterdir() if p.is_file() and p.suffix.lower() in [".dwg", ".dxf"]])
    캐드파일들 = sorted(list(set(모든_캐드파일)))
    if not 캐드파일들:
        print("[CAD ] 폴더 내에 처리할 도면 파일이 없습니다."); return pd.DataFrame(columns=["파일명", "도면번호(DWG)", "도면명(DWG)", "축척_A1(DWG)", "축척_A3(DWG)"])

    print(f"\n[CAD ] 총 {len(캐드파일들)}개의 도면 분석 중... (터보 모드 가동 🚀)")
    최종_데이터 = []
    with concurrent.futures.ProcessPoolExecutor() as executor:
        futures = {executor.submit(_process_single_dwg, (path, block_name.strip().lower(), roi_cfg, base_w, base_h)): path for path in 캐드파일들}
        for i, future in enumerate(concurrent.futures.as_completed(futures), 1):
            경로 = futures[future]
            try:
                결과, 에러 = future.result()
                if 결과: 최종_데이터.extend(결과)
                print(f"   [{i}/{len(캐드파일들)}] {'완료' if 결과 else '패스'}: {os.path.basename(경로)} ({에러 if 에러 else '성공'})")
            except Exception as e:
                print(f"   [{i}/{len(캐드파일들)}] 시스템 오류: {os.path.basename(경로)} ({e})")
    
    if not 최종_데이터:
        return pd.DataFrame(columns=["파일명", "도면번호(DWG)", "도면명(DWG)", "축척_A1(DWG)", "축척_A3(DWG)"])
    return pd.DataFrame(최종_데이터)

# ============================================================================
# 4. 리포트 생성
# ============================================================================
def build_report(list_df: pd.DataFrame, dwg_df: pd.DataFrame, out_path: str):
    if list_df.empty and dwg_df.empty:
        print("[알림] 추출된 데이터가 없어 엑셀 리포트를 생성하지 않습니다.")
        return

    lst, dwg = list_df.copy(), dwg_df.copy()
    
    if "도면번호(LIST)" not in lst.columns: lst["도면번호(LIST)"] = ""
    if "도면번호(DWG)" not in dwg.columns: dwg["도면번호(DWG)"] = ""

    lst["KEY"] = lst["도면번호(LIST)"].astype(str).str.upper().str.replace(r"[\s\-_]", "", regex=True)
    dwg["KEY"] = dwg["도면번호(DWG)"].astype(str).str.upper().str.replace(r"[\s\-_]", "", regex=True)
    
    결과 = pd.merge(lst, dwg, on="KEY", how="outer", indicator=True)
    결과["상태"] = 결과["_merge"].map({"both": "일치", "left_only": "DWG 누락", "right_only": "목록표 누락"})
    
    cols = ["도면번호(LIST)", "도면명(LIST)", "축척_A1(LIST)", "축척_A3(LIST)", "도면번호(DWG)", "도면명(DWG)", "축척_A1(DWG)", "축척_A3(DWG)", "파일명", "상태"]
    for c in cols: 
        if c not in 결과.columns: 결과[c] = ""
    
    결과[cols].fillna("X").to_excel(out_path, index=False)
    
    wb = load_workbook(out_path); ws = wb.active
    빨간색 = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
    h = {cell.value: cell.column for cell in ws[1]}
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, h["상태"]).value != "일치":
            for c in range(1, len(cols)+1): ws.cell(row, c).fill = 빨간색
        else:
            val_list = re.sub(r"[\s\-_]", "", str(ws.cell(row, h.get("도면번호(LIST)")).value).upper())
            val_dwg = re.sub(r"[\s\-_]", "", str(ws.cell(row, h.get("도면번호(DWG)")).value).upper())
            
            if val_list != val_dwg:
                ws.cell(row, h.get("도면번호(LIST)")).fill = 빨간색
                ws.cell(row, h.get("도면번호(DWG)")).fill = 빨간색
            for s in ["A1", "A3"]:
                p_v = str(ws.cell(row, h[f"축척_{s}(LIST)"]).value).replace(" ","")
                d_v = str(ws.cell(row, h[f"축척_{s}(DWG)"]).value).replace(" ","")
                if p_v != d_v:
                    ws.cell(row, h[f"축척_{s}(LIST)"]).fill = 빨간색; ws.cell(row, h[f"축척_{s}(DWG)"]).fill = 빨간색
    wb.save(out_path)
    print(f"\n[XLSX] 리포트 저장 완료: {out_path}")

# ============================================================================
# 5. 메인 함수 
# ============================================================================
def main():
    print("=" * 72)
    print(" AutoDWG Cross-Checker v_1.4 (LISP Connected - Magic Rotation Radar)")
    print("=" * 72)

    check_oda_installation()

    blk_name = input("1. 도곽 블록 이름을 입력하세요: ").strip()
    roi_config = load_roi_config(blk_name)

    if not roi_config:
        print(f"\n[오류] '{blk_name}'에 대한 구역 설정 파일이 없습니다!")
        
        config_dir = os.path.join(os.environ.get('APPDATA', ''), 'AutoDWG_Checker')
        if os.path.exists(config_dir):
            saved_files = [f.replace('.json', '') for f in os.listdir(config_dir) if f.endswith('.json')]
            if saved_files:
                print("\n💡 [참고] 현재 저장되어 있는 도곽 목록은 다음과 같습니다:")
                for sf in saved_files: print(f"   - {sf}")
                print("\n(이름의 띄어쓰기, 대소문자, 괄호 모양이 정확히 일치해야 합니다!)")
        
        print("\n캐드에서 SET_ROI 명령어로 구역을 먼저 지정해 주세요.")
        input("\n엔터를 누르면 종료됩니다...")
        return

    base_w = float(roi_config.get('base_w', 841.0))
    base_h = float(roi_config.get('base_h', 594.0))

    print(f"\n[성공] '{blk_name}' 설정을 로드했습니다. (원본크기: {base_w}x{base_h})")

    print("\n2. 도면 목록표 DWG 파일의 경로를 입력하세요.")
    목록표_경로 = input("   경로: ").strip().strip('"')
    if not os.path.isfile(목록표_경로):
        print("[ERROR] 유효한 파일이 아닙니다. 종료합니다."); return

    dwg_dirs = []
    print("\n3. 검토할 개별 도면(DWG) 폴더 경로를 입력하세요. (여러 개 입력 가능)")
    print("   더 이상 없다면 'N'을 입력하세요.")
    while True:
        path_input = input(f"   [{len(dwg_dirs) + 1}번째 폴더] (또는 N): ").strip().strip('"')
        if path_input.upper() == 'N': break
        if os.path.isdir(path_input): dwg_dirs.append(path_input)

    print("-" * 72)

    if getattr(sys, 'frozen', False): 실행폴더 = os.path.dirname(sys.executable)
    else: 실행폴더 = os.path.dirname(os.path.abspath(__file__))
    최종_저장경로 = os.path.join(실행폴더, 리포트_이름)

    try:
        list_데이터 = extract_dwg_list_table(목록표_경로, blk_name, base_w, base_h)
        dwg_데이터 = extract_dwg_data_multiprocess(dwg_dirs, blk_name, roi_config, base_w, base_h)

        build_report(list_데이터, dwg_데이터, 최종_저장경로)
        
        print("-" * 72)
        print(f"[DONE] 검토 완료! 리포트가 프로그램과 같은 폴더에 저장되었습니다.")
        
    except PermissionError:
        print("\n[ERROR] 엑셀 파일이 이미 켜져 있습니다. 창을 닫고 다시 실행해 주세요.")
    except Exception as e:
        print(f"\n[ERROR] 알 수 없는 오류 발생: {e}")

    print("=" * 72)
    input("\n[안내] 모든 작업이 끝났습니다. 엔터키를 누르면 창이 닫힙니다...")


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()