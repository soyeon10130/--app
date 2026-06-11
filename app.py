import streamlit as st
import pandas as pd
import io
import datetime
import re
from datetime import timedelta, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KST_OFFSET = timedelta(hours=9)
SET1_DEST  = {'DAD','BKK','HKG','NRT'}
SET2_DEST  = {'DAC','LAX','EWR','SFO'}
SET3P_DEST = {'HNL'}
INSTR_DC   = {'LIP','LCP','DLCP','I','I*'}
INSTR_EXCL = {'강용학','김문배','박충근','박형득','서세규'}  # 교관수당 제외 대상

# DH 탑승자는 Roster 파일에서 자동 추출 (parse_roster_dh_exclude 함수 참조)

# ═══════════════════════════════════════════
# 공통 유틸
# ═══════════════════════════════════════════
def fmt_hhmm(h):
    if h is None or (isinstance(h, float) and pd.isna(h)) or h == 0:
        return "-"
    neg = h < 0; h = abs(h)
    hours = int(h); mins = round((h - hours) * 60)
    if mins == 60: hours += 1; mins = 0
    return f"{'-' if neg else ''}{hours:02d}:{mins:02d}"

def fmt_time(t):
    if t is None or (isinstance(t, float) and pd.isna(t)): return "-"
    if isinstance(t, (time, datetime.datetime)): return t.strftime("%H:%M")
    return str(t)

def parse_hhmm(val):
    if pd.isna(val): return 0.0
    val = str(val).strip()
    if ":" in val:
        p = val.split(":")
        try: return int(p[0]) + int(p[1]) / 60
        except: return 0.0
    try: return float(val)
    except: return 0.0

def td_to_hours(val):
    if isinstance(val, datetime.timedelta): return val.total_seconds() / 3600
    return 0.0

def make_border(color="BBBBBB"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def style_hdr(ws, row, headers, bg="1F4E79", height=20):
    bd = make_border()
    hf = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hb = PatternFill("solid", fgColor=bg)
    ct = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font=hf; c.fill=hb; c.alignment=ct; c.border=bd
    ws.row_dimensions[row].height = height

# ═══════════════════════════════════════════
# Roster에서 DH 탑승자 자동 추출
# ═══════════════════════════════════════════
def parse_roster_dh_exclude(uploaded):
    """
    Roster 파일을 읽어 DC 컬럼이 'DH'인 행의
    승무원 이름 + 날짜(Date 원본) + 편명(숫자부분)을 추출.
    반환: dict { (date_str, flight_num_str): {이름, ...}, ... }
    """
    df = pd.read_excel(uploaded)
    raw = df.copy()
    name_indices = raw[raw.iloc[:,0].astype(str).str.match(r'^[가-힣]{2,5}:$', na=False)].index.tolist()
    name_indices.append(len(raw))

    dh_exclude = {}  # (date_str, flight_no) -> set of names

    for idx_i, name_idx in enumerate(name_indices[:-1]):
        crew_name = str(raw.iloc[name_idx, 0]).replace(":", "").strip()
        next_idx = name_indices[idx_i + 1]
        hdr_rows = raw.iloc[name_idx:next_idx][raw.iloc[name_idx:next_idx, 0] == "Date"].index
        if len(hdr_rows) == 0:
            continue
        hdr_idx = hdr_rows[0]
        data = raw.iloc[hdr_idx+1:next_idx].copy()
        data.columns = ["Date","Pairing","DC","CI_L","CO_L","Activity",
                        "From","Start_L","To","Finish_L","AC_Hotel","BH","FDP","Blhr"]
        data = data.reset_index(drop=True)
        data["Date_ff"] = data["Date"].ffill()
        data["Pairing_ff"] = data["Pairing"].ffill()

        # DC가 DH인 행이 포함된 그룹에서 YP 실제비행 편명 추출
        data["group_id"] = data["Pairing"].notna().cumsum()

        for gid, grp in data.groupby("group_id"):
            if not grp["DC"].astype(str).str.strip().str.upper().eq("DH").any():
                continue
            # 실제 비행편(YP로 시작)의 날짜+편명 수집
