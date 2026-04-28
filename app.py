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
# 파일1: DHC / DAYOFF 파싱
# ═══════════════════════════════════════════
def parse_dhc_file(uploaded):
    df = pd.read_excel(uploaded)
    data = df.iloc[4:].copy()
    data.columns = ["Crew Code","Position","Name","Counter","Mar26","Total"]
    data = data[data["Counter"] != "Counter"].copy()
    block  = data[data["Counter"]=="Block"][["Crew Code","Mar26"]].copy()
    dayoff = data[data["Counter"]=="dayoff"][["Crew Code","Mar26"]].copy()
    dhc    = data[data["Counter"]=="DHC"][["Crew Code","Mar26"]].copy()
    block.columns=["Crew Code","Block"]; block["Block"]=block["Block"].apply(parse_hhmm)
    dayoff.columns=["Crew Code","Dayoff"]; dayoff["Dayoff"]=pd.to_numeric(dayoff["Dayoff"],errors="coerce").fillna(0).astype(int)
    dhc.columns=["Crew Code","DHC"]; dhc["DHC"]=dhc["DHC"].apply(parse_hhmm)
    return block.merge(dayoff, on="Crew Code").merge(dhc, on="Crew Code")

# ═══════════════════════════════════════════
# 파일2: OBCA / OBFO 파싱
# ═══════════════════════════════════════════
def parse_obca_file(uploaded):
    df = pd.read_excel(uploaded)
    data = df.iloc[6:].copy()
    data.columns=["Crew Code","AC Type","Position","Block Hours","Cruise Time","Sectors","Valid From","Valid To"]
    data = data[data["Crew Code"]!="Crew Code"].copy()
    data["Crew Code"] = data["Crew Code"].ffill()
    ob = data[data["Position"].isin(["OBFO","OBCA"])].copy()
    ob["OB_hrs"] = ob["Block Hours"].apply(td_to_hours)
    ob_sum = ob.groupby("Crew Code")["OB_hrs"].sum().reset_index()
    ob_sum.columns=["Crew Code","OBCA_OBFO"]
    return ob_sum

def calc_summary(base_df, ob_df):
    df = base_df.merge(ob_df, on="Crew Code", how="left")
    df["OBCA_OBFO"] = df["OBCA_OBFO"].fillna(0)
    df["DHC_50"]    = df["DHC"] * 0.5
    df["Total_Flt"] = (df["Block"] + df["DHC_50"] - df["OBCA_OBFO"]).clip(lower=0)
    df["Dayoff_Under8"] = df["Dayoff"] < 8
    return df

# ═══════════════════════════════════════════
# 파일3: Roster 교관 수당 파싱
# ═══════════════════════════════════════════
AIRPORT_UTC = {
    'ICN': 9, 'NRT': 9, 'HKG': 8, 'BKK': 7, 'DAD': 7,
    'LAX': -8, 'EWR': -5, 'SFO': -8, 'DAC': 6, 'HNL': -10,
}

def classify_route(from_val, to_val):
    """FROM 또는 TO 기준으로 Set 분류 (ICN 제외한 공항 기준)"""
    for v in [from_val, to_val]:
        v = str(v).strip().upper()
        if v == 'ICN': continue
        if v in SET1_DEST:  return "1set"
        if v in SET2_DEST:  return "2set"
        if v in SET3P_DEST: return "3P"
    return None

def local_to_kst(date_str, time_str, from_city):
    """현지 출발시간 → KST datetime (공항별 UTC 오프셋 적용)"""
    try:
        base = datetime.datetime.strptime(str(date_str).strip(), "%d%b%y")
    except:
        return None
    s = str(time_str).strip() if not pd.isna(time_str) else ""
    if not s or s == "nan":
        return datetime.datetime(base.year, base.month, base.day, 0, 0)
    next_day = "+1" in s
    s_clean = s.replace("+1", "").strip()
    try:
        h = int(s_clean[:2]); m = int(s_clean[2:4])
    except:
        return datetime.datetime(base.year, base.month, base.day, 0, 0)
    local_dt = datetime.datetime(base.year, base.month, base.day, h, m)
    if next_day: local_dt += timedelta(days=1)
    offset = AIRPORT_UTC.get(str(from_city).strip().upper(), 9)
    return local_dt + timedelta(hours=(9 - offset))

def calc_instr_hrs(blhr_h, set_type):
    if set_type == "1set":  return blhr_h
    if set_type == "2set":  return blhr_h / 2
    if set_type == "3P":    return blhr_h / 3
    return 0.0

def is_actual_flight(activity):
    """YP로 시작하는 실제 비행편 여부"""
    if pd.isna(activity): return False
    return str(activity).strip().upper().startswith("YP")

def parse_roster_file(uploaded, target_month=None, target_year=None):
    df = pd.read_excel(uploaded)
    raw = df.copy()
    name_indices = raw[raw.iloc[:,0].astype(str).str.match(r'^[가-힣]{2,5}:$', na=False)].index.tolist()
    name_indices.append(len(raw))

    detail_rows = []
    for idx_i, name_idx in enumerate(name_indices[:-1]):
        crew_name = str(raw.iloc[name_idx, 0]).replace(":", "").strip()
        if crew_name in INSTR_EXCL:
            continue
        next_idx = name_indices[idx_i + 1]
        hdr_rows = raw.iloc[name_idx:next_idx][raw.iloc[name_idx:next_idx, 0] == "Date"].index
        if len(hdr_rows) == 0: continue
        hdr_idx = hdr_rows[0]
        data = raw.iloc[hdr_idx+1:next_idx].copy()
        data.columns = ["Date","Pairing","DC","CI_L","CO_L","Activity",
                        "From","Start_L","To","Finish_L","AC_Hotel","BH","FDP","Blhr"]
        data = data.reset_index(drop=True)
        data["Pairing_ff"] = data["Pairing"].ffill()
        data["Date_ff"]    = data["Date"].ffill()

        # ★ raw Pairing 열에 값이 등장할 때마다 새 그룹 (운항 단위 정확 분리)
        data["group_id"] = data["Pairing"].notna().cumsum()
        first_valid = data["Pairing_ff"].first_valid_index()
        if first_valid is None: continue
        data = data[(data.index >= first_valid) & (data["group_id"] > 0)]

        for gid, grp in data.groupby("group_id"):
            if grp["Pairing_ff"].isna().all(): continue
            if not grp["DC"].isin(INSTR_DC).any(): continue

            flights = grp[grp["Activity"].apply(is_actual_flight)].copy()
            flights["Blhr_h"] = flights["Blhr"].apply(parse_hhmm)
            flights = flights[flights["Blhr_h"] > 0]

            for _, row in flights.iterrows():
                from_val = str(row["From"]).strip() if not pd.isna(row["From"]) else ""
                to_val   = str(row["To"]).strip()   if not pd.isna(row["To"])   else ""
                set_type = classify_route(from_val, to_val)
                if set_type is None: continue

                # ★ 공항별 UTC 오프셋으로 KST 출발 시각 계산 → 월 귀속 판단
                kst_dt = local_to_kst(row["Date_ff"], row["Start_L"], from_val)
                if kst_dt is None: continue
                if target_month and target_year:
                    if kst_dt.month != target_month or kst_dt.year != target_year:
                        continue

                blhr_h  = row["Blhr_h"]
                instr_h = calc_instr_hrs(blhr_h, set_type)
                dc_val  = str(row["DC"]).strip() if not pd.isna(row["DC"]) else "-"

                detail_rows.append({
                    "이름":       crew_name,
                    "날짜(KST)":  kst_dt.strftime("%m/%d"),
                    "DC":         dc_val,
                    "편명":       str(row["Activity"]).strip(),
                    "From":       from_val,
                    "To":         to_val,
                    "Set구분":    set_type,
                    "Blhr(원본)": fmt_hhmm(blhr_h),
                    "Blhr_h":     blhr_h,
                    "교관시간_h":  instr_h,
                    "교관시간":   fmt_hhmm(instr_h),
                })

    return pd.DataFrame(detail_rows)

# ═══════════════════════════════════════════
# FltReport 수당 계산
# ═══════════════════════════════════════════
def combine_dt(date_str, t):
    if pd.isna(t) or t is None: return None
    base = datetime.datetime.strptime(date_str, "%d%b%y")
    return datetime.datetime(base.year, base.month, base.day, t.hour, t.minute, t.second)

def calc_night(ci, co):
    if not ci or not co: return 0.0
    total=0.0; d=ci.date()
    for delta in [-1,0,1]:
        ns=datetime.datetime.combine(d+timedelta(days=delta), time(22,0))
        ne=datetime.datetime.combine(d+timedelta(days=delta+1), time(6,0))
        o1=max(ci,ns); o2=min(co,ne)
        if o2>o1: total+=(o2-o1).total_seconds()/3600
    return total

def calc_ot(atd, ata):
    if not atd or not ata: return 0.0
    return max(0.0, (ata-atd).total_seconds()/3600 - 8)

def blhrs_decimal(t):
    if pd.isna(t) or t is None: return 0.0
    if isinstance(t, time): return t.hour + t.minute/60 + t.second/3600
    return 0.0

def process_flt(df, target_month, target_year):
    sum_rows, det_rows = [], []
    for _, row in df.iterrows():
        crew_str = row["운항 Crew"]
        if pd.isna(crew_str): continue
        atd_utc = combine_dt(row["Date"], row["ATD"])
        if not atd_utc: continue
        atd_kst = atd_utc + KST_OFFSET
        ata_utc = combine_dt(row["Date"], row["ATA"])
        ata_kst = (ata_utc+KST_OFFSET) if ata_utc else None
        ci_utc  = combine_dt(row["Date"], row["운항 C/I"])
        co_utc  = combine_dt(row["Date"], row["운항 C/O"])
        ci_kst  = (ci_utc+KST_OFFSET) if ci_utc else None
        co_kst  = (co_utc+KST_OFFSET) if co_utc else None
        if ata_kst and ata_kst < atd_kst: ata_kst += timedelta(days=1)
        if co_kst and ci_kst and co_kst < ci_kst: co_kst += timedelta(days=1)
        if atd_kst.month!=target_month or atd_kst.year!=target_year: continue
        flight=str(row["Flight"]); bl=blhrs_decimal(row["Bl Hrs"])
        night=calc_night(ci_kst,co_kst); ot=calc_ot(atd_kst,ata_kst)
        p3=bl if flight in ["0151","0152"] else 0.0
        route=f"{row['From']}→{row['To']}"
        for name in str(crew_str).split():
            name=name.strip()
            if not name: continue
            sum_rows.append({"이름":name,"night":night,"ot":ot,"p3":p3})
            det_rows.append({"이름":name,"날짜":atd_kst.strftime("%m/%d"),"편명":flight,"구간":route,
                             "ATD(KST)":fmt_time(atd_kst),"ATA(KST)":fmt_time(ata_kst),
                             "CI(KST)":fmt_time(ci_kst),"CO(KST)":fmt_time(co_kst),
                             "야간(h)":night,"연장(h)":ot,"3P(h)":p3})
    if not sum_rows: return pd.DataFrame(), pd.DataFrame()
    s=pd.DataFrame(sum_rows)
    summary=(s.groupby("이름").agg(야간=("night","sum"),연장=("ot","sum"),P3=("p3","sum"))
             .reset_index().sort_values("이름").reset_index(drop=True))
    detail=pd.DataFrame(det_rows).sort_values(["이름","날짜","편명"]).reset_index(drop=True)
    return summary, detail

# ═══════════════════════════════════════════
# 엑셀 빌드
# ═══════════════════════════════════════════
def build_excel(flt_sum, flt_det, calc_df, instr_det, target_year, target_month):
    wb = Workbook()
    bd=make_border(); center=Alignment(horizontal="center",vertical="center")
    left=Alignment(horizontal="left",vertical="center")
    cf=Font(name="Arial",size=10); tf=Font(name="Arial",bold=True,size=10)
    tbg=PatternFill("solid",fgColor="FFF2CC")
    sbf=Font(name="Arial",bold=True,color="FFFFFF",size=10)
    sbbg=PatternFill("solid",fgColor="2E75B6")
    redf=Font(name="Arial",bold=True,size=10,color="C00000")
    redbg=PatternFill("solid",fgColor="FCEAEA")
    label=f"{target_year}년 {target_month}월"

    def title_row(ws, text, ncols, note=None):
        col_letter = get_column_letter(ncols)
        ws.merge_cells(f"A1:{col_letter}1")
        c=ws["A1"]; c.value=text
        c.font=Font(name="Arial",bold=True,size=13); c.alignment=center
        ws.row_dimensions[1].height=28
        if note:
            ws.merge_cells(f"A2:{col_letter}2")
            c2=ws["A2"]; c2.value=note
            c2.font=Font(name="Arial",italic=True,size=9,color="777777")
            c2.alignment=left

    def tot_row(ws, row_n, ncols, vals_dict):
        for col in range(1,ncols+1):
            c=ws.cell(row=row_n,column=col); c.font=tf; c.fill=tbg; c.alignment=center; c.border=bd
        for col, val in vals_dict.items():
            ws.cell(row=row_n,column=col).value=val

    # ── 시트1: 수당 요약 ─────────────────────
    ws1=wb.active; ws1.title="수당 요약"
    title_row(ws1, f"{label} 운항 수당 정산표 — 요약", 5,
              "※ 야간: 22:00~06:00(KST) | 연장: 일 8시간 초과 | 3P: 편명 0151·0152")
    style_hdr(ws1, 3, ["No","이름","야간 시간","연장 시간","3P 시간"])
    for i,row in flt_sum.iterrows():
        r=i+4
        for col,val in enumerate([i+1,row["이름"],fmt_hhmm(row["야간"]),fmt_hhmm(row["연장"]),fmt_hhmm(row["P3"])],1):
            c=ws1.cell(row=r,column=col,value=val); c.font=cf; c.border=bd
            c.alignment=left if col==2 else center
    tr=len(flt_sum)+4
    ws1.merge_cells(f"A{tr}:B{tr}")
    tot_row(ws1,tr,5,{1:"합계",3:fmt_hhmm(flt_sum["야간"].sum()),4:fmt_hhmm(flt_sum["연장"].sum()),5:fmt_hhmm(flt_sum["P3"].sum())})
    for i,w in enumerate([5,14,11,11,11],1): ws1.column_dimensions[get_column_letter(i)].width=w
    ws1.freeze_panes="A4"

    # ── 시트2: 개인별 상세 ───────────────────
    ws2=wb.create_sheet("개인별 상세")
    title_row(ws2, f"{label} 운항 수당 정산표 — 개인별 비행 상세", 11)
    style_hdr(ws2,2,["이름","날짜","편명","구간","ATD(KST)","ATA(KST)","CI(KST)","CO(KST)","야간(h)","연장(h)","3P(h)"])
    names=sorted(flt_det["이름"].unique()); cur=3; ci=0; fills=["FFFFFF","F7FBFF"]
    for name in names:
        grp=flt_det[flt_det["이름"]==name].reset_index(drop=True)
        rf=PatternFill("solid",fgColor=fills[ci%2]); ci+=1
        for _,dr in grp.iterrows():
            for col,key in enumerate(["이름","날짜","편명","구간","ATD(KST)","ATA(KST)","CI(KST)","CO(KST)","야간(h)","연장(h)","3P(h)"],1):
                val=dr[key]
                if key in ["야간(h)","연장(h)","3P(h)"]: val=fmt_hhmm(val)
                c=ws2.cell(row=cur,column=col,value=val); c.font=cf; c.fill=rf; c.border=bd
                c.alignment=left if col in [1,4] else center
            cur+=1
        sv=[name+" 소계","","","","","","","",fmt_hhmm(grp["야간(h)"].sum()),fmt_hhmm(grp["연장(h)"].sum()),fmt_hhmm(grp["3P(h)"].sum())]
        for col,val in enumerate(sv,1):
            c=ws2.cell(row=cur,column=col,value=val); c.font=sbf; c.fill=sbbg; c.alignment=center; c.border=bd
        cur+=1
    for col in range(1,12):
        c=ws2.cell(row=cur,column=col); c.font=tf; c.fill=tbg; c.alignment=center; c.border=bd
    ws2.cell(row=cur,column=1).value="전체 합계"
    ws2.cell(row=cur,column=9).value=fmt_hhmm(flt_det["야간(h)"].sum())
    ws2.cell(row=cur,column=10).value=fmt_hhmm(flt_det["연장(h)"].sum())
    ws2.cell(row=cur,column=11).value=fmt_hhmm(flt_det["3P(h)"].sum())
    for i,w in enumerate([14,8,8,12,10,10,10,10,10,10,10],1): ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.freeze_panes="A3"

    # ── 시트3: 총비행시간 ────────────────────
    ws3=wb.create_sheet("총비행시간")
    hdrs3=["No","이름","Block(h)","DHC(h)","DHC×50%(h)","OBCA/OBFO(h)","총비행시간(h)"]
    title_row(ws3,f"{label} 개인별 총 비행시간",len(hdrs3),"※ 총비행시간 = Block + DHC×50% - OBCA/OBFO")
    style_hdr(ws3,3,hdrs3)
    ds=calc_df.sort_values("Crew Code").reset_index(drop=True)
    for i,row in ds.iterrows():
        r=i+4
        for col,val in enumerate([i+1,row["Crew Code"],fmt_hhmm(row["Block"]),fmt_hhmm(row["DHC"]),
                                   fmt_hhmm(row["DHC_50"]),fmt_hhmm(row["OBCA_OBFO"]),fmt_hhmm(row["Total_Flt"])],1):
            c=ws3.cell(row=r,column=col,value=val); c.font=cf; c.border=bd
            c.alignment=left if col==2 else center
    tr3=len(ds)+4; ws3.merge_cells(f"A{tr3}:B{tr3}")
    tot_row(ws3,tr3,len(hdrs3),{1:"합계",3:fmt_hhmm(ds["Block"].sum()),4:fmt_hhmm(ds["DHC"].sum()),
             5:fmt_hhmm(ds["DHC_50"].sum()),6:fmt_hhmm(ds["OBCA_OBFO"].sum()),7:fmt_hhmm(ds["Total_Flt"].sum())})
    for i,w in enumerate([5,14,11,11,11,13,13],1): ws3.column_dimensions[get_column_letter(i)].width=w
    ws3.freeze_panes="A4"

    # ── 시트4: DAYOFF 미달 ───────────────────
    ws4=wb.create_sheet("DAYOFF 8회 미만")
    hdrs4=["No","이름","DAYOFF 횟수","비고"]
    title_row(ws4,f"{label} DAYOFF 8회 미만자",len(hdrs4))
    style_hdr(ws4,2,hdrs4)
    u8=calc_df[calc_df["Dayoff_Under8"]].sort_values("Crew Code").reset_index(drop=True)
    if len(u8)==0:
        ws4.merge_cells("A3:D3")
        c=ws4["A3"]; c.value="8회 미만자 없음"
        c.font=Font(name="Arial",italic=True,size=10,color="777777"); c.alignment=center
    else:
        for i,row in u8.iterrows():
            r=i+3
            for col,val in enumerate([i+1,row["Crew Code"],int(row["Dayoff"]),f"기준 미달 ({int(row['Dayoff'])}회 / 8회)"],1):
                c=ws4.cell(row=r,column=col,value=val); c.font=redf; c.fill=redbg; c.border=bd
                c.alignment=left if col in [2,4] else center
    for i,w in enumerate([5,14,14,24],1): ws4.column_dimensions[get_column_letter(i)].width=w
    ws4.freeze_panes="A3"

    # ── 시트5: 교관 수당 요약 ────────────────
    ws5=wb.create_sheet("교관수당 요약")
    hdrs5=["No","이름","1set 합계","2set 합계","3P 합계","교관시간 총계"]
    title_row(ws5,f"{label} 교관 비행 수당 요약",len(hdrs5),
              "※ 1set=BH전체(DAD·BKK·HKG·NRT) | 2set=BH×1/2(DAC·LAX·EWR·SFO) | 3P=BH×1/3(HNL)")
    style_hdr(ws5,3,hdrs5)
    if not instr_det.empty:
        instr_sum=(instr_det.groupby("이름").agg(
            s1=("교관시간_h", lambda x: x[instr_det.loc[x.index,"Set구분"]=="1set"].sum()),
            s2=("교관시간_h", lambda x: x[instr_det.loc[x.index,"Set구분"]=="2set"].sum()),
            s3=("교관시간_h", lambda x: x[instr_det.loc[x.index,"Set구분"]=="3P"].sum()),
            tot=("교관시간_h","sum")
        ).reset_index().sort_values("이름").reset_index(drop=True))
        for i,row in instr_sum.iterrows():
            r=i+4
            for col,val in enumerate([i+1,row["이름"],fmt_hhmm(row["s1"]),fmt_hhmm(row["s2"]),fmt_hhmm(row["s3"]),fmt_hhmm(row["tot"])],1):
                c=ws5.cell(row=r,column=col,value=val); c.font=cf; c.border=bd
                c.alignment=left if col==2 else center
        tr5=len(instr_sum)+4; ws5.merge_cells(f"A{tr5}:B{tr5}")
        tot_row(ws5,tr5,len(hdrs5),{1:"합계",
            3:fmt_hhmm(instr_det[instr_det["Set구분"]=="1set"]["교관시간_h"].sum()),
            4:fmt_hhmm(instr_det[instr_det["Set구분"]=="2set"]["교관시간_h"].sum()),
            5:fmt_hhmm(instr_det[instr_det["Set구분"]=="3P"]["교관시간_h"].sum()),
            6:fmt_hhmm(instr_det["교관시간_h"].sum())})
    for i,w in enumerate([5,14,12,12,12,14],1): ws5.column_dimensions[get_column_letter(i)].width=w
    ws5.freeze_panes="A4"

    # ── 시트6: 교관 수당 상세 ────────────────
    ws6=wb.create_sheet("교관수당 상세")
    hdrs6=["이름","날짜(KST)","DC","편명","From","To","Set구분","Blhr(원본)","교관시간","산출방식"]
    title_row(ws6,f"{label} 교관 비행 수당 상세",len(hdrs6))
    style_hdr(ws6,2,hdrs6)
    cur6=3; ci6=0; fills6=["FFFFFF","EDF4FF"]
    if not instr_det.empty:
        names6=sorted(instr_det["이름"].unique())
        for name in names6:
            grp=instr_det[instr_det["이름"]==name].reset_index(drop=True)
            rf=PatternFill("solid",fgColor=fills6[ci6%2]); ci6+=1
            for _,dr in grp.iterrows():
                st=dr["Set구분"]
                if st=="1set":   formula="BH × 1 (1set)"
                elif st=="2set": formula="BH × 1/2 (2set)"
                else:            formula="BH × 1/3 (3P)"
                vals=[dr["이름"],dr["날짜(KST)"],dr["DC"],dr["편명"],dr["From"],dr["To"],
                      st,dr["Blhr(원본)"],dr["교관시간"],formula]
                for col,val in enumerate(vals,1):
                    c=ws6.cell(row=cur6,column=col,value=val); c.font=cf; c.fill=rf; c.border=bd
                    c.alignment=left if col in [1,4,10] else center
                cur6+=1
            sv=[name+" 소계","","","","","","",
                "",fmt_hhmm(grp["교관시간_h"].sum()),""]
            for col,val in enumerate(sv,1):
                c=ws6.cell(row=cur6,column=col,value=val); c.font=sbf; c.fill=sbbg; c.alignment=center; c.border=bd
            cur6+=1
        for col in range(1,len(hdrs6)+1):
            c=ws6.cell(row=cur6,column=col); c.font=tf; c.fill=tbg; c.alignment=center; c.border=bd
        ws6.cell(row=cur6,column=1).value="전체 합계"
        ws6.cell(row=cur6,column=9).value=fmt_hhmm(instr_det["교관시간_h"].sum())
    for i,w in enumerate([14,9,8,10,6,6,8,10,10,16],1): ws6.column_dimensions[get_column_letter(i)].width=w
    ws6.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════
# Streamlit UI
# ═══════════════════════════════════════════
st.set_page_config(page_title="운항 수당 정산기", page_icon="✈️", layout="centered")
st.title("✈️ 운항 수당 정산기")
st.caption("4개 파일 업로드 → 월 선택 → 정산 실행 → 엑셀 다운로드 (시트 6개)")
st.divider()

c1,c2 = st.columns(2)
with c1:
    flt_file  = st.file_uploader("📂 FltReport.xlsx", type=["xlsx"])
    dhc_file  = st.file_uploader("📂 월DHC_DAYOFF_총비행시간.xlsx", type=["xlsx"])
with c2:
    ob_file   = st.file_uploader("📂 OBCA.xlsx", type=["xlsx"])
    rost_file = st.file_uploader("📂 Roster.xlsx (교관수당)", type=["xlsx"])

all_uploaded = flt_file and dhc_file and ob_file and rost_file

if all_uploaded:
    try:
        flt_df = pd.read_excel(flt_file)
        required={"Date","Flight","ATD","ATA","Bl Hrs","운항 C/I","운항 C/O","운항 Crew"}
        missing=required-set(flt_df.columns)
        if missing: st.error(f"FltReport 필수 열 없음: {missing}"); st.stop()

        flt_df["date_parsed"]=pd.to_datetime(flt_df["Date"],format="%d%b%y")
        available=sorted(flt_df[flt_df["운항 Crew"].notna()]["date_parsed"].dt.to_period("M").unique(),reverse=True)

        base_df   = parse_dhc_file(dhc_file)
        ob_df     = parse_obca_file(ob_file)
        calc_df   = calc_summary(base_df, ob_df)
        instr_det = parse_roster_file(rost_file)  # 월 필터는 정산 실행 시 적용

        n_instr = instr_det["이름"].nunique() if not instr_det.empty else 0
        st.success(f"✅ 4개 파일 로드 완료 — FltReport {len(flt_df):,}행 · 승무원 {len(base_df)}명 · 교관 {n_instr}명")

        selected_str=st.selectbox("📅 정산할 월 선택",[str(p) for p in available])
        sel=pd.Period(selected_str,freq="M"); target_year,target_month=sel.year,sel.month

        with st.expander("ℹ️ 계산 기준 보기"):
            st.markdown("""
| 항목 | 기준 |
|------|------|
| **야간** | 운항 C/I~C/O(KST) 중 22:00~06:00 겹치는 시간 |
| **연장** | ATD~ATA(KST) 8시간 초과분 |
| **3P** | 편명 0151·0152 BI Hrs |
| **총비행시간** | Block + DHC×50% - OBCA/OBFO |
| **DAYOFF 미달** | 월 DAYOFF 8회 미만 |
| **교관수당 1set** | BH 전체 (DAD·BKK·HKG·NRT) |
| **교관수당 2set** | BH × 1/2 (DAC·LAX·EWR·SFO) |
| **교관수당 3P** | BH × 1/3 (HNL) |
| **교관 DC** | LIP·LCP·DLCP·I·I* (BH 없으면 무시) |
            """)

        if st.button("🚀 정산 실행", type="primary", use_container_width=True):
            with st.spinner("계산 중..."):
                flt_sum, flt_det = process_flt(flt_df, target_month, target_year)
                instr_det = parse_roster_file(rost_file, target_month, target_year)

            if flt_sum.empty:
                st.warning("FltReport에서 해당 월 데이터가 없습니다.")
            else:
                st.success(f"✅ {target_year}년 {target_month}월 정산 완료")

                tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs(
                    ["📊 수당 요약","📋 개인별 상세","⏱ 총비행시간","📅 DAYOFF 미달","🎓 교관수당 요약","🎓 교관수당 상세"])

                with tab1:
                    d=flt_sum.copy()
                    for k in ["야간","연장","P3"]: d[k]=d[k].apply(fmt_hhmm)
                    d.index+=1; d.columns=["이름","야간 시간","연장 시간","3P 시간"]
                    st.dataframe(d,use_container_width=True,height=360)
                    c1,c2,c3=st.columns(3)
                    c1.metric("야간",fmt_hhmm(flt_sum["야간"].sum()))
                    c2.metric("연장",fmt_hhmm(flt_sum["연장"].sum()))
                    c3.metric("3P",fmt_hhmm(flt_sum["P3"].sum()))

                with tab2:
                    nl=["전체"]+sorted(flt_det["이름"].unique().tolist())
                    sn=st.selectbox("승무원 선택",nl)
                    vw=flt_det if sn=="전체" else flt_det[flt_det["이름"]==sn]
                    d2=vw.copy()
                    for col in ["야간(h)","연장(h)","3P(h)"]: d2[col]=d2[col].apply(fmt_hhmm)
                    st.dataframe(d2.reset_index(drop=True),use_container_width=True,height=380)

                with tab3:
                    d3=calc_df.sort_values("Crew Code").reset_index(drop=True).copy()
                    for col in ["Block","DHC","DHC_50","OBCA_OBFO","Total_Flt"]: d3[col]=d3[col].apply(fmt_hhmm)
                    d3=d3[["Crew Code","Block","DHC","DHC_50","OBCA_OBFO","Total_Flt","Dayoff"]]
                    d3.columns=["이름","Block(h)","DHC(h)","DHC×50%(h)","OBCA/OBFO(h)","총비행시간(h)","DAYOFF"]
                    d3.index+=1
                    st.dataframe(d3,use_container_width=True,height=380)

                with tab4:
                    u8=calc_df[calc_df["Dayoff_Under8"]].sort_values("Crew Code").reset_index(drop=True)
                    if len(u8)==0: st.success("8회 미만자 없음 ✅")
                    else:
                        st.error(f"⚠️ DAYOFF 8회 미만자 {len(u8)}명")
                        d4=u8[["Crew Code","Dayoff"]].copy(); d4.columns=["이름","DAYOFF 횟수"]; d4.index+=1
                        st.dataframe(d4,use_container_width=True)

                with tab5:
                    if instr_det.empty:
                        st.info("교관 비행 데이터 없음")
                    else:
                        instr_sum=(instr_det.groupby("이름").agg(
                            s1=("교관시간_h",lambda x:x[instr_det.loc[x.index,"Set구분"]=="1set"].sum()),
                            s2=("교관시간_h",lambda x:x[instr_det.loc[x.index,"Set구분"]=="2set"].sum()),
                            s3=("교관시간_h",lambda x:x[instr_det.loc[x.index,"Set구분"]=="3P"].sum()),
                            tot=("교관시간_h","sum")
                        ).reset_index().sort_values("이름").reset_index(drop=True))
                        d5=instr_sum.copy()
                        for col in ["s1","s2","s3","tot"]: d5[col]=d5[col].apply(fmt_hhmm)
                        d5.index+=1; d5.columns=["이름","1set 합계","2set 합계","3P 합계","교관시간 총계"]
                        st.dataframe(d5,use_container_width=True,height=360)
                        c1,c2,c3,c4=st.columns(4)
                        c1.metric("1set",fmt_hhmm(instr_det[instr_det["Set구분"]=="1set"]["교관시간_h"].sum()))
                        c2.metric("2set",fmt_hhmm(instr_det[instr_det["Set구분"]=="2set"]["교관시간_h"].sum()))
                        c3.metric("3P",fmt_hhmm(instr_det[instr_det["Set구분"]=="3P"]["교관시간_h"].sum()))
                        c4.metric("전체",fmt_hhmm(instr_det["교관시간_h"].sum()))

                with tab6:
                    if instr_det.empty:
                        st.info("교관 비행 데이터 없음")
                    else:
                        nl2=["전체"]+sorted(instr_det["이름"].unique().tolist())
                        sn2=st.selectbox("교관 선택",nl2,key="instr_sel")
                        vw2=instr_det if sn2=="전체" else instr_det[instr_det["이름"]==sn2]
                        d6=vw2[["이름","날짜(KST)","DC","편명","From","To","Set구분","Blhr(원본)","교관시간"]].copy()
                        d6.index+=1
                        st.dataframe(d6,use_container_width=True,height=400)
                        if sn2!="전체":
                            cc1,cc2,cc3=st.columns(3)
                            cc1.metric("1set",fmt_hhmm(vw2[vw2["Set구분"]=="1set"]["교관시간_h"].sum()))
                            cc2.metric("2set",fmt_hhmm(vw2[vw2["Set구분"]=="2set"]["교관시간_h"].sum()))
                            cc3.metric("3P",fmt_hhmm(vw2[vw2["Set구분"]=="3P"]["교관시간_h"].sum()))

                excel_buf=build_excel(flt_sum,flt_det,calc_df,instr_det,target_year,target_month)
                fname=f"{target_year}년{target_month:02d}월_운항정산.xlsx"
                st.download_button(
                    label="⬇️ 엑셀 다운로드 (6개 시트: 수당요약·개인상세·총비행시간·DAYOFF·교관요약·교관상세)",
                    data=excel_buf, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    except Exception as e:
        st.error(f"오류 발생: {e}"); st.exception(e)
else:
    missing=[]
    if not flt_file:  missing.append("FltReport.xlsx")
    if not dhc_file:  missing.append("월DHC_DAYOFF_총비행시간.xlsx")
    if not ob_file:   missing.append("OBCA.xlsx")
    if not rost_file: missing.append("Roster.xlsx")
    st.info(f"아래 파일을 모두 업로드해주세요: {' · '.join(missing)}")
