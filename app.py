import streamlit as st
import pandas as pd
import io
import datetime
from datetime import timedelta, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KST_OFFSET = timedelta(hours=9)

# ═══════════════════════════════════════════
# 공통 유틸
# ═══════════════════════════════════════════
def fmt_hhmm(h):
    if h is None or (isinstance(h, float) and pd.isna(h)) or h == 0:
        return "-"
    hours = int(h); mins = round((h - hours) * 60)
    if mins == 60: hours += 1; mins = 0
    return f"{hours:02d}:{mins:02d}"

def fmt_time(t):
    if t is None or (isinstance(t, float) and pd.isna(t)): return "-"
    if isinstance(t, (time, datetime.datetime)): return t.strftime("%H:%M")
    return str(t)

def parse_hhmm(val):
    if pd.isna(val): return 0.0
    val = str(val).strip()
    if ":" in val:
        p = val.split(":")
        return int(p[0]) + int(p[1]) / 60
    try: return float(val)
    except: return 0.0

def td_to_hours(val):
    if isinstance(val, datetime.timedelta): return val.total_seconds() / 3600
    return 0.0

def make_border():
    thin = Side(style="thin", color="BBBBBB")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

# ═══════════════════════════════════════════
# 파일1: DHC / DAYOFF / 총비행시간 파싱
# ═══════════════════════════════════════════
def parse_dhc_file(uploaded):
    df = pd.read_excel(uploaded)
    data = df.iloc[4:].copy()
    data.columns = ["Crew Code", "Position", "Name", "Counter", "Mar26", "Total"]
    data = data[data["Counter"] != "Counter"].copy()

    block  = data[data["Counter"] == "Block"][["Crew Code","Mar26"]].copy()
    dayoff = data[data["Counter"] == "dayoff"][["Crew Code","Mar26"]].copy()
    dhc    = data[data["Counter"] == "DHC"][["Crew Code","Mar26"]].copy()

    block.columns  = ["Crew Code", "Block"]
    dayoff.columns = ["Crew Code", "Dayoff"]
    dhc.columns    = ["Crew Code", "DHC"]

    block["Block"]   = block["Block"].apply(parse_hhmm)
    dayoff["Dayoff"] = pd.to_numeric(dayoff["Dayoff"], errors="coerce").fillna(0).astype(int)
    dhc["DHC"]       = dhc["DHC"].apply(parse_hhmm)

    merged = block.merge(dayoff, on="Crew Code").merge(dhc, on="Crew Code")
    return merged

# ═══════════════════════════════════════════
# 파일2: OBCA / OBFO 파싱
# ═══════════════════════════════════════════
def parse_obca_file(uploaded):
    df = pd.read_excel(uploaded)
    data = df.iloc[6:].copy()
    data.columns = ["Crew Code","AC Type","Position","Block Hours","Cruise Time","Sectors","Valid From","Valid To"]
    data = data[data["Crew Code"] != "Crew Code"].copy()
    data["Crew Code"] = data["Crew Code"].ffill()
    ob = data[data["Position"].isin(["OBFO","OBCA"])].copy()
    ob["OB_hrs"] = ob["Block Hours"].apply(td_to_hours)
    ob_sum = ob.groupby("Crew Code")["OB_hrs"].sum().reset_index()
    ob_sum.columns = ["Crew Code", "OBCA_OBFO"]
    return ob_sum

# ═══════════════════════════════════════════
# 종합 산출 로직
# ═══════════════════════════════════════════
def calc_summary(base_df, ob_df):
    df = base_df.merge(ob_df, on="Crew Code", how="left")
    df["OBCA_OBFO"] = df["OBCA_OBFO"].fillna(0)

    # 총비행시간 = Block + DHC×50% - OBCA_OBFO
    df["DHC_50"]     = df["DHC"] * 0.5
    df["Total_Flt"]  = df["Block"] + df["DHC_50"] - df["OBCA_OBFO"]
    df["Total_Flt"]  = df["Total_Flt"].clip(lower=0)

    # DAYOFF 8회 미만
    df["Dayoff_Under8"] = df["Dayoff"] < 8
    return df

# ═══════════════════════════════════════════
# FltReport 수당 계산 (기존 로직)
# ═══════════════════════════════════════════
def combine_date_time_utc(date_str, t):
    if pd.isna(t) or t is None: return None
    base = datetime.datetime.strptime(date_str, "%d%b%y")
    return datetime.datetime(base.year, base.month, base.day, t.hour, t.minute, t.second)

def calc_night_hours(ci_kst, co_kst):
    if ci_kst is None or co_kst is None: return 0.0
    total = 0.0; ci_date = ci_kst.date()
    for d in [-1, 0, 1]:
        ns = datetime.datetime.combine(ci_date + timedelta(days=d), time(22, 0))
        ne = datetime.datetime.combine(ci_date + timedelta(days=d+1), time(6, 0))
        os_ = max(ci_kst, ns); oe_ = min(co_kst, ne)
        if oe_ > os_: total += (oe_ - os_).total_seconds() / 3600
    return total

def calc_overtime_hours(atd_kst, ata_kst):
    if atd_kst is None or ata_kst is None: return 0.0
    return max(0.0, (ata_kst - atd_kst).total_seconds() / 3600 - 8)

def blhrs_to_decimal(t):
    if pd.isna(t) or t is None: return 0.0
    if isinstance(t, time): return t.hour + t.minute/60 + t.second/3600
    return 0.0

def process_flt(df, target_month, target_year):
    summary_rows, detail_rows = [], []
    for _, row in df.iterrows():
        crew_str = row["운항 Crew"]
        if pd.isna(crew_str): continue
        date_str = row["Date"]
        atd_utc = combine_date_time_utc(date_str, row["ATD"])
        if atd_utc is None: continue
        atd_kst = atd_utc + KST_OFFSET
        ata_utc = combine_date_time_utc(date_str, row["ATA"])
        ata_kst = (ata_utc + KST_OFFSET) if ata_utc else None
        ci_utc  = combine_date_time_utc(date_str, row["운항 C/I"])
        co_utc  = combine_date_time_utc(date_str, row["운항 C/O"])
        ci_kst  = (ci_utc + KST_OFFSET) if ci_utc else None
        co_kst  = (co_utc + KST_OFFSET) if co_utc else None
        if ata_kst and ata_kst < atd_kst: ata_kst += timedelta(days=1)
        if co_kst and ci_kst and co_kst < ci_kst: co_kst += timedelta(days=1)
        if atd_kst.month != target_month or atd_kst.year != target_year: continue
        flight = str(row["Flight"]); bl_hrs = blhrs_to_decimal(row["Bl Hrs"])
        is_3p = flight in ["0151","0152"]
        night = calc_night_hours(ci_kst, co_kst)
        ot    = calc_overtime_hours(atd_kst, ata_kst)
        p3    = bl_hrs if is_3p else 0.0
        route = f"{row['From']}→{row['To']}"
        for name in str(crew_str).split():
            name = name.strip()
            if not name: continue
            summary_rows.append({"이름": name, "night": night, "ot": ot, "p3": p3})
            detail_rows.append({
                "이름": name, "날짜": atd_kst.strftime("%m/%d"),
                "편명": flight, "구간": route,
                "ATD(KST)": fmt_time(atd_kst), "ATA(KST)": fmt_time(ata_kst),
                "CI(KST)": fmt_time(ci_kst), "CO(KST)": fmt_time(co_kst),
                "야간(h)": night, "연장(h)": ot, "3P(h)": p3,
            })
    if not summary_rows:
        return pd.DataFrame(), pd.DataFrame()
    s_df = pd.DataFrame(summary_rows)
    summary = (s_df.groupby("이름")
        .agg(야간=("night","sum"), 연장=("ot","sum"), P3=("p3","sum"))
        .reset_index().sort_values("이름").reset_index(drop=True))
    detail = pd.DataFrame(detail_rows).sort_values(["이름","날짜","편명"]).reset_index(drop=True)
    return summary, detail

# ═══════════════════════════════════════════
# 엑셀 생성
# ═══════════════════════════════════════════
def write_sheet_header(ws, title, note=None):
    ws.merge_cells(f"A1:{get_column_letter(ws.max_column or 10)}1")
    c = ws["A1"]
    c.value = title
    c.font = Font(name="Arial", bold=True, size=13)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    if note:
        ws.merge_cells(f"A2:{get_column_letter(ws.max_column or 10)}2")
        c2 = ws["A2"]
        c2.value = note
        c2.font = Font(name="Arial", italic=True, size=9, color="777777")
        c2.alignment = Alignment(horizontal="left", vertical="center")

def style_header_row(ws, row, headers, bg="1F4E79"):
    bd = make_border()
    hdr_f = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_bg = PatternFill("solid", fgColor=bg)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = hdr_f; c.fill = hdr_bg; c.alignment = center; c.border = bd
    ws.row_dimensions[row].height = 20

def build_excel(flt_summary, flt_detail, calc_df, target_year, target_month):
    wb = Workbook()
    bd = make_border()
    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left", vertical="center")
    cel_f  = Font(name="Arial", size=10)
    tot_f  = Font(name="Arial", bold=True, size=10)
    tot_bg = PatternFill("solid", fgColor="FFF2CC")
    sub_f  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    sub_bg = PatternFill("solid", fgColor="2E75B6")
    red_f  = Font(name="Arial", bold=True, size=10, color="C00000")
    label  = f"{target_year}년 {target_month}월"

    # ── 시트1: 수당 요약 ─────────────────────
    ws1 = wb.active; ws1.title = "수당 요약"
    style_header_row(ws1, 1, [""] * 5)  # placeholder, overwrite below
    ws1.merge_cells("A1:E1")
    c = ws1["A1"]
    c.value = f"{label} 운항 수당 정산표 — 요약"
    c.font  = Font(name="Arial", bold=True, size=13)
    c.alignment = center
    ws1.row_dimensions[1].height = 28
    ws1.merge_cells("A2:E2")
    c2 = ws1["A2"]
    c2.value = "※ 야간: 22:00~06:00(KST) | 연장: 일 8시간 초과 | 3P: 편명 0151·0152"
    c2.font = Font(name="Arial", italic=True, size=9, color="777777")
    c2.alignment = left

    style_header_row(ws1, 3, ["No","이름","야간 시간","연장 시간","3P 시간"])
    for i, row in flt_summary.iterrows():
        r = i + 4
        for col, val in enumerate([i+1, row["이름"], fmt_hhmm(row["야간"]),
                                    fmt_hhmm(row["연장"]), fmt_hhmm(row["P3"])], 1):
            c = ws1.cell(row=r, column=col, value=val)
            c.font=cel_f; c.border=bd; c.alignment=left if col==2 else center
    tr = len(flt_summary)+4
    ws1.merge_cells(f"A{tr}:B{tr}")
    for col in range(1,6):
        c = ws1.cell(row=tr, column=col); c.font=tot_f; c.fill=tot_bg; c.alignment=center; c.border=bd
    ws1["A"+str(tr)].value = "합계"
    for col,key in zip([3,4,5],["야간","연장","P3"]):
        ws1.cell(row=tr, column=col).value = fmt_hhmm(flt_summary[key].sum())
    for i,w in enumerate([5,14,11,11,11],1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A4"

    # ── 시트2: 개인별 상세 ───────────────────
    ws2 = wb.create_sheet("개인별 상세")
    ws2.merge_cells("A1:K1")
    c = ws2["A1"]; c.value = f"{label} 운항 수당 정산표 — 개인별 비행 상세"
    c.font=Font(name="Arial",bold=True,size=13); c.alignment=center
    ws2.row_dimensions[1].height=28
    style_header_row(ws2, 2, ["이름","날짜","편명","구간","ATD(KST)","ATA(KST)","CI(KST)","CO(KST)","야간(h)","연장(h)","3P(h)"])
    names = sorted(flt_detail["이름"].unique())
    cur_row = 3; fill_colors=["FFFFFF","F7FBFF"]; cidx=0
    for name in names:
        grp = flt_detail[flt_detail["이름"]==name].reset_index(drop=True)
        row_fill = PatternFill("solid", fgColor=fill_colors[cidx%2]); cidx+=1
        for _, dr in grp.iterrows():
            for col, key in enumerate(["이름","날짜","편명","구간","ATD(KST)","ATA(KST)","CI(KST)","CO(KST)","야간(h)","연장(h)","3P(h)"],1):
                val = dr[key]
                if key in ["야간(h)","연장(h)","3P(h)"]: val = fmt_hhmm(val)
                c = ws2.cell(row=cur_row, column=col, value=val)
                c.font=cel_f; c.fill=row_fill; c.border=bd
                c.alignment = left if col in [1,4] else center
            cur_row+=1
        sv = [name+" 소계","","","","","","","",
              fmt_hhmm(grp["야간(h)"].sum()), fmt_hhmm(grp["연장(h)"].sum()), fmt_hhmm(grp["3P(h)"].sum())]
        for col,val in enumerate(sv,1):
            c=ws2.cell(row=cur_row,column=col,value=val); c.font=sub_f; c.fill=sub_bg; c.alignment=center; c.border=bd
        cur_row+=1
    for col in range(1,12):
        c=ws2.cell(row=cur_row,column=col); c.font=tot_f; c.fill=tot_bg; c.alignment=center; c.border=bd
    ws2.cell(row=cur_row,column=1).value="전체 합계"
    ws2.cell(row=cur_row,column=9).value=fmt_hhmm(flt_detail["야간(h)"].sum())
    ws2.cell(row=cur_row,column=10).value=fmt_hhmm(flt_detail["연장(h)"].sum())
    ws2.cell(row=cur_row,column=11).value=fmt_hhmm(flt_detail["3P(h)"].sum())
    for i,w in enumerate([14,8,8,12,10,10,10,10,10,10,10],1):
        ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.freeze_panes="A3"

    # ── 시트3: 총비행시간 ────────────────────
    ws3 = wb.create_sheet("총비행시간")
    hdrs3 = ["No","이름","Block(h)","DHC(h)","DHC×50%(h)","OBCA/OBFO(h)","총비행시간(h)"]
    ws3.merge_cells(f"A1:{get_column_letter(len(hdrs3))}1")
    c=ws3["A1"]; c.value=f"{label} 개인별 총 비행시간"
    c.font=Font(name="Arial",bold=True,size=13); c.alignment=center; ws3.row_dimensions[1].height=28
    ws3.merge_cells(f"A2:{get_column_letter(len(hdrs3))}2")
    c2=ws3["A2"]; c2.value="※ 총비행시간 = Block + DHC×50% - OBCA/OBFO"
    c2.font=Font(name="Arial",italic=True,size=9,color="777777"); c2.alignment=left
    style_header_row(ws3, 3, hdrs3)
    df_sorted = calc_df.sort_values("Crew Code").reset_index(drop=True)
    for i, row in df_sorted.iterrows():
        r = i+4
        vals = [i+1, row["Crew Code"], fmt_hhmm(row["Block"]), fmt_hhmm(row["DHC"]),
                fmt_hhmm(row["DHC_50"]), fmt_hhmm(row["OBCA_OBFO"]), fmt_hhmm(row["Total_Flt"])]
        for col,val in enumerate(vals,1):
            c=ws3.cell(row=r,column=col,value=val)
            c.font=cel_f; c.border=bd; c.alignment=left if col==2 else center
    tr3=len(df_sorted)+4
    ws3.merge_cells(f"A{tr3}:B{tr3}")
    for col in range(1,len(hdrs3)+1):
        c=ws3.cell(row=tr3,column=col); c.font=tot_f; c.fill=tot_bg; c.alignment=center; c.border=bd
    ws3["A"+str(tr3)].value="합계"
    ws3.cell(row=tr3,column=3).value=fmt_hhmm(df_sorted["Block"].sum())
    ws3.cell(row=tr3,column=4).value=fmt_hhmm(df_sorted["DHC"].sum())
    ws3.cell(row=tr3,column=5).value=fmt_hhmm(df_sorted["DHC_50"].sum())
    ws3.cell(row=tr3,column=6).value=fmt_hhmm(df_sorted["OBCA_OBFO"].sum())
    ws3.cell(row=tr3,column=7).value=fmt_hhmm(df_sorted["Total_Flt"].sum())
    for i,w in enumerate([5,14,11,11,11,13,13],1):
        ws3.column_dimensions[get_column_letter(i)].width=w
    ws3.freeze_panes="A4"

    # ── 시트4: DAYOFF 8회 미만 ───────────────
    ws4 = wb.create_sheet("DAYOFF 8회 미만")
    hdrs4=["No","이름","DAYOFF 횟수","비고"]
    ws4.merge_cells(f"A1:{get_column_letter(len(hdrs4))}1")
    c=ws4["A1"]; c.value=f"{label} DAYOFF 8회 미만자"
    c.font=Font(name="Arial",bold=True,size=13); c.alignment=center; ws4.row_dimensions[1].height=28
    style_header_row(ws4, 2, hdrs4)
    under8 = calc_df[calc_df["Dayoff_Under8"]].sort_values("Crew Code").reset_index(drop=True)
    if len(under8)==0:
        ws4.merge_cells("A3:D3")
        c=ws4["A3"]; c.value="8회 미만자 없음"
        c.font=Font(name="Arial",italic=True,size=10,color="777777"); c.alignment=center
    else:
        alert_bg = PatternFill("solid", fgColor="FCEAEA")
        for i, row in under8.iterrows():
            r=i+3
            vals=[i+1, row["Crew Code"], int(row["Dayoff"]), f"기준 미달 ({int(row['Dayoff'])}회 / 8회)"]
            for col,val in enumerate(vals,1):
                c=ws4.cell(row=r,column=col,value=val)
                c.font=red_f; c.fill=alert_bg; c.border=bd; c.alignment=left if col in [2,4] else center
    for i,w in enumerate([5,14,14,24],1):
        ws4.column_dimensions[get_column_letter(i)].width=w
    ws4.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════
# Streamlit UI
# ═══════════════════════════════════════════
st.set_page_config(page_title="운항 수당 정산기", page_icon="✈️", layout="centered")
st.title("✈️ 운항 수당 정산기")
st.caption("3개 파일 업로드 → 월 선택 → 정산 실행 → 엑셀 다운로드 (시트 4개)")
st.divider()

col1, col2, col3 = st.columns(3)
with col1:
    flt_file = st.file_uploader("📂 FltReport.xlsx", type=["xlsx"])
with col2:
    dhc_file = st.file_uploader("📂 월DHC_DAYOFF_총비행시간.xlsx", type=["xlsx"])
with col3:
    ob_file  = st.file_uploader("📂 OBCA.xlsx", type=["xlsx"])

all_uploaded = flt_file and dhc_file and ob_file

if all_uploaded:
    try:
        # FltReport 파싱
        flt_df = pd.read_excel(flt_file)
        required = {"Date","Flight","ATD","ATA","Bl Hrs","운항 C/I","운항 C/O","운항 Crew"}
        missing = required - set(flt_df.columns)
        if missing: st.error(f"FltReport 필수 열 없음: {missing}"); st.stop()

        flt_df["date_parsed"] = pd.to_datetime(flt_df["Date"], format="%d%b%y")
        available = sorted(
            flt_df[flt_df["운항 Crew"].notna()]["date_parsed"].dt.to_period("M").unique(),
            reverse=True)

        # DHC / OBCA 파싱
        base_df = parse_dhc_file(dhc_file)
        ob_df   = parse_obca_file(ob_file)
        calc_df = calc_summary(base_df, ob_df)

        st.success(f"✅ 3개 파일 로드 완료 — FltReport {len(flt_df):,}행 · 승무원 {len(base_df)}명")

        selected_str = st.selectbox("📅 정산할 월 선택", [str(p) for p in available])
        sel = pd.Period(selected_str, freq="M")
        target_year, target_month = sel.year, sel.month

        with st.expander("ℹ️ 수당 계산 기준"):
            st.markdown("""
| 항목 | 기준 |
|------|------|
| **야간** | 운항 C/I~C/O(KST) 중 22:00~06:00+1 겹치는 시간 |
| **연장** | ATD~ATA(KST) 8시간 초과분 |
| **3P** | 편명 0151·0152 BI Hrs |
| **총비행시간** | Block + DHC×50% - OBCA/OBFO |
| **DAYOFF 미달** | 월 DAYOFF 8회 미만 |
            """)

        if st.button("🚀 정산 실행", type="primary", use_container_width=True):
            with st.spinner("계산 중..."):
                flt_summary, flt_detail = process_flt(flt_df, target_month, target_year)

            if flt_summary.empty:
                st.warning("FltReport에서 해당 월 데이터가 없습니다.")
            else:
                st.success(f"✅ {target_year}년 {target_month}월 — 수당 {len(flt_summary)}명 · 총비행시간 {len(calc_df)}명 정산 완료")

                tab1, tab2, tab3, tab4 = st.tabs(["📊 수당 요약","📋 개인별 상세","⏱ 총비행시간","📅 DAYOFF 미달"])

                with tab1:
                    disp=flt_summary.copy()
                    for k in ["야간","연장","P3"]: disp[k]=disp[k].apply(fmt_hhmm)
                    disp.index+=1; disp.columns=["이름","야간 시간","연장 시간","3P 시간"]
                    st.dataframe(disp, use_container_width=True, height=380)
                    c1,c2,c3=st.columns(3)
                    c1.metric("야간 합계", fmt_hhmm(flt_summary["야간"].sum()))
                    c2.metric("연장 합계", fmt_hhmm(flt_summary["연장"].sum()))
                    c3.metric("3P 합계",   fmt_hhmm(flt_summary["P3"].sum()))

                with tab2:
                    name_list=["전체"]+sorted(flt_detail["이름"].unique().tolist())
                    sel_name=st.selectbox("승무원 선택", name_list)
                    view=flt_detail if sel_name=="전체" else flt_detail[flt_detail["이름"]==sel_name]
                    disp2=view.copy()
                    for col in ["야간(h)","연장(h)","3P(h)"]: disp2[col]=disp2[col].apply(fmt_hhmm)
                    st.dataframe(disp2.reset_index(drop=True), use_container_width=True, height=380)
                    if sel_name!="전체":
                        c1,c2,c3=st.columns(3)
                        c1.metric("야간", fmt_hhmm(view["야간(h)"].sum()))
                        c2.metric("연장", fmt_hhmm(view["연장(h)"].sum()))
                        c3.metric("3P",   fmt_hhmm(view["3P(h)"].sum()))

                with tab3:
                    disp3=calc_df.sort_values("Crew Code").reset_index(drop=True).copy()
                    disp3["Block"]=disp3["Block"].apply(fmt_hhmm)
                    disp3["DHC"]=disp3["DHC"].apply(fmt_hhmm)
                    disp3["DHC_50"]=disp3["DHC_50"].apply(fmt_hhmm)
                    disp3["OBCA_OBFO"]=disp3["OBCA_OBFO"].apply(fmt_hhmm)
                    disp3["Total_Flt"]=disp3["Total_Flt"].apply(fmt_hhmm)
                    disp3=disp3[["Crew Code","Block","DHC","DHC_50","OBCA_OBFO","Total_Flt","Dayoff"]]
                    disp3.columns=["이름","Block(h)","DHC(h)","DHC×50%(h)","OBCA/OBFO(h)","총비행시간(h)","DAYOFF"]
                    disp3.index+=1
                    st.dataframe(disp3, use_container_width=True, height=380)
                    c1,c2=st.columns(2)
                    c1.metric("총비행시간 합계", fmt_hhmm(calc_df["Total_Flt"].sum()))
                    c2.metric("OBCA/OBFO 제거 합계", fmt_hhmm(calc_df["OBCA_OBFO"].sum()))

                with tab4:
                    under8=calc_df[calc_df["Dayoff_Under8"]].sort_values("Crew Code").reset_index(drop=True)
                    if len(under8)==0:
                        st.success("8회 미만자 없음 ✅")
                    else:
                        st.error(f"⚠️ DAYOFF 8회 미만자 총 {len(under8)}명")
                        disp4=under8[["Crew Code","Dayoff"]].copy()
                        disp4.columns=["이름","DAYOFF 횟수"]
                        disp4.index+=1
                        st.dataframe(disp4, use_container_width=True, height=300)

                # 다운로드
                excel_buf = build_excel(flt_summary, flt_detail, calc_df, target_year, target_month)
                fname = f"{target_year}년{target_month:02d}월_운항정산.xlsx"
                st.download_button(
                    label="⬇️ 엑셀 다운로드 (수당요약 + 개인별상세 + 총비행시간 + DAYOFF미달)",
                    data=excel_buf, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"오류 발생: {e}"); st.exception(e)
else:
    st.info("위 3개 파일을 모두 업로드해주세요.")
    missing_files = []
    if not flt_file: missing_files.append("FltReport.xlsx")
    if not dhc_file: missing_files.append("월DHC_DAYOFF_총비행시간.xlsx")
    if not ob_file:  missing_files.append("OBCA.xlsx")
    if missing_files:
        st.warning(f"미업로드 파일: {', '.join(missing_files)}")
