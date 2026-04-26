import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KST_OFFSET = timedelta(hours=9)

def combine_date_time_utc(date_str, t):
    if pd.isna(t) or t is None:
        return None
    base = datetime.strptime(date_str, "%d%b%y")
    return datetime(base.year, base.month, base.day, t.hour, t.minute, t.second)

def calc_night_hours(ci_kst, co_kst):
    if ci_kst is None or co_kst is None:
        return 0.0
    total = 0.0
    ci_date = ci_kst.date()
    for d in [-1, 0, 1]:
        ns = datetime.combine(ci_date + timedelta(days=d), time(22, 0))
        ne = datetime.combine(ci_date + timedelta(days=d + 1), time(6, 0))
        os_ = max(ci_kst, ns)
        oe_ = min(co_kst, ne)
        if oe_ > os_:
            total += (oe_ - os_).total_seconds() / 3600
    return total

def calc_overtime_hours(atd_kst, ata_kst):
    if atd_kst is None or ata_kst is None:
        return 0.0
    dur = (ata_kst - atd_kst).total_seconds() / 3600
    return max(0.0, dur - 8)

def blhrs_to_decimal(t):
    if pd.isna(t) or t is None:
        return 0.0
    if isinstance(t, time):
        return t.hour + t.minute / 60 + t.second / 3600
    return 0.0

def fmt_hhmm(h):
    if h == 0:
        return "-"
    hours = int(h)
    mins = round((h - hours) * 60)
    if mins == 60:
        hours += 1; mins = 0
    return f"{hours:02d}:{mins:02d}"

def fmt_time(t):
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return "-"
    if isinstance(t, time):
        return t.strftime("%H:%M")
    if isinstance(t, datetime):
        return t.strftime("%H:%M")
    return str(t)

def process(df, target_month, target_year):
    summary_rows = []
    detail_rows = []

    for _, row in df.iterrows():
        crew_str = row["운항 Crew"]
        if pd.isna(crew_str):
            continue
        date_str = row["Date"]
        atd_utc = combine_date_time_utc(date_str, row["ATD"])
        ata_utc = combine_date_time_utc(date_str, row["ATA"])
        ci_utc  = combine_date_time_utc(date_str, row["운항 C/I"])
        co_utc  = combine_date_time_utc(date_str, row["운항 C/O"])
        if atd_utc is None:
            continue

        atd_kst = atd_utc + KST_OFFSET
        ata_kst = (ata_utc + KST_OFFSET) if ata_utc else None
        ci_kst  = (ci_utc  + KST_OFFSET) if ci_utc  else None
        co_kst  = (co_utc  + KST_OFFSET) if co_utc  else None

        if ata_kst and ata_kst < atd_kst:
            ata_kst += timedelta(days=1)
        if co_kst and ci_kst and co_kst < ci_kst:
            co_kst += timedelta(days=1)

        if atd_kst.month != target_month or atd_kst.year != target_year:
            continue

        flight  = str(row["Flight"])
        bl_hrs  = blhrs_to_decimal(row["Bl Hrs"])
        is_3p   = flight in ["0151", "0152"]
        night   = calc_night_hours(ci_kst, co_kst)
        ot      = calc_overtime_hours(atd_kst, ata_kst)
        p3      = bl_hrs if is_3p else 0.0
        route   = f"{row['From']}→{row['To']}"

        crews = [c.strip() for c in str(crew_str).split() if c.strip()]
        for name in crews:
            summary_rows.append({"이름": name, "night": night, "ot": ot, "p3": p3})
            detail_rows.append({
                "이름":       name,
                "날짜":       atd_kst.strftime("%m/%d"),
                "편명":       flight,
                "구간":       route,
                "ATD(KST)":  fmt_time(atd_kst),
                "ATA(KST)":  fmt_time(ata_kst),
                "CI(KST)":   fmt_time(ci_kst),
                "CO(KST)":   fmt_time(co_kst),
                "야간(h)":    night,
                "연장(h)":    ot,
                "3P(h)":      p3,
            })

    if not summary_rows:
        return pd.DataFrame(), pd.DataFrame()

    s_df = pd.DataFrame(summary_rows)
    summary = (
        s_df.groupby("이름")
        .agg(야간=("night","sum"), 연장=("ot","sum"), P3=("p3","sum"))
        .reset_index().sort_values("이름").reset_index(drop=True)
    )
    detail = (
        pd.DataFrame(detail_rows)
        .sort_values(["이름","날짜","편명"])
        .reset_index(drop=True)
    )
    return summary, detail

def border_cell(ws, row, col, value, font, fill, alignment, border):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font; c.fill = fill; c.alignment = alignment; c.border = border
    return c

def build_excel(summary, detail, target_year, target_month):
    wb = Workbook()

    thin   = Side(style="thin", color="BBBBBB")
    bd     = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center")
    hdr_f  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_bg = PatternFill("solid", fgColor="1F4E79")
    sub_f  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    sub_bg = PatternFill("solid", fgColor="2E75B6")
    tot_f  = Font(name="Arial", bold=True, size=10)
    tot_bg = PatternFill("solid", fgColor="FFF2CC")
    cel_f  = Font(name="Arial", size=10)
    no_fill= PatternFill(fill_type=None)

    label  = f"{target_year}년 {target_month}월"

    # ── 시트1: 요약 ──────────────────────────────
    ws1 = wb.active
    ws1.title = "요약"

    ws1.merge_cells("A1:E1")
    c = ws1["A1"]
    c.value = f"{label} 운항 수당 정산표 — 요약"
    c.font  = Font(name="Arial", bold=True, size=13)
    c.alignment = center
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:E2")
    c = ws1["A2"]
    c.value = "※ 야간: 22:00~06:00(KST) | 연장: 일 8시간 초과 | 3P: 편명 0151·0152"
    c.font  = Font(name="Arial", italic=True, size=9, color="777777")
    c.alignment = left

    for col, h in enumerate(["No","이름","야간 시간","연장 시간","3P 시간"], 1):
        c = ws1.cell(row=3, column=col, value=h)
        c.font=hdr_f; c.fill=hdr_bg; c.alignment=center; c.border=bd
    ws1.row_dimensions[3].height = 20

    for i, row in summary.iterrows():
        r = i + 4
        for col, val in enumerate([i+1, row["이름"], fmt_hhmm(row["야간"]),
                                    fmt_hhmm(row["연장"]), fmt_hhmm(row["P3"])], 1):
            c = ws1.cell(row=r, column=col, value=val)
            c.font=cel_f; c.border=bd
            c.alignment = left if col==2 else center

    tr = len(summary) + 4
    ws1.merge_cells(f"A{tr}:B{tr}")
    for col in range(1,6):
        c = ws1.cell(row=tr, column=col)
        c.font=tot_f; c.fill=tot_bg; c.alignment=center; c.border=bd
    ws1["A"+str(tr)].value = "합계"
    for col, key in zip([3,4,5],["야간","연장","P3"]):
        ws1.cell(row=tr, column=col).value = fmt_hhmm(summary[key].sum())

    for i,w in enumerate([5,14,11,11,11],1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A4"

    # ── 시트2: 개인별 상세 ───────────────────────
    ws2 = wb.create_sheet("개인별 상세")

    ws2.merge_cells("A1:K1")
    c = ws2["A1"]
    c.value = f"{label} 운항 수당 정산표 — 개인별 비행 상세"
    c.font  = Font(name="Arial", bold=True, size=13)
    c.alignment = center
    ws2.row_dimensions[1].height = 28

    hdrs = ["이름","날짜","편명","구간","ATD(KST)","ATA(KST)","CI(KST)","CO(KST)","야간(h)","연장(h)","3P(h)"]
    for col, h in enumerate(hdrs, 1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font=hdr_f; c.fill=hdr_bg; c.alignment=center; c.border=bd
    ws2.row_dimensions[2].height = 20

    # 이름별 그룹핑하여 소계 포함
    names = detail["이름"].unique()
    cur_row = 3
    fill_colors = ["FFFFFF", "F7FBFF"]
    color_idx = 0

    for name in sorted(names):
        grp = detail[detail["이름"]==name].reset_index(drop=True)
        row_fill = PatternFill("solid", fgColor=fill_colors[color_idx % 2])
        color_idx += 1

        for _, dr in grp.iterrows():
            for col, key in enumerate(["이름","날짜","편명","구간",
                                        "ATD(KST)","ATA(KST)","CI(KST)","CO(KST)",
                                        "야간(h)","연장(h)","3P(h)"], 1):
                val = dr[key]
                if key in ["야간(h)","연장(h)","3P(h)"]:
                    val = fmt_hhmm(val)
                c = ws2.cell(row=cur_row, column=col, value=val)
                c.font=cel_f; c.fill=row_fill; c.border=bd
                c.alignment = left if col in [1,4] else center
            cur_row += 1

        # 소계 행
        sub_vals = [name+" 소계", "", "", "", "", "", "", "",
                    fmt_hhmm(grp["야간(h)"].sum()),
                    fmt_hhmm(grp["연장(h)"].sum()),
                    fmt_hhmm(grp["3P(h)"].sum())]
        for col, val in enumerate(sub_vals, 1):
            c = ws2.cell(row=cur_row, column=col, value=val)
            c.font=sub_f; c.fill=sub_bg; c.alignment=center; c.border=bd
        cur_row += 1

    # 전체 합계
    for col in range(1,12):
        c = ws2.cell(row=cur_row, column=col)
        c.font=tot_f; c.fill=tot_bg; c.alignment=center; c.border=bd
    ws2.cell(row=cur_row, column=1).value = "전체 합계"
    ws2.cell(row=cur_row, column=9).value  = fmt_hhmm(detail["야간(h)"].sum())
    ws2.cell(row=cur_row, column=10).value = fmt_hhmm(detail["연장(h)"].sum())
    ws2.cell(row=cur_row, column=11).value = fmt_hhmm(detail["3P(h)"].sum())

    col_widths = [14,8,8,12,10,10,10,10,10,10,10]
    for i,w in enumerate(col_widths,1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

# ── Streamlit UI ─────────────────────────────
st.set_page_config(page_title="운항 수당 정산기", page_icon="✈️", layout="centered")
st.title("✈️ 운항 수당 정산기")
st.caption("FltReport.xlsx 업로드 → 월 선택 → 정산 실행 → 엑셀 다운로드 (요약 + 개인별 상세)")
st.divider()

uploaded = st.file_uploader("📂 FltReport.xlsx 업로드", type=["xlsx"])

if uploaded:
    try:
        df = pd.read_excel(uploaded)
        required = {"Date","Flight","ATD","ATA","Bl Hrs","운항 C/I","운항 C/O","운항 Crew"}
        missing = required - set(df.columns)
        if missing:
            st.error(f"필수 열이 없습니다: {missing}"); st.stop()

        df["date_parsed"] = pd.to_datetime(df["Date"], format="%d%b%y")
        available = sorted(
            df[df["운항 Crew"].notna()]["date_parsed"].dt.to_period("M").unique(),
            reverse=True
        )
        st.success(f"파일 로드 완료 — 총 {len(df):,}행 · {len(available)}개월 데이터 감지")

        selected_str = st.selectbox("📅 정산할 월 선택", [str(p) for p in available])
        sel = pd.Period(selected_str, freq="M")
        target_year, target_month = sel.year, sel.month

        with st.expander("📋 원본 데이터 미리보기 (상위 10행)"):
            st.dataframe(df.head(10), use_container_width=True)

        if st.button("🚀 정산 실행", type="primary", use_container_width=True):
            with st.spinner("계산 중..."):
                summary, detail = process(df, target_month, target_year)

            if summary.empty:
                st.warning("해당 월의 데이터가 없습니다.")
            else:
                st.success(f"✅ {target_year}년 {target_month}월 — 총 **{len(summary)}명** 정산 완료")

                tab1, tab2 = st.tabs(["📊 요약", "📋 개인별 상세"])

                with tab1:
                    disp = summary.copy()
                    disp["야간"] = summary["야간"].apply(fmt_hhmm)
                    disp["연장"] = summary["연장"].apply(fmt_hhmm)
                    disp["P3"]   = summary["P3"].apply(fmt_hhmm)
                    disp.index += 1
                    disp.columns = ["이름","야간 시간","연장 시간","3P 시간"]
                    st.dataframe(disp, use_container_width=True, height=400)
                    c1,c2,c3 = st.columns(3)
                    c1.metric("야간 합계", fmt_hhmm(summary["야간"].sum()))
                    c2.metric("연장 합계", fmt_hhmm(summary["연장"].sum()))
                    c3.metric("3P 합계",   fmt_hhmm(summary["P3"].sum()))

                with tab2:
                    name_list = ["전체"] + sorted(detail["이름"].unique().tolist())
                    sel_name = st.selectbox("승무원 선택", name_list)
                    view = detail if sel_name=="전체" else detail[detail["이름"]==sel_name]
                    disp2 = view.copy()
                    for col in ["야간(h)","연장(h)","3P(h)"]:
                        disp2[col] = disp2[col].apply(fmt_hhmm)
                    st.dataframe(disp2.reset_index(drop=True), use_container_width=True, height=420)
                    if sel_name != "전체":
                        c1,c2,c3 = st.columns(3)
                        c1.metric("야간", fmt_hhmm(view["야간(h)"].sum()))
                        c2.metric("연장", fmt_hhmm(view["연장(h)"].sum()))
                        c3.metric("3P",   fmt_hhmm(view["3P(h)"].sum()))

                excel_buf = build_excel(summary, detail, target_year, target_month)
                fname = f"{target_year}년{target_month:02d}월_수당정산.xlsx"
                st.download_button(
                    label="⬇️ 엑셀 다운로드 (요약 + 개인별 상세 시트)",
                    data=excel_buf, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"오류 발생: {e}"); st.exception(e)
else:
    st.info("FltReport.xlsx 파일을 업로드해주세요.")
    with st.expander("ℹ️ 수당 계산 기준"):
        st.markdown("""
| 수당 | 기준 열 | 로직 |
|------|--------|------|
| **야간** | 운항 C/I → C/O (KST) | 22:00~06:00+1 겹치는 시간 |
| **연장** | ATD → ATA (KST) | 8시간 초과분 |
| **3P**  | Bl Hrs | 편명 0151·0152만 해당 |
        """)
