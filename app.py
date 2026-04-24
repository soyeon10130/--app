import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────
# 계산 로직
# ─────────────────────────────────────────
KST_OFFSET = timedelta(hours=9)

def combine_date_time_utc(date_str, t):
    if pd.isna(t) or t is None:
        return None
    base = datetime.strptime(date_str, "%d%b%y")
    return datetime(base.year, base.month, base.day, t.hour, t.minute, t.second)

def calc_night_hours(ci_kst, co_kst):
    """22:00~06:00+1 구간과 CI~CO 겹치는 시간(h) 합산"""
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
    """ATD~ATA 중 8시간 초과분(h)"""
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

def decimal_to_hhmm(h):
    hours = int(h)
    mins = round((h - hours) * 60)
    if mins == 60:
        hours += 1
        mins = 0
    return f"{hours:02d}:{mins:02d}"

def process(df, target_month, target_year):
    df = df.copy()
    df["date_parsed"] = pd.to_datetime(df["Date"], format="%d%b%y")

    rows = []
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

        rows.append({
            "name":      str(crew_str).strip(),
            "night_hrs": calc_night_hours(ci_kst, co_kst),
            "ot_hrs":    calc_overtime_hours(atd_kst, ata_kst),
            "p3_hrs":    bl_hrs if is_3p else 0.0,
        })

    if not rows:
        return pd.DataFrame(columns=["이름", "야간 시간(h)", "연장 시간(h)", "3P 시간(h)"])

    exploded = []
    for r in rows:
        for name in r["name"].split():
            name = name.strip()
            if name:
                exploded.append({"이름": name, "night": r["night_hrs"],
                                 "ot": r["ot_hrs"], "p3": r["p3_hrs"]})

    exp_df = pd.DataFrame(exploded)
    summary = (
        exp_df.groupby("이름")
        .agg(야간=("night", "sum"), 연장=("ot", "sum"), P3=("p3", "sum"))
        .reset_index()
        .sort_values("이름")
        .reset_index(drop=True)
    )
    return summary

def build_excel(summary, target_year, target_month):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{target_year}년 {target_month}월 수당정산"

    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    hdr_fill  = PatternFill("solid", fgColor="1F4E79")
    tot_fill  = PatternFill("solid", fgColor="FFF2CC")
    tot_font  = Font(name="Arial", bold=True, size=10)
    cell_font = Font(name="Arial", size=10)
    center    = Alignment(horizontal="center", vertical="center")
    left      = Alignment(horizontal="left",   vertical="center")
    thin      = Side(style="thin", color="AAAAAA")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 제목
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value     = f"{target_year}년 {target_month}월 운항 수당 정산표"
    c.font      = Font(name="Arial", bold=True, size=14)
    c.alignment = center
    ws.row_dimensions[1].height = 30

    # 주석
    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value     = "※ 야간: 22:00~06:00(KST) | 연장: 일 8시간 초과 | 3P: YP151(0151)/YP152(0152)편 BI Hrs"
    c.font      = Font(name="Arial", italic=True, size=9, color="666666")
    c.alignment = left
    ws.row_dimensions[2].height = 16

    # 헤더
    for col, h in enumerate(["No", "이름", "야간 시간", "연장 시간", "3P 시간"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = border
    ws.row_dimensions[3].height = 22

    # 데이터
    for i, row in summary.iterrows():
        r = i + 4
        for col, val in enumerate(
            [i + 1, row["이름"],
             decimal_to_hhmm(row["야간"]),
             decimal_to_hhmm(row["연장"]),
             decimal_to_hhmm(row["P3"])], 1
        ):
            c = ws.cell(row=r, column=col, value=val)
            c.font = cell_font; c.border = border
            c.alignment = left if col == 2 else center

    # 합계
    tr = len(summary) + 4
    ws.merge_cells(f"A{tr}:B{tr}")
    for col in range(1, 6):
        c = ws.cell(row=tr, column=col)
        c.font = tot_font; c.fill = tot_fill
        c.alignment = center; c.border = border
    ws["A" + str(tr)].value = "합계"
    for col, key in zip([3, 4, 5], ["야간", "연장", "P3"]):
        ws.cell(row=tr, column=col).value = decimal_to_hhmm(summary[key].sum())

    for i, w in enumerate([6, 14, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────
st.set_page_config(page_title="운항 수당 정산기", page_icon="✈️", layout="centered")

st.title("✈️ 운항 수당 정산기")
st.caption("FltReport.xlsx 업로드 → 월 선택 → 정산 실행 → 엑셀 다운로드")

st.divider()

# ① 파일 업로드
uploaded = st.file_uploader("📂 FltReport.xlsx 업로드", type=["xlsx"])

if uploaded:
    try:
        df = pd.read_excel(uploaded)
        required = {"Date", "Flight", "ATD", "ATA", "Bl Hrs", "운항 C/I", "운항 C/O", "운항 Crew"}
        missing = required - set(df.columns)
        if missing:
            st.error(f"필수 열이 없습니다: {missing}")
            st.stop()

        df["date_parsed"] = pd.to_datetime(df["Date"], format="%d%b%y")
        available = (
            df[df["운항 Crew"].notna()]["date_parsed"]
            .dt.to_period("M")
            .unique()
        )
        available_sorted = sorted(available, reverse=True)

        st.success(f"파일 로드 완료 — 총 {len(df):,}행 · {len(available_sorted)}개월 데이터 감지")

        # ② 월 선택
        col1, col2 = st.columns([2, 1])
        with col1:
            month_options = [str(p) for p in available_sorted]
            selected_str  = st.selectbox("📅 정산할 월 선택", month_options)
        with col2:
            st.write("")  # 높이 맞춤

        sel_period = pd.Period(selected_str, freq="M")
        target_year  = sel_period.year
        target_month = sel_period.month

        # ③ 요약 미리보기 (Expander)
        with st.expander("📋 원본 데이터 미리보기 (상위 10행)"):
            st.dataframe(df.head(10), use_container_width=True)

        # ④ 정산 실행
        if st.button("🚀 정산 실행", type="primary", use_container_width=True):
            with st.spinner("계산 중..."):
                summary = process(df, target_month, target_year)

            if summary.empty:
                st.warning("해당 월의 데이터가 없습니다.")
            else:
                st.success(f"✅ {target_year}년 {target_month}월 — 총 **{len(summary)}명** 정산 완료")

                # 표시용 데이터프레임
                display = summary.copy()
                display["야간"] = summary["야간"].apply(decimal_to_hhmm)
                display["연장"] = summary["연장"].apply(decimal_to_hhmm)
                display["P3"]   = summary["P3"].apply(decimal_to_hhmm)
                display.index  += 1
                display.columns = ["이름", "야간 시간", "연장 시간", "3P 시간"]
                st.dataframe(display, use_container_width=True, height=400)

                # 소계 메트릭
                c1, c2, c3 = st.columns(3)
                c1.metric("야간 합계", decimal_to_hhmm(summary["야간"].sum()))
                c2.metric("연장 합계", decimal_to_hhmm(summary["연장"].sum()))
                c3.metric("3P 합계",   decimal_to_hhmm(summary["P3"].sum()))

                # ⑤ 엑셀 다운로드
                excel_buf = build_excel(summary, target_year, target_month)
                fname = f"{target_year}년{target_month:02d}월_수당정산.xlsx"
                st.download_button(
                    label="⬇️ 엑셀 다운로드",
                    data=excel_buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"오류 발생: {e}")
        st.exception(e)
else:
    st.info("왼쪽 위 또는 여기에 FltReport.xlsx 파일을 업로드해주세요.")

    with st.expander("ℹ️ 수당 계산 기준 보기"):
        st.markdown("""
| 수당 종류 | 계산 기준 |
|----------|----------|
| **야간 수당** | `운항 C/I` ~ `운항 C/O` (KST) 중 **22:00 ~ 06:00+1** 겹치는 시간 |
| **연장 수당** | `ATD` ~ `ATA` (KST) 구간이 **8시간 초과**하는 분량 |
| **3P 수당**  | 편명 **0151 / 0152** (YP151·YP152) 의 BI Hrs |
| **KST 변환** | UTC 열 모두 **+9시간** 적용, 자정 초과 구간 자동 보정 |
        """)
