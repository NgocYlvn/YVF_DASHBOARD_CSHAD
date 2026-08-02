from pathlib import Path
import html
import io
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# APP CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="YVF Adoption Dashboard - CS HAD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "YVF Adoption Dashboard – CS HAD"
DEFAULT_DATA_FILE = "YVF_Adoption_Data.xlsx"
FY_LABEL = "FY2026"
ONBOARDING_TARGET = 6
BOOKING_TARGET = 800

DAILY_BOOKING_TITLE = "DAILY BOOKING VOLUME"
MONTHLY_BOOKING_TITLE = "MONTHLY BOOKING VOLUME"

NAV_ITEMS = [
    "Overview",
    "Customer Adoption",
    "Booking Performance",
    "User Issues & Improvements",
]


# ============================================================
# STYLE
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --navy: #083B82;
        --blue: #0B63CE;
        --orange: #ED6B21;
        --green: #169B62;
        --amber: #F59E0B;
        --red: #DC2626;
        --text: #172033;
        --muted: #667085;
        --line: #DCE5F0;
        --panel: #FFFFFF;
        --page: #F7F9FC;
    }

    html, body, [class*="css"] {
        font-family: Arial, "Segoe UI", sans-serif;
    }

    .stApp {
        background: var(--page);
        color: var(--text);
    }

    [data-testid="stHeader"] {
        height: 3.25rem;
        background: var(--page);
    }

    [data-testid="stToolbar"] {
        top: 0.35rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #073472 0%, #0B4D9B 100%);
    }

    [data-testid="stSidebar"] * {
        color: #FFFFFF;
    }

    [data-testid="stSidebar"] .stRadio label {
        font-weight: 600;
    }

    .block-container {
        max-width: 1600px;
        padding-top: 1.65rem;
        padding-bottom: 2rem;
    }

    .dashboard-title {
        display: block;
        position: static;
        font-size: 1.9rem;
        line-height: 1.15;
        font-weight: 800;
        color: #083B82 !important;
        opacity: 1 !important;
        filter: none !important;
        text-shadow: none !important;
        -webkit-text-fill-color: #083B82 !important;
        margin: 0 0 0.35rem 0;
        letter-spacing: -0.02em;
    }

    .dashboard-subtitle {
        display: block;
        position: relative;
        z-index: 5;
        color: #667085 !important;
        opacity: 1 !important;
        font-size: 0.82rem;
        line-height: 1.2;
        margin: 0 0 0.75rem 0;
    }

    .section-title {
        background: var(--navy);
        color: #FFFFFF;
        padding: 0.52rem 0.8rem;
        border-radius: 10px 10px 0 0;
        font-weight: 750;
        margin-top: 0.25rem;
    }

    .kpi-card {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 12px;
        height: 142px;
        min-height: 142px;
        max-height: 142px;
        display:flex;
        flex-direction:column;
        justify-content:center;
        align-items:center;
        box-sizing:border-box;
        overflow:hidden;
        box-shadow:0 2px 10px rgba(28,54,89,.05);
    }

    .kpi-label {
        color: var(--navy);
        font-size:0.88rem;
        font-weight:700;
        text-align:center;
        margin-bottom:12px;
    }

    .kpi-value {
        font-size:2.25rem;
        font-weight:800;
        line-height:1;
        color:var(--blue);
        text-align:center;
        margin:0;
    }

    .kpi-note {
        color: var(--muted);
        font-size: 0.71rem;
        line-height: 1.2;
        margin-top: 0.48rem;
        text-align: center;
        padding: 0 0.4rem;
    }

    .accent-orange .kpi-value { color: var(--orange); }
    .accent-green .kpi-value { color: var(--green); }
    .accent-amber .kpi-value { color: var(--amber); }
    .accent-red .kpi-value { color: var(--red); }

    .target-card {
        padding: 0 12px;
    }

    .target-card .kpi-label {
        margin-bottom: 14px;
    }

    .target-grid {
        width: 100%;
        display: grid;
        grid-template-columns: minmax(0, 1fr) 1px minmax(0, 1fr);
        align-items: center;
        justify-items: stretch;
        column-gap: 12px;
    }

    .target-item {
        min-width: 0;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
    }

    .target-value {
        color: var(--orange);
        font-size: 1.9rem;
        font-weight: 800;
        line-height: 1;
        margin: 0;
    }

    .target-caption {
        color: var(--muted);
        font-size: 0.72rem;
        line-height: 1.15;
        margin-top: 0.4rem;
        white-space: nowrap;
    }

    .target-divider {
        width: 1px;
        height: 52px;
        background: var(--line);
        justify-self: center;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid var(--line);
        border-radius: 10px;
        overflow: hidden;
    }

    div[data-testid="stAlert"] {
        margin-top: 0.4rem;
    }

    .wrapped-table-container {
        width: 100%;
        overflow-x: auto;
        border: 1px solid var(--line);
        border-radius: 10px;
        background: #FFFFFF;
    }

    .wrapped-data-table {
        width: 100%;
        min-width: 1180px;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 0.78rem;
        color: var(--text);
    }

    .wrapped-data-table th {
        background: #F7F9FC;
        color: var(--muted);
        font-weight: 500;
        text-align: left;
        padding: 0.65rem 0.7rem;
        border-right: 1px solid var(--line);
        border-bottom: 1px solid var(--line);
        vertical-align: middle;
        white-space: normal;
        overflow-wrap: anywhere;
    }

    .wrapped-data-table td {
        padding: 0.68rem 0.7rem;
        border-right: 1px solid #E4EAF2;
        border-bottom: 1px solid #E4EAF2;
        vertical-align: top;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: normal;
        line-height: 1.45;
    }

    .wrapped-data-table th:last-child,
    .wrapped-data-table td:last-child {
        border-right: none;
    }

    .wrapped-data-table tbody tr:last-child td {
        border-bottom: none;
    }

    .wrapped-data-table tbody tr:nth-child(even) {
        background: #FBFCFE;
    }

    .footer-note {
        color: var(--muted);
        font-size: 0.76rem;
        padding-top: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        re.sub(r"\s+", " ", str(c).replace("\n", " ")).strip()
        for c in df.columns
    ]
    return df


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    return df.dropna(how="all").reset_index(drop=True)


def normalize_status(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    required = {
        "Customer_Volume": 2,
        "Booking_Records": 1,
        "Onboarded_Customers": 1,
        "Improvement Proposals": 1,
        "Feedback": 1,
        "Dashboard_Overview": None,
    }

    missing = [s for s in required if s not in excel.sheet_names]
    if missing:
        raise ValueError("Missing sheet(s): " + ", ".join(missing))

    data = {}
    for sheet, header_row in required.items():
        if sheet == "Dashboard_Overview":
            data[sheet] = pd.read_excel(
                excel,
                sheet_name=sheet,
                header=None,
            )
        else:
            df = pd.read_excel(
                excel,
                sheet_name=sheet,
                header=header_row,
            )
            data[sheet] = drop_empty_rows(clean_columns(df))
    return data


def read_source_file(uploaded_file=None):
    """Automatically locate the Excel source file deployed with the app."""
    app_dir = Path(__file__).resolve().parent

    preferred_names = [
        DEFAULT_DATA_FILE,
        "37efe85c-1e56-4030-86c8-b68b1fb857b5.xlsx",
        "YVF_Adoption_Dashboard_CS_HAD.xlsx",
        "YVF_Adoption_Dashboard_CS_HAD.xlsm",
    ]

    all_excel_files = [
        p for p in app_dir.rglob("*")
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and not p.name.startswith("~$")
    ]

    by_lower_name = {p.name.lower(): p for p in all_excel_files}
    for preferred_name in preferred_names:
        matched = by_lower_name.get(preferred_name.lower())
        if matched is not None:
            return matched.read_bytes(), matched.name

    required_sheets = {
        "Customer_Volume",
        "Booking_Records",
        "Onboarded_Customers",
        "Improvement Proposals",
        "Feedback",
        "Dashboard_Overview",
    }

    valid_candidates = []
    for candidate in sorted(
        all_excel_files,
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    ):
        try:
            sheet_names = pd.ExcelFile(candidate).sheet_names
            if required_sheets.issubset(set(sheet_names)):
                valid_candidates.append(candidate)
        except Exception:
            continue

    if valid_candidates:
        selected = valid_candidates[0]
        return selected.read_bytes(), selected.name

    st.error("Không tìm thấy file Excel dữ liệu hợp lệ trong GitHub Repository.")
    st.markdown(
        """
        **Cách khắc phục**

        1. Upload file `YVF_Adoption_Data.xlsx` vào cùng Repository với `app.py`.
        2. Đảm bảo file Excel có đủ các sheet dữ liệu.
        3. Vào Streamlit Community Cloud và chọn **Reboot app**.
        """
    )

    detected = [str(p.relative_to(app_dir)) for p in all_excel_files]
    if detected:
        st.warning("Đã tìm thấy file Excel nhưng cấu trúc sheet chưa đúng:")
        st.code("\n".join(detected))
    else:
        st.warning("Repository hiện không có file `.xlsx` hoặc `.xlsm`.")

    st.stop()


def safe_divide(numerator, denominator):
    if denominator in (0, None) or pd.isna(denominator):
        return 0.0
    return float(numerator) / float(denominator)


def format_percent(value, decimals=0):
    return f"{value * 100:.{decimals}f}%"


def overview_number(overview: pd.DataFrame, row: int, col: int, cell: str):
    """Read a numeric value from Dashboard_Overview."""
    value = pd.to_numeric(overview.iloc[row, col], errors="coerce")
    if pd.isna(value):
        raise ValueError(
            f"Dashboard_Overview!{cell} must contain a valid number."
        )
    return float(value)


def overview_rate(overview: pd.DataFrame, row: int, col: int, cell: str):
    """Read a percentage stored as either 40%/0.4 or the number 40."""
    value = overview_number(overview, row, col, cell)
    if value > 1:
        value = value / 100
    return value


def kpi_card(label, value, note="", accent=""):
    note_html = (
        f'<div class="kpi-note">{html.escape(str(note))}</div>'
        if note
        else ""
    )
    st.markdown(
        f"""
        <div class="kpi-card {accent}">
            <div class="kpi-label">{html.escape(str(label))}</div>
            <div class="kpi-value">{html.escape(str(value))}</div>
            {note_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def prepare_data(data):
    customer = data["Customer_Volume"].copy()
    booking = data["Booking_Records"].copy()
    onboarded = data["Onboarded_Customers"].copy()
    proposals = data["Improvement Proposals"].copy()
    feedback = data["Feedback"].copy()

    customer = customer[
        pd.to_numeric(customer.get("No."), errors="coerce").notna()
    ].copy()
    customer["YVF Status"] = normalize_status(customer["YVF Status"])
    customer["Customer Name"] = customer["Customer Name"].astype(str).str.strip()
    customer["Total Volume"] = pd.to_numeric(
        customer.get("Total Volume"), errors="coerce"
    ).fillna(0)

    booking["Booking Date"] = pd.to_datetime(
        booking["Booking Date"], errors="coerce"
    )
    booking["Bookings"] = pd.to_numeric(
        booking["Bookings"], errors="coerce"
    ).fillna(0)
    booking["Processing Time (min)"] = pd.to_numeric(
        booking["Processing Time (min)"], errors="coerce"
    )
    booking["YVF Used"] = normalize_status(booking["YVF Used"])
    booking["Status"] = normalize_status(booking["Status"])
    booking["Customer Name"] = booking["Customer Name"].astype(str).str.strip()
    booking["Transport Mode"] = booking["Transport Mode"].astype(str).str.strip()
    booking["Month Start"] = booking["Booking Date"].dt.to_period("M").dt.to_timestamp()
    booking["Month Label"] = booking["Booking Date"].dt.strftime("%b-%Y")

    onboarded["YVF Booking Status"] = normalize_status(
        onboarded["YVF Booking Status"]
    )
    onboarded["Customer Name"] = onboarded["Customer Name"].astype(str).str.strip()

    for df, date_col in [
        (proposals, "Proposal Date"),
        (feedback, "Date"),
    ]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    feedback["Feedback Type"] = normalize_status(
        feedback.get("Feedback Type", pd.Series(index=feedback.index, dtype=str))
    )
    feedback["Status"] = normalize_status(
        feedback.get("Status", pd.Series(index=feedback.index, dtype=str))
    )

    # Negative feedback is treated as a user issue for management follow-up.
    issues = feedback[
        feedback["Feedback Type"].str.casefold().eq("negative")
    ].copy()

    return customer, booking, onboarded, proposals, feedback, issues


def calculate_metrics(customer, booking, onboarded):
    eligible = int(
        (
            customer["YVF Status"].ne("")
            & customer["YVF Status"].str.casefold().ne("declined")
        ).sum()
    )
    total_onboarded = int(
        customer["YVF Status"].str.startswith("Account Approved", na=False).sum()
    )
    new_onboarded = int(
        customer["YVF Status"]
        .str.contains(r"Account Approved.*2026FY", case=False, regex=True, na=False)
        .sum()
    )
    active = int(
        onboarded["YVF Booking Status"].str.casefold().eq("fully booking").sum()
    )
    pending = max(ONBOARDING_TARGET - new_onboarded, 0)

    yvf_booking = booking[
        booking["YVF Used"].str.casefold().eq("yes")
    ].copy()
    ytd_bookings = int(yvf_booking["Bookings"].sum())
    avg_time = yvf_booking["Processing Time (min)"].mean()
    avg_time = 0 if pd.isna(avg_time) else round(float(avg_time))

    return {
        "eligible": eligible,
        "total_onboarded": total_onboarded,
        "overall_rate": safe_divide(total_onboarded, eligible),
        "new_onboarded": new_onboarded,
        "new_rate": safe_divide(new_onboarded, eligible),
        "active": active,
        "pending": pending,
        "ytd_bookings": ytd_bookings,
        "avg_time": avg_time,
        "onboarding_target": ONBOARDING_TARGET,
        "booking_target": BOOKING_TARGET,
        "onboarding_achievement": safe_divide(new_onboarded, ONBOARDING_TARGET),
        "booking_achievement": safe_divide(ytd_bookings, BOOKING_TARGET),
        "active_rate": safe_divide(active, total_onboarded),
    }


def gauge_chart(value, title, detail, color="#ed6b21"):
    percentage = min(max(value * 100, 0), 100)
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=percentage,
            number={
                "suffix": "%",
                "font": {"size": 29, "color": "#ed6b21"},
            },
            title={
                "text": (
                    f"<b>{title}</b>"
                    f"<br><span style='font-size:11px;color:#667085'>{detail}</span>"
                ),
                "font": {"size": 14, "color": "#083b82"},
            },
            gauge={
                "shape": "angular",
                "axis": {
                    "range": [0, 100],
                    "tickwidth": 1,
                    "tickfont": {"size": 9},
                    "tickvals": [0, 50, 100],
                    "ticktext": ["0%", "50%", "100%"],
                },
                "bar": {"color": color, "thickness": 0.32},
                "bgcolor": "#e9eef5",
                "borderwidth": 0,
                "steps": [{"range": [0, 100], "color": "#e9eef5"}],
            },
        )
    )
    fig.update_layout(
        height=300,
        margin=dict(l=5, r=5, t=35, b=0),
        paper_bgcolor="white",
        font={"color": "#172033"},
    )
    return fig


def standard_chart_layout(fig, height=350):
    fig.update_layout(
        height=height,
        margin=dict(l=15, r=15, t=45, b=25),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#172033"),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
        hoverlabel=dict(bgcolor="white"),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#e9eef5")
    return fig


def render_wrapped_html_table(
    df: pd.DataFrame,
    column_widths: dict[str, str] | None = None,
    date_columns: list[str] | None = None,
):
    """Render a read-only HTML table with wrapped text and automatic row height."""
    if df.empty:
        st.info("No data available.")
        return

    display_df = df.copy()
    column_widths = column_widths or {}
    date_columns = date_columns or []

    for column in date_columns:
        if column in display_df.columns:
            display_df[column] = pd.to_datetime(
                display_df[column],
                errors="coerce",
            ).dt.strftime("%d-%b-%Y")
            display_df[column] = display_df[column].fillna("")

    display_df = display_df.fillna("")

    colgroup = "".join(
        f'<col style="width:{column_widths.get(str(column), "auto")};">'
        for column in display_df.columns
    )

    header_html = "".join(
        f"<th>{html.escape(str(column))}</th>"
        for column in display_df.columns
    )

    body_rows = []
    for _, row in display_df.iterrows():
        cells = "".join(
            "<td>"
            + html.escape(str(value)).replace("\n", "<br>")
            + "</td>"
            for value in row
        )
        body_rows.append(f"<tr>{cells}</tr>")

    table_html = f"""
    <div class="wrapped-table-container">
        <table class="wrapped-data-table">
            <colgroup>{colgroup}</colgroup>
            <thead><tr>{header_html}</tr></thead>
            <tbody>{''.join(body_rows)}</tbody>
        </table>
    </div>
    """

    st.markdown(table_html, unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.markdown("## 📊 YVF Dashboard")
st.sidebar.caption("CS HAD")

uploaded_file = None

page = st.sidebar.radio(
    "Navigation",
    NAV_ITEMS,
    index=0,
)


# ============================================================
# LOAD DATA
# ============================================================
try:
    source_bytes, source_name = read_source_file(uploaded_file)
    raw_data = load_workbook(source_bytes)
    customer, booking, onboarded, proposals, feedback, issues = prepare_data(raw_data)
    metrics = calculate_metrics(customer, booking, onboarded)

    # ========================================================
    # OVERVIEW KPI SOURCE: Dashboard_Overview!A4:J4
    # ========================================================
    overview = raw_data["Dashboard_Overview"]

    metrics["eligible"] = int(
        overview_number(overview, 3, 0, "A4")
    )
    metrics["total_onboarded"] = int(
        overview_number(overview, 3, 1, "B4")
    )
    metrics["overall_rate"] = overview_rate(
        overview, 3, 2, "C4"
    )
    metrics["new_onboarded"] = int(
        overview_number(overview, 3, 3, "D4")
    )
    metrics["new_rate"] = overview_rate(
        overview, 3, 4, "E4"
    )
    metrics["active"] = int(
        overview_number(overview, 3, 5, "F4")
    )
    metrics["pending"] = int(
        overview_number(overview, 3, 5, "F4")
    )
    metrics["ytd_bookings"] = int(
        overview_number(overview, 3, 6, "G4")
    )
    metrics["avg_time"] = overview_number(
        overview, 3, 7, "H4"
    )
    metrics["onboarding_target"] = int(
        overview_number(overview, 3, 8, "I4")
    )
    metrics["booking_target"] = int(
        overview_number(overview, 3, 9, "J4")
    )

    # Ratios used by other dashboard elements.
    metrics["onboarding_achievement"] = safe_divide(
        metrics["new_onboarded"],
        metrics["onboarding_target"],
    )
    metrics["booking_achievement"] = safe_divide(
        metrics["ytd_bookings"],
        metrics["booking_target"],
    )
    metrics["active_rate"] = safe_divide(
        metrics["active"],
        metrics["total_onboarded"],
    )
except Exception as exc:
    st.error("Không thể đọc dữ liệu nguồn của Dashboard.")
    st.exception(exc)
    st.info(
        "Vui lòng kiểm tra file Excel đã được upload lên GitHub, "
        "đúng tên sheet và không bị đặt mật khẩu."
    )
    st.stop()

latest_booking_date = booking["Booking Date"].max()
data_date = (
    latest_booking_date.strftime("%d %b %Y")
    if pd.notna(latest_booking_date)
    else "Not available"
)

st.markdown("<div style='height:0.15rem'></div>", unsafe_allow_html=True)
st.markdown(
    f'<div class="dashboard-title">{APP_TITLE}</div>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="dashboard-subtitle">📅 Last Updated: {data_date}</div>',
    unsafe_allow_html=True,
)


# ============================================================
# PAGE 1: OVERVIEW
# ============================================================
if page == "Overview":
    cols = st.columns(5, gap="small")
    with cols[0]:
        kpi_card("Eligible Customers", f"{metrics['eligible']}")
    with cols[1]:
        kpi_card(
            "Total Onboarded",
            f"{metrics['total_onboarded']}",
        )
    with cols[2]:
        kpi_card(
            "FY2026 Onboarded",
            f"{metrics['new_onboarded']}",
        )
    with cols[3]:
        kpi_card(
            "YTD Bookings",
            f"{metrics['ytd_bookings']}",
        )
    with cols[4]:
        targets_html = f"""
        <div class="kpi-card target-card">
            <div class="kpi-label">{FY_LABEL} Targets</div>
            <div class="target-grid">
                <div class="target-item">
                    <div class="target-value">{metrics["onboarding_target"]}</div>
                    <div class="target-caption">Customers</div>
                </div>
                <div class="target-divider"></div>
                <div class="target-item">
                    <div class="target-value">{metrics["booking_target"]}</div>
                    <div class="target-caption">Bookings</div>
                </div>
            </div>
        </div>
        """
        st.markdown(targets_html, unsafe_allow_html=True)


    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([1.0, 1.8], gap="medium")

    with left:
        st.markdown('<div class="section-title">FY2026 ACTUAL VS TARGET</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2, gap="small")
        with c1:
            st.plotly_chart(
                gauge_chart(
                    metrics["onboarding_achievement"],
                    "Onboarding",
                    f"{metrics['new_onboarded']} / {metrics['onboarding_target']} customers",
                ),
                use_container_width=True,
                config={"displayModeBar": False},
            )
        with c2:
            st.plotly_chart(
                gauge_chart(
                    metrics["booking_achievement"],
                    "Bookings",
                    f"{metrics['ytd_bookings']} / {metrics['booking_target']} bookings",
                ),
                use_container_width=True,
                config={"displayModeBar": False},
            )

    with right:
        st.markdown('<div class="section-title">BOOKING TREND</div>', unsafe_allow_html=True)

        month_axis = pd.date_range("2026-07-01", "2027-03-01", freq="MS")
        monthly_actual = (
            booking[booking["YVF Used"].str.casefold().eq("yes")]
            .groupby("Month Start", as_index=False)["Bookings"]
            .sum()
        )
        monthly = pd.DataFrame({"Month Start": month_axis}).merge(
            monthly_actual,
            on="Month Start",
            how="left",
        )
        monthly["Bookings"] = monthly["Bookings"].fillna(0).astype(int)
        monthly["Month Label"] = monthly["Month Start"].dt.strftime("%b %y")

        fig = px.bar(
            monthly,
            x="Month Label",
            y="Bookings",
            text="Bookings",
            category_orders={"Month Label": monthly["Month Label"].tolist()},
        )
        fig.update_traces(
            marker_color="#0b63ce",
            textposition="outside",
            cliponaxis=False,
        )
        monthly_target = round(metrics["booking_target"] / 12)
        fig.add_hline(
            y=monthly_target,
            line_dash="dash",
            line_color="#ed6b21",
            annotation_text=f"Monthly target: {monthly_target}",
            annotation_position="top right",
        )
        standard_chart_layout(fig, 300)
        fig.update_xaxes(
            categoryorder="array",
            categoryarray=monthly["Month Label"].tolist(),
            tickangle=0,
            tickfont={"size": 11},
        )
        fig.update_yaxes(rangemode="tozero")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">BOOKINGS BY ONBOARDED CUSTOMER</div>', unsafe_allow_html=True)

    onboarded_names = (
        onboarded[["Customer Name"]]
        .dropna()
        .assign(**{"Customer Name": lambda x: x["Customer Name"].astype(str).str.strip()})
        .drop_duplicates()
    )
    booking_volume = (
        booking[booking["YVF Used"].str.casefold().eq("yes")]
        .groupby("Customer Name", as_index=False)["Bookings"]
        .sum()
    )
    all_onboarded_booking = onboarded_names.merge(
        booking_volume,
        on="Customer Name",
        how="left",
    )
    all_onboarded_booking["Bookings"] = (
        all_onboarded_booking["Bookings"].fillna(0).astype(int)
    )
    all_onboarded_booking = all_onboarded_booking.sort_values(
        ["Bookings", "Customer Name"],
        ascending=[True, False],
    )

    if all_onboarded_booking.empty:
        st.info("No onboarded customer data available.")
    else:
        chart_height = max(300, 44 * len(all_onboarded_booking))
        fig = px.bar(
            all_onboarded_booking,
            x="Bookings",
            y="Customer Name",
            orientation="h",
            text="Bookings",
        )
        fig.update_traces(
            marker_color="#0b63ce",
            textposition="outside",
            cliponaxis=False,
        )
        standard_chart_layout(fig, chart_height)
        fig.update_yaxes(
            categoryorder="array",
            categoryarray=all_onboarded_booking["Customer Name"].tolist(),
            tickfont={"size": 11},
            automargin=True,
        )
        fig.update_xaxes(rangemode="tozero")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


# ============================================================
# PAGE 2: CUSTOMER ADOPTION
# ============================================================
elif page == "Customer Adoption":
    st.markdown("### Customer Adoption")

    status_options = sorted(customer["YVF Status"].dropna().unique().tolist())
    selected_status = st.multiselect(
        "Filter by YVF status",
        options=status_options,
        default=status_options,
    )

    filtered_customer = customer[customer["YVF Status"].isin(selected_status)].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Eligible Customers", metrics["eligible"])
    with c2:
        kpi_card("Total Onboarded", metrics["total_onboarded"])
    with c3:
        kpi_card("FY2026 Pending Onboard", metrics["pending"], accent="accent-amber")

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([1.6, 1])

    with left:
        st.markdown('<div class="section-title">CUSTOMER STATUS & VOLUME</div>', unsafe_allow_html=True)
        display_cols = ["Customer Name", "Total Volume", "YVF Status"]
        adoption_table = filtered_customer[display_cols].sort_values(
            "Total Volume", ascending=False
        )
        st.dataframe(
            adoption_table,
            hide_index=True,
            use_container_width=True,
            height=520,
            column_config={
                "Customer Name": st.column_config.TextColumn(
                    "Customer Name",
                    width="medium",
                ),
                "Total Volume": st.column_config.NumberColumn(
                    "Export HBL Volume",
                    format="%d",
                    width="small",
                ),
                "YVF Status": st.column_config.TextColumn(
                    "YVF Status",
                    width="medium",
                ),
            },
        )

    with right:
        st.markdown('<div class="section-title">CUSTOMER ONBOARDING STATUS</div>', unsafe_allow_html=True)
        status_count = (
            filtered_customer.groupby("YVF Status", as_index=False)
            .size()
            .rename(columns={"size": "Customers"})
            .sort_values("Customers", ascending=False)
        )
        fig = px.bar(
            status_count,
            x="Customers",
            y="YVF Status",
            orientation="h",
            text="Customers",
        )
        fig.update_traces(
            textposition="outside",
            cliponaxis=False,
        )
        standard_chart_layout(fig, 400)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">ONBOARDED CUSTOMER DETAILS</div>', unsafe_allow_html=True)

    booking_by_customer = (
        booking[booking["YVF Used"].str.casefold().eq("yes")]
        .groupby("Customer Name", as_index=False)
        .agg(
            YTD_Bookings=("Bookings", "sum"),
            Last_Booking=("Booking Date", "max"),
            Avg_Processing_Time=("Processing Time (min)", "mean"),
        )
    )

    onboard_detail = onboarded.merge(
        booking_by_customer,
        on="Customer Name",
        how="left",
    )
    onboard_detail["YTD_Bookings"] = onboard_detail["YTD_Bookings"].fillna(0).astype(int)
    onboard_detail["Avg_Processing_Time"] = onboard_detail[
        "Avg_Processing_Time"
    ].round(1)
    onboard_detail = onboard_detail.rename(
        columns={
            "YVF Booking Status": "Booking Status",
            "YTD_Bookings": "YTD Bookings",
            "Last_Booking": "Last Booking",
            "Avg_Processing_Time": "Avg. Processing Time per Booking",
        }
    )

    st.dataframe(
        onboard_detail[
            [
                "Customer Name",
                "Transport Mode",
                "Booking Status",
                "YTD Bookings",
                "Last Booking",
                "Avg. Processing Time per Booking",
                "Remarks",
            ]
        ],
        hide_index=True,
        use_container_width=True,
        height=430,
        column_config={
            "Customer Name": st.column_config.TextColumn(
                "Customer Name",
                width="medium",
            ),
            "Transport Mode": st.column_config.TextColumn(
                "Transport Mode",
                width="small",
            ),
            "Booking Status": st.column_config.TextColumn(
                "Booking Status",
                width="medium",
            ),
            "YTD Bookings": st.column_config.NumberColumn(
                "YTD Bookings",
                width="small",
                format="%d",
            ),
            "Last Booking": st.column_config.DateColumn(
                "Last Booking",
                width="small",
                format="DD-MMM-YYYY",
            ),
            "Avg. Processing Time per Booking": st.column_config.NumberColumn(
                "Avg. Processing Time per Booking",
                width="medium",
                format="%.1f min",
            ),
            "Remarks": st.column_config.TextColumn(
                "Remarks",
                width="large",
            ),
        },
    )


# ============================================================
# PAGE 3: BOOKING PERFORMANCE
# ============================================================
elif page == "Booking Performance":
    st.markdown("### Booking Performance")

    min_date = booking["Booking Date"].min()
    max_date = booking["Booking Date"].max()

    cfilter1, cfilter2, cfilter3 = st.columns([1.2, 1, 1])
    with cfilter1:
        date_range = st.date_input(
            "Booking date range",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )
    with cfilter2:
        mode_options = sorted(booking["Transport Mode"].dropna().unique().tolist())
        modes = st.multiselect(
            "Transport mode",
            mode_options,
            default=mode_options,
        )
    with cfilter3:
        customer_options = sorted(booking["Customer Name"].dropna().unique().tolist())
        customers = st.multiselect(
            "Customer",
            customer_options,
            default=customer_options,
        )

    filtered_booking = booking[
        booking["Transport Mode"].isin(modes)
        & booking["Customer Name"].isin(customers)
        & booking["YVF Used"].str.casefold().eq("yes")
    ].copy()

    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        start_date, end_date = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
        filtered_booking = filtered_booking[
            filtered_booking["Booking Date"].between(start_date, end_date)
        ]

    total_bookings = int(filtered_booking["Bookings"].sum())
    processing_series = filtered_booking["Processing Time (min)"].dropna()

    avg_processing = processing_series.mean()
    avg_processing = 0 if pd.isna(avg_processing) else round(float(avg_processing), 1)

    fastest_processing = processing_series.min()
    fastest_processing = 0 if pd.isna(fastest_processing) else round(float(fastest_processing), 1)

    slowest_processing = processing_series.max()
    slowest_processing = 0 if pd.isna(slowest_processing) else round(float(slowest_processing), 1)

    booking_achievement = safe_divide(total_bookings, metrics["booking_target"])

    c1, c2, c3, c4, c5 = st.columns(5, gap="small")
    with c1:
        kpi_card("Booking Volume", total_bookings)
    with c2:
        kpi_card(
            "Target Achievement",
            format_percent(booking_achievement, 1),
            accent="accent-orange",
        )
    with c3:
        kpi_card("Avg. Processing Time / Booking", f"{avg_processing:.1f} min")
    with c4:
        kpi_card("Fastest Processing Time / Booking", f"{fastest_processing:.1f} min")
    with c5:
        kpi_card(
            "Slowest Processing Time / Booking",
            f"{slowest_processing:.1f} min",
            accent="accent-red",
        )

    # ========================================================
    # ROW 1: TIME-BASED PERFORMANCE
    # ========================================================
    st.markdown("<br>", unsafe_allow_html=True)
    volume_col, processing_col = st.columns([1, 1], gap="medium")

    with volume_col:
        # Use the selected filter period to decide Daily or Monthly view.
        if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
            selected_start = pd.Timestamp(date_range[0]).normalize()
            selected_end = pd.Timestamp(date_range[1]).normalize()
        else:
            available_dates = filtered_booking["Booking Date"].dropna()
            if not available_dates.empty:
                selected_start = available_dates.min().normalize()
                selected_end = available_dates.max().normalize()
            else:
                selected_start = pd.NaT
                selected_end = pd.NaT

        selected_days = (
            (selected_end - selected_start).days + 1
            if pd.notna(selected_start) and pd.notna(selected_end)
            else 0
        )

        use_monthly = selected_days > 31

        if use_monthly:
            chart_title = MONTHLY_BOOKING_TITLE

            chart_data = (
                filtered_booking.assign(
                    Month_Start=filtered_booking["Booking Date"]
                    .dt.to_period("M")
                    .dt.to_timestamp()
                )
                .groupby("Month_Start", as_index=False)["Bookings"]
                .sum()
                .sort_values("Month_Start")
            )

            chart_data["Period Label"] = (
                chart_data["Month_Start"].dt.strftime("%b %Y")
            )
            custom_date_col = "Month_Start"
            tick_angle = -30 if len(chart_data) > 12 else 0
            hover_template = (
                "%{customdata[0]|%b %Y}"
                "<br>Bookings: %{y}"
                "<extra></extra>"
            )
        else:
            chart_title = DAILY_BOOKING_TITLE

            chart_data = (
                filtered_booking.groupby(
                    "Booking Date",
                    as_index=False,
                )["Bookings"]
                .sum()
                .sort_values("Booking Date")
            )

            chart_data["Booking Date"] = pd.to_datetime(
                chart_data["Booking Date"],
                errors="coerce",
            )
            chart_data = chart_data.dropna(subset=["Booking Date"])
            chart_data["Period Label"] = (
                chart_data["Booking Date"].dt.strftime("%d %b")
            )
            custom_date_col = "Booking Date"
            tick_angle = -45 if len(chart_data) > 15 else 0
            hover_template = (
                "%{customdata[0]|%d %b %Y}"
                "<br>Bookings: %{y}"
                "<extra></extra>"
            )

        st.markdown(
            f'<div class="section-title">{chart_title}</div>',
            unsafe_allow_html=True,
        )

        if chart_data.empty:
            st.info("No booking data available for the selected period.")
        else:
            fig = px.bar(
                chart_data,
                x="Period Label",
                y="Bookings",
                text="Bookings",
                custom_data=[custom_date_col],
                category_orders={
                    "Period Label": chart_data["Period Label"].tolist()
                },
            )

            fig.update_traces(
                marker_color="#0b63ce",
                textposition="outside",
                cliponaxis=False,
                hovertemplate=hover_template,
            )

            standard_chart_layout(fig, 350)

            fig.update_xaxes(
                type="category",
                categoryorder="array",
                categoryarray=chart_data["Period Label"].tolist(),
                tickangle=tick_angle,
                title_text="",
                tickfont={
                    "size": 10 if len(chart_data) > 20 else 11
                },
            )

            fig.update_yaxes(
                title_text="",
                rangemode="tozero",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )


    with processing_col:
        processing_title = (
            "MONTHLY AVERAGE PROCESSING TIME / BOOKING"
            if use_monthly
            else "DAILY AVERAGE PROCESSING TIME / BOOKING"
        )

        st.markdown(
            f'<div class="section-title">{processing_title}</div>',
            unsafe_allow_html=True,
        )

        processing_source = filtered_booking[
            (filtered_booking["Bookings"] > 0)
            & filtered_booking["Processing Time (min)"].notna()
            & filtered_booking["Booking Date"].notna()
        ].copy()

        processing_source["Weighted Processing Time"] = (
            processing_source["Processing Time (min)"]
            * processing_source["Bookings"]
        )

        if use_monthly:
            processing_source["Period Start"] = (
                processing_source["Booking Date"]
                .dt.to_period("M")
                .dt.to_timestamp()
            )
            period_format = "%b %Y"
            tick_angle = -30 if processing_source["Period Start"].nunique() > 12 else 0
            hover_date_format = "%b %Y"
        else:
            processing_source["Period Start"] = (
                processing_source["Booking Date"].dt.normalize()
            )
            period_format = "%d %b"
            tick_angle = -45 if processing_source["Period Start"].nunique() > 15 else 0
            hover_date_format = "%d %b %Y"

        processing = (
            processing_source
            .groupby("Period Start", as_index=False)
            .agg(
                Total_Weighted_Time=("Weighted Processing Time", "sum"),
                Booking_Volume=("Bookings", "sum"),
            )
            .sort_values("Period Start")
        )

        processing["Avg_Processing_Time"] = np.where(
            processing["Booking_Volume"] > 0,
            processing["Total_Weighted_Time"] / processing["Booking_Volume"],
            np.nan,
        )

        processing = processing.dropna(
            subset=["Period Start", "Avg_Processing_Time"]
        )

        processing["Period Label"] = (
            processing["Period Start"].dt.strftime(period_format)
        )

        if processing.empty:
            st.info(
                "No processing-time data available for the selected period."
            )
        else:
            fig = px.line(
                processing,
                x="Period Label",
                y="Avg_Processing_Time",
                markers=True,
                custom_data=[
                    "Period Start",
                    "Booking_Volume",
                ],
                category_orders={
                    "Period Label": processing["Period Label"].tolist()
                },
            )

            fig.update_traces(
                line_color="#ed6b21",
                marker_color="#ed6b21",
                line_width=2.5,
                marker_size=8,
                hovertemplate=(
                    f"%{{customdata[0]|{hover_date_format}}}"
                    "<br>Average processing time: %{y:.1f} min"
                    "<br>Bookings: %{customdata[1]:.0f}"
                    "<extra></extra>"
                ),
            )

            standard_chart_layout(fig, 350)

            fig.update_xaxes(
                type="category",
                categoryorder="array",
                categoryarray=processing["Period Label"].tolist(),
                tickangle=tick_angle,
                title_text="",
                showticklabels=True,
                showline=True,
                ticks="outside",
                tickfont={
                    "size": 10 if len(processing) > 20 else 11
                },
            )

            fig.update_yaxes(
                title_text="Minutes (min)",
                rangemode="tozero",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )



    # ========================================================
    # ROW 2: BOOKING BREAKDOWN
    # ========================================================
    st.markdown("<br>", unsafe_allow_html=True)
    customer_col, distribution_col = st.columns([1.4, 1], gap="medium")

    with customer_col:
        st.markdown('<div class="section-title">BOOKINGS BY CUSTOMER</div>', unsafe_allow_html=True)
        by_customer = (
            filtered_booking.groupby("Customer Name", as_index=False)["Bookings"]
            .sum()
            .sort_values("Bookings", ascending=True)
        )
        fig = px.bar(
            by_customer,
            x="Bookings",
            y="Customer Name",
            orientation="h",
            text="Bookings",
        )
        fig.update_traces(marker_color="#0b63ce", textposition="outside")
        standard_chart_layout(fig, 330)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


    with distribution_col:
        st.markdown('<div class="section-title">BOOKING DISTRIBUTION</div>', unsafe_allow_html=True)
        by_mode = (
            filtered_booking.groupby("Transport Mode", as_index=False)["Bookings"]
            .sum()
            .sort_values("Bookings", ascending=False)
        )
        fig = px.pie(
            by_mode,
            names="Transport Mode",
            values="Bookings",
            hole=0.55,
        )
        fig.update_traces(textposition="inside", textinfo="label+percent")
        fig.update_layout(
            height=330,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor="white",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">BOOKING RECORDS</div>',
        unsafe_allow_html=True,
    )

    booking_records_display = filtered_booking[
        [
            "Booking Date",
            "Customer Name",
            "Transport Mode",
            "Bookings",
            "Processing Time (min)",
            "Status",
            "Handled By",
            "Remarks",
        ]
    ].sort_values("Booking Date", ascending=False).copy()

    booking_records_display["Bookings"] = pd.to_numeric(
        booking_records_display["Bookings"],
        errors="coerce",
    ).fillna(0).astype(int)

    booking_records_display["Processing Time (min)"] = pd.to_numeric(
        booking_records_display["Processing Time (min)"],
        errors="coerce",
    ).map(lambda value: f"{value:.1f}" if pd.notna(value) else "")

    booking_records_display["Remarks"] = (
        booking_records_display["Remarks"]
        .replace({None: "", "None": "", np.nan: ""})
        .fillna("")
    )

    render_wrapped_html_table(
        booking_records_display,
        column_widths={
            "Booking Date": "12%",
            "Customer Name": "16%",
            "Transport Mode": "13%",
            "Bookings": "9%",
            "Processing Time (min)": "15%",
            "Status": "11%",
            "Handled By": "11%",
            "Remarks": "13%",
        },
        date_columns=["Booking Date"],
    )


# ============================================================
# PAGE 4: USER ISSUES & IMPROVEMENTS
# ============================================================
else:
    st.markdown("### User Issues & Improvements")
    st.caption(
        "Track user feedback, reported issues, and improvement actions."
    )

    # --------------------------------------------------------
    # PAGE-SPECIFIC STYLE
    # --------------------------------------------------------
    st.markdown(
        """
        <style>
        .issue-kpi-card {
            background: #FFFFFF;
            border: 1px solid #DCE5F0;
            border-radius: 12px;
            min-height: 128px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            box-shadow: 0 2px 10px rgba(28,54,89,.05);
        }

        .issue-kpi-label {
            color: #083B82;
            font-size: 0.82rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 0.55rem;
            text-align: center;
        }

        .issue-kpi-value {
            font-size: 2.35rem;
            font-weight: 850;
            line-height: 1;
        }

        .issue-kpi-note {
            color: #667085;
            font-size: 0.74rem;
            margin-top: 0.55rem;
            text-align: center;
        }

        .issue-total .issue-kpi-value { color: #0B63CE; }
        .issue-completed .issue-kpi-value { color: #169B62; }
        .issue-open .issue-kpi-value { color: #DC2626; }
        .issue-rate .issue-kpi-value { color: #ED6B21; }

        .feedback-summary {
            background: #FFFFFF;
            border: 1px solid #DCE5F0;
            border-radius: 10px;
            overflow: hidden;
            height: 334px;
            min-height: 334px;
            max-height: 334px;
            box-sizing: border-box;
            box-shadow: 0 2px 10px rgba(28,54,89,.05);
        }

        .feedback-header {
            background: #083B82;
            color: #FFFFFF;
            padding: 0.52rem 0.8rem;
            font-size: 0.95rem;
            font-weight: 800;
            letter-spacing: 0.02em;
            text-transform: uppercase;
        }

        .feedback-content {
            padding: 0.75rem 1.1rem 1rem;
        }

        .feedback-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 10px;
        }

        .feedback-box {
            border: 1px solid #E4EAF2;
            border-radius: 10px;
            min-height: 105px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            background: #FAFCFF;
        }

        .feedback-value {
            font-size: 1.85rem;
            font-weight: 850;
            line-height: 1;
        }

        .feedback-label {
            color: #667085;
            font-size: 0.74rem;
            margin-top: 0.45rem;
        }

        .feedback-total .feedback-value { color: #0B63CE; }
        .feedback-positive .feedback-value { color: #169B62; }
        .feedback-negative .feedback-value { color: #DC2626; }

        .record-bar {
            background: linear-gradient(90deg, #083B82 0%, #0B63CE 100%);
            color: #FFFFFF;
            padding: 0.64rem 0.9rem;
            border-radius: 9px 9px 0 0;
            font-weight: 800;
            letter-spacing: 0.03em;
            margin-top: 1rem;
        }

        div[data-testid="stTabs"] button {
            font-weight: 700;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # STANDARDIZE SOURCE DATA
    # --------------------------------------------------------
    proposals["Status"] = normalize_status(
        proposals.get(
            "Status",
            pd.Series(index=proposals.index, dtype=str),
        )
    )
    feedback["Feedback Type"] = normalize_status(
        feedback.get(
            "Feedback Type",
            pd.Series(index=feedback.index, dtype=str),
        )
    )
    feedback["Status"] = normalize_status(
        feedback.get(
            "Status",
            pd.Series(index=feedback.index, dtype=str),
        )
    )

    completed_statuses = {
        "completed",
        "closed",
        "resolved",
        "done",
        "implemented",
    }

    proposal_status = proposals["Status"].str.casefold()
    feedback_type = feedback["Feedback Type"].str.casefold()

    # --------------------------------------------------------
    # KPI CALCULATIONS
    # Improvement Proposals sheet represents the issue-to-action log:
    # User Issue -> Improvement Proposal -> Status.
    # --------------------------------------------------------
    total_issues = int(len(proposals))
    completed_issues = int(proposal_status.isin(completed_statuses).sum())
    open_issues = max(total_issues - completed_issues, 0)
    issue_completion_rate = safe_divide(completed_issues, total_issues)

    total_feedback = int(len(feedback))
    positive_feedback = int(feedback_type.eq("positive").sum())
    negative_feedback = int(feedback_type.eq("negative").sum())
    other_feedback = max(
        total_feedback - positive_feedback - negative_feedback,
        0,
    )

    # --------------------------------------------------------
    # USER ISSUE SUMMARY
    # --------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4, gap="small")

    def issue_kpi(column, label, value, note, css_class):
        with column:
            st.markdown(
                f"""
                <div class="issue-kpi-card {css_class}">
                    <div class="issue-kpi-label">{label}</div>
                    <div class="issue-kpi-value">{value}</div>
                    <div class="issue-kpi-note">{note}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    issue_kpi(
        k1,
        "Total User Issues",
        total_issues,
        "Issues identified from user experience",
        "issue-total",
    )
    issue_kpi(
        k2,
        "Completed",
        completed_issues,
        "Improvement completed",
        "issue-completed",
    )
    issue_kpi(
        k3,
        "Open",
        open_issues,
        "Improvement action required",
        "issue-open",
    )
    issue_kpi(
        k4,
        "Completion Rate",
        format_percent(issue_completion_rate, 0),
        f"{completed_issues} of {total_issues} issues completed",
        "issue-rate",
    )

    # --------------------------------------------------------
    # ISSUE DISTRIBUTION + USER FEEDBACK
    # --------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    issue_col, feedback_col = st.columns([1.6, 1], gap="small")

    with issue_col:
        st.markdown(
            '<div class="section-title">USER ISSUES BY CATEGORY</div>',
            unsafe_allow_html=True,
        )

        if "Category" in proposals.columns:
            category_count = (
                proposals.assign(
                    Category=proposals["Category"]
                    .fillna("Unclassified")
                    .astype(str)
                    .str.strip()
                    .replace("", "Unclassified")
                )
                .groupby("Category", as_index=False)
                .size()
                .rename(columns={"size": "Issues"})
                .sort_values(["Issues", "Category"], ascending=[True, False])
            )
        else:
            category_count = pd.DataFrame(
                {"Category": ["Unclassified"], "Issues": [total_issues]}
            )

        if category_count.empty:
            st.info("No user issue data available.")
        else:
            fig = px.bar(
                category_count,
                x="Issues",
                y="Category",
                orientation="h",
                text="Issues",
            )
            fig.update_traces(
                marker_color="#0B63CE",
                textposition="outside",
                cliponaxis=False,
                hovertemplate="%{y}: %{x} issue(s)<extra></extra>",
            )
            standard_chart_layout(fig, 294)
            fig.update_layout(
                margin=dict(l=15, r=15, t=5, b=25),
            )
            fig.update_xaxes(
                dtick=1,
                rangemode="tozero",
                title_text="Number of Issues",
            )
            fig.update_yaxes(
                categoryorder="array",
                categoryarray=category_count["Category"].tolist(),
                automargin=True,
            )
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},
            )

    with feedback_col:
        st.markdown(
            f"""
            <div class="feedback-summary">
                <div class="feedback-header">USER FEEDBACK SUMMARY</div>
                <div class="feedback-content">
                    <div class="feedback-grid">
                        <div class="feedback-box feedback-total">
                            <div class="feedback-value">{total_feedback}</div>
                            <div class="feedback-label">Total Feedback</div>
                        </div>
                        <div class="feedback-box feedback-positive">
                            <div class="feedback-value">{positive_feedback}</div>
                            <div class="feedback-label">Positive</div>
                        </div>
                        <div class="feedback-box feedback-negative">
                            <div class="feedback-value">{negative_feedback}</div>
                            <div class="feedback-label">Negative</div>
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,    
        )

    # --------------------------------------------------------
    # RECORD TABLES
    # --------------------------------------------------------
    st.markdown(
        '<div class="record-bar">▣ &nbsp; RECORD</div>',
        unsafe_allow_html=True,
    )

    tab_issues, tab_feedback = st.tabs(
        ["User Issue", "User Feedback"]
    )

    with tab_issues:
        issue_columns = [
            column
            for column in [
                "No.",
                "Proposal Date",
                "Submitted By",
                "Category",
                "Module",
                "Current User Issue",
                "User Impact",
                "Improvement Proposal",
                "Status",
            ]
            if column in proposals.columns
        ]

        issue_display = proposals[issue_columns].copy()
        issue_display = issue_display.rename(
            columns={
                "Proposal Date": "Reported Date",
                "Submitted By": "Reported By",
                "Current User Issue": "Current User Issue",
                "User Impact": "Current User Issue",
                "Improvement Proposal": "Proposed Improvement",
            }
        )

        if "Reported Date" in issue_display.columns:
            issue_display = issue_display.sort_values(
                "Reported Date",
                ascending=False,
            )

        render_wrapped_html_table(
            issue_display,
            column_widths={
                "No.": "5%",
                "Reported Date": "9%",
                "Reported By": "11%",
                "Category": "10%",
                "Module": "8%",
                "Current User Issue": "30%",
                "Proposed Improvement": "28%",
                "Status": "8%",
            },
            date_columns=["Reported Date"],
        )

    with tab_feedback:
        feedback_columns = [
            column
            for column in [
                "Date",
                "Reported By",
                "Feedback",
                "Feedback Type",
                "Status",
            ]
            if column in feedback.columns
        ]

        feedback_display = feedback[feedback_columns].copy()
        if "Date" in feedback_display.columns:
            feedback_display = feedback_display.sort_values(
                "Date",
                ascending=False,
            )

        feedback_display = feedback_display.rename(
            columns={
                "Feedback": "User Feedback",
            }
        )

        render_wrapped_html_table(
            feedback_display,
            column_widths={
                "Date": "12%",
                "Reported By": "15%",
                "User Feedback": "48%",
                "Feedback Type": "13%",
                "Status": "12%",
            },
            date_columns=["Date"],
        )

st.markdown(
    '<div class="footer-note">YVF Adoption Dashboard – CS HAD | '
    '© 2026 CS HAD | Internal Use Only | Version 7.0</div>',
    unsafe_allow_html=True,
)
