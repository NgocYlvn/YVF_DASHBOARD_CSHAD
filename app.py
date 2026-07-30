
from pathlib import Path
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
BOOKING_TARGET = 500

NAV_ITEMS = [
    "Overview",
    "Customer Adoption",
    "Booking Performance",
    "Issues & Improvement",
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
        padding: 0.82rem 0.62rem;
        height: 142px;
        min-height: 142px;
        max-height: 142px;
        box-sizing: border-box;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(28, 54, 89, 0.05);
    }

    .kpi-label {
        color: var(--navy);
        font-size: 0.82rem;
        font-weight: 750;
        line-height: 1.2;
        min-height: 46px;
        display: flex;
        align-items: flex-start;
        justify-content: center;
        text-align: center;
    }

    .kpi-value {
        font-size: 1.78rem;
        font-weight: 800;
        line-height: 1.1;
        color: var(--blue);
        margin-top: 0.35rem;
        text-align: center;
    }

    .kpi-note {
        color: var(--muted);
        font-size: 0.71rem;
        line-height: 1.15;
        margin-top: 0.28rem;
        text-align: center;
    }

    .accent-orange .kpi-value { color: var(--orange); }
    .accent-green .kpi-value { color: var(--green); }
    .accent-amber .kpi-value { color: var(--amber); }
    .accent-red .kpi-value { color: var(--red); }

    div[data-testid="stDataFrame"] {
        border: 1px solid var(--line);
        border-radius: 10px;
        overflow: hidden;
    }

    div[data-testid="stAlert"] {
        margin-top: 0.4rem;
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
        "Customer_Feedback": 1,
        "User Issues": 1,
    }

    missing = [s for s in required if s not in excel.sheet_names]
    if missing:
        raise ValueError("Missing sheet(s): " + ", ".join(missing))

    data = {}
    for sheet, header_row in required.items():
        df = pd.read_excel(excel, sheet_name=sheet, header=header_row)
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
        "Customer_Feedback",
        "User Issues",
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
        st.code("\\n".join(detected))
    else:
        st.warning("Repository hiện không có file `.xlsx` hoặc `.xlsm`.")

    st.stop()



def safe_divide(numerator, denominator):
    if denominator in (0, None) or pd.isna(denominator):
        return 0.0
    return float(numerator) / float(denominator)


def format_percent(value, decimals=0):
    return f"{value * 100:.{decimals}f}%"


def kpi_card(label, value, note="", accent=""):
    st.markdown(
        f"""
        <div class="kpi-card {accent}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def prepare_data(data):
    customer = data["Customer_Volume"].copy()
    booking = data["Booking_Records"].copy()
    onboarded = data["Onboarded_Customers"].copy()
    proposals = data["Improvement Proposals"].copy()
    feedback = data["Customer_Feedback"].copy()
    issues = data["User Issues"].copy()

    # Remove summary rows from customer source.
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
        (feedback, "Feedback Date"),
        (issues, "Date"),
    ]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

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
    # Six executive KPI cards with equal width and height
    cols = st.columns(6, gap="small")
    with cols[0]:
        kpi_card("Eligible Customers", f"{metrics['eligible']}")
    with cols[1]:
        kpi_card(
            "Total Onboarded",
            f"{metrics['total_onboarded']}",
            f"Onboarding rate: {format_percent(metrics['overall_rate'])}",
        )
    with cols[2]:
        kpi_card(
            "New Onboarded (2026)",
            f"{metrics['new_onboarded']}",
            f"Onboarding rate: {format_percent(metrics['new_rate'])}",
        )
    with cols[3]:
        kpi_card(
            "Active Customers",
            f"{metrics['active']}",
            f"Activation rate: {format_percent(metrics['active_rate'])}",
            accent="accent-green",
        )
    with cols[4]:
        kpi_card(
            "YTD Bookings via YVF",
            f"{metrics['ytd_bookings']}",
            f"Target achievement: {format_percent(metrics['booking_achievement'], 1)}",
        )
    with cols[5]:
        kpi_card(
            f"{FY_LABEL} Targets",
            f"{metrics['onboarding_target']} / {metrics['booking_target']}",
            "Customers / Bookings",
            accent="accent-orange",
        )

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([1.0, 1.8], gap="medium")

    with left:
        st.markdown('<div class="section-title">PROGRESS TO TARGET</div>', unsafe_allow_html=True)
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
        st.markdown('<div class="section-title">BOOKING TREND (VIA YVF)</div>', unsafe_allow_html=True)

        # Fixed FY2026 reporting months: Jul-2026 to Mar-2027.
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
        monthly_target = BOOKING_TARGET / 12
        fig.add_hline(
            y=monthly_target,
            line_dash="dash",
            line_color="#ed6b21",
            annotation_text=f"Monthly target: {monthly_target:.1f}",
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

    # Show every customer listed in Onboarded_Customers, including customers with zero bookings.
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

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Eligible Customers", metrics["eligible"])
    with c2:
        kpi_card("Total Onboarded", metrics["total_onboarded"])
    with c3:
        kpi_card("Active Customers", metrics["active"], accent="accent-green")
    with c4:
        kpi_card("Pending Target", metrics["pending"], accent="accent-amber")

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
                "Total Volume": st.column_config.NumberColumn(
                    "Export HBL Volume", format="%d"
                ),
            },
        )

    with right:
        st.markdown('<div class="section-title">CUSTOMERS BY YVF STATUS</div>', unsafe_allow_html=True)
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
        fig.update_traces(marker_color="#0b63ce", textposition="outside")
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
            "Avg_Processing_Time": "Avg. Processing Time",
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
                "Avg. Processing Time",
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
            "Avg. Processing Time": st.column_config.NumberColumn(
                "Avg. Processing Time",
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

    if isinstance(date_range, tuple) and len(date_range) == 2:
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

    active_customers = int(filtered_booking["Customer Name"].nunique())
    booking_achievement = safe_divide(total_bookings, BOOKING_TARGET)

    c1, c2, c3, c4, c5 = st.columns(5, gap="small")
    with c1:
        kpi_card("Bookings in Selection", total_bookings)
    with c2:
        kpi_card("Active Customers", active_customers, accent="accent-green")
    with c3:
        kpi_card("Avg. Processing Time", f"{avg_processing:.1f} min")
    with c4:
        kpi_card("Fastest Processing Time", f"{fastest_processing:.1f} min")
    with c5:
        kpi_card(
            "Booking Achievement",
            format_percent(booking_achievement, 1),
            f"{total_bookings} / {BOOKING_TARGET} bookings",
            accent="accent-orange",
        )

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([1.55, 1])

    with left:
        st.markdown('<div class="section-title">BOOKING TREND BY DATE</div>', unsafe_allow_html=True)
        daily = (
            filtered_booking.groupby("Booking Date", as_index=False)["Bookings"]
            .sum()
            .sort_values("Booking Date")
        )
        fig = px.bar(daily, x="Booking Date", y="Bookings", text="Bookings")
        fig.update_traces(marker_color="#0b63ce", textposition="outside")
        standard_chart_layout(fig, 360)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with right:
        st.markdown('<div class="section-title">BOOKINGS BY MODE</div>', unsafe_allow_html=True)
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
            height=360,
            margin=dict(l=20, r=20, t=35, b=20),
            paper_bgcolor="white",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns(2)

    with left:
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

    with right:
        st.markdown('<div class="section-title">PROCESSING TIME BY DATE</div>', unsafe_allow_html=True)
        processing = (
            filtered_booking.groupby("Booking Date", as_index=False)[
                "Processing Time (min)"
            ]
            .mean()
            .sort_values("Booking Date")
        )
        fig = px.line(
            processing,
            x="Booking Date",
            y="Processing Time (min)",
            markers=True,
        )
        fig.update_traces(line_color="#ed6b21", marker_color="#ed6b21")
        standard_chart_layout(fig, 330)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">BOOKING RECORDS</div>', unsafe_allow_html=True)
    st.dataframe(
        filtered_booking[
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
        ].sort_values("Booking Date", ascending=False),
        hide_index=True,
        use_container_width=True,
        column_config={
            "Booking Date": st.column_config.DateColumn(format="DD-MMM-YYYY"),
            "Processing Time (min)": st.column_config.NumberColumn(format="%.1f"),
        },
    )


# ============================================================
# PAGE 4: ISSUES & IMPROVEMENT
# ============================================================
else:
    st.markdown("### Issues & Improvement")

    open_issues = int(
        issues["Status"].fillna("").astype(str).str.casefold().eq("open").sum()
    )
    completed_issues = int(
        issues["Status"].fillna("").astype(str).str.casefold().eq("completed").sum()
    )
    open_proposals = int(
        proposals["Status"].fillna("").astype(str).str.casefold().eq("open").sum()
    )
    completed_proposals = int(
        proposals["Status"].fillna("").astype(str).str.casefold().eq("completed").sum()
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Open User Issues", open_issues, accent="accent-red")
    with c2:
        kpi_card("Completed Issues", completed_issues, accent="accent-green")
    with c3:
        kpi_card("Open Proposals", open_proposals, accent="accent-amber")
    with c4:
        kpi_card("Completed Proposals", completed_proposals, accent="accent-green")

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns(2)

    with left:
        st.markdown('<div class="section-title">USER ISSUES BY CATEGORY</div>', unsafe_allow_html=True)
        issue_category = (
            issues.groupby("Category", as_index=False)
            .size()
            .rename(columns={"size": "Issues"})
            .sort_values("Issues", ascending=True)
        )
        fig = px.bar(
            issue_category,
            x="Issues",
            y="Category",
            orientation="h",
            text="Issues",
        )
        fig.update_traces(marker_color="#ed6b21", textposition="outside")
        standard_chart_layout(fig, 300)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with right:
        st.markdown('<div class="section-title">PROPOSALS BY PRIORITY</div>', unsafe_allow_html=True)
        proposal_priority = (
            proposals.groupby("Priority", as_index=False)
            .size()
            .rename(columns={"size": "Proposals"})
        )
        fig = px.pie(
            proposal_priority,
            names="Priority",
            values="Proposals",
            hole=0.55,
        )
        fig.update_traces(textinfo="label+value")
        fig.update_layout(
            height=300,
            margin=dict(l=20, r=20, t=35, b=20),
            paper_bgcolor="white",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br>", unsafe_allow_html=True)
    tab1, tab2, tab3 = st.tabs(
        ["User Issues", "Improvement Proposals", "Positive Feedback"]
    )

    with tab1:
        st.dataframe(
            issues[
                ["Date", "Customer", "Issue", "Category", "Status"]
            ].sort_values(["Status", "Date"], ascending=[False, False]),
            hide_index=True,
            use_container_width=True,
            column_config={
                "Date": st.column_config.DateColumn(format="DD-MMM-YYYY"),
                "Issue": st.column_config.TextColumn(width="large"),
            },
        )

    with tab2:
        st.dataframe(
            proposals[
                [
                    "Proposal Date",
                    "Submitted By",
                    "Category",
                    "Module",
                    "Improvement Proposal",
                    "Priority",
                    "Status",
                ]
            ].sort_values(["Status", "Priority"], ascending=[False, True]),
            hide_index=True,
            use_container_width=True,
            column_config={
                "Proposal Date": st.column_config.DateColumn(format="DD-MMM-YYYY"),
                "Improvement Proposal": st.column_config.TextColumn(width="large"),
            },
        )

    with tab3:
        st.dataframe(
            feedback[
                [
                    "Feedback Date",
                    "Customer",
                    "Positive Feedback",
                    "Category",
                    "Business Value",
                ]
            ].sort_values("Feedback Date", ascending=False),
            hide_index=True,
            use_container_width=True,
            column_config={
                "Feedback Date": st.column_config.DateColumn(format="DD-MMM-YYYY"),
                "Positive Feedback": st.column_config.TextColumn(width="large"),
            },
        )


st.markdown(
    '<div class="footer-note">YVF Adoption Dashboard – CS HAD | '
    'Data is calculated directly from the Excel workbook | Version 6.1</div>',
    unsafe_allow_html=True,
)
