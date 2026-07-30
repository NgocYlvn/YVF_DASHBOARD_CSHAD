from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_TITLE = "CS HAD – YVF Adoption Dashboard"
DATA_FILE = Path(__file__).parent / "YVF_Adoption_Dashboard_Source.xlsx"
BLUE = "#0B4F8A"
ORANGE = "#F36F21"
DARK = "#17324D"
GREEN = "#1E9E68"
RED = "#D84A4A"
LIGHT_BLUE = "#EAF3FA"

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
<style>
[data-testid="stAppViewContainer"] {background: #F6F8FB;}
[data-testid="stSidebar"] {background: linear-gradient(180deg, #083F70 0%, #0B4F8A 100%);}
[data-testid="stSidebar"] * {color: white;}
[data-testid="stSidebar"] .stRadio label {padding: .35rem .5rem; border-radius: 8px;}
[data-testid="stSidebar"] .stRadio label:hover {background: rgba(255,255,255,.10);}
.block-container {
    padding-top: 1.25rem;
    padding-bottom: 2rem;
    padding-left: 2.25rem;
    padding-right: 2.25rem;
    max-width: 1500px;
}
[data-testid="stAppViewContainer"] .main {overflow-x: hidden;}
[data-testid="stAppViewContainer"] section.main {min-width: 0;}
[data-testid="collapsedControl"] {left: 0.75rem; top: 0.75rem; z-index: 1000;}
@media (max-width: 1200px) {
    .block-container {padding-left: 4rem; padding-right: 1.25rem;}
}
h1, h2, h3 {color: #17324D;}
.hero {
    background: linear-gradient(110deg,#0B4F8A,#1469A9);
    padding: 12px 22px;      /* trước là 22px 26px */
    border-radius: 14px;
    color: white;
    box-shadow: 0 4px 12px rgba(11,79,138,.12);
    margin-bottom: 12px;
}

.hero h1 {
    color:white;
    margin:0;
    font-size:22px;          /* trước là 30px */
    font-weight:700;
}

.hero p {
    margin:4px 0 0 0;
    opacity:.9;
    font-size:13px;          /* trước là 15px */
}
.kpi-card {background:white; border:1px solid #E5EAF0; border-radius:15px; padding:16px 17px; min-height:118px; height:100%; box-shadow:0 4px 14px rgba(23,50,77,.06); overflow-wrap:anywhere;}
.kpi-label {font-size:13px; color:#68798A; font-weight:600; margin-bottom:10px;}
.kpi-value {font-size:29px; font-weight:750; color:#17324D; line-height:1.05;}
.kpi-note {font-size:12px; color:#7D8B99; margin-top:8px;}
.section-title {font-size:18px; font-weight:750; color:#17324D; margin:8px 0 10px;}
.insight {background:white; border-left:5px solid #F36F21; border-radius:12px; padding:15px 17px; box-shadow:0 4px 14px rgba(23,50,77,.05); margin:4px 0 12px;}
.insight p {margin:0 0 9px 0; line-height:1.65;}
.insight p:last-child {margin-bottom:0;}
.quote-card {background:white; border:1px solid #E7ECF1; border-radius:14px; padding:16px 18px; height:100%; box-shadow:0 3px 12px rgba(23,50,77,.05);}
.quote {font-size:16px; color:#17324D; line-height:1.55;}
.quote-meta {font-size:12px; color:#758696; margin-top:10px; font-weight:650;}
.small-muted {font-size:12px;color:#738495;}
div[data-testid="stDataFrame"] {border:1px solid #E5EAF0; border-radius:12px; overflow:hidden;}
</style>
""",
    unsafe_allow_html=True,
)


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")


@st.cache_data(show_spinner=False)
def load_data(path: str) -> Dict[str, pd.DataFrame]:
    xlsx = pd.ExcelFile(path, engine="openpyxl")
    result: Dict[str, pd.DataFrame] = {}
    for sheet in xlsx.sheet_names:
        result[sheet] = clean_columns(pd.read_excel(path, sheet_name=sheet, header=1, engine="openpyxl"))
    return result


def safe_num(value, default=0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def fmt_int(value) -> str:
    return f"{safe_num(value):,.0f}"


def fmt_pct(value) -> str:
    v = safe_num(value)
    if v > 1.5:
        v /= 100
    return f"{v:.1%}"


def kpi(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>',
        unsafe_allow_html=True,
    )


def style_fig(fig, height: int = 350):
    fig.update_layout(
        height=height,
        margin=dict(l=15, r=15, t=55, b=15),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color=DARK, size=12),
        title_font=dict(size=16, color=DARK),
        legend_title_text="",
        hoverlabel=dict(bgcolor="white"),
    )
    fig.update_xaxes(showgrid=False, linecolor="#E8EDF2")
    fig.update_yaxes(gridcolor="#EEF2F6", zeroline=False)
    return fig


def normalize_dates(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


if not DATA_FILE.exists():
    st.error("Source file not found: YVF_Adoption_Dashboard_Source.xlsx")
    st.stop()

try:
    data = load_data(str(DATA_FILE))
except Exception as exc:
    st.error(f"Unable to read the source workbook: {exc}")
    st.stop()

required = {
    "Dashboard_Overview", "Customer_Volume", "Booking_Records", "Onboarded_Customers",
    "Improvement Proposals", "Customer_Feedback", "User Issues"
}
missing = sorted(required.difference(data))
if missing:
    st.error("Missing source sheets: " + ", ".join(missing))
    st.stop()

overview = data["Dashboard_Overview"]
volume = data["Customer_Volume"]
bookings = normalize_dates(data["Booking_Records"], ["Booking Date"])
onboarded = data["Onboarded_Customers"]
proposals = normalize_dates(data["Improvement Proposals"], ["Proposal Date"])
feedback = normalize_dates(data["Customer_Feedback"], ["Feedback Date"])
issues = normalize_dates(data["User Issues"], ["Date"])

if "No." in volume.columns:
    volume = volume[pd.to_numeric(volume["No."], errors="coerce").notna()].copy()

st.sidebar.markdown("## 📊 CS HAD")
st.sidebar.caption("YVF Adoption Dashboard")
page = st.sidebar.radio(
    "Navigation",
    ["🏠 Overview", "👥 Adoption", "📦 Booking Status", "⚠️ User Issues", "💡Enhancement", "⭐Feedback"],
    label_visibility="collapsed",
)

updated_at = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%d/%m/%Y %H:%M")
st.markdown(
    f"""
    <div class="hero">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;">
            <h1>{APP_TITLE}</h1>
            <span style="font-size:13px;font-weight:600;white-space:nowrap;">
                🕒 Last updated: {updated_at}
            </span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

ov = overview.iloc[-1] if not overview.empty else pd.Series(dtype="object")
eligible = safe_num(ov.get("Eligible Customers"))
total_hbl = safe_num(ov.get("Total Export HBLs"))
onboarded_count = safe_num(ov.get("Onboarded Customers"))
onboarding_rate = onboarded_count / eligible if eligible > 0 else 0
pending = safe_num(ov.get("Pending Customers"))
active = safe_num(ov.get("Active YVF Customers"))
yvf_bookings = safe_num(ov.get("YVF Bookings"))
avg_time = safe_num(ov.get("Avg. Booking Time (min/booking)"))
new_customer_target = safe_num(ov.get("New Customer Target"))
monthly_target = safe_num(ov.get("Monthly Booking Target"))
source_adoption_rate = (
    onboarded_count / eligible
    if eligible > 0
    else 0
)
booking_achievement = yvf_bookings / monthly_target if monthly_target else 0

if page == "🏠 Overview":
    cols = st.columns(6)
    with cols[0]:
        kpi("Eligible Customers", fmt_int(eligible), "Target customer pool")
    with cols[1]:
        kpi("Onboarded Customers", fmt_int(onboarded_count), f"Onboarding rate {fmt_pct(onboarding_rate)}")
    with cols[2]:
        kpi("Active YVF Customers", fmt_int(active), "Customers with YVF bookings")
    with cols[3]:
        kpi("Adoption Rate", fmt_pct(source_adoption_rate), "Onboarded / Eligible")
    with cols[4]:
        kpi("YVF Bookings", fmt_int(yvf_bookings), f"Monthly target {fmt_int(monthly_target)}")
    with cols[5]:
        kpi("Avg. Processing Time", f"{avg_time:.1f} min", "Per booking")

    left, right = st.columns([1.12, 1])
    with left:
        status_order = ["Fully Adopted", "Trial Booking Completed", "No YVF Booking Yet"]
        status_counts = onboarded["YVF Booking Status"].value_counts().reindex(status_order, fill_value=0).reset_index()
        status_counts.columns = ["Status", "Customers"]
        fig = px.bar(status_counts, x="Status", y="Customers", text="Customers", title="Approved Account Adoption Status",
                     color="Status", color_discrete_sequence=[GREEN, ORANGE, "#9AA9B6"])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=booking_achievement * 100,
            number={"suffix": "%", "font": {"size": 34}},
            delta={"reference": 100, "relative": False},
            title={"text": "Monthly Booking Target Achievement"},
            gauge={
                "axis": {"range": [0, max(120, booking_achievement * 120)]},
                "bar": {"color": ORANGE},
                "steps": [{"range": [0, 70], "color": "#FBE9DF"}, {"range": [70, 100], "color": "#DDEAF5"}],
                "threshold": {"line": {"color": GREEN, "width": 4}, "value": 100},
            },
        ))
        st.plotly_chart(style_fig(gauge), use_container_width=True)

    c1, c2 = st.columns([1.25, 1])
    with c1:
        booking_by_customer = bookings.groupby("Customer Name", as_index=False)["Bookings"].sum().sort_values("Bookings", ascending=True)
        fig = px.bar(booking_by_customer, x="Bookings", y="Customer Name", orientation="h", text="Bookings",
                     title="YVF Bookings by Customer", color_discrete_sequence=[BLUE])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig, 330), use_container_width=True)
    with c2:
        open_issues = int((issues["Status"].astype(str).str.lower() == "open").sum())
        completed_issues = int((issues["Status"].astype(str).str.lower() == "completed").sum())
        st.markdown('<div class="section-title">Management Snapshot</div>', unsafe_allow_html=True)
        st.markdown(
            f'''
            <div class="insight">
                <p><b>Adoption:</b> {fmt_int(active)} active customers out of {fmt_int(onboarded_count)} approved accounts.</p>
                <p><b>Bookings:</b> {fmt_int(yvf_bookings)} YVF bookings, equivalent to {fmt_pct(booking_achievement)} of the monthly target.</p>
                <p><b>Issues:</b> {open_issues} open and {completed_issues} completed.</p>
                <p><b>Customer Feedback:</b> {len(feedback)} positive feedback records captured.</p>
            </div>
            ''',
            unsafe_allow_html=True,
        )
        latest = issues.sort_values("Date", ascending=False).head(4)[["Date", "Customer", "Issue", "Status"]].copy()
        latest["Date"] = latest["Date"].dt.strftime("%d-%b-%Y")
        st.dataframe(latest, hide_index=True, use_container_width=True, height=220)

elif page == "👥 Customer Adoption":
    c = st.columns(4)
    with c[0]: kpi("Eligible Customers", fmt_int(eligible), "Customers suitable for promotion")
    with c[1]: kpi("Approved Accounts", fmt_int(len(onboarded)), "YVF accounts approved")
    with c[2]: kpi("Fully Adopted", fmt_int((onboarded["YVF Booking Status"] == "Fully Adopted").sum()), "Using YVF as standard process")
    with c[3]: kpi("Pending Customers", fmt_int(pending), "Pending onboarding")

    left, right = st.columns([1, 1.2])
    with left:
        status_counts = onboarded["YVF Booking Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Customers"]
        fig = px.pie(status_counts, names="Status", values="Customers", hole=.58, title="Approved Accounts by Booking Status",
                     color_discrete_sequence=[GREEN, ORANGE, "#9AA9B6", BLUE])
        fig.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        status_volume = volume.groupby("YVF Status", as_index=False)["Total Volume"].sum().sort_values("Total Volume", ascending=False)
        fig = px.bar(status_volume, x="YVF Status", y="Total Volume", text="Total Volume", title="Shipment Volume by YVF Status",
                     color="YVF Status", color_discrete_sequence=[BLUE, ORANGE, GREEN, "#8394A5", "#B0BBC5"])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)

    st.markdown('<div class="section-title">Approved Customer Accounts</div>', unsafe_allow_html=True)
    st.dataframe(onboarded, hide_index=True, use_container_width=True, height=340)

elif page == "📦 Booking Performance":
    month_options = sorted(bookings["Month"].dropna().astype(str).unique().tolist())
    selected_months = st.multiselect("Month", month_options, default=month_options)
    filtered = bookings[bookings["Month"].astype(str).isin(selected_months)] if selected_months else bookings.iloc[0:0]

    total_bookings = safe_num(filtered["Bookings"].sum())
    weighted_time = (filtered["Bookings"] * filtered["Processing Time (min)"]).sum() / total_bookings if total_bookings else 0
    c = st.columns(4)
    with c[0]: kpi("Bookings", fmt_int(total_bookings), "Selected period")
    with c[1]: kpi("Customers", fmt_int(filtered["Customer Name"].nunique()), "Customers with bookings")
    with c[2]: kpi("Avg. Processing Time", f"{weighted_time:.1f} min", "Weighted by booking volume")
    with c[3]: kpi("Completed Rate", fmt_pct((filtered["Status"] == "Completed").mean() if len(filtered) else 0), "Booking records completed")

    left, right = st.columns([1.25, 1])
    with left:
        daily = filtered.groupby("Booking Date", as_index=False)["Bookings"].sum().sort_values("Booking Date")
        fig = px.line(daily, x="Booking Date", y="Bookings", markers=True, title="Daily YVF Booking Trend", color_discrete_sequence=[BLUE])
        fig.update_traces(line=dict(width=3), marker=dict(size=8))
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        by_mode = filtered.groupby("Transport Mode", as_index=False)["Bookings"].sum()
        fig = px.pie(by_mode, names="Transport Mode", values="Bookings", hole=.55, title="Bookings by Transport Mode", color_discrete_sequence=[BLUE, ORANGE, GREEN])
        st.plotly_chart(style_fig(fig), use_container_width=True)

    left, right = st.columns(2)
    with left:
        by_customer = filtered.groupby("Customer Name", as_index=False)["Bookings"].sum().sort_values("Bookings", ascending=False)
        fig = px.bar(by_customer, x="Customer Name", y="Bookings", text="Bookings", title="Bookings by Customer", color_discrete_sequence=[ORANGE])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        by_handler = filtered.groupby("Handled By", as_index=False)["Bookings"].sum().sort_values("Bookings", ascending=False)
        fig = px.bar(by_handler, x="Handled By", y="Bookings", text="Bookings", title="Bookings by Handler", color_discrete_sequence=[BLUE])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)

    display = filtered.copy()
    display["Booking Date"] = display["Booking Date"].dt.strftime("%d-%b-%Y")
    st.markdown('<div class="section-title">Booking Details</div>', unsafe_allow_html=True)
    st.dataframe(display, hide_index=True, use_container_width=True, height=360)

elif page == "⚠️ User Issues":
    status_filter = st.multiselect("Status", sorted(issues["Status"].dropna().unique()), default=sorted(issues["Status"].dropna().unique()))
    filtered = issues[issues["Status"].isin(status_filter)] if status_filter else issues.iloc[0:0]
    c = st.columns(4)
    with c[0]: kpi("Total Issues", fmt_int(len(issues)), "All recorded issues")
    with c[1]: kpi("Open", fmt_int((issues["Status"] == "Open").sum()), "Require follow-up")
    with c[2]: kpi("Completed", fmt_int((issues["Status"] == "Completed").sum()), "Resolved issues")
    with c[3]: kpi("Resolution Rate", fmt_pct((issues["Status"] == "Completed").mean()), "Completed / total")

    left, right = st.columns(2)
    with left:
        cat = issues["Category"].value_counts().reset_index(); cat.columns=["Category","Issues"]
        fig = px.bar(cat, x="Category", y="Issues", text="Issues", title="Issues by Category", color_discrete_sequence=[ORANGE])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        stat = issues["Status"].value_counts().reset_index(); stat.columns=["Status","Issues"]
        fig = px.pie(stat, names="Status", values="Issues", hole=.58, title="Issue Resolution Status",
                     color="Status", color_discrete_map={"Open": RED, "In Progress": ORANGE, "Completed": GREEN})
        st.plotly_chart(style_fig(fig), use_container_width=True)

    display = filtered.copy(); display["Date"] = display["Date"].dt.strftime("%d-%b-%Y")
    st.dataframe(display, hide_index=True, use_container_width=True, height=380)

elif page == "💡 Improvement Proposals":
    c = st.columns(4)
    with c[0]: kpi("Total Proposals", fmt_int(len(proposals)), "All improvement ideas")
    with c[1]: kpi("Open", fmt_int((proposals["Status"] == "Open").sum()), "Awaiting action")
    with c[2]: kpi("Completed", fmt_int((proposals["Status"] == "Completed").sum()), "Implemented or closed")
    with c[3]: kpi("High Priority", fmt_int((proposals["Priority"] == "High").sum()), "Management attention")

    left, right = st.columns(2)
    with left:
        cat = proposals["Category"].value_counts().reset_index(); cat.columns=["Category","Proposals"]
        fig = px.bar(cat, x="Category", y="Proposals", text="Proposals", title="Proposals by Category", color_discrete_sequence=[BLUE])
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        stat = proposals["Status"].value_counts().reset_index(); stat.columns=["Status","Proposals"]
        fig = px.pie(stat, names="Status", values="Proposals", hole=.58, title="Proposal Status",
                     color="Status", color_discrete_map={"Open": RED, "In Progress": ORANGE, "Completed": GREEN})
        st.plotly_chart(style_fig(fig), use_container_width=True)

    display = proposals.copy(); display["Proposal Date"] = display["Proposal Date"].dt.strftime("%d-%b-%Y")
    st.dataframe(display, hide_index=True, use_container_width=True, height=380)

elif page == "⭐ Customer Feedback":
    c = st.columns(4)
    with c[0]: kpi("Positive Feedback", fmt_int(len(feedback)), "Recorded customer comments")
    with c[1]: kpi("Customers", fmt_int(feedback["Customer"].nunique()), "Customers providing feedback")
    with c[2]: kpi("Feedback Categories", fmt_int(feedback["Category"].nunique()), "Experience themes")
    with c[3]: kpi("Latest Feedback", feedback["Feedback Date"].max().strftime("%d-%b-%Y") if feedback["Feedback Date"].notna().any() else "–", "Most recent record")

    cat = feedback["Category"].value_counts().reset_index(); cat.columns=["Category","Feedback"]
    fig = px.bar(cat, x="Category", y="Feedback", text="Feedback", title="Positive Feedback by Category", color_discrete_sequence=[GREEN])
    fig.update_traces(textposition="outside")
    st.plotly_chart(style_fig(fig, 300), use_container_width=True)

    st.markdown('<div class="section-title">Customer Highlights</div>', unsafe_allow_html=True)
    cards = st.columns(min(3, max(1, len(feedback))))
    for idx, (_, row) in enumerate(feedback.head(3).iterrows()):
        with cards[idx]:
            date_text = row["Feedback Date"].strftime("%d-%b-%Y") if pd.notna(row["Feedback Date"]) else ""
            st.markdown(
                f'<div class="quote-card"><div class="quote">“{row["Positive Feedback"]}”</div>'
                f'<div class="quote-meta">{row["Customer"]} · {row["Category"]} · {date_text}</div>'
                f'<div class="small-muted">Business value: {row["Business Value"]}</div></div>',
                unsafe_allow_html=True,
            )

    display = feedback.copy(); display["Feedback Date"] = display["Feedback Date"].dt.strftime("%d-%b-%Y")
    st.markdown('<div class="section-title">Feedback Records</div>', unsafe_allow_html=True)
    st.dataframe(display, hide_index=True, use_container_width=True, height=300)
