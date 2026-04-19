import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, date
from data_loader import load_all_data, build_monthly_df, MONTHS, WORKER_COLS, get_drive_service

st.set_page_config(
    page_title="Purge This Organize That",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Bold, high-contrast palette ───────────────────────────────────────────────
COLOR_2021 = "#B39DDB"   # soft purple
COLOR_2022 = "#80CBC4"   # teal
COLOR_2023 = "#EF9A9A"   # soft red
COLOR_2024 = "#4FC3F7"   # bright sky blue
COLOR_2025 = "#FFB74D"   # vivid amber
COLOR_2026 = "#F06292"   # hot pink (current year)"
BG_DARK    = "#1C1518"
BG_CARD    = "#251C20"
TEXT_LIGHT = "#F5EDE8"
TEXT_DIM   = "#9E8580"
ACCENT     = "#F06292"
ACCENT2    = "#FFB74D"

WORKER_COLORS = {
    "tracy":    "#F06292",
    "tricia":   "#4FC3F7",
    "kristi":   "#A5D6A7",
    "chandler": "#FFB74D",
    "macy":     "#CE93D8",
    "amber":    "#80DEEA",
}

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,600;1,400&family=Jost:wght@300;400;500&display=swap');
  html, body, [class*="css"] {{ font-family: 'Jost', sans-serif; background-color: {BG_DARK}; color: {TEXT_LIGHT}; }}
  .block-container {{ padding: 0 1.5rem 3rem; max-width: 900px; margin: auto; }}
  .dash-header {{ text-align: center; padding: 1rem 0 1rem; border-bottom: 1px solid #3A2830; margin-bottom: 1rem; margin-top: 0; }}
  .dash-header .tagline-top {{ font-family: 'Jost', sans-serif; font-size: .7rem; font-weight: 500; letter-spacing: .2em; text-transform: uppercase; color: {ACCENT}; margin-bottom: .6rem; }}
  .dash-header h1 {{ font-family: 'Playfair Display', serif; font-size: clamp(1.9rem, 5vw, 2.8rem); color: {TEXT_LIGHT}; margin: 0 0 .4rem; line-height: 1.15; }}
  .dash-header h1 em {{ color: {ACCENT}; font-style: italic; }}
  .dash-header p {{ color: {TEXT_DIM}; font-size: .82rem; font-weight: 300; margin: 0; letter-spacing: .04em; }}
  .flourish {{ text-align: center; color: {ACCENT}; font-size: .9rem; letter-spacing: .3em; margin: .6rem 0 0; opacity: .5; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: .9rem; margin-bottom: 1.8rem; }}
  .kpi-card {{ background: {BG_CARD}; border: 1px solid #3A2830; border-radius: 14px; padding: 1.2rem 1rem; text-align: center; transition: transform .2s, border-color .2s; position: relative; overflow: hidden; }}
  .kpi-card::before {{ content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, {ACCENT2}, {ACCENT}); border-radius: 14px 14px 0 0; }}
  .kpi-card:hover {{ transform: translateY(-3px); border-color: #5A3040; }}
  .kpi-label {{ font-size: .68rem; font-weight: 500; letter-spacing: .1em; text-transform: uppercase; color: {TEXT_DIM}; margin-bottom: .4rem; }}
  .kpi-value {{ font-family: 'Playfair Display', serif; font-size: clamp(1.4rem, 4vw, 2rem); color: {ACCENT}; line-height: 1; }}
  .kpi-delta {{ font-size: .76rem; margin-top: .35rem; font-weight: 400; }}
  .delta-up {{ color: #A5D6A7; }}
  .delta-down {{ color: #EF9A9A; }}
  .delta-neutral {{ color: {TEXT_DIM}; }}
  .section-title {{ font-family: 'Playfair Display', serif; font-size: 1.2rem; color: {TEXT_LIGHT}; margin: 1.8rem 0 .7rem; padding-left: .1rem; }}
  .chart-wrap {{ background: {BG_CARD}; border: 1px solid #3A2830; border-radius: 14px; padding: 1rem .5rem .5rem; margin-bottom: 1.2rem; }}
  .rev-table {{ width: 100%; border-collapse: collapse; font-size: .85rem; }}
  .rev-table th {{ background: #2E1E24; color: {TEXT_DIM}; font-weight: 500; letter-spacing: .07em; font-size: .7rem; text-transform: uppercase; padding: .65rem .9rem; text-align: right; }}
  .rev-table th:first-child {{ text-align: left; }}
  .rev-table td {{ padding: .55rem .9rem; border-bottom: 1px solid #2E1E24; text-align: right; color: {TEXT_LIGHT}; }}
  .rev-table td:first-child {{ text-align: left; color: {TEXT_DIM}; font-weight: 400; }}
  .rev-table tr:hover td {{ background: #2E1E24; }}
  .rev-table tr.current-month td {{ color: {ACCENT}; font-weight: 600; }}
  .rev-table .zero {{ color: #3A2830; }}
  .dash-footer {{ text-align: center; color: {TEXT_DIM}; font-size: .72rem; margin-top: 2.5rem; padding-top: 1rem; border-top: 1px solid #3A2830; letter-spacing: .05em; }}
  div[data-testid="stToolbar"] {{ display: none; }}
  footer {{ display: none; }}
  .stSpinner > div {{ border-top-color: {ACCENT} !important; }}
  .stTabs {{ margin-top: 0 !important; padding-top: 0 !important; }}
  section[data-testid="stMain"] > div {{ padding-top: 0 !important; }}
  .stTabs [data-baseweb="tab-list"] {{ gap: 8px; background-color: {BG_CARD}; border-radius: 12px; padding: 4px; }}
  .stTabs [data-baseweb="tab"] {{ background-color: transparent; color: {TEXT_DIM}; border-radius: 8px; padding: 8px 20px; font-family: 'Jost', sans-serif; font-size: .85rem; }}
  .stTabs [aria-selected="true"] {{ background-color: {ACCENT} !important; color: white !important; }}
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="dash-header">
  <div class="tagline-top">Business Performance Dashboard</div>
  <h1>Purge This <em>Organize That</em></h1>
  <div class="flourish">— ✦ —</div>
  <p style="margin-top:.6rem">Year-over-Year Revenue · Updated {datetime.now().strftime("%B %d, %Y")}</p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab4, tab2, tab3, tab5 = st.tabs(["📊 Dashboard", "📈 Insights", "➕ Log a Job", "💰 Unpaid Jobs", "✏️ Edit Jobs"])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    @st.cache_data(ttl=300)
    def get_data():
        raw = load_all_data()
        return build_monthly_df(raw)

    with st.spinner("Loading your data…"):
        df = get_data()

    now = datetime.now()
    current_month_num = now.month
    current_month_name = MONTHS[current_month_num - 1]

    def ytd(df, year):
        return df[(df.year == year) & (df.month_num <= current_month_num)]["total_revenue"].sum()

    def fmt(v):
        return f"${v:,.0f}"

    def delta_html(curr, prev):
        if prev == 0:
            return '<span class="delta-neutral">—</span>'
        pct = (curr - prev) / prev * 100
        cls = "delta-up" if pct >= 0 else "delta-down"
        arrow = "▲" if pct >= 0 else "▼"
        return f'<span class="{cls}">{arrow} {abs(pct):.1f}% vs prior year</span>'

    def chart_layout(height=300):
        return dict(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Jost", color=TEXT_DIM, size=11),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                        font=dict(size=13, color=TEXT_LIGHT)),
            margin=dict(l=10, r=10, t=30, b=10),
            xaxis=dict(showgrid=False, tickfont=dict(size=10), tickangle=-30),
            yaxis=dict(showgrid=True, gridcolor="#2E1E24"),
            height=height,
        )

    ytd_2026 = ytd(df, 2026)
    ytd_2025 = ytd(df, 2025)
    ytd_2024 = ytd(df, 2024)
    annual_2025 = df[df.year == 2025]["total_revenue"].sum()
    annual_2024 = df[df.year == 2024]["total_revenue"].sum()
    cur_month_2026 = df[(df.year == 2026) & (df.month == current_month_name)]["total_revenue"].sum()
    cur_month_2025 = df[(df.year == 2025) & (df.month == current_month_name)]["total_revenue"].sum()

    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">2026 YTD Revenue</div>
        <div class="kpi-value">{fmt(ytd_2026)}</div>
        <div class="kpi-delta">{delta_html(ytd_2026, ytd_2025)}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">{current_month_name} 2026</div>
        <div class="kpi-value">{fmt(cur_month_2026)}</div>
        <div class="kpi-delta">{delta_html(cur_month_2026, cur_month_2025)}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Full Year 2025</div>
        <div class="kpi-value">{fmt(annual_2025)}</div>
        <div class="kpi-delta">{delta_html(annual_2025, annual_2024)}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">Monthly Revenue — 3-Year View</div>', unsafe_allow_html=True)
    fig_bar = go.Figure()
    for year, color in [(2024, COLOR_2024), (2025, COLOR_2025), (2026, COLOR_2026)]:
        ydata = [df[(df.year == year) & (df.month == m)]["total_revenue"].values[0]
                 if len(df[(df.year == year) & (df.month == m)]) else 0 for m in MONTHS]
        fig_bar.add_trace(go.Bar(name=str(year), x=MONTHS, y=ydata, marker_color=color,
                                 hovertemplate="<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>"))
    layout = chart_layout(300)
    layout["yaxis"]["tickprefix"] = "$"
    layout["yaxis"]["tickformat"] = ","
    fig_bar.update_layout(**layout, barmode="group")
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Jobs Per Month — 3-Year View</div>', unsafe_allow_html=True)
    fig_jobs = go.Figure()
    for year, color in [(2024, COLOR_2024), (2025, COLOR_2025), (2026, COLOR_2026)]:
        ydata = [int(df[(df.year == year) & (df.month == m)]["jobs"].values[0])
                 if len(df[(df.year == year) & (df.month == m)]) else 0 for m in MONTHS]
        fig_jobs.add_trace(go.Bar(name=str(year), x=MONTHS, y=ydata, marker_color=color,
                                  hovertemplate="<b>%{x}</b><br>Jobs: %{y}<extra></extra>"))
    fig_jobs.update_layout(**chart_layout(280), barmode="group")
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_jobs, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Worker Revenue Breakdown — 2026</div>', unsafe_allow_html=True)
    workers_2026 = [w.lower() for w in WORKER_COLS[2026] if w != "Total"]
    fig_workers = go.Figure()
    for w in workers_2026:
        col = f"worker_{w}"
        if col in df.columns:
            ydata = [df[(df.year == 2026) & (df.month == m)][col].values[0]
                     if len(df[(df.year == 2026) & (df.month == m)]) else 0 for m in MONTHS]
            fig_workers.add_trace(go.Bar(
                name=w.capitalize(), x=MONTHS, y=ydata,
                marker_color=WORKER_COLORS.get(w, "#888"),
                hovertemplate=f"<b>%{{x}}</b><br>{w.capitalize()}: $%{{y:,.0f}}<extra></extra>"
            ))
    layout_w = chart_layout(280)
    layout_w["yaxis"]["tickprefix"] = "$"
    layout_w["yaxis"]["tickformat"] = ","
    fig_workers.update_layout(**layout_w, barmode="stack")
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_workers, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Cumulative YTD Revenue</div>', unsafe_allow_html=True)
    fig_line = go.Figure()
    for year, color, width, dash in [(2021, COLOR_2021, 1.5, "dot"), (2022, COLOR_2022, 1.5, "dot"), (2023, COLOR_2023, 1.5, "dot"), (2024, COLOR_2024, 2, "dot"), (2025, COLOR_2025, 2.5, "dash"), (2026, COLOR_2026, 3.5, "solid")]:
        cumvals, running = [], 0
        for i, m in enumerate(MONTHS):
            if year == 2026 and i + 1 > current_month_num:
                break
            val = df[(df.year == year) & (df.month == m)]["total_revenue"].values
            running += val[0] if len(val) else 0
            cumvals.append((m, running))
        fig_line.add_trace(go.Scatter(name=str(year), x=[c[0] for c in cumvals], y=[c[1] for c in cumvals],
            mode="lines+markers", line=dict(color=color, width=width, dash=dash), marker=dict(size=6),
            hovertemplate="<b>%{x}</b><br>Cumulative: $%{y:,.0f}<extra></extra>"))
    layout_l = chart_layout(280)
    layout_l["yaxis"]["tickprefix"] = "$"
    layout_l["yaxis"]["tickformat"] = ","
    fig_line.update_layout(**layout_l)
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_line, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Monthly Revenue Breakdown</div>', unsafe_allow_html=True)
    def cell(v): return '<td class="zero">—</td>' if v == 0 else f"<td>{fmt(v)}</td>"
    def jcell(v): return '<td class="zero">—</td>' if v == 0 else f"<td>{v}</td>"
    rows_html = ""
    all_years = [2021, 2022, 2023, 2024, 2025, 2026]
    for i, month in enumerate(MONTHS):
        is_current = (i + 1 == current_month_num)
        row_class = ' class="current-month"' if is_current else ""
        marker = " ✦" if is_current else ""
        cells = ""
        for yr in all_years:
            r = df[(df.year == yr) & (df.month == month)]
            v = r["total_revenue"].values[0] if len(r) else 0
            j = int(r["jobs"].values[0]) if len(r) else 0
            cells += cell(v) + jcell(j)
        rows_html += f"<tr{row_class}><td>{month}{marker}</td>{cells}</tr>"

    yr_headers = "".join([f'<th colspan="2" style="text-align:center;color:{c};border-bottom:1px solid #3A2830;">{y}</th>' for y, c in [(2021,COLOR_2021),(2022,COLOR_2022),(2023,COLOR_2023),(2024,COLOR_2024),(2025,COLOR_2025),(2026,COLOR_2026)]])
    yr_subheaders = '<th>Revenue</th><th>Jobs</th>' * 6
    yr_totals = ""
    for _y, _c in [(2021,COLOR_2021),(2022,COLOR_2022),(2023,COLOR_2023),(2024,COLOR_2024),(2025,COLOR_2025),(2026,COLOR_2026)]:
        _rev = fmt(df[df.year==_y]["total_revenue"].sum())
        _jobs = int(df[df.year==_y]["jobs"].sum())
        yr_totals += f'<td style="color:{_c};font-weight:600;text-align:right;">{_rev}</td><td style="color:{_c};font-weight:600;text-align:right;">{_jobs}</td>'
    st.markdown(f"""
    <div class="chart-wrap" style="overflow-x:auto;">
    <table class="rev-table">
      <thead>
        <tr><th rowspan="2">Month</th>{yr_headers}</tr>
        <tr>{yr_subheaders}</tr>
      </thead>
      <tbody>{rows_html}</tbody>
      <tfoot><tr style="border-top: 2px solid #5A3040;">
        <td style="color:{TEXT_LIGHT};font-weight:600;font-family:'Playfair Display',serif;">Total</td>
        {yr_totals}
      </tr></tfoot>
    </table>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="dash-footer">
      Purge This Organize That · Data refreshes every 5 minutes · {datetime.now().strftime("%I:%M %p")}
    </div>
    """, unsafe_allow_html=True)



# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — LOG A JOB
# ════════════════════════════════════════════════════════════════════════════════
with tab2:

    st.markdown('<div class="section-title">Log a New Job</div>', unsafe_allow_html=True)

    FILE_ID_2026 = "1iVAghLUz1gIPFK-1Qq77YbdCW-9ILnb5TOANvv1t2G8"
    SHEET_2026   = "2026 PTOT Tracking"

    # Load customer list dynamically from the sheet
    @st.cache_data(ttl=120)
    def load_customers():
        import json, os
        from googleapiclient.discovery import build as build_service
        from google.oauth2 import service_account as sa
        if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
            info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        else:
            info = dict(st.secrets["gcp_service_account"])
        creds = sa.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        svc = build_service("sheets", "v4", credentials=creds)
        result = svc.spreadsheets().values().get(
            spreadsheetId=FILE_ID_2026,
            range=f"{SHEET_2026}!A:H"
        ).execute()
        rows = result.get("values", [])
        customers = {}
        for row in rows[1:]:  # skip header
            if not row or not row[0].strip():
                continue
            name = row[0].strip()
            addr = row[2].strip() if len(row) > 2 else ""
            mil  = float(row[7]) if len(row) > 7 and row[7].strip().replace('.','').isdigit() else 0
            if name not in customers:
                customers[name] = {"address": addr, "mileage": mil}
            else:
                # Update mileage with most recent entry
                if mil > 0:
                    customers[name]["mileage"] = mil
        return customers

    try:
        customers = load_customers()
    except Exception:
        customers = {}

    # Initialize session state
    if "cname" not in st.session_state:
        st.session_state["cname"] = "-- New Customer --"
    if "caddr" not in st.session_state:
        st.session_state["caddr"] = ""
    if "cmileage" not in st.session_state:
        st.session_state["cmileage"] = 0

    # Dropdown of existing customers
    dropdown_options = ["-- New Customer --"] + sorted(customers.keys())
    selected = st.selectbox(
        "Select Existing Customer (or choose New Customer to type a new name)",
        options=dropdown_options,
        key="customer_dropdown"
    )

    # Auto-fill when existing customer selected
    if selected != "-- New Customer --" and selected in customers:
        default_name    = selected
        default_address = customers[selected]["address"]
        default_mileage = int(customers[selected]["mileage"])
    else:
        default_name    = ""
        default_address = ""
        default_mileage = 0

    # Customer name and address fields
    customer_name = st.text_input("Customer Name", value=default_name, placeholder="Type name for new customer")
    address       = st.text_input("Customer Address", value=default_address, placeholder="Street address")

    col1, col2 = st.columns(2)
    with col1:
        job_date = st.date_input("Date of Work", value=date.today())
    with col2:
        description = st.text_input("Job Description", placeholder="e.g. bedroom closet, kitchen")

    col3, col4, col5 = st.columns(3)
    with col3:
        hours = st.number_input("Hours Worked", min_value=0.5, max_value=12.0, value=4.0, step=0.5)
    with col4:
        rate = st.number_input("Hourly Rate ($)", min_value=0, value=65, step=5)
    with col5:
        mileage = st.number_input("Travel Mileage", min_value=0, value=default_mileage, step=1)

    st.markdown("**Helper (optional)**")
    col6, col7, col8 = st.columns(3)
    with col6:
        worker = st.selectbox("Worker", options=["None", "Kristi", "Amber"])
    with col7:
        worker_rate = st.number_input("Worker Rate ($/hr)", min_value=0, value=20, step=5,
                                      disabled=(worker == "None"))
    with col8:
        worker_hours = st.number_input("Worker Hours", min_value=0.0, max_value=12.0, value=4.0, step=0.5)

    income       = hours * rate
    worker_total = (worker_rate * worker_hours) if worker != "None" else 0
    net_revenue  = income - worker_total

    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Income", f"${income:,.0f}")
    col_m2.metric("Worker Pay", f"${worker_total:,.0f}")
    col_m3.metric("Net Revenue", f"${net_revenue:,.0f}")

    with st.form("job_entry_form", clear_on_submit=True):
        submitted = st.form_submit_button("✅ Save Job to Spreadsheet", use_container_width=True)

    if submitted:
        final_name = customer_name.strip()
        final_addr = address.strip()
        if not final_name or not description:
            st.error("Please fill in Customer Name and Job Description.")
        else:
            try:
                with st.spinner("Saving to Google Sheets…"):
                    import json, os
                    from googleapiclient.discovery import build as build_service
                    from google.oauth2 import service_account as sa

                    if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
                        info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
                    else:
                        info = dict(st.secrets["gcp_service_account"])

                    creds = sa.Credentials.from_service_account_info(
                        info,
                        scopes=[
                            "https://www.googleapis.com/auth/drive",
                            "https://www.googleapis.com/auth/spreadsheets"
                        ]
                    )
                    sheets_service = build_service("sheets", "v4", credentials=creds)

                    worker_val = worker if worker != "None" else ""
                    new_row_values = [
                        final_name,
                        job_date.strftime("%Y-%m-%d"),
                        final_addr,
                        description,
                        float(hours),
                        float(rate),
                        float(income),
                        float(mileage),
                        "",
                        "",
                        worker_val,
                        float(worker_rate) if worker != "None" else "",
                        float(worker_hours) if worker != "None" else "",
                        float(worker_total) if worker != "None" else 0,
                        "",
                        float(net_revenue),
                    ]

                    # Find the last row with a customer name in column A
                    col_a = sheets_service.spreadsheets().values().get(
                        spreadsheetId=FILE_ID_2026,
                        range=f"{SHEET_2026}!A:A"
                    ).execute()
                    existing = col_a.get("values", [])
                    # Find last non-empty row in col A
                    last_row = 0
                    for i, row in enumerate(existing):
                        if row and str(row[0]).strip():
                            last_row = i + 1  # 1-indexed
                    next_row = last_row + 1

                    # Write directly to that specific row
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=FILE_ID_2026,
                        range=f"{SHEET_2026}!A{next_row}:P{next_row}",
                        valueInputOption="USER_ENTERED",
                        body={"values": [new_row_values]}
                    ).execute()

                st.session_state["cname"] = ""
                st.session_state["caddr"] = ""
                st.cache_data.clear()
                st.success(f"✅ Job saved! {final_name} · {description} · ${net_revenue:,.0f} net revenue")

            except Exception as e:
                st.error(f"Error saving job: {str(e)}")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3 — UNPAID JOBS
# ════════════════════════════════════════════════════════════════════════════════
with tab3:

    st.markdown('<div class="section-title">Outstanding Payments</div>', unsafe_allow_html=True)

    FILE_ID_2026_T3 = "1iVAghLUz1gIPFK-1Qq77YbdCW-9ILnb5TOANvv1t2G8"
    SHEET_2026_T3   = "2026 PTOT Tracking"

    @st.cache_data(ttl=60)
    def load_unpaid():
        import json, os
        from googleapiclient.discovery import build as build_service
        from google.oauth2 import service_account as sa
        if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
            info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        else:
            info = dict(st.secrets["gcp_service_account"])
        creds = sa.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        svc = build_service("sheets", "v4", credentials=creds)
        result = svc.spreadsheets().values().get(
            spreadsheetId=FILE_ID_2026_T3,
            range=f"{SHEET_2026_T3}!A:I"
        ).execute()
        rows = result.get("values", [])
        unpaid = []
        for i, row in enumerate(rows[1:], start=2):
            if not row or not row[0].strip():
                continue
            name      = row[0].strip()
            date_val  = row[1].strip() if len(row) > 1 else ""
            collected = row[8].strip().upper() if len(row) > 8 else ""
            if collected != "Y":
                unpaid.append({"row": i, "name": name, "date": date_val})
        return unpaid

    def mark_paid(row_num):
        import json, os
        from googleapiclient.discovery import build as build_service
        from google.oauth2 import service_account as sa
        if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
            info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        else:
            info = dict(st.secrets["gcp_service_account"])
        creds = sa.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        svc = build_service("sheets", "v4", credentials=creds)
        svc.spreadsheets().values().update(
            spreadsheetId=FILE_ID_2026_T3,
            range=f"{SHEET_2026_T3}!I{row_num}",
            valueInputOption="USER_ENTERED",
            body={"values": [["Y"]]}
        ).execute()

    try:
        unpaid_jobs = load_unpaid()
    except Exception as e:
        st.error(f"Error loading unpaid jobs: {str(e)}")
        unpaid_jobs = []

    if not unpaid_jobs:
        st.markdown(f'<div style="background:#251C20;border:1px solid #3A2830;border-radius:14px;padding:2rem;text-align:center;margin-top:1rem;"><div style="font-size:2rem;margin-bottom:.5rem;">🎉</div><div style="color:#A5D6A7;font-family:Playfair Display,serif;font-size:1.1rem;">All caught up!</div><div style="color:#9E8580;font-size:.85rem;margin-top:.3rem;">No outstanding payments.</div></div>', unsafe_allow_html=True)
    else:
        count = len(unpaid_jobs)
        st.markdown(f'<div style="color:#9E8580;font-size:.82rem;margin-bottom:1rem;">{count} outstanding payment{"s" if count != 1 else ""} — tap Paid to mark as collected</div>', unsafe_allow_html=True)
        for job in unpaid_jobs:
            col_a, col_b = st.columns([4, 1])
            with col_a:
                st.markdown(f'<div style="background:#251C20;border:1px solid #3A2830;border-radius:10px;padding:.8rem 1rem;margin-bottom:.5rem;"><div style="color:#F5EDE8;font-weight:500;">{job["name"]}</div><div style="color:#9E8580;font-size:.78rem;margin-top:.2rem;">{job["date"]}</div></div>', unsafe_allow_html=True)
            with col_b:
                if st.button("✅ Paid", key=f"paid_{job['row']}", use_container_width=True):
                    try:
                        mark_paid(job["row"])
                        st.cache_data.clear()
                        st.success(f"Marked {job['name']} as paid!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {str(e)}")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 4 — INSIGHTS
# ════════════════════════════════════════════════════════════════════════════════
with tab4:

    @st.cache_data(ttl=300)
    def get_insight_data():
        raw = load_all_data()
        return build_monthly_df(raw)

    df_i = get_insight_data()
    now_i = datetime.now()
    cur_month_i = now_i.month

    def chart_layout_i(height=280):
        return dict(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Jost", color="#9E8580", size=11),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                        font=dict(size=12, color="#F5EDE8")),
            margin=dict(l=10, r=10, t=30, b=10),
            xaxis=dict(showgrid=False, tickfont=dict(size=10), tickangle=-30),
            yaxis=dict(showgrid=True, gridcolor="#2E1E24"),
            height=height,
        )

    # ── 1. YTD Pace ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">YTD Pace — 2026 vs Prior Years</div>', unsafe_allow_html=True)

    fig_pace = go.Figure()
    for year, color, width, dash in [(2024, COLOR_2024, 2, "dot"), (2025, COLOR_2025, 2.5, "dash"), (2026, COLOR_2026, 3.5, "solid")]:
        cumvals, running = [], 0
        for i, m in enumerate(MONTHS):
            if i + 1 > cur_month_i:
                break
            val = df_i[(df_i.year == year) & (df_i.month == m)]["total_revenue"].values
            running += val[0] if len(val) else 0
            cumvals.append((m, running))
        fig_pace.add_trace(go.Scatter(
            name=str(year), x=[c[0] for c in cumvals], y=[c[1] for c in cumvals],
            mode="lines+markers", line=dict(color=color, width=width, dash=dash),
            marker=dict(size=7),
            hovertemplate="<b>%{x}</b><br>YTD: $%{y:,.0f}<extra></extra>"
        ))
    layout_p = chart_layout_i(280)
    layout_p["yaxis"]["tickprefix"] = "$"
    layout_p["yaxis"]["tickformat"] = ","
    fig_pace.update_layout(**layout_p)
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_pace, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 2. Average Revenue per Job ────────────────────────────────────────────
    st.markdown('<div class="section-title">Average Revenue per Job by Year</div>', unsafe_allow_html=True)

    avg_data = []
    for year in [2021, 2022, 2023, 2024, 2025, 2026]:
        yr_df = df_i[df_i.year == year]
        total_rev = yr_df["total_revenue"].sum()
        total_jobs = yr_df["jobs"].sum()
        avg = total_rev / total_jobs if total_jobs > 0 else 0
        avg_data.append((year, avg))

    colors_all = [COLOR_2021, COLOR_2022, COLOR_2023, COLOR_2024, COLOR_2025, COLOR_2026]
    fig_avg = go.Figure(go.Bar(
        x=[str(a[0]) for a in avg_data],
        y=[a[1] for a in avg_data],
        marker_color=colors_all,
        hovertemplate="<b>%{x}</b><br>Avg per job: $%{y:,.0f}<extra></extra>",
        showlegend=False
    ))
    layout_a = chart_layout_i(260)
    layout_a["yaxis"]["tickprefix"] = "$"
    layout_a["yaxis"]["tickformat"] = ","
    fig_avg.update_layout(**layout_a)
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_avg, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 3. Seasonal Monthly Average ───────────────────────────────────────────
    st.markdown('<div class="section-title">Seasonal Pattern — Average Revenue by Month (All Years)</div>', unsafe_allow_html=True)

    seasonal = []
    for m in MONTHS:
        vals = df_i[(df_i.month == m) & (df_i.year >= 2023)]["total_revenue"]
        seasonal.append(vals.mean() if len(vals) else 0)

    fig_seas = go.Figure(go.Bar(
        x=MONTHS, y=seasonal,
        marker_color=COLOR_2026,
        hovertemplate="<b>%{x}</b><br>Avg Revenue: $%{y:,.0f}<extra></extra>",
        showlegend=False
    ))
    layout_s = chart_layout_i(260)
    layout_s["yaxis"]["tickprefix"] = "$"
    layout_s["yaxis"]["tickformat"] = ","
    fig_seas.update_layout(**layout_s)
    st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
    st.plotly_chart(fig_seas, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 4. Top Clients 2026 ───────────────────────────────────────────────────
    st.markdown('<div class="section-title">Top Clients by Revenue — 2026</div>', unsafe_allow_html=True)

    try:
        import io
        from googleapiclient.http import MediaIoBaseDownload
        _svc = get_drive_service()
        _req = _svc.files().export_media(
            fileId="1iVAghLUz1gIPFK-1Qq77YbdCW-9ILnb5TOANvv1t2G8",
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        _buf = io.BytesIO()
        _dl = MediaIoBaseDownload(_buf, _req)
        _done = False
        while not _done:
            _, _done = _dl.next_chunk()
        _buf.seek(0)
        import pandas as _pd
        df_raw = _pd.read_excel(_buf, sheet_name="2026 PTOT Tracking", header=0)
        clients = {}
        freq = {}
        for _, row in df_raw.iterrows():
            name = str(row.iloc[0]).strip() if _pd.notna(row.iloc[0]) else ""
            if not name or name.startswith("Unnamed") or name == "nan":
                continue
            try:
                rev = float(row.iloc[15]) if _pd.notna(row.iloc[15]) else 0
            except (ValueError, TypeError):
                rev = 0
            clients[name] = clients.get(name, 0) + rev
            freq[name] = freq.get(name, 0) + 1

        # Top clients by revenue
        top_clients = sorted(clients.items(), key=lambda x: x[1], reverse=True)[:8]
        fig_clients = go.Figure(go.Bar(
            x=[c[0] for c in top_clients],
            y=[c[1] for c in top_clients],
            marker_color=COLOR_2026,
            hovertemplate="<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>",
            showlegend=False
        ))
        layout_c = chart_layout_i(260)
        layout_c["yaxis"]["tickprefix"] = "$"
        layout_c["yaxis"]["tickformat"] = ","
        layout_c["xaxis"]["tickangle"] = -30
        fig_clients.update_layout(**layout_c)
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(fig_clients, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

        # ── 6. Client Frequency ───────────────────────────────────────────────
        st.markdown('<div class="section-title">Client Visit Frequency — 2026</div>', unsafe_allow_html=True)
        top_freq = sorted(freq.items(), key=lambda x: x[1], reverse=True)[:8]
        fig_freq = go.Figure(go.Bar(
            x=[c[0] for c in top_freq],
            y=[c[1] for c in top_freq],
            marker_color=COLOR_2024,
            hovertemplate="<b>%{x}</b><br>Visits: %{y}<extra></extra>",
            showlegend=False
        ))
        layout_f = chart_layout_i(260)
        layout_f["xaxis"]["tickangle"] = -30
        fig_freq.update_layout(**layout_f)
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(fig_freq, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

        # ── 7. Tracy Take-Home vs Worker Costs ───────────────────────────────
        st.markdown('<div class="section-title">Tracy\'s Revenue vs Worker Costs — 2026</div>', unsafe_allow_html=True)
        result2 = svc.spreadsheets().values().get(
            spreadsheetId="1iVAghLUz1gIPFK-1Qq77YbdCW-9ILnb5TOANvv1t2G8",
            range="2026 PTOT Tracking!A:P"
        ).execute() if False else None

        tracy_by_month = []
        worker_by_month = []
        for m in MONTHS:
            r = df_i[(df_i.year == 2026) & (df_i.month == m)]
            tracy = r["worker_tracy"].values[0] if len(r) and "worker_tracy" in r.columns else 0
            worker_cost = (r["worker_amber"].values[0] if len(r) and "worker_amber" in r.columns else 0) + \
                          (r["worker_kristi"].values[0] if len(r) and "worker_kristi" in r.columns else 0)
            tracy_by_month.append(tracy)
            worker_by_month.append(worker_cost)

        fig_split = go.Figure()
        fig_split.add_trace(go.Bar(name="Tracy", x=MONTHS, y=tracy_by_month,
            marker_color=COLOR_2026, hovertemplate="<b>%{x}</b><br>Tracy: $%{y:,.0f}<extra></extra>"))
        fig_split.add_trace(go.Bar(name="Workers", x=MONTHS, y=worker_by_month,
            marker_color=COLOR_2025, hovertemplate="<b>%{x}</b><br>Workers: $%{y:,.0f}<extra></extra>"))
        layout_sp = chart_layout_i(260)
        layout_sp["yaxis"]["tickprefix"] = "$"
        layout_sp["yaxis"]["tickformat"] = ","
        fig_split.update_layout(**layout_sp, barmode="stack")
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(fig_split, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error loading client data: {str(e)}")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 5 — EDIT LAST 5 JOBS
# ════════════════════════════════════════════════════════════════════════════════
with tab5:

    st.markdown('<div class="section-title">Edit Recent Jobs</div>', unsafe_allow_html=True)

    FILE_ID_2026_E = "1iVAghLUz1gIPFK-1Qq77YbdCW-9ILnb5TOANvv1t2G8"
    SHEET_2026_E   = "2026 PTOT Tracking"

    def load_recent_jobs():
        import io, json, os
        from googleapiclient.http import MediaIoBaseDownload
        import pandas as pd
        svc = get_drive_service()
        request = svc.files().export_media(
            fileId=FILE_ID_2026_E,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        df_e = pd.read_excel(buf, sheet_name=SHEET_2026_E, header=0)
        # Get rows with valid customer names only
        job_rows = df_e[df_e.iloc[:, 0].notna() & 
                        df_e.iloc[:, 0].astype(str).str.strip().ne('') &
                        ~df_e.iloc[:, 0].astype(str).str.startswith('Unnamed')].copy()
        job_rows['_sheet_row'] = job_rows.index + 2  # 1-indexed + header row
        return job_rows.tail(5).reset_index(drop=True)

    def save_job_edit(sheet_row, values):
        import json, os
        from googleapiclient.discovery import build as build_service
        from google.oauth2 import service_account as sa
        if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
            info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        else:
            info = dict(st.secrets["gcp_service_account"])
        creds = sa.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        svc = build_service("sheets", "v4", credentials=creds)
        svc.spreadsheets().values().update(
            spreadsheetId=FILE_ID_2026_E,
            range=f"{SHEET_2026_E}!A{sheet_row}:P{sheet_row}",
            valueInputOption="USER_ENTERED",
            body={"values": [values]}
        ).execute()

    def delete_job_row(sheet_row):
        import json, os
        from googleapiclient.discovery import build as build_service
        from google.oauth2 import service_account as sa
        if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
            info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        else:
            info = dict(st.secrets["gcp_service_account"])
        creds = sa.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        svc = build_service("sheets", "v4", credentials=creds)
        # Get spreadsheet ID for the sheet
        spreadsheet = svc.spreadsheets().get(spreadsheetId=FILE_ID_2026_E).execute()
        sheet_id = None
        for s in spreadsheet["sheets"]:
            if s["properties"]["title"] == SHEET_2026_E:
                sheet_id = s["properties"]["sheetId"]
                break
        # Delete the row
        svc.spreadsheets().batchUpdate(
            spreadsheetId=FILE_ID_2026_E,
            body={"requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": sheet_row - 1,  # 0-indexed
                        "endIndex": sheet_row
                    }
                }
            }]}
        ).execute()

    try:
        recent_jobs = load_recent_jobs()
    except Exception as e:
        st.error(f"Error loading jobs: {str(e)}")
        recent_jobs = None

    if recent_jobs is not None:
        st.markdown(f'<div style="color:#9E8580;font-size:.82rem;margin-bottom:1rem;">Showing last {len(recent_jobs)} job entries — expand a job to edit it</div>', unsafe_allow_html=True)

        for idx, job in recent_jobs.iterrows():
            sheet_row = int(job['_sheet_row'])
            job_name = str(job.iloc[0]).strip()
            job_date_raw = job.iloc[1]
            try:
                job_date_val = pd.to_datetime(job_date_raw).date()
            except Exception:
                job_date_val = date.today()

            expander_label = f"{job_name} — {job_date_val.strftime('%B %d, %Y')}"

            with st.expander(expander_label):
                with st.form(f"edit_form_{sheet_row}", clear_on_submit=False):

                    col_a, col_b = st.columns(2)
                    with col_a:
                        e_name = st.text_input("Customer Name", value=job_name)
                    with col_b:
                        e_date = st.date_input("Date of Work", value=job_date_val)

                    e_address = st.text_input("Customer Address", value=str(job.iloc[2]).strip() if pd.notna(job.iloc[2]) else "")
                    e_desc    = st.text_input("Job Description",  value=str(job.iloc[3]).strip() if pd.notna(job.iloc[3]) else "")

                    col_c, col_d, col_e = st.columns(3)
                    with col_c:
                        e_hours = st.number_input("Hours Worked", min_value=0.5, max_value=12.0,
                            value=float(job.iloc[4]) if pd.notna(job.iloc[4]) else 4.0, step=0.5)
                    with col_d:
                        e_rate = st.number_input("Hourly Rate ($)", min_value=0,
                            value=int(job.iloc[5]) if pd.notna(job.iloc[5]) else 65, step=5)
                    with col_e:
                        e_mileage = st.number_input("Travel Mileage", min_value=0,
                            value=int(float(job.iloc[7])) if pd.notna(job.iloc[7]) else 0, step=1)

                    st.markdown("**Helper (optional)**")
                    col_f, col_g, col_h, col_i = st.columns(4)
                    with col_f:
                        current_worker = str(job.iloc[10]).strip() if pd.notna(job.iloc[10]) else "None"
                        worker_options = ["None", "Kristi", "Amber"]
                        worker_idx = worker_options.index(current_worker) if current_worker in worker_options else 0
                        e_worker = st.selectbox("Worker", options=worker_options, index=worker_idx)
                    with col_g:
                        e_worker_rate = st.number_input("Worker Rate ($/hr)", min_value=0,
                            value=int(float(job.iloc[11])) if pd.notna(job.iloc[11]) and job.iloc[11] != '' else 0, step=5)
                    with col_h:
                        e_worker_hours = st.number_input("Worker Hours", min_value=0.0, max_value=12.0,
                            value=float(job.iloc[12]) if pd.notna(job.iloc[12]) and job.iloc[12] != '' else 0.0, step=0.5)
                    with col_i:
                        e_worker_paid = st.selectbox("Worker Paid?",
                            options=["", "Y"],
                            index=1 if str(job.iloc[14]).strip().upper() == "Y" else 0)

                    # Recalculate
                    e_income       = e_hours * e_rate
                    e_worker_total = (e_worker_rate * e_worker_hours) if e_worker != "None" else 0
                    e_net_revenue  = e_income - e_worker_total

                    st.markdown(f"""
                    <div style="background:#1C1518;border:1px solid #3A2830;border-radius:10px;padding:.8rem;margin:.5rem 0;">
                      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:.5rem;text-align:center;">
                        <div><div style="color:#9E8580;font-size:.65rem;text-transform:uppercase;">Income</div>
                          <div style="color:#F06292;font-size:1.1rem;">${e_income:,.0f}</div></div>
                        <div><div style="color:#9E8580;font-size:.65rem;text-transform:uppercase;">Worker Pay</div>
                          <div style="color:#FFB74D;font-size:1.1rem;">${e_worker_total:,.0f}</div></div>
                        <div><div style="color:#9E8580;font-size:.65rem;text-transform:uppercase;">Net Revenue</div>
                          <div style="color:#A5D6A7;font-size:1.1rem;">${e_net_revenue:,.0f}</div></div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    save_btn = st.form_submit_button("💾 Save Changes", use_container_width=True)

                if save_btn:
                    try:
                        worker_val = e_worker if e_worker != "None" else ""
                        updated_row = [
                            e_name,
                            e_date.strftime("%Y-%m-%d"),
                            e_address,
                            e_desc,
                            float(e_hours),
                            float(e_rate),
                            float(e_income),
                            float(e_mileage),
                            str(job.iloc[8]) if pd.notna(job.iloc[8]) else "",  # $ Collected - preserve
                            str(job.iloc[9]) if pd.notna(job.iloc[9]) else "",  # Referral - preserve
                            worker_val,
                            float(e_worker_rate) if e_worker != "None" else "",
                            float(e_worker_hours) if e_worker != "None" else "",
                            float(e_worker_total) if e_worker != "None" else 0,
                            e_worker_paid,
                            float(e_net_revenue),
                        ]
                        save_job_edit(sheet_row, updated_row)
                        st.success(f"✅ {e_name} updated successfully!")
                        st.cache_data.clear()
                    except Exception as e:
                        st.error(f"Error saving: {str(e)}")

                # ── Delete button outside the form ────────────────────────────
                st.markdown("---")
                col_del1, col_del2 = st.columns([3, 1])
                with col_del2:
                    if st.button("🗑️ Delete Job", key=f"del_{sheet_row}", use_container_width=True):
                        st.session_state[f"confirm_delete_{sheet_row}"] = True

                if st.session_state.get(f"confirm_delete_{sheet_row}", False):
                    st.warning(f"Are you sure you want to delete this job? This cannot be undone.")
                    col_yes, col_no = st.columns(2)
                    with col_yes:
                        if st.button("Yes, delete it", key=f"yes_del_{sheet_row}", use_container_width=True):
                            try:
                                delete_job_row(sheet_row)
                                st.session_state[f"confirm_delete_{sheet_row}"] = False
                                st.cache_data.clear()
                                st.success("Job deleted!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error deleting: {str(e)}")
                    with col_no:
                        if st.button("Cancel", key=f"no_del_{sheet_row}", use_container_width=True):
                            st.session_state[f"confirm_delete_{sheet_row}"] = False
                            st.rerun()