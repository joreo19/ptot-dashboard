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
COLOR_2024 = "#4FC3F7"
COLOR_2025 = "#FFB74D"
COLOR_2026 = "#F06292"
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
  .block-container {{ padding: 1.5rem 1.5rem 3rem; max-width: 900px; margin: auto; }}
  .dash-header {{ text-align: center; padding: 2.2rem 0 1.8rem; border-bottom: 1px solid #3A2830; margin-bottom: 1.8rem; }}
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
tab1, tab2 = st.tabs(["📊 Dashboard", "➕ Log a Job"])

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
    for year, color, width, dash in [(2024, COLOR_2024, 2, "dot"), (2025, COLOR_2025, 2.5, "dash"), (2026, COLOR_2026, 3.5, "solid")]:
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
    rows_html = ""
    for i, month in enumerate(MONTHS):
        r24 = df[(df.year == 2024) & (df.month == month)]
        r25 = df[(df.year == 2025) & (df.month == month)]
        r26 = df[(df.year == 2026) & (df.month == month)]
        v24 = r24["total_revenue"].values[0] if len(r24) else 0
        v25 = r25["total_revenue"].values[0] if len(r25) else 0
        v26 = r26["total_revenue"].values[0] if len(r26) else 0
        j24 = int(r24["jobs"].values[0]) if len(r24) else 0
        j25 = int(r25["jobs"].values[0]) if len(r25) else 0
        j26 = int(r26["jobs"].values[0]) if len(r26) else 0

        def cell(v): return '<td class="zero">—</td>' if v == 0 else f"<td>{fmt(v)}</td>"
        def jcell(v): return '<td class="zero">—</td>' if v == 0 else f"<td>{v}</td>"

        is_current = (i + 1 == current_month_num)
        row_class = ' class="current-month"' if is_current else ""
        marker = " ✦" if is_current else ""
        rows_html += f"<tr{row_class}><td>{month}{marker}</td>{cell(v24)}{jcell(j24)}{cell(v25)}{jcell(j25)}{cell(v26)}{jcell(j26)}</tr>"

    st.markdown(f"""
    <div class="chart-wrap">
    <table class="rev-table">
      <thead>
        <tr>
          <th rowspan="2">Month</th>
          <th colspan="2" style="text-align:center;color:{COLOR_2024};border-bottom:1px solid #3A2830;">2024</th>
          <th colspan="2" style="text-align:center;color:{COLOR_2025};border-bottom:1px solid #3A2830;">2025</th>
          <th colspan="2" style="text-align:center;color:{COLOR_2026};border-bottom:1px solid #3A2830;">2026</th>
        </tr>
        <tr>
          <th>Revenue</th><th>Jobs</th>
          <th>Revenue</th><th>Jobs</th>
          <th>Revenue</th><th>Jobs</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
      <tfoot><tr style="border-top: 2px solid #5A3040;">
        <td style="color:{TEXT_LIGHT};font-weight:600;font-family:'Playfair Display',serif;">Total</td>
        <td style="color:{COLOR_2024};font-weight:600;text-align:right;">{fmt(df[df.year==2024]['total_revenue'].sum())}</td>
        <td style="color:{COLOR_2024};font-weight:600;text-align:right;">{int(df[df.year==2024]['jobs'].sum())}</td>
        <td style="color:{COLOR_2025};font-weight:600;text-align:right;">{fmt(df[df.year==2025]['total_revenue'].sum())}</td>
        <td style="color:{COLOR_2025};font-weight:600;text-align:right;">{int(df[df.year==2025]['jobs'].sum())}</td>
        <td style="color:{COLOR_2026};font-weight:600;text-align:right;">{fmt(df[df.year==2026]['total_revenue'].sum())}</td>
        <td style="color:{COLOR_2026};font-weight:600;text-align:right;">{int(df[df.year==2026]['jobs'].sum())}</td>
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

    with st.form("job_entry_form", clear_on_submit=True):

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
            worker_hours = st.number_input("Worker Hours", min_value=0.0, max_value=12.0, value=0.0, step=0.5,
                                           disabled=(worker == "None"))

        income       = hours * rate
        worker_total = (worker_rate * worker_hours) if worker != "None" else 0
        net_revenue  = income - worker_total

        st.markdown(f"""
        <div style="background:{BG_DARK};border:1px solid #3A2830;border-radius:12px;padding:1rem;margin:1rem 0;">
          <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;text-align:center;">
            <div>
              <div style="color:{TEXT_DIM};font-size:.7rem;text-transform:uppercase;letter-spacing:.1em;">Income</div>
              <div style="color:{COLOR_2026};font-size:1.3rem;font-family:'Playfair Display',serif;">${income:,.0f}</div>
            </div>
            <div>
              <div style="color:{TEXT_DIM};font-size:.7rem;text-transform:uppercase;letter-spacing:.1em;">Worker Pay</div>
              <div style="color:{COLOR_2025};font-size:1.3rem;font-family:'Playfair Display',serif;">${worker_total:,.0f}</div>
            </div>
            <div>
              <div style="color:{TEXT_DIM};font-size:.7rem;text-transform:uppercase;letter-spacing:.1em;">Net Revenue</div>
              <div style="color:#A5D6A7;font-size:1.3rem;font-family:'Playfair Display',serif;">${net_revenue:,.0f}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

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