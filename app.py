"""
MLM Break Time Tracker - Streamlit Dashboard
Supports multiple stations (DFH1, DVB8). Station is auto-detected from
uploaded filenames and can be overridden via the sidebar selector.
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import date, datetime
from analysis import run_analysis, export_excel

# -----------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------
st.set_page_config(
    page_title="MLM Break Time Tracker",
    page_icon="T",
    layout="wide",
    initial_sidebar_state="expanded",
)

KNOWN_STATIONS = ["DFH1", "DVB8"]

# -----------------------------------------------------------------
# CUSTOM CSS
# -----------------------------------------------------------------
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F7F9FC; }
[data-testid="stSidebar"]          { background: #FFFFFF; border-right: 2px solid #1B2A4A; }
[data-testid="stSidebar"] *        { color: #1B2A4A !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3       { color: #1B2A4A !important; }

.metric-row { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
.metric-card {
    flex: 1; min-width: 110px; border-radius: 10px;
    padding: 16px 12px; text-align: center;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
}
.metric-card .val { font-size: 2rem; font-weight: 700; line-height: 1.1; }
.metric-card .lbl { font-size: 0.72rem; text-transform: uppercase; letter-spacing:.05em; margin-top:4px; opacity:.8; }
.c-total   { background:#1B2A4A; color:#fff; }
.c-action  { background:#E87722; color:#fff; }
.c-major   { background:#FDDCDC; color:#C0392B; }
.c-mod     { background:#FFF3CD; color:#856404; }
.c-minor   { background:#FFF8DC; color:#8B6914; }
.c-missing { background:#EDE7F6; color:#6A1B9A; }
.c-match   { background:#D5F5E3; color:#1A7A42; }

.station-badge {
    display:inline-block; background:#E87722; color:#fff;
    border-radius:6px; padding:3px 12px; font-size:0.85rem;
    font-weight:700; letter-spacing:.08em; margin-left:10px;
    vertical-align:middle;
}

.script-card {
    border-radius:10px; padding:16px 20px; margin-bottom:12px;
    box-shadow:0 1px 4px rgba(0,0,0,0.08);
}
.script-card h4 { margin:0 0 6px 0; font-size:1rem; }
.script-card p  { margin:0; font-size:0.92rem; line-height:1.6; }
.s-major   { background:#FDDCDC; border-left:5px solid #C0392B; }
.s-mod     { background:#FFF3CD; border-left:5px solid #E87722; }
.s-minor   { background:#FFF8DC; border-left:5px solid #F39C12; }
.s-missing { background:#EDE7F6; border-left:5px solid #8E44AD; }

.page-header {
    background:linear-gradient(135deg,#1B2A4A 0%,#2E5090 100%);
    border-radius:12px; padding:22px 28px; margin-bottom:24px; color:#fff;
}
.page-header h1 { margin:0; font-size:1.6rem; }
.page-header p  { margin:4px 0 0; opacity:.75; font-size:0.9rem; }

.folder-tip {
    background:#EBF5FB; border-left:4px solid #2E5090;
    border-radius:6px; padding:10px 14px; font-size:0.85rem; color:#1B2A4A;
}
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------
# PASSWORD GATE
# -----------------------------------------------------------------
def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    st.markdown("""
    <div style="max-width:380px;margin:80px auto 0;text-align:center;">
      <div style="font-size:2.8rem;margin-bottom:8px;">T</div>
      <h2 style="color:#1B2A4A;margin-bottom:4px;">MLM Break Time Tracker</h2>
      <p style="color:#6B7280;font-size:0.9rem;">Enter your access password to continue</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pw = st.text_input("Password", type="password",
                           label_visibility="collapsed", placeholder="Enter password")
        if st.button("Sign In", use_container_width=True, type="primary"):
            correct = st.secrets.get("APP_PASSWORD", "mlm2026")
            if pw == correct:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    return False


# -----------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------
SEV_CLASS = {
    "Major":         ("s-major",   "RED"),
    "Moderate":      ("s-mod",     "ORANGE"),
    "Minor":         ("s-minor",   "YELLOW"),
    "Missing Entry": ("s-missing", "PURPLE"),
    "Match":         ("",          "CHECK"),
}

def fmt_time(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return "-"
    try:
        return val.strftime("%-I:%M %p") if hasattr(val, "strftime") else str(val)
    except Exception: return str(val)

def detect_station(adp_name: str, amz_name: str) -> str:
    combined = f"{adp_name} {amz_name}".upper()
    for s in KNOWN_STATIONS:
        if s in combined:
            return s
    return ""

def infer_date(adp_name: str, amz_name: str) -> str:
    for fname in [adp_name, amz_name]:
        m = re.search(r"(\d{1,2})[._-](\d{1,2})[._-]?(\d{2,4})?", fname)
        if m:
            mo, dy = m.group(1), m.group(2)
            yr = m.group(3) or str(date.today().year)
            yr = f"20{yr}" if len(yr) == 2 else yr
            try:
                return datetime(int(yr), int(mo), int(dy)).strftime("%B %d, %Y")
            except Exception:
                pass
    return date.today().strftime("%B %d, %Y")


# -----------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------
def render_sidebar():
    with st.sidebar:
        st.markdown("## Break Time Tracker")
        st.markdown("---")
        st.markdown("### Upload Daily Files")

        adp_file    = st.file_uploader("ADP Timecard Export (.xlsx)",    type=["xlsx"], key="adp_upload")
        amazon_file = st.file_uploader("Amazon Break Utilization (.xlsx)", type=["xlsx"], key="amz_upload")

        st.markdown("### Station")

        auto_station = ""
        if adp_file and amazon_file:
            auto_station = detect_station(adp_file.name, amazon_file.name)

        station_options = KNOWN_STATIONS + ([] if auto_station in KNOWN_STATIONS else [auto_station])
        default_idx     = station_options.index(auto_station) if auto_station in station_options else 0

        station = st.selectbox(
            "Select station",
            options=station_options,
            index=default_idx,
            label_visibility="collapsed",
            help="Auto-detected from filename. Override here if needed.",
        )

        if auto_station and auto_station == station:
            st.caption(f"Auto-detected: {station} from filename")
        elif auto_station and auto_station != station:
            st.caption(f"Auto-detected {auto_station}, overridden to {station}")
        else:
            st.caption("Station not found in filename - please confirm above.")

        st.markdown("---")

        run_btn = st.button(
            "Run Analysis",
            use_container_width=True,
            type="primary",
            disabled=(adp_file is None or amazon_file is None),
        )

        st.markdown("---")
        st.markdown("### How to use")
        st.markdown("""
1. Upload both files above
2. Confirm the station
3. Click **Run Analysis**
4. Review the dashboard
5. Download Excel report
6. Use 5 PM Scripts tab for conversations
        """)

        st.markdown("---")
        if st.button("Sign Out", use_container_width=True):
            st.session_state["authenticated"] = False
            st.session_state.pop("results", None)
            st.rerun()

        st.markdown(
            '<p style="font-size:0.7rem;opacity:.5;text-align:center;margin-top:24px;">'
            "MLM Break Tracker v1.0</p>",
            unsafe_allow_html=True,
        )

    return adp_file, amazon_file, station, run_btn


# -----------------------------------------------------------------
# METRIC CARDS
# -----------------------------------------------------------------
def render_metrics(df: pd.DataFrame):
    counts = df["severity"].value_counts()
    needs  = int(df["needs_action"].sum())
    cards  = [
        ("total",   len(df),                        "Total Employees"),
        ("action",  needs,                          "Need Action"),
        ("major",   counts.get("Major", 0),         "Major"),
        ("mod",     counts.get("Moderate", 0),      "Moderate"),
        ("minor",   counts.get("Minor", 0),         "Minor"),
        ("missing", counts.get("Missing Entry", 0), "Missing Entry"),
        ("match",   counts.get("Match", 0),         "Match"),
    ]
    html = '<div class="metric-row">'
    for key, val, lbl in cards:
        html += (f'<div class="metric-card c-{key}">'
                 f'<div class="val">{val}</div><div class="lbl">{lbl}</div></div>')
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# -----------------------------------------------------------------
# TAB 1 - DISCREPANCY TABLE
# -----------------------------------------------------------------
def render_table(df: pd.DataFrame):
    issues = df[df["needs_action"]].copy()
    if issues.empty:
        st.success("No discrepancies today - all break times match!")
        return

    rows = []
    for _, row in issues.iterrows():
        sev  = row.get("severity", "-")
        a_m  = row.get("amz_break_minutes")
        b_m  = row.get("adp_break_minutes")
        diff = row.get("diff_minutes")
        _, icon = SEV_CLASS.get(sev, ("", ""))
        rows.append({
            "Employee":      str(row.get("adp_name") or row.get("amz_name") or "-"),
            "Severity":      f"{icon} {sev}",
            "Amazon Break":  f"{a_m:.0f} min" if pd.notna(a_m) else "-",
            "ADP Break":     f"{b_m:.0f} min" if pd.notna(b_m) else "-",
            "Difference":    f"{diff:+.1f} min" if pd.notna(diff) else "-",
            "Direction":     str(row.get("direction", "-")),
            "Amz Start":     fmt_time(row.get("amz_break_start")),
            "Amz End":       fmt_time(row.get("amz_break_end")),
            "ADP Start":     fmt_time(row.get("adp_break_start")),
            "ADP End":       fmt_time(row.get("adp_break_end")),
        })

    display_df = pd.DataFrame(rows)

    def color_row(row):
        s = row["Severity"]
        if "Major"    in s: return ["background-color:#FDDCDC;color:#C0392B"] * len(row)
        if "Moderate" in s: return ["background-color:#FFF3CD;color:#856404"] * len(row)
        if "Minor"    in s: return ["background-color:#FFF8DC;color:#8B6914"] * len(row)
        if "Missing"  in s: return ["background-color:#EDE7F6;color:#6A1B9A"] * len(row)
        return [""] * len(row)

    st.markdown(f"**{len(issues)} employees need action before 5 PM today:**")
    st.dataframe(
        display_df.style.apply(color_row, axis=1),
        use_container_width=True,
        height=min(420, 60 + len(issues) * 38),
        hide_index=True,
    )


# -----------------------------------------------------------------
# TAB 2 - CONVERSATION SCRIPTS
# -----------------------------------------------------------------
def render_scripts(df: pd.DataFrame):
    issues = df[df["needs_action"]].reset_index(drop=True)
    if issues.empty:
        st.success("No conversations needed today - all break times match!")
        return

    st.markdown(f"**{len(issues)} employees to speak with at 5 PM.** Expand any card to copy the script.")
    st.markdown("")

    for _, row in issues.iterrows():
        sev    = row.get("severity", "-")
        emp    = str(row.get("adp_name") or row.get("amz_name") or "-")
        script = row.get("conversation_script", "")
        a_m    = row.get("amz_break_minutes")
        b_m    = row.get("adp_break_minutes")
        diff   = row.get("diff_minutes")
        sc, icon = SEV_CLASS.get(sev, ("", ""))

        a_str = f"{a_m:.0f} min" if pd.notna(a_m) else "-"
        b_str = f"{b_m:.0f} min" if pd.notna(b_m) else "-"
        d_str = f"{diff:+.1f} min" if pd.notna(diff) else "-"

        with st.expander(
            f"{icon}  **{emp}** - {sev}  |  Amazon: {a_str}  ADP: {b_str}  Diff: {d_str}"
        ):
            st.markdown(f"""
            <div class="script-card {sc}">
              <h4>Conversation Script</h4>
              <p>{script}</p>
            </div>
            """, unsafe_allow_html=True)
            st.code(script, language=None)


# -----------------------------------------------------------------
# TAB 3 - ALL EMPLOYEES
# -----------------------------------------------------------------
def render_all(df: pd.DataFrame):
    rows = []
    for _, row in df.iterrows():
        sev  = row.get("severity", "-")
        a_m  = row.get("amz_break_minutes")
        b_m  = row.get("adp_break_minutes")
        diff = row.get("diff_minutes")
        _, icon = SEV_CLASS.get(sev, ("", ""))
        rows.append({
            "Employee (ADP)": str(row.get("adp_name") or "-"),
            "Amazon DA Name": str(row.get("amz_name") or "-"),
            "Amazon (min)":   f"{a_m:.0f}" if pd.notna(a_m) else "-",
            "ADP (min)":      f"{b_m:.0f}" if pd.notna(b_m) else "-",
            "Difference":     f"{diff:+.1f}" if pd.notna(diff) else "-",
            "Status":         f"{icon} {sev}",
        })

    full_df = pd.DataFrame(rows)

    def color_row(row):
        s = row["Status"]
        if "Major"    in s: return ["background-color:#FDDCDC;color:#C0392B"] * len(row)
        if "Moderate" in s: return ["background-color:#FFF3CD;color:#856404"] * len(row)
        if "Minor"    in s: return ["background-color:#FFF8DC;color:#8B6914"] * len(row)
        if "Missing"  in s: return ["background-color:#EDE7F6;color:#6A1B9A"] * len(row)
        return ["background-color:#D5F5E3;color:#1A7A42"] * len(row)

    st.markdown(f"**All {len(df)} employees reviewed:**")
    st.dataframe(
        full_df.style.apply(color_row, axis=1),
        use_container_width=True,
        height=min(600, 60 + len(df) * 36),
        hide_index=True,
    )


# -----------------------------------------------------------------
# MAIN APP
# -----------------------------------------------------------------
def main():
    if not check_password():
        return

    adp_file, amazon_file, station, run_btn = render_sidebar()

    if run_btn and adp_file and amazon_file:
        with st.spinner(f"Analyzing break records for station {station}..."):
            try:
                df          = run_analysis(adp_file, amazon_file)
                report_date = infer_date(adp_file.name, amazon_file.name)
                st.session_state.update({
                    "results":     df,
                    "report_date": report_date,
                    "station":     station,
                    "adp_name":    adp_file.name,
                    "amz_name":    amazon_file.name,
                })
            except Exception as e:
                st.error(f"Analysis failed: {e}")
                return

    if "results" not in st.session_state:
        st.markdown("""
        <div class="page-header">
          <h1>MLM Break Time Tracker</h1>
          <p>Upload the ADP and Amazon files in the sidebar, confirm the station, then click <strong>Run Analysis</strong>.</p>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        c1.info("**Step 1:** Upload both files in the sidebar")
        c2.info("**Step 2:** Confirm the station (auto-detected from filename)")
        c3.info("**Step 3:** Click **Run Analysis** and review results")

        st.markdown("---")
        st.markdown("#### Recommended Folder Structure")
        st.markdown("""
        <div class="folder-tip">
        Save your daily downloaded reports to the matching station folder on your computer:<br><br>
        <code>MLM / DFH1 / Reports / Break_Discrepancy_Report_DFH1_February-18-2026.xlsx</code><br>
        <code>MLM / DVB8 / Reports / Break_Discrepancy_Report_DVB8_February-18-2026.xlsx</code>
        </div>
        """, unsafe_allow_html=True)
        return

    df          = st.session_state["results"]
    report_date = st.session_state.get("report_date", "")
    station     = st.session_state.get("station", "")
    disc_count  = int(df["needs_action"].sum())

    st.markdown(f"""
    <div class="page-header">
      <h1>Break Time Discrepancy Report
        <span class="station-badge">{station}</span>
      </h1>
      <p>{report_date}  -
         <strong style="color:#E87722;">{disc_count} employee{"s" if disc_count != 1 else ""}
         need{"" if disc_count != 1 else "s"} correction before 5 PM</strong>
      </p>
    </div>
    """, unsafe_allow_html=True)

    render_metrics(df)

    col_dl, col_tip = st.columns([2, 5])
    with col_dl:
        safe_date    = report_date.replace(" ", "-").replace(",", "")
        excel_bytes  = export_excel(df, report_date, station)
        filename     = f"Break_Discrepancy_Report_{station}_{safe_date}.xlsx"
        st.download_button(
            label="Download Excel Report",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_tip:
        st.markdown(
            f'<div class="folder-tip">Save this report to: '
            f'<code>MLM / {station} / Reports / {filename}</code></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([
        f"Needs Action ({disc_count})",
        "5 PM Scripts",
        f"All Employees ({len(df)})",
    ])
    with tab1: render_table(df)
    with tab2: render_scripts(df)
    with tab3: render_all(df)


if __name__ == "__main__":
    main()
