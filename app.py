from __future__ import annotations

import json
import tempfile
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import pandas as pd
import plotly.express as px
import streamlit as st

from io_utils import save_uploads, list_xlsx_sheets, preview_file, preview_moda_core
from dq_matching import (
    profile_file,
    pairwise_matching,
    moda_file_overlap,
    unmatched_sites,
    scoped_source_site_matching,
)
from pipeline import run_analysis
from analytics import rows_to_df, overview_metrics, cfm_journey, cfm_woreda_table, protection_summary
from outputs import detailed_tracker_xlsx, aggregated_tracker_xlsx, management_report_docx, dq_report_docx, bundle_zip

# -----------------------------------------------------------------------------
# PAGE / BRANDING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="J-MAAP | Jijiga Monitoring, Analytics & Action Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

WFP_BLUE = "#007DBC"
WFP_DARK = "#005B8E"
TEXT_DARK = "#16324F"
BG = "#F5F8FB"
CARD_BORDER = "#DCE6EF"

st.markdown(
    f"""
<style>
:root {{ --wfp-blue:{WFP_BLUE}; --wfp-dark:{WFP_DARK}; --text:{TEXT_DARK}; --bg:{BG}; --border:{CARD_BORDER}; }}
html, body, [class*="css"] {{ font-family: "Open Sans", "Segoe UI", sans-serif; color: var(--text); }}
.stApp {{ background: var(--bg); }}
.block-container {{ padding-top: 0.35rem; padding-bottom: 2rem; max-width: 100%; }}
[data-testid="stSidebar"] {{ background: #F1F6FA; border-right: 1px solid #D7E3EC; }}
[data-testid="stSidebar"] .block-container {{ padding-top: .65rem; }}
[data-testid="stSidebarNav"] {{ display:none; }}
[data-testid="stMetric"] {{ background:white; border:1px solid var(--border); padding:14px 16px; border-radius:12px; box-shadow:0 1px 2px rgba(23,49,79,.05); }}
[data-testid="stMetricLabel"] {{ color:#60758A; font-size:.82rem; }}
[data-testid="stMetricValue"] {{ color:#0B5790; font-size:1.65rem; font-weight:700; }}
[data-testid="stDataFrame"] {{ background:white; border:1px solid var(--border); border-radius:12px; overflow:hidden; }}
.stButton > button, .stDownloadButton > button {{ border-radius:9px; font-weight:600; }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{ background:var(--wfp-blue); border-color:var(--wfp-blue); }}
.wfp-topbar {{
  margin: -.35rem -1rem 1rem -1rem; padding: 15px 28px; color:white;
  background: linear-gradient(90deg, #006AA6 0%, #007DBC 56%, #0067A0 100%);
  display:flex; align-items:center; justify-content:space-between; gap:24px; min-height:74px;
  box-shadow:0 2px 5px rgba(0,66,105,.18);
}}
.wfp-brand {{ display:flex; align-items:center; gap:14px; min-width:255px; }}
.wfp-mark {{ width:48px; height:48px; border:2px solid rgba(255,255,255,.9); border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:20px; font-weight:800; }}
.wfp-brand-title {{ font-weight:700; font-size:17px; line-height:1.1; }}
.wfp-brand-sub {{ font-size:11px; opacity:.9; margin-top:3px; }}
.wfp-center {{ text-align:center; flex:1; }}
.wfp-center h1 {{ margin:0; color:white; font-size:27px; font-weight:750; letter-spacing:.1px; }}
.wfp-center p {{ margin:4px 0 0 0; font-size:12px; opacity:.95; }}
.wfp-period {{ min-width:200px; text-align:right; font-size:12px; }}
.page-title {{ font-size:29px; color:#123D67; font-weight:750; margin: 4px 0 2px 0; }}
.page-subtitle {{ color:#60758A; margin-bottom:14px; font-size:.94rem; }}
.section-title {{ color:#123D67; font-weight:700; font-size:1.08rem; margin:10px 0 8px 0; }}
.card {{ background:white; border:1px solid var(--border); border-radius:12px; padding:16px 18px; box-shadow:0 1px 2px rgba(23,49,79,.04); }}
.card-title {{ color:#123D67; font-size:.98rem; font-weight:700; margin-bottom:4px; }}
.muted {{ color:#6B7F92; font-size:.85rem; }}
.status-ok {{background:#ECFDF3;border:1px solid #ABEFC6;padding:9px 12px;border-radius:9px;color:#176B45;}}
.status-warn {{background:#FFFAEB;border:1px solid #FEDF89;padding:9px 12px;border-radius:9px;color:#8A5A00;}}
.status-info {{background:#EFF8FF;border:1px solid #B2DDFF;padding:9px 12px;border-radius:9px;color:#175C8D;}}
.step-card {{background:white;border:1px solid var(--border);border-radius:11px;padding:12px 14px;margin-bottom:8px;}}
.step-num {{display:inline-block;width:25px;height:25px;border-radius:50%;background:#EAF4FA;color:#006AA6;font-weight:700;text-align:center;line-height:25px;margin-right:8px;}}
.sidebar-brand {{ background:#006EA9; color:white; padding:14px 12px; border-radius:10px; margin-bottom:12px; }}
.sidebar-brand strong {{font-size:15px;}}
.sidebar-brand small {{display:block; margin-top:4px; opacity:.9;}}
hr {{ border-color:#DDE7EF; }}
</style>
""",
    unsafe_allow_html=True,
)


def branded_header(reporting_month: str) -> None:
    month_label = reporting_month
    try:
        month_label = pd.Period(reporting_month, freq="M").strftime("%B %Y")
    except Exception:
        pass
    st.markdown(
        f"""
<div class="wfp-topbar">
  <div class="wfp-brand">
    <div class="wfp-mark">WFP</div>
    <div><div class="wfp-brand-title">World Food Programme</div><div class="wfp-brand-sub">SAVING LIVES · CHANGING LIVES</div></div>
  </div>
  <div class="wfp-center"><h1>J-MAAP – Jijiga Monitoring, Analytics & Action Platform</h1><p>From Data to Action | Accountable Programmes | Stronger Communities</p></div>
  <div class="wfp-period">Reporting Month<br><strong>{month_label}</strong></div>
</div>
""",
        unsafe_allow_html=True,
    )


def page_heading(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="page-title">{title}</div><div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def fmt_pct(x):
    return f"{x:.1%}" if pd.notna(x) else ""


@st.cache_data(show_spinner=False)
def cached_raw_preview(path: str, sheet: str | None, source_type: str):
    if source_type == "MoDa":
        return preview_moda_core(path, 1000)
    return preview_file(path, sheet, 1000)


# -----------------------------------------------------------------------------
# STATE / SIDEBAR
# -----------------------------------------------------------------------------
if "workdir" not in st.session_state:
    st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "upload_signature" not in st.session_state:
    st.session_state.upload_signature = None
if "generated_outputs" not in st.session_state:
    st.session_state.generated_outputs = None

with st.sidebar:
    st.markdown('<div class="sidebar-brand"><strong>Jijiga AO</strong><small>Monthly Monitoring Analytics</small></div>', unsafe_allow_html=True)
    reporting_month = st.text_input("Reporting month", value="2026-08", help="YYYY-MM; reporting period is derived from monitoring date.")
    sub_office = st.text_input("Sub-office", value="Jijiga")
    st.divider()
    nav = st.radio(
        "Navigation",
        [
            "Home",
            "1. Data Upload",
            "2. Data Quality & Matching",
            "3. Raw Data Explorer",
            "4. Monitoring Overview",
            "5. AAP / CFM",
            "6. Protection, Safety & Dignity",
            "7. Activity Analysis",
            "8. Findings & Actions",
            "9. Generate Outputs",
        ],
        index=0,
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("Validation gates")
    st.markdown("1. System signal  \\n2. Field validation  \\n3. Internal WFP agreement  \\n4. Action assignment  \\n5. Verification & closure")
    st.divider()
    st.caption("Version 1.1 · Streamlit deployment package")

branded_header(reporting_month)

# -----------------------------------------------------------------------------
# UPLOAD / PROFILING (persistent control at top)
# -----------------------------------------------------------------------------
uploads = st.file_uploader(
    "Upload monthly raw files",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    help="Upload one or more MoDa exports. RBMF, FRN, logistics/handover and previous action trackers are optional and will be profiled separately.",
    label_visibility="collapsed" if nav != "1. Data Upload" else "visible",
)

if uploads:
    signature = tuple((f.name, getattr(f, "size", len(f.getbuffer()))) for f in uploads)
    if signature != st.session_state.upload_signature:
        st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
        paths = save_uploads(uploads, st.session_state.workdir)
        profiles = []
        with st.spinner("Profiling uploaded files and building matching diagnostics..."):
            for p in paths:
                try:
                    profiles.append(profile_file(p))
                except Exception as e:
                    st.error(f"Could not profile {Path(p).name}: {e}")
            st.session_state.profiles = profiles
            st.session_state.paths = paths
            st.session_state.pairs = pairwise_matching(profiles)
            st.session_state.uuid_mat, st.session_state.site_mat = moda_file_overlap(profiles)
            st.session_state.unmatched = unmatched_sites(profiles)
        st.session_state.analysis = None
        st.session_state.analysis_params = None
        st.session_state.generated_outputs = None
        st.session_state.upload_signature = signature
else:
    profiles = st.session_state.get("profiles", [])
    paths = st.session_state.get("paths", [])

# -----------------------------------------------------------------------------
# HOME
# -----------------------------------------------------------------------------
if nav == "Home":
    page_heading("Monitoring Analytics Control Centre", "A transparent monthly workflow from raw submissions to validated management action.")
    if profiles:
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        a = st.session_state.analysis
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Files loaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Analysis status", "Ready" if a else "Not run")
        c4.metric("Reporting month", reporting_month)
        if a:
            dq = a["dq"]
            st.markdown('<div class="status-ok">Monthly analysis is available. Review the DQ exceptions and field-validation status before generating final outputs.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">Files are loaded. Continue to Data Quality & Matching before running the monthly analysis.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-info">Upload the monthly raw files to begin. MoDa exports drive the monitoring analysis; operational files are matched diagnostically.</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Monthly workflow</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    steps = [
        ("1", "Upload", "Load raw MoDa and optional operational sources."),
        ("2", "Validate", "Inspect schema, duplicates, date logic and source matching."),
        ("3", "Analyse", "Calculate indicators, thematic signals and site-level findings."),
        ("4", "Act", "Validate findings, aggregate actions and generate management outputs."),
    ]
    for col, (n, t, d) in zip(cols, steps):
        with col:
            st.markdown(f'<div class="card"><span class="step-num">{n}</span><b>{t}</b><div class="muted" style="margin-top:8px">{d}</div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# DATA UPLOAD
# -----------------------------------------------------------------------------
elif nav == "1. Data Upload":
    page_heading("1. Data Upload", "Load all monthly source files and confirm that the expected inputs have been recognised.")
    if not profiles:
        st.info("Upload one or more raw files using the uploader above.")
    else:
        inv = pd.DataFrame([{k: p.get(k) for k in ["file", "source_type", "sheet", "preview_rows", "columns", "site_column", "uuid_unique_preview"]} for p in profiles])
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Files uploaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Other sources", len(profiles) - moda_count)
        c4.metric("Reporting month", reporting_month)
        st.dataframe(inv, use_container_width=True, hide_index=True)
        if moda_count:
            st.markdown('<div class="status-ok">MoDa source detected. Continue to Data Quality & Matching before running analysis.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">No MoDa source detected. DQ profiling is available, but findings generation requires MoDa data.</div>', unsafe_allow_html=True)
        with st.expander("Column / schema inventory"):
            for p in profiles:
                st.markdown(f"**{p['file']} — {p['source_type']}**")
                st.write(p.get("column_names", []))

# -----------------------------------------------------------------------------
# DQ & MATCHING
# -----------------------------------------------------------------------------
elif nav == "2. Data Quality & Matching":
    page_heading("2. Data Quality & Matching", "Check source integrity, overlap and match quality before allowing the monthly analysis to proceed.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        pairs = st.session_state.get("pairs", pd.DataFrame())
        if not pairs.empty:
            st.markdown('<div class="section-title">Cross-file reconciliation</div>', unsafe_allow_html=True)
            st.dataframe(pairs, use_container_width=True, hide_index=True)
        uuid_mat, site_mat = st.session_state.get("uuid_mat", pd.DataFrame()), st.session_state.get("site_mat", pd.DataFrame())
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="section-title">MoDa UUID overlap</div>', unsafe_allow_html=True)
            st.dataframe(uuid_mat, use_container_width=True) if not uuid_mat.empty else st.caption("No comparable MoDa UUID matrix available.")
        with c2:
            st.markdown('<div class="section-title">Normalized site overlap</div>', unsafe_allow_html=True)
            st.dataframe(site_mat, use_container_width=True) if not site_mat.empty else st.caption("No comparable site matrix available.")

        um = st.session_state.get("unmatched", pd.DataFrame())
        if not um.empty:
            st.markdown('<div class="section-title">Optional-source site matching against MoDa</div>', unsafe_allow_html=True)
            view = um.copy()
            if "Match %" in view.columns:
                view["Match %"] = view["Match %"].map(fmt_pct)
            st.dataframe(view, use_container_width=True, hide_index=True)
        st.caption("Exact site-name matching is intentionally conservative. Production reconciliation should use the canonical Site Crosswalk for spelling variants and corporate IDs.")

        moda_paths = [p["path"] for p in profiles if p["source_type"] == "MoDa"]
        if moda_paths and st.button("Run controlled monthly analysis", type="primary", use_container_width=True):
            with st.spinner("Running controlled analysis and DQ rules..."):
                st.session_state.analysis = run_analysis(moda_paths, reporting_month, sub_office)
                st.session_state.analysis_params = (reporting_month, sub_office)
                st.session_state.generated_outputs = None
            st.success("Analysis completed. Review DQ exceptions below before relying on findings.")

        if st.session_state.analysis:
            dq = st.session_state.analysis["dq"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Accepted submissions", dq["rows"])
            c2.metric("Unique UUIDs", dq["uuid_unique"])
            c3.metric("Duplicates removed", dq["duplicate_uuid_rows"])
            c4.metric("DQ issues", len(dq["issues"]))
            st.markdown('<div class="section-title">Automated DQ exceptions</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(dq["issues"]), use_container_width=True, hide_index=True)
            scoped = scoped_source_site_matching(profiles, st.session_state.analysis["rows"])
            if not scoped.empty:
                scoped_view = scoped.copy()
                scoped_view["Match %"] = scoped_view["Match %"].map(fmt_pct)
                st.markdown('<div class="section-title">Matching to selected month / sub-office cohort</div>', unsafe_allow_html=True)
                st.dataframe(scoped_view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# RAW EXPLORER
# -----------------------------------------------------------------------------
elif nav == "3. Raw Data Explorer":
    page_heading("3. Raw Data Explorer", "Explore uploaded records before analysis. Filters and previews are designed for transparency and spot-checking.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        a = st.session_state.analysis
        if a:
            rows = a["rows"]
            m = overview_metrics(rows)
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total submissions", m["submissions"])
            c2.metric("Sites monitored", m["sites"])
            c3.metric("Woredas", m["woredas"])
            c4.metric("Activities", m["activities"])
            df_a = rows_to_df(rows)
            c5.metric("WFP submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "WFP").sum()))
            c6.metric("TPM submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "TPM").sum()))

        chosen = st.selectbox("Source file", [p["file"] for p in profiles])
        p = next(x for x in profiles if x["file"] == chosen)
        sheets = list_xlsx_sheets(p["path"])
        sheet = st.selectbox("Worksheet", sheets, index=(sheets.index("data") if "data" in sheets else 0)) if sheets else None
        dfraw = cached_raw_preview(p["path"], sheet, p["source_type"]).copy()

        q = st.text_input("Search displayed columns", value="", placeholder="Type a value, site, woreda, activity or UUID fragment...")
        if q and not dfraw.empty:
            mask = dfraw.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)
            dfraw = dfraw[mask]

        if p["source_type"] == "MoDa":
            st.caption("For performance, the preview exposes core raw fields only. The complete uploaded workbook remains the source used by the analysis engine.")

        if not dfraw.empty:
            # Lightweight visual preview from displayed core fields when recognizable.
            date_col = next((c for c in dfraw.columns if str(c).lower() in {"monitoring_date", "date", "start"}), None)
            activity_col = next((c for c in dfraw.columns if "activity" in str(c).lower()), None)
            woreda_col = next((c for c in dfraw.columns if "woreda" in str(c).lower()), None)
            chart_cols = st.columns(3)
            if date_col:
                tmp = pd.to_datetime(dfraw[date_col], errors="coerce").dt.date.value_counts().sort_index().reset_index()
                tmp.columns = ["Date", "Records"]
                with chart_cols[0]: st.plotly_chart(px.bar(tmp, x="Date", y="Records", title="Records by date"), use_container_width=True, config={"displayModeBar":False})
            if activity_col:
                tmp = dfraw[activity_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Activity","Records"]
                with chart_cols[1]: st.plotly_chart(px.pie(tmp, values="Records", names="Activity", hole=.5, title="Activity composition"), use_container_width=True, config={"displayModeBar":False})
            if woreda_col:
                tmp = dfraw[woreda_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Woreda","Records"]
                with chart_cols[2]: st.plotly_chart(px.bar(tmp.sort_values("Records"), x="Records", y="Woreda", orientation="h", title="Top woredas"), use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="section-title">Raw data preview</div>', unsafe_allow_html=True)
        st.dataframe(dfraw, use_container_width=True, height=540, hide_index=True)
        st.download_button("Download current preview as CSV", dfraw.to_csv(index=False).encode("utf-8-sig"), file_name=f"{Path(chosen).stem}_preview.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# MONITORING OVERVIEW
# -----------------------------------------------------------------------------
elif nav == "4. Monitoring Overview":
    page_heading("4. Monitoring Overview", "Summarise the volume, geography and provider composition of the accepted monthly monitoring cohort.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis from Data Quality & Matching.")
    else:
        rows = a["rows"]; df = rows_to_df(rows); m = overview_metrics(rows)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Submissions", m["submissions"]); c2.metric("Sites", m["sites"]); c3.metric("Woredas", m["woredas"]); c4.metric("Activities", m["activities"]); c5.metric("Providers", m["providers"])
        c1, c2 = st.columns(2)
        with c1:
            act = df["activity"].value_counts().reset_index(); act.columns=["Activity","Submissions"]
            st.plotly_chart(px.bar(act.sort_values("Submissions"), x="Submissions", y="Activity", orientation="h", title="Submissions by activity"), use_container_width=True)
        with c2:
            pr = df["provider"].value_counts().reset_index(); pr.columns=["Provider","Submissions"]
            st.plotly_chart(px.pie(pr, values="Submissions", names="Provider", hole=.55, title="WFP / TPM evidence composition"), use_container_width=True)
        if "monitoring_date" in df.columns:
            day = df.dropna(subset=["monitoring_date"]).groupby(df["monitoring_date"].dt.date).size().reset_index(name="Submissions"); day.columns=["Date","Submissions"]
            st.plotly_chart(px.line(day, x="Date", y="Submissions", markers=True, title="Monitoring submissions by date"), use_container_width=True)
        woreda = df["woreda"].replace("", pd.NA).value_counts().head(20).reset_index(); woreda.columns=["Woreda","Submissions"]
        st.plotly_chart(px.bar(woreda.sort_values("Submissions"), x="Submissions", y="Woreda", orientation="h", title="Submissions by woreda — top 20"), use_container_width=True)

# -----------------------------------------------------------------------------
# AAP / CFM
# -----------------------------------------------------------------------------
elif nav == "5. AAP / CFM":
    page_heading("5. AAP / CFM", "Track the community feedback journey from awareness and access through usage, response and satisfaction.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        rows = a["rows"]
        act_label = st.selectbox("Activity", ["Activity 1 (Relief response)", "Activity 2 (Nutrition assistance)", "Activity 3 (Refugee operations)", "Activity 6 (resilience)"], key="cfmact")
        res = cfm_journey(rows, activity=act_label)
        fdf = pd.DataFrame(res["stages"])
        if res["base"]:
            fdf["Percent of base"] = fdf["pct_base"] * 100
            c1, c2 = st.columns([1.15, 1])
            with c1:
                st.plotly_chart(px.funnel(fdf, x="count", y="stage", title=f"CFM journey — base n={res['base']}"), use_container_width=True)
            with c2:
                stage_cards = "".join([f'<div class="step-card"><b>{r.stage}</b><span style="float:right;color:#006EA9;font-weight:700">{r["Percent of base"]:.0f}%</span><div class="muted">{int(r["count"])} records</div></div>' for _, r in fdf.iterrows()])
                st.markdown(stage_cards, unsafe_allow_html=True)
            st.caption("Response and Satisfaction are only shown where supported by the activity module; Response uses a strict nested rule to avoid skip-logic inflation.")

        if act_label == "Activity 1 (Relief response)":
            frames = []
            for mod in ["Food", "Cash"]:
                rr = cfm_journey(rows, activity=act_label, modality=mod)
                for x in rr["stages"]:
                    frames.append({"Modality": mod, "Stage": x["stage"], "Percent of base": 100 * x["pct_base"] if x["pct_base"] is not None else None, "Base": rr["base"]})
            mdf = pd.DataFrame(frames)
            st.plotly_chart(px.line(mdf, x="Stage", y="Percent of base", color="Modality", markers=True, title="Relief CFM journey — Food vs Cash"), use_container_width=True)
            st.caption("Interpret modality differences descriptively: the August cash sample is geographically concentrated and should not be treated as a causal modality effect.")

        heat = cfm_woreda_table(rows, act_label)
        if not heat.empty:
            stages = [x for x in ["Awareness", "Access", "Usage", "Response", "Satisfaction"] if x in heat.columns]
            long = heat.melt(id_vars=["Woreda", "N"], value_vars=stages, var_name="Stage", value_name="Rate").dropna()
            if not long.empty:
                pivot = long.pivot(index="Woreda", columns="Stage", values="Rate") * 100
                st.plotly_chart(px.imshow(pivot, text_auto=".0f", aspect="auto", color_continuous_scale="Blues", title="Woreda CFM stage heatmap (%)"), use_container_width=True)
                st.dataframe(heat, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# PROTECTION
# -----------------------------------------------------------------------------
elif nav == "6. Protection, Safety & Dignity":
    page_heading("6. Protection, Safety & Dignity", "Review protection, inclusion, integrity and dignity signals while separating prevalence from seriousness.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        pdf = protection_summary(a["indicators"])
        if not pdf.empty:
            view = pdf.copy(); view["Issue rate %"] = view["Issue rate"].map(lambda x: 100 * x if pd.notna(x) else None)
            st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y="Finding", color="Activity", orientation="h", hover_data=["Severity", "Applicable", "Affected sites"], title="Protection / integrity assurance signals"), use_container_width=True)
            st.dataframe(view[["Activity", "Theme", "Finding", "Issues", "Applicable", "Issue rate %", "Severity", "Affected sites"]], use_container_width=True, hide_index=True)
            st.info("Frequency and seriousness are interpreted separately. Rare payment, misconduct or stock-control signals remain high-priority verification items even when prevalence is low.")

# -----------------------------------------------------------------------------
# ACTIVITY ANALYSIS
# -----------------------------------------------------------------------------
elif nav == "7. Activity Analysis":
    page_heading("7. Activity Analysis", "Review configured thematic findings by programme activity before moving to management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        idf = pd.DataFrame(a["indicators"])
        if idf.empty:
            st.info("No configured indicators available.")
        else:
            activity_col = "Activity" if "Activity" in idf.columns else ("activity" if "activity" in idf.columns else None)
            if activity_col:
                acts = sorted(idf[activity_col].dropna().astype(str).unique().tolist())
                act = st.selectbox("Activity", acts)
                view = idf[idf[activity_col].astype(str) == act].copy()
            else:
                view = idf.copy()
            rate_col = next((c for c in ["Issue_Rate", "Issue rate", "issue_rate"] if c in view.columns), None)
            if rate_col:
                view["Issue rate %"] = pd.to_numeric(view[rate_col], errors="coerce") * 100
                finding_col = next((c for c in ["Finding", "finding", "Indicator", "indicator"] if c in view.columns), view.columns[0])
                st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y=finding_col, orientation="h", title="Configured issue rates"), use_container_width=True)
            st.dataframe(view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# FINDINGS & ACTIONS
# -----------------------------------------------------------------------------
elif nav == "8. Findings & Actions":
    page_heading("8. Findings & Actions", "Move from provisional site-level signals to aggregated, owned and verifiable management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        ddf = pd.DataFrame(a["detail"]); adf = pd.DataFrame(a["agg"])
        c1, c2, c3 = st.columns(3)
        c1.metric("Detailed finding records", len(ddf)); c2.metric("Aggregated actions", len(adf)); c3.metric("Pending field validation", int((ddf.get("Validation_Status", pd.Series(dtype=str)) == "Pending field validation").sum()))
        st.markdown('<div class="section-title">Aggregated management actions</div>', unsafe_allow_html=True)
        st.dataframe(adf, use_container_width=True, height=330, hide_index=True)
        st.markdown('<div class="section-title">Detailed findings / validation register</div>', unsafe_allow_html=True)
        if not ddf.empty:
            c1, c2 = st.columns(2)
            with c1:
                act_opts = ["All"] + sorted(ddf["Activity"].dropna().unique().tolist()); act = st.selectbox("Activity", act_opts)
            with c2:
                sev_opts = ["All"] + sorted(ddf["Severity"].dropna().unique().tolist()); sev = st.selectbox("Severity", sev_opts)
            filtered = ddf.copy()
            if act != "All": filtered = filtered[filtered["Activity"] == act]
            if sev != "All": filtered = filtered[filtered["Severity"] == sev]
            st.dataframe(filtered, use_container_width=True, height=520, hide_index=True)

# -----------------------------------------------------------------------------
# OUTPUTS
# -----------------------------------------------------------------------------
elif nav == "9. Generate Outputs":
    page_heading("9. Generate Outputs", "Generate the monthly trackers, reports and machine-readable datasets from the validated analytical cohort.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        if st.button("Prepare downloadable outputs", type="primary", use_container_width=True):
            with st.spinner("Building trackers and reports..."):
                detail_xlsx = detailed_tracker_xlsx(a["detail"])
                agg_xlsx = aggregated_tracker_xlsx(a["agg"])
                mgmt_docx = management_report_docx(reporting_month, a["dq"], a["indicators"], a["agg"], a["detail"])
                dq_docx = dq_report_docx(reporting_month, a["dq"], a["indicators"])
                std_csv = rows_to_df(a["rows"]).drop(columns=["date_obj"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
                ind_csv = pd.DataFrame(a["indicators"]).to_csv(index=False).encode("utf-8-sig")
                bundle = bundle_zip({
                    f"{reporting_month}_Detailed_Findings_Tracker.xlsx": detail_xlsx,
                    f"{reporting_month}_Aggregated_Action_Tracker.xlsx": agg_xlsx,
                    f"{reporting_month}_Management_Monitoring_Report.docx": mgmt_docx,
                    f"{reporting_month}_Data_Quality_Report.docx": dq_docx,
                    f"{reporting_month}_Standardized_Submissions.csv": std_csv,
                    f"{reporting_month}_Indicator_Summary.csv": ind_csv,
                    f"{reporting_month}_DQ_Summary.json": json.dumps(a["dq"], default=str, indent=2).encode("utf-8"),
                })
                st.session_state.generated_outputs = {"month": reporting_month, "detail": detail_xlsx, "agg": agg_xlsx, "mgmt": mgmt_docx, "dq": dq_docx, "bundle": bundle}
        out = st.session_state.generated_outputs
        if out and out.get("month") == reporting_month:
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("Detailed findings tracker (.xlsx)", out["detail"], f"{reporting_month}_Detailed_Findings_Tracker.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("Management monitoring report (.docx)", out["mgmt"], f"{reporting_month}_Management_Monitoring_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with c2:
                st.download_button("Aggregated action tracker (.xlsx)", out["agg"], f"{reporting_month}_Aggregated_Action_Tracker.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("Data quality report (.docx)", out["dq"], f"{reporting_month}_Data_Quality_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            st.download_button("Download complete monthly output package (.zip)", out["bundle"], f"{reporting_month}_Monitoring_Output_Package.zip", "application/zip", type="primary", use_container_width=True)
            st.caption("The generated trackers preserve validation and management-agreement fields so automated signals can become agreed actions without losing the evidence trail.")
        else:
            st.markdown('<div class="status-info">Outputs are generated on demand so dashboard interaction remains responsive.</div>', unsafe_allow_html=True)
