"""Client demonstration of the NCVET monthly monitoring pipeline."""
from datetime import datetime
from html import escape
from pathlib import Path
import subprocess
import sys
import time
import tempfile
import os
import pandas as pd
import plotly.express as px
import streamlit as st
from demo_data import ROOT, load_run, metrics

st.set_page_config(page_title="NCVET | Monitoring intelligence", page_icon="◈", layout="wide")
st.markdown("""<style>
.block-container {padding-top:4rem; padding-bottom:3rem; max-width:1500px}
[data-testid="stSidebar"] {background:#102B39}
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label {color:#E8F1F4!important}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {color:#A9C2CC!important}
h1 {font-size:2.45rem!important; letter-spacing:-1.2px; font-weight:750!important}
h2 {font-size:1.5rem!important; letter-spacing:-.4px}
h3 {font-size:1.12rem!important}
[data-testid="stMetric"] {background:white; border:1px solid #DFE6EC; border-radius:14px; padding:19px 22px; min-height:128px}
[data-testid="stMetricLabel"] p {color:#526674; font-size:.88rem}
[data-testid="stMetricLabel"] {white-space:normal!important; height:auto!important}
[data-testid="stMetricLabel"] p {white-space:normal!important; overflow:visible!important; text-overflow:clip!important}
[data-testid="stMetricValue"] {font-weight:700; color:#123B48}
.eyebrow {font-size:11px; font-weight:750; letter-spacing:2.3px; color:#16766D; text-transform:uppercase; margin-bottom:10px}
.hero {background:linear-gradient(110deg,#123B48,#16665F); color:white; border-radius:18px; padding:30px 34px; margin:12px 0 26px}
.hero h2 {color:white; font-size:1.7rem!important; margin:0 0 8px}
.hero p {color:#D6E9E7; max-width:860px; margin:0; line-height:1.6}
.flow {display:flex; gap:10px; margin:18px 0 25px; flex-wrap:wrap}
.flow div {flex:1; min-width:155px; border:1px solid #DFE6EC; border-radius:12px; background:white; padding:16px}
.flow b {display:block; margin-top:5px; font-size:14px; color:#173D4A}
.flow span {font-size:11px; color:#14756D; font-weight:700; letter-spacing:1px}
.valuecard {border:1px solid #DCE5EB; background:white; border-radius:12px; padding:22px; min-height:145px; overflow-wrap:anywhere}
.valuecard small {font-weight:700; color:#6B7D87; letter-spacing:1.2px}
.valuecard p {font-size:1.2rem; margin-top:16px}
.valuecard.after {border-top:4px solid #138779}
.valuecard.before {border-top:4px solid #D8A15D}
.footnote {color:#69808C; font-size:12px; border-top:1px solid #DFE6EC; padding-top:16px; margin-top:25px}
</style>""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=4)
def cached_run(folder, signature):
    return load_run(folder)


def table(frame, **kwargs):
    # Spreadsheet columns can mix dates, numbers and text; display without Arrow coercion.
    st.dataframe(frame.astype(str), hide_index=True, width="stretch", **kwargs)


def chart(fig, height=330, legend=False):
    fig.update_layout(height=height, margin=dict(l=5,r=15,t=20,b=10),
                      font=dict(family="Arial", color="#244553", size=12),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      showlegend=legend, legend=dict(orientation="h", y=-.12, x=0))
    st.plotly_chart(fig, width="stretch", config={"displayModeBar":False})


def heading(kicker, title, subtitle):
    st.markdown(f'<div class="eyebrow">{escape(kicker)}</div>', unsafe_allow_html=True)
    st.title(title)
    st.caption(subtitle)


def start_fresh():
    for key in ["active_run", "completed_run", "selected_run", "next_run"]:
        st.session_state.pop(key, None)
    st.session_state["upload_revision"] = st.session_state.get("upload_revision", 0) + 1
    st.session_state["demo_page"] = "Run pipeline"


def view_results():
    st.session_state["active_run"] = st.session_state["completed_run"]
    st.session_state["demo_page"] = "Overview"


def fresh_run_form():
    import calendar
    heading("New reporting run", "Start with a new file", "Upload your link workbook and select the reporting period to generate a fresh set of outputs.")
    left, right = st.columns(2)
    now = datetime.now()
    month_name = left.selectbox("Reporting month", list(calendar.month_name)[1:], index=(now.month - 2) % 12)
    years = list(range(2000, 2101))
    default_year = now.year if now.month > 1 else now.year - 1
    year = right.selectbox("Reporting year", years, index=years.index(default_year))
    month = f"{month_name} {year}"
    upload = st.file_uploader("Upload a new link workbook", type=["xlsx"], key=f"fresh_upload_{st.session_state.get('upload_revision', 0)}")
    st.caption("Include the agency name, reporting month, whether assessments were conducted, and the submission link. A timestamp is optional.")
    if upload is None:
        st.info("No file selected. Upload a workbook to begin.")
    if st.button("Run validation & cleaning", type="primary", disabled=upload is None):
        st.session_state.pop("completed_run", None)
        st.session_state.pop("active_run", None)
        # Each execution has its own output parent so another session's run cannot be selected.
        job_root = Path(tempfile.mkdtemp(prefix="ncvet_"))
        link_path = job_root / "links.xlsx"
        link_path.write_bytes(upload.getvalue())
        command = [sys.executable,"-u",str(ROOT/"run_pipeline.py"),"--links",str(link_path),"--month",month,"--output",str(job_root),"--aa-list",str(ROOT/"AA List.xlsx"),"--sectors",str(ROOT/"Sectors_List.xlsx"),"--state-district",str(ROOT/"State District List.xlsx"),"--dcf",str(ROOT/"AA Monthly DCF 2627 - V1 - 06 May 2026  (1).xlsx")]
        started = time.monotonic()
        with st.status("Processing the new workbook…", expanded=True) as progress:
            st.write(f"Downloading, validating and cleaning for {month}.")
            try:
                result = subprocess.run(command,cwd=ROOT,capture_output=True,text=True,encoding="utf-8",errors="replace",timeout=900,env={**os.environ, "NCVET_WEB_MODE": "1"})
                (job_root/"execution.log").write_text(result.stdout+result.stderr,encoding="utf-8")
                completed = [p for p in job_root.glob("run_*") if (p/"AA_Cleaned_Master.xlsx").is_file() and (p/"AA_Status_Report.xlsx").is_file()]
                if result.returncode == 0 and len(completed) == 1:
                    st.session_state["completed_run"] = str(completed[0])
                    progress.update(label=f"Completed in {time.monotonic()-started:.1f} seconds",state="complete",expanded=False)
                else:
                    progress.update(label="The run needs attention",state="error")
                    st.error("The new run did not complete. Check the workbook and submission links, then try again.")
                    with st.expander("Run details"):
                        st.code((result.stdout+result.stderr)[-8000:])
            except subprocess.TimeoutExpired:
                progress.update(label="Run timed out",state="error")
                st.error("The run exceeded 15 minutes. Check link access and try again.")
            except OSError as exc:
                progress.update(label="Unable to complete the run",state="error")
                st.error(str(exc))
    if st.session_state.get("completed_run"):
        st.success("Your new reports are ready.")
        st.button("View new results",type="primary",on_click=view_results)


# Migrate already-open demo sessions away from the earlier auto-loaded history.
if not st.session_state.get("fresh_start_enabled"):
    start_fresh()
    st.session_state["fresh_start_enabled"] = True

selected = st.session_state.get("active_run")
with st.sidebar:
    st.markdown("## ◈ NCVET")
    st.caption("ASSESSMENT AGENCY MONITORING")
    st.divider()
    pages = ["Overview", "Before & after", "Effectiveness", "Review queue", "Reports", "Run pipeline"] if selected else ["Run pipeline"]
    page = st.radio("Explore", pages, key="demo_page", label_visibility="collapsed")
    st.divider()
    st.button("Start fresh", on_click=start_fresh)
    st.caption("Only results from your new run appear here.")

if page == "Run pipeline" or not selected:
    fresh_run_form()
    st.stop()

folder = Path(selected)
signature = tuple((p.name, p.stat().st_mtime_ns) for p in folder.glob("*.xlsx"))
try:
    data = cached_run(selected, signature)
except Exception as exc:
    st.error(f"This reporting run could not be loaded: {exc}")
    st.stop()

raw, clean, log = data["raw"], data["clean"], data["log"]
all_metrics = metrics(raw, clean, log)
with st.sidebar:
    st.markdown(f"### {data['period']}")
    st.caption(f"{len(raw):,} consolidated records · {raw[data['agency_column']].nunique()} agencies with rows")
    st.caption("Unresolved values are marked Missing. They remain in the review queue.")


if page == "Overview":
    heading("Monitoring intelligence", "From submissions to structured insight.", f"{data['period']}  /  Assessment agency monitoring")
    st.markdown('<div class="hero"><h2>A visible trail from raw data to review.</h2><p>Bring agency submissions together, standardise against reference lists, and see exactly what changed. Every correction is traceable; unresolved values stay visible.</p></div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for col, label, value, help_text in zip(cols,
        ["Records processed", "Standardised / filled", "Marked Missing", "Flagged rows"],
        [all_metrics["rows"], all_metrics["resolved"], all_metrics["missing"], all_metrics["flagged"]],
        ["Rows in the consolidated and cleaned master.", "Logged changes excluding unresolved values.", "Unmatched cells replaced with Missing; these are not successful corrections.", "Rows with a non-empty Cleaning Flags value."]):
        col.metric(label, f"{value:,}", help=help_text)
    st.markdown('<div class="flow"><div><span>01 / COLLECT</span><b>Download & validate</b></div><div><span>02 / STANDARDISE</span><b>Clean against master lists</b></div><div><span>03 / REVIEW</span><b>Surface unresolved values</b></div><div><span>04 / REPORT</span><b>Download & retain evidence</b></div></div>', unsafe_allow_html=True)
    left, right = st.columns([1.4, 1])
    with left:
        st.subheader("Where the tool intervened")
        counts = log.groupby("Column").size().nlargest(7).sort_values().rename("Cells").reset_index()
        counts["Field"] = counts["Column"].map(lambda x: x if len(x) <= 30 else x[:29] + "…")
        chart(px.bar(counts, x="Cells", y="Field", orientation="h", color_discrete_sequence=["#198579"], hover_data=["Column"]))
    with right:
        st.subheader("Submission coverage")
        if not data["status"].empty:
            counts = data["status"]["Status"].value_counts().rename_axis("Status").reset_index(name="Agencies")
            chart(px.bar(counts, x="Agencies", y="Status", orientation="h", color="Status", color_discrete_map={"Submitted":"#198579", "Data Not Submitted":"#A9BAC5", "No Assessments":"#698AA3", "Not in Format":"#D6A15A", "File Error":"#B75650"}))
            st.caption("From the submission report, including agencies without a response.")
    st.info("Demo route: open Before & after to inspect a correction, then Effectiveness to quantify the work, and Review queue to see what still needs a decision.")

elif page == "Before & after":
    heading("Trace every change", "Before & after", "Compare actual submitted values with the cleaned output, one change or one full record at a time.")
    c1, c2, c3 = st.columns([1.4,1.4,1])
    agency = c1.selectbox("Agency", ["All agencies"] + sorted(log["Agency"].unique()))
    subset = log if agency == "All agencies" else log.loc[log["Agency"].eq(agency)]
    field = c2.selectbox("Field", ["All fields"] + sorted(subset["Column"].unique()))
    if field != "All fields":
        subset = subset.loc[subset["Column"].eq(field)]
    outcome = c3.selectbox("Outcome", ["All outcomes", "Standardised or filled", "Unmatched"])
    if outcome != "All outcomes":
        subset = subset.loc[subset["Match Method"].eq("Unmatched") if outcome == "Unmatched" else subset["Match Method"].ne("Unmatched")]
    st.caption(f"{len(subset):,} audit entries in this selection")
    if subset.empty:
        st.info("No changes match these filters.")
    else:
        selected_change = st.selectbox("Choose a change", list(subset.index), format_func=lambda i: f"Row {subset.at[i, 'Master Data Row']} · {subset.at[i, 'Column']} · {subset.at[i, 'Agency']}")
        change = subset.loc[selected_change]
        before, after = st.columns(2)
        for col, label, value, kind in [(before,"AS SUBMITTED",change["Original Value"],"before"),(after,"AFTER CLEANING",change["Cleaned Value"],"after")]:
            col.markdown(f'<div class="valuecard {kind}"><small>{label}</small><p>{escape(str(value)) if str(value) else "(blank)"}</p></div>', unsafe_allow_html=True)
        st.caption(f"{change['Change Type']} · Match method: {change['Match Method']} · Workbook row: {change['Master Data Row']}")
        if change["Match Method"] == "Unmatched":
            st.warning("This value could not be resolved. Missing is a review placeholder, not a verified replacement.")
        row_id = int(change["Master Data Row"])
        with st.expander("Inspect the complete record and its source"):
            common = list(raw.columns.intersection(clean.columns))
            comparison = pd.DataFrame({"Field":common,"Before":[raw.at[row_id,c] for c in common],"After":[clean.at[row_id,c] for c in common]})
            only_changed = st.checkbox("Show only fields with audit entries", value=True)
            if only_changed:
                comparison = comparison.loc[comparison["Field"].isin(log.loc[log["Master Data Row"].eq(row_id), "Column"])]
            table(comparison)
            st.caption(f"Source: {raw.at[row_id, 'Source File']} · sheet {raw.at[row_id, 'Source Sheet']} · row {raw.at[row_id, 'Source Row']}")
        st.subheader("Change ledger")
        table(subset[["Master Data Row","Agency","Column","Original Value","Cleaned Value","Change Type","Match Method"]], height=350)

elif page == "Effectiveness":
    heading("Measure the intervention", "Cleaning effectiveness", "Observed changes from the audit log. Standardisation does not establish that the original business facts were correct.")
    agency = st.selectbox("Agency", ["All agencies"] + sorted(raw[data["agency_column"]].unique()))
    scoped_raw = raw if agency == "All agencies" else raw.loc[raw[data["agency_column"]].eq(agency)]
    scoped_clean = clean.loc[scoped_raw.index]
    scoped_log = log.loc[log["Master Data Row"].isin(scoped_raw.index)]
    m = metrics(scoped_raw, scoped_clean, scoped_log)
    cols = st.columns(4)
    cols[0].metric("Rows touched", f"{m['affected']:,} / {m['rows']:,}")
    cols[1].metric("Standardised or filled", f"{m['resolved']:,}")
    cols[2].metric("Unresolved cells", f"{m['unmatched']:,}")
    cols[3].metric("Rows without cleaning flags", f"{m['rows']-m['flagged']:,} / {m['rows']:,}")
    left,right = st.columns([1,1.3])
    with left:
        st.subheader("How values were handled")
        counts = scoped_log.groupby("Change Type").size().reset_index(name="Cells")
        if not counts.empty:
            fig = px.pie(counts, names="Change Type", values="Cells", hole=.68, color="Change Type", color_discrete_map={"Standardised":"#198579", "Filled":"#6996B3", "Replaced with Missing":"#D6A15A"})
            fig.update_traces(textinfo="value", textposition="inside")
            chart(fig, height=370, legend=True)
    with right:
        st.subheader("Changes by field")
        pivot = scoped_log.groupby(["Column","Change Type"]).size().reset_index(name="Cells")
        top = scoped_log["Column"].value_counts().head(8).index
        pivot = pivot.loc[pivot["Column"].isin(top)].copy()
        pivot["Field"] = pivot["Column"].map(lambda x: x if len(x) <= 27 else x[:26] + "…")
        chart(px.bar(pivot,x="Cells",y="Field",color="Change Type",orientation="h",color_discrete_map={"Standardised":"#198579", "Filled":"#6996B3", "Replaced with Missing":"#D6A15A"},hover_data=["Column"]))
    st.subheader("What these measures mean")
    table(pd.DataFrame([
        ["Standardised or filled", "Changed cells with a successful match, parse, normalisation or fill; excludes unmatched values."],
        ["Unresolved cells", "Entries the cleaner could not resolve. Missing makes the gap explicit; it does not resolve it."],
        ["Rows without cleaning flags", "Rows with no flags from these cleaning rules. This is not a completeness or accuracy score."],
        ["Fuzzy matches", f"{m['fuzzy']} approximate matches in this selection; review their source values and confidence in the queue."],
    ], columns=["Measure","Interpretation"]))
    with st.expander("Illustrative manual review effort"):
        st.caption("A scenario estimate, not measured time saved. No benchmark of manual work has been collected.")
        seconds = st.slider("Assumed seconds to handle one standardised or filled cell", 5,120,30,5)
        st.metric("Equivalent manual handling effort", f"{m['resolved']*seconds/3600:.1f} hours")
        st.caption(f"{m['resolved']:,} cells × {seconds} seconds. Excludes unresolved values and does not subtract automated runtime or human review.")

elif page == "Review queue":
    heading("Human review", "Make the remaining gaps visible.", "Original values are retained. Missing values and approximate matches are available for follow-up.")
    mode = st.radio("Review type", ["Unmatched values", "Approximate matches", "Flagged records", "Unmatched form agencies"], horizontal=True)
    if mode in ["Unmatched values", "Approximate matches"]:
        queue = log.loc[log["Match Method"].eq("Unmatched" if mode == "Unmatched values" else "Fuzzy")].copy()
        agency = st.selectbox("Agency", ["All agencies"] + sorted(queue["Agency"].unique()))
        if agency != "All agencies":
            queue = queue.loc[queue["Agency"].eq(agency)]
        c1,c2,c3 = st.columns(3)
        c1.metric("Cells to review", len(queue))
        c2.metric("Distinct submitted values", queue[["Column","Original Value"]].drop_duplicates().shape[0])
        c3.metric("Records affected", queue["Master Data Row"].nunique())
        columns = ["Master Data Row","Agency","Column","Original Value","Cleaned Value","Match Method","Confidence"]
        table(queue[columns],height=440)
        if mode == "Approximate matches":
            st.caption("Confidence is the match similarity reported by the cleaner, not a probability that the value is correct.")
        st.download_button("Download this review queue",queue[columns].to_csv(index=False).encode("utf-8-sig"),file_name="review_queue.csv",mime="text/csv")
    elif mode == "Flagged records":
        queue = clean.loc[clean["Cleaning Flags"].astype(str).str.strip().ne("")]
        table(queue[[data["agency_column"],"Batch ID","Cleaning Flags"]].reset_index(),height=480)
    else:
        responses = data["responses"]
        if responses.empty:
            st.info("No response log is available for this run.")
        else:
            queue = responses.loc[responses["Match Method"].eq("Unmatched")]
            table(queue[["Response Row","Submitted Agency Name","Match Method","Submitted Month"]])
            st.caption("These form submissions did not match the agency master and were not processed into the consolidated data.")

elif page == "Reports":
    heading("Evidence & outputs", "Cleaning reports", "Download the submitted data, cleaned records and validation report.")
    st.subheader("Download workbooks")
    labels = {"AA_Consolidated_Master.xlsx":("Raw consolidated data","Submitted rows aligned to the DCF fields."),
              "AA_Cleaned_Master.xlsx":("Cleaned data & change trail","Standardised records, original values, flags and audit log."),
              "AA_Status_Report.xlsx":("Submission status","Coverage, format validation and the response log.")}
    files = [p for p in sorted(folder.glob("*.xlsx")) if p.name in labels]
    columns = st.columns(2)
    for i,path in enumerate(files):
        label,description = labels[path.name]
        with columns[i%2].container(border=True):
            st.subheader(label)
            st.caption(description)
            st.download_button("Download workbook", path.read_bytes(), file_name=path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=path.name)
    with st.expander("Metric provenance"):
        st.write("Counts come from Master Data and Cleaning Log. Submission coverage comes from Validation Details. Before-and-after rows are checked against source file, sheet and row identifiers.")
        if data["manifest"].get("clean", data["manifest"].get("stages",{}).get("2_clean",{})).get("log_truncated"):
            st.warning("This run has a truncated cleaning log. Audit-based metrics represent only the available entries.")


st.markdown(f'<div class="footnote">NCVET assessment agency monitoring · {escape(data["period"])} · Figures reflect the selected reporting run.</div>',unsafe_allow_html=True)
