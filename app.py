"""Web front end: upload a schedule workbook, watch it build, download the result.

Run locally:  streamlit run app.py
"""
import tempfile
from pathlib import Path

import streamlit as st

from scheduler.loader import LoaderError
from scheduler.pipeline import run
from scheduler.template_builder import refresh_template

st.set_page_config(page_title="Pharmacist Scheduler", page_icon="🗓️", layout="centered")

st.title("Pharmacist Scheduler")
st.write(
    "Upload your schedule workbook. Everything you've already entered is kept; "
    "the program fills the blank cells and gives you back a finished copy to download."
)

upload = st.file_uploader("Schedule workbook (.xlsx)", type=["xlsx"])

with st.expander("Options"):
    minutes = st.number_input(
        "Search time (seconds)", min_value=0, max_value=600, value=0, step=30,
        help="0 uses the time on the workbook's Settings sheet. Longer can give a slightly better "
             "schedule; 60 is usually plenty, 180 for a final version.")
    refresh_only = st.button(
        "Only update the live checks (fill nothing)", use_container_width=True, disabled=not upload,
        help="Rebuilds the days/target columns and the Shifts Fulfilled table, e.g. after adding a "
             "pharmacist or a role, without changing the schedule.")

# Forget an old result when a different file is chosen.
file_key = (upload.name, upload.size) if upload else None
if st.session_state.get("file_key") != file_key:
    st.session_state.pop("result", None)
    st.session_state["file_key"] = file_key

if upload and refresh_only:
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / Path(upload.name).name
        src.write_bytes(upload.getvalue())
        out = Path(tmp) / f"{src.stem} - checks updated.xlsx"
        try:
            refresh_template(src, out)
            st.session_state["result"] = {"ok": True, "refresh": True, "name": out.name,
                                          "bytes": out.read_bytes()}
        except LoaderError as e:
            st.session_state["result"] = {"ok": False, "message": str(e), "errors": [],
                                          "warnings": []}

if upload and st.button("Build schedule", type="primary", use_container_width=True):
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / Path(upload.name).name
        src.write_bytes(upload.getvalue())
        out_name = f"{src.stem} - solved.xlsx"
        out = Path(tmp) / out_name

        with st.status("Building the schedule…", expanded=True) as status:
            bar = st.progress(0.0, text="Starting")

            def progress(stage, frac, msg):
                bar.progress(min(max(frac, 0.0), 1.0), text=msg)

            res = run(src, out, time_limit=minutes or None, on_progress=progress)
            if res.ok:
                status.update(label="Schedule built", state="complete", expanded=False)
            else:
                status.update(label="Couldn't build the schedule", state="error", expanded=False)

        st.session_state["result"] = {
            "ok": res.ok,
            "message": res.message,
            "errors": [(i.category, i.message) for i in res.errors],
            "warnings": [(i.category, i.message) for i in res.warnings],
            "summary": res.summary,
            "name": out_name,
            "bytes": out.read_bytes() if res.ok else None,
        }

r = st.session_state.get("result")
if r:
    if not r["ok"]:
        st.error(r["message"])
        if "Couldn't find a cell labeled" in r["message"]:
            st.warning("This doesn't look like a schedule workbook. Upload the department's "
                       "schedule template (the workbook with the Schedule, Aptitude, Requirements "
                       "and Settings sheets).")
        for cat, msg in r["errors"]:
            st.markdown(f"- **{cat}**: {msg}")
        st.info("Fix these in the workbook and upload it again. Your file hasn't been changed.")
    elif r.get("refresh"):
        st.success("Live checks updated. Nothing on the schedule was changed.")
        st.download_button("Download the updated workbook", r["bytes"], file_name=r["name"],
                           type="primary", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        s = r["summary"]
        st.success(f"Done: {s['pharmacists']} pharmacists, {s['start']:%b %-d} – {s['end']:%b %-d, %Y}.")
        st.download_button("Download the finished schedule", r["bytes"], file_name=r["name"],
                           type="primary", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Cells filled", s["filled"])
        c2.metric("Required slots empty", s["required_empty"])
        c3.metric("PRN shifts needed", s["prn_needed"])
        c4.metric("Pay periods off target", s["short_periods"] + s["over_periods"])
        notes = []
        if s["required_empty"]:
            notes.append(f"{s['required_empty']} required slot(s) couldn't be covered: red in the Shifts "
                         f"Fulfilled table under the schedule.")
        if s["prn_needed"]:
            notes.append(f"{s['prn_needed']} PRN shift(s) needed: see the PRN NEEDED row and the PRN "
                         f"Requests sheet.")
        if s["short_periods"]:
            notes.append(f"{s['short_periods']} pay period(s) where someone is short of their target: red "
                         f"in the days/target columns right of the schedule.")
        if s["over_periods"]:
            notes.append(f"{s['over_periods']} pay period(s) over target because of days entered on the "
                         f"template: yellow in the days/target columns.")
        for n in notes:
            st.markdown(f"- {n}")
        if r["warnings"]:
            with st.expander(f"Template warnings ({len(r['warnings'])})"):
                for cat, msg in r["warnings"]:
                    st.markdown(f"- **{cat}**: {msg}")
        if s["status"] != "OPTIMAL" and s.get("gap") is not None:
            st.caption(f"Searched for {s['seconds']:.0f} s. A longer search time (Options) may "
                       f"improve the schedule slightly.")

st.divider()
st.caption("Files are processed only to build your schedule and are deleted right afterward.")
