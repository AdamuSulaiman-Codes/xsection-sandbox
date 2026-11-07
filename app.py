"""Cross-section sandbox UI.  Run:  streamlit run app.py

Note: Streamlit caches imported modules, so restart the server after editing
xsection.py (a rerun alone keeps the old module in memory).
"""
import io

import pandas as pd
import streamlit as st

import xsection as xs

st.set_page_config(page_title="Cross-Section Sandbox", layout="wide")
st.title("Road Cross-Section Sandbox")

# ---------------- sidebar: upload ----------------
sb = st.sidebar
sb.header("Upload")
ups = sb.file_uploader(
    "Total-station CSV, or Excel / stations-ground-parameters files",
    type=["xlsx", "xls", "csv"], accept_multiple_files=True,
)
tpl = io.BytesIO()
xs.make_template(tpl)
sb.download_button("Download workbook template", tpl.getvalue(), "xsection_template.xlsx")
sb.download_button("Download table templates (.zip)", xs.make_csv_template(),
                   "xsection_csv_templates.zip")

pdict = {}
stations, ground = None, None
if ups:
    files = {u.name: u for u in ups}
    notes: list[str] = []
    try:
        stations, ground, pdict = xs.load_uploaded(files, notes)
        parts = [f"{len(stations)} stations"]
        parts.append(f"{len(ground)} ground points" if ground is not None else "no ground points")
        parts.append(f"{len(pdict)} parameter overrides")
        sb.caption("Loaded " + ", ".join(parts) + ".")
        for n in notes:
            sb.warning(n)
    except Exception as e:
        sb.error(f"Could not read file: {e}")
if stations is None:
    stations, ground, pdict = xs.load_workbook(io.BytesIO(tpl.getvalue()))
    sb.info("Showing sample data. Upload your file to replace it.")

base = xs.params_from_dict(pdict)


def num(label, name, step, mn=None, mx=None, fmt="%.2f"):
    v = getattr(base, name)
    if isinstance(v, int):
        return int(sb.number_input(label, value=v, step=int(step)))
    return sb.number_input(label, value=float(v), step=float(step), min_value=None if mn is None else float(mn), max_value=None if mx is None else float(mx), format=fmt)


# ---------------- sidebar: drawing mode ----------------
sb.header("Drawing mode")
mode = sb.radio(
    "What to draw",
    ["Survey + road design", "Surveyed ground only", "Full design sheet"],
    index=0,
    help="Survey + road design draws your surveyed ground with the proposed pavement box "
         "on top (geometry unchanged). Surveyed ground only plots the shots by themselves. "
         "Full design sheet is the older complete drawing with dimensions, drains and "
         "layer call-outs.",
)
show_design = mode.startswith("Survey +")
template = mode.startswith("Full")

p = xs.Params()
if show_design or template:
    sb.header("Road template")
    p.carriageway = num("Carriageway / side (m)", "carriageway", 0.05, 1.0)
    p.shoulder = num("Shoulder / side (m)", "shoulder", 0.05, 0.0)
    p.camber = num("Camber (%)", "camber", 0.1, 0.0, 15.0, "%.1f")
    p.shoulder_fall = num("Shoulder fall (%)", "shoulder_fall", 0.5, 0.0, 20.0, "%.1f")
    sb.header("Pavement (mm)")
    p.wearing = num("Wearing course", "wearing", 5, 0.0, fmt="%.0f")
    p.surface_dressing = num("Surface dressing", "surface_dressing", 1, 0.0, fmt="%.0f")
    p.base = num("Base", "base", 10, 0.0, fmt="%.0f")
    p.subbase = num("Sub-base", "subbase", 10, 0.0, fmt="%.0f")
    p.cbr = num("Sub-grade CBR (%)", "cbr", 1)
if template:
    sb.header("Earthworks")
    p.cut_slope = num("Cut slope (H:1V)", "cut_slope", 0.25, 0.1)
    p.fill_slope = num("Fill slope (H:1V)", "fill_slope", 0.25, 0.1)
    p.drain_width = num("Drain width (m)", "drain_width", 0.05, 0.1)
    p.drain_depth = num("Drain depth (m)", "drain_depth", 0.05, 0.1)
    p.row = num("Right-of-way / side (m)", "row", 1.0, 6.0)
    p.transverse_slope = num("Assumed ground slope when no survey (%)", "transverse_slope", 0.5, -50.0, 50.0, "%.1f")
    p.embankment_h = num("Typical rural embankment height (m)", "embankment_h", 0.25, 0.1)

sb.header("Section drawing")
p.vexag = num("Vertical exaggeration", "vexag", 0.5, 1.0, 10.0, "%.1f")
p.text_h = num("Text height (mm)", "text_h", 10, 40, fmt="%.0f")

# ---------------- main ----------------
if ups:
    st.caption("Using uploaded data: " + ", ".join(files))
else:
    st.caption("No file loaded — upload a total-station CSV in the sidebar.")

tabs = dict(zip(["station", "data"], st.tabs(["Station sections", "Data"])))

# execution order: Data first so edits feed the Station drawings
with tabs["data"]:
    st.caption("Edit values here (live). Offsets are metres from the centre line "
               "(negative = left), levels are metres.")
    st.subheader("Stations")
    stations = st.data_editor(stations, num_rows="dynamic", width="stretch", key="st_ed")
    st.subheader("Ground survey shots")
    gsrc = ground if ground is not None else pd.DataFrame({"Chainage": [], "Offset": [], "Level": []})
    ground = st.data_editor(gsrc, num_rows="dynamic", width="stretch", key="gd_ed")
    if len(ground) == 0:
        ground = None

with tabs["station"]:
    if not ups:
        st.info("Upload your total-station CSV in the sidebar to generate cross-sections "
                "from the surveyed ground.")
    else:
        stations = stations.dropna(subset=["Chainage", "Design", "Existing"])
        labels = [xs.fmt_chainage(c) for c in stations["Chainage"]]
        pick = st.multiselect("Stations to draw", labels, default=labels[:1])
        if pick:
            sel = stations[[l in pick for l in labels]]
            if template:
                doc2 = xs.station_doc(sel, ground, p, cols=1)
            else:
                doc2 = xs.survey_doc(sel, ground, p, cols=1, show_design=show_design)
            st.image(xs.render_png(doc2, width_in=12, dpi=200, max_width_px=1460), width="stretch")
            c1, c2, c3 = st.columns(3)
            c1.download_button("Download DXF (selected)", xs.doc_bytes(doc2), "sections_selected.dxf")
            c2.download_button("Download PNG (selected)", xs.render_png(doc2, width_in=12, dpi=200),
                               "sections_selected.png")
            if c3.button("Build DXF for ALL stations"):
                if template:
                    build = xs.station_doc(stations, ground, p, cols=2)
                else:
                    build = xs.survey_doc(stations, ground, p, cols=2, show_design=show_design)
                st.download_button("Download all.dxf", xs.doc_bytes(build), "sections_all.dxf")

            rows = []
            for _, r in sel.iterrows():
                if template:
                    sec, surveyed, built = xs.station_section(r, ground, p)
                    rows.append({"Chainage": xs.fmt_chainage(r["Chainage"]),
                                 "Zone": "Built-up" if built else "Rural",
                                 "Cut m2": round(sec["cut"], 2), "Fill m2": round(sec["fill"], 2),
                                 "Surveyed": surveyed})
                else:
                    g = xs.station_points(r["Chainage"], ground)
                    row = {"Chainage": xs.fmt_chainage(r["Chainage"])}
                    if show_design:
                        row["Design CL"] = round(float(r["Design"]), 3)
                    if g is None or not len(g):
                        row.update({"Shots": 0, "Left (m)": None, "Right (m)": None,
                                    "Span (m)": None, "Level range (m)": None})
                    else:
                        o, l = g["Offset"], g["Level"]
                        row.update({"Shots": len(g),
                                    "Left (m)": round(min(0.0, o.min()), 2),
                                    "Right (m)": round(max(0.0, o.max()), 2),
                                    "Span (m)": round(o.max() - o.min(), 2),
                                    "Level range (m)": round(l.max() - l.min(), 2)})
                    rows.append(row)
            st.dataframe(pd.DataFrame(rows), width="stretch")
