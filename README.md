# Cross-Section Sandbox

Turn total-station survey points (or a stations/ground table) into road cross-section
drawings, as DXF and PNG, from a small Streamlit app.

![Full design sheet](demo_builtup/preview_full_design_sheet.png)

## Run

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501 and upload a file in the sidebar.

## Input: total-station CSV (primary workflow)

A point file with `North, East, Height` and a side-code column (`C` / `L` / `R`):

```csv
POINT,North,East,Height,Code
P1,997861.2633,543063.4404,297.3123,C
P2,997868.0328,543061.8938,297.2671,L
P3,997855.5552,543064.9443,297.2683,R
```

The app:

- splits the file into roads at text markers / large coordinate jumps;
- pairs every side shot with its nearest centre-line shot (not by row order);
- computes chainage from cumulative centre-line distance;
- computes true offsets, left negative / right positive, from the L→R axis;
- keeps the surveyed levels exactly as shot.

Feature codes (buildings, culverts, poles, river…) are ignored for the section
geometry; only `C`, `L` and `R` shots are used.

## Drawing modes (sidebar)

| Mode | What it draws |
|---|---|
| **Survey + road design** (default) | Surveyed ground plus the proposed pavement box: pavement layers, formation line, cut/fill hatching clipped to the surveyed width, shot markers, cross-falls, offsets, levels and a shot schedule. Survey geometry is never altered. |
| Surveyed ground only | Exactly the shots in the file — profile, markers, offsets, levels, surveyed width. No pavement. |
| Full design sheet | Complete design drawing: dimensions, drains, layer call-outs, cut/fill areas. |

With no design levels in the file, the pavement is placed at the surveyed centre line
(design level = surveyed CL); this is stated on the sheet and in the sidebar.

### Road width vs survey extent

Two different numbers, both labelled on the sheet:

- **formation width** — the design road, `2 × carriageway + 2 × shoulder`
  (e.g. 10.30 m). This is what the `10.30m` and `1.50 / 3.65 / 3.65 / 1.50` dimensions measure.
- **survey extent** — the distance between the outermost survey shots
  (e.g. 12.84 m). It is labelled `... m survey` and is not the road width.

## Use as a library

Calculations are independent of drawing. `compute_section()` returns a
`SectionResult` containing the ground, formation levels, cut/fill areas, daylight
points and the internal geometry; drawings and schedules read that same object.

```python
import xsection as xs

p = xs.Params(carriageway=3.65, shoulder=1.50, camber=2.5, shoulder_fall=5.0)

res = xs.compute_section(
    chainage=0.0,
    ground=[(-5.15, 99.83), (0.0, 99.83), (5.15, 99.83)],   # (offset m, level m)
    design_level=100.0,
    p=p,
    built_up=True,
)

res.formation_width        # 10.30  (the road)
res.survey_extent          # 10.30  (the shots)
res.cut_area               # m² to the daylight points
res.surveyed_cut_area      # m² clipped to surveyed ground
res.fully_covered          # False when daylight lies beyond the shots
res.daylight               # {-1: 7.34, 1: 7.34}
res.table()                # per-shot: offset, ground, formation, cut/fill depth
xs.section_schedule([res]) # one summary row per section

# draw the same geometry
doc = xs.survey_doc(df, ground_df, p, show_design=True)
open("sections.dxf", "wb").write(xs.doc_bytes(doc))
open("sections.png", "wb").write(xs.render_png(doc))
```

Other frozen calculation helpers: `formation_width(p)`, `pavement_thickness(p)`,
`formation_level(offset, design_level, p)`, `stations(start, end, interval, extra=())`.

`SectionResult` fields: `chainage, design_level, formation_width,
pavement_thickness, ground, formation, cut_area, fill_area, surveyed_cut_area,
surveyed_fill_area, fully_covered, daylight, cut_polygons, fill_polygons, geometry`.


## Input: Stations / Ground / Parameters

One workbook with sheets `Stations`, `Ground`, `Parameters`, or three CSVs named
`stations.csv`, `ground.csv`, `parameters.csv`.

- **Stations**: Station, Design, Existing, Zone = Built-up/Rural
- **Ground**: Station, Offset (−left/+right), Level
- **Parameters**: Name, Value (optional)

`demo_builtup.xlsx` (or `demo_builtup/*.csv`) reproduces the built-up typical sheet:
10.30 m formation, 3.65 m lanes, 1.50 m shoulders, 0.60 m drains, 50/200/200 mm
pavement, 2.5 % camber, 5 % shoulder fall, CBR 15 %. See
[`demo_builtup/README.md`](demo_builtup/README.md) for which values came from the
reference sheet and which were deliberately left to defaults.

## CLI

```bash
python xsection.py your.xlsx --out sections.dxf
python xsection.py stations.csv ground.csv parameters.csv --out sections.dxf
python xsection.py "survey points.csv" --out sections.dxf
```

## Notes

- CSV parsing is tolerant: `,` `;` tab or `|` delimiters, UTF-8/UTF-16/Windows
  encodings, decimal commas (`297,325`), unit suffixes (`297.325 m`, `2.5%`),
  `CH 0+000` style stations, and a title row above the headers.
- The sidebar reports how many stations and ground points were loaded and warns
  about unreadable rows, unknown parameter names, or one-sided sections.
- PNG previews are rendered at 1460 px (Streamlit's limit) and downloads at higher
  resolution. DXF text height is set in the sidebar.
- Streamlit caches imported modules: restart the server after editing `xsection.py`.

## Project layout

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `xsection.py` | Engine: calculations (`compute_section` → `SectionResult`), DXF/PNG drawing, loading, CLI |
| `requirements.txt` | `streamlit, ezdxf, pandas, openpyxl, xlrd, matplotlib, numpy` |
| `demo_builtup.xlsx`, `demo_builtup/` | Test data for the built-up typical sheet |
| `typical_sections.dxf` | Example output |
