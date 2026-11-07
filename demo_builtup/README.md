# Built-up typical-section test data

Reproduces the reference sheet "TYPICAL ROAD CROSS-SECTION (BUILT-UP AREAS)".

Upload **demo_builtup.xlsx** (one file), or the three CSVs in `demo_builtup/`
together (`stations.csv` + `ground.csv` + `parameters.csv`).
Then choose **Drawing mode → Full design sheet**.

## Where each value comes from

Taken directly from the reference sheet:

| Parameter | Value | Source on the sheet |
|---|---|---|
| carriageway | 3.65 m | "3.65m" (each side) |
| shoulder | 1.50 m | "1.50m" (each side) |
| camber | 2.5 % | "2.5 %" arrows |
| shoulder_fall | 5.0 % | "5.0 %" arrows |
| wearing | 50 mm | "50mm Wearing Course" |
| base | 200 mm | "200mm Lateritic Base Course" |
| subbase | 200 mm | "200mm Laterite Sub-base Course" |
| drain_width | 0.60 m | "0.60m" over each drain |
| cbr | 15 % | "Sub-grade CBR must not be less than 15%" |

Deliberately **not** in the file, because the sheet does not state them:

- **surface_dressing thickness** — the sheet only says "Single layer
  Surface-dressing"; the app default (15 mm) applies.
- **drain_depth** — the sheet says "Varies"; the app default (0.60 m) is used and
  the drawing labels the depth "Varies".
- cut/fill slope, right-of-way, drain wall thickness, embankment height,
  vertical exaggeration, text height — no values on the sheet; app defaults apply.

## Level datum

The sheet carries no levels at all. For a drawable test file the design CL is an
arbitrary datum of **100.000 m**, and the existing ground is placed at the
shoulder-edge surface level (100.000 − 0.025×3.65 − 0.05×1.50 ≈ 99.829), which is
what the sheet depicts: flat ground with the drains cut in at the edges.

Chainages 0+000 / 0+025 / 0+050 are simply three copies of the same detail so the
station picker has something to select.
