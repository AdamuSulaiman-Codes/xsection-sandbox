"""
xsection.py - road cross-section engine (geometry + DXF drawing).

Units: geometry in metres; DXF drawn in millimetres (x1000).
Layout of the drawing follows the "TYPICAL ROAD CROSS-SECTION" sheet
(dimension strings, CL, slopes, layer call-outs, sub-grade note, drains).
"""
from __future__ import annotations

import csv
import io
import os
import re
import zipfile
from dataclasses import dataclass, asdict, fields

import ezdxf
import numpy as np
import pandas as pd
from ezdxf.enums import TextEntityAlignment as TA


# --------------------------------------------------------------------------
# Parameters
# --------------------------------------------------------------------------
@dataclass
class Params:
    carriageway: float = 3.65      # m, each side of CL
    shoulder: float = 1.50         # m, each side
    camber: float = 2.5            # % carriageway crossfall
    shoulder_fall: float = 5.0     # % shoulder crossfall
    wearing: float = 50.0          # mm wearing course (carriageway only)
    surface_dressing: float = 15.0 # mm single layer surface dressing
    base: float = 200.0            # mm lateritic base
    subbase: float = 200.0         # mm laterite sub-base
    drain_width: float = 0.60      # m (built-up areas)
    drain_depth: float = 0.60      # m below formation edge
    drain_wall: float = 0.10       # m wall thickness
    cut_slope: float = 1.5         # H : 1V
    fill_slope: float = 2.0        # H : 1V
    row: float = 15.0              # m each side of CL (right-of-way)
    transverse_slope: float = 0.0  # % ground fall to the right when no survey offsets
    embankment_h: float = 1.0      # m, only for the typical rural drawing
    vexag: float = 2.0             # vertical exaggeration (drawing only)
    cbr: int = 15                  # sub-grade CBR note
    text_h: float = 240.0          # mm main text height


PARAM_NAMES = [f.name for f in fields(Params)]


def params_from_dict(d: dict) -> Params:
    p = Params()
    for k, v in d.items():
        if k in PARAM_NAMES:
            setattr(p, k, type(getattr(p, k))(v))
    return p


# --------------------------------------------------------------------------
# Excel I/O
# --------------------------------------------------------------------------
def to_number(v, decimal_comma: bool = True) -> float:
    """Best-effort numeric parse: '297.325 m', '297,325', '1,234.56', '1 234,56'.
    decimal_comma=True reads '297,325' as 297.325 (levels); False reads it as 297325 (chainages)."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return float(v) if pd.notna(v) else np.nan
    s = str(v).strip().replace("\u00a0", " ").replace("\u2212", "-").replace("\u2013", "-")
    if not s or s.lower() in {"nan", "none", "null", "-", "--"}:
        return np.nan
    s = s.replace(" ", "")                       # also normalises digit grouping
    s = re.sub(r"(?i)(mm|cm|km|m|ft|in|deg|%|°)$", "", s).strip()
    if not s or s in {"+", "-", ".", ","}:
        return np.nan
    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):          # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:                                    # 1,234.56
            s = s.replace(",", "")
    elif has_comma:
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3}){2,}", s):   # 1,234,567
            s = s.replace(",", "")
        elif decimal_comma:                                # 297,325 -> 297.325
            s = s.replace(",", ".")
        else:                                              # chainage 1,025 -> 1025
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return np.nan


def parse_chainage(v) -> float:
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\u00a0", "")
    s = re.sub(r"(?i)^(ch|sta|stn|km|k|p)\.?", "", s)   # "CH 0+000"
    m = re.fullmatch(r"(\d+)\+(\d+(?:[.,]\d+)?)", s)
    if m:
        rest = to_number(m.group(2))                # 0+325,5 is a decimal comma
        return int(m.group(1)) * 1000 + (0.0 if pd.isna(rest) else rest)
    n = to_number(s, decimal_comma=False)           # plain 1,025 -> 1025
    if pd.isna(n):
        raise ValueError(f"unrecognised chainage {v!r}")
    return n


def fmt_chainage(x: float) -> str:
    k = int(x // 1000)
    r = x - k * 1000
    return f"{k}+{int(round(r)):03d}" if abs(r - round(r)) < 1e-6 else f"{k}+{r:06.2f}"


def _norm(col) -> str:
    return str(col).lower().replace(" ", "").replace("_", "")


def _find(cols, keys):
    for c in cols:
        lc = _norm(c)
        if lc.startswith("unnamed"):        # Excel/CSV index column
            continue
        if any(k in lc for k in keys):
            return c
    return None


def load_workbook(src, notes: list | None = None):
    """Return (stations_df, ground_df|None, params_dict).
    Sheets: Stations (Station, Design, Existing, [Zone]), Ground (Station, Offset, Level),
    Parameters (Name, Value). A total-station point sheet is converted too."""
    xl = pd.ExcelFile(src)
    names = {n.lower(): n for n in xl.sheet_names}
    first = pd.read_excel(xl, sheet_name=names.get("stations", xl.sheet_names[0]))
    if _is_survey_points(first):
        return survey_points_to_model(first, notes)
    out = _stations_table(first, notes)
    ground = _ground_table(pd.read_excel(xl, sheet_name=names["ground"]), notes) if "ground" in names else None
    pdict = _params_table(pd.read_excel(xl, sheet_name=names["parameters"]), notes) if "parameters" in names else {}
    return out, ground, pdict


def _stations_table(df: pd.DataFrame, notes: list | None = None) -> pd.DataFrame:
    cs = _find(df.columns, ["station", "chainage", "ch"])
    cd = _find(df.columns, ["design", "elevdesign", "proposed", "formation"])
    ce = _find(df.columns, ["exist", "ground", "egl"])
    cz = _find(df.columns, ["zone", "area", "type"])
    if not (cs and cd and ce):
        raise ValueError(
            "Stations columns not found. Detected: " + ", ".join(map(str, df.columns)) +
            ". Need Station (or Chainage), Design level and Existing level.")
    out = pd.DataFrame({
        "Chainage": [_chainage_or_nan(v) for v in df[cs]],
        "Design": [to_number(v) for v in df[cd]],
        "Existing": [to_number(v) for v in df[ce]],
        "Zone": _zone_series(df[cz]) if cz else "Built-up",
    })
    if out.empty:
        raise ValueError(_no_rows_message(df, cs, cd, ce))
    usable = out.dropna(subset=["Chainage", "Design", "Existing"]).reset_index(drop=True)
    dropped = len(out) - len(usable)
    if not len(usable):
        raise ValueError(_no_rows_message(df, cs, cd, ce))
    if dropped and notes is not None:
        bad = _first_bad(df, {"Station": cs, "Design": cd, "Existing": ce})
        notes.append(f"{dropped} station row(s) skipped as unreadable ({bad}).")
    return usable


def _chainage_or_nan(v) -> float:
    try:
        return parse_chainage(v)
    except (ValueError, TypeError):
        return np.nan


def _zone_series(s: pd.Series) -> pd.Series:
    z = s.astype(str).str.strip()
    return z.where(~z.str.lower().isin({"", "nan", "none"}), "Built-up")


def _first_bad(df: pd.DataFrame, cols: dict) -> str:
    picks = []
    for label, col in cols.items():
        for v in df[col]:
            try:
                if label == "Station":
                    parse_chainage(v)
                elif pd.isna(to_number(v)):
                    picks.append(f"{label}={v!r}")
                    break
            except (ValueError, TypeError):
                picks.append(f"{label}={v!r}")
                break
    return ", ".join(picks) if picks else "blank cells"


def _no_rows_message(df: pd.DataFrame, cs, cd, ce) -> str:
    sample = _first_bad(df, {"Station": cs, "Design": cd, "Existing": ce})
    return (f"No station rows could be read from {len(df)} row(s). "
            f"Columns used: {cs} / {cd} / {ce}. Unreadable examples: {sample}. "
            "Levels must be numbers (297.325) and stations like 0+000, CH 0+000 or 25.")


def _ground_table(df: pd.DataFrame, notes: list | None = None):
    gs = _find(df.columns, ["station", "chainage", "ch"])
    go = _find(df.columns, ["offset", "dist"])
    gl = _find(df.columns, ["level", "elev", "rl"])
    if not (gs and go and gl):
        if notes is not None:
            notes.append("Ground file ignored: needs Station, Offset and Level columns.")
        return None
    out = pd.DataFrame({
        "Chainage": [_chainage_or_nan(v) for v in df[gs]],
        "Offset": [to_number(v) for v in df[go]],
        "Level": [to_number(v) for v in df[gl]],
    }).dropna().reset_index(drop=True)
    if out.empty:
        if notes is not None:
            notes.append("Ground file has no usable rows; sections will use assumed ground.")
        return None
    return out


def _params_table(df: pd.DataFrame, notes: list | None = None) -> dict:
    kn = _find(df.columns, ["name", "param", "key"])
    kv = _find(df.columns, ["value", "val"])
    if not (kn and kv):
        if df.shape[1] < 2:
            return {}
        kn, kv = df.columns[0], df.columns[1]
    pdict, unknown = {}, []
    for k, v in zip(df[kn], df[kv]):
        name = str(k).strip()
        if name in PARAM_NAMES:
            num = to_number(v)
            if pd.notna(num):
                pdict[name] = num
        elif name and name.lower() not in {"nan", "none", "name", "parameter"}:
            unknown.append(name)
    if notes is not None and unknown:
        notes.append("Ignored unknown parameter name(s): " + ", ".join(unknown[:6]))
    return pdict


_HEADER_HINTS = ("station", "chainage", "ch", "offset", "dist", "level", "elev",
                 "rl", "design", "exist", "zone", "name", "value")


# --------------------------------------------------------------------------
# Total-station point import (North/East/Height + C/L/R codes)
# --------------------------------------------------------------------------
_SIDE_ALIASES = {
    "c": "C", "cl": "C", "c/l": "C", "cln": "C", "center": "C", "centre": "C",
    "l": "L", "lt": "L", "left": "L",
    "r": "R", "rt": "R", "right": "R",
}


def _col_by_name(df, keys):
    for c in df.columns:
        lc = _norm(c)
        if lc.startswith("unnamed"):
            continue
        if any(k in lc for k in keys):
            return c
    return None


def _survey_code_col(df):
    best, best_frac = None, 0.0
    for c in df.columns:
        s = df[c].dropna().astype(str).str.strip().str.lower()
        if s.empty:
            continue
        frac = float(s.isin(_SIDE_ALIASES).mean())
        if frac > best_frac:
            best, best_frac = c, frac
    return best if best_frac >= 0.5 else None


def survey_columns(df):
    """Return (north, east, height, code) for a total-station point table, else None."""
    north = _col_by_name(df, ["northing", "north"])
    east = _col_by_name(df, ["easting", "east"])
    height = _col_by_name(df, ["height", "elevation", "elev", "level", "rl"])
    code = _survey_code_col(df)
    if north and east and height and code:
        return north, east, height, code
    return None


def _is_survey_points(df) -> bool:
    cols = survey_columns(df)
    if not cols or len(df) < 6:
        return False
    sides = df[cols[3]].astype(str).str.strip().str.lower().map(_SIDE_ALIASES)
    return int((sides == "C").sum()) >= 3


def _assign_sides(pts: pd.DataFrame, centre: pd.DataFrame, maxd: float) -> dict:
    """Map each side shot to its nearest centre-line station (one shot per station)."""
    out = {}
    if pts.empty or centre.empty:
        return out
    cp = centre[["N", "E"]].to_numpy(float)
    for i, r in pts.iterrows():
        d = np.hypot(cp[:, 0] - r["N"], cp[:, 1] - r["E"])
        j = int(d.argmin())
        if d[j] <= maxd and (j not in out or d[j] < out[j][1]):
            out[j] = (i, float(d[j]))
    return {j: i for j, (i, _) in out.items()}


def survey_points_to_model(df: pd.DataFrame, notes: list | None = None):
    """Convert a total-station point list (C/L/R coded) into stations + ground.
    Chainage = cumulative distance between centre-line shots; the design level is
    taken as the surveyed centre-line level (no design levels exist in the file)."""
    cols = survey_columns(df)
    if not cols:
        raise ValueError("Not a survey point table (needs North, East, Height and C/L/R codes)")
    north, east, height, code = cols
    work = pd.DataFrame({
        "N": pd.to_numeric(df[north], errors="coerce"),
        "E": pd.to_numeric(df[east], errors="coerce"),
        "H": pd.to_numeric(df[height], errors="coerce"),
        "Side": df[code].astype(str).str.strip().str.lower().map(_SIDE_ALIASES),
    })
    # text rows (road names, legend) break the data into segments
    work["Seg"] = (work["N"].isna() & work["E"].isna()).cumsum()
    work = work.dropna(subset=["N", "E", "H"])
    work = work[work["Side"].notna()].reset_index(drop=True)

    rows, ground, road_ranges = [], [], []
    ch, one_sided = 0.0, 0
    for seg, g in work.groupby("Seg", sort=True):
        c = g[g["Side"] == "C"].reset_index(drop=True)
        if c.empty:
            continue
        cp = c[["N", "E"]].to_numpy(float)
        spacing = np.hypot(*np.diff(cp, axis=0).T) if len(cp) > 1 else np.array([25.0])
        maxd = float(np.clip(np.median(spacing) * 0.7, 8.0, 25.0))
        lmap = _assign_sides(g[g["Side"] == "L"], c, maxd)
        rmap = _assign_sides(g[g["Side"] == "R"], c, maxd)
        start = ch
        prev = None
        for j in range(len(c)):
            row = c.iloc[j]
            if prev is not None:
                ch += float(np.hypot(row["N"] - prev["N"], row["E"] - prev["E"]))
            li, ri = lmap.get(j), rmap.get(j)
            cvec = np.array([row["E"], row["N"]])

            def vec(p):
                return np.array([p["E"], p["N"]]) - cvec

            if li is not None and ri is not None:
                lp, rp = g.loc[li], g.loc[ri]
                u = vec(rp) - vec(lp)
                norm = np.linalg.norm(u)
                u = u / norm if norm else np.array([1.0, 0.0])
            else:
                p0, p1 = c.iloc[max(j - 1, 0)], c.iloc[min(j + 1, len(c) - 1)]
                t = np.array([p1["E"] - p0["E"], p1["N"] - p0["N"]])
                n = np.linalg.norm(t)
                u = np.array([t[1], -t[0]]) / n if n else np.array([1.0, 0.0])
                known = g.loc[li] if li is not None else (g.loc[ri] if ri is not None else None)
                if known is not None:
                    sgn = np.dot(vec(known), u)
                    if (li is not None and sgn > 0) or (ri is not None and sgn < 0):
                        u = -u
            if li is not None and ri is not None and np.dot(vec(g.loc[ri]), u) < 0:
                u = -u

            ch = round(ch, 3)
            design = round(float(row["H"]), 3)
            rows.append({"Chainage": ch, "Design": design, "Existing": design, "Zone": "Rural"})
            pts = [(0.0, design, "C")]
            if li is not None:
                pts.append((round(float(np.dot(vec(g.loc[li]), u)), 3),
                            round(float(g.loc[li]["H"]), 3), "L"))
            if ri is not None:
                pts.append((round(float(np.dot(vec(g.loc[ri]), u)), 3),
                            round(float(g.loc[ri]["H"]), 3), "R"))
            if len(pts) < 3:
                one_sided += 1
            for off, lvl, code in pts:
                ground.append({"Chainage": ch, "Offset": off, "Level": lvl, "Code": code})
            prev = row
        road_ranges.append((start, ch, int(len(c))))
        ch += 25.0                      # keep roads apart in one chainage list

    if not rows:
        raise ValueError("Survey file has no centre-line (C) shots to build sections from")
    stations = pd.DataFrame(rows)
    gdf = pd.DataFrame(ground)
    max_off = float(gdf["Offset"].abs().max())
    row_w = float(np.clip(np.ceil(max_off) + 2.0, 8.0, 20.0))
    if notes is not None:
        spans = "; ".join(f"Road {i + 1}: {fmt_chainage(a)}-{fmt_chainage(b)} ({n} st)"
                          for i, (a, b, n) in enumerate(road_ranges))
        notes.append(f"Survey point file converted: {len(stations)} sections. {spans}.")
        notes.append("Road-template mode only: design level = surveyed centre line, Zone = Rural, "
                     f"right-of-way {row_w:.0f} m (survey reaches ±{max_off:.1f} m).")
        if one_sided:
            notes.append(f"{one_sided} section(s) have only one side shot.")
    return stations, gdf, {"row": row_w}



def _read_bytes(src) -> bytes:
    if isinstance(src, (str, os.PathLike)):
        with open(src, "rb") as fh:
            return fh.read()
    if hasattr(src, "seek"):
        src.seek(0)
    return src.read()


def _decode_csv(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("latin-1", errors="replace")


def _has_known_header(df: pd.DataFrame) -> bool:
    cols = [str(c).lower().replace(" ", "").replace("_", "") for c in df.columns]
    cols = [c for c in cols if not c.startswith("unnamed")]
    return any(any(h in c for h in _HEADER_HINTS) for c in cols)


def _field_count(line: str, sep: str) -> int:
    try:
        return len(next(csv.reader([line], delimiter=sep)))
    except csv.Error:
        return line.count(sep) + 1


def _detect_sep(text: str) -> str:
    """Pick the delimiter that splits every line the same way (handles decimal commas)."""
    lines = [ln for ln in text.splitlines() if ln.strip()][:25]
    best_sep, best = ",", (-1.0, 0, 0)
    for sep in (",", ";", "\t", "|"):
        counts = [_field_count(ln, sep) for ln in lines]
        if not counts:
            continue
        modal = max(set(counts), key=counts.count)
        if modal < 2:
            continue
        score = (counts.count(modal) / len(counts), int(sep in lines[0]) if lines else 0, modal)
        if score > best:
            best_sep, best = sep, score
    return best_sep


def _parse_csv_text(text: str, **kwargs) -> pd.DataFrame:
    return pd.read_csv(io.StringIO(text), sep=_detect_sep(text), **kwargs)


def _read_csv(src) -> pd.DataFrame:
    """Read a CSV regardless of delimiter, encoding, or title rows above the header."""
    text = _decode_csv(_read_bytes(src))
    try:
        df = _parse_csv_text(text)
    except pd.errors.EmptyDataError:
        raise ValueError("CSV file is empty") from None
    if not _has_known_header(df):
        # some exports put a title/blank line above the column headers
        for h in (1, 2, 3):
            try:
                cand = _parse_csv_text(text, header=h)
            except (pd.errors.ParserError, pd.errors.EmptyDataError):
                break
            if _has_known_header(cand):
                return cand
    return df


def _table_from_name(name) -> str | None:
    """Guess which table a CSV holds from its filename (stations/ground/parameters)."""
    base = str(name).lower().replace("\\", "/").rsplit("/", 1)[-1].rsplit(".", 1)[0]
    if "ground" in base:
        return "ground"
    if "param" in base:
        return "parameters"
    if "station" in base or "chainage" in base:
        return "stations"
    return None


def load_csv_files(files: dict, notes: list | None = None):
    """Return (stations_df, ground_df|None, params_dict) from {name: path|buffer}.
    Names must identify the table: stations.csv, ground.csv, parameters.csv.
    A total-station point file (North/East/Height + C/L/R) is auto-detected."""
    tables, unknown = {}, []
    for name, src in files.items():
        df = _read_csv(src)
        if _is_survey_points(df):
            return survey_points_to_model(df, notes)
        key = _table_from_name(name)
        if key is None:
            unknown.append((name, df))
        elif key not in tables:
            tables[key] = df
    if "stations" not in tables:
        if len(files) == 1 and unknown:
            # single untitled CSV: treat it as the stations table
            tables["stations"] = unknown[0][1]
        elif unknown:
            raise ValueError(
                "Could not tell which table each CSV holds: " + ", ".join(n for n, _ in unknown) +
                ". Name them stations.csv, ground.csv and parameters.csv.")
        else:
            raise ValueError("No stations CSV found (expected stations.csv).")
    out = _stations_table(tables["stations"], notes)
    ground = _ground_table(tables["ground"], notes) if "ground" in tables else None
    pdict = _params_table(tables["parameters"], notes) if "parameters" in tables else {}
    return out, ground, pdict


def _is_workbook_bytes(data: bytes) -> bool:
    return data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0"   # xlsx/zip, xls/ole2


def load_uploaded(files: dict, notes: list | None = None):
    """Return (stations_df, ground_df|None, params_dict) for any mix of Excel/CSV.
    Any workbook wins; otherwise the CSVs are combined by filename.
    `notes`, when given, collects human-readable warnings."""
    if not files:
        raise ValueError("No file supplied")
    csvs, books = {}, {}
    for name, src in files.items():
        if str(name).lower().endswith(".csv") and not _is_workbook_bytes(_read_bytes(src)):
            csvs[name] = src
        else:
            books[name] = src
    if books:
        return load_workbook(next(iter(books.values())), notes)
    return load_csv_files(csvs, notes)


def make_template(path_or_buf, params: Params | None = None):
    """Sample workbook (first stations of Road 1 from the setting-out sheet)."""
    st, gd, pr = _template_frames(params)
    with pd.ExcelWriter(path_or_buf, engine="openpyxl") as w:
        st.to_excel(w, sheet_name="Stations", index=False)
        gd.to_excel(w, sheet_name="Ground", index=False)
        pr.to_excel(w, sheet_name="Parameters", index=False)


def make_csv_template(params: Params | None = None) -> bytes:
    """Same sample data as three CSVs zipped together (stations/ground/parameters)."""
    st, gd, pr = _template_frames(params)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("stations.csv", st.to_csv(index=False))
        z.writestr("ground.csv", gd.to_csv(index=False))
        z.writestr("parameters.csv", pr.to_csv(index=False))
    return buf.getvalue()


def _template_frames(params: Params | None = None):
    rows = [
        (0, 297.325, 297.325), (25, 297.195, 296.862), (50, 297.065, 296.847),
        (75, 296.935, 296.801), (100, 296.805, 296.716), (125, 296.675, 296.570),
        (150, 296.545, 296.213), (175, 296.415, 295.741), (200, 296.285, 295.595),
        (225, 296.155, 295.433), (250, 296.025, 295.296), (275, 295.896, 295.122),
        (300, 295.766, 294.840), (325, 295.636, 294.704), (350, 295.506, 294.604),
        (375, 295.376, 294.406), (400, 295.246, 294.354), (425, 295.116, 294.322),
        (450, 294.986, 294.438), (475, 294.856, 294.450), (500, 294.733, 294.227),
    ]
    st = pd.DataFrame({
        "Station": [fmt_chainage(r[0]) for r in rows],
        "Design": [r[1] for r in rows],
        "Existing": [r[2] for r in rows],
        "Zone": ["Built-up" if r[0] <= 400 else "Rural" for r in rows],
    })
    # EXAMPLE survey offsets for two stations (replace with real field data)
    g = []
    for ch, base_l, slope in [(100, 296.716, -1.0), (300, 294.840, 2.0)]:
        for off in (-15, -10, -7.5, -5.15, -2.5, 0, 2.5, 5.15, 7.5, 10, 15):
            g.append((fmt_chainage(ch), off, round(base_l - slope / 100 * off, 3)))
    gd = pd.DataFrame(g, columns=["Station", "Offset", "Level"])
    p = params or Params()
    pr = pd.DataFrame({"Name": PARAM_NAMES, "Value": [getattr(p, n) for n in PARAM_NAMES]})
    return st, gd, pr


# --------------------------------------------------------------------------
# Geometry
# --------------------------------------------------------------------------
def ground_for(ch, ground_df, existing, p: Params):
    """Ground line (x, y) covering +-row. Uses survey offsets if provided."""
    if ground_df is not None:
        g = ground_df[np.isclose(ground_df["Chainage"], ch)].sort_values("Offset")
        if len(g) >= 2:
            gx, gy = g["Offset"].to_numpy(float), g["Level"].to_numpy(float)
            if gx[0] > -p.row:
                gx, gy = np.r_[-p.row, gx], np.r_[gy[0], gy]
            if gx[-1] < p.row:
                gx, gy = np.r_[gx, p.row], np.r_[gy, gy[-1]]
            return gx, gy, True
    gx = np.array([-p.row, p.row])
    gy = existing - p.transverse_slope / 100 * gx
    return gx, gy, False


def build_section(L, gx, gy, p: Params, built_up: bool):
    c, s = p.carriageway, p.shoulder
    xe = c + s
    cam, fall = p.camber / 100, p.shoulder_fall / 100
    wear, sd, base, sub = (v / 1000 for v in (p.wearing, p.surface_dressing, p.base, p.subbase))
    T = wear + sd + base + sub

    def surf(x):
        x = np.abs(x)
        return np.where(x <= c, L - cam * x, L - cam * c - fall * (x - c))

    def S(x):
        return float(surf(x))

    # layer polygons: bands parallel to the surface, per zone -------------
    def band(xs, d0, d1):
        return ([(x, S(x) - d0) for x in xs] + [(x, S(x) - d1) for x in reversed(xs)])
    cz = [-c, 0, c]
    zones = [(cz, [0, wear, wear + sd, wear + sd + base, T]),
             ([-xe, -c], [0, 0, sd, T - sub, T]),
             ([c, xe], [0, 0, sd, T - sub, T])]
    layers = {"wearing": [], "dressing": [], "base": [], "subbase": []}
    for xs, d in zones:
        if d[1] > d[0]:
            layers["wearing"].append(band(xs, d[0], d[1]))
        layers["dressing"].append(band(xs, d[1], d[2]))
        layers["base"].append(band(xs, d[2], d[3]))
        layers["subbase"].append(band(xs, d[3], d[4]))

    gfun = lambda x: np.interp(x, gx, gy)
    step = 0.01
    sides = {}
    for sgn in (-1, 1):
        Fe = S(xe) - T
        ge = float(gfun(sgn * xe))
        mode = "fill" if ge < Fe - 1e-9 else "cut"
        drain = built_up and mode == "cut"
        xo = np.arange(xe, p.row + step, step)
        g_o = gfun(sgn * xo)
        if mode == "fill":
            top = Fe - (xo - xe) / p.fill_slope
            verts = [(xe, Fe)]
        elif drain:
            yb = Fe - p.drain_depth
            top = np.where(xo < xe + p.drain_width, yb, yb + (xo - xe - p.drain_width) / p.cut_slope)
            verts = [(xe, Fe), (xe, yb), (xe + p.drain_width, yb)]
        else:
            top = Fe + (xo - xe) / p.cut_slope
            verts = [(xe, Fe)]
        diff = top - g_o
        sg = np.sign(diff[0]) if diff[0] != 0 else 0
        idx = np.where(np.sign(diff) != sg)[0]
        if sg == 0:
            dx = xe
        elif len(idx):
            i = idx[0]
            t = diff[i - 1] / (diff[i - 1] - diff[i]) if diff[i - 1] != diff[i] else 0
            dx = xo[i - 1] + t * step
        else:
            dx = p.row
        yd = float(gfun(sgn * dx))
        verts.append((dx, yd))
        sides[sgn] = dict(mode=mode, drain=drain, daylight=dx, verts=verts,
                          Fe=Fe, top=top, xo=xo, g_o=g_o)

    # drain outlines
    drains = []
    for sgn, sd_ in sides.items():
        if sd_["drain"]:
            yb = sd_["Fe"] - p.drain_depth
            yo = float(gfun(sgn * (xe + p.drain_width)))
            t = p.drain_wall
            u = [(xe, S(xe)), (xe, yb), (xe + p.drain_width, yb), (xe + p.drain_width, yo),
                 (xe + p.drain_width - t, yo), (xe + p.drain_width - t, yb + t),
                 (xe + t, yb + t), (xe + t, S(xe))]
            drains.append([(sgn * x, y) for x, y in u])

    # full formation profile + areas
    dl, dr = sides[-1]["daylight"], sides[1]["daylight"]
    xp = np.arange(-xe, xe + step / 2, step)
    xl = -sides[-1]["xo"][sides[-1]["xo"] <= dl][::-1]
    xr = sides[1]["xo"][sides[1]["xo"] <= dr]
    tl = sides[-1]["top"][: len(xl)][::-1]
    tr = sides[1]["top"][: len(xr)]
    xf = np.r_[xl, xp, xr]
    tf = np.r_[tl, surf(xp) - T, tr]
    gf = gfun(xf)
    cut_a = float(np.trapezoid(np.clip(gf - tf, 0, None), xf))
    fill_a = float(np.trapezoid(np.clip(tf - gf, 0, None), xf))
    m = np.abs(xf) <= xe
    cut_p = float(np.trapezoid(np.clip(gf[m] - tf[m], 0, None), xf[m]))
    fill_p = float(np.trapezoid(np.clip(tf[m] - gf[m], 0, None), xf[m]))

    def runs(mask):
        out, i = [], 0
        while i < len(mask):
            if mask[i]:
                j = i
                while j + 1 < len(mask) and mask[j + 1]:
                    j += 1
                if j > i:
                    out.append([(xf[k], tf[k]) for k in range(i, j + 1)] +
                               [(xf[k], gf[k]) for k in range(j, i - 1, -1)])
                i = j + 1
            else:
                i += 1
        return out
    dwid = np.where(xf < 0, p.drain_width if sides[-1]["drain"] else 0.0,
                    p.drain_width if sides[1]["drain"] else 0.0)
    outside = np.abs(xf) > xe + dwid
    fill_polys = runs(tf > gf + 1e-6)
    cut_polys = runs((gf > tf + 1e-6) & outside)

    return dict(L=L, xe=xe, T=T, surf=S, layers=layers, drains=drains, sides=sides,
                gfun=gfun, gx=gx, gy=gy, xf=xf, tf=tf, gf=gf,
                dl=dl, dr=dr, cut=cut_a, fill=fill_a, cut_p=cut_p, fill_p=fill_p,
                fill_polys=fill_polys, cut_polys=cut_polys, built_up=built_up,
                wear=wear, sd=sd, base=base, sub=sub)


# --------------------------------------------------------------------------
# DXF drawing
# --------------------------------------------------------------------------
def new_doc():
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    for name, col in [("ROAD", 7), ("PAVE", 7), ("GROUND", 30), ("EARTH", 30), ("SUBGRADE", 30),
                      ("DIM", 7), ("TEXT", 7), ("CL", 1), ("DRAIN", 7), ("NOTE", 250)]:
        if name not in doc.layers:
            doc.layers.add(name, color=col)
    doc.styles.get("Standard").dxf.font = "arial.ttf"
    return doc


def _txt(msp, s, x, y, h, align=TA.MIDDLE_CENTER, layer="TEXT", color=None, rot=0, bold=False):
    t = msp.add_text(s, dxfattribs={"height": h, "layer": layer, "rotation": rot,
                                    **({"color": color} if color else {})})
    t.set_placement((x, y), align=align)
    return t


def _line(msp, a, b, layer="DIM", color=None, lw=None):
    at = {"layer": layer}
    if color:
        at["color"] = color
    if lw:
        at["lineweight"] = lw
    return msp.add_line(a, b, dxfattribs=at)


def _dim_h(msp, x1, x2, y, text, h, y_from, layer="DIM"):
    _line(msp, (x1, y_from), (x1, y + 120))
    _line(msp, (x2, y_from), (x2, y + 120))
    gap = len(text) * h * 0.30 + 70
    xm = (x1 + x2) / 2
    if x2 - x1 > 2 * gap + 200:
        _line(msp, (x1, y), (xm - gap, y))
        _line(msp, (xm + gap, y), (x2, y))
    else:
        _line(msp, (x1, y), (x2, y))
    for x in (x1, x2):
        _line(msp, (x - 110, y - 110), (x + 110, y + 110), lw=35)
    _txt(msp, text, xm, y, h, layer=layer)


def _arrow(msp, x0, x1, y, color=8):
    _line(msp, (x0, y), (x1, y), "NOTE", color)
    d = 1 if x1 > x0 else -1
    _line(msp, (x1, y), (x1 - d * 90, y + 35), "NOTE", color)
    _line(msp, (x1, y), (x1 - d * 90, y - 35), "NOTE", color)


def _hatch(msp, pts, kind, layer="PAVE"):
    h = msp.add_hatch(color=7, dxfattribs={"layer": layer})
    if kind == "solid_dark":
        h.set_solid_fill(color=250)
    elif kind == "solid":
        h.set_solid_fill(color=7)
    else:
        name, scale, ang, col = kind
        h.set_pattern_fill(name, color=col, scale=scale, angle=ang)
    h.paths.add_polyline_path(pts, is_closed=True)
    return h


def draw_section(msp, sec, p: Params, ox=0.0, oy=0.0, title="", subtitle="", info=()):
    """Draw one section; CL top-of-pavement sits at (ox, oy). Returns bottom y (mm)."""
    Lref = sec["L"]
    k = 1000.0
    ky = k * p.vexag
    X = lambda x: ox + x * k
    Y = lambda y: oy + (y - Lref) * ky
    P = lambda pt: (X(pt[0]), Y(pt[1]))
    h = p.text_h
    xe = sec["xe"]
    c = p.carriageway

    # earthworks hatch (fill / cut)
    for poly in sec["fill_polys"]:
        _hatch(msp, [P(q) for q in poly], ("ANSI31", 90, 45, 9), "EARTH")
    for poly in sec["cut_polys"]:
        _hatch(msp, [P(q) for q in poly], ("ANSI37", 140, 0, 30), "EARTH")

    # pavement layers
    kinds = {"wearing": "solid_dark", "dressing": "solid", "base": ("AR-CONC", 4, 0, 7),
             "subbase": ("ANSI31", 30, 45, 8)}
    for name, polys in sec["layers"].items():
        for poly in polys:
            pts = [P(q) for q in poly]
            _hatch(msp, pts, kinds[name])
            msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": "PAVE", "lineweight": 25})

    # drains
    for d in sec["drains"]:
        msp.add_lwpolyline([P(q) for q in d], close=True, dxfattribs={"layer": "DRAIN", "lineweight": 35})
    for sgn, sd_ in sec["sides"].items():
        if sd_["drain"]:
            xa, xb = sgn * xe, sgn * (xe + p.drain_width)
            yb_ = sd_["Fe"] - p.drain_depth
            go = float(sec["gfun"](xb))
            _txt(msp, "DRAIN", X((xa + xb) / 2), Y(yb_ + p.drain_depth * 0.55), h * 0.38, TA.MIDDLE_CENTER, "NOTE")
            # width dimension above ground
            yy = Y(max(go, sec["surf"](xe))) + 800
            x1_, x2_ = sorted((X(xa), X(xb)))
            _line(msp, (x1_, yy - 650), (x1_, yy + 120))
            _line(msp, (x2_, yy - 650), (x2_, yy + 120))
            _line(msp, (x1_, yy), (x2_, yy))
            for xx in (x1_, x2_):
                _line(msp, (xx - 90, yy - 90), (xx + 90, yy + 90), lw=35)
            _txt(msp, f"{p.drain_width:.2f}m", sgn * 1 * 0 + (x1_ + x2_) / 2 + sgn * 330, yy + h * 1.0, h * 1.25,
                 TA.MIDDLE_CENTER, "DIM")
            # "Varies" vertical dimension on outer wall
            xv = X(sgn * (xe + p.drain_width + 0.35))
            y_top_, y_bot_ = Y(go), Y(yb_)
            _line(msp, (xv, y_top_), (xv, y_bot_), "DIM")
            for yy_ in (y_top_, y_bot_):
                _line(msp, (xv - 90, yy_ - 90), (xv + 90, yy_ + 90), lw=35)
            _txt(msp, "Varies", xv + sgn * 220, (y_top_ + y_bot_) / 2 - 150, h * 1.4,
                 TA.MIDDLE_LEFT if sgn > 0 else TA.MIDDLE_RIGHT, "TEXT")

    # formation / sub-grade line (dashed orange) incl. crown under the pavement
    pts = []
    for sgn in (-1, 1):
        pts.append([(sgn * x, y) for x, y in sec["sides"][sgn]["verts"]])
    full = list(reversed(pts[0])) + [(0.0, Lref - sec["T"])] + pts[1]
    msp.add_lwpolyline([P(q) for q in full], dxfattribs={"layer": "SUBGRADE", "linetype": "DASHED",
                                                         "ltscale": 40, "lineweight": 25})
    # ground line
    x0, x1 = -sec["dl"] - 1.2, sec["dr"] + 1.2
    gxs = np.r_[x0, sec["gx"][(sec["gx"] > x0) & (sec["gx"] < x1)], x1]
    gl = [(x, float(sec["gfun"](x))) for x in gxs]
    msp.add_lwpolyline([P(q) for q in gl], dxfattribs={"layer": "GROUND", "lineweight": 50})
    for sgn in (-1, 1):
        xt_m = sgn * (max(sec["dl"], sec["dr"]) + 1.9)
        yt = Y(float(sec["gfun"](xt_m))) + 330
        _txt(msp, "Exist. Ground", X(xt_m), yt, h * 1.1, TA.MIDDLE_CENTER, "NOTE")

    ymin = min([Y(q[1]) for q in full] + [Y(float(sec["gfun"](0)))] +
               [Y(q[1]) for d in sec["drains"] for q in d])
    ytop = Y(Lref)

    # CL
    _line(msp, (ox, ymin - 1900), (ox, ytop + 2900), "CL", 1)
    _txt(msp, "CL", ox + 110, ytop + 330, h, TA.MIDDLE_LEFT, "CL", 1)

    # dimension strings (overall on top, then segments)
    ef = ytop + 500
    y2 = ytop + 1900
    y1 = ytop + 2500
    _dim_h(msp, X(-xe), X(xe), y1, f"{2 * xe:.2f}m", h * 1.4, ef)
    for a, b, w in [(-xe, -c, p.shoulder), (-c, 0, c), (0, c, c), (c, xe, p.shoulder)]:
        _dim_h(msp, X(a), X(b), y2, f"{w:.2f}m", h * 1.4, ef)

    # LHS / RHS, zone labels, crossfalls
    for xm, t in ((-c / 2, "LHS"), (c / 2, "RHS")):
        _txt(msp, t, X(xm), ytop + 1450, h * 0.55, TA.MIDDLE_CENTER, "NOTE")
    _line(msp, (X(-c / 2), ytop + 1300), (X(-c / 2), ytop + 1050), "NOTE", 8)
    _line(msp, (X(c / 2), ytop + 1050), (X(c / 2), ytop + 1300), "NOTE", 8)
    for sgn in (-1, 1):
        xm = sgn * (c + xe) / 2
        _txt(msp, "SHOULDER", X(xm), ytop + 800, h * 0.62, TA.MIDDLE_CENTER, "NOTE")
        _txt(msp, "CARRIAGEWAY", X(sgn * c / 2), ytop + 800, h * 0.62, TA.MIDDLE_CENTER, "NOTE")
        _txt(msp, f"{p.shoulder_fall:.1f}%", X(xm), Y(sec["surf"](abs(xm))) + 430, h * 1.1, TA.MIDDLE_CENTER, "TEXT")
        _arrow(msp, X(sgn * (c + 0.15)), X(sgn * (xe - 0.15)), Y(sec["surf"](abs(xm))) + 230)
        _txt(msp, f"{p.camber:.1f}%", X(sgn * c / 2), Y(sec["surf"](c / 2)) + 430, h * 1.1, TA.MIDDLE_CENTER, "TEXT")
        _arrow(msp, X(sgn * 0.3), X(sgn * (c - 0.3)), Y(sec["surf"](c / 2)) + 230)

    # sub-grade notes
    _txt(msp, "SUB-GRADE", X(0.55), Y(Lref - sec["T"]) - 330, h * 0.4, TA.MIDDLE_LEFT, "NOTE")
    _txt(msp, f"(Sub-grade CBR must not be less than {p.cbr}%)", X(-0.2), ymin - 200,
         h * 1.0, TA.MIDDLE_LEFT, "TEXT")

    # layer call-outs (right-aligned shoulder table, left-aligned carriageway table)
    th = h * 0.85
    lh = th * 2.5
    ytab = ymin - 750

    def block(x_leader_m, width, items, y_target, right_aligned):
        xl_ = X(x_leader_m)
        xa, xb = (xl_ - width, xl_) if right_aligned else (xl_, xl_ + width)
        _line(msp, (xl_, y_target), (xl_, ytab + lh * 0.45), "NOTE", 7)
        for i, t in enumerate(items):
            yy = ytab - i * lh
            _line(msp, (xa, yy - lh * 0.45), (xb, yy - lh * 0.45), "NOTE", 7)
            if right_aligned:
                _txt(msp, t, xb - 60, yy, th, TA.MIDDLE_RIGHT, "TEXT")
            else:
                _txt(msp, t, xa + 60, yy, th, TA.MIDDLE_LEFT, "TEXT")
    sh_items = ["Single layer Surface-dressing", f"{p.base:.0f}mm Lateritic Base Course",
                f"{p.subbase:.0f}mm Laterite Sub-base Course"]
    cw_items = [f"{p.wearing:.0f}mm Wearing Course", "Single layer Surface-dressing",
                f"{p.base:.0f}mm Lateritic Base Course", f"{p.subbase:.0f}mm Laterite Sub-base Course"]
    xs_sh = -(c + 0.55 * p.shoulder)
    block(xs_sh, 3300, sh_items, Y(sec["surf"](-xs_sh) - sec["sd"] - 0.01), True)
    block(-2.2, 4300, cw_items, Y(sec["surf"](2.2) - sec["wear"] - 0.01), False)

    # title block
    ybt = ytab - 4 * lh - 700
    _txt(msp, title, ox, ybt, h * 1.9, TA.MIDDLE_CENTER, "TEXT")
    if subtitle:
        _txt(msp, subtitle, ox, ybt - h * 2.5, h * 1.9, TA.MIDDLE_CENTER, "TEXT")
    for i, line in enumerate(info):
        _txt(msp, line, ox, ybt - h * 2.5 * (2 if subtitle else 1) - (i + 0.2) * h * 1.9, h * 1.1,
             TA.MIDDLE_CENTER, "NOTE")
    return ybt - h * 2.5 * (2 if subtitle else 1) - (len(info) + 1) * h * 1.9 - h


# --------------------------------------------------------------------------
# Survey-only drawing (the CSV is the whole model: no road template)
# --------------------------------------------------------------------------
def station_points(ch, ground_df):
    """Survey shots for one chainage, sorted left to right, or None."""
    if ground_df is None or not len(ground_df):
        return None
    g = ground_df[np.isclose(ground_df["Chainage"].to_numpy(float), float(ch))]
    return g.sort_values("Offset") if len(g) else None


def draw_survey_section(msp, ch, g, p: Params, ox=0.0, oy=0.0, design=None):
    """Plot the surveyed shots for one station, optionally with the road design
    (pavement box) laid on the same geometry. Returns bottom y (mm)."""
    k = 1000.0
    ky = k * p.vexag
    h = p.text_h

    if g is None or len(g) < 1:
        _txt(msp, f"CH {fmt_chainage(ch)} — no survey shots", ox, oy, h * 1.0,
             TA.MIDDLE_CENTER, "TEXT")
        return oy - 2400

    g = g.sort_values("Offset")
    offs = g["Offset"].to_numpy(float)
    lvls = g["Level"].to_numpy(float)
    codes = ([str(c).strip() for c in g["Code"]] if "Code" in g.columns
             else [""] * len(g))
    lref = float(lvls.max())
    if design is not None:
        lref = max(lref, float(design))
    X = lambda o: ox + o * k
    Y = lambda l: oy - (lref - l) * ky
    x0_, x1_ = X(float(offs.min())), X(float(offs.max()))

    # ---- road design overlaid on the survey (pavement layers + formation) ----
    sec = None
    if design is not None and len(offs) > 1:
        sec = build_section(float(design), offs, lvls, p, False)
        # cut / fill between formation and surveyed ground, clipped to the survey width
        xa = max(float(offs.min()), -sec["xe"])
        xb = min(float(offs.max()), sec["xe"])
        if xb - xa > 0.2:
            gxs = np.linspace(xa, xb, 41)
            gnd = np.interp(gxs, offs, lvls)
            form = np.array([float(sec["surf"](x)) - sec["T"] for x in gxs])
            diff = form - gnd

            def _runs(mask):
                out, i = [], 0
                while i < len(mask):
                    if mask[i]:
                        j = i
                        while j + 1 < len(mask) and mask[j + 1]:
                            j += 1
                        if j > i:
                            out.append(range(i, j + 1))
                        i = j + 1
                    else:
                        i += 1
                return out

            for idxs in _runs(diff > 1e-3):          # fill
                poly = ([(X(gxs[k]), Y(form[k])) for k in idxs]
                        + [(X(gxs[k]), Y(gnd[k])) for k in reversed(idxs)])
                _hatch(msp, poly, ("ANSI31", 90, 45, 9), "EARTH")
            for idxs in _runs(diff < -1e-3):         # cut
                poly = ([(X(gxs[k]), Y(form[k])) for k in idxs]
                        + [(X(gxs[k]), Y(gnd[k])) for k in reversed(idxs)])
                _hatch(msp, poly, ("ANSI37", 140, 0, 30), "EARTH")

        kinds = {"wearing": "solid_dark", "dressing": "solid",
                 "base": ("AR-CONC", 4, 0, 7), "subbase": ("ANSI31", 30, 45, 8)}
        for name, polys in sec["layers"].items():
            for poly in polys:
                pts = [(X(x), Y(l)) for x, l in poly]
                _hatch(msp, pts, kinds[name])
                msp.add_lwpolyline(pts, close=True,
                                   dxfattribs={"layer": "PAVE", "lineweight": 25})
        # formation line under the pavement
        fl = ([(X(-x), Y(l)) for x, l in reversed(sec["sides"][-1]["verts"])]
              + [(X(0.0), Y(float(design) - sec["T"]))]
              + [(X(x), Y(l)) for x, l in sec["sides"][1]["verts"]])
        msp.add_lwpolyline(fl, dxfattribs={"layer": "SUBGRADE", "linetype": "DASHED",
                                           "ltscale": 40, "lineweight": 25})

    # ---- level grid (derived only from the shot levels) ----
    lo, hi = float(lvls.min()), float(lvls.max())
    room = (hi - lo) * ky
    if room >= 1.5 * h:                       # skip on near-flat sections (labels would collide)
        levels = [lo, hi] if room < 3 * h else [lo, (lo + hi) / 2, hi]
        xL = x0_ - 900
        for lv in levels:
            yy = Y(lv)
            msp.add_line((xL, yy), (x1_, yy),
                         dxfattribs={"layer": "NOTE", "color": 9, "linetype": "DOT",
                                     "ltscale": 8})
            _txt(msp, f"{lv:.3f}", xL - 150, yy, h * 0.42, TA.MIDDLE_RIGHT, "NOTE")

    # ---- ground profile exactly as surveyed (no padding, no template) ----
    pts = [(X(o), Y(l)) for o, l in zip(offs, lvls)]
    if len(pts) > 1:
        msp.add_lwpolyline(pts, dxfattribs={"layer": "GROUND", "lineweight": 50})
    for o, l in zip(offs, lvls):
        msp.add_circle((X(o), Y(l)), radius=80, dxfattribs={"layer": "GROUND"})

    # ---- per-segment cross-fall, computed from the shots ----
    for i in range(len(offs) - 1):
        dx = float(offs[i + 1] - offs[i])
        if abs(dx) < 1e-9:
            continue
        pct = (float(lvls[i + 1] - lvls[i]) / dx) * 100.0
        _txt(msp, f"{pct:+.2f}%", (X(offs[i]) + X(offs[i + 1])) / 2,
             (Y(lvls[i]) + Y(lvls[i + 1])) / 2 + 230, h * 0.48,
             TA.MIDDLE_CENTER, "NOTE")

    # ---- shot labels: code above, offset + level below ----
    for (o, l), c in zip(zip(offs, lvls), codes):
        label = c if c and c.lower() not in {"nan", "none"} else ""
        if label:
            _txt(msp, label, X(o), Y(l) + 430, h * 0.85, TA.MIDDLE_CENTER, "TEXT")
        _txt(msp, f"{o:+.2f}", X(o), Y(l) - 420, h * 0.60, TA.MIDDLE_CENTER, "NOTE")
        _txt(msp, f"{l:.3f}", X(o), Y(l) - 880, h * 0.60, TA.MIDDLE_CENTER, "NOTE")

    # ---- centre line ----
    jc = int(np.argmin(np.abs(offs)))
    if abs(offs[jc]) < 0.25:
        xc = X(float(offs[jc]))
        _line(msp, (xc, Y(lo) - 1900), (xc, oy + 1000), "CL", 1)
        _txt(msp, "CL", xc + 170, oy + 980, h * 0.72, TA.MIDDLE_LEFT, "CL", 1)

    # ---- dimension across the surveyed width ----
    if len(offs) > 1:
        ytop = oy + 1500
        yfrom = float(min(Y(lvls[0]), Y(lvls[-1]))) + 200
        _dim_h(msp, x0_, x1_, ytop, f"{float(offs.max() - offs.min()):.2f}m",
               h * 1.05, yfrom)

    # ---- title block ----
    ybt = Y(lo) - 2500
    _txt(msp, f"CROSS-SECTION AT CH {fmt_chainage(ch)}", ox, ybt, h * 1.6,
         TA.MIDDLE_CENTER, "TEXT")
    sub = "(SURVEYED GROUND + ROAD DESIGN)" if sec is not None else "(SURVEYED GROUND)"
    _txt(msp, sub, ox, ybt - h * 2.2, h * 1.15, TA.MIDDLE_CENTER, "TEXT")
    if len(offs) > 1:
        info = (f"{len(offs)} shots  ·  offsets {offs.min():+.2f} to {offs.max():+.2f} m  ·  "
                f"V.E. ×{p.vexag:g}")
    else:
        info = f"single shot only — no profile  ·  V.E. ×{p.vexag:g}"
    _txt(msp, info, ox, ybt - h * 4.0, h * 0.70, TA.MIDDLE_CENTER, "NOTE")
    if sec is not None:
        t_mm = p.wearing + p.surface_dressing + p.base + p.subbase
        _txt(msp, f"Road design: CL {float(design):.3f}  ·  {p.carriageway:.2f} m lanes  ·  "
                  f"{p.shoulder:.2f} m shoulders  ·  {p.camber:.1f}% camber  ·  {t_mm:.0f} mm pavement",
             ox, ybt - h * 5.3, h * 0.60, TA.MIDDLE_CENTER, "NOTE")

    # ---- shot schedule: everything is measured, nothing assumed ----
    ytab = ybt - h * 6.6
    colsx = [ox - 4500, ox - 1700, ox + 1800, ox + 5000]
    for cx, t in zip(colsx, ["POINT", "OFFSET m", "LEVEL m", "Δ CL m"]):
        _txt(msp, t, cx, ytab, h * 0.50, TA.MIDDLE_CENTER, "TEXT")
    _line(msp, (colsx[0] - 1100, ytab - h * 0.55), (colsx[-1] + 1100, ytab - h * 0.55),
          "NOTE", 7)
    for i, ((o, l), c) in enumerate(zip(zip(offs, lvls), codes)):
        yy = ytab - h * (1.3 + 1.2 * i)
        vals = [c or "·", f"{o:+.2f}", f"{l:.3f}", f"{l - lvls[jc]:+.3f}"]
        for cx, t in zip(colsx, vals):
            _txt(msp, t, cx, yy, h * 0.52, TA.MIDDLE_CENTER, "TEXT")
    yend = ytab - h * (1.3 + 1.2 * (len(offs) - 1)) - h * 0.8

    # ---- scale bar ----
    bx = ox - 4600
    ybar = yend - h * 1.1
    _line(msp, (bx, ybar), (bx + 1000, ybar), "NOTE", 7, lw=50)
    _line(msp, (bx, ybar - 100), (bx, ybar + 100), "NOTE", 7, lw=35)
    _line(msp, (bx + 1000, ybar - 100), (bx + 1000, ybar + 200), "NOTE", 7, lw=35)
    _txt(msp, "1 m", bx + 1000, ybar + 220, h * 0.50, TA.MIDDLE_LEFT, "NOTE")
    return ybar - h * 1.2


# --------------------------------------------------------------------------
# Public helpers
# --------------------------------------------------------------------------
def typical_sections(p: Params):
    """The two typical drawings (built-up + rural) like the reference sheet."""
    L = 100.0
    secs = []
    S = lambda x: L - (p.camber / 100) * min(abs(x), p.carriageway) - (p.shoulder_fall / 100) * max(abs(x) - p.carriageway, 0)
    T = (p.wearing + p.surface_dressing + p.base + p.subbase) / 1000
    xe = p.carriageway + p.shoulder
    gx = np.array([-p.row, p.row])
    # built-up: ground at shoulder-edge level (drain cut into it)
    secs.append((build_section(L, gx, np.full(2, S(xe) - 0.005), p, True),
                 "TYPICAL ROAD CROSS-SECTION", "(BUILT-UP AREAS)"))
    # rural: embankment
    gy = np.full(2, S(xe) - T - p.embankment_h)
    secs.append((build_section(L, gx, gy, p, False),
                 "TYPICAL ROAD CROSS-SECTION", "(RURAL AREAS)"))
    return secs


def station_section(row, ground_df, p: Params):
    gx, gy, surveyed = ground_for(row["Chainage"], ground_df, row["Existing"], p)
    built = str(row.get("Zone", "Built-up")).lower().startswith("b")
    return build_section(row["Design"], gx, gy, p, built), surveyed, built


def station_info(sec, row, surveyed):
    return [f"Design CL {row['Design']:.3f}   Existing CL {row['Existing']:.3f}   "
            f"Diff {row['Existing'] - row['Design']:+.3f} m",
            f"Cut {sec['cut']:.2f} m2 (platform {sec['cut_p']:.2f})   Fill {sec['fill']:.2f} m2 "
            f"(platform {sec['fill_p']:.2f})" + ("" if surveyed else "   [ground ASSUMED]")]


def typical_doc(p: Params):
    doc = new_doc()
    msp = doc.modelspace()
    oy = 0.0
    for sec, t, st in typical_sections(p):
        yb = draw_section(msp, sec, p, 0, oy, t, st)
        oy = yb - 3700
    return doc


def survey_doc(df: pd.DataFrame, ground_df, p: Params, cols=1, show_design=False):
    """One section per station, drawn from the survey shots (optionally + design)."""
    doc = new_doc()
    msp = doc.modelspace()
    secs = [(r, station_points(r["Chainage"], ground_df)) for _, r in df.iterrows()]
    span = 10.0
    for _, g in secs:
        if g is not None and len(g) > 1:
            span = max(span, float(g["Offset"].max() - g["Offset"].min()))
    dx = (span + 8.0) * 1000
    oy, rowbot = 0.0, None
    for i, (r, g) in enumerate(secs):
        if i and i % cols == 0:
            oy, rowbot = (rowbot if rowbot is not None else oy) - 2800, None
        design = None
        if show_design and pd.notna(r.get("Design")):
            design = float(r["Design"])
        yb = draw_survey_section(msp, r["Chainage"], g, p, (i % cols) * dx, oy, design=design)
        rowbot = yb if rowbot is None else min(rowbot, yb)
    return doc


def station_doc(df: pd.DataFrame, ground_df, p: Params, cols=3):
    doc = new_doc()
    msp = doc.modelspace()
    dx = (2 * p.row + 6) * 1000
    oy, rowbot = 0.0, 0.0
    for i, (_, r) in enumerate(df.iterrows()):
        if i and i % cols == 0:
            oy, rowbot = rowbot - 3700, 0.0
            oy = oy if oy < 0 else oy
        sec, surveyed, built = station_section(r, ground_df, p)
        yb = draw_section(msp, sec, p, (i % cols) * dx, oy,
                          f"CROSS-SECTION AT CH {fmt_chainage(r['Chainage'])}",
                          "(BUILT-UP AREA)" if built else "(RURAL AREA)", station_info(sec, r, surveyed))
        rowbot = yb if (i % cols == 0) else min(rowbot, yb)
    return doc


def doc_bytes(doc) -> bytes:
    s = io.StringIO()
    doc.write(s)
    return s.getvalue().encode("utf-8")


def render_png(doc, dpi=200, width_in=14, max_width_px=None) -> bytes:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from ezdxf import bbox
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

    if max_width_px:
        dpi = min(dpi, max_width_px / width_in)   # Streamlit resamples above 1460 px

    msp = doc.modelspace()
    ext = bbox.extents(msp)
    w, hh = ext.size.x, ext.size.y
    pad = 0.03 * max(w, hh)
    fig = plt.figure(figsize=(width_in, max(3, width_in * (hh + 2 * pad) / (w + 2 * pad))), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.COLOR,
                        lineweight_scaling=1.6, min_lineweight=0.3)
    Frontend(RenderContext(doc), MatplotlibBackend(ax), config=cfg).draw_layout(msp)
    ax.set_xlim(ext.extmin.x - pad, ext.extmax.x + pad)
    ax.set_ylim(ext.extmin.y - pad, ext.extmax.y + pad)
    ax.set_aspect("equal")
    ax.axis("off")
    # the ezdxf backend resizes the figure; restore a predictable output size
    fig.set_size_inches(width_in, max(1.0, width_in * (hh + 2 * pad) / (w + 2 * pad)))
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white", dpi=dpi)
    plt.close(fig)
    return buf.getvalue()


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Excel or CSV -> cross-section DXF")
    ap.add_argument("inputs", nargs="+",
                    help="a workbook .xlsx, or stations.csv [ground.csv] [parameters.csv]")
    ap.add_argument("--out", default="sections.dxf")
    a = ap.parse_args()
    files = {os.path.basename(str(p)): p for p in a.inputs}
    stations_df, ground_df, pdict = load_uploaded(files)
    station_doc(stations_df, ground_df, params_from_dict(pdict)).saveas(a.out)
    print("saved", a.out)
