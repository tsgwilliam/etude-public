# TODO export legend with png. Currently it is just in Streamlit and not in plotly

# =============================================================================
# SECTION 1 — PAGE SETUP AND IMPORTS
# Imports external dependencies and defines the page shell plus top-level upload
# widgets shown before project data is processed.
# =============================================================================
import io
import re
import tempfile
import colorsys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from oneclick_detail import (
    load_detail_report,
    material_breakdown,
    plot_material_treemap_emissions,
    plot_material_treemap_mass,
)

def render_page_shell_and_uploads():
    st.set_page_config(page_title="OneClickLCA Detail Report Grapher", layout="wide")
    st.title("OneClickLCA → standard graphs (Detail Report)")

    st.write(
        "Upload one or more OneClick detailReport exports (.xls/.xlsx/.csv). "
        "Assign each file a building name and GIA, then choose which buildings to combine."
    )

    uploaded_files = st.file_uploader(
        "Upload OneClick detailReport exports (.xls/.xlsx/.csv)",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True,
    )

    st.subheader("Manual additions (kgCO2e/m² GIA)")
    manual_file = st.file_uploader(
        "Upload manual entries CSV",
        type=["csv"],
        key="manual_upload",
    )
    st.caption("Manual CSV must include a 'Building Name' column matching the building names in the table below.")

    return uploaded_files, manual_file

# =============================================================================
# SECTION 2 — GLOBAL CONSTANTS
# Stores fixed ordering rules, labels, and export constants used across charts,
# legends, and GLA outputs.
# =============================================================================
RICS_HIGH_ORDER = [
    "Deconstruction",
    "Substructure",
    "Superstructure",
    "Finishes",
    "Fittings, furnishings and equipment",
    "Services",
    "Prefabricated buildings and building units",
    "Work to existing buildings",
    "External works",
    "Main contractor preliminaries",
]

GLA_TABLE1_CATEGORY_ORDER = [
    "0.1 Toxic Mat.",
    "0.2 Demolition",
    "0.3 Supports",
    "0.4 Groundworks",
    "0.5 Diversion",
    "1 Substructure",
    "2.1 Frame",
    "2.2 Upper Floors",
    "2.3 Roof",
    "2.4 Stairs & Ramps",
    "2.5 Ext. Walls",
    "2.6 Windows & Ext. Doors",
    "2.7 Int. Walls & Partitions",
    "2.8 Int. Doors",
    "3 Finishes",
    "4 Fittings, furnishings & equipments",
    "5 Services (MEP)",
    "6 Prefabricated",
    "7 Existing bldg",
    "8 Ext. works",
    "Other or overall site construction",
    "Unclassified / Other",
]

GLA_STAGE_ORDER = [
    "A1-A3", "A4", "A5",
    "B1", "B2", "B3", "B4", "B5", "B6", "B7",
    "C1", "C2", "C3", "C4", "D",
]

CANONICAL_RICS_LEGEND_SYSTEM = "streamlit_high_level"

# =============================================================================
# SECTION 3 — SMALL GENERIC UTILITIES
# Holds simple reusable helpers that do not depend on Streamlit page flow.
# Covers labels, validation, temporary files, colours, and text cleanup.
# =============================================================================
def format_selected_buildings_title(selected_buildings: list[str], total_gia: float) -> str:
    if not selected_buildings:
        buildings_txt = "No buildings selected"
    elif len(selected_buildings) == 1:
        buildings_txt = selected_buildings[0]
    elif len(selected_buildings) <= 3:
        buildings_txt = ", ".join(selected_buildings)
    else:
        buildings_txt = f"{len(selected_buildings)} buildings"

    return f"{buildings_txt} — GIA {total_gia:,.0f} m²"

def validate_required_columns(df: pd.DataFrame, required: list[str], context: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"{context} missing required columns: {missing}")

def default_building_name(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem

def format_building_label(selected_buildings: list[str]) -> str:
    names = [str(x).strip() for x in selected_buildings if str(x).strip()]
    if not names:
        return "No building selected"
    if len(names) == 1:
        return names[0]
    if len(names) <= 4:
        return " + ".join(names)
    return f"{len(names)} buildings combined"

def clean_rics_label(s: str) -> str:
    if s is None:
        return ""

    s = str(s).strip()

    # remove leading numbering (e.g. "2", "2.1", "2.1.3")
    s = re.sub(r"^\d+(?:\.\d+)*\s*", "", s)

    # remove any leftover leading punctuation (., -, :, etc)
    s = re.sub(r"^[\.\-–—:]+\s*", "", s)

    return s.strip()

def parse_rics_numeric_tuple(s: str) -> tuple:
    if s is None:
        return (9999,)

    s = str(s).strip()
    match = re.match(r"^(\d+(?:\.\d+)*)", s)
    if not match:
        return (9999,)

    return tuple(int(part) for part in match.group(1).split(".")) or (9999,)

def _hex_to_rgb01(hex_color: str) -> tuple[float, float, float]:
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Expected 6-digit hex colour, got: {hex_color}")
    return int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0

def _rgb01_to_hex(rgb: tuple[float, float, float]) -> str:
    r, g, b = rgb
    return "#{:02x}{:02x}{:02x}".format(
        int(max(0, min(1, r)) * 255),
        int(max(0, min(1, g)) * 255),
        int(max(0, min(1, b)) * 255),
    )

def _get_partial_target_line_x_range(bar_scope: str) -> tuple[float, float]:
    scope = str(bar_scope).strip().lower()
    if scope == "upfront":
        return 0.0, 0.5
    if scope == "whole life cycle":
        return 0.5, 1.0
    return 0.0, 1.0

def shade_palette(base_hex: str, n: int) -> list[str]:
    if n <= 1:
        return [base_hex]
    r, g, b = _hex_to_rgb01(base_hex)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l_min = max(0.25, l - 0.22)
    l_max = min(0.75, l + 0.22)
    out = []
    for i in range(n):
        li = l_min + (l_max - l_min) * (i / (n - 1))
        rr, gg, bb = colorsys.hls_to_rgb(h, li, s)
        out.append(_rgb01_to_hex((rr, gg, bb)))
    return out

def luminance_from_hex(hex_color: str) -> float:
    r0, g0, b0 = _hex_to_rgb01(hex_color)
    return 0.2126 * r0 + 0.7152 * g0 + 0.0722 * b0

def get_default_rics_colour(h: str) -> str:
    if not h:
        return "#4c78a8"

    h_low = str(h).lower()

    if "superstructure" in h_low:
        return "#23A846"
    if "substructure" in h_low:
        return "#E8C602"
    if "finish" in h_low:
        return "#EC8447"
    if "service" in h_low:
        return "#6BA1D5"
    if "fitting" in h_low:
        return "#A56CC1"
    if "prefabricated" in h_low:
        return "#00A6A6"
    if "existing" in h_low:
        return "#C06C84"
    if "external" in h_low:
        return "#F67280"
    if "prelim" in h_low:
        return "#6C6C6C"
    if "deconstruction" in h_low:
        return "#999999"

    return "#4c78a8"

def pick_chart_value_col(df: pd.DataFrame, em_col: str) -> str:
    if "rics_allocated_value" in df.columns:
        return "rics_allocated_value"
    return em_col

def make_png_bytes(fig: go.Figure, width_px: int, height_px: int) -> bytes | None:
    try:
        return fig.to_image(format="png", width=width_px, height=height_px, scale=1)
    except Exception:
        return None

# =============================================================================
# SECTION 4 — INPUT PARSING AND DATA ASSEMBLY
# Reads uploaded files and manual CSV data, normalises returned rows, and
# combines multiple selected buildings into project-level datasets.
# =============================================================================
def _read_uploaded_to_tempfile(uploaded_file) -> Path:
    name_l = uploaded_file.name.lower()
    if name_l.endswith(".csv"):
        suffix = ".csv"
    elif name_l.endswith(".xlsx"):
        suffix = ".xlsx"
    else:
        suffix = ".xls"

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.flush()
    tmp.close()
    return Path(tmp.name)

def drop_oneclick_total_rows(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if "section" in out.columns:
        sec = out["section"].astype(str).str.strip().str.lower()
        out = out[~sec.isin({"total", "subtotal", "sum"})]

    if "Resource" in out.columns:
        res = out["Resource"].astype(str).str.strip()
        res_l = res.str.lower()
        out = out[
            ~(
                res_l.eq("deconstruction and demolition process (per gia)")
                | res_l.str.contains(r"\bprocess\s*\(per\s*gia\)\b", regex=True, na=False)
                | res_l.str.match(r"^total\b", na=False)
            )
        ]

    descriptor_cols = [
        c
        for c in [
            "Resource",
            "Design name",
            "Indicator name",
            "Comment",
            "element_name",
            "User input",
            "User input unit",
        ]
        if c in out.columns
    ]
    if descriptor_cols:
        tmp = out[descriptor_cols].copy()
        for c in descriptor_cols:
            tmp[c] = tmp[c].astype(str).str.strip()
            tmp.loc[tmp[c].isin(["", "nan", "None"]), c] = None
        all_missing = tmp.isna().all(axis=1)

        num_cols = out.select_dtypes("number").columns.tolist()
        if num_cols:
            has_any_number = out[num_cols].notna().any(axis=1) & (out[num_cols].abs().sum(axis=1) != 0)
            out = out[~(all_missing & has_any_number)]
        else:
            out = out[~all_missing]

    return out

def pick_emissions_col(df: pd.DataFrame) -> str:
    preferred = ["kgco2e", "gwp_kgco2e", "gwp", "GWP", "GWP total", "GWP [kgCO2e]"]
    for c in preferred:
        if c in df.columns:
            return c

    candidates = []
    for c in df.columns:
        cl = str(c).lower()
        if any(k in cl for k in ["gwp", "co2e", "kgco2", "kg co2", "kg co2e", "kgco2e", "carbon"]):
            candidates.append(c)
    if candidates:
        return candidates[0]

    numeric = df.select_dtypes("number").columns.tolist()
    if numeric:
        return numeric[0]

    raise KeyError("No numeric emissions column found.")

def parse_oneclick_uploaded_file(uploaded_file):
    tmp_path = _read_uploaded_to_tempfile(uploaded_file)

    result = load_detail_report(tmp_path)
    if isinstance(result, tuple) and len(result) == 2:
        rows, meta = result
        entries = None
    else:
        rows = getattr(result, "rows", None)
        meta = getattr(result, "meta", None)
        entries = getattr(result, "entries", None)
        if rows is None:
            raise TypeError(f"Unexpected return type from load_detail_report: {type(result)}")

    rows = drop_oneclick_total_rows(rows)
    em_col = pick_emissions_col(rows)

    if "_source_row_id" in rows.columns:
        rows["_source_row_id"] = rows["_source_row_id"].astype(str)

    return rows, meta, entries, em_col

def build_combined_rows(uploaded_files_, building_meta_df: pd.DataFrame, selected_buildings: list[str]):
    rows_all = []
    entries_all = []
    meta_by_building = {}
    em_col_global = None

    file_lookup = {f.name: f for f in uploaded_files_}
    selected_meta = building_meta_df[building_meta_df["building_name"].isin(selected_buildings)].copy()

    for _, r in selected_meta.iterrows():
        uploaded_file = file_lookup[r["file_name"]]
        rows, meta, entries, em_col = parse_oneclick_uploaded_file(uploaded_file)

        if em_col_global is None:
            em_col_global = em_col

        rows = rows.copy()
        rows["building_name"] = r["building_name"]
        rows["building_gia_m2"] = float(r["gia_m2"] or 0.0)
        rows["source_file"] = str(r["file_name"])

        if "_source_row_id" in rows.columns:
            rows["_source_row_id"] = rows["_source_row_id"].astype(str).map(
                lambda x: f"{r['building_name']}__{x}"
            )

        rows_all.append(rows)
        meta_by_building[r["building_name"]] = meta

        if entries is not None and len(entries):
            entries = entries.copy()
            entries["building_name"] = r["building_name"]
            entries["source_file"] = str(r["file_name"])
            entries_all.append(entries)

    combined_rows = pd.concat(rows_all, ignore_index=True) if rows_all else pd.DataFrame()
    combined_entries = pd.concat(entries_all, ignore_index=True) if entries_all else None

    return combined_rows, combined_entries, meta_by_building, em_col_global, selected_meta

def load_manual_intensity_rows(
    manual_file,
    building_meta_df: pd.DataFrame,
    selected_buildings: list[str],
    em_col: str,
) -> pd.DataFrame:
    df = pd.read_csv(manual_file)

    required = [
        "Building Name",
        "manual_item_id",
        "stack_label",
        "rics_high_label",
        "rics_level2_label",
        "section",
        "kgco2e_per_m2_gia",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Manual CSV missing columns: {missing}")

    df["Building Name"] = df["Building Name"].astype(str).str.strip()
    df = df[df["Building Name"].isin(selected_buildings)].copy()

    if df.empty:
        return pd.DataFrame()

    meta = building_meta_df[["building_name", "gia_m2"]].copy()
    meta["building_name"] = meta["building_name"].astype(str).str.strip()

    df = df.merge(
        meta,
        left_on="Building Name",
        right_on="building_name",
        how="left",
    )

    if df["gia_m2"].isna().any():
        missing_names = sorted(df.loc[df["gia_m2"].isna(), "Building Name"].unique().tolist())
        raise ValueError(
            "Manual CSV contains Building Name values that do not match the building table or have no GIA: "
            + ", ".join(missing_names)
        )

    df["section"] = df["section"].astype(str).str.strip()

    df["kgco2e"] = (
        pd.to_numeric(df["kgco2e_per_m2_gia"], errors="coerce").fillna(0.0)
        * pd.to_numeric(df["gia_m2"], errors="coerce").fillna(0.0)
    )

    if "biogenic_kgco2e_per_m2_gia" in df.columns:
        df["biogenic_kgco2e"] = (
            pd.to_numeric(df["biogenic_kgco2e_per_m2_gia"], errors="coerce").fillna(0.0)
            * pd.to_numeric(df["gia_m2"], errors="coerce").fillna(0.0)
        )
    else:
        df["biogenic_kgco2e"] = 0.0

    out = pd.DataFrame()
    out["section"] = df["section"]
    out["rics_high_label"] = df["rics_high_label"].astype(str).str.strip()
    out["rics_level2_label"] = df["rics_level2_label"].astype(str).str.strip()
    out["rics_alloc_label"] = out["rics_level2_label"]
    out["rics_detail"] = df["stack_label"]
    out["element_name"] = df["stack_label"]
    out["Comment"] = df.get("comment", "Manual entry")
    out[em_col] = df["kgco2e"]
    out["rics_allocated_value"] = out[em_col]
    out["_source_row_id"] = (
        "manual__"
        + df["Building Name"].astype(str)
        + "__"
        + df["manual_item_id"].astype(str)
    )
    out["biogenic_kgco2e"] = df["biogenic_kgco2e"]
    out["building_name"] = df["Building Name"]
    out["building_gia_m2"] = pd.to_numeric(df["gia_m2"], errors="coerce").fillna(0.0)
    out["source_file"] = "manual_upload"

    return out

# =============================================================================
# SECTION 5 — ROW ENRICHMENT AND CLASSIFICATION
# Derives labels, modules, chart values, contributors, and biogenic totals for
# charts, legends, audits, and exports.
# =============================================================================
def expand_modules(modules: list[str]) -> list[str]:
    expanded = []
    for m in modules:
        if m == "A1-A5":
            expanded.extend(["A1-A5", "A1-A3", "A4", "A5"])
        else:
            expanded.append(m)
    return list(dict.fromkeys(expanded))

def derive_level2_label(rows: pd.DataFrame) -> pd.DataFrame:
    out = rows.copy()

    if "rics_level2_label" in out.columns:
        return out

    if "rics_detail" in out.columns:
        detail = out["rics_detail"].astype(str)
        out["rics_level2_label"] = (
            detail.str.extract(r"^(\d+\.\d+\s+[^:;|]+)", expand=False).fillna(detail)
        )
        return out

    if "rics_high_label" in out.columns:
        out["rics_level2_label"] = out["rics_high_label"].astype(str)
    else:
        out["rics_level2_label"] = pd.Series("", index=out.index, dtype="string")

    return out

def find_biogenic_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        cl = str(c).lower()
        if cl in ["biogenic_kgco2e", "biogenic"]:
            return c
    return None

def apply_stack_label_mapping(
    df: pd.DataFrame,
    display_map: dict[str, str],
    group_map: dict[str, str],
    detail_col: str = "rics_detail",
    out_prefix: str = "lvl2",
) -> pd.DataFrame:
    out = df.copy()
    id_col = f"{out_prefix}_id"
    display_col = f"{out_prefix}_display"
    group_col = f"{out_prefix}_group"
    final_col = f"{out_prefix}_final"

    out[id_col] = out[detail_col].map(clean_rics_label).astype(str)
    out[display_col] = out[id_col].map(display_map).fillna(out[id_col])
    out[group_col] = out[id_col].map(group_map)
    out[final_col] = out[group_col].where(
        out[group_col].notna() & (out[group_col].astype(str).str.len() > 0),
        out[display_col],
    )
    return out

def build_rics_high_order(agg: pd.DataFrame, clean_col: str = "rics_high_clean") -> list[str]:
    highs_present = agg[clean_col].dropna().unique().tolist()
    return [h for h in RICS_HIGH_ORDER if h in highs_present] + [
        h for h in highs_present if h not in RICS_HIGH_ORDER
    ]

def compute_top_biogenic_contributors(
    rows: pd.DataFrame,
    em_col: str,
    top_n: int = 5,
) -> str:
    rr = rows.copy()

    if "section" not in rr.columns:
        return ""

    rr = rr[rr["section"].astype(str).str.strip().str.lower().eq("bioc")].copy()
    if rr.empty:
        return ""

    if "element_name" in rr.columns:
        rr["true_name"] = rr["element_name"].fillna("")
    elif "Comment" in rr.columns:
        rr["true_name"] = rr["Comment"].fillna("")
    else:
        rr["true_name"] = rr.get("Resource", "(unknown)").fillna("(unknown)")

    vals = (
        rr.groupby("true_name", dropna=False)[em_col]
        .sum()
        .abs()
        .sort_values(ascending=False)
    )

    total = vals.sum()
    if total <= 0:
        return ""

    lines = []
    for name, val in vals.head(top_n).items():
        share = 100.0 * val / total
        lines.append(f"{name}: {share:.1f}%")

    return "<br>".join(lines)

def compute_biogenic_total(rows: pd.DataFrame, modules: list[str], em_col: str, scale: float = 1.0) -> float:
    rr = rows.copy()

    if "section" in rr.columns:
        sec = rr["section"].astype(str).str.strip().str.lower()
        bio_rows = rr[sec.eq("bioc")].copy()

        if not bio_rows.empty:
            if "_source_row_id" in bio_rows.columns:
                total = (
                    bio_rows.groupby("_source_row_id", dropna=False)[em_col]
                    .first()
                    .pipe(pd.to_numeric, errors="coerce")
                    .fillna(0.0)
                    .sum()
                )
            else:
                total = pd.to_numeric(bio_rows[em_col], errors="coerce").fillna(0.0).sum()

            return float(abs(total) * scale)

        modules_expanded = expand_modules(modules)
        rr = rr[rr["section"].isin(modules_expanded)]

    bio_col = find_biogenic_col(rr)
    if bio_col is None:
        return 0.0

    if "_source_row_id" in rr.columns:
        total = (
            rr.groupby("_source_row_id", dropna=False)[bio_col]
            .first()
            .pipe(pd.to_numeric, errors="coerce")
            .fillna(0.0)
            .sum()
        )
    else:
        total = pd.to_numeric(rr[bio_col], errors="coerce").fillna(0.0).sum()

    return float(abs(total) * scale)

def compute_top_contributors_by_stack(
    rows: pd.DataFrame,
    modules: list[str],
    value_col: str,
    display_map: dict[str, str],
    group_map: dict[str, str],
    top_n: int = 5,
) -> dict[str, str]:
    modules_expanded = expand_modules(modules)
    rr = rows.copy()
    rr = rr[rr["section"].isin(modules_expanded)].copy()
    if rr.empty:
        return {}

    if "element_name" in rr.columns:
        rr["true_name"] = rr["element_name"].fillna("")
    elif "Comment" in rr.columns:
        rr["true_name"] = rr["Comment"].fillna("")
    else:
        rr["true_name"] = rr.get("Resource", "(unknown)").fillna("(unknown)")

    rr = apply_stack_label_mapping(rr, display_map=display_map, group_map=group_map)

    totals = rr.groupby("lvl2_final", dropna=False)[value_col].sum()

    by_elem = (
        rr.groupby(["lvl2_final", "true_name"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .sort_values(["lvl2_final", value_col], ascending=[True, False])
    )

    out = {}
    for stack, total in totals.items():
        if not total:
            continue
        top = by_elem[by_elem["lvl2_final"] == stack].head(top_n)
        lines = []
        for _, rrr in top.iterrows():
            share = (rrr[value_col] / total) * 100.0
            lines.append(f"{rrr['true_name']}: {share:.1f}%")
        out[str(stack)] = "<br>".join(lines)

    return out

def build_stack_by_module_export(
    rows: pd.DataFrame,
    em_col: str,
    display_map: dict[str, str],
    group_map: dict[str, str],
) -> pd.DataFrame:
    rr = derive_level2_label(rows).copy()
    value_col = pick_chart_value_col(rr, em_col)

    rr = apply_stack_label_mapping(
        rr,
        display_map=display_map,
        group_map=group_map,
        out_prefix="stack",
    )

    out = (
        rr.groupby(["stack_final", "section"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .pivot_table(
            index="stack_final",
            columns="section",
            values=value_col,
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    out.columns = [str(c) for c in out.columns]
    return out

# =============================================================================
# SECTION 6 — RECONCILIATION AND COVERAGE CHECKS
# Compares source totals with allocated totals and prepares audit outputs used by
# the UI.
# =============================================================================
def build_reconciliation_table(rows: pd.DataFrame, em_col: str) -> pd.DataFrame:
    rr = rows.copy()

    if "_source_row_id" not in rr.columns:
        raise KeyError("Expected '_source_row_id' in rows for reconciliation.")
    if "section" not in rr.columns:
        rr["section"] = ""

    alloc_col = "rics_allocated_value" if "rics_allocated_value" in rr.columns else em_col
    key_cols = ["_source_row_id", "section"]

    original = (
        rr.groupby(key_cols, dropna=False)[em_col]
        .first()
        .reset_index()
        .rename(columns={em_col: "original_kgco2e"})
    )

    allocated = (
        rr.groupby(key_cols, dropna=False)[alloc_col]
        .sum()
        .reset_index()
        .rename(columns={alloc_col: "allocated_kgco2e"})
    )

    desc_cols = [c for c in ["element_name", "Resource", "Comment", "rics_detail", "building_name"] if c in rr.columns]
    desc = (
        rr.groupby(key_cols, dropna=False)[desc_cols].first().reset_index()
        if desc_cols else pd.DataFrame(columns=key_cols)
    )

    out = original.merge(allocated, on=key_cols, how="outer").merge(desc, on=key_cols, how="left")
    out["original_kgco2e"] = pd.to_numeric(out["original_kgco2e"], errors="coerce").fillna(0.0)
    out["allocated_kgco2e"] = pd.to_numeric(out["allocated_kgco2e"], errors="coerce").fillna(0.0)
    out["gap_kgco2e"] = out["original_kgco2e"] - out["allocated_kgco2e"]
    out["gap_pct"] = out["gap_kgco2e"] / out["original_kgco2e"].replace(0, pd.NA)
    out["gap_pct"] = out["gap_pct"].fillna(0.0)
    return out

def summarise_reconciliation(recon: pd.DataFrame, tolerance_kg: float = 0.01) -> pd.DataFrame:
    r = recon.copy()
    r["is_flagged"] = r["gap_kgco2e"].abs() > tolerance_kg

    summary = (
        r.groupby("section", dropna=False)
        .agg(
            original_kgco2e=("original_kgco2e", "sum"),
            allocated_kgco2e=("allocated_kgco2e", "sum"),
            gap_kgco2e=("gap_kgco2e", "sum"),
            flagged_rows=("is_flagged", "sum"),
            total_rows=("_source_row_id", "count"),
        )
        .reset_index()
    )
    summary["gap_pct"] = summary["gap_kgco2e"] / summary["original_kgco2e"].replace(0, pd.NA)
    summary["gap_pct"] = summary["gap_pct"].fillna(0.0)
    return summary

def build_unassigned_rows(rows: pd.DataFrame, em_col: str, tolerance_kg: float = 0.01) -> pd.DataFrame:
    recon = build_reconciliation_table(rows, em_col)
    flagged = recon[recon["gap_kgco2e"].abs() > tolerance_kg].copy()
    flagged = flagged.sort_values("gap_kgco2e", key=lambda s: s.abs(), ascending=False)
    return flagged

def compute_chart_coverage(
    rows: pd.DataFrame,
    modules: list[str],
    em_col: str,
    display_map: dict[str, str],
    group_map: dict[str, str],
    contingency_map: dict[str, float],
    project_contingency_pct: float,
    use_intensity: bool,
    gia_m2: float,
) -> tuple[dict, pd.DataFrame]:
    rr = derive_level2_label(rows).copy()
    modules_expanded = expand_modules(modules)
    rr = rr[rr["section"].isin(modules_expanded)]

    value_col = pick_chart_value_col(rr, em_col)

    scale = 1.0
    if use_intensity and gia_m2 and gia_m2 > 0:
        scale = 1.0 / float(gia_m2)

    rr = apply_stack_label_mapping(rr, display_map=display_map, group_map=group_map)

    seg_df = (
        rr.groupby(["lvl2_id", "lvl2_final"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .rename(columns={value_col: "base_value"})
    )

    cont_df = (
        rr.groupby("lvl2_id", dropna=False)[value_col]
        .sum()
        .reset_index()
        .rename(columns={value_col: "base_value"})
    )

    project_frac = float(project_contingency_pct or 0.0) / 100.0
    cont_df["contingency_frac"] = cont_df["lvl2_id"].map(contingency_map).fillna(0.0) + project_frac
    cont_df["contingency_value"] = cont_df["base_value"] * cont_df["contingency_frac"]

    total_source = 0.0
    if "_source_row_id" in rr.columns:
        total_source = float(
            rr.groupby(["_source_row_id", "section"], dropna=False)[em_col].first().sum()
        )
    else:
        total_source = float(rr[em_col].sum())

    plotted_base = float(seg_df["base_value"].sum())
    plotted_cont = float(cont_df["contingency_value"].sum())

    alloc_gap_df = pd.DataFrame()
    if "_source_row_id" in rr.columns and "rics_allocated_value" in rr.columns:
        src = (
            rr.groupby(["_source_row_id", "section"], dropna=False)[[em_col, "rics_allocated_value"]]
            .first()
            .reset_index()
        )
        alloc = (
            rr.groupby(["_source_row_id", "section"], dropna=False)["rics_allocated_value"]
            .sum()
            .reset_index()
        )
        src = src.drop(columns=["rics_allocated_value"]).merge(alloc, on=["_source_row_id", "section"], how="left")
        src = src.rename(columns={em_col: "source_value", "rics_allocated_value": "allocated_value"})
        src["gap"] = src["source_value"] - src["allocated_value"]
        src["gap_abs"] = src["gap"].abs()

        detail_cols = [c for c in ["element_name", "rics_detail", "Resource", "Comment", "building_name"] if c in rr.columns]
        rr_detail = rr.copy()
        rr_detail["_source_row_id_sort"] = rr_detail["_source_row_id"].astype(str)

        detail = (
            rr_detail.sort_values("_source_row_id_sort")
            .groupby(["_source_row_id", "section"], dropna=False)[detail_cols]
            .first()
            .reset_index()
        )
        alloc_gap_df = src.merge(detail, on=["_source_row_id", "section"], how="left")
        alloc_gap_df = alloc_gap_df.sort_values("gap_abs", ascending=False)

    uncovered_positive = 0.0
    if not alloc_gap_df.empty:
        uncovered_positive = float(alloc_gap_df.loc[alloc_gap_df["gap"] > 1e-9, "gap"].sum())

    coverage = {
        "value_col_used": value_col,
        "total_source": total_source,
        "total_source_scaled": total_source * scale,
        "plotted_base": plotted_base * scale,
        "plotted_contingency": plotted_cont * scale,
        "uncovered_positive": uncovered_positive * scale,
        "coverage_ratio": (plotted_base / total_source) if total_source else 1.0,
        "fully_allocated": uncovered_positive < 1e-6,
    }
    return coverage, alloc_gap_df

# =============================================================================
# SECTION 7 — CHART ITEM EDITOR STATE
# Builds and persists the editable subcategory map used for grouping, visibility,
# display-name overrides, and contingency by subcategory.
# =============================================================================
def sync_building_meta(uploaded_files_) -> pd.DataFrame:
    cols = ["file_name", "building_name", "gia_m2"]

    fresh = pd.DataFrame(
        [
            {
                "file_name": f.name,
                "building_name": default_building_name(f.name),
                "gia_m2": 0.0,
            }
            for f in uploaded_files_
        ],
        columns=cols,
    )

    if "building_meta_df" not in st.session_state:
        st.session_state["building_meta_df"] = fresh.copy()
        return st.session_state["building_meta_df"]

    existing = st.session_state["building_meta_df"][cols].copy()
    out = fresh[["file_name"]].merge(existing, on="file_name", how="left", suffixes=("", "_old"))

    default_name = dict(zip(fresh["file_name"], fresh["building_name"]))
    default_gia = dict(zip(fresh["file_name"], fresh["gia_m2"]))

    out["building_name"] = out["building_name"].fillna(out["file_name"].map(default_name))
    out["gia_m2"] = pd.to_numeric(out["gia_m2"], errors="coerce").fillna(out["file_name"].map(default_gia))

    out = out[cols].copy()
    st.session_state["building_meta_df"] = out
    return out

def build_subcat_editor_df(rows: pd.DataFrame) -> pd.DataFrame:
    rr = rows.copy()

    if "rics_detail" not in rr.columns:
        raise KeyError("Expected 'rics_detail' column in parsed rows.")

    rr["subcategory_id"] = rr["rics_detail"].map(clean_rics_label).astype(str)
    rr["_order_key"] = rr["rics_detail"].astype(str).str.extract(r"^(\d+(?:\.\d+)*)", expand=False).fillna("999")

    if "rics_level2_label" in rr.columns:
        rr["rics_category"] = rr["rics_level2_label"].astype(str).str.strip()
    elif "rics_high_label" in rr.columns:
        rr["rics_category"] = rr["rics_high_label"].astype(str).str.strip()
    else:
        rr["rics_category"] = ""

    subcats = (
        rr[["subcategory_id", "_order_key", "rics_category"]]
        .dropna()
        .drop_duplicates(subset=["subcategory_id"])
        .sort_values(["_order_key", "subcategory_id"])
    )

    ids = subcats["subcategory_id"].tolist()
    rics_labels = subcats["rics_category"].tolist()

    return pd.DataFrame(
        {
            "subcategory_id": ids,
            "rics_category": rics_labels,
            "display_name": ids,
            "group_name": ["" for _ in ids],
            "contingency_pct": [0.0 for _ in ids],
            "show_on_chart": [True for _ in ids],
        }
    )

def sync_subcat_editor_state(rows: pd.DataFrame) -> pd.DataFrame:
    fresh_df = build_subcat_editor_df(rows)
    cols = ["subcategory_id", "rics_category", "display_name", "group_name", "contingency_pct", "show_on_chart"]

    if "subcat_map_df" not in st.session_state:
        st.session_state["subcat_map_df"] = fresh_df[cols].copy()
        return st.session_state["subcat_map_df"]

    existing = st.session_state["subcat_map_df"].copy()
    for col in cols:
        if col not in existing.columns:
            if col == "rics_category":
                existing[col] = ""
            elif col == "show_on_chart":
                existing[col] = True
            elif col == "contingency_pct":
                existing[col] = 0.0
            else:
                existing[col] = ""

    existing = existing[cols].copy()

    out = fresh_df[["subcategory_id", "rics_category"]].merge(
        existing.drop(columns=["rics_category"]),
        on="subcategory_id",
        how="left",
    )

    default_display = dict(zip(fresh_df["subcategory_id"], fresh_df["display_name"]))
    default_group = dict(zip(fresh_df["subcategory_id"], fresh_df["group_name"]))
    default_cont = dict(zip(fresh_df["subcategory_id"], fresh_df["contingency_pct"]))

    out["display_name"] = out["display_name"].fillna(out["subcategory_id"].map(default_display))
    out["group_name"] = out["group_name"].fillna(out["subcategory_id"].map(default_group))
    out["contingency_pct"] = pd.to_numeric(out["contingency_pct"], errors="coerce").fillna(
        out["subcategory_id"].map(default_cont)
    )
    out["show_on_chart"] = out["show_on_chart"].fillna(True).astype(bool)

    out = out[cols].copy()
    st.session_state["subcat_map_df"] = out
    return out

# =============================================================================
# SECTION 8 — LEGEND DATA BUILDERS
# Builds high-level legend tables for one-stack and two-stack views and keeps
# legend summarisation separate from chart drawing.
# =============================================================================
def _render_legend_items_grid(
    legend_df: pd.DataFrame,
    n_cols: int,
    render_text,
) -> None:
    rows_of_items = [legend_df.iloc[i:i + n_cols] for i in range(0, len(legend_df), n_cols)]

    for chunk in rows_of_items:
        cols = st.columns(n_cols)
        for col, (_, row) in zip(cols, chunk.iterrows()):
            with col:
                swatch, textcol = st.columns([1, 12])
                with swatch:
                    st.markdown(
                        f"""
                        <div style="
                            width:16px;
                            height:16px;
                            background:{row['colour']};
                            border:1px solid #999;
                            border-radius:2px;
                            margin-top:4px;
                        "></div>
                        """,
                        unsafe_allow_html=True,
                    )
                with textcol:
                    render_text(row)

def build_high_level_legend_df(
    rows: pd.DataFrame,
    modules: list[str],
    em_col: str,
    colour_map: dict[str, str],
    use_intensity: bool,
    gia_m2: float,
    contingency_map: dict[str, float] | None = None,
    project_contingency_pct: float = 0.0,
    contingency_colour: str = "#6C6C6C",
    biogenic_colour: str = "#2B0FC9",
) -> pd.DataFrame:
    rr = derive_level2_label(rows).copy()
    modules_expanded = expand_modules(modules)
    rr = rr[rr["section"].isin(modules_expanded)].copy()

    scale = 1.0
    if use_intensity:
        if not gia_m2 or gia_m2 <= 0:
            raise ValueError("Intensity is enabled but GIA is not set (> 0).")
        scale = 1.0 / float(gia_m2)

    if rr.empty:
        base_out = pd.DataFrame(columns=["rics_high_clean", "value", "colour"])
    else:
        value_col = pick_chart_value_col(rr, em_col)
        rr["rics_high_clean"] = rr["rics_high_label"].map(clean_rics_label)

        base_out = (
            rr.groupby("rics_high_clean", dropna=False)[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "value"})
        )
        base_out["value"] = base_out["value"] * scale
        base_out["colour"] = base_out["rics_high_clean"].map(lambda x: colour_map.get(x, "#4c78a8"))

    legend_rows = [base_out]

    contingency_total = 0.0
    if not rr.empty:
        value_col = pick_chart_value_col(rr, em_col)
        rr["lvl2_id"] = rr["rics_detail"].map(clean_rics_label).astype(str)

        base_by_id = (
            rr.groupby("lvl2_id", dropna=False)[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "base_value"})
        )

        project_frac = float(project_contingency_pct or 0.0) / 100.0
        contingency_map = contingency_map or {}

        for _, row in base_by_id.iterrows():
            sid = str(row["lvl2_id"])
            base_val = float(row["base_value"] or 0.0)
            subcat_frac = float(contingency_map.get(sid, 0.0) or 0.0)
            contingency_total += base_val * (project_frac + subcat_frac)

        contingency_total *= scale

    if abs(contingency_total) > 1e-9:
        legend_rows.append(
            pd.DataFrame(
                [{
                    "rics_high_clean": "Contingency",
                    "value": contingency_total,
                    "colour": contingency_colour,
                }]
            )
        )

    bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
    if abs(bio_total) > 1e-9:
        legend_rows.append(
            pd.DataFrame(
                [{
                    "rics_high_clean": "Biogenic",
                    "value": -abs(bio_total),
                    "colour": biogenic_colour,
                }]
            )
        )

    out = pd.concat(legend_rows, ignore_index=True) if legend_rows else pd.DataFrame(
        columns=["rics_high_clean", "value", "colour"]
    )

    legend_order = RICS_HIGH_ORDER + ["Contingency", "Biogenic"]
    out["order"] = out["rics_high_clean"].apply(
        lambda x: legend_order.index(x) if x in legend_order else 999
    )
    out = out.sort_values(["order", "rics_high_clean"]).reset_index(drop=True)

    return out.drop(columns=["order"])

def build_high_level_legend_df_two_charts(
    rows: pd.DataFrame,
    upfront_modules: list[str],
    whole_life_modules: list[str],
    em_col: str,
    colour_map: dict[str, str],
    use_intensity: bool,
    gia_m2: float,
    contingency_map: dict[str, float] | None = None,
    project_contingency_pct: float = 0.0,
    contingency_colour: str = "#6C6C6C",
    biogenic_colour: str = "#E657C7",
) -> pd.DataFrame:
    upfront_df = build_high_level_legend_df(
        rows=rows,
        modules=upfront_modules,
        em_col=em_col,
        colour_map=colour_map,
        use_intensity=use_intensity,
        gia_m2=gia_m2,
        contingency_map=contingency_map,
        project_contingency_pct=project_contingency_pct,
        contingency_colour=contingency_colour,
        biogenic_colour=biogenic_colour,
    ).rename(columns={"value": "upfront_value"})

    whole_life_df = build_high_level_legend_df(
        rows=rows,
        modules=whole_life_modules,
        em_col=em_col,
        colour_map=colour_map,
        use_intensity=use_intensity,
        gia_m2=gia_m2,
        contingency_map=contingency_map,
        project_contingency_pct=project_contingency_pct,
        contingency_colour=contingency_colour,
        biogenic_colour=biogenic_colour,
    ).rename(columns={"value": "whole_life_value"})

    out = upfront_df.merge(
        whole_life_df[["rics_high_clean", "whole_life_value"]],
        on="rics_high_clean",
        how="outer",
    )

    colour_lookup = {
        **dict(zip(upfront_df["rics_high_clean"], upfront_df["colour"])),
        **dict(zip(whole_life_df["rics_high_clean"], whole_life_df["colour"])),
    }
    out["colour"] = out["rics_high_clean"].map(colour_lookup)

    out["upfront_value"] = pd.to_numeric(out["upfront_value"], errors="coerce").fillna(0.0)
    out["whole_life_value"] = pd.to_numeric(out["whole_life_value"], errors="coerce").fillna(0.0)

    legend_order = RICS_HIGH_ORDER + ["Contingency", "Biogenic"]
    out["order"] = out["rics_high_clean"].apply(
        lambda x: legend_order.index(x) if x in legend_order else 999
    )
    out = out.sort_values(["order", "rics_high_clean"]).reset_index(drop=True)

    return out.drop(columns=["order"])

def _render_high_level_legend_generic(
    legend_df: pd.DataFrame,
    use_intensity: bool,
    render_text,
) -> None:
    if legend_df.empty:
        return

    st.markdown("**High-level RICS category totals**")
    _render_legend_items_grid(legend_df=legend_df, n_cols=3, render_text=lambda row: render_text(row, use_intensity))

def render_high_level_legend(legend_df: pd.DataFrame, use_intensity: bool):
    def render_text(row, use_intensity_flag: bool):
        unit_label = "kgCO₂e/m² GIA" if use_intensity_flag else "kgCO₂e"
        st.markdown(f"**{row['rics_high_clean']}**  \n{row['value']:.0f} {unit_label}")

    _render_high_level_legend_generic(legend_df=legend_df, use_intensity=use_intensity, render_text=render_text)

def render_high_level_legend_two_charts(legend_df: pd.DataFrame, use_intensity: bool):
    def render_text(row, use_intensity_flag: bool):
        unit_label = "kgCO₂e/m² GIA" if use_intensity_flag else "kgCO₂e"
        st.markdown(
            f"**{row['rics_high_clean']}**  \n"
            f"Upfront: {row['upfront_value']:.0f} {unit_label}  \n"
            f"Life cycle: {row['whole_life_value']:.0f} {unit_label}"
        )

    _render_high_level_legend_generic(legend_df=legend_df, use_intensity=use_intensity, render_text=render_text)

def should_render_rics_streamlit_legend(show_legend: bool) -> bool:
    return bool(show_legend)

def should_show_plotly_rics_legend() -> bool:
    return False

# =============================================================================
# SECTION 9 — PLOTLY CHART BUILDERS
# Builds the single-stack and two-stack RICS charts while keeping plotting logic
# isolated from Streamlit page orchestration.
# =============================================================================
def plot_rics_single_stack(
    rows: pd.DataFrame,
    modules: list[str],
    em_col: str,
    colour_map: dict[str, str],
    height_px: int,
    bar_width: float,
    small_segment_threshold: float,
    segment_border_px: float,
    title: str,
    use_intensity: bool,
    gia_m2: float,
    display_map: dict[str, str],
    group_map: dict[str, str],
    contingency_map: dict[str, float],
    project_contingency_pct: float,
    contingency_colour: str,
    biogenic_colour: str,
    show_legend: bool,
    y_min: float | None = None,
    y_max: float | None = None,
    y_dtick: float | None = None,
    collapsed_high_level_labels: bool = False,
    target_line_value: float | None = None,
    target_line_label: str = "",
    target_line_colour: str = "#D62728",
    target_line_value_2: float | None = None,
    target_line_label_2: str = "",
    target_line_colour_2: str = "#1F77B4",
    bar_label_font_size: int = 12,
) -> go.Figure:
    rr = derive_level2_label(rows).copy()
    modules_expanded = expand_modules(modules)
    rr = rr[rr["section"].isin(modules_expanded)]
    if rr.empty:
        raise ValueError("No rows found for the selected module set.")

    value_col = pick_chart_value_col(rr, em_col)

    if use_intensity:
        if not gia_m2 or gia_m2 <= 0:
            raise ValueError("Intensity is enabled but GIA is not set (> 0).")
        scale = 1.0 / float(gia_m2)
        y_unit = "kgCO₂e/m² GIA"
    else:
        scale = 1.0
        y_unit = "kgCO₂e"

    rr["rics_high_orig"] = rr["rics_high_label"].astype(str)
    rr["rics_level2_orig"] = rr["rics_level2_label"].astype(str)

    rr["rics_high_clean"] = rr["rics_high_label"].map(clean_rics_label)
    rr = apply_stack_label_mapping(rr, display_map=display_map, group_map=group_map)

    agg = (
        rr.groupby(
            ["rics_high_orig", "rics_high_clean", "rics_level2_orig", "lvl2_id", "lvl2_final"],
            dropna=False,
        )[value_col]
        .sum()
        .reset_index()
    )
    agg["high_key"] = agg["rics_high_orig"].map(parse_rics_numeric_tuple)
    agg["lvl2_key"] = agg["rics_level2_orig"].map(parse_rics_numeric_tuple)

    high_order = build_rics_high_order(agg)

    agg_order = (
        agg.groupby(["rics_high_clean", "lvl2_final"], dropna=False)["lvl2_key"]
        .min()
        .reset_index()
        .rename(columns={"lvl2_key": "lvl2_key_min"})
    )
    agg = agg.merge(agg_order, on=["rics_high_clean", "lvl2_final"], how="left")

    lvl2_order = {}
    for h in high_order:
        sub = agg[agg["rics_high_clean"] == h].sort_values("lvl2_key_min", ascending=True)
        lvl2_order[h] = [s for s in sub["lvl2_final"].astype(str).unique().tolist() if s and s.lower() != "nan"]

    top5_map = compute_top_contributors_by_stack(
        rows=rows,
        modules=modules,
        value_col=value_col,
        display_map=display_map,
        group_map=group_map,
        top_n=5,
    )

    palettes = {}
    shade_index = {}
    for h in high_order:
        base = colour_map.get(h, "#4c78a8")
        segs = lvl2_order.get(h, [])
        palettes[h] = shade_palette(base, max(1, len(segs)))
        shade_index[h] = {seg: i for i, seg in enumerate(segs)}

    fig = go.Figure()
    x = [""]
    seen_legend_groups = set()
    base_segments = []
    cumulative = 0.0

    for h in high_order:
        for seg in lvl2_order.get(h, []):
            v_raw = float(
                agg.loc[(agg["rics_high_clean"] == h) & (agg["lvl2_final"] == seg), value_col]
                .fillna(0)
                .sum()
            )
            if v_raw <= 1e-9:
                continue

            v = v_raw * scale
            i = shade_index[h].get(seg, 0)
            colour = palettes[h][min(i, len(palettes[h]) - 1)]

            mid = cumulative + v / 2.0
            cumulative += v
            base_segments.append((h, seg, v, mid, colour))

    base_by_id = (
        rr.groupby("lvl2_id", dropna=False)[value_col]
        .sum()
        .reset_index()
        .rename(columns={value_col: "base_value"})
    )

    contingency_total_raw = 0.0
    project_frac = float(project_contingency_pct or 0.0) / 100.0

    for _, row in base_by_id.iterrows():
        sid = str(row["lvl2_id"])
        base_val = float(row["base_value"] or 0.0)
        subcat_frac = float(contingency_map.get(sid, 0.0) or 0.0)
        total_frac = project_frac + subcat_frac
        contingency_total_raw += base_val * total_frac

    contingency_total = contingency_total_raw * scale

    segments = list(base_segments)
    if contingency_total > 1e-9:
        mid = cumulative + contingency_total / 2.0
        cumulative += contingency_total
        segments.append(("Contingency", "Contingency", contingency_total, mid, contingency_colour))

    total_pos = sum(v for (_, _, v, _, _) in segments) if segments else 0.0

    small_set = set()
    for h, seg, v, _, _ in segments:
        if h == "Contingency":
            continue
        share = (v / total_pos) if total_pos else 0.0
        if share < small_segment_threshold:
            small_set.add((h, seg))

    if collapsed_high_level_labels:
        segment_line_color = "rgba(255,255,255,0.25)"
        segment_line_width = 0.5
    else:
        segment_line_color = "white"
        segment_line_width = segment_border_px

    include_plotly_legend = should_show_plotly_rics_legend()

    for h, seg, v, _, colour in segments:
        is_cont = h == "Contingency"
        is_small = (h, seg) in small_set
        text_colour = "white" if luminance_from_hex(colour) < 0.5 else "black"

        if collapsed_high_level_labels:
            text = None
            textfont = None
        else:
            if (not is_small) or is_cont:
                text = [f"{seg} {v:,.0f}"]
                textfont = dict(color=text_colour, size=bar_label_font_size)
            else:
                text = None
                textfont = None

        hover = (
            "<b>Total</b><br>"
            f"High-level: {h}<br>"
            f"Sub-category: {seg}<br>"
            f"Segment value: %{{y:,.0f}} {y_unit}<br>"
        )
        if h not in {"Contingency", "Biogenic (total)"}:
            hover += "<br><b>Top contributors (by Comment)</b><br>%{customdata}"
        hover += "<extra></extra>"

        legend_name = h
        showlegend_flag = include_plotly_legend and (h not in seen_legend_groups)

        fig.add_trace(
            go.Bar(
                x=x,
                y=[v],
                width=bar_width,
                marker=dict(
                    color=colour,
                    line=dict(color=segment_line_color, width=segment_line_width),
                ),
                name=legend_name,
                legendgroup=h,
                showlegend=showlegend_flag,
                customdata=[top5_map.get(seg, "")],
                hovertemplate=hover,
                text=text,
                textangle=0,
                textposition="inside",
                insidetextanchor="middle",
                textfont=textfont,
            )
        )

        if showlegend_flag:
            seen_legend_groups.add(h)

    if collapsed_high_level_labels:
        high_level_df = (
            agg.groupby("rics_high_clean", dropna=False)[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "high_value_raw"})
        )

        high_level_df["order"] = high_level_df["rics_high_clean"].apply(
            lambda x: high_order.index(x) if x in high_order else 999
        )
        high_level_df = high_level_df.sort_values("order").reset_index(drop=True)
        high_level_df["high_value"] = high_level_df["high_value_raw"] * scale

        high_segments = []
        running_y = 0.0
        for _, row in high_level_df.iterrows():
            high_name = row["rics_high_clean"]
            high_value = float(row["high_value"] or 0.0)

            if high_value <= 1e-9:
                continue

            mid_y = running_y + high_value / 2.0
            running_y += high_value

            base_colour = colour_map.get(high_name, "#4c78a8")
            text_colour = "white" if luminance_from_hex(base_colour) < 0.7 else "black"

            high_segments.append(
                {
                    "name": high_name,
                    "value": high_value,
                    "mid_y": mid_y,
                    "colour": base_colour,
                    "text_colour": text_colour,
                }
            )

        if contingency_total > 1e-9:
            cont_mid_y = running_y + contingency_total / 2.0
            running_y += contingency_total
            cont_text_colour = "white" if luminance_from_hex(contingency_colour) < 0.7 else "black"

            high_segments.append(
                {
                    "name": "Contingency",
                    "value": contingency_total,
                    "mid_y": cont_mid_y,
                    "colour": contingency_colour,
                    "text_colour": cont_text_colour,
                }
            )

        high_small_threshold = max(small_segment_threshold, 0.06)

        small_high_segments = []
        for seg_row in high_segments:
            share = (seg_row["value"] / total_pos) if total_pos else 0.0

            if share >= high_small_threshold:
                fig.add_annotation(
                    x=0.5,
                    y=seg_row["mid_y"],
                    xref="paper",
                    yref="y",
                    text=f"{seg_row['name']} {seg_row['value']:,.0f}",
                    showarrow=False,
                    font=dict(size=bar_label_font_size + 2, color=seg_row["text_colour"]),
                    align="center",
                    xanchor="center",
                    yanchor="middle",
                )
            else:
                small_high_segments.append(seg_row)

        if small_high_segments:
            small_high_segments = sorted(small_high_segments, key=lambda d: d["mid_y"], reverse=True)

            y_top = max(s["mid_y"] for s in small_high_segments)
            y_bottom = min(s["mid_y"] for s in small_high_segments)
            n = len(small_high_segments)

            if n == 1:
                y_slots = [small_high_segments[0]["mid_y"]]
            else:
                pad = max((y_top - y_bottom) * 0.15, 12.0)
                y_slots = [
                    (y_top + pad) - ((y_top - y_bottom) + 2 * pad) * (i / (n - 1))
                    for i in range(n)
                ]

            x_bar = 0.50
            x_text = 0.90

            for seg_row, y_label in zip(small_high_segments, y_slots):
                fig.add_shape(
                    type="line",
                    xref="paper",
                    yref="y",
                    x0=x_bar,
                    y0=seg_row["mid_y"],
                    x1=x_text - 0.01,
                    y1=y_label,
                    line=dict(color="rgba(60,60,60,0.7)", width=1),
                )

                fig.add_annotation(
                    x=x_text,
                    xref="paper",
                    y=y_label,
                    yref="y",
                    text=f"{seg_row['name']} {seg_row['value']:,.0f}",
                    showarrow=False,
                    bgcolor="rgba(255,255,255,0.85)",
                    font=dict(size=bar_label_font_size, color="black"),
                    align="center",
                    xanchor="left",
                )

    bio_top = compute_top_biogenic_contributors(rows, em_col=em_col, top_n=5)
    bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
    bio_mid = None
    if bio_total > 1e-9:
        v = -abs(bio_total)
        bio_mid = v / 2.0
        fig.add_trace(
            go.Bar(
                x=x,
                y=[v],
                width=bar_width,
                marker=dict(
                    color=biogenic_colour,
                    line=dict(color=segment_line_color, width=segment_line_width),
                ),
                name="Biogenic (total)",
                customdata=[bio_top],
                textangle=0,
                hovertemplate=(
                    "<b>Biogenic</b><br>"
                    f"Value: %{{y:,.0f}} {y_unit}<br>"
                    "<br><b>Top contributors (within biogenic)</b><br>%{customdata}"
                    "<extra></extra>"
                ),
                text=None,
                showlegend=False,
            )
        )

    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        xaxis_title="",
        yaxis_title=y_unit,
        margin=dict(l=40, r=340, t=70, b=80),
        xaxis=dict(tickfont=dict(size=17), title=dict(font=dict(size=19))),
        yaxis=dict(tickfont=dict(size=17), title=dict(font=dict(size=19))),
        legend=dict(font=dict(size=17)),
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )

    if bio_total > 1e-9:
        bio_floor = -abs(bio_total) * 1.15
    else:
        bio_floor = None

    final_y_min = y_min
    if final_y_min is not None and bio_floor is not None:
        final_y_min = min(final_y_min, bio_floor)
    elif final_y_min is None and bio_floor is not None:
        final_y_min = bio_floor

    yaxis_kwargs = {}
    if final_y_min is not None or y_max is not None:
        yaxis_kwargs["range"] = [final_y_min if final_y_min is not None else 0.0, y_max if y_max is not None else None]
    if y_dtick is not None:
        yaxis_kwargs["dtick"] = y_dtick

    if yaxis_kwargs:
        fig.update_yaxes(**yaxis_kwargs)

    small_segments = [(h, seg, v, mid, colour) for (h, seg, v, mid, colour) in segments if (h, seg) in small_set]
    small_sorted = sorted(small_segments, key=lambda t: t[3], reverse=True)

    if (not collapsed_high_level_labels) and small_sorted and total_pos > 0:
        y_top = total_pos * 0.98
        y_bottom = total_pos * 0.60
        n = len(small_sorted)
        y_slots = [small_sorted[0][3]] if n == 1 else [
            y_top - (y_top - y_bottom) * (i / (n - 1)) for i in range(n)
        ]

        x_bar = 0.50
        x_text = 0.90

        for (_, seg, v, mid, _), y_label in zip(small_sorted, y_slots):
            fig.add_shape(
                type="line",
                xref="paper",
                yref="y",
                x0=x_bar,
                y0=mid,
                x1=x_text - 0.01,
                y1=y_label,
                line=dict(color="rgba(60,60,60,0.7)", width=1),
            )
            fig.add_annotation(
                x=x_text,
                xref="paper",
                y=y_label,
                yref="y",
                text=f"{seg} {v:,.0f}",
                showarrow=False,
                bgcolor="rgba(255,255,255,0.85)",
                font=dict(size=max(bar_label_font_size - 1, 8), color="black"),
                align="center",
                xanchor="left",
            )

    if bio_total > 1e-9 and bio_mid is not None:
        bio_label_y = -abs(bio_total) * 0.55
        x_bar = 0.50
        x_text = 0.90

        fig.add_shape(
            type="line",
            xref="paper",
            yref="y",
            x0=x_bar,
            y0=bio_mid,
            x1=x_text - 0.01,
            y1=bio_label_y,
            line=dict(color="rgba(60,60,60,0.7)", width=1),
        )

        fig.add_annotation(
            x=x_text,
            xref="paper",
            y=bio_label_y,
            yref="y",
            text=f"Biogenic {-abs(bio_total):,.0f}",
            showarrow=False,
            bgcolor="rgba(255,255,255,0.85)",
            font=dict(size=max(bar_label_font_size - 1, 8), color="black"),
            align="center",
            xanchor="left",
        )

    if target_line_value is not None:
        fig.add_hline(
            y=target_line_value,
            line_width=2,
            line_dash="dot",
            line_color=target_line_colour,
        )

        if target_line_label.strip():
            fig.add_annotation(
                x=0.5,
                xref="paper",
                y=target_line_value,
                yref="y",
                text=target_line_label,
                showarrow=False,
                font=dict(size=bar_label_font_size, color=target_line_colour),
                bgcolor="rgba(255,255,255,0.75)",
                xanchor="center",
                yanchor="bottom",
                align="center",
            )

    if target_line_value_2 is not None:
        fig.add_hline(
            y=target_line_value_2,
            line_width=2,
            line_dash="dot",
            line_color=target_line_colour_2,
        )

        if target_line_label_2.strip():
            fig.add_annotation(
                x=0.5,
                xref="paper",
                y=target_line_value_2,
                yref="y",
                text=target_line_label_2,
                showarrow=False,
                font=dict(size=bar_label_font_size, color=target_line_colour_2),
                bgcolor="rgba(255,255,255,0.75)",
                xanchor="center",
                yanchor="bottom",
                align="center",
            )

    return fig

def plot_rics_two_stacks(
    rows: pd.DataFrame,
    upfront_modules: list[str],
    whole_life_modules: list[str],
    em_col: str,
    colour_map: dict[str, str],
    height_px: int,
    bar_width: float,
    segment_border_px: float,
    title: str,
    use_intensity: bool,
    gia_m2: float,
    display_map: dict[str, str],
    group_map: dict[str, str],
    contingency_map: dict[str, float],
    project_contingency_pct: float,
    contingency_colour: str,
    biogenic_colour: str,
    show_legend: bool,
    y_min: float | None = None,
    y_max: float | None = None,
    y_dtick: float | None = None,
    small_segment_threshold: float = 0.03,
    target_line_value: float | None = None,
    target_line_label: str = "",
    target_line_colour: str = "#D62728",
    target_line_value_2: float | None = None,
    target_line_label_2: str = "",
    target_line_colour_2: str = "#1F77B4",
    collapsed_high_level_labels: bool = False,
    bar_label_font_size: int = 12,
    limit_target_line_to_one_bar: bool = False,
    target_line_bar_scope: str = "Upfront",
    limit_target_line_2_to_one_bar: bool = False,
    target_line_2_bar_scope: str = "Whole life cycle",
) -> go.Figure:
    def build_stack_data(
        rows_in: pd.DataFrame,
        modules: list[str],
        add_biogenic: bool,
    ):
        rr = derive_level2_label(rows_in).copy()
        modules_expanded = expand_modules(modules)
        rr = rr[rr["section"].isin(modules_expanded)].copy()
        if rr.empty:
            return [], [], 0.0

        value_col = pick_chart_value_col(rr, em_col)

        if use_intensity:
            if not gia_m2 or gia_m2 <= 0:
                raise ValueError("Intensity is enabled but GIA is not set (> 0).")
            scale = 1.0 / float(gia_m2)
        else:
            scale = 1.0

        rr["rics_high_orig"] = rr["rics_high_label"].astype(str)
        rr["rics_level2_orig"] = rr["rics_level2_label"].astype(str)
        rr["rics_high_clean"] = rr["rics_high_label"].map(clean_rics_label)
        rr = apply_stack_label_mapping(rr, display_map=display_map, group_map=group_map)

        agg = (
            rr.groupby(
                ["rics_high_orig", "rics_high_clean", "rics_level2_orig", "lvl2_id", "lvl2_final"],
                dropna=False,
            )[value_col]
            .sum()
            .reset_index()
        )
        agg["high_key"] = agg["rics_high_orig"].map(parse_rics_numeric_tuple)
        agg["lvl2_key"] = agg["rics_level2_orig"].map(parse_rics_numeric_tuple)

        high_order = build_rics_high_order(agg)

        agg_order = (
            agg.groupby(["rics_high_clean", "lvl2_final"], dropna=False)["lvl2_key"]
            .min()
            .reset_index()
            .rename(columns={"lvl2_key": "lvl2_key_min"})
        )
        agg = agg.merge(agg_order, on=["rics_high_clean", "lvl2_final"], how="left")

        lvl2_order = {}
        for h in high_order:
            sub = agg[agg["rics_high_clean"] == h].sort_values("lvl2_key_min", ascending=True)
            lvl2_order[h] = [
                s for s in sub["lvl2_final"].astype(str).unique().tolist()
                if s and s.lower() != "nan"
            ]

        palettes = {}
        shade_index = {}
        for h in high_order:
            base = colour_map.get(h, "#4c78a8")
            segs = lvl2_order.get(h, [])
            palettes[h] = shade_palette(base, max(1, len(segs)))
            shade_index[h] = {seg: i for i, seg in enumerate(segs)}

        top5_map = compute_top_contributors_by_stack(
            rows=rows_in,
            modules=modules,
            value_col=value_col,
            display_map=display_map,
            group_map=group_map,
            top_n=5,
        )

        segments = []
        cumulative = 0.0

        for h in high_order:
            for seg in lvl2_order.get(h, []):
                v_raw = float(
                    agg.loc[
                        (agg["rics_high_clean"] == h) & (agg["lvl2_final"] == seg),
                        value_col,
                    ]
                    .fillna(0)
                    .sum()
                )
                if v_raw <= 1e-9:
                    continue

                v = v_raw * scale
                i = shade_index[h].get(seg, 0)
                colour = palettes[h][min(i, len(palettes[h]) - 1)]
                mid = cumulative + v / 2.0
                cumulative += v

                segments.append(
                    {
                        "high": h,
                        "seg": seg,
                        "value": v,
                        "mid": mid,
                        "colour": colour,
                        "customdata": top5_map.get(seg, ""),
                    }
                )

        base_by_id = (
            rr.groupby("lvl2_id", dropna=False)[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "base_value"})
        )
        contingency_total_raw = 0.0
        project_frac = float(project_contingency_pct or 0.0) / 100.0

        for _, row in base_by_id.iterrows():
            sid = str(row["lvl2_id"])
            base_val = float(row["base_value"] or 0.0)
            subcat_frac = float(contingency_map.get(sid, 0.0) or 0.0)
            contingency_total_raw += base_val * (project_frac + subcat_frac)

        contingency_total = contingency_total_raw * scale
        if contingency_total > 1e-9:
            mid = cumulative + contingency_total / 2.0
            cumulative += contingency_total
            segments.append(
                {
                    "high": "Contingency",
                    "seg": "Contingency",
                    "value": contingency_total,
                    "mid": mid,
                    "colour": contingency_colour,
                    "customdata": "",
                }
            )

        high_level_df = (
            agg.groupby("rics_high_clean", dropna=False)[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "high_value_raw"})
        )
        high_level_df["order"] = high_level_df["rics_high_clean"].apply(
            lambda x: high_order.index(x) if x in high_order else 999
        )
        high_level_df = high_level_df.sort_values("order").reset_index(drop=True)
        high_level_df["high_value"] = high_level_df["high_value_raw"] * scale

        high_segments = []
        running_y = 0.0
        for _, row in high_level_df.iterrows():
            high_name = row["rics_high_clean"]
            high_value = float(row["high_value"] or 0.0)
            if high_value <= 1e-9:
                continue

            mid_y = running_y + high_value / 2.0
            running_y += high_value
            base_colour = colour_map.get(high_name, "#4c78a8")
            text_colour = "white" if luminance_from_hex(base_colour) < 0.7 else "black"

            high_segments.append(
                {
                    "name": high_name,
                    "value": high_value,
                    "mid_y": mid_y,
                    "colour": base_colour,
                    "text_colour": text_colour,
                }
            )

        if contingency_total > 1e-9:
            cont_mid_y = running_y + contingency_total / 2.0
            running_y += contingency_total
            cont_text_colour = "white" if luminance_from_hex(contingency_colour) < 0.7 else "black"

            high_segments.append(
                {
                    "name": "Contingency",
                    "value": contingency_total,
                    "mid_y": cont_mid_y,
                    "colour": contingency_colour,
                    "text_colour": cont_text_colour,
                }
            )

        bio_value = 0.0
        if add_biogenic:
            bio_total = compute_biogenic_total(rows_in, modules=modules, em_col=em_col, scale=scale)
            if bio_total > 1e-9:
                bio_value = -abs(bio_total)

        return segments, high_segments, bio_value

    if use_intensity:
        y_unit = "kgCO₂e/m² GIA"
    else:
        y_unit = "kgCO₂e"

    upfront_segments, upfront_high_segments, upfront_bio = build_stack_data(rows, upfront_modules, add_biogenic=True)
    whole_life_segments, whole_life_high_segments, whole_life_bio = build_stack_data(rows, whole_life_modules, add_biogenic=True)

    upfront_total = sum(s["value"] for s in upfront_segments) + upfront_bio
    whole_life_total = sum(s["value"] for s in whole_life_segments) + whole_life_bio

    def fmt_total(v: float) -> str:
        return f"{v:,.0f} {y_unit}"

    fig = go.Figure()

    x_upfront = 0.0
    x_wlc = 1.25
    x_labels = {
        x_upfront: f"<span style='font-size:20px'>Upfront</span><br><span style='font-size:16px'>{fmt_total(upfront_total)}</span>",
        x_wlc: f"<span style='font-size:20px'>Whole life cycle</span><br><span style='font-size:16px'>{fmt_total(whole_life_total)}</span>",
    }
    actual_bar_width = 0.58 * bar_width

    if collapsed_high_level_labels:
        segment_line_color = "rgba(255,255,255,0.20)"
        segment_line_width = 0.5
    else:
        segment_line_color = "white"
        segment_line_width = segment_border_px

    def add_bar_segments(x_pos: float, segments: list[dict], total_pos: float):
        small_segments = []

        for seg_row in segments:
            seg = seg_row["seg"]
            v = seg_row["value"]
            colour = seg_row["colour"]
            share = (v / total_pos) if total_pos else 0.0
            is_small = seg_row["high"] != "Contingency" and share < small_segment_threshold
            text_colour = "white" if luminance_from_hex(colour) < 0.5 else "black"

            hover = (
                "<b>Total</b><br>"
                f"High-level: {seg_row['high']}<br>"
                f"Sub-category: {seg}<br>"
                f"Segment value: %{{y:,.0f}} {y_unit}<br>"
            )
            if seg_row["high"] not in {"Contingency", "Biogenic (total)"}:
                hover += "<br><b>Top contributors (by Comment)</b><br>%{customdata}"
            hover += "<extra></extra>"

            if collapsed_high_level_labels:
                text = None
                textfont = None
            else:
                text = [f"{seg} {v:,.0f}"] if not is_small else None
                textfont = dict(color=text_colour, size=bar_label_font_size) if not is_small else None

            fig.add_trace(
                go.Bar(
                    x=[x_pos],
                    y=[v],
                    width=actual_bar_width,
                    marker=dict(color=colour, line=dict(color=segment_line_color, width=segment_line_width)),
                    customdata=[seg_row["customdata"]],
                    hovertemplate=hover,
                    text=text,
                    textangle=0,
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=textfont,
                    showlegend=False,
                )
            )

            if is_small:
                small_segments.append(seg_row)

        return small_segments

    upfront_total_pos = sum(s["value"] for s in upfront_segments)
    whole_life_total_pos = sum(s["value"] for s in whole_life_segments)

    upfront_small = add_bar_segments(x_upfront, upfront_segments, upfront_total_pos)
    whole_life_small = add_bar_segments(x_wlc, whole_life_segments, whole_life_total_pos)

    if upfront_bio < 0:
        fig.add_trace(
            go.Bar(
                x=[x_upfront],
                y=[upfront_bio],
                width=actual_bar_width,
                marker=dict(color=biogenic_colour, line=dict(color=segment_line_color, width=segment_line_width)),
                hovertemplate=f"<b>Biogenic</b> Value: %{{y:,.0f}} {y_unit}<extra></extra>",
                showlegend=False,
            )
        )

    if whole_life_bio < 0:
        fig.add_trace(
            go.Bar(
                x=[x_wlc],
                y=[whole_life_bio],
                width=actual_bar_width,
                marker=dict(color=biogenic_colour, line=dict(color=segment_line_color, width=segment_line_width)),
                hovertemplate=f"<b>Biogenic</b> Value: %{{y:,.0f}} {y_unit}<extra></extra>",
                showlegend=False,
            )
        )

    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        xaxis_title="",
        yaxis_title=y_unit,
        margin=dict(l=80, r=80, t=70, b=80),
        xaxis=dict(
            tickmode="array",
            tickvals=[x_upfront, x_wlc],
            ticktext=[x_labels[x_upfront], x_labels[x_wlc]],
            tickfont=dict(size=17),
            title=dict(font=dict(size=19)),
            range=[-0.75, 2.0],
        ),
        yaxis=dict(tickfont=dict(size=17), title=dict(font=dict(size=19))),
        legend=dict(font=dict(size=17)),
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )

    bio_values = [v for v in [upfront_bio, whole_life_bio] if v < 0]
    bio_floor = min(bio_values) * 1.15 if bio_values else None

    final_y_min = y_min
    if final_y_min is not None and bio_floor is not None:
        final_y_min = min(final_y_min, bio_floor)
    elif final_y_min is None and bio_floor is not None:
        final_y_min = bio_floor

    yaxis_kwargs = {}
    if final_y_min is not None or y_max is not None:
        yaxis_kwargs["range"] = [
            final_y_min if final_y_min is not None else 0.0,
            y_max if y_max is not None else None,
        ]
    if y_dtick is not None:
        yaxis_kwargs["dtick"] = y_dtick
    if yaxis_kwargs:
        fig.update_yaxes(**yaxis_kwargs)

    def add_target_line(
        value: float | None,
        label: str,
        colour: str,
        limit_to_one_bar: bool,
        bar_scope: str,
    ):
        if value is None:
            return

        if limit_to_one_bar:
            midpoint = (x_upfront + x_wlc) / 2.0
            if str(bar_scope).strip().lower().startswith("whole"):
                x0 = midpoint
                x1 = 2.0
            else:
                x0 = -0.75
                x1 = midpoint

            fig.add_shape(
                type="line",
                xref="x",
                yref="y",
                x0=x0,
                x1=x1,
                y0=value,
                y1=value,
                line=dict(width=2, dash="dot", color=colour),
            )
        else:
            fig.add_hline(
                y=value,
                line_width=2,
                line_dash="dot",
                line_color=colour,
            )

        if label.strip():
            fig.add_annotation(
                x=(x_upfront + x_wlc) / 2.0,
                xref="x",
                y=value,
                yref="y",
                text=label,
                showarrow=False,
                font=dict(size=bar_label_font_size, color=colour),
                bgcolor="rgba(255,255,255,0.75)",
                xanchor="center",
                yanchor="bottom",
                align="center",
            )

    add_target_line(
        value=target_line_value,
        label=target_line_label,
        colour=target_line_colour,
        limit_to_one_bar=limit_target_line_to_one_bar,
        bar_scope=target_line_bar_scope,
    )

    add_target_line(
        value=target_line_value_2,
        label=target_line_label_2,
        colour=target_line_colour_2,
        limit_to_one_bar=limit_target_line_2_to_one_bar,
        bar_scope=target_line_2_bar_scope,
    )

    def add_small_segment_leaders(
        x_pos: float,
        small_segments: list[dict],
        total_pos: float,
        side: str,
    ):
        if not small_segments or total_pos <= 0:
            return

        small_segments = sorted(small_segments, key=lambda d: d["mid"], reverse=True)

        y_top = total_pos * 0.98
        y_bottom = total_pos * 0.72
        n = len(small_segments)
        y_slots = [small_segments[0]["mid"]] if n == 1 else [
            y_top - (y_top - y_bottom) * (i / (n - 1)) for i in range(n)
        ]

        if side == "right":
            bar_edge = x_pos + actual_bar_width / 2.0
            text_x = x_pos + 0.42
            xanchor = "left"
            x1 = text_x - 0.03
        else:
            bar_edge = x_pos - actual_bar_width / 2.0
            text_x = x_pos - 0.42
            xanchor = "right"
            x1 = text_x + 0.03

        for seg_row, y_label in zip(small_segments, y_slots):
            fig.add_shape(
                type="line",
                xref="x",
                yref="y",
                x0=bar_edge,
                y0=seg_row["mid"],
                x1=x1,
                y1=y_label,
                line=dict(color="rgba(60,60,60,0.7)", width=1),
            )
            fig.add_annotation(
                x=text_x,
                xref="x",
                y=y_label,
                yref="y",
                text=f"{seg_row['seg']} {seg_row['value']:,.0f}",
                showarrow=False,
                bgcolor="rgba(255,255,255,0.85)",
                font=dict(size=max(bar_label_font_size - 1, 8), color="black"),
                align="center",
                xanchor=xanchor,
            )

    def add_high_level_labels(
        x_pos: float,
        high_segments: list[dict],
        total_pos: float,
        side_for_small: str,
    ):
        if not collapsed_high_level_labels:
            return

        high_small_threshold = max(small_segment_threshold, 0.06)
        small_high_segments = []

        for seg_row in high_segments:
            share = (seg_row["value"] / total_pos) if total_pos else 0.0
            if share >= high_small_threshold:
                fig.add_annotation(
                    x=x_pos,
                    xref="x",
                    y=seg_row["mid_y"],
                    yref="y",
                    text=f"{seg_row['name']} {seg_row['value']:,.0f}",
                    showarrow=False,
                    font=dict(size=bar_label_font_size + 2, color=seg_row["text_colour"]),
                    align="center",
                    xanchor="center",
                    yanchor="middle",
                )
            else:
                small_high_segments.append(seg_row)

        if not small_high_segments:
            return

        small_high_segments = sorted(small_high_segments, key=lambda d: d["mid_y"], reverse=True)

        y_top = max(s["mid_y"] for s in small_high_segments)
        y_bottom = min(s["mid_y"] for s in small_high_segments)
        n = len(small_high_segments)

        if n == 1:
            y_slots = [small_high_segments[0]["mid_y"]]
        else:
            pad = max((y_top - y_bottom) * 0.15, 12.0)
            y_slots = [
                (y_top + pad) - ((y_top - y_bottom) + 2 * pad) * (i / (n - 1))
                for i in range(n)
            ]

        if side_for_small == "right":
            bar_edge = x_pos + actual_bar_width / 2.0
            text_x = x_pos + 0.42
            xanchor = "left"
            x1 = text_x - 0.03
        else:
            bar_edge = x_pos - actual_bar_width / 2.0
            text_x = x_pos - 0.42
            xanchor = "right"
            x1 = text_x + 0.03

        for seg_row, y_label in zip(small_high_segments, y_slots):
            fig.add_shape(
                type="line",
                xref="x",
                yref="y",
                x0=bar_edge,
                y0=seg_row["mid_y"],
                x1=x1,
                y1=y_label,
                line=dict(color="rgba(60,60,60,0.7)", width=1),
            )

            fig.add_annotation(
                x=text_x,
                xref="x",
                y=y_label,
                yref="y",
                text=f"{seg_row['name']} {seg_row['value']:,.0f}",
                showarrow=False,
                bgcolor="rgba(255,255,255,0.85)",
                font=dict(size=bar_label_font_size, color="black"),
                align="center",
                xanchor=xanchor,
            )

    if collapsed_high_level_labels:
        add_high_level_labels(x_upfront, upfront_high_segments, upfront_total_pos, side_for_small="left")
        add_high_level_labels(x_wlc, whole_life_high_segments, whole_life_total_pos, side_for_small="right")
    else:
        add_small_segment_leaders(x_upfront, upfront_small, upfront_total_pos, side="left")
        add_small_segment_leaders(x_wlc, whole_life_small, whole_life_total_pos, side="right")

    def add_bio_label(x_pos: float, bio_value: float, side: str):
        if bio_value >= 0:
            return

        bio_mid = bio_value / 2.0
        if side == "right":
            bar_edge = x_pos + actual_bar_width / 2.0
            text_x = x_pos + 0.42
            xanchor = "left"
            x1 = text_x - 0.03
        else:
            bar_edge = x_pos - actual_bar_width / 2.0
            text_x = x_pos - 0.42
            xanchor = "right"
            x1 = text_x + 0.03

        bio_label_y = bio_value * 0.55

        fig.add_shape(
            type="line",
            xref="x",
            yref="y",
            x0=bar_edge,
            y0=bio_mid,
            x1=x1,
            y1=bio_label_y,
            line=dict(color="rgba(60,60,60,0.7)", width=1),
        )
        fig.add_annotation(
            x=text_x,
            xref="x",
            y=bio_label_y,
            yref="y",
            text=f"Biogenic {bio_value:,.0f}",
            showarrow=False,
            bgcolor="rgba(255,255,255,0.85)",
            font=dict(size=max(bar_label_font_size - 1, 8), color="black"),
            align="center",
            xanchor=xanchor,
        )

    add_bio_label(x_upfront, upfront_bio, side="left")
    add_bio_label(x_wlc, whole_life_bio, side="right")

    return fig

# =============================================================================
# SECTION 10 — MATERIALS CHART HELPERS
# Prepares rows for material treemaps and wraps imported material chart functions
# so materials logic stays separate from the RICS chart builders.
# =============================================================================
def _compute_materials_module_rows(rows: pd.DataFrame, module: str, em_col: str) -> pd.DataFrame:
    if module == "A1-A5":
        needed = {"A1-A3", "A4", "A5"}
        have = set(rows["section"].dropna().unique())
        if not needed.issubset(have):
            st.warning("A1-A5 requested, but one or more of A1-A3/A4/A5 is missing. Falling back to A1-A3 only.")
    return rows

def render_materials_charts(
    rows: pd.DataFrame,
    em_col: str,
    mat_module: str,
    building_label: str,
    chart_height: int,
):
    chart_title = f"{mat_module} emissions by material category — {building_label}"

    st.subheader("Materials breakdown (emissions and mass)")
    rows_for_mat = _compute_materials_module_rows(rows, module=mat_module, em_col=em_col)

    try:
        em_df, mass_df, mass_col = material_breakdown(rows_for_mat, module=mat_module)
    except Exception as e:
        st.error(f"Failed to compute materials breakdown: {e}")
        st.stop()

    current_fig = None
    c1, c2 = st.columns(2)
    with c1:
        fig3a = plot_material_treemap_emissions(em_df, title=chart_title)
        fig3a.update_layout(
            height=chart_height,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        current_fig = fig3a
        st.plotly_chart(fig3a, width="stretch")

    with c2:
        if mass_col:
            fig3b = plot_material_treemap_mass(
                mass_df,
                title=f"Mass by material category ({mass_col}) — {building_label}",
            )
            fig3b.update_layout(
                height=chart_height,
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig3b, width="stretch")
        else:
            st.info("No numeric mass column detected in this export, so the mass treemap is hidden.")

    return current_fig

# =============================================================================
# SECTION 11 — GLA EXPORT HELPERS
# Maps rows and entries into GLA Table 1 and Table 2 outputs and builds workbook
# and debug exports without mixing in Streamlit rendering logic.
# =============================================================================
def _collapse_gla_code(text: str) -> str:
    txt = str(text or "").strip()
    m = re.match(r"^(\d+(?:\.\d+)*)", txt)
    if not m:
        return ""

    code = m.group(1)

    if code.startswith("0.1"):
        return "0.1 Toxic Mat."
    if code.startswith("0.2"):
        return "0.2 Demolition"
    if code.startswith("0.3"):
        return "0.3 Supports"
    if code.startswith("0.4"):
        return "0.4 Groundworks"
    if code.startswith("0.5"):
        return "0.5 Diversion"

    if code == "1" or code.startswith("1."):
        return "1 Substructure"

    if code == "2.1" or code.startswith("2.1."):
        return "2.1 Frame"
    if code == "2.2" or code.startswith("2.2."):
        return "2.2 Upper Floors"
    if code == "2.3" or code.startswith("2.3."):
        return "2.3 Roof"
    if code == "2.4" or code.startswith("2.4."):
        return "2.4 Stairs & Ramps"
    if code == "2.5" or code.startswith("2.5."):
        return "2.5 Ext. Walls"
    if code == "2.6" or code.startswith("2.6."):
        return "2.6 Windows & Ext. Doors"
    if code == "2.7" or code.startswith("2.7."):
        return "2.7 Int. Walls & Partitions"
    if code == "2.8" or code.startswith("2.8."):
        return "2.8 Int. Doors"

    if code == "3" or code.startswith("3."):
        return "3 Finishes"
    if code == "4" or code.startswith("4."):
        return "4 Fittings, furnishings & equipments"
    if code == "5" or code.startswith("5."):
        return "5 Services (MEP)"
    if code == "6" or code.startswith("6."):
        return "6 Prefabricated"
    if code == "7" or code.startswith("7."):
        return "7 Existing bldg"
    if code == "8" or code.startswith("8."):
        return "8 Ext. works"

    return "Unclassified / Other"

def _collapse_gla_code_short(text_value: str) -> str:
    txt = str(text_value or "").strip()
    m = re.match(r"^(\d+(?:\.\d+)*)", txt)
    if not m:
        return ""
    code = m.group(1)

    if code.startswith("0."):
        return code
    if code == "1" or code.startswith("1."):
        return "1"
    if code == "2.1" or code.startswith("2.1."):
        return "2.1"
    if code == "2.2" or code.startswith("2.2."):
        return "2.2"
    if code == "2.3" or code.startswith("2.3."):
        return "2.3"
    if code == "2.4" or code.startswith("2.4."):
        return "2.4"
    if code == "2.5" or code.startswith("2.5."):
        return "2.5"
    if code == "2.6" or code.startswith("2.6."):
        return "2.6"
    if code == "2.7" or code.startswith("2.7."):
        return "2.7"
    if code == "2.8" or code.startswith("2.8."):
        return "2.8"
    if code == "3" or code.startswith("3."):
        return "3"
    if code == "4" or code.startswith("4."):
        return "4"
    if code == "5" or code.startswith("5."):
        return "5"
    if code == "6" or code.startswith("6."):
        return "6"
    if code == "7" or code.startswith("7."):
        return "7"
    if code == "8" or code.startswith("8."):
        return "8"
    return ""

def _infer_module_b(material_type: str, gla_category: str) -> str:
    mt = str(material_type or "").lower()
    cat = str(gla_category or "")

    if mt in ["concrete", "masonry", "reinforcement / steel", "structural steel"]:
        return "No replacement assumed within 60-year RSP"
    if mt == "timber":
        if cat.startswith("2.6") or cat.startswith("2.8") or cat.startswith("3"):
            return "Replacement may occur during study period"
        return "No replacement assumed within 60-year RSP"
    if mt in ["glass", "aluminium"]:
        return "Replacement assumed once during study period"
    if mt in ["plasterboard / gypsum"]:
        return "Replacement assumed once during study period"
    if mt in ["membranes / bituminous products"]:
        return "Periodic replacement assumed (20-25 years)"
    if mt in ["insulation"]:
        return "No replacement assumed (unless disturbed)"
    return "Standard replacement assumptions per RICS WLC v2"

def _backfill_material_from_eol(material_type: str, eol: str) -> str:
    mt = str(material_type or "").strip()
    if mt and mt != "Other":
        return mt
    e = str(eol or "").lower()
    if "wood" in e:
        return "Timber"
    if "brick" in e or "stone" in e:
        return "Masonry"
    if "steel" in e:
        return "Structural steel"
    if "glass" in e:
        return "Glass"
    if "aluminium" in e:
        return "Aluminium"
    if "gypsum" in e:
        return "Plasterboard / gypsum"
    return mt or "Other"

def _assign_gla_category_from_entry_row(row) -> str:
    txt = str(row.get("rics_detail_base", "") or "").strip()
    gla_code = _collapse_gla_code_short(txt)
    if gla_code:
        return gla_code

    q = str(row.get("Question", "") or "").strip().lower()
    if "substructure" in q or "foundation" in q:
        return "1"
    if "frame" in q:
        return "2.1"
    if "upper floor" in q or "upper floors" in q:
        return "2.2"
    if "roof" in q:
        return "2.3"
    if "stair" in q or "ramp" in q:
        return "2.4"
    if "external wall" in q:
        return "2.5"
    if "window" in q or "external door" in q:
        return "2.6"
    if "internal wall" in q or "partition" in q:
        return "2.7"
    if "internal door" in q:
        return "2.8"
    if "finish" in q:
        return "3"
    if "fitting" in q or "furnishing" in q or "equipment" in q:
        return "4"
    if "service" in q or "mep" in q:
        return "5"
    if "prefabricated" in q:
        return "6"
    if "existing building" in q:
        return "7"
    if "external works" in q:
        return "8"
    return ""

def _normalise_material_family_label(resource: str) -> str:
    txt = str(resource or "").strip()
    if not txt:
        return ""

    txt_l = txt.lower()

    patterns = [
        ("Concrete", [r"\bconcrete\b", r"\bcement\b", r"\bgrout\b", r"\bscreed\b", r"\bprecast\b"]),
        ("Reinforcement / steel", [r"\brebar\b", r"reinforc", r"\bmesh\b"]),
        ("Structural steel", [r"structural steel", r"\buc\b", r"\bub\b", r"\brhs\b", r"\bchs\b", r"\bshs\b", r"\bsteel\b"]),
        ("Aluminium", [r"aluminium", r"aluminum"]),
        ("Glass", [r"\bglass\b", r"glazing"]),
        ("Timber", [r"\btimber\b", r"\bwood\b", r"glulam", r"clt", r"plywood", r"mdf", r"osb"]),
        ("Plasterboard / gypsum", [r"plasterboard", r"gypsum", r"drywall"]),
        ("Masonry", [r"brick", r"block", r"masonry", r"stone"]),
        ("Insulation", [r"insulation", r"mineral wool", r"rockwool", r"glass wool", r"eps", r"xps", r"pir", r"pur"]),
        ("Membranes / bituminous products", [r"membrane", r"bitumen", r"felt", r"dpm"]),
        ("Plastics", [r"\bpvc\b", r"plastic", r"polyeth", r"polyprop"]),
        ("Copper", [r"\bcopper\b"]),
        ("Plant / equipment", [r"boiler", r"heat pump", r"mvhr", r"ahu", r"fan", r"chiller", r"pump"]),
    ]

    for family, pats in patterns:
        if any(re.search(p, txt_l) for p in pats):
            return family

    return "Other"

def build_gla_table1_export(rows: pd.DataFrame) -> pd.DataFrame:
    rr = rows.copy()
    if "section" not in rr.columns:
        return pd.DataFrame(
            columns=["gla_category", "Sequestered (or biogenic) carbon"] + GLA_STAGE_ORDER + ["A-C total", "A-D total"]
        )

    rr["section"] = rr["section"].astype(str).str.strip()
    rr.loc[rr["section"] == "A1-A5", "section"] = "A1-A3"

    value_col = "rics_allocated_value" if "rics_allocated_value" in rr.columns else pick_chart_value_col(rr, pick_emissions_col(rr))
    rr[value_col] = pd.to_numeric(rr[value_col], errors="coerce").fillna(0.0)

    rr["gla_category"] = ""
    if "rics_alloc_label" in rr.columns:
        rr["gla_category"] = rr["rics_alloc_label"].map(_collapse_gla_code)
    if "rics_level2_label" in rr.columns:
        rr["gla_category"] = rr["gla_category"].where(
            rr["gla_category"].astype(str).str.strip() != "",
            rr["rics_level2_label"].map(_collapse_gla_code),
        )
    if "rics_high_label" in rr.columns:
        rr["gla_category"] = rr["gla_category"].where(
            rr["gla_category"].astype(str).str.strip() != "",
            rr["rics_high_label"].map(_collapse_gla_code),
        )

    rr["gla_category"] = rr["gla_category"].fillna("").astype(str).str.strip()

    unmapped = rr[rr["gla_category"] == ""].copy()
    build_gla_table1_export.last_unmapped = unmapped.copy()

    mapped = rr[rr["gla_category"] != ""].copy()
    if mapped.empty:
        return pd.DataFrame(
            columns=["gla_category", "Sequestered (or biogenic) carbon"] + GLA_STAGE_ORDER + ["A-C total", "A-D total"]
        )

    bio_total_col = pd.Series(0.0, index=mapped.index)

    bioc_mask = mapped["section"].astype(str).str.strip().str.lower().eq("bioc")
    if bioc_mask.any():
        bio_total_col.loc[bioc_mask] = pd.to_numeric(
            mapped.loc[bioc_mask, value_col],
            errors="coerce",
        ).fillna(0.0)
    elif "biogenic_kgco2e" in mapped.columns:
        bio_total_col = pd.to_numeric(mapped["biogenic_kgco2e"], errors="coerce").fillna(0.0)

    mapped["_bio_value"] = bio_total_col
    print("tidy comparable stage total", rr.loc[rr["section"].isin(GLA_STAGE_ORDER), value_col].sum())
    print("unmapped comparable stage total", unmapped.loc[unmapped["section"].isin(GLA_STAGE_ORDER), value_col].sum())
    print("mapped comparable stage total", mapped.loc[mapped["section"].isin(GLA_STAGE_ORDER), value_col].sum())

    stage_check = (
        rr.groupby(["section", "gla_category"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .sort_values(value_col, ascending=False)
    )
    print(stage_check.head(100).to_string())
    modules_only = mapped[~mapped["section"].str.lower().eq("bioc")].copy()
    modules_only = modules_only[modules_only["section"].isin(GLA_STAGE_ORDER)].copy()

    out = (
        modules_only.groupby(["gla_category", "section"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .pivot_table(index="gla_category", columns="section", values=value_col, aggfunc="sum", fill_value=0.0)
        .reset_index()
    )
    out.columns = [str(c) for c in out.columns]

    bio_by_cat = (
        mapped.groupby("gla_category", dropna=False)["_bio_value"]
        .sum()
        .reset_index()
        .rename(columns={"_bio_value": "Sequestered (or biogenic) carbon"})
    )

    out = bio_by_cat.merge(out, on="gla_category", how="outer").fillna(0.0)

    for c in GLA_STAGE_ORDER:
        if c not in out.columns:
            out[c] = 0.0

    out["A-C total"] = out[[c for c in GLA_STAGE_ORDER if c != "D"]].sum(axis=1)
    out["A-D total"] = out[GLA_STAGE_ORDER].sum(axis=1)

    order_map = {name: i for i, name in enumerate(GLA_TABLE1_CATEGORY_ORDER)}
    out["_order"] = out["gla_category"].map(lambda x: order_map.get(x, 999))
    out = out.sort_values(["_order", "gla_category"]).drop(columns="_order").reset_index(drop=True)

    return out[["gla_category", "Sequestered (or biogenic) carbon"] + GLA_STAGE_ORDER + ["A-C total", "A-D total"]]

def build_gla_table2_export(
    entries: pd.DataFrame | None = None,
    rows: pd.DataFrame | None = None,
    min_group_mass_kg: float = 100.0,
    min_group_share_pct: float = 1.0,
) -> pd.DataFrame:
    cols = [
        "GLA category",
        "Material type",
        "Mass of raw materials kg",
        "Module B assumptions",
        "Material end of life scenarios (Module C)",
        "Estimated reusable materials kg",
        "Estimated recyclable materials kg",
        "Data quality flags",
    ]
    if entries is None or len(entries) == 0:
        return pd.DataFrame(columns=cols)

    ee = entries.copy()
    ee["resource"] = ee["Resource"].astype("string").fillna("").str.strip() if "Resource" in ee.columns else ""
    ee["mass_kg"] = pd.to_numeric(ee["mass_value"], errors="coerce").fillna(0.0) if "mass_value" in ee.columns else 0.0
    ee["reusable_kg"] = pd.to_numeric(ee["reusable_kg"], errors="coerce").fillna(0.0) if "reusable_kg" in ee.columns else 0.0
    ee["recyclable_kg"] = pd.to_numeric(ee["recyclable_kg"], errors="coerce").fillna(0.0) if "recyclable_kg" in ee.columns else 0.0
    ee["eol_process"] = ee["eol_process"].astype("string").fillna("").str.strip() if "eol_process" in ee.columns else ""

    ee = ee[(ee["mass_kg"] > 0) & (ee["resource"] != "")].copy()
    if ee.empty:
        return pd.DataFrame(columns=cols)

    ee["GLA category"] = ee.apply(_assign_gla_category_from_entry_row, axis=1)
    ee["Material type"] = ee["resource"].map(_normalise_material_family_label)
    ee["Material type"] = ee.apply(lambda r: _backfill_material_from_eol(r["Material type"], r["eol_process"]), axis=1)

    grouped = (
        ee.groupby(["GLA category", "Material type", "eol_process"], dropna=False)
        .agg(
            **{
                "Mass of raw materials kg": ("mass_kg", "sum"),
                "Estimated reusable materials kg": ("reusable_kg", "sum"),
                "Estimated recyclable materials kg": ("recyclable_kg", "sum"),
            }
        )
        .reset_index()
        .rename(columns={"eol_process": "Material end of life scenarios (Module C)"})
    )

    def build_quality_flags(row):
        flags = []
        mass = float(row.get("Mass of raw materials kg", 0.0) or 0.0)
        reuse = float(row.get("Estimated reusable materials kg", 0.0) or 0.0)
        recycle = float(row.get("Estimated recyclable materials kg", 0.0) or 0.0)
        eol = str(row.get("Material end of life scenarios (Module C)", "") or "").lower()
        mt = str(row.get("Material type", "") or "").lower()

        if str(row.get("GLA category", "") or "").strip() == "":
            flags.append("Unmapped category")
        if mt in {"", "other"}:
            flags.append("Unclear material")
        if not eol.strip():
            flags.append("Missing EOL")

        if mass > 0 and (reuse > mass or recycle > mass or (reuse + recycle) > mass):
            if mt in {"plasterboard / gypsum", "membranes / bituminous products", "insulation", "glass", "aluminium", "timber"}:
                flags.append("Lifecycle EOL > initial mass (replacement assumed)")
            elif mt in {"concrete", "masonry", "reinforcement / steel", "structural steel"}:
                flags.append("Check: EOL mass exceeds initial (unexpected)")
            else:
                flags.append("EOL mass exceeds initial (uncertain cause)")

        if "glass" in eol and mt not in {"glass", "aluminium"}:
            flags.append("Composite EOL (system-level)")
        if "wood" in eol and mt not in {"timber"}:
            flags.append("Composite EOL (system-level)")

        return "; ".join(flags)

    grouped["Module B assumptions"] = grouped.apply(
        lambda r: _infer_module_b(r["Material type"], r["GLA category"]),
        axis=1,
    )
    grouped["Data quality flags"] = grouped.apply(build_quality_flags, axis=1)

    grouped["Mass of raw materials kg"] = pd.to_numeric(grouped["Mass of raw materials kg"], errors="coerce").fillna(0.0)
    totals = grouped.groupby("GLA category", dropna=False)["Mass of raw materials kg"].transform("sum")
    grouped["_share"] = grouped["Mass of raw materials kg"] / totals.replace(0, pd.NA)
    grouped["_share"] = grouped["_share"].fillna(0.0)

    keep = (grouped["Mass of raw materials kg"] >= float(min_group_mass_kg)) | (grouped["_share"] >= float(min_group_share_pct) / 100.0)
    grouped = grouped[keep].copy()

    grouped = grouped[cols].sort_values(
        by=["GLA category", "Mass of raw materials kg", "Material type"],
        ascending=[True, False, True],
    ).reset_index(drop=True)

    return grouped

def build_gla_export_workbook_bytes(table1: pd.DataFrame, table2: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        table1.to_excel(writer, sheet_name="GLA Table 1", index=False)
        table2.to_excel(writer, sheet_name="GLA Table 2", index=False)
    output.seek(0)
    return output.getvalue()

def build_gla_table2_unmapped_debug(
    entries: pd.DataFrame | None = None,
    rows: pd.DataFrame | None = None,
) -> pd.DataFrame:
    columns_entries = [
        "source_file", "building_name", "_material_entry_id", "Resource", "Comment",
        "Question", "rics_detail_base", "mass_value", "eol_process", "reusable_kg", "recyclable_kg",
    ]
    columns_rows = [
        "source_file", "building_name", "_source_row_id", "Resource", "Comment", "Question",
        "section", "mass_value", "rics_alloc_label", "rics_level2_label", "rics_high_label",
    ]

    if entries is not None and len(entries):
        ee = entries.copy()
        for c in ["source_file", "building_name", "Resource", "Comment", "Question", "rics_detail_base", "eol_process"]:
            if c not in ee.columns:
                ee[c] = ""
            ee[c] = ee[c].astype("string").fillna("").str.strip()
        for c in ["mass_value", "reusable_kg", "recyclable_kg"]:
            if c not in ee.columns:
                ee[c] = 0.0
            ee[c] = pd.to_numeric(ee[c], errors="coerce").fillna(0.0)

        ee = ee[(ee["mass_value"] > 0) & (ee["Resource"] != "")].copy()
        ee["GLA category"] = ee.apply(_assign_gla_category_from_entry_row, axis=1)
        unmapped = ee[ee["GLA category"].astype(str).str.strip() == ""].copy()
        if unmapped.empty:
            return pd.DataFrame(columns=columns_entries)
        cols = [c for c in columns_entries if c in unmapped.columns]
        return unmapped[cols].sort_values(
            by=[c for c in ["source_file", "building_name", "Question", "Resource"] if c in unmapped.columns]
        ).reset_index(drop=True)

    if rows is not None and len(rows):
        rr = rows.copy()
        if "_source_row_id" in rr.columns:
            rr = rr.sort_values("_source_row_id").groupby("_source_row_id", dropna=False).first().reset_index()

        for c in ["source_file", "building_name", "Resource", "Comment", "Question", "section", "rics_alloc_label", "rics_level2_label", "rics_high_label"]:
            if c not in rr.columns:
                rr[c] = ""
            rr[c] = rr[c].astype("string").fillna("").str.strip()
        if "mass_value" not in rr.columns:
            rr["mass_value"] = 0.0
        rr["mass_value"] = pd.to_numeric(rr["mass_value"], errors="coerce").fillna(0.0)

        gla_col = pd.Series("", index=rr.index, dtype="string")
        if "rics_alloc_label" in rr.columns:
            gla_col = rr["rics_alloc_label"]
        if "rics_level2_label" in rr.columns:
            gla_col = gla_col.where(gla_col != "", rr["rics_level2_label"])
        if "rics_high_label" in rr.columns:
            gla_col = gla_col.where(gla_col != "", rr["rics_high_label"])

        rr["GLA category"] = gla_col.map(_collapse_gla_code_short)
        unmapped = rr[(rr["mass_value"] > 0) & (rr["Resource"] != "") & (rr["GLA category"] == "")].copy()
        if unmapped.empty:
            return pd.DataFrame(columns=columns_rows)
        cols = [c for c in columns_rows if c in unmapped.columns]
        return unmapped[cols].sort_values(
            by=[c for c in ["source_file", "building_name", "Question", "Resource"] if c in unmapped.columns]
        ).reset_index(drop=True)

    return pd.DataFrame(columns=columns_entries)

# =============================================================================
# SECTION 12 — STREAMLIT SIDEBAR CONTROLS
# Defines all sidebar inputs in one place and returns control bundles used by the
# main app flow.
# =============================================================================
def render_primary_sidebar_controls() -> dict:
    with st.sidebar:
        st.header("Controls")

        chart_choice = st.radio(
            "Chart",
            options=[
                "1) Upfront (A1–A5)",
                "2) Life cycle embodied (excl. B6, B7 and D)",
                "3) Upfront + life cycle embodied (1+2)",
                "4) Materials (A1–A3 emissions + mass) [WIP don't use]",
            ],
            index=0,
        )

        st.subheader("Sizing")
        chart_height = st.slider("Chart height in app (px)", 400, 1800, 1200, 50)
        bar_width = st.slider("Bar width (relative)", 0.2, 1.0, 0.66, 0.02)
        segment_border_px = st.slider("Segment border (px)", 0.0, 4.0, 1.0, 0.25)

        st.subheader("Label text")
        bar_label_font_size = st.slider("Bar label font size", 8, 24, 12, 1)

        st.subheader("Leader labels")
        small_segment_threshold = st.slider(
            "Label segments smaller than this share of total",
            min_value=0.0,
            max_value=0.10,
            value=0.03,
            step=0.005,
        )

        show_legend = st.checkbox("Show legend", value=True, key="show_legend_main")

        use_intensity = st.checkbox("Show intensity (per m² GIA)", value=True)

        st.subheader("Benchmark / target line")
        show_target_line = st.checkbox("Show benchmark / target line", value=False)

        target_line_value = st.number_input(
            "Benchmark / target value",
            value=0.0,
            step=1.0,
            help="Y-axis value for the horizontal benchmark / target line.",
        )

        target_line_label = st.text_input(
            "Benchmark / target label",
            value="Target",
        )

        target_line_colour = st.color_picker(
            "Benchmark / target colour",
            value="#D62728",
            key="target_line_colour",
        )

        limit_target_line_to_one_bar = st.checkbox(
            "Limit benchmark / target line to one bar in chart 3",
            value=False,
            help="Only affects '3) Upfront + life cycle embodied (1+2)'. The label stays centred between the two bars.",
        )

        target_line_bar_scope = st.selectbox(
            "Chart 3 benchmark / target line bar",
            options=["Upfront", "Whole life cycle"],
            index=0,
            disabled=not limit_target_line_to_one_bar,
        )

        st.subheader("Benchmark / target line 2")
        show_target_line_2 = st.checkbox("Show benchmark / target line 2", value=False)

        target_line_value_2 = st.number_input(
            "Benchmark / target value 2",
            value=0.0,
            step=1.0,
            help="Y-axis value for the second horizontal benchmark / target line.",
        )

        target_line_label_2 = st.text_input(
            "Benchmark / target label 2",
            value="Target 2",
        )

        target_line_colour_2 = st.color_picker(
            "Benchmark / target colour 2",
            value="#1F77B4",
            key="target_line_colour_2",
        )

        limit_target_line_2_to_one_bar = st.checkbox(
            "Limit benchmark / target line 2 to one bar in chart 3",
            value=False,
            help="Only affects '3) Upfront + life cycle embodied (1+2)'. The label stays centred between the two bars.",
        )

        target_line_2_bar_scope = st.selectbox(
            "Chart 3 benchmark / target line 2 bar",
            options=["Upfront", "Whole life cycle"],
            index=1,
            disabled=not limit_target_line_2_to_one_bar,
        )

        st.subheader("Y-axis")
        manual_y_axis = st.checkbox("Manual y-axis min/max", value=False)

        y_axis_min = st.number_input(
            "Y-axis min",
            value=0.0,
            step=10.0,
            disabled=not manual_y_axis,
        )
        y_axis_max = st.number_input(
            "Y-axis max",
            value=600.0,
            step=10.0,
            disabled=not manual_y_axis,
        )
        y_axis_dtick = st.number_input(
            "Y-axis increment",
            min_value=0.0,
            value=50.0,
            step=10.0,
            disabled=not manual_y_axis,
            help="Set to 0 to let Plotly choose automatically.",
        )

        st.subheader("PNG export")
        export_width_px = st.number_input("PNG width (px)", min_value=200, value=1600, step=100)
        export_height_px = st.number_input("PNG height (px)", min_value=200, value=1200, step=100)

        st.subheader("Materials")
        mat_module = st.selectbox("Module for materials treemap", options=["A1-A3", "A1-A5"], index=0)
        st.caption("A1–A5 is computed as A1-A3 + A4 + A5 when available.")

        st.subheader("Reconciliation")
        recon_tolerance_kg = st.number_input(
            "Allocation tolerance (kgCO₂e)",
            min_value=0.0,
            value=0.01,
            step=0.01,
            help="Rows with absolute gap above this are flagged.",
        )

    return {
        "chart_choice": chart_choice,
        "chart_height": chart_height,
        "bar_width": bar_width,
        "segment_border_px": segment_border_px,
        "bar_label_font_size": bar_label_font_size,
        "small_segment_threshold": small_segment_threshold,
        "show_legend": show_legend,
        "use_intensity": use_intensity,
        "show_target_line": show_target_line,
        "target_line_value": target_line_value,
        "target_line_label": target_line_label,
        "target_line_colour": target_line_colour,
        "limit_target_line_to_one_bar": limit_target_line_to_one_bar,
        "target_line_bar_scope": target_line_bar_scope,
        "show_target_line_2": show_target_line_2,
        "target_line_value_2": target_line_value_2,
        "target_line_label_2": target_line_label_2,
        "target_line_colour_2": target_line_colour_2,
        "limit_target_line_2_to_one_bar": limit_target_line_2_to_one_bar,
        "target_line_2_bar_scope": target_line_2_bar_scope,
        "manual_y_axis": manual_y_axis,
        "y_axis_min": y_axis_min,
        "y_axis_max": y_axis_max,
        "y_axis_dtick": y_axis_dtick,
        "export_width_px": export_width_px,
        "export_height_px": export_height_px,
        "mat_module": mat_module,
        "recon_tolerance_kg": recon_tolerance_kg,
    }
def render_buildings_sidebar_controls(building_meta_df: pd.DataFrame) -> tuple[list[str], bool]:
    with st.sidebar:
        st.subheader("Buildings")
        selected_buildings = st.multiselect(
            "Select building(s)",
            options=building_meta_df["building_name"].tolist(),
            default=building_meta_df["building_name"].tolist(),
            key="selected_buildings",
        )

        st.subheader("Labels")
        collapsed_high_level_labels = st.checkbox(
            "Show high-level RICS labels only",
            value=False,
            key="collapsed_high_level_labels",
        )

    return selected_buildings, collapsed_high_level_labels

def render_colour_controls_sidebar() -> tuple[dict[str, str], str, str]:
    colour_map = {}
    with st.sidebar:
        st.subheader("Colours (by high-level RICS category)")
        for h in RICS_HIGH_ORDER:
            default_colour = get_default_rics_colour(h)
            colour_map[h] = st.color_picker(h, value=default_colour, key=f"col_{h}")

        st.subheader("Other colours")
        contingency_colour = st.color_picker("Contingency colour", value="#6C6C6C", key="cont_colour")
        biogenic_colour = st.color_picker("Biogenic colour", value="#E657C7", key="bio_colour")

    return colour_map, contingency_colour, biogenic_colour

# =============================================================================
# SECTION 13 — STREAMLIT MAIN PAGE FLOW
# Orchestrates uploads, parsing, editing, chart rendering, and downloads in the
# order the user experiences the app.
# =============================================================================
def render_buildings_editor(building_meta_df: pd.DataFrame) -> pd.DataFrame:
    st.subheader("Buildings")
    st.caption("Assign a building name and GIA to each uploaded OneClick export.")

    edited_buildings = st.data_editor(
        building_meta_df.set_index("file_name", drop=False),
        width="stretch",
        num_rows="fixed",
        column_config={
            "file_name": st.column_config.TextColumn("Uploaded file", disabled=True),
            "building_name": st.column_config.TextColumn("Building name"),
            "gia_m2": st.column_config.NumberColumn("GIA (m²)", min_value=0.0, step=1.0),
        },
        hide_index=True,
        key="building_meta_editor",
    ).reset_index(drop=True)

    edited_buildings["building_name"] = edited_buildings["building_name"].fillna("").astype(str).str.strip()
    edited_buildings["gia_m2"] = pd.to_numeric(edited_buildings["gia_m2"], errors="coerce").fillna(0.0)
    st.session_state["building_meta_df"] = edited_buildings.copy()
    return edited_buildings.copy()

def load_project_rows_or_stop(
    uploaded_files_,
    building_meta_df: pd.DataFrame,
    selected_buildings: list[str],
    manual_file_obj,
):
    if not selected_buildings:
        st.info("Select at least one building to continue.")
        st.stop()

    try:
        rows, entries, meta_by_building, em_col, selected_meta = build_combined_rows(
            uploaded_files_=uploaded_files_,
            building_meta_df=building_meta_df,
            selected_buildings=selected_buildings,
        )
    except Exception as e:
        st.error(f"Failed to parse selected building files: {e}")
        st.stop()

    validate_required_columns(
        rows,
        ["section", "rics_detail", "rics_high_label", "rics_level2_label", "_source_row_id", em_col],
        "Parsed OneClick rows",
    )

    if manual_file_obj is not None:
        try:
            manual_rows = load_manual_intensity_rows(
                manual_file=manual_file_obj,
                building_meta_df=building_meta_df,
                selected_buildings=selected_buildings,
                em_col=em_col,
            )
            if not manual_rows.empty:
                rows = pd.concat([rows, manual_rows], ignore_index=True)
                st.success(f"Loaded {len(manual_rows)} manual rows")
            else:
                st.info("Manual CSV uploaded, but no rows matched the currently selected buildings.")
        except Exception as e:
            st.error(f"Manual upload failed: {e}")
            st.stop()

    if rows.empty:
        st.info("No rows found for the selected buildings.")
        st.stop()

    return rows, entries, meta_by_building, em_col, selected_meta

def build_project_summary(selected_buildings: list[str], selected_meta: pd.DataFrame) -> tuple[float, str, str]:
    total_gia = float(pd.to_numeric(selected_meta["gia_m2"], errors="coerce").fillna(0.0).sum())
    building_label = format_building_label(selected_buildings)
    selected_buildings_title = format_selected_buildings_title(selected_buildings, total_gia)
    return total_gia, building_label, selected_buildings_title

def render_project_metadata(meta_by_building: dict) -> None:
    with st.expander("Project metadata", expanded=False):
        st.json(meta_by_building)

def render_reconciliation_and_contingency(
    selected_meta: pd.DataFrame,
    total_gia: float,
    recon_summary: pd.DataFrame,
    recon_flagged: pd.DataFrame,
) -> float:
    with st.expander("Allocation reconciliation / coverage check", expanded=True):
        st.write(
            "This compares the original source-row carbon values against the sum of the horizontal RICS allocation values. "
            "A gap means carbon has not been fully assigned to categories, or has been over-assigned."
        )
        st.dataframe(recon_summary, width="stretch")

        overall_original = recon_summary["original_kgco2e"].sum()
        overall_allocated = recon_summary["allocated_kgco2e"].sum()
        overall_gap = recon_summary["gap_kgco2e"].sum()
        overall_gap_pct = (overall_gap / overall_original) if overall_original else 0.0

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Selected buildings", str(len(selected_meta)))
        c2.metric("Combined GIA", f"{total_gia:,.0f} m²")
        c3.metric("Original total", f"{overall_original:,.2f}")
        c4.metric("Allocated total", f"{overall_allocated:,.2f}")
        c5.metric("Gap", f"{overall_gap:,.2f}")
        st.caption(f"Gap %: {overall_gap_pct:.2%}")

        if len(recon_flagged):
            st.warning(f"{len(recon_flagged)} source rows are not fully reconciled.")
            st.dataframe(recon_flagged.head(200), width="stretch")
        else:
            st.success("All source rows reconcile within tolerance.")

        st.subheader("Contingency")
        project_contingency_pct = st.number_input(
            "Project contingency (%)",
            min_value=0.0,
            max_value=100.0,
            step=0.5,
            key="project_contingency_pct_main",
            help="Applied to all plotted OneClick and manual entries before the separate contingency stack is calculated.",
        )

    return project_contingency_pct

def build_module_selections(rows: pd.DataFrame) -> tuple[list[str], list[str]]:
    available_modules = sorted(set(rows["section"].dropna().unique()))
    default_wlc = [m for m in available_modules if m not in {"B6", "B7", "D", "TOTAL", "bioC"}]
    upfront_modules = ["A1-A5"]
    wlc_modules_selected = default_wlc
    return upfront_modules, wlc_modules_selected

def render_chart_items_editor(rows: pd.DataFrame) -> pd.DataFrame:
    subcat_df = sync_subcat_editor_state(rows).copy()

    with st.expander("Chart items / grouping / visibility", expanded=True):
        st.caption(
            "This table includes OneClick and manual import items. Rename labels, merge items into groups, set contingency, and toggle items on or off for the charts."
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Reset chart items table", key="reset_chart_items"):
                st.session_state["subcat_map_df"] = build_subcat_editor_df(rows).copy()
                subcat_df = st.session_state["subcat_map_df"].copy()
        with c2:
            if st.button("Clear grouping only", key="clear_grouping_only"):
                tmp = st.session_state["subcat_map_df"].copy()
                tmp["group_name"] = ""
                st.session_state["subcat_map_df"] = tmp
                subcat_df = tmp.copy()

        subcat_df = subcat_df.set_index("subcategory_id", drop=False)

        edited = st.data_editor(
            subcat_df,
            width="stretch",
            num_rows="fixed",
            column_config={
                "subcategory_id": st.column_config.TextColumn("ID (source)", disabled=True),
                "rics_category": st.column_config.TextColumn("RICS category", disabled=True),
                "display_name": st.column_config.TextColumn("Label on chart"),
                "group_name": st.column_config.TextColumn("Group (optional)"),
                "contingency_pct": st.column_config.NumberColumn(
                    "Contingency (%)",
                    min_value=0.0,
                    max_value=100.0,
                    step=0.5,
                ),
                "show_on_chart": st.column_config.CheckboxColumn("Show"),
            },
            hide_index=True,
            key="subcat_editor",
        ).reset_index(drop=True)

        edited["group_name"] = edited["group_name"].fillna("").astype(str).str.strip()
        edited["display_name"] = edited["display_name"].fillna("").astype(str).str.strip()
        edited["contingency_pct"] = pd.to_numeric(edited["contingency_pct"], errors="coerce").fillna(0.0)
        edited["show_on_chart"] = edited["show_on_chart"].fillna(True).astype(bool)

        st.session_state["subcat_map_df"] = edited.copy()

    return st.session_state["subcat_map_df"].copy()

def build_chart_mappings(subcat_map_df: pd.DataFrame):
    display_map = dict(zip(subcat_map_df["subcategory_id"], subcat_map_df["display_name"]))
    group_map = {
        str(sid): g.strip()
        for sid, g in zip(subcat_map_df["subcategory_id"], subcat_map_df["group_name"])
        if isinstance(g, str) and g.strip()
    }
    contingency_map = {
        str(sid): float(pct) / 100.0
        for sid, pct in zip(subcat_map_df["subcategory_id"], subcat_map_df["contingency_pct"])
    }
    show_map = dict(zip(subcat_map_df["subcategory_id"], subcat_map_df["show_on_chart"]))
    return display_map, group_map, contingency_map, show_map

def prepare_rows_chart(rows: pd.DataFrame, show_map: dict[str, bool]) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows_all = rows.copy()
    rows_chart = rows_all.copy()
    rows_chart["__subcategory_id"] = rows_chart["rics_detail"].map(clean_rics_label).astype(str)
    rows_chart = rows_chart[rows_chart["__subcategory_id"].map(show_map).fillna(True)].copy()
    rows_chart = rows_chart.drop(columns=["__subcategory_id"], errors="ignore")

    return rows_all, rows_chart

def pick_chart_rows_for_current_view(
    chart_choice: str,
    rows: pd.DataFrame,
    rows_chart: pd.DataFrame,
) -> pd.DataFrame:
    if str(chart_choice).startswith(("1", "2", "3")):
        return rows_chart
    return rows

def resolve_y_axis_settings(
    manual_y_axis: bool,
    y_axis_min: float,
    y_axis_max: float,
    y_axis_dtick: float,
) -> tuple[float | None, float | None, float | None]:
    y_min_val = y_axis_min if manual_y_axis else None
    y_max_val = y_axis_max if manual_y_axis else None
    y_dtick_val = y_axis_dtick if (manual_y_axis and y_axis_dtick > 0) else None
    return y_min_val, y_max_val, y_dtick_val

def resolve_coverage_modules(
    chart_choice: str,
    upfront_modules: list[str],
    wlc_modules_selected: list[str],
) -> list[str]:
    if chart_choice.startswith("1"):
        return upfront_modules
    if chart_choice.startswith("2"):
        return wlc_modules_selected
    if chart_choice.startswith("3"):
        return list(dict.fromkeys(upfront_modules + wlc_modules_selected))
    return upfront_modules

def render_chart_coverage_audit(
    coverage_summary: dict,
    coverage_gaps: pd.DataFrame,
) -> None:
    with st.expander("Chart coverage audit", expanded=False):
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Value column used", coverage_summary["value_col_used"])
        c2.metric("Source total", f"{coverage_summary['total_source_scaled']:,.2f}")
        c3.metric("Plotted base total", f"{coverage_summary['plotted_base']:,.2f}")
        c4.metric("Contingency total", f"{coverage_summary['plotted_contingency']:,.2f}")
        c5.metric("Uncovered positive gap", f"{coverage_summary['uncovered_positive']:,.2f}")

        if coverage_summary["fully_allocated"]:
            st.success("All positive carbon appears to have been allocated to a chart category.")
        else:
            st.warning("Some source rows are not fully allocated to chart categories. Review the table below.")

        audit_df = coverage_gaps.copy()
        if not audit_df.empty:
            show_cols = [
                c for c in [
                    "building_name", "section", "element_name", "rics_detail",
                    "source_value", "allocated_value", "gap", "gap_abs"
                ] if c in audit_df.columns
            ]
            st.dataframe(audit_df[show_cols].head(200), width="stretch")

def render_selected_chart(
    chart_choice: str,
    rows: pd.DataFrame,
    em_col: str,
    total_gia: float,
    building_label: str,
    selected_buildings_title: str,
    colour_map: dict[str, str],
    contingency_colour: str,
    biogenic_colour: str,
    display_map: dict[str, str],
    group_map: dict[str, str],
    contingency_map: dict[str, float],
    project_contingency_pct: float,
    upfront_modules: list[str],
    wlc_modules_selected: list[str],
    chart_height: int,
    bar_width: float,
    small_segment_threshold: float,
    segment_border_px: float,
    use_intensity: bool,
    show_legend: bool,
    collapsed_high_level_labels: bool,
    y_min_val: float | None,
    y_max_val: float | None,
    y_dtick_val: float | None,
    target_line_value,
    target_line_label: str,
    target_line_colour: str,
    target_line_value_2,
    target_line_label_2: str,
    target_line_colour_2: str,
    bar_label_font_size: int,
    mat_module: str,
    limit_target_line_to_one_bar: bool = False,
    target_line_bar_scope: str = "Upfront",
    limit_target_line_2_to_one_bar: bool = False,
    target_line_2_bar_scope: str = "Whole life cycle",
):
    current_fig = None
    render_streamlit_legend = should_render_rics_streamlit_legend(show_legend)

    if chart_choice.startswith("1"):
        st.subheader(f"Upfront embodied carbon (A1–A5) — {selected_buildings_title}")

        coverage = compute_chart_coverage(
            rows=rows,
            modules=upfront_modules,
            em_col=em_col,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            use_intensity=use_intensity,
            gia_m2=total_gia,
        )[0]

        st.caption(
            f"Coverage check — source: {coverage['total_source_scaled']:,.2f}, "
            f"plotted base: {coverage['plotted_base']:,.2f}, "
            f"contingency: {coverage['plotted_contingency']:,.2f}, "
            f"uncovered positive gap: {coverage['uncovered_positive']:,.2f}"
        )

        legend_df = build_high_level_legend_df(
            rows=rows,
            modules=upfront_modules,
            em_col=em_col,
            colour_map=colour_map,
            use_intensity=use_intensity,
            gia_m2=total_gia,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
        )
        if render_streamlit_legend:
            render_high_level_legend(legend_df, use_intensity)

        current_fig = plot_rics_single_stack(
            rows=rows,
            modules=upfront_modules,
            em_col=em_col,
            colour_map=colour_map,
            height_px=chart_height,
            bar_width=bar_width,
            small_segment_threshold=small_segment_threshold,
            segment_border_px=segment_border_px,
            title=f"Upfront embodied carbon (A1–A5) by RICS category — {selected_buildings_title}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
            show_legend=show_legend,
            y_min=y_min_val,
            y_max=y_max_val,
            y_dtick=y_dtick_val,
            collapsed_high_level_labels=collapsed_high_level_labels,
            target_line_value=target_line_value,
            target_line_label=target_line_label,
            target_line_colour=target_line_colour,
            target_line_value_2=target_line_value_2,
            target_line_label_2=target_line_label_2,
            target_line_colour_2=target_line_colour_2,
            bar_label_font_size=bar_label_font_size,
        )
        st.plotly_chart(current_fig, width="stretch")
        return current_fig

    if chart_choice.startswith("2"):
        st.subheader(f"Life cycle embodied carbon (excl. B6, B7 and D) — {selected_buildings_title}")

        coverage = compute_chart_coverage(
            rows=rows,
            modules=wlc_modules_selected,
            em_col=em_col,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            use_intensity=use_intensity,
            gia_m2=total_gia,
        )[0]

        st.caption(
            f"Coverage check — source: {coverage['total_source_scaled']:,.2f}, "
            f"plotted base: {coverage['plotted_base']:,.2f}, "
            f"contingency: {coverage['plotted_contingency']:,.2f}, "
            f"uncovered positive gap: {coverage['uncovered_positive']:,.2f}"
        )

        legend_df = build_high_level_legend_df(
            rows=rows,
            modules=wlc_modules_selected,
            em_col=em_col,
            colour_map=colour_map,
            use_intensity=use_intensity,
            gia_m2=total_gia,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
        )
        if render_streamlit_legend:
            render_high_level_legend(legend_df, use_intensity)

        current_fig = plot_rics_single_stack(
            rows=rows,
            modules=wlc_modules_selected,
            em_col=em_col,
            colour_map=colour_map,
            height_px=chart_height,
            bar_width=bar_width,
            small_segment_threshold=small_segment_threshold,
            segment_border_px=segment_border_px,
            title=f"Life cycle embodied carbon (excl. B6, B7 and D) by RICS category — {selected_buildings_title}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
            show_legend=show_legend,
            y_min=y_min_val,
            y_max=y_max_val,
            y_dtick=y_dtick_val,
            collapsed_high_level_labels=collapsed_high_level_labels,
            target_line_value=target_line_value,
            target_line_label=target_line_label,
            target_line_colour=target_line_colour,
            target_line_value_2=target_line_value_2,
            target_line_label_2=target_line_label_2,
            target_line_colour_2=target_line_colour_2,
            bar_label_font_size=bar_label_font_size,
        )
        st.plotly_chart(current_fig, width="stretch")
        return current_fig

    if chart_choice.startswith("3"):
        st.subheader(f"Upfront + life cycle embodied carbon — {selected_buildings_title}")

        legend_df = build_high_level_legend_df_two_charts(
            rows=rows,
            upfront_modules=upfront_modules,
            whole_life_modules=wlc_modules_selected,
            em_col=em_col,
            colour_map=colour_map,
            use_intensity=use_intensity,
            gia_m2=total_gia,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
        )
        if render_streamlit_legend:
            render_high_level_legend_two_charts(legend_df, use_intensity)

        current_fig = plot_rics_two_stacks(
            rows=rows,
            upfront_modules=upfront_modules,
            whole_life_modules=wlc_modules_selected,
            em_col=em_col,
            colour_map=colour_map,
            height_px=chart_height,
            bar_width=bar_width,
            segment_border_px=segment_border_px,
            title=f"Upfront + life cycle embodied carbon by RICS category — {selected_buildings_title}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
            show_legend=show_legend,
            y_min=y_min_val,
            y_max=y_max_val,
            y_dtick=y_dtick_val,
            small_segment_threshold=small_segment_threshold,
            target_line_value=target_line_value,
            target_line_label=target_line_label,
            target_line_colour=target_line_colour,
            target_line_value_2=target_line_value_2,
            target_line_label_2=target_line_label_2,
            target_line_colour_2=target_line_colour_2,
            collapsed_high_level_labels=collapsed_high_level_labels,
            bar_label_font_size=bar_label_font_size,
            limit_target_line_to_one_bar=limit_target_line_to_one_bar,
            target_line_bar_scope=target_line_bar_scope,
            limit_target_line_2_to_one_bar=limit_target_line_2_to_one_bar,
            target_line_2_bar_scope=target_line_2_bar_scope,
        )
        st.plotly_chart(current_fig, width="stretch")
        return current_fig

    current_fig = render_materials_charts(
        rows=rows,
        em_col=em_col,
        mat_module=mat_module,
        building_label=building_label,
        chart_height=chart_height,
    )
    return current_fig

def render_chart_png_download(
    current_fig: go.Figure | None,
    export_width_px: int,
    export_height_px: int,
    building_label: str,
    chart_choice: str,
) -> None:
    if current_fig is None:
        return

    png_bytes = make_png_bytes(current_fig, int(export_width_px), int(export_height_px))
    if png_bytes is not None:
        st.download_button(
            "Download chart PNG",
            data=png_bytes,
            file_name=f"{re.sub(r'[^A-Za-z0-9._-]+', '_', building_label)}_{chart_choice.split(')')[0]}.png",
            mime="image/png",
        )
    else:
        st.info("PNG export is unavailable in this environment. Install kaleido in the app environment to enable it.")

def render_processed_table_downloads(
    rows: pd.DataFrame,
    entries: pd.DataFrame | None,
    recon: pd.DataFrame,
    recon_summary: pd.DataFrame,
    em_col: str,
    display_map: dict[str, str],
    group_map: dict[str, str],
) -> pd.DataFrame:
    st.subheader("Download processed tables")

    st.download_button(
        "Download tidy rows CSV",
        data=rows.to_csv(index=False).encode("utf-8"),
        file_name="oneclick_detail_tidy_rows.csv",
        mime="text/csv",
    )

    if entries is not None and len(entries):
        st.download_button(
            "Download reconstructed material entries CSV",
            data=entries.to_csv(index=False).encode("utf-8"),
            file_name="oneclick_material_entries.csv",
            mime="text/csv",
        )

    st.download_button(
        "Download reconciliation table CSV",
        data=recon.to_csv(index=False).encode("utf-8"),
        file_name="oneclick_reconciliation_rows.csv",
        mime="text/csv",
    )

    st.download_button(
        "Download reconciliation summary CSV",
        data=recon_summary.to_csv(index=False).encode("utf-8"),
        file_name="oneclick_reconciliation_summary.csv",
        mime="text/csv",
    )

    stack_export = build_stack_by_module_export(
        rows=rows,
        em_col=em_col,
        display_map=display_map,
        group_map=group_map,
    )

    st.download_button(
        "Download current stacks by module CSV",
        data=stack_export.to_csv(index=False).encode("utf-8"),
        file_name="oneclick_current_stacks_by_module.csv",
        mime="text/csv",
    )

    return stack_export

def render_gla_export_settings() -> tuple[str, float, float]:
    with st.expander("GLA export settings", expanded=False):
        gla_scenario_label = st.text_input(
            "GLA scenario label",
            value="current_grid",
            help="Use a short label such as current_grid or decarbonised_grid.",
        )
        gla_min_group_mass_kg = st.number_input(
            "GLA Table 2 minimum grouped mass (kg)",
            min_value=0.0,
            value=100.0,
            step=50.0,
        )
        gla_min_group_share_pct = st.number_input(
            "GLA Table 2 minimum grouped share within category (%)",
            min_value=0.0,
            value=1.0,
            step=0.5,
        )

    return gla_scenario_label, float(gla_min_group_mass_kg), float(gla_min_group_share_pct)

def build_gla_export_artifacts(
    rows: pd.DataFrame,
    entries: pd.DataFrame | None,
    building_label: str,
    gla_scenario_label: str,
    gla_min_group_mass_kg: float,
    gla_min_group_share_pct: float,
) -> dict:
    gla_table1 = build_gla_table1_export(rows)
    gla_table1_unmapped = getattr(build_gla_table1_export, "last_unmapped", pd.DataFrame()).copy()
    gla_table2 = build_gla_table2_export(
        entries=entries,
        rows=rows,
        min_group_mass_kg=float(gla_min_group_mass_kg),
        min_group_share_pct=float(gla_min_group_share_pct),
    )
    gla_table2_unmapped = build_gla_table2_unmapped_debug(entries=entries, rows=rows)
    gla_workbook = build_gla_export_workbook_bytes(gla_table1, gla_table2)

    gla_file_stub = (
        re.sub(r"[^A-Za-z0-9._-]+", "_", building_label)
        + "_"
        + re.sub(r"[^A-Za-z0-9._-]+", "_", str(gla_scenario_label).strip() or "scenario")
    )

    return {
        "table1": gla_table1,
        "table1_unmapped": gla_table1_unmapped,
        "table2": gla_table2,
        "table2_unmapped": gla_table2_unmapped,
        "workbook": gla_workbook,
        "file_stub": gla_file_stub,
    }

def render_gla_export_downloads(gla_artifacts: dict) -> None:
    gla_table1 = gla_artifacts["table1"]
    gla_table1_unmapped = gla_artifacts["table1_unmapped"]
    gla_table2 = gla_artifacts["table2"]
    gla_table2_unmapped = gla_artifacts["table2_unmapped"]
    gla_workbook = gla_artifacts["workbook"]
    gla_file_stub = gla_artifacts["file_stub"]

    st.download_button(
        "Download GLA Table 1 CSV",
        data=gla_table1.to_csv(index=False).encode("utf-8"),
        file_name=f"{gla_file_stub}_gla_table_1.csv",
        mime="text/csv",
    )

    st.download_button(
        "Download GLA Table 2 CSV",
        data=gla_table2.to_csv(index=False).encode("utf-8"),
        file_name=f"{gla_file_stub}_gla_table_2.csv",
        mime="text/csv",
    )

    st.download_button(
        "Download GLA export workbook (.xlsx)",
        data=gla_workbook,
        file_name=f"{gla_file_stub}_gla_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if gla_table1_unmapped is not None and len(gla_table1_unmapped):
        st.warning(
            f"GLA Table 1 excludes {len(gla_table1_unmapped)} rows that do not map to a recognised RICS/NRM category."
        )

        show_cols = [
            c for c in [
                "source_file",
                "building_name",
                "section",
                "Resource",
                "Question",
                "Comment",
                "rics_alloc_label",
                "rics_level2_label",
                "rics_high_label",
                "rics_detail_base",
            ]
            if c in gla_table1_unmapped.columns
        ]
        st.dataframe(gla_table1_unmapped[show_cols], width="stretch")

        st.download_button(
            "Download GLA Table 1 unmapped rows CSV",
            data=gla_table1_unmapped.to_csv(index=False).encode("utf-8"),
            file_name=f"{gla_file_stub}_gla_table_1_unmapped_rows.csv",
            mime="text/csv",
        )

    if gla_table2_unmapped is not None and len(gla_table2_unmapped):
        st.warning(
            f"GLA Table 2 has {len(gla_table2_unmapped)} unmapped source rows. Download the debug CSV to trace the source file(s)."
        )
        st.download_button(
            "Download GLA Table 2 unmapped rows CSV",
            data=gla_table2_unmapped.to_csv(index=False).encode("utf-8"),
            file_name=f"{gla_file_stub}_gla_table_2_unmapped_rows.csv",
            mime="text/csv",
        )

def render_element_export_download(rows: pd.DataFrame, em_col: str) -> None:
    elem_key = [
        c for c in ["building_name", "element_name", "rics_detail", "rics_high_code", "rics_high_label"]
        if c in rows.columns
    ]
    if not (elem_key and "section" in rows.columns):
        return

    value_col_for_export = pick_chart_value_col(rows, em_col)
    elem = (
        rows.groupby(elem_key + ["section"], dropna=False)[value_col_for_export]
        .sum()
        .reset_index()
        .pivot_table(
            index=elem_key,
            columns="section",
            values=value_col_for_export,
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )
    elem.columns = [str(c) for c in elem.columns]

    st.download_button(
        "Download element-by-module CSV",
        data=elem.to_csv(index=False).encode("utf-8"),
        file_name="oneclick_detail_element_by_module.csv",
        mime="text/csv",
    )

def run_main_app(
    uploaded_files,
    manual_file,
    sidebar_controls: dict,
) -> None:
    if uploaded_files:
        building_meta_df = sync_building_meta(uploaded_files).copy()
        building_meta_df = render_buildings_editor(building_meta_df)
        selected_buildings, collapsed_high_level_labels = render_buildings_sidebar_controls(building_meta_df)

        rows, entries, meta_by_building, em_col, selected_meta = load_project_rows_or_stop(
            uploaded_files_=uploaded_files,
            building_meta_df=building_meta_df,
            selected_buildings=selected_buildings,
            manual_file_obj=manual_file,
        )

        total_gia, building_label, selected_buildings_title = build_project_summary(
            selected_buildings=selected_buildings,
            selected_meta=selected_meta,
        )

        recon = build_reconciliation_table(rows, em_col)
        recon_summary = summarise_reconciliation(recon, tolerance_kg=sidebar_controls["recon_tolerance_kg"])
        recon_flagged = build_unassigned_rows(rows, em_col, tolerance_kg=sidebar_controls["recon_tolerance_kg"])

        render_project_metadata(meta_by_building)

        project_contingency_pct = render_reconciliation_and_contingency(
            selected_meta=selected_meta,
            total_gia=total_gia,
            recon_summary=recon_summary,
            recon_flagged=recon_flagged,
        )

        upfront_modules, wlc_modules_selected = build_module_selections(rows)
        colour_map, contingency_colour, biogenic_colour = render_colour_controls_sidebar()
        subcat_map_df = render_chart_items_editor(rows)
        display_map, group_map, contingency_map, show_map = build_chart_mappings(subcat_map_df)
        _, rows_chart = prepare_rows_chart(rows, show_map)

        chart_rows = pick_chart_rows_for_current_view(
            chart_choice=sidebar_controls["chart_choice"],
            rows=rows,
            rows_chart=rows_chart,
        )

        y_min_val, y_max_val, y_dtick_val = resolve_y_axis_settings(
            manual_y_axis=sidebar_controls["manual_y_axis"],
            y_axis_min=sidebar_controls["y_axis_min"],
            y_axis_max=sidebar_controls["y_axis_max"],
            y_axis_dtick=sidebar_controls["y_axis_dtick"],
        )

        coverage_modules = resolve_coverage_modules(
            chart_choice=sidebar_controls["chart_choice"],
            upfront_modules=upfront_modules,
            wlc_modules_selected=wlc_modules_selected,
        )

        coverage_summary, coverage_gaps = compute_chart_coverage(
            rows=chart_rows,
            modules=coverage_modules,
            em_col=em_col,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            use_intensity=sidebar_controls["use_intensity"],
            gia_m2=total_gia,
        )
        render_chart_coverage_audit(coverage_summary, coverage_gaps)

        current_fig = render_selected_chart(
            chart_choice=sidebar_controls["chart_choice"],
            rows=chart_rows,
            em_col=em_col,
            total_gia=total_gia,
            building_label=building_label,
            selected_buildings_title=selected_buildings_title,
            colour_map=colour_map,
            contingency_colour=contingency_colour,
            biogenic_colour=biogenic_colour,
            display_map=display_map,
            group_map=group_map,
            contingency_map=contingency_map,
            project_contingency_pct=project_contingency_pct,
            upfront_modules=upfront_modules,
            wlc_modules_selected=wlc_modules_selected,
            chart_height=sidebar_controls["chart_height"],
            bar_width=sidebar_controls["bar_width"],
            small_segment_threshold=sidebar_controls["small_segment_threshold"],
            segment_border_px=sidebar_controls["segment_border_px"],
            use_intensity=sidebar_controls["use_intensity"],
            show_legend=sidebar_controls["show_legend"],
            collapsed_high_level_labels=collapsed_high_level_labels,
            y_min_val=y_min_val,
            y_max_val=y_max_val,
            y_dtick_val=y_dtick_val,
            target_line_value=sidebar_controls["target_line_value"] if sidebar_controls["show_target_line"] else None,
            target_line_label=sidebar_controls["target_line_label"],
            target_line_colour=sidebar_controls["target_line_colour"],
            target_line_value_2=sidebar_controls["target_line_value_2"] if sidebar_controls["show_target_line_2"] else None,
            target_line_label_2=sidebar_controls["target_line_label_2"],
            target_line_colour_2=sidebar_controls["target_line_colour_2"],
            bar_label_font_size=sidebar_controls["bar_label_font_size"],
            mat_module=sidebar_controls["mat_module"],
            limit_target_line_to_one_bar=sidebar_controls["limit_target_line_to_one_bar"],
            target_line_bar_scope=sidebar_controls["target_line_bar_scope"],
            limit_target_line_2_to_one_bar=sidebar_controls["limit_target_line_2_to_one_bar"],
            target_line_2_bar_scope=sidebar_controls["target_line_2_bar_scope"],
        )
        
        render_chart_png_download(
            current_fig=current_fig,
            export_width_px=sidebar_controls["export_width_px"],
            export_height_px=sidebar_controls["export_height_px"],
            building_label=building_label,
            chart_choice=sidebar_controls["chart_choice"],
        )

        stack_export = render_processed_table_downloads(
            rows=rows,
            entries=entries,
            recon=recon,
            recon_summary=recon_summary,
            em_col=em_col,
            display_map=display_map,
            group_map=group_map,
        )

        gla_scenario_label, gla_min_group_mass_kg, gla_min_group_share_pct = render_gla_export_settings()

        gla_artifacts = build_gla_export_artifacts(
            rows=rows,
            entries=entries,
            building_label=building_label,
            gla_scenario_label=gla_scenario_label,
            gla_min_group_mass_kg=gla_min_group_mass_kg,
            gla_min_group_share_pct=gla_min_group_share_pct,
        )
        render_gla_export_downloads(gla_artifacts)

        render_element_export_download(rows=rows, em_col=em_col)

    else:
        # Empty-state message shown before any OneClick files have been uploaded.
        st.info("Upload one or more detailReport exports to begin.")

uploaded_files, manual_file = render_page_shell_and_uploads()
sidebar_controls = render_primary_sidebar_controls()
run_main_app(
    uploaded_files=uploaded_files,
    manual_file=manual_file,
    sidebar_controls=sidebar_controls,
)