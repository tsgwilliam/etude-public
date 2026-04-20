from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

import pandas as pd
import plotly.express as px


@dataclass
class DetailReport:
    rows: pd.DataFrame
    meta: dict
    entries: pd.DataFrame | None = None

GLA_STAGE_ORDER = [
    "A1-A3", "A4", "A5",
    "B1", "B2", "B3", "B4", "B5", "B6", "B7",
    "C1", "C2", "C3", "C4", "D",
]

# -------------------------------------------------
# Generic helpers
# -------------------------------------------------

def _collapse_rics_alloc_label(text_value: str) -> str:
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

def _read_raw_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return pd.read_excel(path, header=None, engine="xlrd")
    if suffix == ".xlsx":
        return pd.read_excel(path, header=None, engine="openpyxl")
    if suffix == ".csv":
        return pd.read_csv(path, header=None)
    raise ValueError(f"Unsupported file type: {suffix}")


def _read_main_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return pd.read_excel(path, header=2, engine="xlrd")
    if suffix == ".xlsx":
        return pd.read_excel(path, header=2, engine="openpyxl")
    if suffix == ".csv":
        raw = pd.read_csv(path, header=None)
        header_row = 0
        for i in range(min(10, len(raw))):
            first = str(raw.iloc[i, 0]).strip().lower()
            if first == "section":
                header_row = i
                break
        header = raw.iloc[header_row].tolist()
        df = raw.iloc[header_row + 1 :].copy()
        df.columns = header
        df = df.reset_index(drop=True)
        return df
    raise ValueError(f"Unsupported file type: {suffix}")


def _pick_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    return None


def _clean_str_series(s: pd.Series) -> pd.Series:
    return s.astype("string").fillna("").str.strip()


def _coerce_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _normalise_header(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _extract_meta(raw: pd.DataFrame, is_csv: bool) -> dict:
    if is_csv:
        return {}
    meta: dict[str, str] = {}
    try:
        for r in range(min(2, len(raw))):
            vals = raw.iloc[r].dropna().tolist()
            if len(vals) >= 2:
                key = str(vals[0]).strip()
                val = str(vals[1]).strip()
                if key:
                    meta[key] = val
    except Exception:
        pass
    return meta


def _standardise_base_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    rename_candidates = {
        "section": ["section", "Section"],
        "Resource": ["Resource", "resource"],
        "Comment": ["Comment", "comment"],
        "Material": ["Material", "material"],
        "Question": ["Question", "question", "Element", "element", "Category", "category"],
        "Construction": ["Construction", "construction"],
        "User input": ["User input", "user input", "Quantity", "quantity"],
        "User input unit": ["User input unit", "Unit", "unit", "user input unit"],
    }
    for target, candidates in rename_candidates.items():
        found = _pick_column(out, candidates)
        if found and found != target:
            out = out.rename(columns={found: target})

    if "section" not in out.columns:
        out["section"] = ""

    return out


def _find_rics_allocation_columns(df: pd.DataFrame) -> list[str]:
    cols: list[str] = []
    for c in df.columns:
        cs = str(c).strip()
        if re.match(r"^\d+(?:\.\d+)*\s+\S+", cs):
            cols.append(c)
            continue
        if cs.lower() in {"unclassified / other", "unclassified/other", "other"}:
            cols.append(c)
    return cols

def _find_emissions_column(df: pd.DataFrame) -> str:
    preferred = [
        "kgco2e",
        "TOTAL kg CO₂e",
        "TOTAL kg CO2e",
        "Total kg CO₂e",
        "Total kg CO2e",
        "gwp_kgco2e",
        "gwp",
        "GWP total",
        "GWP [kgCO2e]",
        "Global warming potential",
    ]
    for c in preferred:
        if c in df.columns:
            return c
    for c in df.columns:
        cl = str(c).lower()
        if "biogenic" in cl:
            continue
        if any(k in cl for k in ["total kg co", "gwp", "co2e", "kgco2", "kg co2", "carbon"]):
            return c
    numeric = df.select_dtypes("number").columns.tolist()
    if numeric:
        return numeric[0]
    raise KeyError("No emissions column found.")


def _find_mass_column(df: pd.DataFrame) -> str | None:
    preferred = [
        "Mass of raw materials kg",
        "Mass of raw material kg",
        "Mass of raw materials",
        "Mass of materials",
        "Mass of a resource",
        "Mass",
        "mass",
    ]
    for c in preferred:
        if c in df.columns:
            return c
    for c in df.columns:
        cl = str(c).lower()
        if "mass" in cl or "weight" in cl:
            return c
    return None


def _find_first_matching_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for cand in candidates:
        if cand in df.columns:
            return cand
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        found = lower_map.get(str(cand).strip().lower())
        if found is not None:
            return found
    return None


def _extract_optional_numeric_column(df: pd.DataFrame, candidates: list[str]) -> pd.Series:
    col = _find_first_matching_column(df, candidates)
    if col is None:
        return pd.Series(0.0, index=df.index, dtype="float64")
    return _coerce_numeric(df[col])


def _extract_optional_text_column(df: pd.DataFrame, candidates: list[str]) -> pd.Series:
    col = _find_first_matching_column(df, candidates)
    if col is None:
        return pd.Series("", index=df.index, dtype="string")
    return _clean_str_series(df[col])


def _derive_element_name(df: pd.DataFrame) -> pd.Series:
    question = _clean_str_series(df["Question"]) if "Question" in df.columns else pd.Series("", index=df.index, dtype="string")
    comment = _clean_str_series(df["Comment"]) if "Comment" in df.columns else pd.Series("", index=df.index, dtype="string")
    resource = _clean_str_series(df["Resource"]) if "Resource" in df.columns else pd.Series("", index=df.index, dtype="string")
    return question.where(question != "", comment.where(comment != "", resource))


def _derive_material_label(df: pd.DataFrame) -> pd.Series:
    material = _clean_str_series(df["Material"]) if "Material" in df.columns else pd.Series("", index=df.index, dtype="string")
    resource = _clean_str_series(df["Resource"]) if "Resource" in df.columns else pd.Series("", index=df.index, dtype="string")
    construction = _clean_str_series(df["Construction"]) if "Construction" in df.columns else pd.Series("", index=df.index, dtype="string")
    comment = _clean_str_series(df["Comment"]) if "Comment" in df.columns else pd.Series("", index=df.index, dtype="string")

    label = material.where(material != "", resource)
    extra = construction.where(construction != "", comment)
    return label.where(extra == "", label + " — " + extra)


def _derive_rics_detail_base(df: pd.DataFrame) -> pd.Series:
    rics_col = _pick_column(
        df,
        ["RICS category", "RICS Category", "rics category", "NRM category", "NRM Category"],
    )
    if rics_col:
        return _clean_str_series(df[rics_col])
    return pd.Series("", index=df.index, dtype="string")


def _derive_rics_high_from_alloc(alloc_label: pd.Series) -> pd.Series:
    alloc = _clean_str_series(alloc_label)
    top_num = alloc.str.extract(r"^(\d+)", expand=False)
    top_map = {
        "0": "0 Deconstruction",
        "1": "1 Substructure",
        "2": "2 Superstructure",
        "3": "3 Finishes",
        "4": "4 Fittings, furnishings and equipment",
        "5": "5 Services",
        "6": "6 Prefabricated buildings and building units",
        "7": "7 Work to existing buildings",
        "8": "8 External works",
        "9": "9 Main contractor preliminaries",
    }
    return top_num.map(top_map).fillna(alloc)


def _derive_rics_level2_from_alloc(alloc_label: pd.Series, high_label: pd.Series) -> pd.Series:
    alloc = _clean_str_series(alloc_label)
    high = _clean_str_series(high_label)

    full_code = alloc.str.extract(r"^(\d+(?:\.\d+)*)", expand=False).fillna("")
    desc = alloc.str.replace(r"^\d+(?:\.\d+)*\s*", "", regex=True)

    lvl2_code = full_code.str.extract(r"^(\d+\.\d+)", expand=False).fillna("")
    out = lvl2_code.where(lvl2_code == "", lvl2_code + " " + desc)

    return out.where(out != "", high)

def _explode_rics_allocations(df: pd.DataFrame) -> pd.DataFrame:
    alloc_cols = _find_rics_allocation_columns(df)
    if not alloc_cols:
        out = df.copy()
        out["rics_alloc_label"] = _clean_str_series(out["rics_detail_base"]) if "rics_detail_base" in out.columns else ""
        out["rics_allocated_value"] = _coerce_numeric(out["kgco2e"]) if "kgco2e" in out.columns else 0.0
        return out

    parts: list[pd.DataFrame] = []
    allocated_any = pd.Series(False, index=df.index)

    for col in alloc_cols:
        vals = _coerce_numeric(df[col])
        mask = vals != 0
        allocated_any = allocated_any | mask
        if not mask.any():
            continue
        part = df.loc[mask].copy()
        part["rics_alloc_label"] = str(col).strip()
        part["rics_allocated_value"] = vals.loc[mask].values
        parts.append(part)

    unallocated_mask = ~allocated_any
    if unallocated_mask.any():
        part = df.loc[unallocated_mask].copy()
        part["rics_alloc_label"] = _clean_str_series(part["rics_detail_base"]) if "rics_detail_base" in part.columns else ""
        part["rics_allocated_value"] = _coerce_numeric(part["kgco2e"]) if "kgco2e" in part.columns else 0.0
        parts.append(part)

    if not parts:
        out = df.copy()
        out["rics_alloc_label"] = _clean_str_series(out["rics_detail_base"]) if "rics_detail_base" in out.columns else ""
        out["rics_allocated_value"] = _coerce_numeric(out["kgco2e"]) if "kgco2e" in out.columns else 0.0
        return out

    return pd.concat(parts, ignore_index=True)

def _drop_total_like_rows(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    # 1) Drop explicit total/subtotal rows by section label
    sec = _clean_str_series(out["section"]).str.lower()
    out = out[~sec.isin({"total", "subtotal", "sum"})].copy()

    # 2) Drop subtotal-like rows where all descriptor fields are blank
    #    but numeric totals are present
    descriptor_cols = [
        c for c in [
            "Resource",
            "User input",
            "User input unit",
            "Question",
            "Construction",
            "Comment",
            "Material",
        ]
        if c in out.columns
    ]

    if descriptor_cols:
        desc = out[descriptor_cols].copy()

        for c in descriptor_cols:
            desc[c] = _clean_str_series(desc[c])
            desc[c] = desc[c].replace("", pd.NA)

        all_descriptor_blank = desc.isna().all(axis=1)

        numeric_cols = out.select_dtypes("number").columns.tolist()
        if numeric_cols:
            has_nonzero_numeric = out[numeric_cols].fillna(0).abs().sum(axis=1) > 0
            out = out[~(all_descriptor_blank & has_nonzero_numeric)].copy()

    return out

def _build_material_entries(base_df: pd.DataFrame) -> pd.DataFrame:
    working = _drop_total_like_rows(base_df).copy()
    if working.empty:
        return pd.DataFrame()

    mass_col = "mass_value" if "mass_value" in working.columns else _find_mass_column(working)

    descriptor_cols = [c for c in [
        "Resource",
        "User input",
        "User input unit",
        "Question",
        "Construction",
        "Comment",
        "Material",
    ] if c in working.columns]

    for c in descriptor_cols:
        if pd.api.types.is_numeric_dtype(working[c]):
            working[c] = pd.to_numeric(working[c], errors="coerce")
        else:
            working[c] = _clean_str_series(working[c]).replace("", pd.NA)

    if "eol_process" not in working.columns:
        working["eol_process"] = pd.Series("", index=working.index, dtype="string")
    else:
        working["eol_process"] = _clean_str_series(working["eol_process"])

    if "reusable_kg" not in working.columns:
        working["reusable_kg"] = pd.Series(0.0, index=working.index, dtype="float64")
    else:
        working["reusable_kg"] = _coerce_numeric(working["reusable_kg"])

    if "recyclable_kg" not in working.columns:
        working["recyclable_kg"] = pd.Series(0.0, index=working.index, dtype="float64")
    else:
        working["recyclable_kg"] = _coerce_numeric(working["recyclable_kg"])

    working["section_norm"] = _clean_str_series(working["section"])
    working["_section_occurrence"] = working.groupby(descriptor_cols + ["section_norm"], dropna=False).cumcount()

    key_cols = descriptor_cols + ["_section_occurrence"]
    material_id_df = working[key_cols].drop_duplicates().reset_index(drop=True)
    material_id_df["_material_entry_id"] = [f"ME{i+1:06d}" for i in range(len(material_id_df))]
    working = working.merge(material_id_df, on=key_cols, how="left")

    stage_wide = (
        working.pivot_table(
            index="_material_entry_id",
            columns="section_norm",
            values="kgco2e",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    first_cols = [c for c in [
        "Resource",
        "Material",
        "material_label",
        "Question",
        "Construction",
        "Comment",
        "element_name",
        "User input",
        "User input unit",
        "rics_detail_base",
        "eol_process",
    ] if c in working.columns]

    entry_desc = (
        working.sort_values(["_material_entry_id", "_source_row_id"])
        .groupby("_material_entry_id", dropna=False)[first_cols]
        .first()
        .reset_index()
    )

    if mass_col and mass_col in working.columns:
        mass_pref = working.copy()
        sec_rank = mass_pref["section_norm"].map({"A1-A3": 0, "A4": 1, "A5": 2}).fillna(9)
        mass_pref["_mass_rank"] = sec_rank
        mass_num = pd.to_numeric(mass_pref[mass_col], errors="coerce")
        mass_pref = mass_pref[mass_num.notna() & (mass_num != 0)].copy()
        if not mass_pref.empty:
            mass_pref = (
                mass_pref.sort_values(["_material_entry_id", "_mass_rank", "_source_row_id"])
                .groupby("_material_entry_id", dropna=False)[mass_col]
                .first()
                .reset_index()
                .rename(columns={mass_col: "mass_value"})
            )
            entry_desc = entry_desc.merge(mass_pref, on="_material_entry_id", how="left")

    for source_col in ["reusable_kg", "recyclable_kg"]:
        if source_col in working.columns:
            qty_df = (
                working.groupby("_material_entry_id", dropna=False)[source_col]
                .max()
                .reset_index()
            )
            entry_desc = entry_desc.merge(qty_df, on="_material_entry_id", how="left")

    entry_desc = entry_desc.merge(stage_wide, on="_material_entry_id", how="left")

    module_pattern = re.compile(r"^[A-D]\d?(?:-[A-D]\d?)?$|^A1-A3$|^A4$|^A5$|^B\d$|^C\d$|^D$|^bioC$", re.IGNORECASE)

    stage_cols = [c for c in entry_desc.columns if module_pattern.match(str(c))]
    for c in stage_cols:
        entry_desc[c] = pd.to_numeric(entry_desc[c], errors="coerce").fillna(0.0)

    entry_desc["total_excl_bioc"] = entry_desc[[c for c in stage_cols if str(c).lower() != "bioc"]].sum(axis=1)
    for c in stage_cols:
        entry_desc[c] = pd.to_numeric(entry_desc[c], errors="coerce").fillna(0.0)

    if "eol_process" in entry_desc.columns:
        entry_desc["eol_process"] = _clean_str_series(entry_desc["eol_process"])
    if "reusable_kg" in entry_desc.columns:
        entry_desc["reusable_kg"] = _coerce_numeric(entry_desc["reusable_kg"])
    if "recyclable_kg" in entry_desc.columns:
        entry_desc["recyclable_kg"] = _coerce_numeric(entry_desc["recyclable_kg"])

    entry_desc["total_excl_bioc"] = entry_desc[[c for c in stage_cols if str(c).lower() != "bioc"]].sum(axis=1)
    return entry_desc.sort_values("_material_entry_id").reset_index(drop=True)


# -------------------------------------------------
# Loader
# -------------------------------------------------

def load_detail_report(path: Path) -> DetailReport:
    raw = _read_raw_table(path)
    meta = _extract_meta(raw, is_csv=(path.suffix.lower() == ".csv"))

    df = _read_main_table(path)
    df = _normalise_header(df)
    df = _standardise_base_columns(df)

    df["_source_row_id"] = range(len(df))
    df["element_name"] = _derive_element_name(df)
    df["material_label"] = _derive_material_label(df)
    df["rics_detail_base"] = _derive_rics_detail_base(df)

    em_col = _find_emissions_column(df)
    if em_col != "kgco2e":
        df["kgco2e"] = _coerce_numeric(df[em_col])
    else:
        df["kgco2e"] = _coerce_numeric(df["kgco2e"])

    mass_col = _find_mass_column(df)
    if mass_col:
        df["mass_value"] = _coerce_numeric(df[mass_col])

    df["eol_process"] = _extract_optional_text_column(
        df,
        ["EOL Process", "EoL Process", "EOL process", "End of life process"],
    )
    df["reusable_kg"] = _extract_optional_numeric_column(
        df,
        ["Estimated reusable materials kg", "Reusable materials kg", "Estimated reusable kg"],
    )
    df["recyclable_kg"] = _extract_optional_numeric_column(
        df,
        ["Estimated recyclable materials kg", "Recyclable materials kg", "Estimated recyclable kg"],
    )

    entries = _build_material_entries(df)

    if not entries.empty:
        working = _drop_total_like_rows(df).copy()

        descriptor_cols = [c for c in [
            "Resource",
            "User input",
            "User input unit",
            "Question",
            "Construction",
            "Comment",
            "Material",
        ] if c in working.columns]

        for c in descriptor_cols:
            if pd.api.types.is_numeric_dtype(working[c]):
                working[c] = pd.to_numeric(working[c], errors="coerce")
            else:
                working[c] = _clean_str_series(working[c]).replace("", pd.NA)

        if "eol_process" not in working.columns:
            working["eol_process"] = pd.Series("", index=working.index, dtype="string")
        else:
            working["eol_process"] = _clean_str_series(working["eol_process"])

        if "reusable_kg" not in working.columns:
            working["reusable_kg"] = pd.Series(0.0, index=working.index, dtype="float64")
        else:
            working["reusable_kg"] = _coerce_numeric(working["reusable_kg"])

        if "recyclable_kg" not in working.columns:
            working["recyclable_kg"] = pd.Series(0.0, index=working.index, dtype="float64")
        else:
            working["recyclable_kg"] = _coerce_numeric(working["recyclable_kg"])

        working["section_norm"] = _clean_str_series(working["section"])
        working["_section_occurrence"] = working.groupby(descriptor_cols + ["section_norm"], dropna=False).cumcount()
        key_cols = descriptor_cols + ["_section_occurrence"]

        material_id_df = working[key_cols].drop_duplicates().reset_index(drop=True)
        material_id_df["_material_entry_id"] = [f"ME{i+1:06d}" for i in range(len(material_id_df))]
        working = working.merge(material_id_df, on=key_cols, how="left")

        cols_to_merge = ["_source_row_id", "_material_entry_id"]
        for extra_col in ["eol_process", "reusable_kg", "recyclable_kg"]:
            if extra_col in working.columns:
                cols_to_merge.append(extra_col)

        df = df.merge(
            working[cols_to_merge].drop_duplicates(subset=["_source_row_id"]),
            on="_source_row_id",
            how="left",
            suffixes=("", "_from_working"),
        )

        for extra_col in ["eol_process", "reusable_kg", "recyclable_kg"]:
            from_working = f"{extra_col}_from_working"
            if from_working in df.columns:
                if extra_col in {"reusable_kg", "recyclable_kg"}:
                    existing = _coerce_numeric(df[extra_col]) if extra_col in df.columns else pd.Series(0.0, index=df.index, dtype="float64")
                    fallback = _coerce_numeric(df[from_working])
                    df[extra_col] = existing.where(existing != 0, fallback)
                else:
                    existing = _clean_str_series(df[extra_col]) if extra_col in df.columns else pd.Series("", index=df.index, dtype="string")
                    fallback = _clean_str_series(df[from_working])
                    df[extra_col] = existing.where(existing != "", fallback)
                df = df.drop(columns=[from_working])

    tidy = _explode_rics_allocations(df)
    tidy["rics_detail"] = _clean_str_series(tidy["rics_detail_base"])
    tidy["rics_high_label"] = _derive_rics_high_from_alloc(tidy["rics_alloc_label"])
    tidy["rics_high_code"] = _clean_str_series(tidy["rics_high_label"]).str.extract(r"^(\d+)", expand=False).fillna("")
    tidy["rics_level2_label"] = _derive_rics_level2_from_alloc(tidy["rics_alloc_label"], tidy["rics_high_label"])

    tidy["kgco2e"] = _coerce_numeric(tidy["kgco2e"])
    if "mass_value" in tidy.columns:
        tidy["mass_value"] = _coerce_numeric(tidy["mass_value"])
    if "reusable_kg" in tidy.columns:
        tidy["reusable_kg"] = _coerce_numeric(tidy["reusable_kg"])
    else:
        tidy["reusable_kg"] = pd.Series(0.0, index=tidy.index, dtype="float64")
    if "recyclable_kg" in tidy.columns:
        tidy["recyclable_kg"] = _coerce_numeric(tidy["recyclable_kg"])
    else:
        tidy["recyclable_kg"] = pd.Series(0.0, index=tidy.index, dtype="float64")
    if "eol_process" in tidy.columns:
        tidy["eol_process"] = _clean_str_series(tidy["eol_process"])
    else:
        tidy["eol_process"] = pd.Series("", index=tidy.index, dtype="string")

    raw_stage_total = df.loc[df["section"].isin(GLA_STAGE_ORDER + ["bioC"]), "kgco2e"].sum()

    tidy_stage_total_first = (
        tidy.loc[tidy["section"].isin(GLA_STAGE_ORDER + ["bioC"])]
        .groupby("_source_row_id", dropna=False)["kgco2e"]
        .first()
        .sum()
    )

    tidy_alloc_total = (
        tidy.loc[tidy["section"].isin(GLA_STAGE_ORDER + ["bioC"]), "rics_allocated_value"]
        .sum()
    )

    print("raw_stage_total", raw_stage_total)
    print("tidy_stage_total_first", tidy_stage_total_first)
    print("tidy_alloc_total", tidy_alloc_total)

    return DetailReport(rows=tidy, meta=meta, entries=entries)


# -------------------------------------------------
# Material breakdown
# -------------------------------------------------

def material_breakdown(rows: pd.DataFrame, module: str):
    rr = rows.copy()

    if "_source_row_id" in rr.columns:
        rr = (
            rr.sort_values("_source_row_id")
            .groupby("_source_row_id", dropna=False)
            .first()
            .reset_index()
        )

    rr = _drop_total_like_rows(rr)

    if module == "A1-A5":
        mods = {"A1-A3", "A4", "A5"}
    elif module:
        mods = {module}
    else:
        mods = set(rr["section"].dropna().unique())

    rr_mod = rr[rr["section"].isin(mods)].copy()

    if "_material_entry_id" in rr_mod.columns and rr_mod["_material_entry_id"].notna().any():
        if "material_label" in rr_mod.columns:
            label_series = rr_mod["material_label"]
        elif "Material" in rr_mod.columns:
            label_series = rr_mod["Material"]
        else:
            label_series = rr_mod.get("Resource", pd.Series("(unknown)", index=rr_mod.index))
        rr_mod["_material_label"] = _clean_str_series(label_series)

        em_df = (
            rr_mod.groupby(["_material_entry_id", "_material_label"], dropna=False)["kgco2e"]
            .sum()
            .reset_index()
            .rename(columns={"_material_label": "material", "kgco2e": "value"})
        )
        em_df = em_df.groupby("material", dropna=False)["value"].sum().reset_index()

        mass_col = "mass_value" if "mass_value" in rr.columns else _find_mass_column(rr)
        mass_df = None
        if mass_col:
            mass_base = (
                rr_mod.sort_values(["_material_entry_id", "_source_row_id"])
                .groupby("_material_entry_id", dropna=False)[["_material_label", mass_col]]
                .first()
                .reset_index()
                .rename(columns={"_material_label": "material", mass_col: "value"})
            )
            mass_df = mass_base.groupby("material", dropna=False)["value"].sum().reset_index()
        return em_df, mass_df, mass_col

    if "Material" in rr_mod.columns:
        material_col = "Material"
    elif "Resource" in rr_mod.columns:
        material_col = "Resource"
    elif "element_name" in rr_mod.columns:
        material_col = "element_name"
    else:
        material_col = rr_mod.columns[0]

    em_df = (
        rr_mod.groupby(material_col, dropna=False)["kgco2e"]
        .sum()
        .reset_index()
        .rename(columns={material_col: "material", "kgco2e": "value"})
    )

    mass_col = "mass_value" if "mass_value" in rr_mod.columns else _find_mass_column(rr_mod)
    mass_df = None
    if mass_col:
        mass_df = (
            rr_mod.groupby(material_col, dropna=False)[mass_col]
            .sum()
            .reset_index()
            .rename(columns={material_col: "material", mass_col: "value"})
        )
    return em_df, mass_df, mass_col


# -------------------------------------------------
# Plotting
# -------------------------------------------------

def plot_material_treemap_emissions(df: pd.DataFrame, title: str):
    fig = px.treemap(df, path=["material"], values="value", title=title)
    fig.update_layout(margin=dict(t=50, l=10, r=10, b=10))
    return fig


def plot_material_treemap_mass(df: pd.DataFrame, title: str):
    fig = px.treemap(df, path=["material"], values="value", title=title)
    fig.update_layout(margin=dict(t=50, l=10, r=10, b=10))
    return fig