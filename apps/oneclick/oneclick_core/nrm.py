from __future__ import annotations

import re

import pandas as pd

NRM_HIGH_ORDER = [
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

NRM_HIGH_MAP = {
    "0": "Deconstruction",
    "1": "Substructure",
    "2": "Superstructure",
    "3": "Finishes",
    "4": "Fittings, furnishings and equipment",
    "5": "Services",
    "6": "Prefabricated buildings and building units",
    "7": "Work to existing buildings",
    "8": "External works",
    "9": "Main contractor preliminaries",
}

NRM_DEFAULT_LABELS: dict[str, str] = {
    "0": "Deconstruction",
    "1": "Substructure",
    "2": "Superstructure",
    "2.1": "Frame",
    "2.2": "Upper floors",
    "2.3": "Roof",
    "2.4": "Stairs and ramps",
    "2.5": "External walls",
    "2.6": "Windows and external doors",
    "2.7": "Internal walls and partitions",
    "2.8": "Internal doors",
    "3": "Finishes",
    "4": "Fittings, furnishings and equipment",
    "5": "Services",
    "6": "Prefabricated buildings and building units",
    "7": "Work to existing buildings",
    "8": "External works",
}


def extract_nrm_code(text: str) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    match = re.match(r"^(\d+(?:\.\d+)*)", str(text).strip())
    return match.group(1) if match else ""


def nrm_at_level(code: str, level: int) -> str:
    code = str(code or "").strip()
    if not code:
        return ""
    parts = [p for p in code.split(".") if p]
    if not parts:
        return ""
    if level <= 1:
        return parts[0]
    return ".".join(parts[: min(level, len(parts))])


def nrm_high_from_code(code: str) -> str:
    top = extract_nrm_code(code).split(".")[0] if code else ""
    return NRM_HIGH_MAP.get(top, "Unclassified")


def clean_rics_label(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip()
    s = re.sub(r"^\d+(?:\.\d+)*\s*", "", s)
    s = re.sub(r"^[\.\-–—:]+\s*", "", s)
    return s.strip()


def parse_rics_numeric_tuple(text: str) -> tuple[int, ...]:
    code = extract_nrm_code(text)
    if not code:
        return (9999,)
    return tuple(int(part) for part in code.split("."))


def resolve_nrm_label(
    code: str,
    detail_text: str,
    defaults: dict[str, str] | None = None,
    overrides: dict[str, str] | None = None,
) -> str:
    defaults = defaults or {}
    overrides = overrides or {}
    if code in overrides and str(overrides[code]).strip():
        return str(overrides[code]).strip()
    if code in defaults and str(defaults[code]).strip():
        return str(defaults[code]).strip()
    cleaned = clean_rics_label(detail_text)
    if cleaned:
        return cleaned
    if code in NRM_DEFAULT_LABELS:
        return NRM_DEFAULT_LABELS[code]
    return code or "Unclassified"


def add_nrm_columns(
    df: pd.DataFrame,
    nrm_level: int = 3,
    label_defaults: dict[str, str] | None = None,
    label_overrides: dict[str, str] | None = None,
    detail_col: str = "rics_detail",
) -> pd.DataFrame:
    out = df.copy()
    source = out[detail_col].astype(str) if detail_col in out.columns else pd.Series("", index=out.index)
    alloc = out["rics_alloc_label"].astype(str) if "rics_alloc_label" in out.columns else pd.Series("", index=out.index)

    out["nrm_code"] = source.where(
        source.str.extract(r"^(\d)", expand=False).notna(),
        alloc.map(extract_nrm_code),
    ).map(extract_nrm_code)

    missing = out["nrm_code"].eq("")
    if missing.any():
        out.loc[missing, "nrm_code"] = alloc.loc[missing].map(extract_nrm_code)

    out["nrm_level_1"] = out["nrm_code"].map(lambda c: nrm_at_level(c, 1))
    out["nrm_level_2"] = out["nrm_code"].map(lambda c: nrm_at_level(c, 2))
    out["nrm_level_3"] = out["nrm_code"].map(lambda c: nrm_at_level(c, 3))
    out["nrm_level_4"] = out["nrm_code"]

    level_col = {
        1: "nrm_level_1",
        2: "nrm_level_2",
        3: "nrm_level_3",
        4: "nrm_level_4",
    }.get(nrm_level, "nrm_level_3")

    out["chart_segment_code"] = out[level_col]
    out["chart_high"] = out["nrm_code"].map(nrm_high_from_code)
    out["chart_segment_label"] = [
        resolve_nrm_label(code, detail, label_defaults, label_overrides)
        for code, detail in zip(out["chart_segment_code"], source)
    ]
    out["chart_segment_id"] = out["chart_segment_code"].astype(str) + "|" + out["chart_segment_label"].astype(str)
    return out


def build_high_order(values: list[str]) -> list[str]:
    present = [v for v in values if v and str(v).lower() != "nan"]
    ordered = [h for h in NRM_HIGH_ORDER if h in present]
    ordered += [h for h in present if h not in ordered]
    return ordered
