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
    "5.PV": "Photovoltaic systems",
    "6": "Prefabricated buildings and building units",
    "7": "Work to existing buildings",
    "8": "External works",
}

PV_CODE = "5.PV"
PV_LABEL = "Photovoltaic systems"
PV_PATTERN = re.compile(r"photovolta|\bpv\b|solar\s*pv|pv\s*panel", re.IGNORECASE)


def extract_nrm_code(text: str) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    match = re.match(r"^(\d+(?:\.\d+)*)", str(text).strip())
    return match.group(1) if match else ""


def clean_rics_label(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip()
    s = re.sub(r"^\d+(?:\.\d+)*\s*", "", s)
    s = re.sub(r"^[\.\-–—:]+\s*", "", s)
    return s.strip()


def parse_rics_numeric_tuple(text: str) -> tuple[int, ...]:
    code = str(text or "").strip()
    if code == PV_CODE:
        return (5, 999)
    code = extract_nrm_code(code)
    if not code:
        return (9999,)
    return tuple(int(part) for part in code.split(".") if part.isdigit()) or (9999,)


def nrm_at_level(code: str, level: int) -> str:
    code = str(code or "").strip()
    if not code:
        return ""
    # Keep PV as its own segment from level 2+, but roll into Services (5) at level 1.
    if code == PV_CODE:
        return "5" if level <= 1 else PV_CODE
    parts = [p for p in code.split(".") if p]
    if not parts:
        return ""
    if level <= 1:
        return parts[0]
    return ".".join(parts[: min(level, len(parts))])


def nrm_high_from_code(code: str) -> str:
    code = str(code or "").strip()
    if code == PV_CODE:
        return "Services"
    top = extract_nrm_code(code).split(".")[0] if code else ""
    return NRM_HIGH_MAP.get(top, "Unclassified")


def is_pv_text(*parts: str) -> bool:
    blob = " ".join(str(p or "") for p in parts)
    return bool(PV_PATTERN.search(blob))


def available_nrm_levels(nrm_codes: pd.Series) -> list[int]:
    """Return NRM display levels that exist in the data (1..4)."""
    max_depth = 1
    for code in nrm_codes.dropna().astype(str):
        code = code.strip()
        if not code or code.lower() == "nan":
            continue
        if code == PV_CODE:
            max_depth = max(max_depth, 2)
            continue
        depth = len([p for p in code.split(".") if p])
        max_depth = max(max_depth, min(depth, 4))
    return list(range(1, max_depth + 1))


def resolve_nrm_label(
    code: str,
    detail_text: str,
    defaults: dict[str, str] | None = None,
    overrides: dict[str, str] | None = None,
) -> str:
    defaults = defaults or {}
    overrides = overrides or {}
    code = str(code or "").strip()
    if code == PV_CODE:
        return PV_LABEL
    if code in overrides and str(overrides[code]).strip():
        return str(overrides[code]).strip()
    if code in defaults and str(defaults[code]).strip():
        return str(defaults[code]).strip()
    # Prefer the canonical label for the rolled-up code before falling back to
    # the original OneClick detail string (which is often more granular).
    if code in NRM_DEFAULT_LABELS:
        return NRM_DEFAULT_LABELS[code]
    cleaned = clean_rics_label(detail_text)
    if cleaned:
        return cleaned
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

    # Separate PV under Services: keep high-level = Services, distinct segment from level 2+.
    resource = out["Resource"].astype(str) if "Resource" in out.columns else pd.Series("", index=out.index)
    material = out["material_label"].astype(str) if "material_label" in out.columns else pd.Series("", index=out.index)
    pv_mask = [
        is_pv_text(r, m, d, a)
        for r, m, d, a in zip(resource, material, source, alloc)
    ]
    out.loc[pv_mask, "nrm_code"] = PV_CODE
    out["is_pv"] = pv_mask

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
    # Force PV under Services even if original code was 6.
    out.loc[out["is_pv"], "chart_high"] = "Services"
    out["chart_segment_label"] = [
        resolve_nrm_label(code, detail, label_defaults, label_overrides)
        for code, detail in zip(out["chart_segment_code"], source)
    ]
    out["chart_segment_id"] = out["chart_segment_code"].astype(str) + "|" + out["chart_segment_label"].astype(str)
    return out


def resolve_row_contingency_pct(
    nrm_code: str,
    project_pct: float,
    contingency_map: dict[str, float] | None,
) -> float:
    """Manual NRM contingency overrides project-level; otherwise use project pct."""
    contingency_map = contingency_map or {}
    code = str(nrm_code or "").strip()
    if not contingency_map:
        return float(project_pct or 0.0)

    # Longest matching prefix wins (e.g. 1.2.1 beats 1.2 beats 1).
    best_code = ""
    best_pct = None
    for rule_code, pct in contingency_map.items():
        rule = str(rule_code).strip()
        if not rule:
            continue
        if code == rule or code.startswith(rule + "."):
            if len(rule) >= len(best_code):
                best_code = rule
                best_pct = float(pct)
    if best_pct is not None:
        return best_pct
    return float(project_pct or 0.0)


def apply_contingency_rates(
    df: pd.DataFrame,
    project_pct: float = 0.0,
    contingency_map: dict[str, float] | None = None,
) -> pd.DataFrame:
    out = df.copy()
    codes = out["nrm_code"] if "nrm_code" in out.columns else pd.Series("", index=out.index)
    out["contingency_pct"] = [
        resolve_row_contingency_pct(code, project_pct, contingency_map) for code in codes
    ]
    return out


def build_high_order(values: list[str]) -> list[str]:
    present = [v for v in values if v and str(v).lower() != "nan"]
    ordered = [h for h in NRM_HIGH_ORDER if h in present]
    ordered += [h for h in present if h not in ordered]
    return ordered
