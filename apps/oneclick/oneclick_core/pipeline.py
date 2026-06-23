from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import tempfile
from typing import BinaryIO

import pandas as pd

from oneclick_detail import load_detail_report
from oneclick_core.config import ProjectConfig
from oneclick_core.mappings import apply_rule_column, load_label_defaults, load_mapping_csv
from oneclick_core.nrm import add_nrm_columns, extract_nrm_code


@dataclass
class ProjectDataset:
    rows: pd.DataFrame
    meta_by_building: dict[str, dict]
    em_col: str = "kgco2e"


def default_building_name(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[_\-]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


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
    return out


def parse_oneclick_path(path: Path) -> tuple[pd.DataFrame, dict]:
    report = load_detail_report(path)
    rows = drop_oneclick_total_rows(report.rows)
    return rows, report.meta or {}


def parse_uploaded_bytes(name: str, data: bytes) -> tuple[pd.DataFrame, dict]:
    suffix = Path(name).suffix.lower() or ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(data)
        tmp.flush()
        return parse_oneclick_path(Path(tmp.name))


def _building_lookup(project: ProjectConfig, uploaded_names: list[str]) -> pd.DataFrame:
    cols = {c.lower(): c for c in project.buildings.columns} if not project.buildings.empty else {}
    if not cols:
        return pd.DataFrame(
            {
                "file_name": uploaded_names,
                "building_name": [default_building_name(n) for n in uploaded_names],
                "gia_m2": [0.0] * len(uploaded_names),
                "include": ["yes"] * len(uploaded_names),
            }
        )

    df = project.included_buildings().copy()
    rename = {}
    for key, original in cols.items():
        if key == "file_name":
            rename[original] = "file_name"
        elif key == "building_name":
            rename[original] = "building_name"
        elif key == "gia_m2":
            rename[original] = "gia_m2"
        elif key == "include":
            rename[original] = "include"
    df = df.rename(columns=rename)

    for col in ["file_name", "building_name", "gia_m2"]:
        if col not in df.columns:
            if col == "gia_m2":
                df[col] = 0.0
            else:
                df[col] = ""

    if df.empty:
        return pd.DataFrame(
            {
                "file_name": uploaded_names,
                "building_name": [default_building_name(n) for n in uploaded_names],
                "gia_m2": [0.0] * len(uploaded_names),
            }
        )

    known = set(df["file_name"].astype(str))
    extras = [n for n in uploaded_names if n not in known]
    if extras:
        extra_df = pd.DataFrame(
            {
                "file_name": extras,
                "building_name": [default_building_name(n) for n in extras],
                "gia_m2": [0.0] * len(extras),
            }
        )
        df = pd.concat([df, extra_df], ignore_index=True)
    return df


def _manual_rows(project: ProjectConfig, building_lookup: pd.DataFrame, em_col: str) -> pd.DataFrame:
    manual = project.manual_additions
    if manual.empty:
        return pd.DataFrame()

    cols = {c.lower(): c for c in manual.columns}
    required = ["building_name", "life_stage", "nrm_code", "label", "kgco2e_per_m2_gia"]
    if not all(k in cols for k in required):
        return pd.DataFrame()

    gia_map = dict(
        zip(
            building_lookup["building_name"].astype(str).str.strip(),
            pd.to_numeric(building_lookup["gia_m2"], errors="coerce").fillna(0.0),
        )
    )

    rows = []
    for i, row in manual.iterrows():
        building = str(row[cols["building_name"]]).strip()
        gia = float(gia_map.get(building, 0.0))
        intensity = float(pd.to_numeric(row[cols["kgco2e_per_m2_gia"]], errors="coerce") or 0.0)
        nrm_code = str(row[cols["nrm_code"]]).strip()
        label = str(row[cols["label"]]).strip()
        life_stage = str(row[cols["life_stage"]]).strip()
        etude_group = str(row[cols["etude_group"]]).strip() if "etude_group" in cols else ""

        rows.append(
            {
                "section": life_stage,
                "rics_detail": f"{nrm_code}.{label}" if nrm_code and label else label or nrm_code,
                "rics_high_label": "",
                "rics_level2_label": "",
                "rics_alloc_label": nrm_code,
                "rics_allocated_value": intensity * gia,
                em_col: intensity * gia,
                "element_name": label,
                "Comment": "Manual entry",
                "building_name": building,
                "building_gia_m2": gia,
                "source_file": "manual_workbook",
                "_source_row_id": f"manual__{building}__{i}",
                "is_manual": True,
                "nrm_code": extract_nrm_code(nrm_code),
                "etude_group_seed": etude_group,
            }
        )
    return pd.DataFrame(rows)


def build_canonical_dataset(
    uploads: list[tuple[str, bytes]],
    project: ProjectConfig,
    config_dir: Path,
    selected_buildings: list[str] | None = None,
    nrm_level: int = 2,
) -> ProjectDataset:
    uploaded_names = [name for name, _ in uploads]
    building_lookup = _building_lookup(project, uploaded_names)

    if selected_buildings:
        building_lookup = building_lookup[
            building_lookup["building_name"].astype(str).isin(selected_buildings)
        ].copy()

    material_rules = load_mapping_csv(config_dir / "material_family_map.csv")
    etude_rules = load_mapping_csv(config_dir / "etude_group_rules.csv")
    label_defaults = load_label_defaults(config_dir / "nrm_label_defaults.csv")
    label_overrides = project.label_override_map

    all_rows: list[pd.DataFrame] = []
    meta_by_building: dict[str, dict] = {}

    lookup = building_lookup.set_index("file_name", drop=False)
    for file_name, payload in uploads:
        if file_name not in lookup.index:
            continue
        building_name = str(lookup.at[file_name, "building_name"]).strip()
        gia = float(pd.to_numeric(lookup.at[file_name, "gia_m2"], errors="coerce") or 0.0)

        rows, meta = parse_uploaded_bytes(file_name, payload)
        rows = rows.copy()
        rows["building_name"] = building_name
        rows["building_gia_m2"] = gia
        rows["source_file"] = file_name
        rows["is_manual"] = False
        all_rows.append(rows)
        meta_by_building[building_name] = meta

    if not all_rows:
        manual_only = _manual_rows(project, building_lookup, em_col="kgco2e")
        if manual_only.empty:
            return ProjectDataset(rows=pd.DataFrame(), meta_by_building={})
        combined = manual_only
    else:
        combined = pd.concat(all_rows, ignore_index=True)
        manual = _manual_rows(project, building_lookup, em_col="kgco2e")
        if not manual.empty:
            combined = pd.concat([combined, manual], ignore_index=True)

    if "rics_allocated_value" not in combined.columns:
        combined["rics_allocated_value"] = pd.to_numeric(combined.get("kgco2e", 0.0), errors="coerce").fillna(0.0)

    combined = add_nrm_columns(
        combined,
        nrm_level=nrm_level,
        label_defaults=label_defaults,
        label_overrides=label_overrides,
    )

    combined["match_text"] = (
        combined.get("material_label", pd.Series("", index=combined.index)).astype(str)
        + " "
        + combined.get("Resource", pd.Series("", index=combined.index)).astype(str)
        + " "
        + combined.get("element_name", pd.Series("", index=combined.index)).astype(str)
    ).str.strip()

    combined["material_family"] = apply_rule_column(
        combined,
        material_rules,
        target_col="material_family",
        value_col="material_family",
        default="Other",
    )

    if "etude_group_seed" in combined.columns:
        seeded = combined["etude_group_seed"].astype(str).str.strip()
        combined["etude_group"] = seeded.where(seeded != "", None)
    else:
        combined["etude_group"] = None

    assigned = combined["etude_group"].notna() & combined["etude_group"].astype(str).str.strip().ne("")
    if not assigned.all() and not etude_rules.empty:
        inferred = apply_rule_column(
            combined.loc[~assigned],
            etude_rules,
            target_col="etude_group",
            value_col="etude_group",
            default="",
        )
        combined.loc[~assigned, "etude_group"] = inferred

    combined["etude_group"] = combined["etude_group"].fillna("").astype(str)

    return ProjectDataset(rows=combined, meta_by_building=meta_by_building, em_col="kgco2e")
