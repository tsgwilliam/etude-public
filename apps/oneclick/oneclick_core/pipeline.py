from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
import tempfile

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
    warnings: list[str] = field(default_factory=list)


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
    if not data:
        raise ValueError(f"Upload '{name}' is empty (0 bytes). Re-upload the OneClick detailReport file.")

    suffix = Path(name).suffix.lower() or ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(data)
        tmp.flush()
        rows, meta = parse_oneclick_path(Path(tmp.name))
    if rows.empty:
        raise ValueError(
            f"Upload '{name}' did not produce any data rows. "
            "Check that this is a OneClick **detailReport** export (not a summary/results export)."
        )
    return rows, meta


def _normalize_buildings_df(project: ProjectConfig) -> pd.DataFrame:
    if project.buildings.empty:
        return pd.DataFrame(columns=["file_name", "building_name", "gia_m2"])

    cols = {c.lower(): c for c in project.buildings.columns}
    df = project.included_buildings().copy()
    rename = {
        cols[k]: k
        for k in ["file_name", "building_name", "gia_m2", "include"]
        if k in cols
    }
    df = df.rename(columns=rename)

    for col, default in [("file_name", ""), ("building_name", ""), ("gia_m2", 0.0)]:
        if col not in df.columns:
            df[col] = default

    df["file_name"] = df["file_name"].astype(str).str.strip()
    df["building_name"] = df["building_name"].astype(str).str.strip()
    df["gia_m2"] = pd.to_numeric(df["gia_m2"], errors="coerce").fillna(0.0)
    return df


def resolve_upload_building_pairs(
    project: ProjectConfig,
    uploaded_names: list[str],
) -> tuple[pd.DataFrame, list[str]]:
    """Map each uploaded filename to a building name and GIA from the workbook."""
    warnings: list[str] = []
    buildings = _normalize_buildings_df(project)

    if buildings.empty:
        return pd.DataFrame(
            {
                "file_name": uploaded_names,
                "building_name": [default_building_name(n) for n in uploaded_names],
                "gia_m2": [0.0] * len(uploaded_names),
            }
        ), warnings

    pairs: list[dict] = []
    unmatched_uploads = list(uploaded_names)
    assigned_building_idx: set[int] = set()

    for idx, brow in buildings.iterrows():
        workbook_name = str(brow["file_name"]).strip()
        if workbook_name and workbook_name in unmatched_uploads:
            pairs.append(
                {
                    "file_name": workbook_name,
                    "building_name": str(brow["building_name"]).strip(),
                    "gia_m2": float(brow["gia_m2"]),
                }
            )
            unmatched_uploads.remove(workbook_name)
            assigned_building_idx.add(idx)

    remaining_buildings = buildings.loc[~buildings.index.isin(assigned_building_idx)].reset_index(drop=True)
    for i, upload_name in enumerate(unmatched_uploads):
        if i < len(remaining_buildings):
            brow = remaining_buildings.iloc[i]
            workbook_name = str(brow["file_name"]).strip()
            pairs.append(
                {
                    "file_name": upload_name,
                    "building_name": str(brow["building_name"]).strip(),
                    "gia_m2": float(brow["gia_m2"]),
                }
            )
            if workbook_name and workbook_name != upload_name:
                warnings.append(
                    f"Linked upload **{upload_name}** to building **{brow['building_name']}** "
                    f"(workbook `file_name` was `{workbook_name}`, which did not match the upload)."
                )
        else:
            pairs.append(
                {
                    "file_name": upload_name,
                    "building_name": default_building_name(upload_name),
                    "gia_m2": 0.0,
                }
            )
            warnings.append(
                f"No building row available for upload **{upload_name}**; using default name and GIA 0."
            )

    return pd.DataFrame(pairs), warnings


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
        raw_building = row[cols["building_name"]]
        building = "" if pd.isna(raw_building) else str(raw_building).strip()
        intensity = float(pd.to_numeric(row[cols["kgco2e_per_m2_gia"]], errors="coerce") or 0.0)
        nrm_code = str(row[cols["nrm_code"]]).strip()
        label = str(row[cols["label"]]).strip()
        life_stage = str(row[cols["life_stage"]]).strip()
        etude_group = str(row[cols["etude_group"]]).strip() if "etude_group" in cols else ""

        if not building or building.lower() == "nan":
            target_buildings = building_lookup["building_name"].astype(str).str.strip().tolist()
        else:
            target_buildings = [building]

        if intensity == 0.0 and not label:
            continue

        for bname in target_buildings:
            if not bname:
                continue
            gia = float(gia_map.get(bname, 0.0))
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
                    "building_name": bname,
                    "building_gia_m2": gia,
                    "source_file": "manual_workbook",
                    "_source_row_id": f"manual__{bname}__{i}",
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
    building_lookup, warnings = resolve_upload_building_pairs(project, uploaded_names)

    material_rules = load_mapping_csv(config_dir / "material_family_map.csv")
    etude_rules = load_mapping_csv(config_dir / "etude_group_rules.csv")
    label_defaults = load_label_defaults(config_dir / "nrm_label_defaults.csv")
    label_overrides = project.label_override_map

    all_rows: list[pd.DataFrame] = []
    meta_by_building: dict[str, dict] = {}

    lookup = building_lookup.set_index("file_name", drop=False)
    for file_name, payload in uploads:
        if file_name not in lookup.index:
            warnings.append(f"Upload **{file_name}** was not linked to a building and was skipped.")
            continue
        building_name = str(lookup.at[file_name, "building_name"]).strip()
        gia = float(pd.to_numeric(lookup.at[file_name, "gia_m2"], errors="coerce") or 0.0)

        if selected_buildings and building_name not in selected_buildings:
            continue

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
            return ProjectDataset(rows=pd.DataFrame(), meta_by_building={}, warnings=warnings)
        combined = manual_only
    else:
        combined = pd.concat(all_rows, ignore_index=True)
        manual = _manual_rows(project, building_lookup, em_col="kgco2e")
        if not manual.empty:
            if selected_buildings:
                manual = manual[manual["building_name"].astype(str).isin(selected_buildings)]
            if not manual.empty:
                combined = pd.concat([combined, manual], ignore_index=True)

    if selected_buildings:
        combined = combined[combined["building_name"].astype(str).isin(selected_buildings)].copy()

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

    return ProjectDataset(rows=combined, meta_by_building=meta_by_building, warnings=warnings, em_col="kgco2e")
