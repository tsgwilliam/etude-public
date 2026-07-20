from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
import re
from typing import BinaryIO

import pandas as pd


def _truthy(value) -> bool:
    if pd.isna(value):
        return False
    return str(value).strip().lower() in {"yes", "y", "true", "1", "include", "x"}


def _read_sheet(source: Path | BinaryIO, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl")


@dataclass
class ProjectConfig:
    buildings: pd.DataFrame = field(default_factory=pd.DataFrame)
    manual_additions: pd.DataFrame = field(default_factory=pd.DataFrame)
    label_overrides: pd.DataFrame = field(default_factory=pd.DataFrame)

    @property
    def label_override_map(self) -> dict[str, str]:
        if self.label_overrides.empty:
            return {}
        cols = {c.lower(): c for c in self.label_overrides.columns}
        code_col = cols.get("nrm_code")
        label_col = cols.get("display_name")
        if not code_col or not label_col:
            return {}
        out = {}
        for _, row in self.label_overrides.iterrows():
            code = str(row[code_col]).strip()
            label = str(row[label_col]).strip()
            if code and label:
                out[code] = label
        return out

    def included_buildings(self) -> pd.DataFrame:
        if self.buildings.empty:
            return self.buildings
        cols = {c.lower(): c for c in self.buildings.columns}
        include_col = cols.get("include")
        if include_col is None:
            return self.buildings
        return self.buildings[self.buildings[include_col].map(_truthy)].copy()


def load_project_workbook(source: Path | BinaryIO | None) -> ProjectConfig:
    if source is None:
        return ProjectConfig()

    buildings = pd.DataFrame()
    manual = pd.DataFrame()
    overrides = pd.DataFrame()

    try:
        xl = pd.ExcelFile(source, engine="openpyxl")
    except Exception:
        return ProjectConfig()

    if "Buildings" in xl.sheet_names:
        buildings = _read_sheet(source, "Buildings")
    if "Manual_Additions" in xl.sheet_names:
        manual = _read_sheet(source, "Manual_Additions")
    if "Label_Overrides" in xl.sheet_names:
        overrides = _read_sheet(source, "Label_Overrides")

    return ProjectConfig(
        buildings=buildings,
        manual_additions=manual,
        label_overrides=overrides,
    )


def _default_building_name(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[_\-]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def _normalize_buildings_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["file_name", "building_name", "gia_m2", "include"])

    cols = {c.lower(): c for c in df.columns}
    out = df.copy()
    rename = {cols[k]: k for k in ["file_name", "building_name", "gia_m2", "include"] if k in cols}
    out = out.rename(columns=rename)

    for col, default in [("file_name", ""), ("building_name", ""), ("gia_m2", 0.0), ("include", "yes")]:
        if col not in out.columns:
            out[col] = default

    out["file_name"] = out["file_name"].astype(str).str.strip()
    out["building_name"] = out["building_name"].astype(str).str.strip()
    out["gia_m2"] = pd.to_numeric(out["gia_m2"], errors="coerce").fillna(0.0)
    out["include"] = out["include"].astype(str).str.strip().replace("", "yes")
    return out


def _buildings_from_uploads(
    upload_file_names: list[str],
    existing: ProjectConfig | None = None,
) -> pd.DataFrame:
    existing_df = _normalize_buildings_table(existing.buildings) if existing and not existing.buildings.empty else pd.DataFrame()

    rows: list[dict] = []
    for index, file_name in enumerate(upload_file_names):
        file_name = str(file_name).strip()
        building_name = _default_building_name(file_name)
        gia_m2 = 0.0

        exact = existing_df[existing_df["file_name"] == file_name] if not existing_df.empty else pd.DataFrame()
        if not exact.empty:
            building_name = str(exact.iloc[0]["building_name"]).strip() or building_name
            gia_m2 = float(exact.iloc[0]["gia_m2"])
        elif not existing_df.empty and index < len(existing_df):
            brow = existing_df.iloc[index]
            building_name = str(brow["building_name"]).strip() or building_name
            gia_m2 = float(brow["gia_m2"])

        rows.append(
            {
                "file_name": file_name,
                "building_name": building_name,
                "gia_m2": gia_m2,
                "include": "yes",
            }
        )

    if rows:
        return pd.DataFrame(rows)

    return pd.DataFrame(
        {
            "file_name": ["detailReport_example.xls"],
            "building_name": ["Building 1"],
            "gia_m2": [0.0],
            "include": ["yes"],
        }
    )


def building_names_for_uploads(
    upload_file_names: list[str],
    existing: ProjectConfig | None = None,
) -> list[str]:
    return (
        _buildings_from_uploads(upload_file_names, existing)["building_name"]
        .astype(str)
        .str.strip()
        .tolist()
    )


def create_project_workbook_bytes(
    upload_file_names: list[str] | None = None,
    existing: ProjectConfig | None = None,
) -> bytes:
    """Build a project workbook, pre-filling Buildings from uploaded OneClick filenames."""
    buildings = _buildings_from_uploads(upload_file_names or [], existing)

    if existing is not None and not existing.manual_additions.empty:
        manual = existing.manual_additions.copy()
    else:
        manual = pd.DataFrame(
            columns=[
                "building_name",
                "life_stage",
                "nrm_code",
                "label",
                "kgco2e_per_m2_gia",
                "etude_group",
                "notes",
            ]
        )

    if existing is not None and not existing.label_overrides.empty:
        overrides = existing.label_overrides.copy()
    else:
        overrides = pd.DataFrame(
            {
                "nrm_code": ["2.5.1"],
                "display_name": ["External walls (above ground)"],
            }
        )

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        buildings.to_excel(writer, sheet_name="Buildings", index=False)
        manual.to_excel(writer, sheet_name="Manual_Additions", index=False)
        overrides.to_excel(writer, sheet_name="Label_Overrides", index=False)
    return buffer.getvalue()


def create_blank_project_workbook_bytes() -> bytes:
    return create_project_workbook_bytes(upload_file_names=None, existing=None)
