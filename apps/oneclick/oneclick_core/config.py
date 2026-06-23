from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
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


def create_blank_project_workbook_bytes() -> bytes:
    buildings = pd.DataFrame(
        {
            "file_name": ["detailReport_example.xls"],
            "building_name": ["Building 1"],
            "gia_m2": [0.0],
            "include": ["yes"],
        }
    )
    manual = pd.DataFrame(
        {
            "building_name": ["Building 1"],
            "life_stage": ["A1-A3"],
            "nrm_code": ["5.3"],
            "label": ["Manual benchmark example"],
            "kgco2e_per_m2_gia": [0.0],
            "etude_group": [""],
            "notes": ["Set kgCO2e/m2 and building name, or leave blank to skip"],
        }
    )
    overrides = pd.DataFrame({"nrm_code": ["2.5.1"], "display_name": ["External walls (above ground)"]})

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        buildings.to_excel(writer, sheet_name="Buildings", index=False)
        manual.to_excel(writer, sheet_name="Manual_Additions", index=False)
        overrides.to_excel(writer, sheet_name="Label_Overrides", index=False)
    return buffer.getvalue()
