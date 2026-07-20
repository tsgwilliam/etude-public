from __future__ import annotations

from io import BytesIO

import pandas as pd

from oneclick_core.charts import compute_contingency_total
from oneclick_core.config import create_project_workbook_bytes, load_project_workbook
from oneclick_core.nrm import (
    PV_CODE,
    add_nrm_columns,
    available_nrm_levels,
    nrm_at_level,
    resolve_row_contingency_pct,
)


def test_contingency_manual_overrides_project():
    project_pct = 9.0
    contingency_map = {"1.2.1": 5.0}

    assert resolve_row_contingency_pct("1.2.1", project_pct, contingency_map) == 5.0
    assert resolve_row_contingency_pct("1.2.1.3", project_pct, contingency_map) == 5.0
    assert resolve_row_contingency_pct("1.2.2", project_pct, contingency_map) == 9.0
    assert resolve_row_contingency_pct("2.5.1", project_pct, contingency_map) == 9.0


def test_contingency_zero_manual_overrides_project():
    assert resolve_row_contingency_pct("1.2.1", 9.0, {"1.2.1": 0.0}) == 0.0


def test_contingency_longest_prefix_wins():
    contingency_map = {"1": 10.0, "1.2": 7.0, "1.2.1": 5.0}
    assert resolve_row_contingency_pct("1.2.1", 9.0, contingency_map) == 5.0
    assert resolve_row_contingency_pct("1.2.3", 9.0, contingency_map) == 7.0
    assert resolve_row_contingency_pct("1.3", 9.0, contingency_map) == 10.0


def test_compute_contingency_total_does_not_double_apply():
    rows = pd.DataFrame(
        {
            "nrm_code": ["1.2.1", "2.5.1"],
            "kgco2e": [1000.0, 2000.0],
        }
    )
    # 5% on 1.2.1 only; 9% project on everything else → 50 + 180 = 230
    total = compute_contingency_total(
        rows,
        value_col="kgco2e",
        scale=1.0,
        project_contingency_pct=9.0,
        contingency_map={"1.2.1": 5.0},
    )
    assert abs(total - 230.0) < 1e-9


def test_project_workbook_includes_contingency_sheet():
    workbook = create_project_workbook_bytes(upload_file_names=["a.xls"])
    project = load_project_workbook(BytesIO(workbook))
    assert list(project.contingency.columns) == ["nrm_code", "contingency_pct", "notes"]
    assert project.contingency_map == {}

    # Round-trip with a manual override row
    project.contingency = pd.DataFrame(
        {"nrm_code": ["1.2.1"], "contingency_pct": [5.0], "notes": ["test"]}
    )
    workbook2 = create_project_workbook_bytes(upload_file_names=["a.xls"], existing=project)
    loaded = load_project_workbook(BytesIO(workbook2))
    assert loaded.contingency_map == {"1.2.1": 5.0}


def test_pv_maps_under_services_and_separates_from_level_2():
    df = pd.DataFrame(
        {
            "rics_detail": ["5.8.1 Electrical installations", "5.8.1 Electrical installations"],
            "rics_alloc_label": ["5 Services", "5 Services"],
            "Resource": ["Photovoltaic panel system", "Cable tray"],
            "material_label": ["PV modules", "Steel"],
            "kgco2e": [100.0, 50.0],
        }
    )
    out = add_nrm_columns(df, nrm_level=2)
    pv = out[out["Resource"].str.contains("Photovoltaic", case=False)]
    other = out[out["Resource"].str.contains("Cable", case=False)]

    assert pv["nrm_code"].iloc[0] == PV_CODE
    assert pv["chart_high"].iloc[0] == "Services"
    assert "Photovoltaic" in str(pv["chart_segment_label"].iloc[0])
    assert nrm_at_level(PV_CODE, 1) == "5"
    assert nrm_at_level(PV_CODE, 2) == PV_CODE

    level1 = add_nrm_columns(df, nrm_level=1)
    assert level1.loc[level1["is_pv"], "chart_segment_code"].iloc[0] == "5"
    assert other["nrm_code"].iloc[0] != PV_CODE


def test_available_nrm_levels_only_populated_depths():
    codes = pd.Series(["1", "2.5", "2.5.1", "5.PV"])
    assert available_nrm_levels(codes) == [1, 2, 3]

    shallow = pd.Series(["1", "2", "5"])
    assert available_nrm_levels(shallow) == [1]
