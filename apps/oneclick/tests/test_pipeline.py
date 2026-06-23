from __future__ import annotations

from pathlib import Path

import pytest

from oneclick_core.config import load_project_workbook
from oneclick_core.nrm import add_nrm_columns, extract_nrm_code, nrm_at_level
from oneclick_core.pipeline import build_canonical_dataset, default_building_name


SAMPLE_XLS = Path(
    "/home/ubuntu/.cursor/projects/workspace/uploads/"
    "detailReport_12.05.2026_11_26_26_WITH_FINISHES_AMENDED_7199.xls"
)


def test_nrm_code_helpers():
    assert extract_nrm_code("2.5.1.External walls") == "2.5.1"
    assert nrm_at_level("2.5.1", 1) == "2"
    assert nrm_at_level("2.5.1", 2) == "2.5"
    assert nrm_at_level("2.5.1", 3) == "2.5.1"


@pytest.mark.skipif(not SAMPLE_XLS.exists(), reason="sample export not available")
def test_build_canonical_dataset_from_sample():
    config_dir = Path(__file__).resolve().parents[1] / "config"
    payload = SAMPLE_XLS.read_bytes()
    dataset = build_canonical_dataset(
        uploads=[(SAMPLE_XLS.name, payload)],
        project=load_project_workbook(None),
        config_dir=config_dir,
        nrm_level=2,
    )
    assert not dataset.rows.empty
    assert "material_family" in dataset.rows.columns
    assert "chart_segment_label" in dataset.rows.columns
    assert dataset.rows["nrm_code"].astype(str).str.len().gt(0).any()
    assert dataset.rows["material_family"].ne("Other").any()


def test_default_building_name():
    assert "detailReport" in default_building_name("detailReport_12.05.2026.xls")
