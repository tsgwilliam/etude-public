from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest

from oneclick_core.config import create_project_workbook_bytes, load_project_workbook
from oneclick_core.nrm import add_nrm_columns, extract_nrm_code, nrm_at_level
from oneclick_core.pipeline import (
    build_canonical_dataset,
    default_building_name,
    validate_detail_report_bytes,
)


@pytest.fixture
def uploads_dir() -> Path:
    return Path("/home/ubuntu/.cursor/projects/workspace/uploads")


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


def test_workbook_file_name_mismatch_still_pairs_upload():
    """Workbook file_name need not match upload if there is one row per upload."""
    from oneclick_core.config import create_project_workbook_bytes

    config_dir = Path(__file__).resolve().parents[1] / "config"
    sample = Path(
        "/home/ubuntu/.cursor/projects/workspace/uploads/"
        "detailReport_02.06.2026_11_30_33_71b5.xls"
    )
    if not sample.exists():
        pytest.skip("user sample not available")

    project = load_project_workbook(BytesIO(create_project_workbook_bytes()))
    # Simulate user putting a label instead of the real upload filename
    project.buildings.loc[0, "file_name"] = "Test 1"
    project.buildings.loc[0, "building_name"] = "Building 1"
    project.buildings.loc[0, "gia_m2"] = 10000

    dataset = build_canonical_dataset(
        uploads=[("detailReport_02.06.2026_11_30_33.xls", sample.read_bytes())],
        project=project,
        config_dir=config_dir,
        selected_buildings=["Building 1"],
        nrm_level=2,
    )
    assert len(dataset.rows) > 100
    assert dataset.rows["building_name"].eq("Building 1").all()
    assert float(dataset.rows["rics_allocated_value"].sum()) > 0


def test_default_building_name():
    assert "detailReport" in default_building_name("detailReport_12.05.2026.xls")


def test_detail_reports_ingest(uploads_dir: Path):
    files = [
        "detailReport_02.06.2026_11_30_33_29aa.xls",
        "detailReport_08.05.2026_16_00_10_35c0.xls",
        "detailReport_12.05.2026_11_26_26_WITH_FINISHES_AMENDED_a61d.xls",
    ]
    uploads = []
    for name in files:
        path = uploads_dir / name
        if not path.exists():
            pytest.skip(f"missing {name}")
        uploads.append((name, path.read_bytes()))

    dataset = build_canonical_dataset(
        uploads=uploads,
        project=load_project_workbook(None),
        config_dir=Path(__file__).resolve().parents[1] / "config",
    )
    assert len(dataset.rows) > 3000


def test_summary_export_rejected(uploads_dir: Path):
    summary = uploads_dir / "detailReport_02.06.2026_11_30_11_3f89.xls"
    if not summary.exists():
        pytest.skip("summary sample not available")
    with pytest.raises(ValueError, match="summary/results"):
        validate_detail_report_bytes(summary.name, summary.read_bytes())


def test_subtotal_rows_do_not_double_count_john_lobb():
    sample = Path(
        "/home/ubuntu/.cursor/projects/workspace/uploads/"
        "detailReport_20.07.2026_16_34_32_9972.xls"
    )
    if not sample.exists():
        pytest.skip("John Lobb detail report not available")

    dataset = build_canonical_dataset(
        uploads=[(sample.name, sample.read_bytes())],
        project=load_project_workbook(None),
        config_dir=Path(__file__).resolve().parents[1] / "config",
        nrm_level=1,
    )
    upfront = dataset.rows[dataset.rows["section"].isin(["A1-A3", "A4", "A5"])]
    total = float(upfront["rics_allocated_value"].sum())
    # Spreadsheet New Build GWP A1-A5 structure total before contingency/manual overrides
    assert 2_900_000 < total < 3_200_000, total
    sub = float(
        upfront.loc[upfront["chart_high"] == "Substructure", "rics_allocated_value"].sum()
    )
    assert abs(sub - 1_536_916.5) < 1.0, sub
