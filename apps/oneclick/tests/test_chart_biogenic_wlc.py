from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest

from oneclick_core.charts import (
    DEFAULT_WLC_MODULES,
    UPFRONT_MODULES,
    aggregate_chart_data,
    filter_rows_for_modules,
    pick_chart_value_col,
    plot_rics_single_stack,
    plot_rics_two_stacks,
)
from oneclick_core.config import create_project_workbook_bytes, load_project_workbook
from oneclick_core.pipeline import build_canonical_dataset

SAMPLE = Path(
    "/home/ubuntu/.cursor/projects/workspace/uploads/"
    "detailReport_20.07.2026_16_34_32_9972.xls"
)
GIA = 5541.0


@pytest.fixture
def john_lobb_rows():
    if not SAMPLE.exists():
        pytest.skip("John Lobb detail report not available")
    wb = create_project_workbook_bytes(upload_file_names=[SAMPLE.name])
    project = load_project_workbook(BytesIO(wb))
    project.buildings.loc[0, "building_name"] = "John Lobb"
    project.buildings.loc[0, "gia_m2"] = GIA
    dataset = build_canonical_dataset(
        uploads=[(SAMPLE.name, SAMPLE.read_bytes())],
        project=project,
        config_dir=Path(__file__).resolve().parents[1] / "config",
        nrm_level=1,
    )
    return dataset.rows


def test_stacks_exclude_bioc_section(john_lobb_rows):
    for modules in (UPFRONT_MODULES, DEFAULT_WLC_MODULES):
        rr = filter_rows_for_modules(john_lobb_rows, modules)
        assert "bioC" not in set(rr["section"].astype(str))


def test_upfront_gross_intensities(john_lobb_rows):
    rr = filter_rows_for_modules(john_lobb_rows, UPFRONT_MODULES)
    vc = pick_chart_value_col(rr)
    by_high = rr.groupby("chart_high")[vc].sum() / GIA
    assert abs(float(by_high["Substructure"]) - 277.4) < 1.0
    assert abs(float(by_high["Superstructure"]) - 178.1) < 1.0


def test_biogenic_bar_labeled_on_upfront_and_wlc(john_lobb_rows):
    for modules in (UPFRONT_MODULES, DEFAULT_WLC_MODULES):
        fig = plot_rics_single_stack(
            john_lobb_rows,
            modules,
            title="t",
            use_intensity=True,
            gia_m2=GIA,
        )
        bio_texts = [
            t.text[0]
            for t in fig.data
            if getattr(t, "text", None) and t.text and str(t.text[0]).startswith("Biogenic")
        ]
        assert bio_texts, f"Expected labeled biogenic bar for {modules}"


def test_two_stack_shows_biogenic_on_both(john_lobb_rows):
    fig = plot_rics_two_stacks(
        john_lobb_rows,
        UPFRONT_MODULES,
        DEFAULT_WLC_MODULES,
        title="combined",
        use_intensity=True,
        gia_m2=GIA,
        project_contingency_pct=9.0,
    )
    bio_traces = [
        t
        for t in fig.data
        if getattr(t, "text", None) and t.text and str(t.text[0]).startswith("Biogenic")
    ]
    assert len(bio_traces) == 2
    bio_floor = min(float(t.y[0]) for t in bio_traces)
    total_anns = [a for a in fig.layout.annotations if a.text and ("Upfront" in str(a.text) or "Whole life" in str(a.text))]
    assert total_anns
    assert all(float(a.y) < bio_floor for a in total_anns)


def test_aggregate_export_includes_biogenic_row(john_lobb_rows):
    agg = aggregate_chart_data(john_lobb_rows, DEFAULT_WLC_MODULES)
    bio = agg[agg["chart_high"] == "Biogenic"]
    assert len(bio) == 1
    assert float(bio["value"].iloc[0]) < 0
