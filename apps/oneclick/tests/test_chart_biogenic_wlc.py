from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

from oneclick_core.charts import (
    DEFAULT_WLC_MODULES,
    UPFRONT_MODULES,
    filter_rows_for_modules,
    modules_include_lifecycle,
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


def test_modules_include_lifecycle():
    assert modules_include_lifecycle(UPFRONT_MODULES) is False
    assert modules_include_lifecycle(DEFAULT_WLC_MODULES) is True


def test_wlc_nets_biogenic_into_categories(john_lobb_rows):
    rr = filter_rows_for_modules(john_lobb_rows, DEFAULT_WLC_MODULES)
    assert "bioC" in set(rr["section"].astype(str))
    vc = pick_chart_value_col(rr)
    by_high = rr.groupby("chart_high")[vc].sum() / GIA
    # Matches John Lobb spreadsheet Total excl B6&B7 (bio netted per category).
    assert abs(float(by_high["Substructure"]) - 288.9) < 1.0
    assert abs(float(by_high["Superstructure"]) - 199.6) < 3.0


def test_upfront_excludes_biogenic_from_stack(john_lobb_rows):
    rr = filter_rows_for_modules(john_lobb_rows, UPFRONT_MODULES)
    assert "bioC" not in set(rr["section"].astype(str).str.lower())
    vc = pick_chart_value_col(rr)
    by_high = rr.groupby("chart_high")[vc].sum() / GIA
    assert abs(float(by_high["Substructure"]) - 277.4) < 1.0
    assert abs(float(by_high["Superstructure"]) - 178.1) < 1.0


def test_biogenic_bar_labeled_on_upfront_only(john_lobb_rows):
    fig_up = plot_rics_single_stack(
        john_lobb_rows,
        UPFRONT_MODULES,
        title="up",
        use_intensity=True,
        gia_m2=GIA,
    )
    bio_texts = [
        t.text[0]
        for t in fig_up.data
        if getattr(t, "text", None) and t.text and str(t.text[0]).startswith("Biogenic")
    ]
    assert bio_texts, "Upfront chart should label the biogenic bar"

    fig_wlc = plot_rics_single_stack(
        john_lobb_rows,
        DEFAULT_WLC_MODULES,
        title="wlc",
        use_intensity=True,
        gia_m2=GIA,
    )
    bio_texts_wlc = [
        t.text[0]
        for t in fig_wlc.data
        if getattr(t, "text", None) and t.text and str(t.text[0]).startswith("Biogenic")
    ]
    assert not bio_texts_wlc, "WLC nets biogenic into categories — no separate bio bar"


def test_two_stack_totals_sit_below_biogenic(john_lobb_rows):
    fig = plot_rics_two_stacks(
        john_lobb_rows,
        UPFRONT_MODULES,
        DEFAULT_WLC_MODULES,
        title="combined",
        use_intensity=True,
        gia_m2=GIA,
        project_contingency_pct=9.0,
    )
    # Biogenic bar is negative on the upfront stack only.
    bio_ys = [float(t.y[0]) for t in fig.data if float(t.y[0]) < -1]
    assert bio_ys, "Expected a biogenic bar under upfront"
    bio_floor = min(bio_ys)
    total_anns = [a for a in fig.layout.annotations if a.text and "Upfront" in str(a.text)]
    assert total_anns
    assert all(float(a.y) < bio_floor for a in total_anns)
