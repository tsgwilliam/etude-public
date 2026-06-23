from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from oneclick_core.charts import (
    DEFAULT_WLC_MODULES,
    UPFRONT_MODULES,
    aggregate_chart_data,
    plot_rics_single_stack,
    plot_rics_two_stacks,
)
from oneclick_core.config import create_blank_project_workbook_bytes, load_project_workbook
from oneclick_core.export import build_results_workbook_bytes
from oneclick_core.pipeline import build_canonical_dataset, default_building_name

APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = APP_DIR / "config"


def _uploads_from_session() -> list[tuple[str, bytes]]:
    files = st.session_state.get("oneclick_uploads", [])
    return [(name, data) for name, data in files]


def _remember_uploads(uploaded_files) -> None:
    stored: list[tuple[str, bytes]] = []
    for uploaded in uploaded_files or []:
        uploaded.seek(0)
        stored.append((uploaded.name, uploaded.getvalue()))
    st.session_state["oneclick_uploads"] = stored


def _project_config_bytes() -> bytes | None:
    project_file = st.session_state.get("project_workbook_bytes")
    return project_file


def main() -> None:
    st.set_page_config(page_title="Etude OneClick LCA", layout="wide")
    st.title("Etude OneClick LCA reporting")
    st.caption(
        "Upload OneClick detailReport exports and an optional project workbook. "
        "Global mappings live in the repo CSVs; project-specific GIA and manual rows live in the workbook."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        uploaded_files = st.file_uploader(
            "OneClick detailReport exports (.xls/.xlsx)",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
        )
    with col_b:
        project_file = st.file_uploader(
            "Project workbook (.xlsx)",
            type=["xlsx"],
            help="Use the template for Buildings, Manual_Additions, and Label_Overrides sheets.",
        )

    st.download_button(
        "Download blank project workbook template",
        data=create_blank_project_workbook_bytes(),
        file_name="Etude_OneClick_Project.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if uploaded_files:
        _remember_uploads(uploaded_files)
    if project_file is not None:
        project_file.seek(0)
        st.session_state["project_workbook_bytes"] = project_file.getvalue()

    uploads = _uploads_from_session()
    if not uploads:
        st.info("Upload at least one OneClick detailReport export to begin.")
        return

    project = load_project_workbook(BytesIO(_project_config_bytes()) if _project_config_bytes() else None)

    building_names = []
    if not project.buildings.empty:
        cols = {c.lower(): c for c in project.buildings.columns}
        if "building_name" in cols:
            building_names = (
                project.included_buildings()[cols["building_name"]].astype(str).str.strip().tolist()
            )
    if not building_names:
        building_names = [default_building_name(name) for name, _ in uploads]

    with st.sidebar:
        st.header("Controls")
        chart_choice = st.radio(
            "Chart",
            options=[
                "1) Upfront (A1–A5)",
                "2) Life cycle embodied (excl. B6, B7 and D)",
                "3) Upfront + life cycle embodied",
            ],
            index=0,
        )
        nrm_level = st.selectbox("NRM display level", options=[1, 2, 3, 4], index=1)
        selected_buildings = st.multiselect("Buildings", options=building_names, default=building_names)
        use_intensity = st.checkbox("Show intensity (per m² GIA)", value=True)
        chart_height = st.slider("Chart height (px)", 500, 1600, 900, 50)
        collapsed_high_level = st.checkbox("Show only high-level NRM labels", value=nrm_level == 1)
        show_target = st.checkbox("Show benchmark line", value=False)
        target_value = st.number_input("Benchmark value", value=0.0, step=1.0, disabled=not show_target)
        target_label = st.text_input("Benchmark label", value="Target", disabled=not show_target)

    try:
        dataset = build_canonical_dataset(
            uploads=uploads,
            project=project,
            config_dir=CONFIG_DIR,
            selected_buildings=selected_buildings,
            nrm_level=int(nrm_level),
        )
    except Exception as exc:
        st.error(f"Failed to build dataset: {exc}")
        return

    for warning in dataset.warnings:
        st.warning(warning)

    rows = dataset.rows
    if rows.empty:
        st.warning("No rows matched the selected buildings.")
        return

    gia_by_building = (
        rows.groupby("building_name", dropna=False)["building_gia_m2"]
        .first()
        .astype(float)
        .fillna(0.0)
    )
    total_gia = float(gia_by_building.sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows", f"{len(rows):,}")
    c2.metric("Buildings", str(len(selected_buildings)))
    c3.metric("Combined GIA", f"{total_gia:,.0f} m²")

    if use_intensity and total_gia <= 0:
        st.warning("Set GIA values in the project workbook Buildings sheet to show kgCO₂e/m² GIA.")

    title_suffix = ", ".join(selected_buildings) if len(selected_buildings) <= 3 else f"{len(selected_buildings)} buildings"
    target = float(target_value) if show_target else None

    if chart_choice.startswith("1"):
        fig = plot_rics_single_stack(
            rows=rows,
            modules=UPFRONT_MODULES,
            title=f"Upfront embodied carbon (A1–A5) — {title_suffix}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            collapsed_high_level=collapsed_high_level,
            height_px=chart_height,
            target_line_value=target,
            target_line_label=target_label,
        )
    elif chart_choice.startswith("2"):
        fig = plot_rics_single_stack(
            rows=rows,
            modules=DEFAULT_WLC_MODULES,
            title=f"Life cycle embodied carbon — {title_suffix}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            collapsed_high_level=collapsed_high_level,
            height_px=chart_height,
            target_line_value=target,
            target_line_label=target_label,
        )
    else:
        fig = plot_rics_two_stacks(
            rows=rows,
            upfront_modules=UPFRONT_MODULES,
            whole_life_modules=DEFAULT_WLC_MODULES,
            title=f"Upfront + life cycle embodied carbon — {title_suffix}",
            use_intensity=use_intensity,
            gia_m2=total_gia,
            collapsed_high_level=collapsed_high_level,
            height_px=chart_height,
        )

    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Chart data preview"):
        modules = UPFRONT_MODULES if chart_choice.startswith("1") else DEFAULT_WLC_MODULES
        st.dataframe(aggregate_chart_data(rows, modules), use_container_width=True)

    with st.expander("Canonical data preview"):
        preview_cols = [
            c
            for c in [
                "building_name",
                "section",
                "nrm_code",
                "chart_segment_label",
                "material_family",
                "etude_group",
                "rics_allocated_value",
                "material_label",
            ]
            if c in rows.columns
        ]
        st.dataframe(rows[preview_cols].head(200), use_container_width=True)

    project_name = selected_buildings[0] if len(selected_buildings) == 1 else "combined_project"
    excel_bytes = build_results_workbook_bytes(
        rows=rows,
        project_name=project_name,
        nrm_level=int(nrm_level),
        total_gia=total_gia,
        selected_buildings=selected_buildings,
    )
    st.download_button(
        "Download Excel results workbook",
        data=excel_bytes,
        file_name=f"{project_name}_OneClick_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
