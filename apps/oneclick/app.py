from __future__ import annotations

import sys
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# Allow `streamlit run apps/oneclick/app.py` from the repo root (Streamlit Cloud).
APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from oneclick_core.charts import (
    DEFAULT_WLC_MODULES,
    UPFRONT_MODULES,
    aggregate_chart_data,
    plot_rics_single_stack,
    plot_rics_two_stacks,
)
from oneclick_core.config import building_names_for_uploads, create_project_workbook_bytes, load_project_workbook
from oneclick_core.export import build_results_workbook_bytes
from oneclick_core.pipeline import build_canonical_dataset
from oneclick_core import nrm as nrm_mod

try:
    from oneclick_core.config import CONFIG_MODULE_VERSION, contingency_map_from_table
except ImportError:  # pragma: no cover - stale Streamlit Cloud module cache
    import re as _re

    CONFIG_MODULE_VERSION = "stale-config"

    def contingency_map_from_table(contingency) -> dict[str, float]:
        if contingency is None or not isinstance(contingency, pd.DataFrame) or contingency.empty:
            return {}
        cols = {c.lower(): c for c in contingency.columns}
        code_col = cols.get("nrm_code")
        pct_col = cols.get("contingency_pct") or cols.get("contingency_%") or cols.get("pct")
        if not code_col or not pct_col:
            return {}
        out: dict[str, float] = {}
        for _, row in contingency.iterrows():
            code = str(row[code_col]).strip()
            if _re.fullmatch(r"\d+\.0", code):
                code = code[:-2]
            pct_raw = pd.to_numeric(row[pct_col], errors="coerce")
            if not code or code.lower() == "nan" or pd.isna(pct_raw):
                continue
            out[code] = float(pct_raw)
        return out

CONFIG_DIR = APP_DIR / "config"
APP_DEPLOY_ID = f"oneclick-{CONFIG_MODULE_VERSION}"


def _available_nrm_levels(nrm_codes: pd.Series) -> list[int]:
    """Prefer package helper; keep a local fallback for stale Cloud deploys."""
    helper = getattr(nrm_mod, "available_nrm_levels", None)
    if callable(helper):
        return helper(nrm_codes)
    max_depth = 1
    for code in nrm_codes.dropna().astype(str):
        code = code.strip()
        if not code or code.lower() == "nan":
            continue
        if code == getattr(nrm_mod, "PV_CODE", "5.PV"):
            max_depth = max(max_depth, 2)
            continue
        depth = len([p for p in code.split(".") if p])
        max_depth = max(max_depth, min(depth, 4))
    return list(range(1, max_depth + 1))


def _project_contingency_map(project) -> dict[str, float]:
    """Resolve per-NRM contingency without assuming a fresh ProjectConfig class."""
    try:
        raw = getattr(project, "contingency_map", None)
        if isinstance(raw, dict):
            return dict(raw)
    except AttributeError:
        pass
    return contingency_map_from_table(getattr(project, "contingency", None))


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
        "Global mappings live in the repo CSVs; project-specific GIA, manual rows, "
        "and per-NRM contingency live in the workbook."
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
            help=(
                "Sheets: Buildings, Manual_Additions, Label_Overrides, Contingency. "
                "Contingency rows set a % per NRM code and override the optional "
                "project-level contingency for matching codes only."
            ),
        )

    if uploaded_files:
        _remember_uploads(uploaded_files)
    if project_file is not None:
        project_file.seek(0)
        st.session_state["project_workbook_bytes"] = project_file.getvalue()

    uploads = _uploads_from_session()
    project_for_template = load_project_workbook(
        BytesIO(_project_config_bytes()) if _project_config_bytes() else None
    )
    template_label = (
        "Download project workbook (pre-filled from uploads)"
        if uploads
        else "Download blank project workbook template"
    )
    st.download_button(
        template_label,
        data=create_project_workbook_bytes(
            upload_file_names=[name for name, _ in uploads],
            existing=project_for_template if not project_for_template.buildings.empty else None,
        ),
        file_name="Etude_OneClick_Project.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if not uploads:
        st.info("Upload at least one OneClick detailReport export to begin.")
        return

    project = load_project_workbook(BytesIO(_project_config_bytes()) if _project_config_bytes() else None)

    if not project.buildings.empty:
        cols = {c.lower(): c for c in project.buildings.columns}
        if "building_name" in cols:
            building_names = (
                project.included_buildings()[cols["building_name"]].astype(str).str.strip().tolist()
            )
        else:
            building_names = []
    else:
        building_names = []

    if not building_names and uploads:
        building_names = building_names_for_uploads(
            [name for name, _ in uploads],
            existing=project if not project.buildings.empty else None,
        )

    nrm_level_options = st.session_state.get("available_nrm_levels", [1, 2, 3, 4])
    if "nrm_level" in st.session_state and st.session_state["nrm_level"] not in nrm_level_options:
        st.session_state["nrm_level"] = nrm_level_options[min(1, len(nrm_level_options) - 1)]

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
        default_level = 2 if 2 in nrm_level_options else nrm_level_options[0]
        if "nrm_level" not in st.session_state:
            st.session_state["nrm_level"] = default_level
        nrm_level = st.selectbox(
            "NRM display level",
            options=nrm_level_options,
            key="nrm_level",
            help="Only levels present in the uploaded data are listed.",
        )
        selected_buildings = st.multiselect("Buildings", options=building_names, default=building_names)
        use_intensity = st.checkbox("Show intensity (per m² GIA)", value=True)
        chart_height = st.slider("Chart height (px)", 500, 1600, 900, 50)
        collapsed_high_level = st.checkbox("Show only high-level NRM labels", value=int(nrm_level) == 1)
        show_target = st.checkbox("Show benchmark line", value=False)
        target_value = st.number_input("Benchmark value", value=0.0, step=1.0, disabled=not show_target)
        target_label = st.text_input("Benchmark label", value="Target", disabled=not show_target)

        st.divider()
        st.subheader("Contingency")
        project_contingency_pct = st.number_input(
            "Project contingency (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.0,
            step=0.5,
            help=(
                "Optional. Applied to every NRM category that does not have a "
                "manual % on the Contingency sheet. Manual per-code values override "
                "this project rate (they do not stack)."
            ),
        )
        contingency_colour = st.color_picker("Contingency colour", value="#6C6C6C")
        cont_map = _project_contingency_map(project)
        if cont_map:
            st.caption(
                "Workbook Contingency overrides: "
                + ", ".join(f"{code} → {pct:g}%" for code, pct in sorted(cont_map.items()))
            )
        else:
            st.caption("No per-NRM contingency rows in the project workbook (optional).")
        st.caption(f"Deploy: {APP_DEPLOY_ID}")

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

    levels = _available_nrm_levels(rows["nrm_code"])
    if levels and levels != st.session_state.get("available_nrm_levels"):
        st.session_state["available_nrm_levels"] = levels
        if int(nrm_level) not in levels:
            st.session_state["nrm_level"] = levels[min(1, len(levels) - 1)]
            st.rerun()

    contingency_map = _project_contingency_map(project)

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

    chart_kwargs = dict(
        use_intensity=use_intensity,
        gia_m2=total_gia,
        collapsed_high_level=collapsed_high_level,
        height_px=chart_height,
        project_contingency_pct=float(project_contingency_pct),
        contingency_map=contingency_map,
        contingency_colour=contingency_colour,
    )

    if chart_choice.startswith("1"):
        fig = plot_rics_single_stack(
            rows=rows,
            modules=UPFRONT_MODULES,
            title=f"Upfront embodied carbon (A1–A5) — {title_suffix}",
            target_line_value=target,
            target_line_label=target_label,
            **chart_kwargs,
        )
    elif chart_choice.startswith("2"):
        fig = plot_rics_single_stack(
            rows=rows,
            modules=DEFAULT_WLC_MODULES,
            title=f"Life cycle embodied carbon — {title_suffix}",
            target_line_value=target,
            target_line_label=target_label,
            **chart_kwargs,
        )
    else:
        fig = plot_rics_two_stacks(
            rows=rows,
            upfront_modules=UPFRONT_MODULES,
            whole_life_modules=DEFAULT_WLC_MODULES,
            title=f"Upfront + life cycle embodied carbon — {title_suffix}",
            **chart_kwargs,
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
