import math
import base64
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    import PySAM.Pvwattsv8 as pvwatts
except ImportError:
    pvwatts = None


st.set_page_config(page_title="Etude PV Requirement Calculator", layout="wide")

STANDARD_PANEL_EFFICIENCY_KWP_PER_M2 = 0.22
FHS_REQUIRED_AREA_FRACTION = 0.40

DEFAULT_OFFSETS_M = {
    "ridge": 0.6,
    "roof_edge": 0.5,
    "party_wall": 0.75,
}

DEFAULT_SPREADSHEET_MARGIN_M = 0.3
PITCHED_INTER_MODULE_GAP_M = 0.025
FLAT_INTER_MODULE_GAP_M = 0.10

# -----------------------------------------------------------------------------
# SAP compliance placeholders
# These are temporary placeholder annual yield figures for the Part L section.
# They should be replaced with checked SAP / approved-methodology values.
# -----------------------------------------------------------------------------
SAP_PLACEHOLDER_SPECIFIC_YIELD = {
    "England - North": 800,
    "England - Midlands": 860,
    "England - South": 920,
    "England - South West": 900,
    "England - London / South East": 930,
}
SAP_PLACEHOLDER_NOTE = (
    "SAP annual generation inputs in this section are placeholders only and must "
    "be checked against the approved Part L / SAP methodology."
)

# -----------------------------------------------------------------------------
# Module assumptions
# -----------------------------------------------------------------------------
FIXED_MODULE_WIDTH_M = 1.134

MODULE_LENGTH_OPTIONS_M = [
    1.722,
    1.762,
    1.900,
    2.000,
    2.278,
]

DEFAULT_MODULE_LENGTH_M = 1.722
DEFAULT_MODULE_EFFICIENCY_PCT = 22.0
MIN_MODULE_EFFICIENCY_PCT = 19.0
MAX_MODULE_EFFICIENCY_PCT = 25.0
MODULE_EFFICIENCY_STEP_PCT = 0.1

# -----------------------------------------------------------------------------
# PySAM / weather file controls
# -----------------------------------------------------------------------------
EPW_DIRECTORY = Path("resources") / "epw"

PYSAM_DC_AC_RATIO = 1.1
PYSAM_SYSTEM_LOSSES_PCT = 14.0
PYSAM_ARRAY_TYPE_FIXED_ROOF = 1
PYSAM_MODULE_TYPE_STANDARD = 0
PYSAM_GCR = 0.4

# -----------------------------------------------------------------------------
# Graph styling controls
# -----------------------------------------------------------------------------
CHART_HEIGHT = 750
CHART_TITLE = "Part L target vs available PV area"
CHART_BACKGROUND_COLOUR = "white"
CHART_PLOT_BACKGROUND_COLOUR = "white"
CHART_BAR_COLOURS = [
    "#4F67FF",
    "#F05A3A",
    "#17C497",
]
CHART_BAR_WIDTH = 0.42
CHART_BARGAP = 0.55
CHART_SHOW_LEGEND = False
CHART_MARGIN = dict(l=20, r=20, t=50, b=20)
CHART_GRID_COLOUR = "#E6E6E6"
CHART_AXIS_LINE_COLOUR = "#BFBFBF"
CHART_FONT_COLOUR = "#333333"

# -----------------------------------------------------------------------------
# Headline results card styling controls
# -----------------------------------------------------------------------------
SUMMARY_CARD_BACKGROUND = "#FFFFFF"
SUMMARY_CARD_BORDER_COLOUR = "#E6E6E6"
SUMMARY_CARD_BORDER_RADIUS = "12px"
SUMMARY_CARD_PADDING = "16px 18px 14px 18px"
SUMMARY_CARD_MIN_HEIGHT = "110px"
SUMMARY_CARD_SHADOW = "0 1px 3px rgba(0,0,0,0.06)"

SUMMARY_LABEL_FONT_SIZE = "14px"
SUMMARY_LABEL_FONT_WEIGHT = "500"
SUMMARY_LABEL_COLOUR = "#666666"
SUMMARY_LABEL_MARGIN_BOTTOM = "8px"

SUMMARY_VALUE_FONT_SIZE = "42px"
SUMMARY_VALUE_FONT_WEIGHT = "700"
SUMMARY_VALUE_COLOUR = "#222222"
SUMMARY_VALUE_LINE_HEIGHT = "1.05"

SUMMARY_UNIT_FONT_SIZE = "0.72em"
SUMMARY_UNIT_FONT_WEIGHT = "600"
SUMMARY_UNIT_COLOUR = "#444444"

SUMMARY_TEXT_ALIGN = "center"
SUMMARY_VERTICAL_ALIGN = "flex-start"
SUMMARY_ROW_GAP_TOP = "4px"
SUMMARY_ROW_GAP_BOTTOM = "10px"

# -----------------------------------------------------------------------------
# Divider line
# -----------------------------------------------------------------------------
DIVIDER_COLOUR = "#D9D9D9"
DIVIDER_WIDTH_PX = 1
DIVIDER_MIN_HEIGHT_PX = 900

# -----------------------------------------------------------------------------
# Logo
# -----------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = APP_DIR / "resources"
EPW_DIRECTORY = RESOURCE_DIR / "epw"
LOGO_PATH = RESOURCE_DIR / "Etude-logo-animation-v005 Single spin.gif"
LOGO_WIDTH = 180


@dataclass
class RoofGeometry:
    slope_depth_m: float
    gross_area_per_slope_m2: float
    usable_area_per_slope_m2: float
    total_gross_area_m2: float
    total_usable_area_m2: float
    portrait_modules_max: int
    landscape_modules_max: int
    better_layout: str
    better_module_count: int
    better_area_m2: float


def get_base64_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


def safe_floor(value: float) -> int:
    return max(int(math.floor(value)), 0)


def render_summary_card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div style="
            background: {SUMMARY_CARD_BACKGROUND};
            border: 1px solid {SUMMARY_CARD_BORDER_COLOUR};
            border-radius: {SUMMARY_CARD_BORDER_RADIUS};
            padding: {SUMMARY_CARD_PADDING};
            min-height: {SUMMARY_CARD_MIN_HEIGHT};
            box-shadow: {SUMMARY_CARD_SHADOW};
            display: flex;
            flex-direction: column;
            justify-content: {SUMMARY_VERTICAL_ALIGN};
            text-align: {SUMMARY_TEXT_ALIGN};
        ">
            <div style="
                font-size: {SUMMARY_LABEL_FONT_SIZE};
                font-weight: {SUMMARY_LABEL_FONT_WEIGHT};
                color: {SUMMARY_LABEL_COLOUR};
                margin-bottom: {SUMMARY_LABEL_MARGIN_BOTTOM};
            ">
                {label}
            </div>
            <div style="
                font-size: {SUMMARY_VALUE_FONT_SIZE};
                font-weight: {SUMMARY_VALUE_FONT_WEIGHT};
                color: {SUMMARY_VALUE_COLOUR};
                line-height: {SUMMARY_VALUE_LINE_HEIGHT};
            ">
                {value}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def calc_party_wall_count(house_form: str) -> int:
    if house_form == "Detached":
        return 0
    if house_form in {"Semi-detached", "End terrace"}:
        return 1
    if house_form == "Mid terrace":
        return 2
    return 0


def calc_spreadsheet_gfa(
    width_parallel_to_ridge_m: float,
    depth_perpendicular_to_ridge_m: float,
    wall_thickness_m: float,
    house_form: str,
) -> float:
    if house_form == "Detached":
        width_deduction = 2.0 * wall_thickness_m
    elif house_form == "Semi-detached":
        width_deduction = 1.5 * wall_thickness_m
    else:
        width_deduction = 1.0 * wall_thickness_m

    usable_width = max(width_parallel_to_ridge_m - width_deduction, 0.0)
    usable_depth = max(depth_perpendicular_to_ridge_m - 2.0 * wall_thickness_m, 0.0)
    return usable_width * usable_depth


def count_modules_fit(
    usable_x_m: float,
    usable_y_m: float,
    module_x_m: float,
    module_y_m: float,
    gap_x_m: float,
    gap_y_m: float,
) -> int:
    count_x = safe_floor((usable_x_m + gap_x_m) / (module_x_m + gap_x_m))
    count_y = safe_floor((usable_y_m + gap_y_m) / (module_y_m + gap_y_m))
    return count_x * count_y


def calc_pitched_roof_geometry(
    roof_form: str,
    length_m: float,
    width_m: float,
    pitch_deg: float,
    module_length_m: float,
    module_width_m: float,
    excluded_area_total_m2: float,
    perimeter_margin_m: float,
    gap_x_m: float,
    gap_y_m: float,
) -> RoofGeometry:
    pitch_rad = math.radians(pitch_deg)

    if roof_form == "Mono-pitch":
        slope_depth = width_m / max(math.cos(pitch_rad), 1e-6)
        slopes = 1
    else:
        slope_depth = (width_m / 2.0) / max(math.cos(pitch_rad), 1e-6)
        slopes = 2

    gross_area_per_slope = length_m * slope_depth
    usable_length = max(length_m - 2.0 * perimeter_margin_m, 0.0)
    usable_depth = max(slope_depth - 2.0 * perimeter_margin_m, 0.0)

    usable_area_per_slope_pre_exclusions = usable_length * usable_depth
    excluded_area_per_slope = excluded_area_total_m2 / max(slopes, 1)
    usable_area_per_slope = max(usable_area_per_slope_pre_exclusions - excluded_area_per_slope, 0.0)

    portrait_count_per_slope = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_width_m,
        module_y_m=module_length_m,
        gap_x_m=gap_x_m,
        gap_y_m=gap_y_m,
    )
    landscape_count_per_slope = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_length_m,
        module_y_m=module_width_m,
        gap_x_m=gap_x_m,
        gap_y_m=gap_y_m,
    )

    total_portrait = portrait_count_per_slope * slopes
    total_landscape = landscape_count_per_slope * slopes
    better_layout = "Portrait" if total_portrait >= total_landscape else "Landscape"
    better_modules = max(total_portrait, total_landscape)
    better_area = better_modules * module_length_m * module_width_m

    return RoofGeometry(
        slope_depth_m=slope_depth,
        gross_area_per_slope_m2=gross_area_per_slope,
        usable_area_per_slope_m2=usable_area_per_slope,
        total_gross_area_m2=gross_area_per_slope * slopes,
        total_usable_area_m2=usable_area_per_slope * slopes,
        portrait_modules_max=total_portrait,
        landscape_modules_max=total_landscape,
        better_layout=better_layout,
        better_module_count=better_modules,
        better_area_m2=better_area,
    )


def calc_flat_roof_geometry(
    length_m: float,
    width_m: float,
    module_length_m: float,
    module_width_m: float,
    excluded_area_total_m2: float,
    perimeter_margin_m: float,
    flat_layout: str,
) -> RoofGeometry:
    usable_length = max(length_m - 2.0 * perimeter_margin_m, 0.0)
    usable_depth = max(width_m - 2.0 * perimeter_margin_m, 0.0)
    gross_area = length_m * width_m
    usable_area_pre_exclusions = usable_length * usable_depth
    usable_area = max(usable_area_pre_exclusions - excluded_area_total_m2, 0.0)

    if flat_layout == "Front/back facing":
        portrait_gap_x = PITCHED_INTER_MODULE_GAP_M
        portrait_gap_y = FLAT_INTER_MODULE_GAP_M
        landscape_gap_x = PITCHED_INTER_MODULE_GAP_M
        landscape_gap_y = FLAT_INTER_MODULE_GAP_M
    else:
        portrait_gap_x = FLAT_INTER_MODULE_GAP_M
        portrait_gap_y = PITCHED_INTER_MODULE_GAP_M
        landscape_gap_x = FLAT_INTER_MODULE_GAP_M
        landscape_gap_y = PITCHED_INTER_MODULE_GAP_M

    portrait_modules = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_width_m,
        module_y_m=module_length_m,
        gap_x_m=portrait_gap_x,
        gap_y_m=portrait_gap_y,
    )
    landscape_modules = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_length_m,
        module_y_m=module_width_m,
        gap_x_m=landscape_gap_x,
        gap_y_m=landscape_gap_y,
    )

    better_layout = "Portrait" if portrait_modules >= landscape_modules else "Landscape"
    better_modules = max(portrait_modules, landscape_modules)
    better_area = better_modules * module_length_m * module_width_m

    return RoofGeometry(
        slope_depth_m=width_m,
        gross_area_per_slope_m2=gross_area,
        usable_area_per_slope_m2=usable_area,
        total_gross_area_m2=gross_area,
        total_usable_area_m2=usable_area,
        portrait_modules_max=portrait_modules,
        landscape_modules_max=landscape_modules,
        better_layout=better_layout,
        better_module_count=better_modules,
        better_area_m2=better_area,
    )


def calc_hipped_roof_detailed(
    length_m: float,
    width_m: float,
    pitch_deg: float,
    house_form: str,
    module_length_m: float,
    module_width_m: float,
    excluded_area_total_m2: float,
    ridge_offset_m: float,
    edge_offset_m: float,
    party_wall_offset_m: float,
) -> RoofGeometry:
    pitch_rad = math.radians(pitch_deg)
    half_span = width_m / 2.0
    slope_depth = half_span / max(math.cos(pitch_rad), 1e-6)
    gross_area_per_slope = length_m * slope_depth

    party_walls = calc_party_wall_count(house_form)
    effective_hip_reduction = min(width_m * 0.25, 2.0)
    usable_length = max(
        length_m - (2 * edge_offset_m) - effective_hip_reduction - (party_walls * party_wall_offset_m * 0.5),
        0.0,
    )
    usable_depth = max(slope_depth - ridge_offset_m - edge_offset_m, 0.0)
    raw_usable_area = usable_length * usable_depth * 0.92
    usable_area_per_slope = max(raw_usable_area - (excluded_area_total_m2 / 2.0), 0.0)

    portrait_count_per_slope = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_width_m,
        module_y_m=module_length_m,
        gap_x_m=PITCHED_INTER_MODULE_GAP_M,
        gap_y_m=PITCHED_INTER_MODULE_GAP_M,
    )
    landscape_count_per_slope = count_modules_fit(
        usable_x_m=usable_length,
        usable_y_m=usable_depth,
        module_x_m=module_length_m,
        module_y_m=module_width_m,
        gap_x_m=PITCHED_INTER_MODULE_GAP_M,
        gap_y_m=PITCHED_INTER_MODULE_GAP_M,
    )

    total_portrait = portrait_count_per_slope * 2
    total_landscape = landscape_count_per_slope * 2
    better_layout = "Portrait" if total_portrait >= total_landscape else "Landscape"
    better_modules = max(total_portrait, total_landscape)
    better_area = better_modules * module_length_m * module_width_m

    return RoofGeometry(
        slope_depth_m=slope_depth,
        gross_area_per_slope_m2=gross_area_per_slope,
        usable_area_per_slope_m2=usable_area_per_slope,
        total_gross_area_m2=gross_area_per_slope * 2,
        total_usable_area_m2=usable_area_per_slope * 2,
        portrait_modules_max=total_portrait,
        landscape_modules_max=total_landscape,
        better_layout=better_layout,
        better_module_count=better_modules,
        better_area_m2=better_area,
    )


def build_target_chart(
    required_area_m2: float,
    usable_area_m2: float,
    packed_area_m2: float,
) -> go.Figure:
    fig = go.Figure()

    categories = [
        "Part L required PV area",
        "Usable roof area",
        "Packed module area",
    ]
    values = [
        required_area_m2,
        usable_area_m2,
        packed_area_m2,
    ]

    fig.add_bar(
        x=categories,
        y=values,
        width=[CHART_BAR_WIDTH] * len(categories),
        marker_color=CHART_BAR_COLOURS,
        name="Area",
    )

    fig.update_layout(
        height=CHART_HEIGHT,
        margin=CHART_MARGIN,
        yaxis_title="m²",
        title=CHART_TITLE,
        showlegend=CHART_SHOW_LEGEND,
        bargap=CHART_BARGAP,
        paper_bgcolor=CHART_BACKGROUND_COLOUR,
        plot_bgcolor=CHART_PLOT_BACKGROUND_COLOUR,
        font=dict(color=CHART_FONT_COLOUR),
    )
    fig.update_xaxes(showgrid=False, linecolor=CHART_AXIS_LINE_COLOUR)
    fig.update_yaxes(gridcolor=CHART_GRID_COLOUR, linecolor=CHART_AXIS_LINE_COLOUR)

    return fig


def format_pass_fail(actual: float, target: float) -> str:
    return "Pass" if actual >= target else "Shortfall"


def get_available_epw_files(epw_dir: Path) -> dict[str, Path]:
    if not epw_dir.exists():
        return {}
    epw_files = sorted(epw_dir.glob("*.epw"))
    return {f.stem.replace("_", " "): f for f in epw_files}


def get_pysam_tilt_deg(roof_form: str, roof_pitch_deg: float) -> float:
    if roof_form == "Flat":
        return 10.0
    return float(roof_pitch_deg)


def run_pysam_pvwatts(
    system_capacity_kw: float,
    weather_file: Path,
    tilt_deg: float,
    azimuth_deg: float,
    losses_pct: float = PYSAM_SYSTEM_LOSSES_PCT,
    dc_ac_ratio: float = PYSAM_DC_AC_RATIO,
    array_type: int = PYSAM_ARRAY_TYPE_FIXED_ROOF,
    module_type: int = PYSAM_MODULE_TYPE_STANDARD,
    gcr: float = PYSAM_GCR,
) -> dict:
    model = pvwatts.new()

    model.SolarResource.solar_resource_file = str(weather_file)

    model.SystemDesign.system_capacity = float(system_capacity_kw)
    model.SystemDesign.tilt = float(tilt_deg)
    model.SystemDesign.azimuth = float(azimuth_deg)
    model.SystemDesign.losses = float(losses_pct)
    model.SystemDesign.dc_ac_ratio = float(dc_ac_ratio)
    model.SystemDesign.array_type = int(array_type)
    model.SystemDesign.module_type = int(module_type)
    model.SystemDesign.gcr = float(gcr)

    model.execute()

    return {
        "annual_ac_kwh": float(model.Outputs.ac_annual),
        "capacity_factor_pct": float(model.Outputs.capacity_factor),
        "monthly_ac_kwh": list(model.Outputs.ac_monthly),
    }


def format_module_length_label(length_m: float) -> str:
    return f"{length_m * 1000:.0f} mm"


header_left, header_right = st.columns([5, 1.5])

with header_left:
    st.title("Etude Part L 2026 PV requirement calculator")

with header_right:
    if LOGO_PATH.exists():
        st.markdown(
            f"""
            <div style="display:flex; justify-content:flex-end; align-items:flex-start; margin-top:8px;">
                <img src="data:image/gif;base64,{get_base64_image(LOGO_PATH)}" width="{LOGO_WIDTH}">
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div style="
                height: 90px;
                border: 1px dashed #bdbdbd;
                border-radius: 8px;
                display: flex;
                align-items: center;
                justify-content: center;
                color: #7a7a7a;
                font-size: 14px;
                margin-top: 8px;
            ">
                Logo not found
            </div>
            """,
            unsafe_allow_html=True,
        )

metrics_placeholder = st.container()

st.markdown("### Part L compliance inputs")
row1 = st.columns(5)
with row1[0]:
    house_form = st.selectbox(
        "House form",
        ["Detached", "Semi-detached", "End terrace", "Mid terrace"],
        index=0,
    )
with row1[1]:
    roof_form = st.selectbox(
        "Roof form",
        ["Mono-pitch", "Duo-pitch", "Hipped", "Flat"],
        index=1,
    )
with row1[2]:
    gfa_input_mode = st.selectbox(
        "Ground floor area method",
        ["Enter explicitly", "Derive from geometry (spreadsheet-style)"],
        index=0,
    )
with row1[3]:
    offset_mode = st.selectbox(
        "Roof margin method",
        ["Spreadsheet-style perimeter margin", "Detailed offsets (app-only)"],
        index=0,
    )
with row1[4]:
    sap_compliance_region = st.selectbox(
        "SAP compliance region",
        list(SAP_PLACEHOLDER_SPECIFIC_YIELD.keys()),
        index=4,
    )

st.caption(SAP_PLACEHOLDER_NOTE)

left, divider, right = st.columns([1.2, 0.03, 1.0])

with left:
    st.markdown("")

    if gfa_input_mode == "Enter explicitly":
        ground_floor_area_m2 = st.slider(
            "Ground floor area (m²)",
            min_value=20.0,
            max_value=500.0,
            value=72.0,
            step=1.0,
        )
        gfa_source_text = "Entered explicitly"
    else:
        st.markdown("**Ground floor area derivation**")
        ridge_parallel_width_for_gfa_m = st.slider(
            "Width parallel to ridge / long side (m)",
            min_value=4.0,
            max_value=25.0,
            value=9.0,
            step=0.1,
        )
        depth_for_gfa_m = st.slider(
            "Depth perpendicular to ridge / short side (m)",
            min_value=4.0,
            max_value=25.0,
            value=8.0,
            step=0.1,
        )
        wall_thickness_m = st.slider(
            "External wall thickness (m)",
            min_value=0.10,
            max_value=0.60,
            value=0.30,
            step=0.01,
        )
        ground_floor_area_m2 = calc_spreadsheet_gfa(
            width_parallel_to_ridge_m=ridge_parallel_width_for_gfa_m,
            depth_perpendicular_to_ridge_m=depth_for_gfa_m,
            wall_thickness_m=wall_thickness_m,
            house_form=house_form,
        )
        gfa_source_text = "Derived from geometry"

    st.markdown("**Roof geometry**")
    roof_plan_length_m = st.slider(
        "Roof plan length / width parallel to ridge (m)",
        min_value=5.0,
        max_value=20.0,
        value=9.0,
        step=0.1,
    )
    roof_plan_width_m = st.slider(
        "Roof plan width / depth perpendicular to ridge (m)",
        min_value=4.0,
        max_value=15.0,
        value=8.0,
        step=0.1,
    )

    if roof_form in {"Mono-pitch", "Duo-pitch", "Hipped"}:
        roof_pitch_deg = st.slider(
            "Roof pitch (degrees)",
            min_value=10,
            max_value=60,
            value=35,
            step=1,
        )
    else:
        roof_pitch_deg = 0

    excluded_area_total_m2 = st.slider(
        "Roof area blocked by windows / vents / plant etc. (total m²)",
        min_value=0.0,
        max_value=25.0,
        value=2.0,
        step=0.1,
    )

    if roof_form == "Flat":
        flat_layout = st.selectbox(
            "Flat roof arrangement",
            ["Front/back facing", "Facing either side"],
            index=0,
        )
    else:
        flat_layout = None

    st.markdown("**PV modules**")
    module_length_m = st.select_slider(
        "Module length (m)",
        options=MODULE_LENGTH_OPTIONS_M,
        value=DEFAULT_MODULE_LENGTH_M,
        format_func=lambda x: format_module_length_label(x),
    )
    module_width_m = FIXED_MODULE_WIDTH_M
    module_efficiency_pct = st.slider(
        "Module efficiency (%)",
        min_value=MIN_MODULE_EFFICIENCY_PCT,
        max_value=MAX_MODULE_EFFICIENCY_PCT,
        value=DEFAULT_MODULE_EFFICIENCY_PCT,
        step=MODULE_EFFICIENCY_STEP_PCT,
    )

with divider:
    st.markdown(
        f"""
        <div style="
            border-left: {DIVIDER_WIDTH_PX}px solid {DIVIDER_COLOUR};
            height: 100%;
            min-height: {DIVIDER_MIN_HEIGHT_PX}px;
            margin: 0 auto;
        "></div>
        """,
        unsafe_allow_html=True,
    )

with right:
    st.markdown("")
    chart_placeholder = st.empty()

if offset_mode == "Spreadsheet-style perimeter margin":
    with st.expander("Perimeter margin assumptions", expanded=False):
        st.caption("Spreadsheet-style mode uses a single perimeter margin around the PV array.")
        perimeter_margin_m = st.number_input(
            "Perimeter margin around PV array (m)",
            min_value=0.0,
            max_value=2.0,
            value=DEFAULT_SPREADSHEET_MARGIN_M,
            step=0.05,
        )
    ridge_offset_m = None
    edge_offset_m = None
    party_wall_offset_m = None
else:
    with st.expander("Detailed offset assumptions (app-only)", expanded=False):
        st.caption("Detailed offsets are useful for custom layouts. Hipped roofs use this mode most naturally.")
        offset_cols = st.columns(3)
        with offset_cols[0]:
            ridge_offset_m = st.number_input(
                "Ridge offset (m)",
                min_value=0.0,
                max_value=2.0,
                value=DEFAULT_OFFSETS_M["ridge"],
                step=0.05,
            )
        with offset_cols[1]:
            edge_offset_m = st.number_input(
                "Roof edge offset (m)",
                min_value=0.0,
                max_value=2.0,
                value=DEFAULT_OFFSETS_M["roof_edge"],
                step=0.05,
            )
        with offset_cols[2]:
            party_wall_offset_m = st.number_input(
                "Party wall offset (m)",
                min_value=0.0,
                max_value=2.0,
                value=DEFAULT_OFFSETS_M["party_wall"],
                step=0.05,
            )
    perimeter_margin_m = None

required_pv_area_m2 = ground_floor_area_m2 * FHS_REQUIRED_AREA_FRACTION
reference_required_kwp = required_pv_area_m2 * STANDARD_PANEL_EFFICIENCY_KWP_PER_M2

module_area_m2 = module_length_m * module_width_m
module_efficiency_fraction = module_efficiency_pct / 100.0
module_power_kwp = module_area_m2 * module_efficiency_fraction
module_power_wp = module_power_kwp * 1000.0
module_power_density_kwp_per_m2 = module_power_kwp / module_area_m2
modules_required = math.ceil(reference_required_kwp / module_power_kwp) if module_power_kwp > 0 else 0

if roof_form == "Hipped":
    if offset_mode == "Spreadsheet-style perimeter margin":
        geometry = calc_hipped_roof_detailed(
            length_m=roof_plan_length_m,
            width_m=roof_plan_width_m,
            pitch_deg=roof_pitch_deg,
            house_form=house_form,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            excluded_area_total_m2=excluded_area_total_m2,
            ridge_offset_m=perimeter_margin_m,
            edge_offset_m=perimeter_margin_m,
            party_wall_offset_m=perimeter_margin_m,
        )
    else:
        geometry = calc_hipped_roof_detailed(
            length_m=roof_plan_length_m,
            width_m=roof_plan_width_m,
            pitch_deg=roof_pitch_deg,
            house_form=house_form,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            excluded_area_total_m2=excluded_area_total_m2,
            ridge_offset_m=ridge_offset_m,
            edge_offset_m=edge_offset_m,
            party_wall_offset_m=party_wall_offset_m,
        )
elif roof_form in {"Mono-pitch", "Duo-pitch"}:
    if offset_mode == "Spreadsheet-style perimeter margin":
        geometry = calc_pitched_roof_geometry(
            roof_form=roof_form,
            length_m=roof_plan_length_m,
            width_m=roof_plan_width_m,
            pitch_deg=roof_pitch_deg,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            excluded_area_total_m2=excluded_area_total_m2,
            perimeter_margin_m=perimeter_margin_m,
            gap_x_m=PITCHED_INTER_MODULE_GAP_M,
            gap_y_m=PITCHED_INTER_MODULE_GAP_M,
        )
    else:
        approx_margin = max(
            edge_offset_m,
            ridge_offset_m,
            party_wall_offset_m if calc_party_wall_count(house_form) > 0 else 0.0,
        )
        geometry = calc_pitched_roof_geometry(
            roof_form=roof_form,
            length_m=roof_plan_length_m,
            width_m=roof_plan_width_m,
            pitch_deg=roof_pitch_deg,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            excluded_area_total_m2=excluded_area_total_m2,
            perimeter_margin_m=approx_margin,
            gap_x_m=PITCHED_INTER_MODULE_GAP_M,
            gap_y_m=PITCHED_INTER_MODULE_GAP_M,
        )
else:
    if offset_mode == "Spreadsheet-style perimeter margin":
        flat_margin = perimeter_margin_m
    else:
        flat_margin = edge_offset_m

    geometry = calc_flat_roof_geometry(
        length_m=roof_plan_length_m,
        width_m=roof_plan_width_m,
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        excluded_area_total_m2=excluded_area_total_m2,
        perimeter_margin_m=flat_margin,
        flat_layout=flat_layout,
    )

achievable_kwp_from_packed_modules = geometry.better_module_count * module_power_kwp
achievable_kwp_from_usable_area = geometry.total_usable_area_m2 * module_power_density_kwp_per_m2

sap_placeholder_specific_yield = SAP_PLACEHOLDER_SPECIFIC_YIELD[sap_compliance_region]
part_l_required_generation_kwh = reference_required_kwp * sap_placeholder_specific_yield
part_l_proposed_generation_kwh = achievable_kwp_from_packed_modules * sap_placeholder_specific_yield

area_based_status = format_pass_fail(geometry.better_area_m2, required_pv_area_m2)
kwp_based_status = format_pass_fail(achievable_kwp_from_packed_modules, reference_required_kwp)
modules_based_status = "Pass" if geometry.better_module_count >= modules_required else "Shortfall"
generation_based_status = format_pass_fail(part_l_proposed_generation_kwh, part_l_required_generation_kwh)

with metrics_placeholder:
    st.markdown(f"<div style='height:{SUMMARY_ROW_GAP_TOP};'></div>", unsafe_allow_html=True)

    top_col_1, top_col_2, top_col_3, top_col_4 = st.columns(4)

    with top_col_1:
        render_summary_card(
            "Part L required generation",
            f"{part_l_required_generation_kwh:,.0f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWh/a</span>",
        )
    with top_col_2:
        render_summary_card(
            "Actual building annual generation",
            f"{part_l_proposed_generation_kwh:,.0f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWh/a</span>",
        )
    with top_col_3:
        render_summary_card(
            "Modules required",
            f"{modules_required:,.0f}",
        )
    with top_col_4:
        render_summary_card(
            "Modules provided",
            f"{geometry.better_module_count:,.0f}",
        )

    st.markdown(f"<div style='height:{SUMMARY_ROW_GAP_BOTTOM};'></div>", unsafe_allow_html=True)

with chart_placeholder.container():
    st.plotly_chart(
        build_target_chart(
            required_area_m2=required_pv_area_m2,
            usable_area_m2=geometry.total_usable_area_m2,
            packed_area_m2=geometry.better_area_m2,
        ),
        theme=None,
        use_container_width=True,
    )

st.markdown("### Part L compliance results")
part_l_results_df = pd.DataFrame(
    [
        ("Ground floor area", f"{ground_floor_area_m2:,.1f} m²"),
        ("Ground floor area source", gfa_source_text),
        ("SAP compliance region", sap_compliance_region),
        ("SAP placeholder annual yield used", f"{sap_placeholder_specific_yield:,.0f} kWh/kWp·a"),
        ("Part L required PV area (40% of ground floor area)", f"{required_pv_area_m2:,.1f} m²"),
        ("Part L reference panel power density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:,.2f} kWp/m²"),
        ("Part L reference kWp", f"{reference_required_kwp:,.2f} kWp"),
        ("Part L required annual generation", f"{part_l_required_generation_kwh:,.0f} kWh/a"),
        ("Part L proposed annual generation", f"{part_l_proposed_generation_kwh:,.0f} kWh/a"),
        ("Generation check", generation_based_status),
        ("Module width used", f"{module_width_m * 1000:.0f} mm"),
        ("Module length used", f"{module_length_m * 1000:.0f} mm"),
        ("Module efficiency used", f"{module_efficiency_pct:,.1f} %"),
        ("Derived module power", f"{module_power_wp:,.0f} Wp"),
        ("Modules required", f"{modules_required}"),
        ("Relevant roof area (gross)", f"{geometry.total_gross_area_m2:,.1f} m²"),
        ("Relevant roof area (usable)", f"{geometry.total_usable_area_m2:,.1f} m²"),
        ("Packed module area", f"{geometry.better_area_m2:,.1f} m²"),
        ("Area check", area_based_status),
        ("Selected module power density", f"{module_power_density_kwp_per_m2:,.3f} kWp/m²"),
        ("Max modules - portrait", f"{geometry.portrait_modules_max}"),
        ("Max modules - landscape", f"{geometry.landscape_modules_max}"),
        ("Best packing layout", geometry.better_layout),
        ("Modules provided", f"{geometry.better_module_count}"),
        ("Module-count check", modules_based_status),
        ("Achievable packed kWp", f"{achievable_kwp_from_packed_modules:,.2f} kWp"),
        ("Achievable kWp from usable area", f"{achievable_kwp_from_usable_area:,.2f} kWp"),
        ("kWp check", kwp_based_status),
    ],
    columns=["Metric", "Value"],
)
st.dataframe(part_l_results_df, hide_index=True, use_container_width=True)

part_l_methodology_df = pd.DataFrame(
    [
        ("Methodology used for Part L section", "Spreadsheet-aligned roof-fit method plus SAP placeholder annual yield basis"),
        ("Part L reference basis", "40% of ground floor area and 0.22 kWp/m² reference panel density"),
        ("Compliance-region basis", "Temporary SAP region placeholder"),
        ("Radiation / annual yield source", "SAP placeholder values in code - needs checking"),
    ],
    columns=["Topic", "Description"],
)
st.dataframe(part_l_methodology_df, hide_index=True, use_container_width=True)

if (
    generation_based_status == "Shortfall"
    or modules_based_status == "Shortfall"
    or kwp_based_status == "Shortfall"
    or area_based_status == "Shortfall"
):
    st.warning(
        "This geometry appears short of the Part L reference benchmark under the current placeholder SAP compliance basis. "
        "The SAP irradiation values in this section still need to be checked."
    )
else:
    st.success(
        "This geometry appears capable of meeting the Part L reference benchmark under the current placeholder SAP compliance basis."
    )

st.divider()

st.markdown("## PySAM annual generation forecast")
st.caption("This section is secondary. It does not affect the Part L compliance calculation above.")

epw_lookup = get_available_epw_files(EPW_DIRECTORY)
epw_labels = ["None"] + list(epw_lookup.keys())

pysam_input_cols = st.columns(2)
with pysam_input_cols[0]:
    epw_label = st.selectbox("PV yield weather file (EPW)", epw_labels, index=0)
with pysam_input_cols[1]:
    azimuth_deg = st.slider(
        "PV azimuth for PySAM generation (degrees, 180 = south)",
        min_value=0,
        max_value=359,
        value=180,
        step=1,
    )

pysam_result = None
pysam_message = None
selected_epw = None

if epw_label == "None":
    pysam_message = "No EPW selected."
elif pvwatts is None:
    pysam_message = "PySAM is not installed in this environment."
else:
    selected_epw = epw_lookup[epw_label]
    pysam_tilt_deg = get_pysam_tilt_deg(roof_form=roof_form, roof_pitch_deg=roof_pitch_deg)

    try:
        pysam_result = run_pysam_pvwatts(
            system_capacity_kw=achievable_kwp_from_packed_modules,
            weather_file=selected_epw,
            tilt_deg=pysam_tilt_deg,
            azimuth_deg=azimuth_deg,
        )
    except Exception as exc:
        pysam_message = f"PySAM run failed: {exc}"

if pysam_result is not None:
    annual_gen_df = pd.DataFrame(
        [
            ("Generation methodology", "PySAM PVWatts v8"),
            ("Radiation / weather source", selected_epw.name if selected_epw else ""),
            ("Selected EPW", epw_label),
            ("PV azimuth", f"{azimuth_deg}°"),
            ("PV tilt used in PySAM", f"{get_pysam_tilt_deg(roof_form, roof_pitch_deg):.1f}°"),
            ("System capacity used", f"{achievable_kwp_from_packed_modules:,.2f} kWp"),
            ("Annual AC generation", f"{pysam_result['annual_ac_kwh']:,.0f} kWh/a"),
            ("Capacity factor", f"{pysam_result['capacity_factor_pct']:,.1f} %"),
        ],
        columns=["Metric", "Value"],
    )
    st.dataframe(annual_gen_df, hide_index=True, use_container_width=True)

    pysam_assumptions_df = pd.DataFrame(
        [
            ("Performance / system losses", f"{PYSAM_SYSTEM_LOSSES_PCT:.1f} %"),
            ("DC/AC ratio", f"{PYSAM_DC_AC_RATIO:.2f}"),
            ("Array type", "Fixed roof mount"),
            ("Module type", "Standard"),
            ("Ground coverage ratio", f"{PYSAM_GCR:.2f}"),
        ],
        columns=["Assumption", "Value"],
    )
    st.dataframe(pysam_assumptions_df, hide_index=True, use_container_width=True)

    monthly_df = pd.DataFrame(
        {
            "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            "AC generation (kWh)": pysam_result["monthly_ac_kwh"],
        }
    )
    st.dataframe(monthly_df, hide_index=True, use_container_width=True)
else:
    st.info(pysam_message or "Annual generation not available.")

st.markdown("### Method summary")
st.markdown(
    """
This version now separates the tool into two distinct parts.

The **top section** is the Part L 2026 compliance section. It uses the spreadsheet-aligned roof-fit method, fixed reference panel density of **0.22 kWp/m²**, and a PV area equal to **40% of ground floor area**. It also includes a **SAP compliance region** input and converts the reference system into an annual generation target in **kWh/year** using temporary SAP placeholder annual yield values. These placeholders still need to be checked against the approved methodology.

The roof-fit logic remains based on:
- roof form;
- building and roof geometry;
- module dimensions and efficiency;
- margin / offset assumptions; and
- portrait versus landscape packing.

The **bottom section** is a separate PySAM forecast. It uses PySAM PVWatts and a selected EPW weather file to estimate a more realistic annual output for the packed system. This forecast does **not** affect the Part L result above.

The module inputs are now constrained more tightly. Width is fixed at **1134 mm**. Length is chosen from a stepped list of typical module sizes. Module power is not entered directly. Instead, the tool derives module power from module area and selected efficiency.

Both sections now state their methodology, weather / radiation source and key assumptions more explicitly.
"""
)

with st.expander("Current limits", expanded=False):
    st.markdown(
        """
- Houses only in this version.
- Hipped roof logic remains app-specific rather than spreadsheet-equivalent.
- The spreadsheet-style ground floor area derivation uses end terrace and mid terrace under the same terraced rule.
- Blocked roof area is still handled as a simple area deduction rather than explicit obstacle geometry.
- The Part L section currently uses placeholder SAP annual yield values and these must be checked.
- The Part L section is still a simplified roof-fit / sizing tool rather than a full approved-methodology implementation.
- Annual generation in the PySAM section is secondary and does not form part of the Part L benchmark logic.
- EPW selection currently depends on locally stored weather files in `resources/epw/`.
"""
    )