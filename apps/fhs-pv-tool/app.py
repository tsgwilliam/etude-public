import math
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import base64


st.set_page_config(page_title="Etude PV Requirement Calculator", layout="wide")

STANDARD_PANEL_EFFICIENCY_KWP_PER_M2 = 0.22
FHS_REQUIRED_AREA_FRACTION = 0.40

DEFAULT_OFFSETS_M = {
    "ridge": 0.6,
    "roof_edge": 0.5,
    "party_wall": 0.75,
}

HEURISTIC_SPECIFIC_YIELD = {
    "Scotland / North": 800,
    "North England": 830,
    "Midlands": 860,
    "Wales": 850,
    "East England": 900,
    "South England": 910,
    "South West": 920,
    "London / South East": 930,
}

OVERSHADING_FACTOR = {
    "None": 1.00,
    "Light": 0.95,
    "Moderate": 0.88,
    "Heavy": 0.75,
}

# -----------------------------------------------------------------------------
# Graph styling controls
# Edit these directly in code as needed.
# -----------------------------------------------------------------------------
CHART_HEIGHT = 750
CHART_TITLE = "FHS target vs available PV area"
CHART_BACKGROUND_COLOUR = "white"
CHART_PLOT_BACKGROUND_COLOUR = "white"
CHART_BAR_COLOURS = [
    "#4F67FF",  # FHS required PV area
    "#F05A3A",  # Usable roof area
    "#17C497",  # Packed module area
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
# Edit these directly in code as needed.
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

SUMMARY_TEXT_ALIGN = "center"      # left / center / right
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
# LOGO
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


def calc_gable_roof(
    length_m: float,
    width_m: float,
    pitch_deg: float,
    house_form: str,
    module_length_m: float,
    module_width_m: float,
    excluded_area_per_slope_m2: float,
    ridge_offset_m: float,
    edge_offset_m: float,
    party_wall_offset_m: float,
) -> RoofGeometry:
    pitch_rad = math.radians(pitch_deg)
    half_span = width_m / 2.0
    slope_depth = half_span / max(math.cos(pitch_rad), 1e-6)
    gross_area_per_slope = length_m * slope_depth

    party_walls = calc_party_wall_count(house_form)
    usable_length = max(length_m - (2 * edge_offset_m) - (party_walls * party_wall_offset_m), 0.0)
    usable_depth = max(slope_depth - ridge_offset_m - edge_offset_m, 0.0)
    raw_usable_area = usable_length * usable_depth
    usable_area_per_slope = max(raw_usable_area - excluded_area_per_slope_m2, 0.0)

    portrait_count_per_slope = (
        safe_floor(usable_length / module_width_m) * safe_floor(usable_depth / module_length_m)
    )
    landscape_count_per_slope = (
        safe_floor(usable_length / module_length_m) * safe_floor(usable_depth / module_width_m)
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


def calc_hipped_roof(
    length_m: float,
    width_m: float,
    pitch_deg: float,
    house_form: str,
    module_length_m: float,
    module_width_m: float,
    excluded_area_per_slope_m2: float,
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
    usable_area_per_slope = max(raw_usable_area - excluded_area_per_slope_m2, 0.0)

    portrait_count_per_slope = (
        safe_floor(usable_length / module_width_m) * safe_floor(usable_depth / module_length_m)
    )
    landscape_count_per_slope = (
        safe_floor(usable_length / module_length_m) * safe_floor(usable_depth / module_width_m)
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
        "FHS required PV area",
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

st.markdown("")
row1 = st.columns(4)
with row1[0]:
    house_form = st.selectbox(
        "House form",
        ["Detached", "Semi-detached", "End terrace", "Mid terrace"],
        index=0,
    )
with row1[1]:
    roof_type = st.selectbox("Roof type", ["Gable", "Hipped"], index=0)
with row1[2]:
    region = st.selectbox(
        "Region",
        list(HEURISTIC_SPECIFIC_YIELD.keys()),
        index=7,
    )
with row1[3]:
    overshading = st.selectbox(
        "Overshading",
        list(OVERSHADING_FACTOR.keys()),
        index=0,
    )

left, divider, right = st.columns([1.2, 0.03, 1.0])

with left:
    st.markdown("")

    ground_floor_area_m2 = st.slider(
        "Ground floor area (m²)",
        min_value=20.0,
        max_value=500.0,
        value=72.0,
        step=1.0,
    )

    st.markdown("**Roof geometry**")
    roof_plan_length_m = st.slider(
        "Roof plan length (m)",
        min_value=5.0,
        max_value=20.0,
        value=9.0,
        step=0.1,
    )
    roof_plan_width_m = st.slider(
        "Roof plan width (m)",
        min_value=4.0,
        max_value=15.0,
        value=8.0,
        step=0.1,
    )
    roof_pitch_deg = st.slider(
        "Roof pitch (degrees)",
        min_value=10,
        max_value=60,
        value=35,
        step=1,
    )
    excluded_area_total_m2 = st.slider(
        "Roof area blocked by windows / vents / plant etc. (total m²)",
        min_value=0.0,
        max_value=25.0,
        value=2.0,
        step=0.1,
    )

    st.markdown("**PV modules**")
    module_length_m = st.slider(
        "Module length (m)",
        min_value=1.0,
        max_value=3.0,
        value=1.72,
        step=0.01,
    )
    module_width_m = st.slider(
        "Module width (m)",
        min_value=0.5,
        max_value=1.5,
        value=1.13,
        step=0.01,
    )
    module_power_wp = st.slider(
        "Module power (Wp)",
        min_value=250,
        max_value=800,
        value=450,
        step=5,
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

required_pv_area_m2 = ground_floor_area_m2 * FHS_REQUIRED_AREA_FRACTION
reference_required_kwp = required_pv_area_m2 * STANDARD_PANEL_EFFICIENCY_KWP_PER_M2

with st.expander("Advanced offset assumptions (optional)", expanded=False):
    st.caption("These values update the results above.")
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

module_area_m2 = module_length_m * module_width_m
module_power_kwp = module_power_wp / 1000.0
module_power_density_kwp_per_m2 = module_power_kwp / module_area_m2
excluded_area_per_slope_m2 = excluded_area_total_m2 / 2.0

if roof_type == "Gable":
    geometry = calc_gable_roof(
        length_m=roof_plan_length_m,
        width_m=roof_plan_width_m,
        pitch_deg=roof_pitch_deg,
        house_form=house_form,
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        excluded_area_per_slope_m2=excluded_area_per_slope_m2,
        ridge_offset_m=ridge_offset_m,
        edge_offset_m=edge_offset_m,
        party_wall_offset_m=party_wall_offset_m,
    )
else:
    geometry = calc_hipped_roof(
        length_m=roof_plan_length_m,
        width_m=roof_plan_width_m,
        pitch_deg=roof_pitch_deg,
        house_form=house_form,
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        excluded_area_per_slope_m2=excluded_area_per_slope_m2,
        ridge_offset_m=ridge_offset_m,
        edge_offset_m=edge_offset_m,
        party_wall_offset_m=party_wall_offset_m,
    )

achievable_kwp_from_packed_modules = geometry.better_module_count * module_power_kwp
achievable_kwp_from_usable_area = geometry.total_usable_area_m2 * module_power_density_kwp_per_m2

specific_yield_kwh_per_kwp = HEURISTIC_SPECIFIC_YIELD[region]
indicative_generation_kwh_per_year = (
    achievable_kwp_from_packed_modules
    * specific_yield_kwh_per_kwp
    * OVERSHADING_FACTOR[overshading]
)

area_based_status = format_pass_fail(geometry.better_area_m2, required_pv_area_m2)
kwp_based_status = format_pass_fail(achievable_kwp_from_packed_modules, reference_required_kwp)

with metrics_placeholder:
    st.markdown(f"<div style='height:{SUMMARY_ROW_GAP_TOP};'></div>", unsafe_allow_html=True)

    top_col_1, top_col_2, top_col_3, top_col_4 = st.columns(4)

    with top_col_1:
        render_summary_card("Ground floor area", f"{ground_floor_area_m2:,.1f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>m²</span>")

    with top_col_2:
        render_summary_card("FHS required PV area", f"{required_pv_area_m2:,.1f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>m²</span>")

    with top_col_3:
        render_summary_card("FHS reference kWp", f"{reference_required_kwp:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>")

    with top_col_4:
        render_summary_card("Achievable packed kWp", f"{achievable_kwp_from_packed_modules:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>")

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

st.markdown("### Results")
results_df = pd.DataFrame(
    [
        ("Ground floor area", f"{ground_floor_area_m2:,.1f} m²"),
        ("FHS required PV area (40% of ground floor area)", f"{required_pv_area_m2:,.1f} m²"),
        ("FHS reference panel power density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:,.2f} kWp/m²"),
        ("FHS reference kWp", f"{reference_required_kwp:,.2f} kWp"),
        ("Relevant roof slope area (gross)", f"{geometry.total_gross_area_m2:,.1f} m²"),
        ("Relevant roof slope area (usable)", f"{geometry.total_usable_area_m2:,.1f} m²"),
        ("Packed module area", f"{geometry.better_area_m2:,.1f} m²"),
        ("Area check", area_based_status),
        ("Selected module power density", f"{module_power_density_kwp_per_m2:,.3f} kWp/m²"),
        ("Max modules - portrait", f"{geometry.portrait_modules_max}"),
        ("Max modules - landscape", f"{geometry.landscape_modules_max}"),
        ("Best packing layout", geometry.better_layout),
        ("Recommended module count", f"{geometry.better_module_count}"),
        ("Achievable packed kWp", f"{achievable_kwp_from_packed_modules:,.2f} kWp"),
        ("Achievable kWp from usable area", f"{achievable_kwp_from_usable_area:,.2f} kWp"),
        ("Indicative annual generation", f"{indicative_generation_kwh_per_year:,.0f} kWh/a"),
        ("kWp check", kwp_based_status),
    ],
    columns=["Metric", "Value"],
)
st.dataframe(results_df, hide_index=True, use_container_width=True)

if area_based_status == "Shortfall" or kwp_based_status == "Shortfall":
    st.warning(
        "This geometry appears short of the reference benchmark. In practice that may point to a redesign, "
        "higher-output modules, a different roof arrangement, or an argument based on reasonably practicable roof area."
    )
else:
    st.success("This geometry appears capable of meeting the prototype benchmark.")

st.markdown("### Method summary")
st.markdown(
    """
This prototype is focused on the house route in Approved Document L 2026.

The benchmark part of the calculation is fixed. For a dwellinghouse, the reference PV array area is **40% of the ground floor area**. The reference panel power density is **0.22 kWp/m²**. That figure is prescribed in the published guidance for the reference comparison, so it is kept fixed here rather than treated as a user variable.

The app then carries out a separate roof-fit test. It uses the entered roof plan dimensions and pitch to estimate the roof slope area, applies deductions for ridge clearance, roof edges, party walls and blocked areas, and then checks how many modules fit in portrait and landscape.

The chart and results table compare:
- the **FHS required PV area**;
- the **usable roof area** after deductions; and
- the **packed module area** and achievable installed capacity based on the selected modules.

Region and overshading are retained only for the **indicative annual generation** output. They do not change the FHS benchmark itself.

This is useful because the policy target and the practical roof capacity are not the same thing. A house can have a large required PV target but a more limited usable roof area once layout constraints are allowed for.

This remains an early-stage advisory tool. It is not a formal compliance calculation and it does not replace the approved methodology.
"""
)

with st.expander("Current limits", expanded=False):
    st.markdown(
        """
- Houses only in this version.
- Gable and hipped roofs only.
- Hipped roof logic is simplified.
- The FHS benchmark is fixed, but the indicative annual generation output is heuristic only.
- Dormers, multiple roof facets, chimneys and complex setbacks are only approximated through the blocked-area input.
- Flats and buildings containing dwellings need a separate allocation logic and are not included yet.
"""
    )