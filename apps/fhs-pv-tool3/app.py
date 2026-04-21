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


st.set_page_config(page_title="Etude Part L 2026 PV requirement calculator", layout="wide")

# -----------------------------------------------------------------------------
# Part L reference assumptions
# -----------------------------------------------------------------------------
STANDARD_PANEL_EFFICIENCY_KWP_PER_M2 = 0.22
FHS_REQUIRED_AREA_FRACTION = 0.40

SAP_PLACEHOLDER_SPECIFIC_YIELD = {
    "England - North": 800,
    "England - Midlands": 860,
    "England - South": 920,
    "England - South West": 900,
    "England - London / South East": 930,
}
SAP_PLACEHOLDER_NOTE = (
    "SAP annual generation inputs in this tool are placeholders only and must be "
    "checked against the approved Part L / SAP methodology."
)

# -----------------------------------------------------------------------------
# Roof reduction assumptions for Section 2
# -----------------------------------------------------------------------------
DEFAULT_SIMPLE_PERIMETER_MARGIN_M = 0.30
DEFAULT_RIDGE_OFFSET_M = 0.60
DEFAULT_EDGE_OFFSET_M = 0.50
DEFAULT_PARTY_WALL_OFFSET_M = 0.75
DEFAULT_BLOCKED_AREA_M2 = 2.0

FLAT_ROOF_FIXED_TILT_DEG = 12.0

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

STANDARDISED_MODULE_LENGTH_M = 1.722
STANDARDISED_MODULE_WIDTH_M = 1.134
STANDARDISED_MODULE_EFFICIENCY_PCT = 22.0

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
# Styling
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

APP_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = APP_DIR / "resources"
EPW_DIRECTORY = RESOURCE_DIR / "epw"
LOGO_PATH = RESOURCE_DIR / "Etude-logo-animation-v005 Single spin.gif"
LOGO_WIDTH = 180

CHART_HEIGHT = 420
CHART_BACKGROUND_COLOUR = "white"
CHART_PLOT_BACKGROUND_COLOUR = "white"
CHART_FONT_COLOUR = "#333333"
CHART_GRID_COLOUR = "#E6E6E6"
CHART_AXIS_LINE_COLOUR = "#BFBFBF"
CHART_MARGIN = dict(l=20, r=20, t=40, b=20)
CHART_BAR_COLOURS = ["#4F67FF", "#F05A3A"]
CHART_BAR_WIDTH = 0.38
CHART_BARGAP = 0.45


@dataclass
class ArrayDefinition:
    name: str
    azimuth_deg: float
    tilt_deg: float
    area_share_fraction: float


def get_base64_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


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
            justify-content: flex-start;
            text-align: center;
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


def format_module_length_label(length_m: float) -> str:
    return f"{length_m * 1000:.0f} mm"


def module_power_kwp_from_inputs(length_m: float, width_m: float, efficiency_pct: float) -> float:
    return length_m * width_m * (efficiency_pct / 100.0)


def get_available_epw_files(epw_dir: Path) -> dict[str, Path]:
    if not epw_dir.exists():
        return {}
    epw_files = sorted(epw_dir.glob("*.epw"))
    return {f.stem.replace("_", " "): f for f in epw_files}


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


def sap_orientation_factor_placeholder(azimuth_deg: float) -> float:
    deviation = abs(azimuth_deg - 180.0)
    deviation = min(deviation, 360.0 - deviation)
    factor = 1.0 - 0.25 * (deviation / 180.0)
    return max(0.70, min(1.0, factor))


def sap_tilt_factor_placeholder(tilt_deg: float) -> float:
    deviation = abs(tilt_deg - 45.0)
    factor = 1.0 - 0.15 * min(deviation / 45.0, 1.0)
    return max(0.85, min(1.0, factor))


def allocate_integer_counts(total_count: int, share_fractions: list[float]) -> list[int]:
    if total_count <= 0:
        return [0] * len(share_fractions)

    raw = [total_count * share for share in share_fractions]
    base = [math.floor(v) for v in raw]
    remainder = total_count - sum(base)

    order = sorted(
        range(len(share_fractions)),
        key=lambda i: (raw[i] - base[i]),
        reverse=True,
    )

    for i in range(remainder):
        base[order[i % len(order)]] += 1

    return base


def build_comparison_chart(
    required_generation_kwh: float,
    actual_generation_kwh: float,
) -> go.Figure:
    fig = go.Figure()

    categories = ["Part L required", "Actual building"]
    values = [required_generation_kwh, actual_generation_kwh]

    fig.add_bar(
        x=categories,
        y=values,
        width=[CHART_BAR_WIDTH, CHART_BAR_WIDTH],
        marker_color=CHART_BAR_COLOURS,
        name="Annual generation",
    )

    fig.update_layout(
        height=CHART_HEIGHT,
        margin=CHART_MARGIN,
        yaxis_title="kWh/a",
        title="Part L annual generation comparison",
        showlegend=False,
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


def calc_party_wall_count(house_form: str) -> int:
    if house_form == "Detached":
        return 0
    if house_form in {"Semi-detached", "End terrace"}:
        return 1
    if house_form == "Mid terrace":
        return 2
    return 0


def count_modules_on_rectangle(
    usable_length_m: float,
    usable_depth_m: float,
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
) -> int:
    if mount_orientation == "Portrait":
        module_x_m = module_width_m
        module_y_m = module_length_m
    else:
        module_x_m = module_length_m
        module_y_m = module_width_m

    count_x = math.floor(usable_length_m / module_x_m) if module_x_m > 0 else 0
    count_y = math.floor(usable_depth_m / module_y_m) if module_y_m > 0 else 0
    return max(count_x, 0) * max(count_y, 0)


def build_section2_planes(
    roof_form: str,
    plan_length_along_ridge_m: float,
    plan_length_ridge_to_eaves_m: float,
    pitch_deg: float,
    azimuth_deg: float | None,
    blocked_area_total_m2: float,
    perimeter_margin_m: float,
    ridge_offset_m: float,
    edge_offset_m: float,
    party_wall_offset_m: float,
    house_form: str,
) -> list[dict]:
    blocked_area_total_m2 = max(blocked_area_total_m2, 0.0)
    planes: list[dict] = []

    if roof_form == "Flat":
        gross_length_m = plan_length_along_ridge_m
        gross_depth_m = plan_length_ridge_to_eaves_m / 2.0

        if perimeter_margin_m > 0:
            usable_length_m = max(gross_length_m - 2.0 * perimeter_margin_m, 0.0)
            usable_depth_m = max(gross_depth_m - 2.0 * perimeter_margin_m, 0.0)
        else:
            usable_length_m = max(gross_length_m - 2.0 * edge_offset_m, 0.0)
            usable_depth_m = max(gross_depth_m - 2.0 * edge_offset_m, 0.0)

        for name, plane_azimuth in [("East-facing array", 90.0), ("West-facing array", 270.0)]:
            planes.append(
                {
                    "name": name,
                    "azimuth_deg": plane_azimuth,
                    "tilt_deg": FLAT_ROOF_FIXED_TILT_DEG,
                    "gross_length_m": gross_length_m,
                    "gross_depth_m": gross_depth_m,
                    "usable_length_m": usable_length_m,
                    "usable_depth_m": usable_depth_m,
                    "gross_area_m2": gross_length_m * gross_depth_m,
                    "usable_area_before_blocked_m2": usable_length_m * usable_depth_m,
                }
            )

    else:
        pitch_rad = math.radians(pitch_deg)
        gross_length_m = plan_length_along_ridge_m
        gross_depth_m = plan_length_ridge_to_eaves_m / max(math.cos(pitch_rad), 1e-6)

        if perimeter_margin_m > 0:
            usable_length_m = max(gross_length_m - 2.0 * perimeter_margin_m, 0.0)
            usable_depth_m = max(gross_depth_m - 2.0 * perimeter_margin_m, 0.0)
        else:
            party_wall_count = calc_party_wall_count(house_form)
            usable_length_m = max(
                gross_length_m - (2.0 * edge_offset_m) - (party_wall_count * party_wall_offset_m),
                0.0,
            )
            usable_depth_m = max(gross_depth_m - ridge_offset_m - edge_offset_m, 0.0)

        if roof_form == "Mono-pitch":
            plane_azimuths = [float(azimuth_deg)]
            plane_names = ["Roof plane 1"]
        else:
            azimuth_1 = float(azimuth_deg)
            azimuth_2 = (azimuth_1 + 180.0) % 360.0
            plane_azimuths = [azimuth_1, azimuth_2]
            plane_names = ["Roof plane 1", "Roof plane 2"]

        for name, plane_azimuth in zip(plane_names, plane_azimuths):
            planes.append(
                {
                    "name": name,
                    "azimuth_deg": plane_azimuth,
                    "tilt_deg": float(pitch_deg),
                    "gross_length_m": gross_length_m,
                    "gross_depth_m": gross_depth_m,
                    "usable_length_m": usable_length_m,
                    "usable_depth_m": usable_depth_m,
                    "gross_area_m2": gross_length_m * gross_depth_m,
                    "usable_area_before_blocked_m2": usable_length_m * usable_depth_m,
                }
            )

    blocked_area_per_plane_m2 = blocked_area_total_m2 / max(len(planes), 1)
    for plane in planes:
        plane["blocked_area_m2"] = blocked_area_per_plane_m2
        plane["usable_area_after_blocked_m2"] = max(
            plane["usable_area_before_blocked_m2"] - blocked_area_per_plane_m2,
            0.0,
        )

    return planes


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

# -----------------------------------------------------------------------------
# 1) The Part L requirement
# -----------------------------------------------------------------------------
st.markdown("## 1) The Part L requirement")

part_l_top = st.columns(3)

with part_l_top[0]:
    house_form = st.selectbox(
        "House form",
        ["Detached", "Semi-detached", "End terrace", "Mid terrace"],
        index=0,
        key="part_l_house_form",
    )

with part_l_top[1]:
    gfa_input_mode = st.selectbox(
        "Ground floor area method",
        ["Enter explicitly", "Derive from geometry"],
        index=0,
        key="part_l_gfa_mode",
    )

with part_l_top[2]:
    sap_compliance_region = st.selectbox(
        "SAP compliance region",
        list(SAP_PLACEHOLDER_SPECIFIC_YIELD.keys()),
        index=4,
        key="sap_region",
    )

st.caption(SAP_PLACEHOLDER_NOTE)

if gfa_input_mode == "Enter explicitly":
    ground_floor_area_m2 = st.slider(
        "Ground floor area (m²)",
        min_value=20.00,
        max_value=500.00,
        value=72.00,
        step=0.01,
        key="part_l_gfa_direct",
    )
    gfa_source_text = "Entered explicitly"
else:
    gfa_geom_cols = st.columns(3)
    with gfa_geom_cols[0]:
        ridge_parallel_width_for_gfa_m = st.slider(
            "Width parallel to ridge / long side (m)",
            min_value=4.00,
            max_value=25.00,
            value=9.00,
            step=0.01,
            key="part_l_gfa_width",
        )
    with gfa_geom_cols[1]:
        depth_for_gfa_m = st.slider(
            "Depth perpendicular to ridge / short side (m)",
            min_value=4.00,
            max_value=25.00,
            value=8.00,
            step=0.01,
            key="part_l_gfa_depth",
        )
    with gfa_geom_cols[2]:
        wall_thickness_m = st.slider(
            "External wall thickness (m)",
            min_value=0.10,
            max_value=0.60,
            value=0.30,
            step=0.01,
            key="part_l_gfa_wall",
        )
    ground_floor_area_m2 = calc_spreadsheet_gfa(
        width_parallel_to_ridge_m=ridge_parallel_width_for_gfa_m,
        depth_perpendicular_to_ridge_m=depth_for_gfa_m,
        wall_thickness_m=wall_thickness_m,
        house_form=house_form,
    )
    gfa_source_text = "Derived from geometry"

sap_placeholder_specific_yield = SAP_PLACEHOLDER_SPECIFIC_YIELD[sap_compliance_region]

part_l_required_kwp = ground_floor_area_m2 * FHS_REQUIRED_AREA_FRACTION * STANDARD_PANEL_EFFICIENCY_KWP_PER_M2
part_l_required_generation_kwh = part_l_required_kwp * sap_placeholder_specific_yield

standardised_module_power_kwp = module_power_kwp_from_inputs(
    length_m=STANDARDISED_MODULE_LENGTH_M,
    width_m=STANDARDISED_MODULE_WIDTH_M,
    efficiency_pct=STANDARDISED_MODULE_EFFICIENCY_PCT,
)
part_l_required_panel_count = math.ceil(part_l_required_kwp / standardised_module_power_kwp)

part_l_summary_cols = st.columns(3)
with part_l_summary_cols[0]:
    render_summary_card(
        "Part L required generation",
        f"{part_l_required_generation_kwh:,.0f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWh/a</span>",
    )
with part_l_summary_cols[1]:
    render_summary_card(
        "Part L required kWp",
        f"{part_l_required_kwp:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>",
    )
with part_l_summary_cols[2]:
    render_summary_card(
        "Part L required panel count",
        f"{part_l_required_panel_count:,.0f}",
    )

part_l_requirement_df = pd.DataFrame(
    [
        ("Ground floor area", f"{ground_floor_area_m2:,.2f} m²"),
        ("Ground floor area source", gfa_source_text),
        ("SAP compliance region", sap_compliance_region),
        ("SAP placeholder annual yield used", f"{sap_placeholder_specific_yield:,.0f} kWh/kWp·a"),
        ("Part L required generation", f"{part_l_required_generation_kwh:,.0f} kWh/a"),
        ("Part L required kWp", f"{part_l_required_kwp:,.2f} kWp"),
        ("Standardised module width", f"{STANDARDISED_MODULE_WIDTH_M * 1000:.0f} mm"),
        ("Standardised module length", f"{STANDARDISED_MODULE_LENGTH_M * 1000:.0f} mm"),
        ("Standardised module efficiency", f"{STANDARDISED_MODULE_EFFICIENCY_PCT:,.1f} %"),
        ("Standardised module power", f"{standardised_module_power_kwp * 1000:,.0f} Wp"),
        ("Part L required panel count", f"{part_l_required_panel_count:,.0f}"),
    ],
    columns=["Metric", "Value"],
)
st.dataframe(part_l_requirement_df, hide_index=True, use_container_width=True)

st.divider()

# -----------------------------------------------------------------------------
# 2) Your Building measured by Part L
# -----------------------------------------------------------------------------
st.markdown("## 2) Your Building measured by Part L")

building_top = st.columns(2)
with building_top[0]:
    actual_roof_form = st.selectbox(
        "Roof type",
        ["Mono-pitch", "Duo-pitch", "Flat"],
        index=1,
        key="actual_roof_form",
    )
with building_top[1]:
    if actual_roof_form == "Flat":
        st.markdown("Flat roof uses one whole rectangular roof in plan.")
    elif actual_roof_form == "Mono-pitch":
        st.markdown("Mono-pitch uses one rectangular roof plane defined in plan.")
    else:
        st.markdown("Duo-pitch duplicates one entered roof plane and rotates the second by 180°.")

if actual_roof_form == "Flat":
    roof_geom_cols = st.columns(2)
    with roof_geom_cols[0]:
        plan_length_along_ridge_m = st.slider(
            "Whole roof length in plan (m)",
            min_value=2.00,
            max_value=40.00,
            value=10.00,
            step=0.01,
            key="flat_roof_length",
        )
    with roof_geom_cols[1]:
        plan_length_ridge_to_eaves_m = st.slider(
            "Whole roof width in plan (m)",
            min_value=2.00,
            max_value=40.00,
            value=8.00,
            step=0.01,
            key="flat_roof_width",
        )
    mono_or_duo_azimuth_deg = None
    mono_or_duo_pitch_deg = 0.0

    st.info(
        "Flat roof assumption: panels are pitched at 12° in back-to-back rows alternating east-west. "
        "TODO: row spacing still needs to be thought through."
    )

else:
    roof_geom_cols = st.columns(4)
    with roof_geom_cols[0]:
        plan_length_along_ridge_m = st.slider(
            "Ridge length in plan (m)",
            min_value=2.00,
            max_value=40.00,
            value=10.00,
            step=0.01,
            key="pitched_ridge_length",
        )
    with roof_geom_cols[1]:
        plan_length_ridge_to_eaves_m = st.slider(
            "Ridge-to-eaves length in plan (m)",
            min_value=1.00,
            max_value=20.00,
            value=4.00,
            step=0.01,
            key="pitched_eaves_length_plan",
        )
    with roof_geom_cols[2]:
        mono_or_duo_azimuth_deg = st.slider(
            "Roof plane azimuth (degrees, 180 = south)",
            min_value=0,
            max_value=359,
            value=180,
            step=1,
            key="pitched_azimuth",
        )
    with roof_geom_cols[3]:
        mono_or_duo_pitch_deg = st.slider(
            "Roof pitch (degrees)",
            min_value=1,
            max_value=60,
            value=35,
            step=1,
            key="pitched_pitch",
        )

st.markdown("**Roof reductions**")

reduction_top = st.columns(2)
with reduction_top[0]:
    blocked_area_total_m2 = st.slider(
        "Area blocked by windows / vents / plant etc. (m²)",
        min_value=0.0,
        max_value=25.0,
        value=DEFAULT_BLOCKED_AREA_M2,
        step=0.1,
        key="actual_blocked_area",
    )
with reduction_top[1]:
    offset_mode_section_2 = st.selectbox(
        "Roof reduction method",
        ["Simple perimeter margin", "Detailed offsets"],
        index=0,
        key="actual_offset_mode",
    )

if offset_mode_section_2 == "Simple perimeter margin":
    perimeter_margin_m = st.number_input(
        "Perimeter margin around PV zone (m)",
        min_value=0.0,
        max_value=2.0,
        value=DEFAULT_SIMPLE_PERIMETER_MARGIN_M,
        step=0.05,
        key="actual_perimeter_margin",
    )
    ridge_offset_m = 0.0
    edge_offset_m = 0.0
    party_wall_offset_m = 0.0
else:
    st.caption("Detailed offsets reduce usable roof dimensions in this simplified Section 2 method.")
    offset_cols = st.columns(3)
    with offset_cols[0]:
        ridge_offset_m = st.number_input(
            "Ridge offset (m)",
            min_value=0.0,
            max_value=2.0,
            value=DEFAULT_RIDGE_OFFSET_M,
            step=0.05,
            key="actual_ridge_offset",
        )
    with offset_cols[1]:
        edge_offset_m = st.number_input(
            "Roof edge offset (m)",
            min_value=0.0,
            max_value=2.0,
            value=DEFAULT_EDGE_OFFSET_M,
            step=0.05,
            key="actual_edge_offset",
        )
    with offset_cols[2]:
        party_wall_offset_m = st.number_input(
            "Party wall offset (m)",
            min_value=0.0,
            max_value=2.0,
            value=DEFAULT_PARTY_WALL_OFFSET_M,
            step=0.05,
            key="actual_party_wall_offset",
        )
    perimeter_margin_m = 0.0

roof_planes = build_section2_planes(
    roof_form=actual_roof_form,
    plan_length_along_ridge_m=plan_length_along_ridge_m,
    plan_length_ridge_to_eaves_m=plan_length_ridge_to_eaves_m,
    pitch_deg=mono_or_duo_pitch_deg,
    azimuth_deg=mono_or_duo_azimuth_deg,
    blocked_area_total_m2=blocked_area_total_m2,
    perimeter_margin_m=perimeter_margin_m,
    ridge_offset_m=ridge_offset_m,
    edge_offset_m=edge_offset_m,
    party_wall_offset_m=party_wall_offset_m,
    house_form=house_form,
)

module_cols = st.columns(4)
with module_cols[0]:
    module_length_m = st.select_slider(
        "Module length",
        options=MODULE_LENGTH_OPTIONS_M,
        value=DEFAULT_MODULE_LENGTH_M,
        format_func=lambda x: format_module_length_label(x),
        key="actual_module_length",
    )
with module_cols[1]:
    st.text_input(
        "Module width",
        value=f"{FIXED_MODULE_WIDTH_M * 1000:.0f} mm fixed",
        disabled=True,
        key="actual_module_width_display",
    )
    module_width_m = FIXED_MODULE_WIDTH_M
with module_cols[2]:
    module_efficiency_pct = st.slider(
        "Module efficiency (%)",
        min_value=MIN_MODULE_EFFICIENCY_PCT,
        max_value=MAX_MODULE_EFFICIENCY_PCT,
        value=DEFAULT_MODULE_EFFICIENCY_PCT,
        step=MODULE_EFFICIENCY_STEP_PCT,
        key="actual_module_eff",
    )
with module_cols[3]:
    module_mount_orientation = st.selectbox(
        "Mount orientation",
        ["Portrait", "Landscape"],
        index=0,
        key="mount_orientation",
    )

module_area_m2 = module_length_m * module_width_m
module_power_kwp = module_power_kwp_from_inputs(
    length_m=module_length_m,
    width_m=module_width_m,
    efficiency_pct=module_efficiency_pct,
)
module_power_wp = module_power_kwp * 1000.0

for plane in roof_planes:
    packed_by_dimensions = count_modules_on_rectangle(
        usable_length_m=plane["usable_length_m"],
        usable_depth_m=plane["usable_depth_m"],
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        mount_orientation=module_mount_orientation,
    )
    area_cap_panels = math.floor(plane["usable_area_after_blocked_m2"] / module_area_m2) if module_area_m2 > 0 else 0
    plane["max_feasible_panels"] = min(packed_by_dimensions, area_cap_panels)

max_feasible_panels = sum(plane["max_feasible_panels"] for plane in roof_planes)

if max_feasible_panels > 0:
    actual_arrays = [
        ArrayDefinition(
            name=plane["name"],
            azimuth_deg=plane["azimuth_deg"],
            tilt_deg=plane["tilt_deg"],
            area_share_fraction=plane["max_feasible_panels"] / max_feasible_panels,
        )
        for plane in roof_planes
    ]
else:
    equal_share = 1.0 / len(roof_planes) if roof_planes else 1.0
    actual_arrays = [
        ArrayDefinition(
            name=plane["name"],
            azimuth_deg=plane["azimuth_deg"],
            tilt_deg=plane["tilt_deg"],
            area_share_fraction=equal_share,
        )
        for plane in roof_planes
    ]

if max_feasible_panels < 1:
    st.warning("The usable roof dimensions are too small to fit one module under this dimension-based method.")
    installed_panel_count = 0
    actual_array_panel_counts = [0] * len(actual_arrays)
else:
    installed_panel_count = st.slider(
        "Installed panel count",
        min_value=1,
        max_value=max_feasible_panels + 5,
        value=max_feasible_panels,
        step=1,
        key="installed_panel_count",
    )
    actual_array_panel_counts = allocate_integer_counts(
        total_count=installed_panel_count,
        share_fractions=[arr.area_share_fraction for arr in actual_arrays],
    )

actual_building_kwp = installed_panel_count * module_power_kwp

actual_building_generation_kwh = 0.0
for arr, panel_count in zip(actual_arrays, actual_array_panel_counts):
    array_kwp = panel_count * module_power_kwp
    arr_factor = sap_orientation_factor_placeholder(arr.azimuth_deg) * sap_tilt_factor_placeholder(arr.tilt_deg)
    actual_building_generation_kwh += array_kwp * sap_placeholder_specific_yield * arr_factor

actual_generation_status = format_pass_fail(actual_building_generation_kwh, part_l_required_generation_kwh)
actual_kwp_status = format_pass_fail(actual_building_kwp, part_l_required_kwp)
actual_panel_status = format_pass_fail(float(installed_panel_count), float(part_l_required_panel_count))

actual_summary_cols = st.columns(3)
with actual_summary_cols[0]:
    render_summary_card(
        "Actual building annual generation",
        f"{actual_building_generation_kwh:,.0f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWh/a</span>",
    )
with actual_summary_cols[1]:
    render_summary_card(
        "Actual building kWp",
        f"{actual_building_kwp:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>",
    )
with actual_summary_cols[2]:
    render_summary_card(
        "Installed panel count",
        f"{installed_panel_count:,.0f}",
    )

comparison_fig = build_comparison_chart(
    required_generation_kwh=part_l_required_generation_kwh,
    actual_generation_kwh=actual_building_generation_kwh,
)
st.plotly_chart(comparison_fig, theme=None, use_container_width=True)

total_gross_roof_area_m2 = sum(plane["gross_area_m2"] for plane in roof_planes)
total_usable_before_blocked_m2 = sum(plane["usable_area_before_blocked_m2"] for plane in roof_planes)
usable_available_pv_area_m2 = sum(plane["usable_area_after_blocked_m2"] for plane in roof_planes)
other_reduction_area_m2 = max(total_gross_roof_area_m2 - total_usable_before_blocked_m2, 0.0)

actual_building_df = pd.DataFrame(
    [
        ("Roof type", actual_roof_form),
        ("Roof plane count used", f"{len(roof_planes)}"),
        ("Length along ridge / whole roof length in plan", f"{plan_length_along_ridge_m:,.2f} m"),
        ("Ridge-to-eaves / whole roof width in plan", f"{plan_length_ridge_to_eaves_m:,.2f} m"),
        ("Total gross roof area", f"{total_gross_roof_area_m2:,.2f} m²"),
        ("Blocked area", f"{blocked_area_total_m2:,.2f} m²"),
        ("Other reduction area from margins / offsets", f"{other_reduction_area_m2:,.2f} m²"),
        ("Usable roof area", f"{usable_available_pv_area_m2:,.2f} m²"),
        ("Roof reduction method", offset_mode_section_2),
        ("Module width", f"{module_width_m * 1000:.0f} mm"),
        ("Module length", f"{module_length_m * 1000:.0f} mm"),
        ("Module efficiency", f"{module_efficiency_pct:,.1f} %"),
        ("Derived module power", f"{module_power_wp:,.0f} Wp"),
        ("Mount orientation selected", module_mount_orientation),
        ("Maximum feasible panel count", f"{max_feasible_panels}"),
        ("Installed panel count", f"{installed_panel_count}"),
        ("Actual building annual generation", f"{actual_building_generation_kwh:,.0f} kWh/a"),
        ("Actual building kWp", f"{actual_building_kwp:,.2f} kWp"),
        ("Generation check against Part L requirement", actual_generation_status),
        ("kWp check against Part L requirement", actual_kwp_status),
        ("Panel count check against Part L requirement", actual_panel_status),
    ],
    columns=["Metric", "Value"],
)
st.dataframe(actual_building_df, hide_index=True, use_container_width=True)

array_rows = []
for plane, installed_panels in zip(roof_planes, actual_array_panel_counts):
    arr_factor = sap_orientation_factor_placeholder(plane["azimuth_deg"]) * sap_tilt_factor_placeholder(plane["tilt_deg"])
    array_rows.append(
        {
            "Array": plane["name"],
            "Azimuth (deg)": f"{plane['azimuth_deg']:.0f}",
            "Pitch (deg)": f"{plane['tilt_deg']:.0f}",
            "Gross length (m)": f"{plane['gross_length_m']:.2f}",
            "Gross depth (m)": f"{plane['gross_depth_m']:.2f}",
            "Usable length (m)": f"{plane['usable_length_m']:.2f}",
            "Usable depth (m)": f"{plane['usable_depth_m']:.2f}",
            "Usable area after blocked (m²)": f"{plane['usable_area_after_blocked_m2']:.2f}",
            "Max feasible panels": f"{plane['max_feasible_panels']}",
            "Installed panels": f"{installed_panels}",
            "SAP placeholder factor": f"{arr_factor:.3f}",
        }
    )
st.dataframe(pd.DataFrame(array_rows), hide_index=True, use_container_width=True)

st.divider()

# -----------------------------------------------------------------------------
# 3) The PySAM annual generation forecast
# -----------------------------------------------------------------------------
st.markdown("## 3) The PySAM annual generation forecast")
st.caption(
    "This section is secondary. It reuses the arrays defined in Section 2 and does not affect the Part L calculations above."
)

epw_lookup = get_available_epw_files(EPW_DIRECTORY)
epw_labels = ["None"] + list(epw_lookup.keys())

pysam_input_cols = st.columns(2)
with pysam_input_cols[0]:
    epw_label = st.selectbox("PV yield weather file (EPW)", epw_labels, index=0)
with pysam_input_cols[1]:
    st.text_input(
        "Array definitions source",
        value="Inherited from Section 2",
        disabled=True,
    )

pysam_result = None
pysam_message = None
selected_epw = None

if installed_panel_count < 1:
    pysam_message = "No installed panels defined in Section 2."
elif epw_label == "None":
    pysam_message = "No EPW selected."
elif pvwatts is None:
    pysam_message = "PySAM is not installed in this environment."
else:
    selected_epw = epw_lookup[epw_label]
    monthly_total = [0.0] * 12
    annual_total = 0.0

    try:
        pysam_array_rows = []

        for arr, panel_count in zip(actual_arrays, actual_array_panel_counts):
            if panel_count < 1:
                continue

            array_kwp = panel_count * module_power_kwp
            arr_result = run_pysam_pvwatts(
                system_capacity_kw=array_kwp,
                weather_file=selected_epw,
                tilt_deg=arr.tilt_deg,
                azimuth_deg=arr.azimuth_deg,
            )

            annual_total += arr_result["annual_ac_kwh"]
            monthly_total = [a + b for a, b in zip(monthly_total, arr_result["monthly_ac_kwh"])]

            pysam_array_rows.append(
                {
                    "Array": arr.name,
                    "Azimuth (deg)": f"{arr.azimuth_deg:.0f}",
                    "Pitch (deg)": f"{arr.tilt_deg:.0f}",
                    "Installed panels": f"{panel_count}",
                    "System capacity (kWp)": f"{array_kwp:.2f}",
                    "Annual AC generation (kWh/a)": f"{arr_result['annual_ac_kwh']:.0f}",
                }
            )

        pysam_result = {
            "annual_ac_kwh": annual_total,
            "monthly_ac_kwh": monthly_total,
            "array_rows": pysam_array_rows,
        }

    except Exception as exc:
        pysam_message = f"PySAM run failed: {exc}"

if pysam_result is not None:
    annual_gen_df = pd.DataFrame(
        [
            ("Generation methodology", "PySAM PVWatts v8"),
            ("Radiation / weather source", selected_epw.name if selected_epw else ""),
            ("Selected EPW", epw_label),
            ("Total installed panel count", f"{installed_panel_count}"),
            ("Total system capacity used", f"{actual_building_kwp:,.2f} kWp"),
            ("Annual AC generation", f"{pysam_result['annual_ac_kwh']:,.0f} kWh/a"),
        ],
        columns=["Metric", "Value"],
    )
    st.dataframe(annual_gen_df, hide_index=True, use_container_width=True)

    st.dataframe(pd.DataFrame(pysam_result["array_rows"]), hide_index=True, use_container_width=True)

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
            "AC generation (kWh)": [round(v, 1) for v in pysam_result["monthly_ac_kwh"]],
        }
    )
    st.dataframe(monthly_df, hide_index=True, use_container_width=True)
else:
    st.info(pysam_message or "Annual generation not available.")

st.markdown("### Method summary")
st.markdown(
    """
This tool is split into three sections.

**1) The Part L requirement**  
This section calculates the standardised Part L reference requirement from ground floor area. It outputs:
- required annual generation;
- required kWp; and
- required panel count using fixed standardised module assumptions.

**2) Your Building measured by Part L**  
This section now uses dimension-based roof planes:
- **Flat roof**: the user enters the whole roof length and width in plan;
- **Mono-pitch**: the user enters ridge length in plan and ridge-to-eaves length in plan, then pitch and orientation;
- **Duo-pitch**: the tool uses the same mono-pitch inputs, duplicates the roof plane in the background, and rotates the second plane by 180°.

The plan ridge-to-eaves dimension is converted to a sloping roof dimension using trigonometry for mono-pitch and duo-pitch roofs. Roof reductions can then be applied using blocked area plus either a simple perimeter margin or detailed offsets. Panel counting is now dimension-based, so portrait and landscape can produce different results.

**3) The PySAM annual generation forecast**  
This section reuses the arrays defined in Section 2 and applies PySAM PVWatts to estimate a more realistic annual output using a selected EPW weather file. It does not affect the Part L-side calculations.

The SAP regional annual yields used in Sections 1 and 2 are placeholders and still need to be checked against the approved methodology.
"""
)

with st.expander("Current limits", expanded=False):
    st.markdown(
        """
- The SAP regional annual yields are placeholders and need checking.
- Section 2 still uses simplified roof reductions and does not attempt full obstacle geometry.
- Flat roofs assume 50/50 east-west arrays at 12° tilt.
- Flat-roof row spacing still needs to be thought through properly.
- Hipped roofs are not included in the simplified Section 2 method.
- Blocked area is applied as an area deduction rather than a geometric cut-out.
- PySAM uses locally stored EPW files in `resources/epw/`.
"""
    )