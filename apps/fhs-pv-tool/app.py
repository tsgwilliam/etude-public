#TODO Graph is not very good, flat roof row spacing, postcode-level SAP region mapping, hipped/asymmetric roofs and more detailed module sizing checks

import math
import json
import base64
from dataclasses import asdict, dataclass
from pathlib import Path
from copy import deepcopy

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    import PySAM.Pvwattsv8 as pvwatts
except ImportError:
    pvwatts = None


st.set_page_config(page_title="Part L 2026 Photovoltaic Array Calculator", layout="wide")

# -----------------------------------------------------------------------------
# Part L reference assumptions
# -----------------------------------------------------------------------------
STANDARD_PANEL_EFFICIENCY_KWP_PER_M2 = 0.22
FHS_REQUIRED_AREA_FRACTION = 0.40


# -----------------------------------------------------------------------------
# Usable roof area assumptions
# -----------------------------------------------------------------------------
DEFAULT_SIMPLE_SETBACK_M = 0.30
DEFAULT_RIDGE_OFFSET_M = 0.60
DEFAULT_EDGE_OFFSET_M = 0.50
DEFAULT_PARTY_WALL_OFFSET_M = 0.75

DEFAULT_FLAT_PANEL_PITCH_DEG = 12.0

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
APP_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = APP_DIR / "resources"
EPW_DIRECTORY = RESOURCE_DIR / "epw"
REQUIREMENTS_PATH = APP_DIR / "requirements.txt"

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
SUMMARY_CARD_PADDING = "16px 18px 18px 18px"
SUMMARY_CARD_MIN_HEIGHT = "122px"
SUMMARY_CARD_SHADOW = "0 1px 3px rgba(0,0,0,0.06)"

SUMMARY_LABEL_FONT_SIZE = "14px"
SUMMARY_LABEL_FONT_WEIGHT = "500"
SUMMARY_LABEL_COLOUR = "#666666"
SUMMARY_LABEL_MARGIN_BOTTOM = "10px"

SUMMARY_VALUE_FONT_SIZE = "42px"
SUMMARY_VALUE_FONT_WEIGHT = "700"
SUMMARY_VALUE_COLOUR = "#222222"
SUMMARY_VALUE_LINE_HEIGHT = "1.10"

SUMMARY_UNIT_FONT_SIZE = "0.72em"
SUMMARY_UNIT_FONT_WEIGHT = "600"
SUMMARY_UNIT_COLOUR = "#444444"

LOGO_PATH = RESOURCE_DIR / "Etude-logo-animation-v005 Single spin.gif"
LOGO_WIDTH = 180


# -----------------------------------------------------------------------------
# Section theme colours from logo
# -----------------------------------------------------------------------------
ETUDE_PURPLE = "#6A4BA3"
ETUDE_ORANGE = "#D97C3F"
ETUDE_SAGE = "#B8C95E"
ETUDE_BLUE = "#79AFCB"
ETUDE_PINK = "#D14B8F"

SECTION_THEME_COLOURS = {
    "dwelling_inputs": ETUDE_BLUE,
    "part_l_target": ETUDE_PURPLE,
    "building_pv_layout": ETUDE_ORANGE,
    "generation_estimate": ETUDE_SAGE,
    "inputs_summary": ETUDE_BLUE,
    "calculation_assumptions": ETUDE_PINK,
    "roof_plane_table": ETUDE_PURPLE,
    "editor_json": ETUDE_ORANGE,
    "method_summary": ETUDE_BLUE,
    "current_limits": ETUDE_SAGE,
}

# -----------------------------------------------------------------------------
# Diagram styling
# -----------------------------------------------------------------------------
DIAGRAM_SINGLE_HEIGHT = 760
DIAGRAM_DUO_HEIGHT = 760
DIAGRAM_MARGIN = dict(l=30, r=30, t=30, b=30)
DIAGRAM_PANEL_COLOUR = "#17C497"
DIAGRAM_PANEL_USER_COLOUR = "#17C497"
DIAGRAM_PANEL_BLOCKED_COLOUR = "#9E9E9E"
DIAGRAM_PANEL_INVALID_COLOUR = "#E95B54"
DIAGRAM_OBSTACLE_COLOUR = "#6F6F6F"
DIAGRAM_ROOF_COLOUR = "#DCE6F9"
DIAGRAM_USABLE_COLOUR = "#FBE3D6"
DIAGRAM_LINE_COLOUR = "#444444"
DIAGRAM_TEXT_COLOUR = "#333333"
DIAGRAM_ZONE_LINE_COLOUR = "#4F67FF"
DIAGRAM_ZONE_FILL_COLOUR = "rgba(79,103,255,0.08)"
DIAGRAM_PLANE_GAP_X = 2.5
DIAGRAM_PLANE_GAP_Y = 3.3

OBSTACLE_NUDGE_STEP_DEFAULT = 0.10

# -----------------------------------------------------------------------------
# Geometry / state models
# -----------------------------------------------------------------------------
@dataclass
class Rect:
    x: float
    y: float
    w: float
    h: float


@dataclass
class PvArray:
    """
    Common array model used by both input routes:
    - visual roof layout
    - manual array input

    This is now the handover object between the layout section and the
    energy-generation section.
    """
    name: str
    capacity_kwp: float
    azimuth_deg: float
    tilt_deg: float
    shading_factor: float
    panel_count: int | None
    source: str


@dataclass
class ArrayDefinition:
    """
    Internal geometry allocation model used by the visual layout route.

    Keep this for now because the existing visual editor uses it to split panel
    counts across roof planes. It is not the final user-facing array model.
    """
    name: str
    azimuth_deg: float
    tilt_deg: float
    area_share_fraction: float


@dataclass
class RoofPlaneGeometry:
    plane_id: str
    name: str
    roof_form: str
    azimuth_deg: float
    tilt_deg: float
    gross_length_m: float
    gross_depth_m: float
    margin_left_m: float
    margin_right_m: float
    margin_top_m: float
    margin_bottom_m: float
    usable_length_m: float
    usable_depth_m: float
    gross_area_m2: float
    usable_area_m2: float
    packing_length_m: float
    packing_depth_m: float

    def gross_rect(self) -> Rect:
        return Rect(0.0, 0.0, self.gross_length_m, self.gross_depth_m)

    def usable_rect(self) -> Rect:
        return Rect(
            self.margin_left_m,
            self.margin_top_m,
            self.usable_length_m,
            self.usable_depth_m,
        )

    def packing_rect(self) -> Rect:
        return Rect(
            self.margin_left_m,
            self.margin_top_m,
            self.packing_length_m,
            self.packing_depth_m,
        )


@dataclass
class PanelLayout:
    panel_w_m: float
    panel_h_m: float
    cols: int
    rows: int
    count: int
    used_width_m: float
    used_height_m: float


@dataclass
class RoofGeometryBundle:
    roof_form: str
    planes: list[RoofPlaneGeometry]


@dataclass
class EditorPanelInstance:
    panel_id: str
    plane_id: str
    x: float
    y: float
    w: float
    h: float
    rotation_deg: float
    status: str
    zone_id: str | None = None


@dataclass
class EditorObstacle:
    obstacle_id: str
    plane_id: str
    x: float
    y: float
    w: float
    h: float
    obstacle_type: str


@dataclass
class EditorPvZone:
    zone_id: str
    plane_id: str
    x: float
    y: float
    w: float
    h: float
    label: str
    status: str = "valid"


@dataclass
class EditorPlaneState:
    plane_id: str
    name: str
    roof_form: str
    azimuth_deg: float
    tilt_deg: float
    gross_rect: dict
    usable_rect: dict
    packing_rect: dict
    max_feasible_panels: int
    displayed_panels: int
    obstacles: list[dict]
    pv_zones: list[dict]
    panels: list[dict]


@dataclass
class RoofEditorState:
    schema_version: str
    roof_form: str
    module: dict
    planes: list[dict]

# -----------------------------------------------------------------------------
# Styling helpers
# -----------------------------------------------------------------------------
def hex_to_rgba(hex_colour: str, alpha: float) -> str:
    hex_colour = hex_colour.lstrip("#")
    r = int(hex_colour[0:2], 16)
    g = int(hex_colour[2:4], 16)
    b = int(hex_colour[4:6], 16)
    return f"rgba({r}, {g}, {b}, {alpha})"


def render_section_title(section_key: str, title: str) -> None:
    colour = SECTION_THEME_COLOURS[section_key]
    st.markdown(
        f"""
        <div style="
            border: 1px solid {colour};
            background: {hex_to_rgba(colour, 0.14)};
            border-radius: 12px;
            padding: 14px 16px;
            font-size: 24px;
            font-weight: 700;
            color: #222222;
            margin: 18px 0 10px 0;
            line-height: 1.2;
        ">
            {title}
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Generic helpers
# -----------------------------------------------------------------------------
def get_base64_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()

def clear_widget_state_if_outside_range(
    key: str,
    min_value: float,
    max_value: float,
) -> None:
    if key not in st.session_state:
        return

    try:
        current_value = float(st.session_state[key])
    except (TypeError, ValueError):
        del st.session_state[key]
        return

    if current_value < min_value or current_value > max_value:
        del st.session_state[key]

def clear_widget_state_if_not_in_options(
    key: str,
    valid_options: list,
) -> None:
    if key in st.session_state and st.session_state[key] not in valid_options:
        del st.session_state[key]

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

# -----------------------------------------------------------------------------
# Shared photovoltaic array helpers
# -----------------------------------------------------------------------------
SAP_ORIENTATION_OPTIONS = {
    "North": 0.0,
    "North East": 45.0,
    "East": 90.0,
    "South East": 135.0,
    "South": 180.0,
    "South West": 225.0,
    "West": 270.0,
    "North West": 315.0,
}

SAP_TILT_OPTIONS_DEG = [
    0,
    30,
    45,
    60,
    90,
]

SAP_SHADING_OPTIONS = {
    "None or very little": 1.00,
    "Modest": 0.80,
    "Significant": 0.65,
    "Heavy": 0.50,
}


def normalise_azimuth_deg(azimuth_deg: float) -> float:
    return float(azimuth_deg) % 360.0


def get_orientation_label_from_azimuth(azimuth_deg: float) -> str:
    azimuth = normalise_azimuth_deg(azimuth_deg)
    closest_label = min(
        SAP_ORIENTATION_OPTIONS,
        key=lambda label: abs(((SAP_ORIENTATION_OPTIONS[label] - azimuth + 180.0) % 360.0) - 180.0),
    )
    return closest_label


def get_shading_label_from_factor(shading_factor: float) -> str:
    shading_factor = float(shading_factor)
    closest_label = min(
        SAP_SHADING_OPTIONS,
        key=lambda label: abs(SAP_SHADING_OPTIONS[label] - shading_factor),
    )
    return closest_label


def get_total_array_capacity_kwp(pv_arrays: list[PvArray]) -> float:
    return sum(max(float(array.capacity_kwp), 0.0) for array in pv_arrays)


def get_total_array_panel_count(pv_arrays: list[PvArray]) -> int:
    return sum(int(array.panel_count or 0) for array in pv_arrays)


def get_enabled_pv_arrays(pv_arrays: list[PvArray]) -> list[PvArray]:
    return [
        array for array in pv_arrays
        if float(array.capacity_kwp) > 0.0
    ]


def build_pv_arrays_from_visual_layout(
    actual_arrays: list[ArrayDefinition],
    actual_array_panel_counts: list[int],
    module_power_kwp: float,
    source_label: str = "Visual roof layout",
) -> list[PvArray]:
    """
    Converts the current visual-layout array allocation into the shared PvArray model.

    For flat roofs, the existing model still represents this as east/west arrays.
    For pitched roofs, the arrays follow the generated roof planes.
    """
    pv_arrays: list[PvArray] = []

    for arr, panel_count in zip(actual_arrays, actual_array_panel_counts):
        panel_count = int(panel_count)
        capacity_kwp = panel_count * module_power_kwp

        if panel_count <= 0 or capacity_kwp <= 0:
            continue

        pv_arrays.append(
            PvArray(
                name=arr.name,
                capacity_kwp=capacity_kwp,
                azimuth_deg=normalise_azimuth_deg(arr.azimuth_deg),
                tilt_deg=float(arr.tilt_deg),
                shading_factor=1.0,
                panel_count=panel_count,
                source=source_label,
            )
        )

    return pv_arrays


def build_pv_arrays_from_manual_inputs(
    manual_array_inputs: list[dict],
    source_label: str = "Manual array input",
) -> list[PvArray]:
    """
    Converts manual Streamlit input records into the shared PvArray model.

    Expected record keys:
    - enabled
    - name
    - capacity_kwp
    - orientation_label
    - tilt_deg
    - shading_label
    - panel_count
    """
    pv_arrays: list[PvArray] = []

    for idx, record in enumerate(manual_array_inputs, start=1):
        if not bool(record.get("enabled", False)):
            continue

        capacity_kwp = max(float(record.get("capacity_kwp", 0.0)), 0.0)
        if capacity_kwp <= 0:
            continue

        orientation_label = str(record.get("orientation_label", "South"))
        shading_label = str(record.get("shading_label", "None or very little"))

        panel_count_raw = record.get("panel_count")
        panel_count = None
        if panel_count_raw not in {None, ""}:
            panel_count = max(int(panel_count_raw), 0)

        pv_arrays.append(
            PvArray(
                name=str(record.get("name", "")).strip() or f"Array {idx}",
                capacity_kwp=capacity_kwp,
                azimuth_deg=SAP_ORIENTATION_OPTIONS.get(orientation_label, 180.0),
                tilt_deg=float(record.get("tilt_deg", 30.0)),
                shading_factor=SAP_SHADING_OPTIONS.get(shading_label, 1.0),
                panel_count=panel_count,
                source=source_label,
            )
        )

    return pv_arrays


def build_array_summary_rows(pv_arrays: list[PvArray]) -> list[dict]:
    rows = []

    for array in pv_arrays:
        rows.append(
            {
                "Array": array.name,
                "Input source": array.source,
                "Capacity (kWp)": f"{array.capacity_kwp:.2f}",
                "Orientation": get_orientation_label_from_azimuth(array.azimuth_deg),
                "Azimuth (deg)": f"{array.azimuth_deg:.0f}",
                "Tilt (deg)": f"{array.tilt_deg:.0f}",
                "Shading": get_shading_label_from_factor(array.shading_factor),
                "Shading factor": f"{array.shading_factor:.2f}",
                "Panel count": "" if array.panel_count is None else f"{array.panel_count}",
            }
        )

    return rows


def build_empty_pv_array_summary_rows() -> list[dict]:
    return [
        {
            "Array": "No arrays defined",
            "Input source": "",
            "Capacity (kWp)": "0.00",
            "Orientation": "",
            "Azimuth (deg)": "",
            "Tilt (deg)": "",
            "Shading": "",
            "Shading factor": "",
            "Panel count": "",
        }
    ]


def get_part_l_capacity_progress_pct(
    pv_arrays: list[PvArray],
    required_kwp: float,
) -> float:
    if required_kwp <= 0:
        return 0.0
    return get_total_array_capacity_kwp(pv_arrays) / required_kwp * 100.0


def get_array_capacity_status(
    pv_arrays: list[PvArray],
    required_kwp: float,
) -> str:
    return format_pass_fail(
        actual=get_total_array_capacity_kwp(pv_arrays),
        target=required_kwp,
    )

def get_available_epw_files(epw_dir: Path) -> dict[str, Path]:
    if not epw_dir.exists():
        return {}
    epw_files = sorted(epw_dir.glob("*.epw"))
    return {f.stem.replace("_", " "): f for f in epw_files}


def get_pysam_missing_message() -> str:
    if REQUIREMENTS_PATH.exists():
        return (
            "PySAM is not installed in this Python environment. "
            "Install the app dependencies in the same environment that is running Streamlit, "
            "then restart the app."
        )
    return (
        "PySAM is not installed in this Python environment. "
        "Install it in the same environment that is running Streamlit, then restart the app."
    )


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

# -----------------------------------------------------------------------------
# SAP Appendix U photovoltaic generation helpers
# -----------------------------------------------------------------------------
SAP_APPENDIX_U_MONTHS = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

SAP_APPENDIX_U_MONTH_DAYS = [
    31, 28, 31, 30, 31, 30,
    31, 31, 30, 31, 30, 31,
]

# SAP Appendix U Table U3 solar declination, degrees.
# The first value in SAP Table U3 is the monthly mean horizontal irradiance.
# Declination is then used in U3.2 for inclined/oriented surfaces.
SAP_APPENDIX_U_SOLAR_DECLINATION_DEG = [
    -20.7, -12.8, -1.8, 9.8, 18.8, 23.1,
    21.2, 13.7, 2.9, -8.7, -18.4, -23.0,
]

# Broad app regions mapped to representative SAP climate regions.
# These are not postcode-district precise. They are suitable for this simplified tool,
# but a compliance implementation should use SAP postcode-region mapping / PCDB data.
SAP_APPENDIX_U_REGION_DATA = {
    "England - London / South East": {
        "sap_region": 1,
        "sap_region_name": "Thames",
        "latitude_deg": 51.6,
        "horizontal_irradiance_w_m2": [30, 56, 98, 157, 195, 217, 203, 173, 127, 73, 39, 24],
    },
    "England - South": {
        "sap_region": 3,
        "sap_region_name": "Southern England",
        "latitude_deg": 50.9,
        "horizontal_irradiance_w_m2": [35, 62, 109, 172, 209, 235, 217, 185, 138, 80, 44, 27],
    },
    "England - South West": {
        "sap_region": 4,
        "sap_region_name": "South West England",
        "latitude_deg": 50.5,
        "horizontal_irradiance_w_m2": [36, 63, 111, 174, 210, 233, 204, 182, 136, 78, 44, 28],
    },
    "England - Midlands": {
        "sap_region": 6,
        "sap_region_name": "Midlands",
        "latitude_deg": 52.6,
        "horizontal_irradiance_w_m2": [28, 55, 97, 153, 191, 208, 194, 163, 121, 69, 35, 23],
    },
    "England - North": {
        "sap_region": 10,
        "sap_region_name": "North East England",
        "latitude_deg": 54.4,
        "horizontal_irradiance_w_m2": [25, 51, 95, 152, 196, 198, 190, 156, 115, 64, 32, 20],
    },
}

# SAP Appendix U Table U5 constants.
# Symmetric orientation groups use the same constants.
SAP_APPENDIX_U_TABLE_U5_CONSTANTS = {
    "North": {
        "k1": 26.3,
        "k2": -38.5,
        "k3": 14.8,
        "k4": -16.5,
        "k5": 27.3,
        "k6": -11.9,
        "k7": -1.06,
        "k8": 0.0872,
        "k9": -0.191,
    },
    "North East": {
        "k1": 0.165,
        "k2": -3.68,
        "k3": 3.0,
        "k4": 6.38,
        "k5": -4.53,
        "k6": -0.405,
        "k7": -4.38,
        "k8": 4.89,
        "k9": -1.99,
    },
    "East": {
        "k1": 1.44,
        "k2": -2.36,
        "k3": 1.07,
        "k4": -0.514,
        "k5": 1.89,
        "k6": -1.64,
        "k7": -0.542,
        "k8": -0.757,
        "k9": 0.604,
    },
    "South East": {
        "k1": -2.95,
        "k2": 2.89,
        "k3": 1.17,
        "k4": 5.67,
        "k5": -3.54,
        "k6": -4.28,
        "k7": -2.72,
        "k8": -0.25,
        "k9": 3.07,
    },
    "South": {
        "k1": -0.66,
        "k2": -0.106,
        "k3": 2.93,
        "k4": 3.63,
        "k5": -0.374,
        "k6": -7.4,
        "k7": -2.71,
        "k8": -0.991,
        "k9": 4.59,
    },
    "South West": {
        "k1": -2.95,
        "k2": 2.89,
        "k3": 1.17,
        "k4": 5.67,
        "k5": -3.54,
        "k6": -4.28,
        "k7": -2.72,
        "k8": -0.25,
        "k9": 3.07,
    },
    "West": {
        "k1": 1.44,
        "k2": -2.36,
        "k3": 1.07,
        "k4": -0.514,
        "k5": 1.89,
        "k6": -1.64,
        "k7": -0.542,
        "k8": -0.757,
        "k9": 0.604,
    },
    "North West": {
        "k1": 0.165,
        "k2": -3.68,
        "k3": 3.0,
        "k4": 6.38,
        "k5": -4.53,
        "k6": -0.405,
        "k7": -4.38,
        "k8": 4.89,
        "k9": -1.99,
    },
}


def get_appendix_u_orientation_label(azimuth_deg: float) -> str:
    return get_orientation_label_from_azimuth(azimuth_deg)


def get_appendix_u_tilt_label(tilt_deg: float) -> int:
    # This is retained for display only. Appendix U itself can calculate any tilt.
    tilt = float(tilt_deg)
    return int(round(tilt))


def has_appendix_u_table_data(region: str) -> bool:
    return region in SAP_APPENDIX_U_REGION_DATA


def get_appendix_u_region_data(region: str) -> dict:
    if region not in SAP_APPENDIX_U_REGION_DATA:
        available_regions = ", ".join(SAP_APPENDIX_U_REGION_DATA.keys())
        raise ValueError(
            f"No SAP Appendix U region data has been entered for '{region}'. "
            f"Available regions: {available_regions}."
        )

    region_data = SAP_APPENDIX_U_REGION_DATA[region]

    if len(region_data["horizontal_irradiance_w_m2"]) != 12:
        raise ValueError(
            f"SAP Appendix U horizontal irradiance data for '{region}' must contain 12 monthly values."
        )

    return region_data


def get_appendix_u_orientation_constants(orientation_label: str) -> dict:
    if orientation_label not in SAP_APPENDIX_U_TABLE_U5_CONSTANTS:
        available_orientations = ", ".join(SAP_APPENDIX_U_TABLE_U5_CONSTANTS.keys())
        raise ValueError(
            f"No SAP Appendix U Table U5 constants have been entered for '{orientation_label}'. "
            f"Available orientations: {available_orientations}."
        )

    return SAP_APPENDIX_U_TABLE_U5_CONSTANTS[orientation_label]


def calculate_appendix_u_rh_inc(
    orientation_label: str,
    tilt_deg: float,
    latitude_deg: float,
    declination_deg: float,
) -> float:
    constants = get_appendix_u_orientation_constants(orientation_label)

    half_tilt_rad = math.radians(float(tilt_deg) / 2.0)
    sin_half_tilt = math.sin(half_tilt_rad)
    sin_half_tilt_2 = sin_half_tilt ** 2
    sin_half_tilt_3 = sin_half_tilt ** 3

    a_value = (
        constants["k1"] * sin_half_tilt_3
        + constants["k2"] * sin_half_tilt_2
        + constants["k3"] * sin_half_tilt
    )
    b_value = (
        constants["k4"] * sin_half_tilt_3
        + constants["k5"] * sin_half_tilt_2
        + constants["k6"] * sin_half_tilt
    )
    c_value = (
        constants["k7"] * sin_half_tilt_3
        + constants["k8"] * sin_half_tilt_2
        + constants["k9"] * sin_half_tilt
        + 1.0
    )

    latitude_minus_declination_rad = math.radians(float(latitude_deg) - float(declination_deg))
    cos_term = math.cos(latitude_minus_declination_rad)

    rh_inc = a_value * (cos_term ** 2) + b_value * cos_term + c_value

    return max(float(rh_inc), 0.0)


def calculate_appendix_u_monthly_surface_irradiance(
    region: str,
    orientation_label: str,
    tilt_deg: float,
) -> list[float]:
    region_data = get_appendix_u_region_data(region)
    latitude_deg = float(region_data["latitude_deg"])
    horizontal_irradiance = region_data["horizontal_irradiance_w_m2"]

    monthly_surface_irradiance = []

    for horizontal_flux, declination_deg in zip(
        horizontal_irradiance,
        SAP_APPENDIX_U_SOLAR_DECLINATION_DEG,
    ):
        rh_inc = calculate_appendix_u_rh_inc(
            orientation_label=orientation_label,
            tilt_deg=float(tilt_deg),
            latitude_deg=latitude_deg,
            declination_deg=float(declination_deg),
        )
        monthly_surface_irradiance.append(float(horizontal_flux) * rh_inc)

    return monthly_surface_irradiance


def calculate_appendix_u_annual_surface_irradiation_kwh_m2(
    monthly_surface_irradiance_w_m2: list[float],
) -> float:
    if len(monthly_surface_irradiance_w_m2) != 12:
        raise ValueError("Monthly surface irradiance must contain 12 values.")

    return 0.024 * sum(
        days * irradiance
        for days, irradiance in zip(SAP_APPENDIX_U_MONTH_DAYS, monthly_surface_irradiance_w_m2)
    )


def calculate_sap_appendix_u_array_generation(
    array: PvArray,
    region: str,
    inverter_efficiency: float = 0.95,
) -> dict:
    orientation_label = get_appendix_u_orientation_label(array.azimuth_deg)
    tilt_label = get_appendix_u_tilt_label(array.tilt_deg)
    shading_factor = max(min(float(array.shading_factor), 1.0), 0.0)

    monthly_surface_irradiance_w_m2 = calculate_appendix_u_monthly_surface_irradiance(
        region=region,
        orientation_label=orientation_label,
        tilt_deg=float(array.tilt_deg),
    )

    monthly_surface_irradiation_kwh_m2 = [
        0.024 * days * irradiance
        for days, irradiance in zip(SAP_APPENDIX_U_MONTH_DAYS, monthly_surface_irradiance_w_m2)
    ]

    monthly_generation = [
        float(array.capacity_kwp) * monthly_irradiation * shading_factor * inverter_efficiency
        for monthly_irradiation in monthly_surface_irradiation_kwh_m2
    ]

    annual_surface_irradiation_kwh_m2 = sum(monthly_surface_irradiation_kwh_m2)
    annual_generation = sum(monthly_generation)

    region_data = get_appendix_u_region_data(region)

    return {
        "array_name": array.name,
        "source": array.source,
        "capacity_kwp": float(array.capacity_kwp),
        "panel_count": array.panel_count,
        "azimuth_deg": float(array.azimuth_deg),
        "orientation_label": orientation_label,
        "tilt_deg": float(array.tilt_deg),
        "tilt_label": tilt_label,
        "shading_factor": shading_factor,
        "inverter_efficiency": float(inverter_efficiency),
        "sap_region": region_data["sap_region"],
        "sap_region_name": region_data["sap_region_name"],
        "latitude_deg": region_data["latitude_deg"],
        "monthly_surface_irradiance_w_m2": monthly_surface_irradiance_w_m2,
        "monthly_surface_irradiation_kwh_m2": monthly_surface_irradiation_kwh_m2,
        "annual_surface_irradiation_kwh_m2": annual_surface_irradiation_kwh_m2,
        "monthly_generation_kwh": monthly_generation,
        "annual_generation_kwh": annual_generation,
    }


def calculate_sap_appendix_u_generation(
    pv_arrays: list[PvArray],
    region: str,
    inverter_efficiency: float = 0.95,
) -> dict:
    enabled_arrays = get_enabled_pv_arrays(pv_arrays)

    monthly_total = [0.0] * 12
    annual_total = 0.0
    array_results = []

    for array in enabled_arrays:
        array_result = calculate_sap_appendix_u_array_generation(
            array=array,
            region=region,
            inverter_efficiency=inverter_efficiency,
        )

        annual_total += array_result["annual_generation_kwh"]
        monthly_total = [
            existing + new
            for existing, new in zip(monthly_total, array_result["monthly_generation_kwh"])
        ]
        array_results.append(array_result)

    region_data = get_appendix_u_region_data(region)

    return {
        "method": "SAP Appendix U",
        "region": region,
        "sap_region": region_data["sap_region"],
        "sap_region_name": region_data["sap_region_name"],
        "latitude_deg": region_data["latitude_deg"],
        "annual_generation_kwh": annual_total,
        "monthly_generation_kwh": monthly_total,
        "array_results": array_results,
    }


def build_sap_appendix_u_array_rows(array_results: list[dict]) -> list[dict]:
    rows = []

    for result in array_results:
        rows.append(
            {
                "Array": result["array_name"],
                "Input source": result["source"],
                "Capacity (kWp)": f"{result['capacity_kwp']:.2f}",
                "Panel count": "" if result["panel_count"] is None else f"{result['panel_count']}",
                "Azimuth (deg)": f"{result['azimuth_deg']:.0f}",
                "Appendix U orientation": result["orientation_label"],
                "Tilt entered (deg)": f"{result['tilt_deg']:.0f}",
                "SAP region": f"{result['sap_region']} - {result['sap_region_name']}",
                "Representative latitude": f"{result['latitude_deg']:.1f}°N",
                "Annual surface irradiation (kWh/m²)": f"{result['annual_surface_irradiation_kwh_m2']:.0f}",
                "Shading factor": f"{result['shading_factor']:.2f}",
                "Inverter efficiency": f"{result['inverter_efficiency']:.2f}",
                "Annual generation (kWh/a)": f"{result['annual_generation_kwh']:.0f}",
            }
        )

    return rows

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

def build_part_l_capacity_progress_chart(required_kwp: float, actual_kwp: float) -> go.Figure:
    achieved_pct = (actual_kwp / required_kwp * 100.0) if required_kwp > 0 else 0.0
    achieved_pct_capped = min(achieved_pct, 100.0)
    shortfall_pct = max(100.0 - achieved_pct_capped, 0.0)

    fig = go.Figure()

    fig.add_bar(
        x=["Part L photovoltaic target"],
        y=[achieved_pct_capped],
        name="Achieved",
        marker_color="#17C497",
    )

    fig.add_bar(
        x=["Part L photovoltaic target"],
        y=[shortfall_pct],
        name="Shortfall",
        marker_color="#E6E6E6",
    )

    fig.update_layout(
        barmode="stack",
        height=300,
        margin=dict(l=60, r=20, t=24, b=40),
        yaxis_title="% of target kWp",
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#333333"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )

    fig.update_yaxes(
        range=[0, max(110.0, achieved_pct * 1.10)],
        gridcolor="#E6E6E6",
    )
    fig.update_xaxes(showgrid=False)

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


def rect_to_dict(rect: Rect) -> dict:
    return {"x": rect.x, "y": rect.y, "w": rect.w, "h": rect.h}


def dict_to_rect(rect_dict: dict) -> Rect:
    return Rect(
        x=float(rect_dict["x"]),
        y=float(rect_dict["y"]),
        w=float(rect_dict["w"]),
        h=float(rect_dict["h"]),
    )


def get_plane_displayed_panel_counts(
    roof_form: str,
    installed_panel_count: int,
    actual_array_panel_counts: list[int],
) -> list[int]:
    if roof_form == "Flat":
        return [installed_panel_count]
    return actual_array_panel_counts


def rect_contains(outer: Rect, inner: Rect, tol: float = 1e-9) -> bool:
    return (
        inner.x >= outer.x - tol
        and inner.y >= outer.y - tol
        and inner.x + inner.w <= outer.x + outer.w + tol
        and inner.y + inner.h <= outer.y + outer.h + tol
    )


def rects_intersect(a: Rect, b: Rect, tol: float = 1e-9) -> bool:
    return not (
        a.x + a.w <= b.x + tol
        or b.x + b.w <= a.x + tol
        or a.y + a.h <= b.y + tol
        or b.y + b.h <= a.y + tol
    )


# -----------------------------------------------------------------------------
# Geometry helpers
# -----------------------------------------------------------------------------
MIN_ZONE_DIM_M = 0.05


def rect_intersection(a: Rect, b: Rect) -> Rect | None:
    x0 = max(a.x, b.x)
    y0 = max(a.y, b.y)
    x1 = min(a.x + a.w, b.x + b.w)
    y1 = min(a.y + a.h, b.y + b.h)

    if x1 <= x0 or y1 <= y0:
        return None

    return Rect(x=x0, y=y0, w=x1 - x0, h=y1 - y0)


def subtract_rect_from_rect(
    source: Rect,
    cutter: Rect,
    min_dim: float = MIN_ZONE_DIM_M,
) -> list[Rect]:
    overlap = rect_intersection(source, cutter)
    if overlap is None:
        return [source]

    pieces = []

    if overlap.y > source.y:
        pieces.append(
            Rect(
                x=source.x,
                y=source.y,
                w=source.w,
                h=overlap.y - source.y,
            )
        )

    source_bottom = source.y + source.h
    overlap_bottom = overlap.y + overlap.h
    if overlap_bottom < source_bottom:
        pieces.append(
            Rect(
                x=source.x,
                y=overlap_bottom,
                w=source.w,
                h=source_bottom - overlap_bottom,
            )
        )

    if overlap.x > source.x:
        pieces.append(
            Rect(
                x=source.x,
                y=overlap.y,
                w=overlap.x - source.x,
                h=overlap.h,
            )
        )

    source_right = source.x + source.w
    overlap_right = overlap.x + overlap.w
    if overlap_right < source_right:
        pieces.append(
            Rect(
                x=overlap_right,
                y=overlap.y,
                w=source_right - overlap_right,
                h=overlap.h,
            )
        )

    return [
        piece for piece in pieces
        if piece.w >= min_dim and piece.h >= min_dim
    ]


def apply_obstacles_to_pv_zones(editor_state: dict) -> dict:
    editor_state = deepcopy(editor_state)

    for plane in editor_state["planes"]:
        plane_id = plane["plane_id"]
        packing_rect = dict_to_rect(plane["packing_rect"])
        usable_rect = dict_to_rect(plane["usable_rect"])

        obstacles = [
            clamp_obstacle_to_rect(obstacle, usable_rect)
            for obstacle in normalise_obstacle_records(
                obstacle_records=plane.get("obstacles", []),
                plane_id=plane_id,
            )
        ]

        source_zones = [
            clamp_pv_zone_to_rect(zone, packing_rect)
            for zone in normalise_pv_zone_records(
                pv_zone_records=plane.get("pv_zones", []),
                plane_id=plane_id,
            )
        ]

        effective_rects: list[Rect] = []

        for zone in source_zones:
            remaining = [
                Rect(
                    x=float(zone["x"]),
                    y=float(zone["y"]),
                    w=float(zone["w"]),
                    h=float(zone["h"]),
                )
            ]

            for obstacle in obstacles:
                obstacle_rect = Rect(
                    x=float(obstacle["x"]),
                    y=float(obstacle["y"]),
                    w=float(obstacle["w"]),
                    h=float(obstacle["h"]),
                )

                next_remaining: list[Rect] = []
                for rect in remaining:
                    next_remaining.extend(subtract_rect_from_rect(rect, obstacle_rect))
                remaining = next_remaining

            effective_rects.extend(remaining)

        effective_zones = []
        for idx, rect in enumerate(sorted(effective_rects, key=lambda r: (r.y, r.x)), start=1):
            effective_zones.append(
                {
                    "zone_id": f"{plane_id}_effective_zone_{idx}",
                    "plane_id": plane_id,
                    "x": rect.x,
                    "y": rect.y,
                    "w": rect.w,
                    "h": rect.h,
                    "label": f"Array zone {idx}",
                    "status": "valid",
                }
            )

        plane["pv_zones"] = effective_zones

    return editor_state


def get_module_footprint(
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
    roof_form: str,
    flat_panel_pitch_deg: float,
) -> tuple[float, float]:
    if roof_form == "Flat":
        projected_long_dim_m = module_length_m * math.cos(math.radians(flat_panel_pitch_deg))
    else:
        projected_long_dim_m = module_length_m

    if mount_orientation == "Portrait":
        return module_width_m, projected_long_dim_m

    return projected_long_dim_m, module_width_m


def calc_plane_panel_layout(
    packing_length_m: float,
    packing_depth_m: float,
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
    roof_form: str,
    flat_panel_pitch_deg: float,
) -> PanelLayout:
    panel_w_m, panel_h_m = get_module_footprint(
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        mount_orientation=mount_orientation,
        roof_form=roof_form,
        flat_panel_pitch_deg=flat_panel_pitch_deg,
    )

    cols = max(math.floor(packing_length_m / panel_w_m), 0) if panel_w_m > 0 else 0
    rows = max(math.floor(packing_depth_m / panel_h_m), 0) if panel_h_m > 0 else 0
    count = cols * rows

    used_width_m = cols * panel_w_m
    used_height_m = rows * panel_h_m

    return PanelLayout(
        panel_w_m=panel_w_m,
        panel_h_m=panel_h_m,
        cols=cols,
        rows=rows,
        count=count,
        used_width_m=used_width_m,
        used_height_m=used_height_m,
    )


def _build_single_plane(
    plane_id: str,
    name: str,
    roof_form: str,
    azimuth_deg: float,
    tilt_deg: float,
    gross_length_m: float,
    gross_depth_m: float,
    margin_left_m: float,
    margin_right_m: float,
    margin_top_m: float,
    margin_bottom_m: float,
) -> RoofPlaneGeometry:
    usable_length_m = max(gross_length_m - margin_left_m - margin_right_m, 0.0)
    usable_depth_m = max(gross_depth_m - margin_top_m - margin_bottom_m, 0.0)
    usable_area_m2 = usable_length_m * usable_depth_m

    return RoofPlaneGeometry(
        plane_id=plane_id,
        name=name,
        roof_form=roof_form,
        azimuth_deg=azimuth_deg,
        tilt_deg=tilt_deg,
        gross_length_m=gross_length_m,
        gross_depth_m=gross_depth_m,
        margin_left_m=margin_left_m,
        margin_right_m=margin_right_m,
        margin_top_m=margin_top_m,
        margin_bottom_m=margin_bottom_m,
        usable_length_m=usable_length_m,
        usable_depth_m=usable_depth_m,
        gross_area_m2=gross_length_m * gross_depth_m,
        usable_area_m2=usable_area_m2,
        packing_length_m=usable_length_m,
        packing_depth_m=usable_depth_m,
    )


def build_roof_geometry(
    roof_form: str,
    plan_length_along_ridge_m: float,
    plan_length_ridge_to_eaves_m: float,
    pitch_deg: float,
    azimuth_deg: float | None,
    simple_setback_m: float,
    ridge_offset_m: float,
    edge_offset_m: float,
    party_wall_offset_m: float,
    house_form: str,
) -> RoofGeometryBundle:
    planes: list[RoofPlaneGeometry] = []

    if roof_form == "Flat":
        gross_length_m = plan_length_along_ridge_m
        gross_depth_m = plan_length_ridge_to_eaves_m

        if simple_setback_m > 0:
            margin_left = simple_setback_m
            margin_right = simple_setback_m
            margin_top = simple_setback_m
            margin_bottom = simple_setback_m
        else:
            margin_left = edge_offset_m
            margin_right = edge_offset_m
            margin_top = edge_offset_m
            margin_bottom = edge_offset_m

        planes.append(
            _build_single_plane(
                plane_id="plane_1",
                name="Roof",
                roof_form=roof_form,
                azimuth_deg=0.0,
                tilt_deg=0.0,
                gross_length_m=gross_length_m,
                gross_depth_m=gross_depth_m,
                margin_left_m=margin_left,
                margin_right_m=margin_right,
                margin_top_m=margin_top,
                margin_bottom_m=margin_bottom,
            )
        )

        return RoofGeometryBundle(roof_form=roof_form, planes=planes)

    pitch_rad = math.radians(pitch_deg)
    gross_length_m = plan_length_along_ridge_m
    gross_depth_m = plan_length_ridge_to_eaves_m / max(math.cos(pitch_rad), 1e-6)

    if simple_setback_m > 0:
        margin_left = simple_setback_m
        margin_right = simple_setback_m
        margin_top = simple_setback_m
        margin_bottom = simple_setback_m
    else:
        party_wall_count = calc_party_wall_count(house_form)
        side_extra = (party_wall_count * party_wall_offset_m) / 2.0

        margin_left = edge_offset_m + side_extra
        margin_right = edge_offset_m + side_extra
        margin_top = ridge_offset_m
        margin_bottom = edge_offset_m

    if roof_form == "Mono-pitch":
        plane_azimuths = [float(azimuth_deg)]
        plane_names = ["Roof"]
    else:
        azimuth_1 = float(azimuth_deg)
        azimuth_2 = (azimuth_1 + 180.0) % 360.0
        plane_azimuths = [azimuth_1, azimuth_2]
        plane_names = ["Roof plane 1", "Roof plane 2"]

    for idx, (name, plane_azimuth) in enumerate(zip(plane_names, plane_azimuths), start=1):
        planes.append(
            _build_single_plane(
                plane_id=f"plane_{idx}",
                name=name,
                roof_form=roof_form,
                azimuth_deg=plane_azimuth,
                tilt_deg=float(pitch_deg),
                gross_length_m=gross_length_m,
                gross_depth_m=gross_depth_m,
                margin_left_m=margin_left,
                margin_right_m=margin_right,
                margin_top_m=margin_top,
                margin_bottom_m=margin_bottom,
            )
        )

    return RoofGeometryBundle(roof_form=roof_form, planes=planes)


def build_actual_arrays_for_generation(
    roof_form: str,
    mono_or_duo_azimuth_deg: float | None,
    mono_or_duo_pitch_deg: float,
    flat_panel_pitch_deg: float,
) -> list[ArrayDefinition]:
    if roof_form == "Mono-pitch":
        return [
            ArrayDefinition(
                name="Roof",
                azimuth_deg=float(mono_or_duo_azimuth_deg),
                tilt_deg=float(mono_or_duo_pitch_deg),
                area_share_fraction=1.0,
            )
        ]

    if roof_form == "Duo-pitch":
        azimuth_1 = float(mono_or_duo_azimuth_deg)
        azimuth_2 = (azimuth_1 + 180.0) % 360.0

        return [
            ArrayDefinition(
                name="Roof plane 1",
                azimuth_deg=azimuth_1,
                tilt_deg=float(mono_or_duo_pitch_deg),
                area_share_fraction=0.5,
            ),
            ArrayDefinition(
                name="Roof plane 2",
                azimuth_deg=azimuth_2,
                tilt_deg=float(mono_or_duo_pitch_deg),
                area_share_fraction=0.5,
            ),
        ]

    return [
        ArrayDefinition(
            name="East-facing array",
            azimuth_deg=90.0,
            tilt_deg=float(flat_panel_pitch_deg),
            area_share_fraction=0.5,
        ),
        ArrayDefinition(
            name="West-facing array",
            azimuth_deg=270.0,
            tilt_deg=float(flat_panel_pitch_deg),
            area_share_fraction=0.5,
        ),
    ]


def build_panel_instances_for_plane(
    plane: RoofPlaneGeometry,
    layout: PanelLayout,
    displayed_count: int,
) -> list[EditorPanelInstance]:
    if layout.count <= 0 or displayed_count <= 0:
        return []

    packing_rect = plane.packing_rect()
    start_x = packing_rect.x + max((packing_rect.w - layout.used_width_m) / 2.0, 0.0)
    start_y = packing_rect.y + max((packing_rect.h - layout.used_height_m) / 2.0, 0.0)

    panels: list[EditorPanelInstance] = []
    drawn = 0

    for row in range(layout.rows):
        for col in range(layout.cols):
            if drawn >= displayed_count:
                break

            panels.append(
                EditorPanelInstance(
                    panel_id=f"{plane.plane_id}_panel_{drawn + 1}",
                    plane_id=plane.plane_id,
                    x=start_x + col * layout.panel_w_m,
                    y=start_y + row * layout.panel_h_m,
                    w=layout.panel_w_m,
                    h=layout.panel_h_m,
                    rotation_deg=0.0,
                    status="auto",
                )
            )
            drawn += 1

        if drawn >= displayed_count:
            break

    return panels


def build_default_pv_zones_for_plane(plane: RoofPlaneGeometry) -> list[EditorPvZone]:
    packing = plane.packing_rect()

    if packing.w <= 0 or packing.h <= 0:
        return []

    return [
        EditorPvZone(
            zone_id=f"{plane.plane_id}_zone_1",
            plane_id=plane.plane_id,
            x=packing.x,
            y=packing.y,
            w=packing.w,
            h=packing.h,
            label="Array zone 1",
            status="valid",
        )
    ]


def serialize_plane_for_editor(
    plane: RoofPlaneGeometry,
    layout: PanelLayout,
    displayed_panels: int,
) -> dict:
    pv_zone_instances = build_default_pv_zones_for_plane(plane)
    panel_instances = build_panel_instances_for_plane(
        plane=plane,
        layout=layout,
        displayed_count=displayed_panels,
    )

    editor_plane = EditorPlaneState(
        plane_id=plane.plane_id,
        name=plane.name,
        roof_form=plane.roof_form,
        azimuth_deg=plane.azimuth_deg,
        tilt_deg=plane.tilt_deg,
        gross_rect=rect_to_dict(plane.gross_rect()),
        usable_rect=rect_to_dict(plane.usable_rect()),
        packing_rect=rect_to_dict(plane.packing_rect()),
        max_feasible_panels=layout.count,
        displayed_panels=displayed_panels,
        obstacles=[],
        pv_zones=[asdict(zone) for zone in pv_zone_instances],
        panels=[asdict(panel) for panel in panel_instances],
    )

    return asdict(editor_plane)


def build_roof_editor_state(
    geometry: RoofGeometryBundle,
    plane_layouts: dict[str, PanelLayout],
    displayed_panel_counts: list[int],
    module_length_m: float,
    module_width_m: float,
    module_efficiency_pct: float,
    module_mount_orientation: str,
    flat_panel_pitch_deg: float,
) -> dict:
    serialized_planes: list[dict] = []

    for plane, displayed_panels in zip(geometry.planes, displayed_panel_counts):
        layout = plane_layouts[plane.plane_id]
        serialized_planes.append(
            serialize_plane_for_editor(
                plane=plane,
                layout=layout,
                displayed_panels=displayed_panels,
            )
        )

    state = RoofEditorState(
        schema_version="0.2.0",
        roof_form=geometry.roof_form,
        module={
            "length_m": module_length_m,
            "width_m": module_width_m,
            "efficiency_pct": module_efficiency_pct,
            "mount_orientation": module_mount_orientation,
            "flat_panel_pitch_deg": flat_panel_pitch_deg,
        },
        planes=serialized_planes,
    )

    return asdict(state)


def get_editor_state_signature(
    geometry: RoofGeometryBundle,
    plane_layouts: dict[str, PanelLayout],
    displayed_panel_counts: list[int],
    module_length_m: float,
    module_width_m: float,
    module_efficiency_pct: float,
    module_mount_orientation: str,
    flat_panel_pitch_deg: float,
) -> str:
    signature_payload = {
        "roof_form": geometry.roof_form,
        "module": {
            "length_m": module_length_m,
            "width_m": module_width_m,
            "efficiency_pct": module_efficiency_pct,
            "mount_orientation": module_mount_orientation,
            "flat_panel_pitch_deg": flat_panel_pitch_deg,
        },
        "displayed_panel_counts": displayed_panel_counts,
        "planes": [
            {
                "plane_id": plane.plane_id,
                "name": plane.name,
                "azimuth_deg": plane.azimuth_deg,
                "tilt_deg": plane.tilt_deg,
                "gross_rect": rect_to_dict(plane.gross_rect()),
                "usable_rect": rect_to_dict(plane.usable_rect()),
                "packing_rect": rect_to_dict(plane.packing_rect()),
                "max_feasible_panels": plane_layouts[plane.plane_id].count,
            }
            for plane in geometry.planes
        ],
    }

    return json.dumps(signature_payload, sort_keys=True)

# -----------------------------------------------------------------------------
# Editor validation and metrics
# -----------------------------------------------------------------------------
def normalise_obstacle_records(
    obstacle_records: list[dict],
    plane_id: str,
) -> list[dict]:
    clean_records = []

    for idx, record in enumerate(obstacle_records, start=1):
        obstacle_id = str(record.get("obstacle_id", "")).strip() or f"{plane_id}_obstacle_{idx}"
        obstacle_type = str(record.get("obstacle_type", "generic")).strip() or "generic"

        try:
            x = float(record.get("x", 0.0))
            y = float(record.get("y", 0.0))
            w = max(float(record.get("w", 0.0)), 0.0)
            h = max(float(record.get("h", 0.0)), 0.0)
        except (TypeError, ValueError):
            continue

        if w <= 0 or h <= 0:
            continue

        clean_records.append(
            {
                "obstacle_id": obstacle_id,
                "plane_id": plane_id,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "obstacle_type": obstacle_type,
            }
        )

    return clean_records


def clamp_obstacle_to_rect(obstacle: dict, rect: Rect) -> dict:
    obstacle = deepcopy(obstacle)
    obstacle["w"] = max(float(obstacle["w"]), 0.0)
    obstacle["h"] = max(float(obstacle["h"]), 0.0)

    max_x = rect.x + max(rect.w - obstacle["w"], 0.0)
    max_y = rect.y + max(rect.h - obstacle["h"], 0.0)

    obstacle["x"] = min(max(float(obstacle["x"]), rect.x), max_x)
    obstacle["y"] = min(max(float(obstacle["y"]), rect.y), max_y)
    return obstacle


def normalise_pv_zone_records(
    pv_zone_records: list[dict],
    plane_id: str,
) -> list[dict]:
    clean_records = []

    for idx, record in enumerate(pv_zone_records, start=1):
        zone_id = str(record.get("zone_id", "")).strip() or f"{plane_id}_zone_{idx}"
        label = str(record.get("label", "")).strip() or f"Array zone {idx}"
        status = str(record.get("status", "valid")).strip() or "valid"

        if status not in {"valid", "blocked", "invalid"}:
            status = "valid"

        try:
            x = float(record.get("x", 0.0))
            y = float(record.get("y", 0.0))
            w = max(float(record.get("w", 0.0)), 0.0)
            h = max(float(record.get("h", 0.0)), 0.0)
        except (TypeError, ValueError):
            continue

        if w <= 0 or h <= 0:
            continue

        clean_records.append(
            {
                "zone_id": zone_id,
                "plane_id": plane_id,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "label": label,
                "status": status,
            }
        )

    return clean_records


def clamp_pv_zone_to_rect(pv_zone: dict, rect: Rect) -> dict:
    pv_zone = deepcopy(pv_zone)
    pv_zone["w"] = max(float(pv_zone["w"]), 0.0)
    pv_zone["h"] = max(float(pv_zone["h"]), 0.0)

    max_x = rect.x + max(rect.w - pv_zone["w"], 0.0)
    max_y = rect.y + max(rect.h - pv_zone["h"], 0.0)

    pv_zone["x"] = min(max(float(pv_zone["x"]), rect.x), max_x)
    pv_zone["y"] = min(max(float(pv_zone["y"]), rect.y), max_y)
    return pv_zone


def make_obstacle_id(plane_id: str, existing_obstacle_ids: list[str]) -> str:
    i = 1
    existing = set(existing_obstacle_ids)
    while f"{plane_id}_obstacle_{i}" in existing:
        i += 1
    return f"{plane_id}_obstacle_{i}"


def make_zone_id(plane_id: str, existing_zone_ids: list[str]) -> str:
    i = 1
    existing = set(existing_zone_ids)
    while f"{plane_id}_zone_{i}" in existing:
        i += 1
    return f"{plane_id}_zone_{i}"


def regenerate_panels_from_pv_zones(editor_state: dict) -> dict:
    editor_state = deepcopy(editor_state)

    module = editor_state["module"]
    module_length_m = float(module["length_m"])
    module_width_m = float(module["width_m"])
    module_mount_orientation = str(module["mount_orientation"])
    flat_panel_pitch_deg = float(module.get("flat_panel_pitch_deg", DEFAULT_FLAT_PANEL_PITCH_DEG))
    roof_form = str(editor_state["roof_form"])

    for plane in editor_state["planes"]:
        plane_id = plane["plane_id"]
        packing_rect = dict_to_rect(plane["packing_rect"])

        clean_zones = normalise_pv_zone_records(
            pv_zone_records=plane.get("pv_zones", []),
            plane_id=plane_id,
        )

        plane["pv_zones"] = [clamp_pv_zone_to_rect(zone, packing_rect) for zone in clean_zones]

        rebuilt_panels = []
        panel_counter = 1
        remaining = int(plane.get("displayed_panels", 0))

        for zone in plane["pv_zones"]:
            if remaining <= 0:
                break

            zone_layout = calc_plane_panel_layout(
                packing_length_m=float(zone["w"]),
                packing_depth_m=float(zone["h"]),
                module_length_m=module_length_m,
                module_width_m=module_width_m,
                mount_orientation=module_mount_orientation,
                roof_form=roof_form,
                flat_panel_pitch_deg=flat_panel_pitch_deg,
            )

            if zone_layout.count <= 0:
                continue

            draw_count = min(zone_layout.count, remaining)

            for row in range(zone_layout.rows):
                for col in range(zone_layout.cols):
                    if draw_count <= 0:
                        break

                    px = float(zone["x"]) + col * zone_layout.panel_w_m
                    py = float(zone["y"]) + row * zone_layout.panel_h_m

                    rebuilt_panels.append(
                        {
                            "panel_id": f"{plane_id}_panel_{panel_counter}",
                            "plane_id": plane_id,
                            "x": px,
                            "y": py,
                            "w": zone_layout.panel_w_m,
                            "h": zone_layout.panel_h_m,
                            "rotation_deg": 0.0,
                            "status": "auto",
                            "zone_id": zone["zone_id"],
                        }
                    )
                    panel_counter += 1
                    remaining -= 1
                    draw_count -= 1

                if draw_count <= 0:
                    break

        plane["panels"] = rebuilt_panels

    return editor_state


def validate_editor_state(editor_state: dict) -> dict:
    module = editor_state["module"]
    module_length_m = float(module["length_m"])
    module_width_m = float(module["width_m"])
    mount_orientation = str(module["mount_orientation"])
    flat_panel_pitch_deg = float(module.get("flat_panel_pitch_deg", DEFAULT_FLAT_PANEL_PITCH_DEG))
    roof_form = str(editor_state["roof_form"])

    panel_w_m, panel_h_m = get_module_footprint(
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        mount_orientation=mount_orientation,
        roof_form=roof_form,
        flat_panel_pitch_deg=flat_panel_pitch_deg,
    )

    for plane in editor_state["planes"]:
        packing_rect = dict_to_rect(plane["packing_rect"])

        clean_obstacles = normalise_obstacle_records(
            obstacle_records=plane.get("obstacles", []),
            plane_id=plane["plane_id"],
        )
        clean_obstacles = [
            clamp_obstacle_to_rect(obstacle, dict_to_rect(plane["usable_rect"]))
            for obstacle in clean_obstacles
        ]
        plane["obstacles"] = clean_obstacles
        obstacle_rects = [dict_to_rect(obs) for obs in clean_obstacles]

        clean_zones = normalise_pv_zone_records(
            pv_zone_records=plane.get("pv_zones", []),
            plane_id=plane["plane_id"],
        )
        clean_zones = [clamp_pv_zone_to_rect(zone, packing_rect) for zone in clean_zones]
        plane["pv_zones"] = clean_zones

        derived_panels = []
        panel_counter = 1
        remaining = int(plane.get("displayed_panels", 0))

        for zone in clean_zones:
            if remaining <= 0:
                break

            cols = max(math.floor(float(zone["w"]) / panel_w_m), 0) if panel_w_m > 0 else 0
            rows = max(math.floor(float(zone["h"]) / panel_h_m), 0) if panel_h_m > 0 else 0
            zone_capacity = cols * rows
            draw_count = min(zone_capacity, remaining)

            if draw_count <= 0:
                continue

            for row in range(rows):
                for col in range(cols):
                    if draw_count <= 0:
                        break

                    px = float(zone["x"]) + col * panel_w_m
                    py = float(zone["y"]) + row * panel_h_m

                    derived_panels.append(
                        {
                            "panel_id": f"{plane['plane_id']}_panel_{panel_counter}",
                            "plane_id": plane["plane_id"],
                            "x": px,
                            "y": py,
                            "w": panel_w_m,
                            "h": panel_h_m,
                            "rotation_deg": 0.0,
                            "status": "auto",
                            "zone_id": zone["zone_id"],
                        }
                    )
                    panel_counter += 1
                    remaining -= 1
                    draw_count -= 1

                if draw_count <= 0:
                    break

        cleaned_panels = []
        panel_rects = []

        for idx, panel in enumerate(derived_panels, start=1):
            clean_panel = deepcopy(panel)
            clean_panel["panel_id"] = (
                str(clean_panel.get("panel_id", "")).strip()
                or f"{plane['plane_id']}_panel_{idx}"
            )
            clean_panel["plane_id"] = plane["plane_id"]

            try:
                clean_panel["x"] = float(clean_panel.get("x", 0.0))
                clean_panel["y"] = float(clean_panel.get("y", 0.0))
                clean_panel["w"] = max(float(clean_panel.get("w", 0.0)), 0.0)
                clean_panel["h"] = max(float(clean_panel.get("h", 0.0)), 0.0)
                clean_panel["rotation_deg"] = float(clean_panel.get("rotation_deg", 0.0))
            except (TypeError, ValueError):
                clean_panel["x"] = 0.0
                clean_panel["y"] = 0.0
                clean_panel["w"] = 0.0
                clean_panel["h"] = 0.0
                clean_panel["rotation_deg"] = 0.0

            panel_rect = Rect(
                x=clean_panel["x"],
                y=clean_panel["y"],
                w=clean_panel["w"],
                h=clean_panel["h"],
            )

            if panel_rect.w <= 0 or panel_rect.h <= 0:
                clean_panel["status"] = "invalid"
            elif not rect_contains(packing_rect, panel_rect):
                clean_panel["status"] = "invalid"
            elif any(rects_intersect(panel_rect, obstacle_rect) for obstacle_rect in obstacle_rects):
                clean_panel["status"] = "blocked"
            else:
                clean_panel["status"] = "auto"

            cleaned_panels.append(clean_panel)
            panel_rects.append(panel_rect)

        overlapping_panel_indices = set()
        for i in range(len(cleaned_panels)):
            if cleaned_panels[i]["status"] not in {"auto", "user"}:
                continue
            for j in range(i + 1, len(cleaned_panels)):
                if cleaned_panels[j]["status"] not in {"auto", "user"}:
                    continue
                if rects_intersect(panel_rects[i], panel_rects[j]):
                    overlapping_panel_indices.add(i)
                    overlapping_panel_indices.add(j)

        for idx in overlapping_panel_indices:
            cleaned_panels[idx]["status"] = "invalid"

        plane["panels"] = cleaned_panels

    return editor_state


def get_editor_metrics(
    editor_state: dict,
    module_power_kwp: float,
) -> dict:
    fitted_panels = 0
    blocked_panels = 0
    invalid_panels = 0
    total_panels = 0
    fitted_panels_by_plane = {}

    for plane in editor_state["planes"]:
        plane_id = plane["plane_id"]
        fitted_count = 0

        for panel in plane.get("panels", []):
            total_panels += 1
            status = panel.get("status", "auto")

            if status in {"auto", "user"}:
                fitted_panels += 1
                fitted_count += 1
            elif status == "blocked":
                blocked_panels += 1
            else:
                invalid_panels += 1

        fitted_panels_by_plane[plane_id] = fitted_count

    return {
        "total_panels": total_panels,
        "fitted_panels": fitted_panels,
        "blocked_panels": blocked_panels,
        "invalid_panels": invalid_panels,
        "fitted_kwp": fitted_panels * module_power_kwp,
        "fitted_panels_by_plane": fitted_panels_by_plane,
    }

def build_plane_table_rows(
    roof_planes: list[RoofPlaneGeometry],
    display_panel_counts_for_planes: list[int],
    plane_layouts: dict[str, PanelLayout],
) -> list[dict]:
    plane_rows = []

    for plane, installed_panels in zip(roof_planes, display_panel_counts_for_planes):
        plane_rows.append(
            {
                "Plane": plane.name,
                "Azimuth (deg)": f"{plane.azimuth_deg:.0f}",
                "Roof tilt (deg)": f"{plane.tilt_deg:.0f}",
                "Gross length (m)": f"{plane.gross_length_m:.2f}",
                "Gross depth (m)": f"{plane.gross_depth_m:.2f}",
                "Left margin / setback (m)": f"{plane.margin_left_m:.2f}",
                "Right margin / setback (m)": f"{plane.margin_right_m:.2f}",
                "Top margin / setback (m)": f"{plane.margin_top_m:.2f}",
                "Bottom margin / setback (m)": f"{plane.margin_bottom_m:.2f}",
                "Usable length (m)": f"{plane.usable_length_m:.2f}",
                "Usable depth (m)": f"{plane.usable_depth_m:.2f}",
                "Maximum feasible panels": f"{plane_layouts[plane.plane_id].count}",
                "Panels associated with plane": f"{installed_panels}",
            }
        )

    return plane_rows
# -----------------------------------------------------------------------------
# Visual editor helpers
# -----------------------------------------------------------------------------
def get_plane_state(editor_state: dict, plane_id: str) -> dict | None:
    for plane in editor_state["planes"]:
        if plane["plane_id"] == plane_id:
            return plane
    return None


def persist_editor_source_state(editor_state: dict) -> None:
    st.session_state["section2_editor_source_state"] = deepcopy(editor_state)
    st.rerun()


def add_visual_obstacle_to_plane(
    editor_state: dict,
    plane_id: str,
    obstacle_type: str,
    width_m: float,
    height_m: float,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    packing_rect = dict_to_rect(plane["packing_rect"])
    usable_rect = dict_to_rect(plane["usable_rect"])

    width_m = min(max(width_m, 0.01), usable_rect.w)
    height_m = min(max(height_m, 0.01), usable_rect.h)

    new_obstacle = {
        "obstacle_id": make_obstacle_id(
            plane_id=plane_id,
            existing_obstacle_ids=[o["obstacle_id"] for o in plane.get("obstacles", [])],
        ),
        "plane_id": plane_id,
        "x": packing_rect.x,
        "y": packing_rect.y,
        "w": width_m,
        "h": height_m,
        "obstacle_type": obstacle_type,
    }
    new_obstacle = clamp_obstacle_to_rect(new_obstacle, usable_rect)

    plane.setdefault("obstacles", []).append(new_obstacle)
    return editor_state


def move_obstacle_in_plane(
    editor_state: dict,
    plane_id: str,
    obstacle_id: str,
    dx: float,
    dy: float,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    usable_rect = dict_to_rect(plane["usable_rect"])

    for obstacle in plane.get("obstacles", []):
        if obstacle["obstacle_id"] == obstacle_id:
            obstacle["x"] = float(obstacle["x"]) + dx
            obstacle["y"] = float(obstacle["y"]) + dy
            clamped = clamp_obstacle_to_rect(obstacle, usable_rect)
            obstacle.update(clamped)
            break

    return editor_state


def delete_obstacle_from_plane(editor_state: dict, plane_id: str, obstacle_id: str) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    plane["obstacles"] = [obs for obs in plane.get("obstacles", []) if obs["obstacle_id"] != obstacle_id]
    return editor_state


def update_obstacle_size_type(
    editor_state: dict,
    plane_id: str,
    obstacle_id: str,
    new_width: float,
    new_height: float,
    new_type: str,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    usable_rect = dict_to_rect(plane["usable_rect"])

    for obstacle in plane.get("obstacles", []):
        if obstacle["obstacle_id"] == obstacle_id:
            obstacle["w"] = min(max(float(new_width), 0.01), usable_rect.w)
            obstacle["h"] = min(max(float(new_height), 0.01), usable_rect.h)
            obstacle["obstacle_type"] = new_type
            clamped = clamp_obstacle_to_rect(obstacle, usable_rect)
            obstacle.update(clamped)
            break

    return editor_state


def add_visual_zone_to_plane(
    editor_state: dict,
    plane_id: str,
    width_m: float,
    height_m: float,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    packing_rect = dict_to_rect(plane["packing_rect"])
    width_m = min(max(width_m, 0.01), packing_rect.w)
    height_m = min(max(height_m, 0.01), packing_rect.h)

    zone_id = make_zone_id(
        plane_id=plane_id,
        existing_zone_ids=[z["zone_id"] for z in plane.get("pv_zones", [])],
    )

    new_zone = {
        "zone_id": zone_id,
        "plane_id": plane_id,
        "x": packing_rect.x,
        "y": packing_rect.y,
        "w": width_m,
        "h": height_m,
        "label": f"Array zone {len(plane.get('pv_zones', [])) + 1}",
        "status": "valid",
    }
    new_zone = clamp_pv_zone_to_rect(new_zone, packing_rect)

    plane.setdefault("pv_zones", []).append(new_zone)
    return editor_state


def move_zone_in_plane(
    editor_state: dict,
    plane_id: str,
    zone_id: str,
    dx: float,
    dy: float,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    packing_rect = dict_to_rect(plane["packing_rect"])

    for zone in plane.get("pv_zones", []):
        if zone["zone_id"] == zone_id:
            zone["x"] = float(zone["x"]) + dx
            zone["y"] = float(zone["y"]) + dy
            zone.update(clamp_pv_zone_to_rect(zone, packing_rect))
            break

    return editor_state


def update_zone_properties(
    editor_state: dict,
    plane_id: str,
    zone_id: str,
    new_width: float,
    new_height: float,
    new_label: str,
) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    packing_rect = dict_to_rect(plane["packing_rect"])

    for zone in plane.get("pv_zones", []):
        if zone["zone_id"] == zone_id:
            zone["w"] = min(max(float(new_width), 0.01), packing_rect.w)
            zone["h"] = min(max(float(new_height), 0.01), packing_rect.h)
            zone["label"] = str(new_label).strip() or zone["zone_id"]
            zone.update(clamp_pv_zone_to_rect(zone, packing_rect))
            break

    return editor_state


def delete_zone_from_plane(editor_state: dict, plane_id: str, zone_id: str) -> dict:
    plane = get_plane_state(editor_state, plane_id)
    if plane is None:
        return editor_state

    plane["pv_zones"] = [z for z in plane.get("pv_zones", []) if z["zone_id"] != zone_id]
    return editor_state


def build_visual_obstacle_editor(editor_state: dict) -> dict:
    editor_state = deepcopy(editor_state)
    planes = editor_state.get("planes", [])
    roof_form = str(editor_state.get("roof_form", "Flat"))

    if not planes:
        return editor_state

    show_plane_selector = roof_form == "Duo-pitch" and len(planes) > 1

    if show_plane_selector:
        top_cols = st.columns([2.2, 1.0])

        plane_options = {plane["name"]: plane["plane_id"] for plane in planes}

        with top_cols[0]:
            selected_plane_name = st.selectbox(
                "Plane to edit",
                list(plane_options.keys()),
                key="visual_obstacle_plane_select",
            )

        selected_plane_id = plane_options[selected_plane_name]
        selected_plane = get_plane_state(editor_state, selected_plane_id) or planes[0]

        with top_cols[1]:
            clear_widget_state_if_outside_range(
                key=f"editor_move_step_{selected_plane_id}",
                min_value=0.01,
                max_value=1.0,
            )
            move_step = st.number_input(
                "Movement step (m)",
                min_value=0.01,
                max_value=1.0,
                value=OBSTACLE_NUDGE_STEP_DEFAULT,
                step=0.01,
                key=f"editor_move_step_{selected_plane_id}",
            )
    else:
        selected_plane = planes[0]
        selected_plane_id = selected_plane["plane_id"]

        clear_widget_state_if_outside_range(
            key=f"editor_move_step_{selected_plane_id}",
            min_value=0.01,
            max_value=1.0,
        )
        move_step = st.number_input(
            "Movement step (m)",
            min_value=0.01,
            max_value=1.0,
            value=OBSTACLE_NUDGE_STEP_DEFAULT,
            step=0.01,
            key=f"editor_move_step_{selected_plane_id}",
        )

    selected_usable_rect = dict_to_rect(selected_plane["usable_rect"])
    selected_packing_rect = dict_to_rect(selected_plane["packing_rect"])

    if (
        selected_usable_rect.w < 0.01
        or selected_usable_rect.h < 0.01
        or selected_packing_rect.w < 0.01
        or selected_packing_rect.h < 0.01
    ):
        st.warning("The usable roof area is too small for the visual editor controls.")
        return editor_state

    default_obstacle_width = min(1.00, selected_usable_rect.w)
    default_obstacle_height = min(1.00, selected_usable_rect.h)
    default_zone_width = min(2.50, selected_packing_rect.w)
    default_zone_height = min(2.50, selected_packing_rect.h)

    st.markdown("**Step 1: Obstacles (optional)**")

    obstacle_add_cols = st.columns(4)

    with obstacle_add_cols[0]:
        new_obstacle_type = st.selectbox(
            "Obstacle type",
            ["generic", "window", "vent", "plant"],
            key=f"new_obstacle_type_{selected_plane_id}",
        )

    with obstacle_add_cols[1]:
        clear_widget_state_if_outside_range(
            key=f"new_obstacle_width_{selected_plane_id}",
            min_value=0.01,
            max_value=float(selected_usable_rect.w),
        )
        new_obstacle_width = st.number_input(
            "Obstacle width (m)",
            min_value=0.01,
            max_value=float(selected_usable_rect.w),
            value=float(default_obstacle_width),
            step=0.05,
            key=f"new_obstacle_width_{selected_plane_id}",
        )

    with obstacle_add_cols[2]:
        clear_widget_state_if_outside_range(
            key=f"new_obstacle_height_{selected_plane_id}",
            min_value=0.01,
            max_value=float(selected_usable_rect.h),
        )
        new_obstacle_height = st.number_input(
            "Obstacle height (m)",
            min_value=0.01,
            max_value=float(selected_usable_rect.h),
            value=float(default_obstacle_height),
            step=0.05,
            key=f"new_obstacle_height_{selected_plane_id}",
        )

    with obstacle_add_cols[3]:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        if st.button(
            "Add obstacle",
            key=f"add_visual_obstacle_button_{selected_plane_id}",
            width="stretch",
        ):
            editor_state = add_visual_obstacle_to_plane(
                editor_state=editor_state,
                plane_id=selected_plane_id,
                obstacle_type=new_obstacle_type,
                width_m=float(new_obstacle_width),
                height_m=float(new_obstacle_height),
            )
            persist_editor_source_state(editor_state)

    selected_plane = get_plane_state(editor_state, selected_plane_id) or selected_plane
    user_obstacles = selected_plane.get("obstacles", [])

    if user_obstacles:
        obstacle_lookup = {
            obs["obstacle_id"]: obs
            for obs in user_obstacles
        }

        obstacle_cols = st.columns([1.8, 1.0, 1.0, 1.0, 1.0, 1.0])

        with obstacle_cols[0]:
            selected_obstacle_key = f"selected_obstacle_{selected_plane_id}"
            obstacle_options = list(obstacle_lookup.keys())

            clear_widget_state_if_not_in_options(
                key=selected_obstacle_key,
                valid_options=obstacle_options,
            )

            selected_obstacle_id = st.selectbox(
                "Obstacle to edit",
                options=obstacle_options,
                format_func=lambda oid: f"{oid} ({obstacle_lookup[oid]['obstacle_type']})",
                key=selected_obstacle_key,
            )

        obstacle = obstacle_lookup[selected_obstacle_id]

        obstacle_type_options = ["generic", "window", "vent", "plant"]
        obstacle_type_value = obstacle.get("obstacle_type", "generic")
        if obstacle_type_value not in obstacle_type_options:
            obstacle_type_value = "generic"

        with obstacle_cols[1]:
            obs_type = st.selectbox(
                "Type",
                obstacle_type_options,
                index=obstacle_type_options.index(obstacle_type_value),
                key=f"obs_type_edit_{selected_obstacle_id}",
            )

        with obstacle_cols[2]:
            clear_widget_state_if_outside_range(
                key=f"obs_width_edit_{selected_obstacle_id}",
                min_value=0.01,
                max_value=float(selected_usable_rect.w),
            )
            obs_w = st.number_input(
                "Width (m)",
                min_value=0.01,
                max_value=float(selected_usable_rect.w),
                value=min(float(obstacle["w"]), float(selected_usable_rect.w)),
                step=0.05,
                key=f"obs_width_edit_{selected_obstacle_id}",
            )

        with obstacle_cols[3]:
            clear_widget_state_if_outside_range(
                key=f"obs_height_edit_{selected_obstacle_id}",
                min_value=0.01,
                max_value=float(selected_usable_rect.h),
            )
            obs_h = st.number_input(
                "Height (m)",
                min_value=0.01,
                max_value=float(selected_usable_rect.h),
                value=min(float(obstacle["h"]), float(selected_usable_rect.h)),
                step=0.05,
                key=f"obs_height_edit_{selected_obstacle_id}",
            )

        with obstacle_cols[4]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button(
                "Apply changes",
                key=f"apply_obstacle_changes_{selected_obstacle_id}",
                width="stretch",
            ):
                editor_state = update_obstacle_size_type(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                    new_width=float(obs_w),
                    new_height=float(obs_h),
                    new_type=obs_type,
                )
                persist_editor_source_state(editor_state)

        with obstacle_cols[5]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button(
                "Delete selected",
                key=f"delete_obstacle_{selected_obstacle_id}",
                width="stretch",
            ):
                editor_state = delete_obstacle_from_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                )
                persist_editor_source_state(editor_state)

        move_cols = st.columns(4)

        with move_cols[0]:
            if st.button("←", key=f"move_left_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                    dx=-float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)

        with move_cols[1]:
            if st.button("→", key=f"move_right_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                    dx=float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)

        with move_cols[2]:
            if st.button("↑", key=f"move_up_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                    dx=0.0,
                    dy=-float(move_step),
                )
                persist_editor_source_state(editor_state)

        with move_cols[3]:
            if st.button("↓", key=f"move_down_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    obstacle_id=selected_obstacle_id,
                    dx=0.0,
                    dy=float(move_step),
                )
                persist_editor_source_state(editor_state)
    else:
        st.info("No user-added obstacles defined.")

    st.markdown("---")
    st.markdown("**Step 2: Array zones**")

    zone_add_cols = st.columns(4)

    with zone_add_cols[0]:
        clear_widget_state_if_outside_range(
            key=f"new_zone_width_{selected_plane_id}",
            min_value=0.01,
            max_value=float(selected_packing_rect.w),
        )
        new_zone_width = st.number_input(
            "Array zone width (m)",
            min_value=0.01,
            max_value=float(selected_packing_rect.w),
            value=float(default_zone_width),
            step=0.05,
            key=f"new_zone_width_{selected_plane_id}",
        )

    with zone_add_cols[1]:
        clear_widget_state_if_outside_range(
            key=f"new_zone_height_{selected_plane_id}",
            min_value=0.01,
            max_value=float(selected_packing_rect.h),
        )
        new_zone_height = st.number_input(
            "Array zone height (m)",
            min_value=0.01,
            max_value=float(selected_packing_rect.h),
            value=float(default_zone_height),
            step=0.05,
            key=f"new_zone_height_{selected_plane_id}",
        )

    with zone_add_cols[2]:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        if st.button(
            "Add array zone",
            key=f"add_zone_button_{selected_plane_id}",
            width="stretch",
        ):
            editor_state = add_visual_zone_to_plane(
                editor_state=editor_state,
                plane_id=selected_plane_id,
                width_m=float(new_zone_width),
                height_m=float(new_zone_height),
            )
            persist_editor_source_state(editor_state)

    with zone_add_cols[3]:
        st.empty()

    selected_plane = get_plane_state(editor_state, selected_plane_id) or selected_plane
    zones = selected_plane.get("pv_zones", [])

    if zones:
        zone_lookup = {
            zone["zone_id"]: zone
            for zone in zones
        }

        zone_cols = st.columns([1.8, 1.2, 1.0, 1.0, 1.0, 1.0])

        with zone_cols[0]:
            selected_zone_key = f"selected_zone_{selected_plane_id}"
            zone_options = list(zone_lookup.keys())

            clear_widget_state_if_not_in_options(
                key=selected_zone_key,
                valid_options=zone_options,
            )

            selected_zone_id = st.selectbox(
                "Array zone to edit",
                options=zone_options,
                format_func=lambda zid: zone_lookup[zid].get("label", zid) or zid,
                key=selected_zone_key,
            )

        zone = zone_lookup[selected_zone_id]

        with zone_cols[1]:
            zone_label = st.text_input(
                "Label",
                value=zone.get("label", ""),
                key=f"zone_label_{selected_zone_id}",
            )

        with zone_cols[2]:
            clear_widget_state_if_outside_range(
                key=f"zone_w_{selected_zone_id}",
                min_value=0.01,
                max_value=float(selected_packing_rect.w),
            )
            zone_w = st.number_input(
                "Width (m)",
                min_value=0.01,
                max_value=float(selected_packing_rect.w),
                value=min(float(zone["w"]), float(selected_packing_rect.w)),
                step=0.05,
                key=f"zone_w_{selected_zone_id}",
            )

        with zone_cols[3]:
            clear_widget_state_if_outside_range(
                key=f"zone_h_{selected_zone_id}",
                min_value=0.01,
                max_value=float(selected_packing_rect.h),
            )
            zone_h = st.number_input(
                "Height (m)",
                min_value=0.01,
                max_value=float(selected_packing_rect.h),
                value=min(float(zone["h"]), float(selected_packing_rect.h)),
                step=0.05,
                key=f"zone_h_{selected_zone_id}",
            )

        with zone_cols[4]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button(
                "Apply changes",
                key=f"zone_apply_{selected_zone_id}",
                width="stretch",
            ):
                editor_state = update_zone_properties(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                    new_width=float(zone_w),
                    new_height=float(zone_h),
                    new_label=zone_label,
                )
                persist_editor_source_state(editor_state)

        with zone_cols[5]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button(
                "Delete selected",
                key=f"zone_delete_{selected_zone_id}",
                width="stretch",
            ):
                editor_state = delete_zone_from_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                )
                persist_editor_source_state(editor_state)

        zone_move_cols = st.columns(4)

        with zone_move_cols[0]:
            if st.button("←", key=f"zone_left_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                    dx=-float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)

        with zone_move_cols[1]:
            if st.button("→", key=f"zone_right_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                    dx=float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)

        with zone_move_cols[2]:
            if st.button("↑", key=f"zone_up_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                    dx=0.0,
                    dy=-float(move_step),
                )
                persist_editor_source_state(editor_state)

        with zone_move_cols[3]:
            if st.button("↓", key=f"zone_down_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state=editor_state,
                    plane_id=selected_plane_id,
                    zone_id=selected_zone_id,
                    dx=0.0,
                    dy=float(move_step),
                )
                persist_editor_source_state(editor_state)
    else:
        st.warning("No array zones defined.")

    return editor_state
# -----------------------------------------------------------------------------
# Diagram helpers
# -----------------------------------------------------------------------------

def get_plane_label_text(roof_form: str, plane_name: str) -> str:
    roof_type_label = f"{roof_form} roof"
    if roof_form in {"Flat", "Mono-pitch"}:
        return roof_type_label
    return f"{roof_type_label} - {plane_name}"


def build_editor_roof_packing_diagram(editor_state: dict) -> go.Figure:
    fig = go.Figure()

    planes = editor_state.get("planes", [])
    roof_form = editor_state.get("roof_form", "Flat")
    flat_panel_pitch_deg = float(editor_state["module"].get("flat_panel_pitch_deg", DEFAULT_FLAT_PANEL_PITCH_DEG))

    if not planes:
        fig.update_layout(
            height=250,
            margin=DIAGRAM_MARGIN,
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
        )
        return fig

    max_x = 0.0
    max_y = 0.0

    def draw_one_plane(plane: dict, origin_x: float, origin_y: float) -> tuple[float, float]:
        gross = dict_to_rect(plane["gross_rect"])
        usable = dict_to_rect(plane["usable_rect"])
        packing = dict_to_rect(plane["packing_rect"])

        roof_x0 = origin_x + gross.x
        roof_y0 = origin_y + gross.y
        roof_x1 = roof_x0 + gross.w
        roof_y1 = roof_y0 + gross.h

        usable_x0 = origin_x + usable.x
        usable_y0 = origin_y + usable.y
        usable_x1 = usable_x0 + usable.w
        usable_y1 = usable_y0 + usable.h

        packing_x0 = origin_x + packing.x
        packing_y0 = origin_y + packing.y
        packing_x1 = packing_x0 + packing.w
        packing_y1 = packing_y0 + packing.h

        fig.add_shape(
            type="rect",
            x0=roof_x0,
            y0=roof_y0,
            x1=roof_x1,
            y1=roof_y1,
            line=dict(color=DIAGRAM_LINE_COLOUR, width=2),
            fillcolor=DIAGRAM_ROOF_COLOUR,
        )

        fig.add_shape(
            type="rect",
            x0=usable_x0,
            y0=usable_y0,
            x1=usable_x1,
            y1=usable_y1,
            line=dict(color=DIAGRAM_LINE_COLOUR, width=1, dash="dash"),
            fillcolor=DIAGRAM_USABLE_COLOUR,
        )

        fig.add_shape(
            type="rect",
            x0=packing_x0,
            y0=packing_y0,
            x1=packing_x1,
            y1=packing_y1,
            line=dict(color=DIAGRAM_LINE_COLOUR, width=1, dash="dot"),
            fillcolor="rgba(0,0,0,0)",
        )

        gross_w = gross.w
        gross_h = gross.h
        usable_w = usable.w
        usable_h = usable.h

        label_text = get_plane_label_text(roof_form, plane["name"])

        if roof_form == "Flat":
            detail_text = (
                f"Azimuth 0° roof tilt 0°<br>"
                f"Panel pitch above horizontal {flat_panel_pitch_deg:.0f}°<br>"
                f"Gross {gross_w:.2f} × {gross_h:.2f} m<br>"
                f"Usable {usable_w:.2f} × {usable_h:.2f} m<br>"
                f"Array zones {len(plane.get('pv_zones', []))} / max {plane['max_feasible_panels']} panels"
            )
        else:
            detail_text = (
                f"Azimuth {float(plane['azimuth_deg']):.0f}° roof tilt {float(plane['tilt_deg']):.0f}°<br>"
                f"Gross {gross_w:.2f} × {gross_h:.2f} m<br>"
                f"Usable {usable_w:.2f} × {usable_h:.2f} m<br>"
                f"Array zones {len(plane.get('pv_zones', []))} / max {plane['max_feasible_panels']} panels"
            )

        fig.add_annotation(
            x=(roof_x0 + roof_x1) / 2.0,
            y=roof_y0 - 0.65,
            showarrow=False,
            text=f"<b>{label_text}</b>",
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=13),
        )
        fig.add_annotation(
            x=(roof_x0 + roof_x1) / 2.0,
            y=roof_y1 + 0.22,
            showarrow=False,
            text=f"{gross_w:.2f} m",
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=11),
        )
        fig.add_annotation(
            x=roof_x1 + 0.25,
            y=(roof_y0 + roof_y1) / 2.0,
            showarrow=False,
            text=f"{gross_h:.2f} m",
            textangle=90,
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=11),
        )
        fig.add_annotation(
            x=(roof_x0 + roof_x1) / 2.0,
            y=roof_y1 + 1.05,
            showarrow=False,
            text=detail_text,
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=11),
            align="center",
        )

        for obstacle in plane.get("obstacles", []):
            ox0 = origin_x + float(obstacle["x"])
            oy0 = origin_y + float(obstacle["y"])
            ox1 = ox0 + float(obstacle["w"])
            oy1 = oy0 + float(obstacle["h"])

            fig.add_shape(
                type="rect",
                x0=ox0,
                y0=oy0,
                x1=ox1,
                y1=oy1,
                line=dict(color=DIAGRAM_LINE_COLOUR, width=1),
                fillcolor=DIAGRAM_OBSTACLE_COLOUR,
            )

        for zone in plane.get("pv_zones", []):
            zx0 = origin_x + float(zone["x"])
            zy0 = origin_y + float(zone["y"])
            zx1 = zx0 + float(zone["w"])
            zy1 = zy0 + float(zone["h"])

            fig.add_shape(
                type="rect",
                x0=zx0,
                y0=zy0,
                x1=zx1,
                y1=zy1,
                line=dict(color=DIAGRAM_ZONE_LINE_COLOUR, width=2, dash="dot"),
                fillcolor=DIAGRAM_ZONE_FILL_COLOUR,
            )

        for panel in plane.get("panels", []):
            px0 = origin_x + float(panel["x"])
            py0 = origin_y + float(panel["y"])
            px1 = px0 + float(panel["w"])
            py1 = py0 + float(panel["h"])

            status = panel.get("status", "auto")
            fill = DIAGRAM_PANEL_COLOUR
            if status == "user":
                fill = DIAGRAM_PANEL_USER_COLOUR
            elif status == "blocked":
                fill = DIAGRAM_PANEL_BLOCKED_COLOUR
            elif status == "invalid":
                fill = DIAGRAM_PANEL_INVALID_COLOUR

            fig.add_shape(
                type="rect",
                x0=px0,
                y0=py0,
                x1=px1,
                y1=py1,
                line=dict(color=DIAGRAM_LINE_COLOUR, width=1),
                fillcolor=fill,
            )

        fig.add_trace(
            go.Scatter(
                x=[(roof_x0 + roof_x1) / 2.0],
                y=[(roof_y0 + roof_y1) / 2.0],
                mode="markers",
                marker=dict(size=18, color="rgba(0,0,0,0)"),
                hovertemplate=label_text + "<extra></extra>",
                showlegend=False,
            )
        )

        return roof_x1, roof_y1 + 1.6

    if roof_form == "Duo-pitch" and len(planes) == 2:
        current_x = 0.0
        for plane in planes:
            roof_x1, content_y1 = draw_one_plane(plane, current_x, 0.0)
            max_x = max(max_x, roof_x1)
            max_y = max(max_y, content_y1)
            current_x = roof_x1 + DIAGRAM_PLANE_GAP_X
    else:
        current_y = 0.0
        for plane in planes:
            roof_x1, content_y1 = draw_one_plane(plane, 0.0, current_y)
            max_x = max(max_x, roof_x1)
            max_y = max(max_y, content_y1)
            current_y = content_y1 + DIAGRAM_PLANE_GAP_Y

    fig.update_layout(
        height=DIAGRAM_DUO_HEIGHT if roof_form == "Duo-pitch" else DIAGRAM_SINGLE_HEIGHT,
        margin=DIAGRAM_MARGIN,
        paper_bgcolor="white",
        plot_bgcolor="white",
        showlegend=False,
        hovermode="closest",
    )

    fig.update_xaxes(
        visible=False,
        range=[-0.8, max_x + 1.2],
        scaleanchor="y",
        scaleratio=1,
        fixedrange=False,
    )
    fig.update_yaxes(
        visible=False,
        range=[max_y + 0.8, -1.5],
        fixedrange=False,
    )

    return fig



# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
header_left, header_right = st.columns([5, 1.5])

with header_left:
    st.title("Part L 2026 Photovoltaic Array Calculator")

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
# Public-facing introduction
# -----------------------------------------------------------------------------
st.markdown(
    """
Roof-mounted solar PV is expected to play an important role in meeting the Future Homes Standard / Part L 2026 requirements. 
This early Etude PV tool provides an initial estimate of the photovoltaic array capacity likely to be needed under the emerging Part L 2026 approach.

The tool is intended as a simple guide rather than a substitute for full compliance modelling. Please carry out your own checks and let us know if anything looks inconsistent, so that we can review and improve the tool.

If you have questions about Part L 2026, SAP 10.3 or the Home Energy Model (HEM), contact Etude at london@etude.co.uk. We can help assess how your design is likely to perform, including energy demand, energy use and renewable energy generation.
"""
)

# -----------------------------------------------------------------------------
# Method information
# -----------------------------------------------------------------------------
with st.expander("Method summary", expanded=False):
    st.markdown(
        """
This tool is split into four main sections.

**Dwelling inputs** captures the house form and ground floor area. The ground floor area can be entered directly or derived using a simplified geometry helper.

**Part L photovoltaic target** calculates the target photovoltaic array capacity from ground floor area. The current target assumption is that photovoltaic capacity is based on 40% of ground floor area at 0.22 kWp/m².

The Part L target check is based on installed photovoltaic capacity in kWp. Annual generation is estimated separately and is not used for the Part L capacity target check.

**Photovoltaic array layout** creates the array capacity that is passed forward to the generation calculation. Two input routes are available:

- **Visual roof layout**: uses simplified roof geometry, array zones and optional obstacles to fit panels.
- **Manual array input**: bypasses roof geometry where the array capacity, orientation, tilt and shading are already known.

The visual layout route currently uses simplified roof-plane geometry:

- **Flat**: one rectangular roof in plan with user-defined panel pitch above horizontal.
- **Mono-pitch**: one sloping plane derived from plan dimensions and roof pitch.
- **Duo-pitch**: two identical sloping planes, with the second rotated by 180°.

The array layout editor works in two stages:

- obstacles can be added first, but this step is optional;
- array zones are then defined within the usable packing area;
- panels are regenerated automatically within those array zones;
- movement and resizing are handled through native Streamlit controls;
- panel validity is checked against the packing area, obstacle intersections and panel overlap.

**Photovoltaic array energy generation** reuses the arrays from the layout section. It can estimate annual generation using either:

- **PySAM PVWatts**, with a selected local EPW weather file; or
- **SAP Appendix U**, using coded monthly irradiance, declination, representative latitude and orientation constants.

The generation result is reported separately in kWh/year.
"""
    )

with st.expander("Current limits", expanded=False):
    st.markdown(
        """
- The Part L target basis is a working assumption and should be checked against the final approved Part L 2026 / SAP methodology when published.
- The SAP / weather regions are broad app-level regions and are not yet postcode-district precise.
- The Appendix U implementation uses representative SAP climate regions rather than a full postcode-to-SAP-region lookup.
- The roof fit still uses simplified roof reductions rather than a full geometric roof model.
- Flat roofs use a single roof plane in the editor, while generation is still split 50/50 east-west.
- Flat-roof row spacing still needs to be developed properly.
- Hipped roofs are not included in this simplified method.
- Array zones are rectangular only and do not yet support polygon drawing.
- Panels are regenerated from array zones rather than dragged individually.
- The editor uses native Streamlit controls for movement and resizing.
- PySAM generation uses locally stored EPW files in `resources/epw/`.
- SAP Appendix U and PySAM generation are deliberately separated from the Part L photovoltaic capacity target check.
"""
    )

# -----------------------------------------------------------------------------
# Dwelling inputs
# -----------------------------------------------------------------------------
render_section_title("dwelling_inputs", "Dwelling inputs")
with st.container(border=True):
    dwelling_top = st.columns(2)

    with dwelling_top[0]:
        house_form = st.selectbox(
            "House form",
            ["Detached", "Semi-detached", "End terrace", "Mid terrace"],
            index=0,
            key="dwelling_house_form",
        )

    with dwelling_top[1]:
        gfa_input_mode = st.selectbox(
            "Ground floor area method",
            ["Enter explicitly", "Derive from geometry"],
            index=0,
            key="dwelling_gfa_mode",
        )

    if gfa_input_mode == "Enter explicitly":
        ground_floor_area_m2 = st.slider(
            "Ground floor area (m²)",
            min_value=20.00,
            max_value=500.00,
            value=72.00,
            step=0.01,
            key="dwelling_gfa_direct",
        )
        gfa_source_text = "Entered explicitly"

    else:
        st.caption(
            "Derived GFA is a simplified input helper. Use the explicit GFA input where a measured or assessed value is available."
        )

        gfa_geom_cols = st.columns(3)

        with gfa_geom_cols[0]:
            ridge_parallel_width_for_gfa_m = st.slider(
                "Width parallel to ridge / long side (m)",
                min_value=4.00,
                max_value=25.00,
                value=9.00,
                step=0.01,
                key="dwelling_gfa_width",
            )

        with gfa_geom_cols[1]:
            depth_for_gfa_m = st.slider(
                "Depth perpendicular to ridge / short side (m)",
                min_value=4.00,
                max_value=25.00,
                value=8.00,
                step=0.01,
                key="dwelling_gfa_depth",
            )

        with gfa_geom_cols[2]:
            wall_thickness_m = st.slider(
                "External wall thickness (m)",
                min_value=0.10,
                max_value=0.60,
                value=0.30,
                step=0.01,
                key="dwelling_gfa_wall",
            )

        ground_floor_area_m2 = calc_spreadsheet_gfa(
            width_parallel_to_ridge_m=ridge_parallel_width_for_gfa_m,
            depth_perpendicular_to_ridge_m=depth_for_gfa_m,
            wall_thickness_m=wall_thickness_m,
            house_form=house_form,
        )
        gfa_source_text = "Derived from geometry"


# -----------------------------------------------------------------------------
# Part L target
# -----------------------------------------------------------------------------
render_section_title("part_l_target", "Part L photovoltaic target")
with st.container(border=True):
    st.caption(
        "Indicative photovoltaic array capacity target based on ground floor area and the current working Part L 2026 target basis."
    )

    part_l_required_kwp = (
        ground_floor_area_m2
        * FHS_REQUIRED_AREA_FRACTION
        * STANDARD_PANEL_EFFICIENCY_KWP_PER_M2
    )

    standardised_module_power_kwp = module_power_kwp_from_inputs(
        length_m=STANDARDISED_MODULE_LENGTH_M,
        width_m=STANDARDISED_MODULE_WIDTH_M,
        efficiency_pct=STANDARDISED_MODULE_EFFICIENCY_PCT,
    )

    part_l_required_panel_count = math.ceil(
        part_l_required_kwp / standardised_module_power_kwp
    )

    part_l_summary_cols = st.columns(2)

    with part_l_summary_cols[0]:
        render_summary_card(
            "Target photovoltaic array capacity",
            f"{part_l_required_kwp:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>",
        )

    with part_l_summary_cols[1]:
        render_summary_card(
            "Equivalent standardised panel count",
            f"{part_l_required_panel_count:,.0f}",
        )

    target_assumption_rows = [
        ("Required PV area fraction", f"{FHS_REQUIRED_AREA_FRACTION:.2f} of ground floor area"),
        ("Standard panel efficiency density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:.2f} kWp/m²"),
        ("Target capacity formula", "Ground floor area × required PV area fraction × standard panel efficiency density"),
        ("Standardised module size", f"{STANDARDISED_MODULE_LENGTH_M * 1000:.0f} × {STANDARDISED_MODULE_WIDTH_M * 1000:.0f} mm"),
        ("Standardised module efficiency", f"{STANDARDISED_MODULE_EFFICIENCY_PCT:.1f} %"),
    ]

    with st.expander("Show target assumptions", expanded=False):
        st.dataframe(
            pd.DataFrame(target_assumption_rows, columns=["Assumption", "Value"]),
            hide_index=True,
            width="stretch",
        )

# -----------------------------------------------------------------------------
# Building PV layout
# -----------------------------------------------------------------------------
render_section_title("building_pv_layout", "Photovoltaic array layout")
with st.container(border=True):
    array_input_mode = st.radio(
        "Array input method",
        ["Visual roof layout", "Manual array input"],
        index=0,
        horizontal=True,
        key="array_input_mode",
    )

    # Defaults used by downstream summary sections.
    actual_roof_form = "Manual input"
    plan_length_along_ridge_m = 0.0
    plan_length_ridge_to_eaves_m = 0.0
    flat_panel_pitch_deg = float(DEFAULT_FLAT_PANEL_PITCH_DEG)
    mono_or_duo_azimuth_deg = 180.0
    mono_or_duo_pitch_deg = 30.0
    offset_mode_section_2 = "Not applicable"
    simple_setback_m = 0.0
    ridge_offset_m = 0.0
    edge_offset_m = 0.0
    party_wall_offset_m = 0.0
    roof_geometry = RoofGeometryBundle(roof_form="Manual input", planes=[])
    roof_planes = []
    plane_layouts: dict[str, PanelLayout] = {}
    max_feasible_panels = 0
    display_panel_counts_for_planes = []
    actual_arrays = []
    actual_array_panel_counts = []
    installed_panel_count = 0
    module_length_m = DEFAULT_MODULE_LENGTH_M
    module_width_m = FIXED_MODULE_WIDTH_M
    module_efficiency_pct = DEFAULT_MODULE_EFFICIENCY_PCT
    module_mount_orientation = "Portrait"
    module_power_kwp = module_power_kwp_from_inputs(
        length_m=module_length_m,
        width_m=module_width_m,
        efficiency_pct=module_efficiency_pct,
    )
    module_power_wp = module_power_kwp * 1000.0
    actual_building_kwp = 0.0
    actual_kwp_status = "Shortfall"
    actual_panel_status = "Shortfall"

    editor_metrics = {
        "total_panels": 0,
        "fitted_panels": 0,
        "blocked_panels": 0,
        "invalid_panels": 0,
        "fitted_kwp": 0.0,
        "fitted_panels_by_plane": {},
    }

    editor_kwp_status = "Shortfall"
    editor_panel_status = "Shortfall"
    roof_editor_state = {
        "schema_version": "0.2.0",
        "roof_form": "Manual input",
        "module": {
            "length_m": module_length_m,
            "width_m": module_width_m,
            "efficiency_pct": module_efficiency_pct,
            "mount_orientation": module_mount_orientation,
            "flat_panel_pitch_deg": flat_panel_pitch_deg,
        },
        "planes": [],
    }

    if array_input_mode == "Visual roof layout":
        building_top = st.columns([1.15, 1.0])
        with building_top[0]:
            actual_roof_form = st.selectbox(
                "Roof type",
                ["Mono-pitch", "Duo-pitch", "Flat"],
                index=1,
                key="actual_roof_form",
            )
        with building_top[1]:
            helper_text = {
                "Flat": "Flat roof uses one whole rectangular roof in plan.",
                "Mono-pitch": "Mono-pitch uses one rectangular roof plane defined in plan.",
                "Duo-pitch": "Duo-pitch duplicates one entered roof plane and rotates the second by 180°.",
            }[actual_roof_form]
            st.markdown(
                f"<div style='padding-top:2px; line-height:38px;'>{helper_text}</div>",
                unsafe_allow_html=True,
            )

        st.markdown("**Roof parameters**")

        if actual_roof_form == "Flat":
            roof_geom_cols = st.columns(3)
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
            with roof_geom_cols[2]:
                flat_panel_pitch_deg = st.slider(
                    "Panel pitch above horizontal (degrees)",
                    min_value=1,
                    max_value=45,
                    value=int(DEFAULT_FLAT_PANEL_PITCH_DEG),
                    step=1,
                    key="flat_panel_pitch",
                )

            mono_or_duo_azimuth_deg = None
            mono_or_duo_pitch_deg = 0.0

            st.info(
                "Flat roof assumption: the roof itself is horizontal. Panels are split 50/50 east-west for generation. "
                "The panel pitch set above is applied to the PV arrays. TODO: row spacing still needs to be thought through."
            )
        else:
            flat_panel_pitch_deg = float(DEFAULT_FLAT_PANEL_PITCH_DEG)

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

        reduction_row = st.columns(4)
        with reduction_row[0]:
            offset_mode_section_2 = st.selectbox(
                "Usable roof area method",
                ["Simple setback", "Detailed offsets"],
                index=0,
                key="actual_offset_mode",
            )

        if offset_mode_section_2 == "Simple setback":
            with reduction_row[1]:
                simple_setback_m = st.number_input(
                    "Setback around usable array area (m)",
                    min_value=0.0,
                    max_value=2.0,
                    value=DEFAULT_SIMPLE_SETBACK_M,
                    step=0.05,
                    key="actual_simple_setback",
                )
            with reduction_row[2]:
                st.empty()
            with reduction_row[3]:
                st.empty()

            ridge_offset_m = 0.0
            edge_offset_m = 0.0
            party_wall_offset_m = 0.0
        else:
            with reduction_row[1]:
                ridge_offset_m = st.number_input(
                    "Ridge offset (m)",
                    min_value=0.0,
                    max_value=2.0,
                    value=DEFAULT_RIDGE_OFFSET_M,
                    step=0.05,
                    key="actual_ridge_offset",
                )
            with reduction_row[2]:
                edge_offset_m = st.number_input(
                    "Roof edge offset (m)",
                    min_value=0.0,
                    max_value=2.0,
                    value=DEFAULT_EDGE_OFFSET_M,
                    step=0.05,
                    key="actual_edge_offset",
                )
            with reduction_row[3]:
                party_wall_offset_m = st.number_input(
                    "Party wall offset (m)",
                    min_value=0.0,
                    max_value=2.0,
                    value=DEFAULT_PARTY_WALL_OFFSET_M,
                    step=0.05,
                    key="actual_party_wall_offset",
                )

            simple_setback_m = 0.0
        roof_geometry = build_roof_geometry(
            roof_form=actual_roof_form,
            plan_length_along_ridge_m=plan_length_along_ridge_m,
            plan_length_ridge_to_eaves_m=plan_length_ridge_to_eaves_m,
            pitch_deg=mono_or_duo_pitch_deg,
            azimuth_deg=mono_or_duo_azimuth_deg,
            simple_setback_m=simple_setback_m,
            ridge_offset_m=ridge_offset_m,
            edge_offset_m=edge_offset_m,
            party_wall_offset_m=party_wall_offset_m,
            house_form=house_form,
        )
        roof_planes = roof_geometry.planes

        st.markdown("**Module parameters**")

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

        module_power_kwp = module_power_kwp_from_inputs(
            length_m=module_length_m,
            width_m=module_width_m,
            efficiency_pct=module_efficiency_pct,
        )
        module_power_wp = module_power_kwp * 1000.0

        plane_layouts = {}
        max_feasible_panels = 0
        for plane in roof_planes:
            layout = calc_plane_panel_layout(
                packing_length_m=plane.packing_length_m,
                packing_depth_m=plane.packing_depth_m,
                module_length_m=module_length_m,
                module_width_m=module_width_m,
                mount_orientation=module_mount_orientation,
                roof_form=actual_roof_form,
                flat_panel_pitch_deg=flat_panel_pitch_deg,
            )
            plane_layouts[plane.plane_id] = layout
            max_feasible_panels += layout.count

        actual_arrays = build_actual_arrays_for_generation(
            roof_form=actual_roof_form,
            mono_or_duo_azimuth_deg=mono_or_duo_azimuth_deg,
            mono_or_duo_pitch_deg=mono_or_duo_pitch_deg,
            flat_panel_pitch_deg=flat_panel_pitch_deg,
        )

        if max_feasible_panels < 1:
            st.warning("The usable roof dimensions are too small to fit one module under this dimension-based method.")
            installed_panel_count = 0
            actual_array_panel_counts = [0] * len(actual_arrays)
        else:
            clear_widget_state_if_outside_range(
                key="installed_panel_count",
                min_value=1,
                max_value=max_feasible_panels,
            )

            installed_panel_count = st.slider(
                "Target panel count to fit",
                min_value=1,
                max_value=max_feasible_panels,
                value=max_feasible_panels,
                step=1,
                key="installed_panel_count",
            )

            actual_array_panel_counts = allocate_integer_counts(
                total_count=installed_panel_count,
                share_fractions=[arr.area_share_fraction for arr in actual_arrays],
            )

        display_panel_counts_for_planes = get_plane_displayed_panel_counts(
            roof_form=actual_roof_form,
            installed_panel_count=installed_panel_count,
            actual_array_panel_counts=actual_array_panel_counts,
        )

        default_editor_state = build_roof_editor_state(
            geometry=roof_geometry,
            plane_layouts=plane_layouts,
            displayed_panel_counts=display_panel_counts_for_planes,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            module_efficiency_pct=module_efficiency_pct,
            module_mount_orientation=module_mount_orientation,
            flat_panel_pitch_deg=flat_panel_pitch_deg,
        )

        editor_signature = get_editor_state_signature(
            geometry=roof_geometry,
            plane_layouts=plane_layouts,
            displayed_panel_counts=display_panel_counts_for_planes,
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            module_efficiency_pct=module_efficiency_pct,
            module_mount_orientation=module_mount_orientation,
            flat_panel_pitch_deg=flat_panel_pitch_deg,
        )

        if (
            "section2_editor_source_state" not in st.session_state
            or "section2_editor_signature" not in st.session_state
            or st.session_state["section2_editor_signature"] != editor_signature
        ):
            st.session_state["section2_editor_source_state"] = deepcopy(default_editor_state)
            st.session_state["section2_editor_signature"] = editor_signature

        source_state = deepcopy(st.session_state["section2_editor_source_state"])

        st.markdown("**Interactive array layout editor (beta)**")
        st.caption(
            "Use array zones to define the areas where panels should be fitted. "
            "Leave the default array zone unchanged if the whole usable roof area is available. "
            "Optional obstacles can be added to remove unavailable areas from the layout. "
            "Movement step defaults to 0.10 m."
        )

        source_state = build_visual_obstacle_editor(source_state)
        st.session_state["section2_editor_source_state"] = deepcopy(source_state)

        roof_editor_state = deepcopy(source_state)
        roof_editor_state = apply_obstacles_to_pv_zones(roof_editor_state)
        roof_editor_state = regenerate_panels_from_pv_zones(roof_editor_state)
        roof_editor_state = validate_editor_state(roof_editor_state)

        st.markdown("**Array layout preview**")
        st.plotly_chart(
            build_editor_roof_packing_diagram(roof_editor_state),
            theme=None,
            width="stretch",
            key="section2_editor_chart",
        )

        editor_metrics = get_editor_metrics(
            editor_state=roof_editor_state,
            module_power_kwp=module_power_kwp,
        )

        # The shared object used by downstream sections.
        visual_panel_counts_by_plane = editor_metrics.get("fitted_panels_by_plane", {})
        actual_array_panel_counts = [
            int(visual_panel_counts_by_plane.get(f"plane_{idx}", 0))
            for idx in range(1, len(actual_arrays) + 1)
        ]

        # For flat roofs, one visual plane maps onto two generation arrays.
        # Keep the existing east/west split for generation/array handover.
        if actual_roof_form == "Flat":
            actual_array_panel_counts = allocate_integer_counts(
                total_count=int(editor_metrics["fitted_panels"]),
                share_fractions=[arr.area_share_fraction for arr in actual_arrays],
            )

        pv_arrays = build_pv_arrays_from_visual_layout(
            actual_arrays=actual_arrays,
            actual_array_panel_counts=actual_array_panel_counts,
            module_power_kwp=module_power_kwp,
            source_label="Visual roof layout",
        )

        actual_building_kwp = get_total_array_capacity_kwp(pv_arrays)
        actual_kwp_status = get_array_capacity_status(pv_arrays, part_l_required_kwp)
        actual_panel_status = format_pass_fail(
            float(get_total_array_panel_count(pv_arrays)),
            float(part_l_required_panel_count),
        )

        editor_kwp_status = actual_kwp_status
        editor_panel_status = actual_panel_status

    else:
        st.markdown("**Manual array input**")
        st.caption(
            "Use this route where the PV array capacity, orientation and tilt are already known. "
            "This bypasses the roof geometry editor and passes the manually entered arrays into the same downstream calculation object."
        )

        manual_array_count = st.number_input(
            "Number of PV arrays",
            min_value=1,
            max_value=8,
            value=2,
            step=1,
            key="manual_array_count",
        )

        manual_array_inputs = []

        for idx in range(1, int(manual_array_count) + 1):
            with st.container(border=True):
                st.markdown(f"**Array {idx}**")

                manual_cols = st.columns([0.55, 1.25, 1.0, 1.0, 1.0, 1.0, 1.0])

                with manual_cols[0]:
                    enabled = st.checkbox(
                        "Use",
                        value=True if idx == 1 else False,
                        key=f"manual_array_enabled_{idx}",
                    )

                with manual_cols[1]:
                    name = st.text_input(
                        "Array name",
                        value=f"Array {idx}",
                        key=f"manual_array_name_{idx}",
                    )

                with manual_cols[2]:
                    capacity_kwp = st.number_input(
                        "Capacity (kWp)",
                        min_value=0.00,
                        max_value=100.00,
                        value=2.50 if idx == 1 else 0.00,
                        step=0.10,
                        key=f"manual_array_capacity_{idx}",
                    )

                with manual_cols[3]:
                    orientation_label = st.selectbox(
                        "Orientation",
                        list(SAP_ORIENTATION_OPTIONS.keys()),
                        index=list(SAP_ORIENTATION_OPTIONS.keys()).index("South"),
                        key=f"manual_array_orientation_{idx}",
                    )

                with manual_cols[4]:
                    tilt_deg = st.selectbox(
                        "Tilt (deg)",
                        SAP_TILT_OPTIONS_DEG,
                        index=SAP_TILT_OPTIONS_DEG.index(30),
                        key=f"manual_array_tilt_{idx}",
                    )

                with manual_cols[5]:
                    shading_label = st.selectbox(
                        "Shading",
                        list(SAP_SHADING_OPTIONS.keys()),
                        index=0,
                        key=f"manual_array_shading_{idx}",
                    )

                with manual_cols[6]:
                    panel_count = st.number_input(
                        "Panel count",
                        min_value=0,
                        max_value=500,
                        value=0,
                        step=1,
                        key=f"manual_array_panel_count_{idx}",
                    )

                manual_array_inputs.append(
                    {
                        "enabled": enabled,
                        "name": name,
                        "capacity_kwp": capacity_kwp,
                        "orientation_label": orientation_label,
                        "tilt_deg": tilt_deg,
                        "shading_label": shading_label,
                        "panel_count": panel_count if int(panel_count) > 0 else None,
                    }
                )

        pv_arrays = build_pv_arrays_from_manual_inputs(manual_array_inputs)

        installed_panel_count = get_total_array_panel_count(pv_arrays)
        actual_building_kwp = get_total_array_capacity_kwp(pv_arrays)
        actual_kwp_status = get_array_capacity_status(pv_arrays, part_l_required_kwp)
        actual_panel_status = format_pass_fail(float(installed_panel_count), float(part_l_required_panel_count))

        editor_metrics = {
            "total_panels": installed_panel_count,
            "fitted_panels": installed_panel_count,
            "blocked_panels": 0,
            "invalid_panels": 0,
            "fitted_kwp": actual_building_kwp,
            "fitted_panels_by_plane": {},
        }

        editor_kwp_status = actual_kwp_status
        editor_panel_status = actual_panel_status

    pv_arrays = get_enabled_pv_arrays(pv_arrays)
    part_l_capacity_progress_pct = get_part_l_capacity_progress_pct(
        pv_arrays=pv_arrays,
        required_kwp=part_l_required_kwp,
    )

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    editor_summary_cols = st.columns(4)

    with editor_summary_cols[0]:
        render_summary_card(
            "Array capacity",
            f"{get_total_array_capacity_kwp(pv_arrays):,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>",
        )

    with editor_summary_cols[1]:
        render_summary_card(
            "Part L target achieved",
            f"{part_l_capacity_progress_pct:,.0f}<span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>%</span>",
        )

    with editor_summary_cols[2]:
        render_summary_card(
            "Panels fitted / declared",
            f"{get_total_array_panel_count(pv_arrays)}",
        )

    with editor_summary_cols[3]:
        render_summary_card(
            "Panels blocked / invalid",
            f"{editor_metrics['blocked_panels'] + editor_metrics['invalid_panels']}",
        )

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    capacity_progress_fig = build_part_l_capacity_progress_chart(
        required_kwp=part_l_required_kwp,
        actual_kwp=get_total_array_capacity_kwp(pv_arrays),
    )

    st.plotly_chart(
        capacity_progress_fig,
        theme=None,
        width="stretch",
        key="part_l_capacity_progress_chart",
    )

    st.markdown("**Array summary passed to generation calculation**")
    array_summary_rows = build_array_summary_rows(pv_arrays)
    if not array_summary_rows:
        array_summary_rows = build_empty_pv_array_summary_rows()
    st.dataframe(pd.DataFrame(array_summary_rows), hide_index=True, width="stretch")

# -----------------------------------------------------------------------------
# Photovoltaic array energy generation
# -----------------------------------------------------------------------------
render_section_title("generation_estimate", "Photovoltaic array energy generation")
with st.container(border=True):
    st.caption(
        "This section estimates annual electricity generation in kWh from the PV arrays passed from either "
        "the visual roof layout or the manual array input route. The user can choose PySAM PVWatts with an EPW "
        "weather file, or the SAP Appendix U lookup method."
    )

    generation_method = st.radio(
        "Generation calculation method",
        ["PySAM PVWatts", "SAP Appendix U"],
        index=0,
        horizontal=True,
        key="generation_method",
    )

    enabled_generation_arrays = get_enabled_pv_arrays(pv_arrays)
    total_generation_capacity_kwp = get_total_array_capacity_kwp(enabled_generation_arrays)
    total_generation_panel_count = get_total_array_panel_count(enabled_generation_arrays)

    pysam_result = None
    pysam_message = None
    selected_epw = None

    sap_appendix_u_result = None
    sap_appendix_u_message = None
    sap_appendix_u_inverter_efficiency = 0.95

    generation_result = None
    generation_message = None
    generation_result_annual_kwh = 0.0

    if generation_method == "PySAM PVWatts":
        epw_lookup = get_available_epw_files(EPW_DIRECTORY)
        epw_labels = ["None"] + list(epw_lookup.keys())

        pysam_input_cols = st.columns(2)
        with pysam_input_cols[0]:
            epw_label = st.selectbox(
                "Weather file for energy generation estimate (EPW)",
                epw_labels,
                index=0,
                key="pysam_epw_label",
            )
        with pysam_input_cols[1]:
            st.text_input(
                "Array definitions source",
                value="Inherited from photovoltaic array layout",
                disabled=True,
                key="pysam_array_source_display",
            )

        if total_generation_capacity_kwp <= 0:
            pysam_message = "No PV array capacity has been defined in the photovoltaic array layout section."
        elif epw_label == "None":
            pysam_message = "No EPW selected."
        elif pvwatts is None:
            pysam_message = get_pysam_missing_message()
        else:
            selected_epw = epw_lookup[epw_label]
            monthly_total = [0.0] * 12
            annual_total = 0.0

            try:
                pysam_array_rows = []

                for array in enabled_generation_arrays:
                    if array.capacity_kwp <= 0:
                        continue

                    arr_result = run_pysam_pvwatts(
                        system_capacity_kw=array.capacity_kwp,
                        weather_file=selected_epw,
                        tilt_deg=array.tilt_deg,
                        azimuth_deg=array.azimuth_deg,
                    )

                    shading_factor = max(min(float(array.shading_factor), 1.0), 0.0)
                    shaded_annual_ac_kwh = arr_result["annual_ac_kwh"] * shading_factor
                    shaded_monthly_ac_kwh = [
                        month_value * shading_factor
                        for month_value in arr_result["monthly_ac_kwh"]
                    ]

                    annual_total += shaded_annual_ac_kwh
                    monthly_total = [
                        existing + new
                        for existing, new in zip(monthly_total, shaded_monthly_ac_kwh)
                    ]

                    pysam_array_rows.append(
                        {
                            "Array": array.name,
                            "Input source": array.source,
                            "Azimuth (deg)": f"{array.azimuth_deg:.0f}",
                            "Orientation": get_orientation_label_from_azimuth(array.azimuth_deg),
                            "Tilt / panel pitch (deg)": f"{array.tilt_deg:.0f}",
                            "Panel count": "" if array.panel_count is None else f"{array.panel_count}",
                            "System capacity (kWp)": f"{array.capacity_kwp:.2f}",
                            "Shading factor": f"{shading_factor:.2f}",
                            "Annual AC generation before shading (kWh/a)": f"{arr_result['annual_ac_kwh']:.0f}",
                            "Annual AC generation after shading (kWh/a)": f"{shaded_annual_ac_kwh:.0f}",
                        }
                    )

                pysam_result = {
                    "method": "PySAM PVWatts v8",
                    "annual_ac_kwh": annual_total,
                    "monthly_ac_kwh": monthly_total,
                    "array_rows": pysam_array_rows,
                }

                generation_result = pysam_result
                generation_result_annual_kwh = annual_total

            except Exception as exc:
                pysam_message = f"PySAM run failed: {exc}"

        if pysam_result is not None:
            annual_gen_df = pd.DataFrame(
                [
                    ("Generation methodology", "PySAM PVWatts v8"),
                    ("Radiation / weather source", selected_epw.name if selected_epw else ""),
                    ("Selected EPW", epw_label),
                    ("Array source", array_input_mode),
                    ("Total declared / fitted panel count", f"{total_generation_panel_count}"),
                    ("Total system capacity used", f"{total_generation_capacity_kwp:,.2f} kWp"),
                    ("Annual AC generation", f"{pysam_result['annual_ac_kwh']:,.0f} kWh/a"),
                ],
                columns=["Metric", "Value"],
            )
            st.dataframe(annual_gen_df, hide_index=True, width="stretch")

            st.dataframe(
                pd.DataFrame(pysam_result["array_rows"]),
                hide_index=True,
                width="stretch",
            )

            pysam_assumptions_df = pd.DataFrame(
                [
                    ("Performance / system losses", f"{PYSAM_SYSTEM_LOSSES_PCT:.1f} %"),
                    ("DC/AC ratio", f"{PYSAM_DC_AC_RATIO:.2f}"),
                    ("Array type", "Fixed roof mount"),
                    ("Module type", "Standard"),
                    ("Ground coverage ratio", f"{PYSAM_GCR:.2f}"),
                    ("Shading treatment", "Array-level shading factor applied to PySAM AC output"),
                ],
                columns=["Assumption", "Value"],
            )
            st.dataframe(pysam_assumptions_df, hide_index=True, width="stretch")

            monthly_df = pd.DataFrame(
                {
                    "Month": SAP_APPENDIX_U_MONTHS,
                    "AC generation (kWh)": [round(v, 1) for v in pysam_result["monthly_ac_kwh"]],
                }
            )
            st.dataframe(monthly_df, hide_index=True, width="stretch")
        else:
            generation_message = pysam_message or "Annual generation not available."
            st.info(generation_message)

    else:
        sap_cols = st.columns(3)

        with sap_cols[0]:
            sap_appendix_u_region_options = list(SAP_APPENDIX_U_REGION_DATA.keys())

            sap_appendix_u_region = st.selectbox(
                "SAP Appendix U region",
                sap_appendix_u_region_options,
                index=sap_appendix_u_region_options.index("England - London / South East"),
                key="sap_appendix_u_region",
            )

        with sap_cols[1]:
            sap_appendix_u_inverter_efficiency = st.number_input(
                "Inverter efficiency",
                min_value=0.50,
                max_value=1.00,
                value=0.95,
                step=0.01,
                key="sap_appendix_u_inverter_efficiency",
            )

        with sap_cols[2]:
            st.text_input(
                "Array definitions source",
                value="Inherited from photovoltaic array layout",
                disabled=True,
                key="sap_appendix_u_array_source_display",
            )

        if total_generation_capacity_kwp <= 0:
            sap_appendix_u_message = "No PV array capacity has been defined in the photovoltaic array layout section."
        elif not has_appendix_u_table_data(sap_appendix_u_region):
            sap_appendix_u_message = (
                f"No Appendix U region data has been entered for '{sap_appendix_u_region}'. "
                "Add the region to SAP_APPENDIX_U_REGION_DATA before using this method."
            )
        else:
            try:
                sap_appendix_u_result = calculate_sap_appendix_u_generation(
                    pv_arrays=enabled_generation_arrays,
                    region=sap_appendix_u_region,
                    inverter_efficiency=float(sap_appendix_u_inverter_efficiency),
                )

                generation_result = sap_appendix_u_result
                generation_result_annual_kwh = sap_appendix_u_result["annual_generation_kwh"]

            except Exception as exc:
                sap_appendix_u_message = f"SAP Appendix U calculation failed: {exc}"

        if sap_appendix_u_result is not None:
            sap_summary_df = pd.DataFrame(
                [
                    ("Generation methodology", "SAP Appendix U"),
                    ("SAP Appendix U region", sap_appendix_u_result["region"]),
                    ("Array source", array_input_mode),
                    ("Total declared / fitted panel count", f"{total_generation_panel_count}"),
                    ("Total system capacity used", f"{total_generation_capacity_kwp:,.2f} kWp"),
                    ("Inverter efficiency", f"{sap_appendix_u_inverter_efficiency:.2f}"),
                    ("Annual generation", f"{sap_appendix_u_result['annual_generation_kwh']:,.0f} kWh/a"),
                ],
                columns=["Metric", "Value"],
            )
            st.dataframe(sap_summary_df, hide_index=True, width="stretch")

            st.dataframe(
                pd.DataFrame(build_sap_appendix_u_array_rows(sap_appendix_u_result["array_results"])),
                hide_index=True,
                width="stretch",
            )

            monthly_df = pd.DataFrame(
                {
                    "Month": SAP_APPENDIX_U_MONTHS,
                    "Generation (kWh)": [
                        round(v, 1)
                        for v in sap_appendix_u_result["monthly_generation_kwh"]
                    ],
                }
            )
            st.dataframe(monthly_df, hide_index=True, width="stretch")
        else:
            generation_message = sap_appendix_u_message or "Annual generation not available."
            st.info(generation_message)

    if generation_result is None:
        generation_result_annual_kwh = 0.0
st.divider()

# -----------------------------------------------------------------------------
# Bottom summary tables
# -----------------------------------------------------------------------------
total_gross_roof_area_m2 = sum(plane.gross_area_m2 for plane in roof_planes)
usable_available_pv_area_m2 = sum(plane.usable_area_m2 for plane in roof_planes)
other_reduction_area_m2 = max(total_gross_roof_area_m2 - usable_available_pv_area_m2, 0.0)

display_panel_counts_for_planes = get_plane_displayed_panel_counts(
    roof_form=actual_roof_form,
    installed_panel_count=installed_panel_count,
    actual_array_panel_counts=actual_array_panel_counts,
)

selected_epw_label = "Not applicable"
selected_epw_file = "Not applicable"
sap_appendix_u_region_summary = "Not applicable"
sap_appendix_u_inverter_efficiency_summary = "Not applicable"

if generation_method == "PySAM PVWatts":
    selected_epw_label = locals().get("epw_label", "None")
    selected_epw_file = selected_epw.name if selected_epw is not None else "Not selected"
elif generation_method == "SAP Appendix U":
    sap_appendix_u_region_summary = locals().get("sap_appendix_u_region", "Not selected")
    sap_appendix_u_inverter_efficiency_summary = f"{float(sap_appendix_u_inverter_efficiency):.2f}"

if generation_result is not None:
    generation_status_text = "Calculated"
    generation_annual_kwh_text = f"{generation_result_annual_kwh:,.0f} kWh/a"
else:
    generation_status_text = generation_message or "Annual generation not available."
    generation_annual_kwh_text = "Not calculated"

user_inputs_rows = [
    ("Dwelling inputs", "House form", house_form),
    ("Dwelling inputs", "Ground floor area method", gfa_input_mode),
    ("Dwelling inputs", "Ground floor area", f"{ground_floor_area_m2:,.2f} m²"),
    ("Dwelling inputs", "Ground floor area source", gfa_source_text),
    ("Photovoltaic array layout", "Array input method", array_input_mode),
    ("Photovoltaic array layout", "Roof type", actual_roof_form),
    ("Photovoltaic array layout", "Length along ridge / whole roof length in plan", f"{plan_length_along_ridge_m:,.2f} m"),
    ("Photovoltaic array layout", "Ridge-to-eaves / whole roof width in plan", f"{plan_length_ridge_to_eaves_m:,.2f} m"),
    ("Photovoltaic array layout", "Usable roof area method", offset_mode_section_2),
]

if actual_roof_form == "Flat":
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Flat panel pitch above horizontal",
            f"{flat_panel_pitch_deg:.0f}°",
        )
    )
elif actual_roof_form in {"Mono-pitch", "Duo-pitch"}:
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Roof plane azimuth",
            f"{mono_or_duo_azimuth_deg:.0f}°",
        )
    )
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Roof pitch",
            f"{mono_or_duo_pitch_deg:.0f}°",
        )
    )
else:
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Roof geometry",
            "Not applicable for manual array input",
        )
    )

if offset_mode_section_2 == "Simple setback":
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Setback around usable array area",
            f"{simple_setback_m:.2f} m",
        )
    )
elif offset_mode_section_2 == "Detailed offsets":
    user_inputs_rows.append(("Photovoltaic array layout", "Ridge offset", f"{ridge_offset_m:.2f} m"))
    user_inputs_rows.append(("Photovoltaic array layout", "Roof edge offset", f"{edge_offset_m:.2f} m"))
    user_inputs_rows.append(("Photovoltaic array layout", "Party wall offset", f"{party_wall_offset_m:.2f} m"))
else:
    user_inputs_rows.append(
        (
            "Photovoltaic array layout",
            "Roof reduction inputs",
            "Not applicable for manual array input",
        )
    )

user_inputs_rows.extend(
    [
        ("Photovoltaic array layout", "Module width", f"{module_width_m * 1000:.0f} mm"),
        ("Photovoltaic array layout", "Module length", f"{module_length_m * 1000:.0f} mm"),
        ("Photovoltaic array layout", "Module efficiency", f"{module_efficiency_pct:,.1f} %"),
        ("Photovoltaic array layout", "Mount orientation", module_mount_orientation),
        ("Photovoltaic array layout", "Panels fitted / declared", f"{get_total_array_panel_count(pv_arrays)}"),
        ("Photovoltaic array layout", "Array capacity", f"{get_total_array_capacity_kwp(pv_arrays):,.2f} kWp"),
        ("Photovoltaic array energy generation", "Generation calculation method", generation_method),
        ("Photovoltaic array energy generation", "Selected EPW", selected_epw_label),
        ("Photovoltaic array energy generation", "Selected EPW file", selected_epw_file),
        ("Photovoltaic array energy generation", "SAP Appendix U region", sap_appendix_u_region_summary),
        ("Photovoltaic array energy generation", "SAP Appendix U inverter efficiency", sap_appendix_u_inverter_efficiency_summary),
    ]
)

calculation_assumption_rows = [
    ("Part L photovoltaic target", "Required PV area fraction", f"{FHS_REQUIRED_AREA_FRACTION:.2f} of ground floor area"),
    ("Part L photovoltaic target", "Standard panel efficiency density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:.2f} kWp/m²"),
    ("Part L photovoltaic target", "Target capacity formula", "Ground floor area × required PV area fraction × standard panel efficiency density"),
    ("Part L photovoltaic target", "Target photovoltaic capacity", f"{part_l_required_kwp:,.2f} kWp"),
    ("Part L photovoltaic target", "Equivalent standardised panel count", f"{part_l_required_panel_count}"),
    ("Part L photovoltaic target", "Standardised module width", f"{STANDARDISED_MODULE_WIDTH_M * 1000:.0f} mm"),
    ("Part L photovoltaic target", "Standardised module length", f"{STANDARDISED_MODULE_LENGTH_M * 1000:.0f} mm"),
    ("Part L photovoltaic target", "Standardised module efficiency", f"{STANDARDISED_MODULE_EFFICIENCY_PCT:.1f} %"),
    ("Photovoltaic array layout", "Roof planes used", f"{len(roof_planes)}"),
    ("Photovoltaic array layout", "Total gross roof area", f"{total_gross_roof_area_m2:,.2f} m²"),
    ("Photovoltaic array layout", "Other reduction area from margins / offsets", f"{other_reduction_area_m2:,.2f} m²"),
    ("Photovoltaic array layout", "Usable roof area for PV", f"{usable_available_pv_area_m2:,.2f} m²"),
    ("Photovoltaic array layout", "Derived module power", f"{module_power_wp:,.0f} Wp"),
    ("Photovoltaic array layout", "Maximum feasible panel count", f"{max_feasible_panels}"),
    ("Photovoltaic array layout", "Actual array capacity", f"{get_total_array_capacity_kwp(pv_arrays):,.2f} kWp"),
    ("Photovoltaic array layout", "kWp check against Part L requirement", actual_kwp_status),
    ("Photovoltaic array layout", "Panel count check against equivalent standardised panel count", actual_panel_status),
    ("Photovoltaic array layout", "Editor fitted kWp", f"{editor_metrics['fitted_kwp']:,.2f} kWp"),
    ("Photovoltaic array layout", "Editor fitted panels", f"{editor_metrics['fitted_panels']}"),
    ("Photovoltaic array layout", "Editor blocked panels", f"{editor_metrics['blocked_panels']}"),
    ("Photovoltaic array layout", "Editor invalid panels", f"{editor_metrics['invalid_panels']}"),
    ("Photovoltaic array layout", "Editor kWp check against Part L requirement", editor_kwp_status),
    ("Photovoltaic array layout", "Editor panel count check against equivalent standardised panel count", editor_panel_status),
    ("Photovoltaic array layout", "Editor movement step default", f"{OBSTACLE_NUDGE_STEP_DEFAULT:.2f} m"),
    ("Photovoltaic array energy generation", "Generation method selected", generation_method),
    ("Photovoltaic array energy generation", "Generation status", generation_status_text),
    ("Photovoltaic array energy generation", "Annual generation result", generation_annual_kwh_text),

]

if generation_method == "PySAM PVWatts":
    calculation_assumption_rows.extend(
        [
            ("Photovoltaic array energy generation", "PySAM availability", "Installed" if pvwatts is not None else "Not installed"),
            ("Photovoltaic array energy generation", "Weather source", selected_epw_file),
            ("Photovoltaic array energy generation", "Performance / system losses", f"{PYSAM_SYSTEM_LOSSES_PCT:.1f} %"),
            ("Photovoltaic array energy generation", "DC/AC ratio", f"{PYSAM_DC_AC_RATIO:.2f}"),
            ("Photovoltaic array energy generation", "Array type", "Fixed roof mount"),
            ("Photovoltaic array energy generation", "Module type", "Standard"),
            ("Photovoltaic array energy generation", "Ground coverage ratio", f"{PYSAM_GCR:.2f}"),
            ("Photovoltaic array energy generation", "Shading treatment", "Array-level shading factor applied to PySAM AC output"),
        ]
    )
elif generation_method == "SAP Appendix U":
    if sap_appendix_u_region_summary in SAP_APPENDIX_U_REGION_DATA:
        appendix_u_region_data = SAP_APPENDIX_U_REGION_DATA[sap_appendix_u_region_summary]
        appendix_u_region_description = (
            f"{appendix_u_region_data['sap_region']} - "
            f"{appendix_u_region_data['sap_region_name']}, "
            f"{appendix_u_region_data['latitude_deg']:.1f}°N"
        )
    else:
        appendix_u_region_description = "Not selected"

    calculation_assumption_rows.extend(
        [
            ("Photovoltaic array energy generation", "SAP Appendix U region mapping", appendix_u_region_description),
            ("Photovoltaic array energy generation", "SAP Appendix U horizontal irradiance source", "Monthly mean horizontal irradiance values coded in SAP_APPENDIX_U_REGION_DATA"),
            ("Photovoltaic array energy generation", "SAP Appendix U declination source", "Monthly solar declination values coded in SAP_APPENDIX_U_SOLAR_DECLINATION_DEG"),
            ("Photovoltaic array energy generation", "SAP Appendix U orientation constants source", "Orientation constants coded in SAP_APPENDIX_U_TABLE_U5_CONSTANTS"),
            ("Photovoltaic array energy generation", "Inverter efficiency", sap_appendix_u_inverter_efficiency_summary),
            ("Photovoltaic array energy generation", "Shading treatment", "Array-level shading factor applied to SAP Appendix U generation"),
        ]
    )

pv_array_summary_rows = build_array_summary_rows(pv_arrays)
if not pv_array_summary_rows:
    pv_array_summary_rows = build_empty_pv_array_summary_rows()

render_section_title("inputs_summary", "Inputs added by user")
with st.container(border=True):
    user_inputs_df = pd.DataFrame(user_inputs_rows, columns=["Section", "Input", "Value"])
    st.dataframe(user_inputs_df, hide_index=True, width="stretch")

render_section_title("calculation_assumptions", "Calculation assumptions")
with st.container(border=True):
    calculation_assumptions_df = pd.DataFrame(
        calculation_assumption_rows,
        columns=["Section", "Assumption", "Value"],
    )
    st.dataframe(calculation_assumptions_df, hide_index=True, width="stretch")

    st.markdown("**PV arrays used by generation calculation**")
    st.dataframe(pd.DataFrame(pv_array_summary_rows), hide_index=True, width="stretch")

render_section_title("roof_plane_table", "Roof plane table")
with st.container(border=True):
    if roof_planes:
        plane_rows = build_plane_table_rows(
            roof_planes=roof_planes,
            display_panel_counts_for_planes=display_panel_counts_for_planes,
            plane_layouts=plane_layouts,
        )
        st.dataframe(pd.DataFrame(plane_rows), hide_index=True, width="stretch")
    else:
        st.info("No roof-plane geometry is used when the manual array input route is selected.")

with st.expander("Roof editor state JSON", expanded=False):
    st.caption("Serialized roof geometry and layout payload for the interactive editor.")
    st.json(roof_editor_state, expanded=False)
    st.code(json.dumps(roof_editor_state, indent=2), language="json")
