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
# Roof reduction assumptions
# -----------------------------------------------------------------------------
DEFAULT_SIMPLE_PERIMETER_MARGIN_M = 0.30
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

CHART_HEIGHT = 420
CHART_BACKGROUND_COLOUR = "white"
CHART_PLOT_BACKGROUND_COLOUR = "white"
CHART_FONT_COLOUR = "#333333"
CHART_GRID_COLOUR = "#E6E6E6"
CHART_AXIS_LINE_COLOUR = "#BFBFBF"
CHART_MARGIN = dict(l=60, r=20, t=24, b=20)
CHART_BAR_COLOURS = ["#4F67FF", "#F05A3A"]
CHART_BAR_WIDTH = 0.19
CHART_BARGAP = 0.45

# -----------------------------------------------------------------------------
# Section theme colours from logo
# -----------------------------------------------------------------------------
ETUDE_PURPLE = "#6A4BA3"
ETUDE_ORANGE = "#D97C3F"
ETUDE_SAGE = "#B8C95E"
ETUDE_BLUE = "#79AFCB"
ETUDE_PINK = "#D14B8F"

SECTION_THEME_COLOURS = {
    "part_l_target": ETUDE_PURPLE,
    "building_pv_layout": ETUDE_ORANGE,
    "pysam_forecast": ETUDE_SAGE,
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
class ArrayDefinition:
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


def round_up_to_nice_step(value: float) -> float:
    if value <= 0:
        return 1000.0
    if value <= 5000:
        step = 500.0
    elif value <= 20000:
        step = 1000.0
    elif value <= 100000:
        step = 5000.0
    else:
        step = 10000.0
    return math.ceil(value / step) * step


def get_comparison_chart_axis_max(
    ground_floor_area_m2: float,
    actual_roof_form: str,
    plan_length_along_ridge_m: float,
    plan_length_ridge_to_eaves_m: float,
    part_l_required_generation_kwh: float,
    max_feasible_panels: int,
    module_power_kwp: float,
    sap_placeholder_specific_yield: float,
) -> float:
    signature = (
        round(ground_floor_area_m2, 2),
        actual_roof_form,
        round(plan_length_along_ridge_m, 2),
        round(plan_length_ridge_to_eaves_m, 2),
    )

    chart_capacity_basis = max_feasible_panels * module_power_kwp * sap_placeholder_specific_yield
    suggested_max = round_up_to_nice_step(max(part_l_required_generation_kwh, chart_capacity_basis) * 1.10)

    if "comparison_chart_signature" not in st.session_state:
        st.session_state.comparison_chart_signature = signature
        st.session_state.comparison_chart_ymax = suggested_max
    elif st.session_state.comparison_chart_signature != signature:
        st.session_state.comparison_chart_signature = signature
        st.session_state.comparison_chart_ymax = suggested_max

    return float(st.session_state.comparison_chart_ymax)


def build_comparison_chart(
    required_generation_kwh: float,
    actual_generation_kwh: float,
    y_axis_max: float,
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
        yaxis_title_standoff=20,
        showlegend=False,
        bargap=CHART_BARGAP,
        paper_bgcolor=CHART_BACKGROUND_COLOUR,
        plot_bgcolor=CHART_PLOT_BACKGROUND_COLOUR,
        font=dict(color=CHART_FONT_COLOUR),
    )
    fig.update_xaxes(showgrid=False, linecolor=CHART_AXIS_LINE_COLOUR)
    fig.update_yaxes(
        gridcolor=CHART_GRID_COLOUR,
        linecolor=CHART_AXIS_LINE_COLOUR,
        range=[0, y_axis_max],
        automargin=True,
    )

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


def subtract_rect_from_rect(source: Rect, cutter: Rect, min_dim: float = MIN_ZONE_DIM_M) -> list[Rect]:
    overlap = rect_intersection(source, cutter)
    if overlap is None:
        return [source]

    pieces = []

    # Top
    if overlap.y > source.y:
        pieces.append(
            Rect(
                x=source.x,
                y=source.y,
                w=source.w,
                h=overlap.y - source.y,
            )
        )

    # Bottom
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

    # Left
    if overlap.x > source.x:
        pieces.append(
            Rect(
                x=source.x,
                y=overlap.y,
                w=overlap.x - source.x,
                h=overlap.h,
            )
        )

    # Right
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

    clean = [
        p for p in pieces
        if p.w >= min_dim and p.h >= min_dim
    ]
    return clean

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
                    "label": f"PV rectangle {idx}",
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
    count = rows * cols

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
    perimeter_margin_m: float,
    ridge_offset_m: float,
    edge_offset_m: float,
    party_wall_offset_m: float,
    house_form: str,
) -> RoofGeometryBundle:
    planes: list[RoofPlaneGeometry] = []

    if roof_form == "Flat":
        gross_length_m = plan_length_along_ridge_m
        gross_depth_m = plan_length_ridge_to_eaves_m

        if perimeter_margin_m > 0:
            margin_left = perimeter_margin_m
            margin_right = perimeter_margin_m
            margin_top = perimeter_margin_m
            margin_bottom = perimeter_margin_m
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

    if perimeter_margin_m > 0:
        margin_left = perimeter_margin_m
        margin_right = perimeter_margin_m
        margin_top = perimeter_margin_m
        margin_bottom = perimeter_margin_m
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

            px = start_x + col * layout.panel_w_m
            py = start_y + row * layout.panel_h_m

            panels.append(
                EditorPanelInstance(
                    panel_id=f"{plane.plane_id}_panel_{drawn + 1}",
                    plane_id=plane.plane_id,
                    x=px,
                    y=py,
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
            label="PV rectangle 1",
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
        pv_zones=[asdict(z) for z in pv_zone_instances],
        panels=[asdict(p) for p in panel_instances],
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
        label = str(record.get("label", "")).strip() or f"PV rectangle {idx}"
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
    sap_placeholder_specific_yield: float,
) -> dict:
    roof_form = editor_state["roof_form"]

    if roof_form == "Flat":
        average_factor = (
            sap_orientation_factor_placeholder(90.0)
            * sap_tilt_factor_placeholder(editor_state["module"]["flat_panel_pitch_deg"])
            + sap_orientation_factor_placeholder(270.0)
            * sap_tilt_factor_placeholder(editor_state["module"]["flat_panel_pitch_deg"])
        ) / 2.0
        plane_factors = {"plane_1": average_factor}
    else:
        plane_factors = {}
        for plane in editor_state["planes"]:
            plane_factors[plane["plane_id"]] = (
                sap_orientation_factor_placeholder(float(plane["azimuth_deg"]))
                * sap_tilt_factor_placeholder(float(plane["tilt_deg"]))
            )

    valid_panels = 0
    blocked_panels = 0
    invalid_panels = 0
    total_panels = 0
    valid_generation_kwh = 0.0
    valid_panels_by_plane = {}

    for plane in editor_state["planes"]:
        plane_id = plane["plane_id"]
        valid_count = 0

        for panel in plane.get("panels", []):
            total_panels += 1
            status = panel.get("status", "auto")

            if status in {"auto", "user"}:
                valid_panels += 1
                valid_count += 1
            elif status == "blocked":
                blocked_panels += 1
            else:
                invalid_panels += 1

        valid_panels_by_plane[plane_id] = valid_count
        valid_generation_kwh += (
            valid_count * module_power_kwp * sap_placeholder_specific_yield * plane_factors.get(plane_id, 1.0)
        )

    return {
        "total_panels": total_panels,
        "valid_panels": valid_panels,
        "blocked_panels": blocked_panels,
        "invalid_panels": invalid_panels,
        "valid_kwp": valid_panels * module_power_kwp,
        "valid_generation_kwh": valid_generation_kwh,
        "valid_panels_by_plane": valid_panels_by_plane,
    }


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
        "label": f"PV rectangle {len(plane.get('pv_zones', [])) + 1}",
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

        move_step = st.number_input(
            "Movement step (m)",
            min_value=0.01,
            max_value=1.0,
            value=OBSTACLE_NUDGE_STEP_DEFAULT,
            step=0.01,
            key=f"editor_move_step_{selected_plane_id}",
        )

    st.markdown("**Step 1: Obstacles (optional)**")

    obstacle_add_cols = st.columns(4)
    with obstacle_add_cols[0]:
        new_obstacle_type = st.selectbox(
            "Obstacle type",
            ["generic", "window", "vent", "plant"],
            key=f"new_obstacle_type_{selected_plane_id}",
        )
    with obstacle_add_cols[1]:
        new_obstacle_width = st.number_input(
            "Obstacle width (m)",
            min_value=0.01,
            max_value=float(dict_to_rect(selected_plane["usable_rect"]).w),
            value=1.00,
            step=0.05,
            key=f"new_obstacle_width_{selected_plane_id}",
        )
    with obstacle_add_cols[2]:
        new_obstacle_height = st.number_input(
            "Obstacle height (m)",
            min_value=0.01,
            max_value=float(dict_to_rect(selected_plane["usable_rect"]).h),
            value=1.00,
            step=0.05,
            key=f"new_obstacle_height_{selected_plane_id}",
        )
    with obstacle_add_cols[3]:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        if st.button("Add obstacle", key=f"add_visual_obstacle_button_{selected_plane_id}", width="stretch"):
            editor_state = add_visual_obstacle_to_plane(
                editor_state=editor_state,
                plane_id=selected_plane_id,
                obstacle_type=new_obstacle_type,
                width_m=float(new_obstacle_width),
                height_m=float(new_obstacle_height),
            )
            persist_editor_source_state(editor_state)

    user_obstacles = selected_plane.get("obstacles", [])

    if user_obstacles:
        obstacle_lookup = {obs["obstacle_id"]: obs for obs in user_obstacles}

        obstacle_cols = st.columns([1.8, 1.0, 1.0, 1.0, 1.0, 1.0])

        with obstacle_cols[0]:
            selected_obstacle_id = st.selectbox(
                "Obstacle to edit",
                options=list(obstacle_lookup.keys()),
                format_func=lambda oid: f"{oid} ({obstacle_lookup[oid]['obstacle_type']})",
                key=f"selected_obstacle_{selected_plane_id}",
            )

        obstacle = obstacle_lookup[selected_obstacle_id]

        with obstacle_cols[1]:
            obs_type = st.selectbox(
                "Type",
                ["generic", "window", "vent", "plant"],
                index=["generic", "window", "vent", "plant"].index(obstacle["obstacle_type"]),
                key=f"obs_type_edit_{selected_obstacle_id}",
            )
        with obstacle_cols[2]:
            obs_w = st.number_input(
                "Width (m)",
                min_value=0.01,
                value=float(obstacle["w"]),
                step=0.05,
                key=f"obs_width_edit_{selected_obstacle_id}",
            )
        with obstacle_cols[3]:
            obs_h = st.number_input(
                "Height (m)",
                min_value=0.01,
                value=float(obstacle["h"]),
                step=0.05,
                key=f"obs_height_edit_{selected_obstacle_id}",
            )
        with obstacle_cols[4]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("Apply changes", key=f"apply_obstacle_changes_{selected_obstacle_id}", width="stretch"):
                editor_state = update_obstacle_size_type(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                    float(obs_w),
                    float(obs_h),
                    obs_type,
                )
                persist_editor_source_state(editor_state)
        with obstacle_cols[5]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("Delete selected", key=f"delete_obstacle_{selected_obstacle_id}", width="stretch"):
                editor_state = delete_obstacle_from_plane(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                )
                persist_editor_source_state(editor_state)

        move_cols = st.columns(4)
        with move_cols[0]:
            if st.button("←", key=f"move_left_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                    dx=-float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)
        with move_cols[1]:
            if st.button("→", key=f"move_right_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                    dx=float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)
        with move_cols[2]:
            if st.button("↑", key=f"move_up_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                    dx=0.0,
                    dy=-float(move_step),
                )
                persist_editor_source_state(editor_state)
        with move_cols[3]:
            if st.button("↓", key=f"move_down_{selected_obstacle_id}", width="stretch"):
                editor_state = move_obstacle_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_obstacle_id,
                    dx=0.0,
                    dy=float(move_step),
                )
                persist_editor_source_state(editor_state)
    else:
        st.info("No user-added obstacles defined.")

    st.markdown("---")
    st.markdown("**Step 2: PV rectangles**")

    zone_add_cols = st.columns(4)
    with zone_add_cols[0]:
        new_zone_width = st.number_input(
            "PV rectangle width (m)",
            min_value=0.01,
            max_value=float(dict_to_rect(selected_plane["packing_rect"]).w),
            value=2.50,
            step=0.05,
            key=f"new_zone_width_{selected_plane_id}",
        )
    with zone_add_cols[1]:
        new_zone_height = st.number_input(
            "PV rectangle height (m)",
            min_value=0.01,
            max_value=float(dict_to_rect(selected_plane["packing_rect"]).h),
            value=2.50,
            step=0.05,
            key=f"new_zone_height_{selected_plane_id}",
        )
    with zone_add_cols[2]:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        if st.button("Add PV rectangle", key=f"add_zone_button_{selected_plane_id}", width="stretch"):
            editor_state = add_visual_zone_to_plane(
                editor_state=editor_state,
                plane_id=selected_plane_id,
                width_m=float(new_zone_width),
                height_m=float(new_zone_height),
            )
            persist_editor_source_state(editor_state)
    with zone_add_cols[3]:
        st.empty()

    zones = selected_plane.get("pv_zones", [])

    if zones:
        zone_lookup = {zone["zone_id"]: zone for zone in zones}

        zone_cols = st.columns([1.8, 1.2, 1.0, 1.0, 1.0, 1.0])

        with zone_cols[0]:
            selected_zone_id = st.selectbox(
                "PV rectangle to edit",
                options=list(zone_lookup.keys()),
                format_func=lambda zid: zone_lookup[zid].get("label", zid) or zid,
                key=f"selected_zone_{selected_plane_id}",
            )

        zone = zone_lookup[selected_zone_id]

        with zone_cols[1]:
            zone_label = st.text_input(
                "Label",
                value=zone.get("label", ""),
                key=f"zone_label_{selected_zone_id}",
            )
        with zone_cols[2]:
            zone_w = st.number_input(
                "Width (m)",
                min_value=0.01,
                value=float(zone["w"]),
                step=0.05,
                key=f"zone_w_{selected_zone_id}",
            )
        with zone_cols[3]:
            zone_h = st.number_input(
                "Height (m)",
                min_value=0.01,
                value=float(zone["h"]),
                step=0.05,
                key=f"zone_h_{selected_zone_id}",
            )
        with zone_cols[4]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("Apply changes", key=f"zone_apply_{selected_zone_id}", width="stretch"):
                editor_state = update_zone_properties(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                    float(zone_w),
                    float(zone_h),
                    zone_label,
                )
                persist_editor_source_state(editor_state)
        with zone_cols[5]:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("Delete selected", key=f"zone_delete_{selected_zone_id}", width="stretch"):
                editor_state = delete_zone_from_plane(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                )
                persist_editor_source_state(editor_state)

        zone_move_cols = st.columns(4)
        with zone_move_cols[0]:
            if st.button("←", key=f"zone_left_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                    dx=-float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)
        with zone_move_cols[1]:
            if st.button("→", key=f"zone_right_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                    dx=float(move_step),
                    dy=0.0,
                )
                persist_editor_source_state(editor_state)
        with zone_move_cols[2]:
            if st.button("↑", key=f"zone_up_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                    dx=0.0,
                    dy=-float(move_step),
                )
                persist_editor_source_state(editor_state)
        with zone_move_cols[3]:
            if st.button("↓", key=f"zone_down_{selected_zone_id}", width="stretch"):
                editor_state = move_zone_in_plane(
                    editor_state,
                    selected_plane_id,
                    selected_zone_id,
                    dx=0.0,
                    dy=float(move_step),
                )
                persist_editor_source_state(editor_state)
    else:
        st.warning("No PV rectangles defined.")

    return editor_state


# -----------------------------------------------------------------------------
# Diagram helpers
# -----------------------------------------------------------------------------
def add_margin_annotations(
    fig: go.Figure,
    roof_x0: float,
    roof_y0: float,
    roof_x1: float,
    roof_y1: float,
    usable_x0: float,
    usable_y0: float,
    usable_x1: float,
    usable_y1: float,
    plane: RoofPlaneGeometry,
) -> None:
    left_margin = plane.margin_left_m
    right_margin = plane.margin_right_m
    top_margin = plane.margin_top_m
    bottom_margin = plane.margin_bottom_m

    if top_margin > 0:
        fig.add_annotation(
            x=(usable_x0 + usable_x1) / 2.0,
            y=(roof_y0 + usable_y0) / 2.0,
            showarrow=False,
            text=f"top margin {top_margin:.2f} m",
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=10),
        )

    if bottom_margin > 0:
        fig.add_annotation(
            x=(usable_x0 + usable_x1) / 2.0,
            y=(usable_y1 + roof_y1) / 2.0,
            showarrow=False,
            text=f"bottom margin {bottom_margin:.2f} m",
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=10),
        )

    if left_margin > 0:
        fig.add_annotation(
            x=(roof_x0 + usable_x0) / 2.0,
            y=(usable_y0 + usable_y1) / 2.0,
            showarrow=False,
            text=f"{left_margin:.2f} m",
            textangle=90,
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=10),
        )

    if right_margin > 0:
        fig.add_annotation(
            x=(usable_x1 + roof_x1) / 2.0,
            y=(usable_y0 + usable_y1) / 2.0,
            showarrow=False,
            text=f"{right_margin:.2f} m",
            textangle=90,
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=10),
        )


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
                f"PV rectangles {len(plane.get('pv_zones', []))} / max {plane['max_feasible_panels']} panels"
            )
        else:
            detail_text = (
                f"Azimuth {float(plane['azimuth_deg']):.0f}° roof tilt {float(plane['tilt_deg']):.0f}°<br>"
                f"Gross {gross_w:.2f} × {gross_h:.2f} m<br>"
                f"Usable {usable_w:.2f} × {usable_h:.2f} m<br>"
                f"PV rectangles {len(plane.get('pv_zones', []))} / max {plane['max_feasible_panels']} panels"
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


def build_plane_table_rows(
    roof_planes: list[RoofPlaneGeometry],
    display_panel_counts_for_planes: list[int],
    actual_roof_form: str,
    flat_panel_pitch_deg: float,
    plane_layouts: dict[str, PanelLayout],
) -> list[dict]:
    plane_rows = []

    for plane, installed_panels in zip(roof_planes, display_panel_counts_for_planes):
        if actual_roof_form == "Flat":
            arr_factor = (
                sap_orientation_factor_placeholder(90.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
                + sap_orientation_factor_placeholder(270.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
            ) / 2.0
        else:
            arr_factor = sap_orientation_factor_placeholder(plane.azimuth_deg) * sap_tilt_factor_placeholder(
                plane.tilt_deg
            )

        plane_rows.append(
            {
                "Plane": plane.name,
                "Azimuth (deg)": f"{plane.azimuth_deg:.0f}",
                "Roof tilt (deg)": f"{plane.tilt_deg:.0f}",
                "Gross length (m)": f"{plane.gross_length_m:.2f}",
                "Gross depth (m)": f"{plane.gross_depth_m:.2f}",
                "Left margin (m)": f"{plane.margin_left_m:.2f}",
                "Right margin (m)": f"{plane.margin_right_m:.2f}",
                "Top margin (m)": f"{plane.margin_top_m:.2f}",
                "Bottom margin (m)": f"{plane.margin_bottom_m:.2f}",
                "Usable length (m)": f"{plane.usable_length_m:.2f}",
                "Usable depth (m)": f"{plane.usable_depth_m:.2f}",
                "Max feasible panels": f"{plane_layouts[plane.plane_id].count}",
                "Displayed / associated panels": f"{installed_panels}",
                "SAP placeholder factor": f"{arr_factor:.3f}",
            }
        )

    return plane_rows


# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
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
# Part L target
# -----------------------------------------------------------------------------
render_section_title("part_l_target", "Part L target")
with st.container(border=True):
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

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Building PV layout
# -----------------------------------------------------------------------------
render_section_title("building_pv_layout", "Building PV layout")
with st.container(border=True):
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
            "Roof reduction method",
            ["Simple perimeter margin", "Detailed offsets"],
            index=0,
            key="actual_offset_mode",
        )

    if offset_mode_section_2 == "Simple perimeter margin":
        with reduction_row[1]:
            perimeter_margin_m = st.number_input(
                "Perimeter margin around PV zone (m)",
                min_value=0.0,
                max_value=2.0,
                value=DEFAULT_SIMPLE_PERIMETER_MARGIN_M,
                step=0.05,
                key="actual_perimeter_margin",
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

        perimeter_margin_m = 0.0

    roof_geometry = build_roof_geometry(
        roof_form=actual_roof_form,
        plan_length_along_ridge_m=plan_length_along_ridge_m,
        plan_length_ridge_to_eaves_m=plan_length_ridge_to_eaves_m,
        pitch_deg=mono_or_duo_pitch_deg,
        azimuth_deg=mono_or_duo_azimuth_deg,
        perimeter_margin_m=perimeter_margin_m,
        ridge_offset_m=ridge_offset_m,
        edge_offset_m=edge_offset_m,
        party_wall_offset_m=party_wall_offset_m,
        house_form=house_form,
    )
    roof_planes = roof_geometry.planes

    st.markdown("**PV panel parameters**")

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

    plane_layouts: dict[str, PanelLayout] = {}
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

    comparison_ymax = get_comparison_chart_axis_max(
        ground_floor_area_m2=ground_floor_area_m2,
        actual_roof_form=actual_roof_form,
        plan_length_along_ridge_m=plan_length_along_ridge_m,
        plan_length_ridge_to_eaves_m=plan_length_ridge_to_eaves_m,
        part_l_required_generation_kwh=part_l_required_generation_kwh,
        max_feasible_panels=max_feasible_panels,
        module_power_kwp=module_power_kwp,
        sap_placeholder_specific_yield=sap_placeholder_specific_yield,
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

    st.markdown("**Interactive roof editor (beta)**")
    st.caption(
        "Use the controls below to move and resize obstacles and PV rectangles. "
        "Movement step defaults to 0.10 m."
    )

    source_state = build_visual_obstacle_editor(source_state)
    st.session_state["section2_editor_source_state"] = deepcopy(source_state)

    roof_editor_state = deepcopy(source_state)
    roof_editor_state = apply_obstacles_to_pv_zones(roof_editor_state)
    roof_editor_state = regenerate_panels_from_pv_zones(roof_editor_state)
    roof_editor_state = validate_editor_state(roof_editor_state)

    editor_metrics = get_editor_metrics(
        editor_state=roof_editor_state,
        module_power_kwp=module_power_kwp,
        sap_placeholder_specific_yield=sap_placeholder_specific_yield,
    )

    editor_generation_status = format_pass_fail(
        editor_metrics["valid_generation_kwh"],
        part_l_required_generation_kwh,
    )
    editor_kwp_status = format_pass_fail(
        editor_metrics["valid_kwp"],
        part_l_required_kwp,
    )
    editor_panel_status = format_pass_fail(
        float(editor_metrics["valid_panels"]),
        float(part_l_required_panel_count),
    )

    editor_fig = build_editor_roof_packing_diagram(roof_editor_state)
    st.plotly_chart(editor_fig, theme=None, width="stretch", key="section2_editor_chart")

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    editor_summary_cols = st.columns(4)
    with editor_summary_cols[0]:
        render_summary_card(
            "Valid annual generation",
            f"{editor_metrics['valid_generation_kwh']:,.0f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWh/a</span>",
        )
    with editor_summary_cols[1]:
        render_summary_card(
            "Valid kWp",
            f"{editor_metrics['valid_kwp']:,.2f} <span style='font-size:{SUMMARY_UNIT_FONT_SIZE}; font-weight:{SUMMARY_UNIT_FONT_WEIGHT}; color:{SUMMARY_UNIT_COLOUR};'>kWp</span>",
        )
    with editor_summary_cols[2]:
        render_summary_card("Valid panels", f"{editor_metrics['valid_panels']}")
    with editor_summary_cols[3]:
        render_summary_card(
            "Blocked / invalid panels",
            f"{editor_metrics['blocked_panels'] + editor_metrics['invalid_panels']}",
        )

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    editor_comparison_fig = build_comparison_chart(
        required_generation_kwh=part_l_required_generation_kwh,
        actual_generation_kwh=editor_metrics["valid_generation_kwh"],
        y_axis_max=comparison_ymax,
    )
    st.plotly_chart(
        editor_comparison_fig,
        theme=None,
        width="stretch",
        key="section2_editor_comparison_chart",
    )

# -----------------------------------------------------------------------------
# PySAM forecast
# -----------------------------------------------------------------------------
render_section_title("pysam_forecast", "PySAM forecast")
with st.container(border=True):
    st.caption(
        "This section is secondary. It reuses the arrays defined above and does not affect the Part L calculations."
    )

    epw_lookup = get_available_epw_files(EPW_DIRECTORY)
    epw_labels = ["None"] + list(epw_lookup.keys())

    pysam_input_cols = st.columns(2)
    with pysam_input_cols[0]:
        epw_label = st.selectbox("PV yield weather file (EPW)", epw_labels, index=0)
    with pysam_input_cols[1]:
        st.text_input(
            "Array definitions source",
            value="Inherited from Building PV layout",
            disabled=True,
        )

    pysam_result = None
    pysam_message = None
    selected_epw = None

    if installed_panel_count < 1:
        pysam_message = "No installed panels defined in the building section."
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
                        "Tilt / panel pitch (deg)": f"{arr.tilt_deg:.0f}",
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
        st.dataframe(annual_gen_df, hide_index=True, width="stretch")

        st.dataframe(pd.DataFrame(pysam_result["array_rows"]), hide_index=True, width="stretch")

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
        st.dataframe(pysam_assumptions_df, hide_index=True, width="stretch")

        monthly_df = pd.DataFrame(
            {
                "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
                "AC generation (kWh)": [round(v, 1) for v in pysam_result["monthly_ac_kwh"]],
            }
        )
        st.dataframe(monthly_df, hide_index=True, width="stretch")
    else:
        st.info(pysam_message or "Annual generation not available.")

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

user_inputs_rows = [
    ("Part L target", "House form", house_form),
    ("Part L target", "Ground floor area method", gfa_input_mode),
    ("Part L target", "Ground floor area", f"{ground_floor_area_m2:,.2f} m²"),
    ("Part L target", "Ground floor area source", gfa_source_text),
    ("Part L target", "SAP compliance region", sap_compliance_region),
    ("Building PV layout", "Roof type", actual_roof_form),
    ("Building PV layout", "Length along ridge / whole roof length in plan", f"{plan_length_along_ridge_m:,.2f} m"),
    ("Building PV layout", "Ridge-to-eaves / whole roof width in plan", f"{plan_length_ridge_to_eaves_m:,.2f} m"),
    ("Building PV layout", "Roof reduction method", offset_mode_section_2),
]

if actual_roof_form == "Flat":
    user_inputs_rows.append(("Building PV layout", "Flat panel pitch above horizontal", f"{flat_panel_pitch_deg:.0f}°"))
else:
    user_inputs_rows.append(("Building PV layout", "Roof plane azimuth", f"{mono_or_duo_azimuth_deg:.0f}°"))
    user_inputs_rows.append(("Building PV layout", "Roof pitch", f"{mono_or_duo_pitch_deg:.0f}°"))

if offset_mode_section_2 == "Simple perimeter margin":
    user_inputs_rows.append(("Building PV layout", "Perimeter margin around PV zone", f"{perimeter_margin_m:.2f} m"))
else:
    user_inputs_rows.append(("Building PV layout", "Ridge offset", f"{ridge_offset_m:.2f} m"))
    user_inputs_rows.append(("Building PV layout", "Roof edge offset", f"{edge_offset_m:.2f} m"))
    user_inputs_rows.append(("Building PV layout", "Party wall offset", f"{party_wall_offset_m:.2f} m"))

user_inputs_rows.extend(
    [
        ("Building PV layout", "Module width", f"{module_width_m * 1000:.0f} mm"),
        ("Building PV layout", "Module length", f"{module_length_m * 1000:.0f} mm"),
        ("Building PV layout", "Module efficiency", f"{module_efficiency_pct:,.1f} %"),
        ("Building PV layout", "Mount orientation", module_mount_orientation),
        ("Building PV layout", "Installed panel count", f"{installed_panel_count}"),
        ("PySAM forecast", "Selected EPW", epw_label),
    ]
)

calculation_assumption_rows = [
    ("Part L target", "Required PV area fraction", f"{FHS_REQUIRED_AREA_FRACTION:.2f} of ground floor area"),
    ("Part L target", "Standard panel efficiency density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:.2f} kWp/m²"),
    ("Part L target", "SAP annual generation basis", "Placeholder annual yields in code"),
    ("Part L target", "Standardised module width", f"{STANDARDISED_MODULE_WIDTH_M * 1000:.0f} mm"),
    ("Part L target", "Standardised module length", f"{STANDARDISED_MODULE_LENGTH_M * 1000:.0f} mm"),
    ("Part L target", "Standardised module efficiency", f"{STANDARDISED_MODULE_EFFICIENCY_PCT:.1f} %"),
    ("Building PV layout", "Roof planes used", f"{len(roof_planes)}"),
    ("Building PV layout", "Total gross roof area", f"{total_gross_roof_area_m2:,.2f} m²"),
    ("Building PV layout", "Other reduction area from margins / offsets", f"{other_reduction_area_m2:,.2f} m²"),
    ("Building PV layout", "Usable roof area for PV", f"{usable_available_pv_area_m2:,.2f} m²"),
    ("Building PV layout", "Derived module power", f"{module_power_wp:,.0f} Wp"),
    ("Building PV layout", "Maximum feasible panel count", f"{max_feasible_panels}"),
    ("Building PV layout", "Actual building annual generation", f"{actual_building_generation_kwh:,.0f} kWh/a"),
    ("Building PV layout", "Actual building kWp", f"{actual_building_kwp:,.2f} kWp"),
    ("Building PV layout", "Generation check against Part L requirement", actual_generation_status),
    ("Building PV layout", "kWp check against Part L requirement", actual_kwp_status),
    ("Building PV layout", "Panel count check against Part L requirement", actual_panel_status),
    ("Building PV layout", "Editor valid annual generation", f"{editor_metrics['valid_generation_kwh']:,.0f} kWh/a"),
    ("Building PV layout", "Editor valid kWp", f"{editor_metrics['valid_kwp']:,.2f} kWp"),
    ("Building PV layout", "Editor valid panels", f"{editor_metrics['valid_panels']}"),
    ("Building PV layout", "Editor blocked panels", f"{editor_metrics['blocked_panels']}"),
    ("Building PV layout", "Editor invalid panels", f"{editor_metrics['invalid_panels']}"),
    ("Building PV layout", "Editor generation check against Part L requirement", editor_generation_status),
    ("Building PV layout", "Editor kWp check against Part L requirement", editor_kwp_status),
    ("Building PV layout", "Editor panel count check against Part L requirement", editor_panel_status),
    ("Building PV layout", "Editor movement step default", f"{OBSTACLE_NUDGE_STEP_DEFAULT:.2f} m"),
    ("PySAM forecast", "PySAM availability", "Installed" if pvwatts is not None else "Not installed"),
    ("PySAM forecast", "System losses", f"{PYSAM_SYSTEM_LOSSES_PCT:.1f} %"),
    ("PySAM forecast", "DC/AC ratio", f"{PYSAM_DC_AC_RATIO:.2f}"),
    ("PySAM forecast", "Array type", "Fixed roof mount"),
    ("PySAM forecast", "Module type", "Standard"),
    ("PySAM forecast", "Ground coverage ratio", f"{PYSAM_GCR:.2f}"),
]

if pysam_result is not None:
    calculation_assumption_rows.extend(
        [
            ("PySAM forecast", "Total system capacity used", f"{actual_building_kwp:,.2f} kWp"),
            ("PySAM forecast", "Annual AC generation", f"{pysam_result['annual_ac_kwh']:,.0f} kWh/a"),
        ]
    )
else:
    calculation_assumption_rows.append(("PySAM forecast", "Run status", pysam_message or "Annual generation not available."))

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

render_section_title("roof_plane_table", "Roof plane table")
with st.container(border=True):
    plane_rows = build_plane_table_rows(
        roof_planes=roof_planes,
        display_panel_counts_for_planes=display_panel_counts_for_planes,
        actual_roof_form=actual_roof_form,
        flat_panel_pitch_deg=flat_panel_pitch_deg,
        plane_layouts=plane_layouts,
    )
    st.dataframe(pd.DataFrame(plane_rows), hide_index=True, width="stretch")

render_section_title("editor_json", "Roof editor state JSON")
with st.container(border=True):
    st.caption("This is the serialized roof geometry and layout payload for the interactive editor.")
    st.json(roof_editor_state, expanded=False)
    st.code(json.dumps(roof_editor_state, indent=2), language="json")

render_section_title("method_summary", "Method summary")
with st.container(border=True):
    st.markdown(
        """
This tool is split into three main sections.

**Part L target** calculates the standardised Part L reference requirement from ground floor area and outputs required annual generation, required kWp and required panel count using fixed standardised module assumptions.

**Building PV layout** uses dimension-based roof planes:
- **Flat**: one rectangular roof in plan with user-defined panel pitch above horizontal.
- **Mono-pitch**: one sloping plane derived from plan dimensions and roof pitch.
- **Duo-pitch**: two identical sloping planes, with the second rotated by 180°.

The roof editor works in two stages:
- obstacles can be added first, but this step is optional;
- PV rectangles are then defined within the usable packing area;
- panels are regenerated automatically within those PV rectangles;
- movement and resizing are handled through native Streamlit controls;
- panel validity is checked against the packing area, obstacle intersections, and panel overlap.

**PySAM forecast** reuses the arrays from the building section and applies PySAM PVWatts using a selected EPW file. It does not affect the Part L-side calculation.
"""
    )

render_section_title("current_limits", "Current limits")
with st.container(border=True):
    st.markdown(
        """
- The SAP regional annual yields are placeholders and need checking.
- The roof fit still uses simplified roof reductions rather than a full geometric roof model.
- Flat roofs use a single roof plane in the editor, while generation is still split 50/50 east-west.
- Flat-roof row spacing still needs to be thought through properly.
- Hipped roofs are not included in this simplified method.
- PV rectangles are rectangular only and do not yet support polygon drawing.
- Panels are regenerated from PV rectangles rather than dragged individually.
- The editor uses native Streamlit controls for movement and resizing.
- PySAM uses locally stored EPW files in `resources/epw/`.
"""
    )