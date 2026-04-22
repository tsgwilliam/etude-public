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

LOGO_PATH = RESOURCE_DIR / "Etude-logo-animation-v005 Single spin.gif"
LOGO_WIDTH = 180

CHART_HEIGHT = 420
CHART_BACKGROUND_COLOUR = "white"
CHART_PLOT_BACKGROUND_COLOUR = "white"
CHART_FONT_COLOUR = "#333333"
CHART_GRID_COLOUR = "#E6E6E6"
CHART_AXIS_LINE_COLOUR = "#BFBFBF"
CHART_MARGIN = dict(l=20, r=20, t=24, b=20)
CHART_BAR_COLOURS = ["#4F67FF", "#F05A3A"]
CHART_BAR_WIDTH = 0.19
CHART_BARGAP = 0.45

# -----------------------------------------------------------------------------
# Section 2 diagram styling
# -----------------------------------------------------------------------------
DIAGRAM_SINGLE_HEIGHT = 760
DIAGRAM_DUO_HEIGHT = 760
DIAGRAM_MARGIN = dict(l=30, r=30, t=30, b=30)
DIAGRAM_PANEL_COLOUR = "#17C497"
DIAGRAM_ROOF_COLOUR = "#DCE6F9"
DIAGRAM_USABLE_COLOUR = "#FBE3D6"
DIAGRAM_BLOCKED_COLOUR = "#BDBDBD"
DIAGRAM_LINE_COLOUR = "#444444"
DIAGRAM_TEXT_COLOUR = "#333333"
DIAGRAM_PLANE_GAP_X = 2.5
DIAGRAM_PLANE_GAP_Y = 3.3


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


def count_modules_on_rectangle(
    usable_length_m: float,
    usable_depth_m: float,
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
    roof_form: str,
    flat_panel_pitch_deg: float,
) -> int:
    module_x_m, module_y_m = get_module_footprint(
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        mount_orientation=mount_orientation,
        roof_form=roof_form,
        flat_panel_pitch_deg=flat_panel_pitch_deg,
    )

    count_x = math.floor(usable_length_m / module_x_m) if module_x_m > 0 else 0
    count_y = math.floor(usable_depth_m / module_y_m) if module_y_m > 0 else 0
    return max(count_x, 0) * max(count_y, 0)


def calc_plane_panel_layout(
    packing_length_m: float,
    packing_depth_m: float,
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
    roof_form: str,
    flat_panel_pitch_deg: float,
) -> dict:
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

    return {
        "panel_w_m": panel_w_m,
        "panel_h_m": panel_h_m,
        "cols": cols,
        "rows": rows,
        "count": count,
        "used_width_m": used_width_m,
        "used_height_m": used_height_m,
    }


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
    plane: dict,
) -> None:
    left_margin = plane["margin_left_m"]
    right_margin = plane["margin_right_m"]
    top_margin = plane["margin_top_m"]
    bottom_margin = plane["margin_bottom_m"]

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


def add_roof_plane_to_figure(
    fig: go.Figure,
    plane: dict,
    layout: dict,
    display_count: int,
    origin_x: float,
    origin_y: float,
    roof_form: str,
    flat_panel_pitch_deg: float,
) -> tuple[float, float]:
    gross_w = plane["gross_length_m"]
    gross_h = plane["gross_depth_m"]
    usable_w = plane["usable_length_m"]
    usable_h = plane["usable_depth_m"]

    roof_x0 = origin_x
    roof_y0 = origin_y
    roof_x1 = roof_x0 + gross_w
    roof_y1 = roof_y0 + gross_h

    fig.add_shape(
        type="rect",
        x0=roof_x0,
        y0=roof_y0,
        x1=roof_x1,
        y1=roof_y1,
        line=dict(color=DIAGRAM_LINE_COLOUR, width=2),
        fillcolor=DIAGRAM_ROOF_COLOUR,
    )

    usable_x0 = roof_x0 + plane["margin_left_m"]
    usable_x1 = roof_x1 - plane["margin_right_m"]
    usable_y0 = roof_y0 + plane["margin_top_m"]
    usable_y1 = roof_y1 - plane["margin_bottom_m"]

    fig.add_shape(
        type="rect",
        x0=usable_x0,
        y0=usable_y0,
        x1=usable_x1,
        y1=usable_y1,
        line=dict(color=DIAGRAM_LINE_COLOUR, width=1, dash="dash"),
        fillcolor=DIAGRAM_USABLE_COLOUR,
    )

    add_margin_annotations(
        fig=fig,
        roof_x0=roof_x0,
        roof_y0=roof_y0,
        roof_x1=roof_x1,
        roof_y1=roof_y1,
        usable_x0=usable_x0,
        usable_y0=usable_y0,
        usable_x1=usable_x1,
        usable_y1=usable_y1,
        plane=plane,
    )

    blocked_band_depth = plane["blocked_band_depth_m"]
    packing_y0 = usable_y0
    packing_y1 = usable_y1

    if blocked_band_depth > 0:
        blocked_y0 = usable_y1 - blocked_band_depth
        blocked_y1 = usable_y1

        fig.add_shape(
            type="rect",
            x0=usable_x0,
            y0=blocked_y0,
            x1=usable_x1,
            y1=blocked_y1,
            line=dict(color=DIAGRAM_LINE_COLOUR, width=1),
            fillcolor=DIAGRAM_BLOCKED_COLOUR,
        )

        fig.add_annotation(
            x=(usable_x0 + usable_x1) / 2.0,
            y=(blocked_y0 + blocked_y1) / 2.0,
            showarrow=False,
            text=f"blocked band {blocked_band_depth:.2f} m",
            font=dict(color=DIAGRAM_TEXT_COLOUR, size=10),
        )

        packing_y1 = blocked_y0

    panel_w = layout["panel_w_m"]
    panel_h = layout["panel_h_m"]
    cols = layout["cols"]
    rows = layout["rows"]
    to_draw = min(display_count, layout["count"])

    used_width = layout["used_width_m"]
    used_height = layout["used_height_m"]

    packing_w = usable_x1 - usable_x0
    packing_h = packing_y1 - packing_y0

    start_x = usable_x0 + max((packing_w - used_width) / 2.0, 0.0)
    start_y = packing_y0 + max((packing_h - used_height) / 2.0, 0.0)

    drawn = 0
    for r in range(rows):
        for c in range(cols):
            if drawn >= to_draw:
                break

            px0 = start_x + c * panel_w
            py0 = start_y + r * panel_h
            px1 = px0 + panel_w
            py1 = py0 + panel_h

            fig.add_shape(
                type="rect",
                x0=px0,
                y0=py0,
                x1=px1,
                y1=py1,
                line=dict(color=DIAGRAM_LINE_COLOUR, width=1),
                fillcolor=DIAGRAM_PANEL_COLOUR,
            )
            drawn += 1
        if drawn >= to_draw:
            break

    roof_type_label = f"{roof_form} roof"
    if roof_form == "Flat":
        label_text = f"{roof_type_label} - single roof plane"
        detail_text = (
            f"Azimuth 0° roof tilt 0°<br>"
            f"Panel pitch above horizontal {flat_panel_pitch_deg:.0f}°<br>"
            f"Gross {gross_w:.2f} × {gross_h:.2f} m<br>"
            f"Usable {usable_w:.2f} × {usable_h:.2f} m<br>"
            f"Fit {layout['cols']} × {layout['rows']} = {layout['count']} panels"
        )
    else:
        label_text = f"{roof_type_label} - {plane['name']}"
        detail_text = (
            f"Azimuth {plane['azimuth_deg']:.0f}° roof tilt {plane['tilt_deg']:.0f}°<br>"
            f"Gross {gross_w:.2f} × {gross_h:.2f} m<br>"
            f"Usable {usable_w:.2f} × {usable_h:.2f} m<br>"
            f"Fit {layout['cols']} × {layout['rows']} = {layout['count']} panels"
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

    hover_text = (
        f"{label_text}<br>"
        f"Gross: {gross_w:.2f} × {gross_h:.2f} m<br>"
        f"Usable: {usable_w:.2f} × {usable_h:.2f} m<br>"
        f"Blocked area: {plane['blocked_area_m2']:.2f} m²<br>"
        f"Max panels: {layout['count']}"
    )
    if roof_form == "Flat":
        hover_text += f"<br>Panel pitch above horizontal: {flat_panel_pitch_deg:.0f}°"

    fig.add_trace(
        go.Scatter(
            x=[(roof_x0 + roof_x1) / 2.0],
            y=[(roof_y0 + roof_y1) / 2.0],
            mode="markers",
            marker=dict(size=18, color="rgba(0,0,0,0)"),
            hovertemplate=hover_text + "<extra></extra>",
            showlegend=False,
        )
    )

    return roof_x1, roof_y1 + 1.6


def build_roof_packing_diagram(
    roof_form: str,
    roof_planes: list[dict],
    module_length_m: float,
    module_width_m: float,
    mount_orientation: str,
    installed_panel_count: int,
    flat_panel_pitch_deg: float,
) -> go.Figure:
    fig = go.Figure()

    if not roof_planes:
        fig.update_layout(
            height=250,
            margin=DIAGRAM_MARGIN,
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
        )
        return fig

    layouts = []
    total_max = 0
    for plane in roof_planes:
        layout = calc_plane_panel_layout(
            packing_length_m=plane["packing_length_m"],
            packing_depth_m=plane["packing_depth_m"],
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            mount_orientation=mount_orientation,
            roof_form=roof_form,
            flat_panel_pitch_deg=flat_panel_pitch_deg,
        )
        layouts.append(layout)
        total_max += layout["count"]

    if total_max > 0 and installed_panel_count > 0:
        displayed_counts = allocate_integer_counts(
            total_count=min(installed_panel_count, total_max),
            share_fractions=[layout["count"] / total_max for layout in layouts],
        )
    else:
        displayed_counts = [0] * len(layouts)

    max_x = 0.0
    max_y = 0.0

    if roof_form == "Duo-pitch" and len(roof_planes) == 2:
        current_x = 0.0
        for plane, layout, display_count in zip(roof_planes, layouts, displayed_counts):
            roof_x1, content_y1 = add_roof_plane_to_figure(
                fig=fig,
                plane=plane,
                layout=layout,
                display_count=display_count,
                origin_x=current_x,
                origin_y=0.0,
                roof_form=roof_form,
                flat_panel_pitch_deg=flat_panel_pitch_deg,
            )
            max_x = max(max_x, roof_x1)
            max_y = max(max_y, content_y1)
            current_x = roof_x1 + DIAGRAM_PLANE_GAP_X
    else:
        current_y = 0.0
        for plane, layout, display_count in zip(roof_planes, layouts, displayed_counts):
            roof_x1, content_y1 = add_roof_plane_to_figure(
                fig=fig,
                plane=plane,
                layout=layout,
                display_count=display_count,
                origin_x=0.0,
                origin_y=current_y,
                roof_form=roof_form,
                flat_panel_pitch_deg=flat_panel_pitch_deg,
            )
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

        usable_length_m = max(gross_length_m - margin_left - margin_right, 0.0)
        usable_depth_m = max(gross_depth_m - margin_top - margin_bottom, 0.0)

        planes.append(
            {
                "name": "Roof plane 1",
                "azimuth_deg": 0.0,
                "tilt_deg": 0.0,
                "gross_length_m": gross_length_m,
                "gross_depth_m": gross_depth_m,
                "margin_left_m": margin_left,
                "margin_right_m": margin_right,
                "margin_top_m": margin_top,
                "margin_bottom_m": margin_bottom,
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

        usable_length_m = max(gross_length_m - margin_left - margin_right, 0.0)
        usable_depth_m = max(gross_depth_m - margin_top - margin_bottom, 0.0)

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
                    "margin_left_m": margin_left,
                    "margin_right_m": margin_right,
                    "margin_top_m": margin_top,
                    "margin_bottom_m": margin_bottom,
                    "usable_length_m": usable_length_m,
                    "usable_depth_m": usable_depth_m,
                    "gross_area_m2": gross_length_m * gross_depth_m,
                    "usable_area_before_blocked_m2": usable_length_m * usable_depth_m,
                }
            )

    blocked_area_per_plane_m2 = blocked_area_total_m2 / max(len(planes), 1)
    for plane in planes:
        plane["blocked_area_m2"] = blocked_area_per_plane_m2

        if plane["usable_length_m"] > 0:
            blocked_band_depth_m = min(
                blocked_area_per_plane_m2 / plane["usable_length_m"],
                plane["usable_depth_m"],
            )
        else:
            blocked_band_depth_m = 0.0

        plane["blocked_band_depth_m"] = blocked_band_depth_m
        plane["packing_length_m"] = plane["usable_length_m"]
        plane["packing_depth_m"] = max(plane["usable_depth_m"] - blocked_band_depth_m, 0.0)
        plane["usable_area_after_blocked_m2"] = plane["packing_length_m"] * plane["packing_depth_m"]

    return planes


def build_actual_arrays_for_generation(
    roof_form: str,
    mono_or_duo_azimuth_deg: float | None,
    mono_or_duo_pitch_deg: float,
    flat_panel_pitch_deg: float,
) -> list[ArrayDefinition]:
    if roof_form == "Mono-pitch":
        return [
            ArrayDefinition(
                name="Roof plane 1",
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
# Part L requirement
# -----------------------------------------------------------------------------
with st.expander("Part L requirement", expanded=True):
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

# -----------------------------------------------------------------------------
# Your Building measured by Part L
# -----------------------------------------------------------------------------
with st.expander("Your Building measured by Part L", expanded=True):
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

    module_area_m2 = module_length_m * module_width_m
    module_power_kwp = module_power_kwp_from_inputs(
        length_m=module_length_m,
        width_m=module_width_m,
        efficiency_pct=module_efficiency_pct,
    )
    module_power_wp = module_power_kwp * 1000.0

    for plane in roof_planes:
        packed_by_dimensions = count_modules_on_rectangle(
            usable_length_m=plane["packing_length_m"],
            usable_depth_m=plane["packing_depth_m"],
            module_length_m=module_length_m,
            module_width_m=module_width_m,
            mount_orientation=module_mount_orientation,
            roof_form=actual_roof_form,
            flat_panel_pitch_deg=flat_panel_pitch_deg,
        )
        plane["max_feasible_panels"] = packed_by_dimensions

    max_feasible_panels = sum(plane["max_feasible_panels"] for plane in roof_planes)

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

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    packing_fig = build_roof_packing_diagram(
        roof_form=actual_roof_form,
        roof_planes=roof_planes,
        module_length_m=module_length_m,
        module_width_m=module_width_m,
        mount_orientation=module_mount_orientation,
        installed_panel_count=installed_panel_count,
        flat_panel_pitch_deg=flat_panel_pitch_deg,
    )
    st.plotly_chart(packing_fig, theme=None, use_container_width=True)

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

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

    comparison_fig = build_comparison_chart(
        required_generation_kwh=part_l_required_generation_kwh,
        actual_generation_kwh=actual_building_generation_kwh,
        y_axis_max=comparison_ymax,
    )
    st.plotly_chart(comparison_fig, theme=None, use_container_width=True)

    display_panel_counts_for_planes = (
        actual_array_panel_counts
        if actual_roof_form != "Flat"
        else [installed_panel_count]
    )

    plane_rows = []
    for plane, installed_panels in zip(roof_planes, display_panel_counts_for_planes):
        if actual_roof_form == "Flat":
            arr_factor = (
                sap_orientation_factor_placeholder(90.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
                + sap_orientation_factor_placeholder(270.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
            ) / 2.0
        else:
            arr_factor = sap_orientation_factor_placeholder(plane["azimuth_deg"]) * sap_tilt_factor_placeholder(
                plane["tilt_deg"]
            )

        plane_rows.append(
            {
                "Plane": plane["name"],
                "Azimuth (deg)": f"{plane['azimuth_deg']:.0f}",
                "Roof tilt (deg)": f"{plane['tilt_deg']:.0f}",
                "Gross length (m)": f"{plane['gross_length_m']:.2f}",
                "Gross depth (m)": f"{plane['gross_depth_m']:.2f}",
                "Left margin (m)": f"{plane['margin_left_m']:.2f}",
                "Right margin (m)": f"{plane['margin_right_m']:.2f}",
                "Top margin (m)": f"{plane['margin_top_m']:.2f}",
                "Bottom margin (m)": f"{plane['margin_bottom_m']:.2f}",
                "Blocked band depth (m)": f"{plane['blocked_band_depth_m']:.2f}",
                "Packing length (m)": f"{plane['packing_length_m']:.2f}",
                "Packing depth (m)": f"{plane['packing_depth_m']:.2f}",
                "Max feasible panels": f"{plane['max_feasible_panels']}",
                "Displayed / associated panels": f"{installed_panels}",
                "SAP placeholder factor": f"{arr_factor:.3f}",
            }
        )
    st.dataframe(pd.DataFrame(plane_rows), hide_index=True, use_container_width=True)

# -----------------------------------------------------------------------------
# PySAM annual generation forecast
# -----------------------------------------------------------------------------
with st.expander("PySAM annual generation forecast", expanded=True):
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
            value="Inherited from Your Building measured by Part L",
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

st.divider()

# -----------------------------------------------------------------------------
# Bottom summary tables
# -----------------------------------------------------------------------------
total_gross_roof_area_m2 = sum(plane["gross_area_m2"] for plane in roof_planes)
total_usable_before_blocked_m2 = sum(plane["usable_area_before_blocked_m2"] for plane in roof_planes)
usable_available_pv_area_m2 = sum(plane["usable_area_after_blocked_m2"] for plane in roof_planes)
other_reduction_area_m2 = max(total_gross_roof_area_m2 - total_usable_before_blocked_m2, 0.0)

display_panel_counts_for_planes = (
    actual_array_panel_counts
    if actual_roof_form != "Flat"
    else [installed_panel_count]
)

user_inputs_rows = [
    ("Part L", "House form", house_form),
    ("Part L", "Ground floor area method", gfa_input_mode),
    ("Part L", "Ground floor area", f"{ground_floor_area_m2:,.2f} m²"),
    ("Part L", "Ground floor area source", gfa_source_text),
    ("Part L", "SAP compliance region", sap_compliance_region),
    ("Building", "Roof type", actual_roof_form),
    ("Building", "Length along ridge / whole roof length in plan", f"{plan_length_along_ridge_m:,.2f} m"),
    ("Building", "Ridge-to-eaves / whole roof width in plan", f"{plan_length_ridge_to_eaves_m:,.2f} m"),
    ("Building", "Blocked area", f"{blocked_area_total_m2:,.2f} m²"),
    ("Building", "Roof reduction method", offset_mode_section_2),
]

if actual_roof_form == "Flat":
    user_inputs_rows.append(("Building", "Flat panel pitch above horizontal", f"{flat_panel_pitch_deg:.0f}°"))
else:
    user_inputs_rows.append(("Building", "Roof plane azimuth", f"{mono_or_duo_azimuth_deg:.0f}°"))
    user_inputs_rows.append(("Building", "Roof pitch", f"{mono_or_duo_pitch_deg:.0f}°"))

if offset_mode_section_2 == "Simple perimeter margin":
    user_inputs_rows.append(("Building", "Perimeter margin around PV zone", f"{perimeter_margin_m:.2f} m"))
else:
    user_inputs_rows.append(("Building", "Ridge offset", f"{ridge_offset_m:.2f} m"))
    user_inputs_rows.append(("Building", "Roof edge offset", f"{edge_offset_m:.2f} m"))
    user_inputs_rows.append(("Building", "Party wall offset", f"{party_wall_offset_m:.2f} m"))

user_inputs_rows.extend(
    [
        ("Building", "Module width", f"{module_width_m * 1000:.0f} mm"),
        ("Building", "Module length", f"{module_length_m * 1000:.0f} mm"),
        ("Building", "Module efficiency", f"{module_efficiency_pct:,.1f} %"),
        ("Building", "Mount orientation", module_mount_orientation),
        ("Building", "Installed panel count", f"{installed_panel_count}"),
        ("PySAM", "Selected EPW", epw_label),
    ]
)

calculation_assumption_rows = [
    ("Part L", "Required PV area fraction", f"{FHS_REQUIRED_AREA_FRACTION:.2f} of ground floor area"),
    ("Part L", "Standard panel efficiency density", f"{STANDARD_PANEL_EFFICIENCY_KWP_PER_M2:.2f} kWp/m²"),
    ("Part L", "SAP annual generation basis", "Placeholder annual yields in code"),
    ("Part L", "Standardised module width", f"{STANDARDISED_MODULE_WIDTH_M * 1000:.0f} mm"),
    ("Part L", "Standardised module length", f"{STANDARDISED_MODULE_LENGTH_M * 1000:.0f} mm"),
    ("Part L", "Standardised module efficiency", f"{STANDARDISED_MODULE_EFFICIENCY_PCT:.1f} %"),
    ("Building", "Roof planes used", f"{len(roof_planes)}"),
    ("Building", "Total gross roof area", f"{total_gross_roof_area_m2:,.2f} m²"),
    ("Building", "Other reduction area from margins / offsets", f"{other_reduction_area_m2:,.2f} m²"),
    ("Building", "Usable roof area after blocked band", f"{usable_available_pv_area_m2:,.2f} m²"),
    ("Building", "Derived module power", f"{module_power_wp:,.0f} Wp"),
    ("Building", "Maximum feasible panel count", f"{max_feasible_panels}"),
    ("Building", "Actual building annual generation", f"{actual_building_generation_kwh:,.0f} kWh/a"),
    ("Building", "Actual building kWp", f"{actual_building_kwp:,.2f} kWp"),
    ("Building", "Generation check against Part L requirement", actual_generation_status),
    ("Building", "kWp check against Part L requirement", actual_kwp_status),
    ("Building", "Panel count check against Part L requirement", actual_panel_status),
    ("PySAM", "PySAM availability", "Installed" if pvwatts is not None else "Not installed"),
    ("PySAM", "System losses", f"{PYSAM_SYSTEM_LOSSES_PCT:.1f} %"),
    ("PySAM", "DC/AC ratio", f"{PYSAM_DC_AC_RATIO:.2f}"),
    ("PySAM", "Array type", "Fixed roof mount"),
    ("PySAM", "Module type", "Standard"),
    ("PySAM", "Ground coverage ratio", f"{PYSAM_GCR:.2f}"),
]

if pysam_result is not None:
    calculation_assumption_rows.extend(
        [
            ("PySAM", "Total system capacity used", f"{actual_building_kwp:,.2f} kWp"),
            ("PySAM", "Annual AC generation", f"{pysam_result['annual_ac_kwh']:,.0f} kWh/a"),
        ]
    )
else:
    calculation_assumption_rows.append(("PySAM", "Run status", pysam_message or "Annual generation not available."))

st.markdown("### Inputs added by user")
user_inputs_df = pd.DataFrame(user_inputs_rows, columns=["Section", "Input", "Value"])
st.dataframe(user_inputs_df, hide_index=True, use_container_width=True)

st.markdown("### Calculation assumptions")
calculation_assumptions_df = pd.DataFrame(
    calculation_assumption_rows,
    columns=["Section", "Assumption", "Value"],
)
st.dataframe(calculation_assumptions_df, hide_index=True, use_container_width=True)

st.markdown("### Roof plane table")
plane_rows = []
for plane, installed_panels in zip(roof_planes, display_panel_counts_for_planes):
    if actual_roof_form == "Flat":
        arr_factor = (
            sap_orientation_factor_placeholder(90.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
            + sap_orientation_factor_placeholder(270.0) * sap_tilt_factor_placeholder(flat_panel_pitch_deg)
        ) / 2.0
    else:
        arr_factor = sap_orientation_factor_placeholder(plane["azimuth_deg"]) * sap_tilt_factor_placeholder(
            plane["tilt_deg"]
        )

    plane_rows.append(
        {
            "Plane": plane["name"],
            "Azimuth (deg)": f"{plane['azimuth_deg']:.0f}",
            "Roof tilt (deg)": f"{plane['tilt_deg']:.0f}",
            "Gross length (m)": f"{plane['gross_length_m']:.2f}",
            "Gross depth (m)": f"{plane['gross_depth_m']:.2f}",
            "Left margin (m)": f"{plane['margin_left_m']:.2f}",
            "Right margin (m)": f"{plane['margin_right_m']:.2f}",
            "Top margin (m)": f"{plane['margin_top_m']:.2f}",
            "Bottom margin (m)": f"{plane['margin_bottom_m']:.2f}",
            "Blocked band depth (m)": f"{plane['blocked_band_depth_m']:.2f}",
            "Packing length (m)": f"{plane['packing_length_m']:.2f}",
            "Packing depth (m)": f"{plane['packing_depth_m']:.2f}",
            "Max feasible panels": f"{plane['max_feasible_panels']}",
            "Displayed / associated panels": f"{installed_panels}",
            "SAP placeholder factor": f"{arr_factor:.3f}",
        }
    )
st.dataframe(pd.DataFrame(plane_rows), hide_index=True, use_container_width=True)

with st.expander("Method summary", expanded=False):
    st.markdown(
        """
This tool is split into three collapsible sections.

**Part L requirement** calculates the standardised Part L reference requirement from ground floor area and outputs required annual generation, required kWp and required panel count using fixed standardised module assumptions.

**Your Building measured by Part L** uses dimension-based roof planes:
- **Flat roof**: whole roof length and width in plan, with user-defined panel pitch above horizontal;
- **Mono-pitch**: ridge length and ridge-to-eaves length in plan, then roof pitch and azimuth;
- **Duo-pitch**: same entered roof plane duplicated and rotated by 180°.

Panel counting is dimension-based, so portrait and landscape can produce different results. The diagram shows gross roof outline, usable inset area, blocked band and fitted modules.

**PySAM annual generation forecast** reuses the arrays from the building section and applies PySAM PVWatts using a selected EPW file. It does not affect the Part L-side calculation.
"""
    )

with st.expander("Current limits", expanded=False):
    st.markdown(
        """
- The SAP regional annual yields are placeholders and need checking.
- Section 2 still uses simplified roof reductions and does not attempt full obstacle geometry.
- Flat roofs use a single roof plane in the diagram, but generation is split 50/50 east-west.
- Flat-roof row spacing still needs to be thought through properly.
- Hipped roofs are not included in the simplified Section 2 method.
- Blocked area is represented as a blocked band for packing, not as a true obstacle cut-out.
- The roof packing diagram is illustrative rather than a formal layout drawing.
- PySAM uses locally stored EPW files in `resources/epw/`.
"""
    )