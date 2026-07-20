from __future__ import annotations

import colorsys

import pandas as pd
import plotly.graph_objects as go

from oneclick_core.nrm import build_high_order, parse_rics_numeric_tuple

UPFRONT_MODULES = ["A1-A5"]
DEFAULT_WLC_MODULES = [
    "A1-A3", "A4", "A5", "B1", "B2", "B3", "B4", "B5", "C1", "C2", "C3", "C4",
]


def expand_modules(modules: list[str]) -> list[str]:
    expanded: list[str] = []
    for module in modules:
        if module == "A1-A5":
            expanded.extend(["A1-A5", "A1-A3", "A4", "A5"])
        else:
            expanded.append(module)
    return list(dict.fromkeys(expanded))


def pick_chart_value_col(df: pd.DataFrame, em_col: str = "kgco2e") -> str:
    if "rics_allocated_value" in df.columns:
        return "rics_allocated_value"
    return em_col


def get_default_rics_colour(name: str) -> str:
    low = str(name).lower()
    mapping = {
        "superstructure": "#23A846",
        "substructure": "#E8C602",
        "finish": "#EC8447",
        "service": "#6BA1D5",
        "fitting": "#A56CC1",
        "prefabricated": "#00A6A6",
        "existing": "#C06C84",
        "external": "#F67280",
        "prelim": "#6C6C6C",
        "deconstruction": "#999999",
    }
    for key, colour in mapping.items():
        if key in low:
            return colour
    return "#4c78a8"


def _hex_to_rgb01(hex_color: str) -> tuple[float, float, float]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0


def _rgb01_to_hex(rgb: tuple[float, float, float]) -> str:
    r, g, b = rgb
    return "#{:02x}{:02x}{:02x}".format(
        int(max(0, min(1, r)) * 255),
        int(max(0, min(1, g)) * 255),
        int(max(0, min(1, b)) * 255),
    )


def shade_palette(base_hex: str, n: int) -> list[str]:
    if n <= 1:
        return [base_hex]
    r, g, b = _hex_to_rgb01(base_hex)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l_min = max(0.25, l - 0.22)
    l_max = min(0.75, l + 0.22)
    return [
        _rgb01_to_hex(colorsys.hls_to_rgb(h, l_min + (l_max - l_min) * (i / (n - 1)), s))
        for i in range(n)
    ]


def luminance_from_hex(hex_color: str) -> float:
    r, g, b = _hex_to_rgb01(hex_color)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def compute_biogenic_total(
    rows: pd.DataFrame,
    modules: list[str],
    em_col: str = "kgco2e",
    scale: float = 1.0,
) -> float:
    rr = rows.copy()
    if "section" not in rr.columns:
        return 0.0

    sec = rr["section"].astype(str).str.strip().str.lower()
    bio_rows = rr[sec.eq("bioc")].copy()
    if not bio_rows.empty:
        if "_source_row_id" in bio_rows.columns:
            total = (
                bio_rows.groupby("_source_row_id", dropna=False)[em_col]
                .first()
                .pipe(pd.to_numeric, errors="coerce")
                .fillna(0.0)
                .sum()
            )
        else:
            total = pd.to_numeric(bio_rows[em_col], errors="coerce").fillna(0.0).sum()
        return float(abs(total) * scale)

    bio_col = None
    for col in rr.columns:
        if str(col).lower() in {"biogenic_kgco2e", "biogenic"}:
            bio_col = col
            break
    if bio_col is None:
        return 0.0

    rr = rr[rr["section"].isin(expand_modules(modules))]
    if "_source_row_id" in rr.columns:
        total = (
            rr.groupby("_source_row_id", dropna=False)[bio_col]
            .first()
            .pipe(pd.to_numeric, errors="coerce")
            .fillna(0.0)
            .sum()
        )
    else:
        total = pd.to_numeric(rr[bio_col], errors="coerce").fillna(0.0).sum()
    return float(abs(total) * scale)


def filter_rows_for_modules(rows: pd.DataFrame, modules: list[str]) -> pd.DataFrame:
    return rows[rows["section"].isin(expand_modules(modules))].copy()


def aggregate_chart_data(rows: pd.DataFrame, modules: list[str], em_col: str = "kgco2e") -> pd.DataFrame:
    rr = filter_rows_for_modules(rows, modules)
    value_col = pick_chart_value_col(rr, em_col)
    if rr.empty:
        return pd.DataFrame(columns=["chart_high", "chart_segment_label", "chart_segment_code", "value"])

    labels = (
        rr.groupby(["chart_high", "chart_segment_code"], dropna=False)["chart_segment_label"]
        .agg(lambda s: next((str(x) for x in s if str(x).strip() and str(x).lower() != "nan"), ""))
        .reset_index()
    )
    values = (
        rr.groupby(["chart_high", "chart_segment_code"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .rename(columns={value_col: "value"})
    )
    return (
        values.merge(labels, on=["chart_high", "chart_segment_code"], how="left")
        .sort_values(["chart_high", "chart_segment_code"])
        .reset_index(drop=True)
    )


def _build_segments(
    agg: pd.DataFrame,
    colour_map: dict[str, str] | None = None,
) -> tuple[list[dict], list[str], dict[str, list[str]]]:
    colour_map = colour_map or {}
    high_order = build_high_order(agg["chart_high"].astype(str).unique().tolist())

    seg_order: dict[str, list[str]] = {}
    for high in high_order:
        sub = agg[agg["chart_high"] == high].copy()
        sub["_sort"] = sub["chart_segment_code"].map(parse_rics_numeric_tuple)
        sub = sub.sort_values("_sort")
        # Prefer unique codes so NRM level roll-up collapses correctly.
        ordered_labels = []
        seen = set()
        for _, row in sub.iterrows():
            label = str(row["chart_segment_label"])
            if label not in seen:
                ordered_labels.append(label)
                seen.add(label)
        seg_order[high] = ordered_labels

    segments: list[dict] = []
    for high in high_order:
        palette = shade_palette(colour_map.get(high, get_default_rics_colour(high)), max(1, len(seg_order.get(high, []))))
        for idx, seg in enumerate(seg_order.get(high, [])):
            value = float(agg.loc[agg["chart_segment_label"] == seg, "value"].sum())
            if value <= 1e-9:
                continue
            segments.append(
                {
                    "high": high,
                    "seg": seg,
                    "value": value,
                    "colour": palette[min(idx, len(palette) - 1)],
                }
            )
    return segments, high_order, seg_order


def _aggregate_for_plot(rr: pd.DataFrame, value_col: str, scale: float) -> pd.DataFrame:
    if rr.empty:
        return pd.DataFrame(columns=["chart_high", "chart_segment_label", "chart_segment_code", "value"])
    labels = (
        rr.groupby(["chart_high", "chart_segment_code"], dropna=False)["chart_segment_label"]
        .agg(lambda s: next((str(x) for x in s if str(x).strip() and str(x).lower() != "nan"), ""))
        .reset_index()
    )
    values = (
        rr.groupby(["chart_high", "chart_segment_code"], dropna=False)[value_col]
        .sum()
        .reset_index()
        .rename(columns={value_col: "value"})
    )
    agg = values.merge(labels, on=["chart_high", "chart_segment_code"], how="left")
    agg["value"] = agg["value"] * scale
    return agg


def plot_rics_single_stack(
    rows: pd.DataFrame,
    modules: list[str],
    title: str,
    use_intensity: bool,
    gia_m2: float,
    collapsed_high_level: bool = False,
    colour_map: dict[str, str] | None = None,
    biogenic_colour: str = "#2B0FC9",
    height_px: int = 900,
    bar_width: float = 0.66,
    small_segment_threshold: float = 0.03,
    target_line_value: float | None = None,
    target_line_label: str = "",
    em_col: str = "kgco2e",
) -> go.Figure:
    rr = filter_rows_for_modules(rows, modules)
    value_col = pick_chart_value_col(rr, em_col)
    scale = 1.0 / float(gia_m2) if use_intensity and gia_m2 > 0 else 1.0
    y_unit = "kgCO₂e/m² GIA" if scale != 1.0 else "kgCO₂e"

    agg = _aggregate_for_plot(rr, value_col, scale)
    segments, high_order, _ = _build_segments(agg, colour_map)
    total_pos = sum(s["value"] for s in segments) or 1.0
    fig = go.Figure()

    for seg in segments:
        share = seg["value"] / total_pos
        is_small = share < small_segment_threshold
        text = None if collapsed_high_level or is_small else f"{seg['seg']} {seg['value']:,.0f}"
        fig.add_trace(
            go.Bar(
                x=[""],
                y=[seg["value"]],
                width=bar_width,
                marker=dict(color=seg["colour"], line=dict(color="white", width=1)),
                showlegend=False,
                text=[text] if text else None,
                textposition="inside",
                insidetextanchor="middle",
                hovertemplate=(
                    f"<b>{seg['high']}</b><br>{seg['seg']}: %{{y:,.0f}} {y_unit}<extra></extra>"
                ),
            )
        )

    if collapsed_high_level:
        running = 0.0
        for high in high_order:
            high_value = float(agg.loc[agg["chart_high"] == high, "value"].sum())
            if high_value <= 1e-9:
                continue
            mid = running + high_value / 2.0
            running += high_value
            colour = colour_map.get(high, get_default_rics_colour(high)) if colour_map else get_default_rics_colour(high)
            fig.add_annotation(
                x=0.5,
                xref="paper",
                y=mid,
                text=f"{high} {high_value:,.0f}",
                showarrow=False,
                font=dict(color="white" if luminance_from_hex(colour) < 0.5 else "black", size=12),
            )

    bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
    if bio_total > 1e-9:
        fig.add_trace(
            go.Bar(
                x=[""],
                y=[-abs(bio_total)],
                width=bar_width,
                marker=dict(color=biogenic_colour),
                showlegend=False,
                hovertemplate=f"Biogenic: %{{y:,.0f}} {y_unit}<extra></extra>",
            )
        )

    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        yaxis_title=y_unit,
        showlegend=False,
        margin=dict(l=40, r=40, t=70, b=40),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )

    if target_line_value is not None:
        fig.add_hline(y=target_line_value, line_dash="dot", line_color="#D62728")
        if target_line_label:
            fig.add_annotation(
                x=0.5,
                xref="paper",
                y=target_line_value,
                text=target_line_label,
                showarrow=False,
                bgcolor="rgba(255,255,255,0.8)",
            )
    return fig


def plot_rics_two_stacks(
    rows: pd.DataFrame,
    upfront_modules: list[str],
    whole_life_modules: list[str],
    title: str,
    use_intensity: bool,
    gia_m2: float,
    collapsed_high_level: bool = False,
    colour_map: dict[str, str] | None = None,
    biogenic_colour: str = "#2B0FC9",
    height_px: int = 900,
    bar_width: float = 0.66,
    em_col: str = "kgco2e",
) -> go.Figure:
    scale = 1.0 / float(gia_m2) if use_intensity and gia_m2 > 0 else 1.0
    y_unit = "kgCO₂e/m² GIA" if scale != 1.0 else "kgCO₂e"

    def stack_data(modules: list[str]) -> tuple[list[dict], pd.DataFrame, list[str]]:
        rr = filter_rows_for_modules(rows, modules)
        value_col = pick_chart_value_col(rr, em_col)
        agg = _aggregate_for_plot(rr, value_col, scale)
        segments, high_order, _ = _build_segments(agg, colour_map)
        return segments, agg, high_order

    upfront_segments, upfront_agg, upfront_high_order = stack_data(upfront_modules)
    wlc_segments, wlc_agg, wlc_high_order = stack_data(whole_life_modules)
    upfront_total = sum(s["value"] for s in upfront_segments)
    wlc_total = sum(s["value"] for s in wlc_segments)

    fig = go.Figure()
    x_upfront, x_wlc = 0.0, 1.25
    actual_width = 0.58 * bar_width

    def add_stack(x_pos: float, segments: list[dict]):
        for seg in segments:
            text = None if collapsed_high_level else f"{seg['seg']} {seg['value']:,.0f}"
            fig.add_trace(
                go.Bar(
                    x=[x_pos],
                    y=[seg["value"]],
                    width=actual_width,
                    marker=dict(color=seg["colour"], line=dict(color="white", width=1)),
                    showlegend=False,
                    text=[text] if text else None,
                    textposition="inside",
                    insidetextanchor="middle",
                    hovertemplate=f"<b>{seg['high']}</b><br>{seg['seg']}: %{{y:,.0f}} {y_unit}<extra></extra>",
                )
            )

    def add_high_labels(x_pos: float, agg: pd.DataFrame, high_order: list[str]):
        if not collapsed_high_level:
            return
        running = 0.0
        for high in high_order:
            high_value = float(agg.loc[agg["chart_high"] == high, "value"].sum())
            if high_value <= 1e-9:
                continue
            mid = running + high_value / 2.0
            running += high_value
            colour = colour_map.get(high, get_default_rics_colour(high)) if colour_map else get_default_rics_colour(high)
            fig.add_annotation(
                x=x_pos,
                y=mid,
                text=f"{high} {high_value:,.0f}",
                showarrow=False,
                font=dict(color="white" if luminance_from_hex(colour) < 0.5 else "black", size=11),
            )

    add_stack(x_upfront, upfront_segments)
    add_stack(x_wlc, wlc_segments)
    add_high_labels(x_upfront, upfront_agg, upfront_high_order)
    add_high_labels(x_wlc, wlc_agg, wlc_high_order)

    for modules, x_pos in [(upfront_modules, x_upfront), (whole_life_modules, x_wlc)]:
        bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
        if bio_total > 1e-9:
            fig.add_trace(
                go.Bar(
                    x=[x_pos],
                    y=[-abs(bio_total)],
                    width=actual_width,
                    marker=dict(color=biogenic_colour),
                    showlegend=False,
                    hovertemplate=f"Biogenic: %{{y:,.0f}} {y_unit}<extra></extra>",
                )
            )

    for label, x_pos in [
        (f"Upfront<br>{upfront_total:,.0f} {y_unit}", x_upfront),
        (f"Whole life cycle<br>{wlc_total:,.0f} {y_unit}", x_wlc),
    ]:
        fig.add_annotation(
            x=x_pos,
            y=-max(upfront_total, wlc_total, 1.0) * 0.08,
            text=label,
            showarrow=False,
            font=dict(size=14),
        )

    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        yaxis_title=y_unit,
        showlegend=False,
        xaxis=dict(tickvals=[x_upfront, x_wlc], ticktext=["Upfront", "Whole life cycle"], range=[-0.75, 2.0]),
        margin=dict(l=40, r=40, t=70, b=80),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig
