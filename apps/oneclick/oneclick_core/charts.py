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


def filter_rows_for_modules(rows: pd.DataFrame, modules: list[str]) -> pd.DataFrame:
    """Filter to chart modules. Biogenic (bioC) is drawn as a separate bar, not stacked here."""
    return rows[rows["section"].isin(expand_modules(modules))].copy()


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
        value_col = pick_chart_value_col(bio_rows, em_col)
        if "_source_row_id" in bio_rows.columns:
            total = (
                bio_rows.groupby("_source_row_id", dropna=False)[value_col]
                .first()
                .pipe(pd.to_numeric, errors="coerce")
                .fillna(0.0)
                .sum()
            )
        else:
            total = pd.to_numeric(bio_rows[value_col], errors="coerce").fillna(0.0).sum()
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


def _total_label_y(positive_total: float, bio_value: float) -> float:
    """Place stack totals below the lowest bar (biogenic or zero)."""
    floor = min(0.0, -abs(bio_value) if abs(bio_value) > 1e-9 else 0.0)
    span = max(abs(positive_total), abs(floor), 1.0)
    return floor - span * 0.14


def _add_biogenic_bar(
    fig: go.Figure,
    *,
    x,
    bio_total: float,
    width: float,
    biogenic_colour: str,
    y_unit: str,
    collapsed_high_level: bool = False,
) -> None:
    if bio_total <= 1e-9:
        return
    bio_y = -abs(bio_total)
    label = None if collapsed_high_level else f"Biogenic {bio_y:,.0f}"
    fig.add_trace(
        go.Bar(
            x=[x],
            y=[bio_y],
            width=width,
            marker=dict(color=biogenic_colour, line=dict(color="white", width=1)),
            showlegend=False,
            text=[label] if label else None,
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(color="white" if luminance_from_hex(biogenic_colour) < 0.55 else "black"),
            hovertemplate=f"Biogenic: %{{y:,.0f}} {y_unit}<extra></extra>",
        )
    )
    if collapsed_high_level:
        fig.add_annotation(
            x=x if not isinstance(x, str) or x != "" else 0.5,
            xref="x" if not isinstance(x, str) or x != "" else "paper",
            y=bio_y / 2.0,
            text=f"Biogenic {bio_y:,.0f}",
            showarrow=False,
            font=dict(color="white" if luminance_from_hex(biogenic_colour) < 0.55 else "black", size=12),
        )


def aggregate_chart_data(rows: pd.DataFrame, modules: list[str], em_col: str = "kgco2e") -> pd.DataFrame:
    rr = filter_rows_for_modules(rows, modules)
    value_col = pick_chart_value_col(rr, em_col)
    if rr.empty:
        agg = pd.DataFrame(columns=["chart_high", "chart_segment_label", "chart_segment_code", "value"])
    else:
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
        agg = (
            values.merge(labels, on=["chart_high", "chart_segment_code"], how="left")
            .sort_values(["chart_high", "chart_segment_code"])
            .reset_index(drop=True)
        )

    # Include biogenic as its own export/preview row (drawn as a separate bar on charts).
    bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=1.0)
    if bio_total > 1e-9:
        bio_row = pd.DataFrame(
            [
                {
                    "chart_high": "Biogenic",
                    "chart_segment_code": "bioC",
                    "chart_segment_label": "Biogenic carbon storage",
                    "value": -abs(bio_total),
                }
            ]
        )
        agg = pd.concat([agg, bio_row], ignore_index=True)
    return agg


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
            if abs(value) <= 1e-9:
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


def compute_contingency_total(
    rr: pd.DataFrame,
    value_col: str,
    scale: float,
    project_contingency_pct: float = 0.0,
    contingency_map: dict[str, float] | None = None,
) -> float:
    if rr.empty:
        return 0.0
    from oneclick_core.nrm import resolve_row_contingency_pct

    contingency_map = contingency_map or {}
    if not contingency_map and not project_contingency_pct:
        return 0.0

    codes = rr["nrm_code"] if "nrm_code" in rr.columns else pd.Series("", index=rr.index)
    values = pd.to_numeric(rr[value_col], errors="coerce").fillna(0.0)
    total = 0.0
    for code, value in zip(codes, values):
        pct = resolve_row_contingency_pct(str(code), project_contingency_pct, contingency_map)
        total += float(value) * (pct / 100.0)
    return total * scale


def _append_contingency_segment(
    segments: list[dict],
    contingency_total: float,
    contingency_colour: str = "#6C6C6C",
) -> list[dict]:
    if contingency_total <= 1e-9:
        return segments
    out = list(segments)
    out.append(
        {
            "high": "Contingency",
            "seg": "Contingency",
            "value": contingency_total,
            "colour": contingency_colour,
        }
    )
    return out


def plot_rics_single_stack(
    rows: pd.DataFrame,
    modules: list[str],
    title: str,
    use_intensity: bool,
    gia_m2: float,
    collapsed_high_level: bool = False,
    colour_map: dict[str, str] | None = None,
    biogenic_colour: str = "#2B0FC9",
    contingency_colour: str = "#6C6C6C",
    height_px: int = 900,
    bar_width: float = 0.66,
    small_segment_threshold: float = 0.03,
    target_line_value: float | None = None,
    target_line_label: str = "",
    project_contingency_pct: float = 0.0,
    contingency_map: dict[str, float] | None = None,
    em_col: str = "kgco2e",
) -> go.Figure:
    rr = filter_rows_for_modules(rows, modules)
    value_col = pick_chart_value_col(rr, em_col)
    scale = 1.0 / float(gia_m2) if use_intensity and gia_m2 > 0 else 1.0
    y_unit = "kgCO₂e/m² GIA" if scale != 1.0 else "kgCO₂e"

    agg = _aggregate_for_plot(rr, value_col, scale)
    segments, high_order, _ = _build_segments(agg, colour_map)
    contingency_total = compute_contingency_total(
        rr, value_col, scale, project_contingency_pct, contingency_map
    )
    segments = _append_contingency_segment(segments, contingency_total, contingency_colour)
    if contingency_total > 1e-9 and "Contingency" not in high_order:
        high_order = list(high_order) + ["Contingency"]

    # Positive share for small-segment threshold (ignore negative slices).
    total_pos = sum(s["value"] for s in segments if s["value"] > 0) or 1.0
    stack_total = sum(s["value"] for s in segments)
    fig = go.Figure()

    for seg in segments:
        share = abs(seg["value"]) / total_pos
        is_small = share < small_segment_threshold and seg["high"] != "Contingency"
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
        label_rows = list(high_order)
        for high in label_rows:
            if high == "Contingency":
                high_value = contingency_total
                colour = contingency_colour
            else:
                high_value = float(agg.loc[agg["chart_high"] == high, "value"].sum())
                colour = colour_map.get(high, get_default_rics_colour(high)) if colour_map else get_default_rics_colour(high)
            if abs(high_value) <= 1e-9:
                continue
            mid = running + high_value / 2.0
            running += high_value
            fig.add_annotation(
                x=0.5,
                xref="paper",
                y=mid,
                text=f"{high} {high_value:,.0f}",
                showarrow=False,
                font=dict(color="white" if luminance_from_hex(colour) < 0.5 else "black", size=12),
            )

    bio_drawn = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
    _add_biogenic_bar(
        fig,
        x="",
        bio_total=bio_drawn,
        width=bar_width,
        biogenic_colour=biogenic_colour,
        y_unit=y_unit,
        collapsed_high_level=collapsed_high_level,
    )

    fig.add_annotation(
        x=0.5,
        xref="paper",
        y=_total_label_y(stack_total, bio_drawn),
        text=f"{stack_total:,.0f} {y_unit}",
        showarrow=False,
        font=dict(size=14),
    )

    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        yaxis_title=y_unit,
        showlegend=False,
        margin=dict(l=40, r=40, t=70, b=100),
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
    contingency_colour: str = "#6C6C6C",
    height_px: int = 900,
    bar_width: float = 0.66,
    project_contingency_pct: float = 0.0,
    contingency_map: dict[str, float] | None = None,
    em_col: str = "kgco2e",
) -> go.Figure:
    scale = 1.0 / float(gia_m2) if use_intensity and gia_m2 > 0 else 1.0
    y_unit = "kgCO₂e/m² GIA" if scale != 1.0 else "kgCO₂e"

    def stack_data(modules: list[str]) -> tuple[list[dict], pd.DataFrame, list[str], float]:
        rr = filter_rows_for_modules(rows, modules)
        value_col = pick_chart_value_col(rr, em_col)
        agg = _aggregate_for_plot(rr, value_col, scale)
        segments, high_order, _ = _build_segments(agg, colour_map)
        contingency_total = compute_contingency_total(
            rr, value_col, scale, project_contingency_pct, contingency_map
        )
        segments = _append_contingency_segment(segments, contingency_total, contingency_colour)
        if contingency_total > 1e-9 and "Contingency" not in high_order:
            high_order = list(high_order) + ["Contingency"]
        return segments, agg, high_order, contingency_total

    upfront_segments, upfront_agg, upfront_high_order, upfront_cont = stack_data(upfront_modules)
    wlc_segments, wlc_agg, wlc_high_order, wlc_cont = stack_data(whole_life_modules)
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

    def add_high_labels(x_pos: float, agg: pd.DataFrame, high_order: list[str], contingency_total: float):
        if not collapsed_high_level:
            return
        running = 0.0
        for high in high_order:
            if high == "Contingency":
                high_value = contingency_total
                colour = contingency_colour
            else:
                high_value = float(agg.loc[agg["chart_high"] == high, "value"].sum())
                colour = colour_map.get(high, get_default_rics_colour(high)) if colour_map else get_default_rics_colour(high)
            if abs(high_value) <= 1e-9:
                continue
            mid = running + high_value / 2.0
            running += high_value
            fig.add_annotation(
                x=x_pos,
                y=mid,
                text=f"{high} {high_value:,.0f}",
                showarrow=False,
                font=dict(color="white" if luminance_from_hex(colour) < 0.5 else "black", size=11),
            )

    add_stack(x_upfront, upfront_segments)
    add_stack(x_wlc, wlc_segments)
    add_high_labels(x_upfront, upfront_agg, upfront_high_order, upfront_cont)
    add_high_labels(x_wlc, wlc_agg, wlc_high_order, wlc_cont)

    bio_by_x: dict[float, float] = {x_upfront: 0.0, x_wlc: 0.0}
    for modules, x_pos in [
        (upfront_modules, x_upfront),
        (whole_life_modules, x_wlc),
    ]:
        bio_total = compute_biogenic_total(rows, modules=modules, em_col=em_col, scale=scale)
        bio_by_x[x_pos] = bio_total
        _add_biogenic_bar(
            fig,
            x=x_pos,
            bio_total=bio_total,
            width=actual_width,
            biogenic_colour=biogenic_colour,
            y_unit=y_unit,
            collapsed_high_level=collapsed_high_level,
        )

    for label, x_pos, stack_total in [
        (f"Upfront<br>{upfront_total:,.0f} {y_unit}", x_upfront, upfront_total),
        (f"Whole life cycle<br>{wlc_total:,.0f} {y_unit}", x_wlc, wlc_total),
    ]:
        fig.add_annotation(
            x=x_pos,
            y=_total_label_y(stack_total, bio_by_x.get(x_pos, 0.0)),
            text=label,
            showarrow=False,
            font=dict(size=14),
            yanchor="top",
        )

    lowest = min(
        _total_label_y(upfront_total, bio_by_x[x_upfront]),
        _total_label_y(wlc_total, bio_by_x[x_wlc]),
    )
    fig.update_layout(
        barmode="relative",
        height=height_px,
        title=title,
        yaxis_title=y_unit,
        showlegend=False,
        xaxis=dict(
            tickvals=[x_upfront, x_wlc],
            ticktext=["Upfront", "Whole life cycle"],
            range=[-0.75, 2.0],
        ),
        yaxis=dict(range=[lowest * 1.35, None]),
        margin=dict(l=40, r=40, t=70, b=110),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig
