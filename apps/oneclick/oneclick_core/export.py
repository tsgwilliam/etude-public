from __future__ import annotations

from io import BytesIO

import pandas as pd

from oneclick_core.charts import aggregate_chart_data, DEFAULT_WLC_MODULES, UPFRONT_MODULES


def build_results_workbook_bytes(
    rows: pd.DataFrame,
    project_name: str,
    nrm_level: int,
    total_gia: float,
    selected_buildings: list[str],
) -> bytes:
    upfront = aggregate_chart_data(rows, UPFRONT_MODULES)
    wlc = aggregate_chart_data(rows, DEFAULT_WLC_MODULES)
    combined = pd.concat(
        [
            upfront.assign(chart_set="Upfront"),
            wlc.assign(chart_set="Whole_life"),
        ],
        ignore_index=True,
    )

    summary = pd.DataFrame(
        [
            {"metric": "Project", "value": project_name},
            {"metric": "Buildings", "value": ", ".join(selected_buildings)},
            {"metric": "Combined GIA (m²)", "value": total_gia},
            {"metric": "NRM display level", "value": nrm_level},
            {"metric": "Row count", "value": len(rows)},
        ]
    )

    unmapped = rows[
        (rows["nrm_code"].astype(str).str.strip() == "")
        | (rows.get("material_family", pd.Series("", index=rows.index)).astype(str) == "Other")
    ].copy()

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        rows.to_excel(writer, sheet_name="Tidy_Rows", index=False)
        upfront.to_excel(writer, sheet_name="Chart1_Upfront", index=False)
        wlc.to_excel(writer, sheet_name="Chart2_WLC", index=False)
        combined.to_excel(writer, sheet_name="Chart3_Combined", index=False)
        if not unmapped.empty:
            unmapped.to_excel(writer, sheet_name="Unmapped", index=False)
    return buffer.getvalue()
