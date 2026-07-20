from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

from oneclick_core.nrm import extract_nrm_code


def load_mapping_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "priority" in df.columns:
        df["priority"] = pd.to_numeric(df["priority"], errors="coerce").fillna(999).astype(int)
        df = df.sort_values("priority", ascending=True)
    return df


def _match_rule(text: str, nrm_code: str, material_family: str, match_type: str, pattern: str) -> bool:
    pattern = str(pattern or "").strip()
    match_type = str(match_type or "").strip().lower()
    text_l = str(text or "").lower()
    nrm_code = str(nrm_code or "")
    material_family_l = str(material_family or "").lower()

    if match_type == "default":
        return pattern in {"", "*"}
    if match_type == "contains":
        return pattern.lower() in text_l
    if match_type == "startswith":
        return text_l.startswith(pattern.lower())
    if match_type == "regex":
        return bool(re.search(pattern, text, flags=re.IGNORECASE))
    if match_type == "nrm_prefix":
        return nrm_code.startswith(pattern.rstrip(".")) and (
            len(nrm_code) == len(pattern.rstrip(".")) or nrm_code[len(pattern.rstrip("."))] == "."
        )
    if match_type == "nrm_code":
        return nrm_code == pattern
    if match_type == "material_family":
        return material_family_l == pattern.lower()
    return False


def apply_rule_column(
    df: pd.DataFrame,
    rules: pd.DataFrame,
    target_col: str,
    value_col: str,
    text_col: str = "match_text",
    nrm_col: str = "nrm_code",
    family_col: str = "material_family",
    default: str = "Other",
) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="string")

    if text_col not in df.columns:
        df = df.copy()
        parts = []
        for col in ["material_label", "Resource", "element_name", "rics_detail"]:
            if col in df.columns:
                parts.append(df[col].astype(str).fillna(""))
        if parts:
            df[text_col] = parts[0]
            for part in parts[1:]:
                df[text_col] = df[text_col] + " " + part
        else:
            df[text_col] = ""

    result = pd.Series(default, index=df.index, dtype="string")
    if rules.empty or value_col not in rules.columns:
        return result

    for _, rule in rules.iterrows():
        mask = pd.Series(False, index=df.index)
        for idx in df.index:
            mask.at[idx] = _match_rule(
                text=df.at[idx, text_col],
                nrm_code=df.at[idx, nrm_col] if nrm_col in df.columns else "",
                material_family=result.at[idx],
                match_type=rule.get("match_type", ""),
                pattern=rule.get("match_pattern", ""),
            )
        matched = mask & result.eq(default)
        result.loc[matched] = str(rule[value_col]).strip()

    return result


def load_label_defaults(path: Path) -> dict[str, str]:
    df = load_mapping_csv(path)
    if df.empty:
        return {}
    if "nrm_code" not in df.columns or "display_name" not in df.columns:
        return {}
    return {
        str(row.nrm_code).strip(): str(row.display_name).strip()
        for row in df.itertuples(index=False)
        if str(row.nrm_code).strip() and str(row.display_name).strip()
    }
