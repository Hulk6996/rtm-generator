"""
parser.py — чтение входного Excel-файла с требованиями.

Ожидаемые листы:
  • Business Requirements  (BR)
  • Functional Requirements (FR)
  • Test Cases             (TC)
"""

import pandas as pd
from pathlib import Path
from config import SHEET_BR, SHEET_FR, SHEET_TC, BR_COLS, FR_COLS, TC_COLS


# ─────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────

def _norm_refs(val) -> list[str]:
    """'FR-001, FR-002' → ['FR-001', 'FR-002']"""
    if pd.isna(val) or str(val).strip() == "":
        return []
    return [r.strip() for r in str(val).split(",") if r.strip()]


def _required_cols(df: pd.DataFrame, cols: dict, sheet: str) -> None:
    missing = [v for v in cols.values() if v not in df.columns]
    if missing:
        raise ValueError(f"[{sheet}] Отсутствуют колонки: {missing}")


# ─────────────────────────────────────────────────────────────
#  Public functions
# ─────────────────────────────────────────────────────────────

def load_requirements(filepath: str) -> dict:
    """
    Читает Excel и возвращает dict:
      {
        'br': list[dict],
        'fr': list[dict],
        'tc': list[dict],
      }
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {filepath}")

    xl = pd.ExcelFile(filepath)

    def read_sheet(name: str) -> pd.DataFrame:
        if name not in xl.sheet_names:
            raise ValueError(
                f"Лист '{name}' не найден. "
                f"Доступные листы: {xl.sheet_names}"
            )
        df = xl.parse(name)
        df.columns = [c.strip() for c in df.columns]
        df = df.dropna(how="all")
        return df

    br_df = read_sheet(SHEET_BR)
    fr_df = read_sheet(SHEET_FR)
    tc_df = read_sheet(SHEET_TC)

    _required_cols(br_df, {k: v for k, v in BR_COLS.items() if k in ("id", "title", "priority", "status")}, SHEET_BR)
    _required_cols(fr_df, {k: v for k, v in FR_COLS.items() if k in ("id", "title", "br_refs", "status")}, SHEET_FR)
    _required_cols(tc_df, {k: v for k, v in TC_COLS.items() if k in ("id", "title", "fr_refs", "result")}, SHEET_TC)

    br_list = _parse_br(br_df)
    fr_list = _parse_fr(fr_df)
    tc_list = _parse_tc(tc_df)

    print(f"  ✔ BR: {len(br_list):3d} | FR: {len(fr_list):3d} | TC: {len(tc_list):3d}")
    return {"br": br_list, "fr": fr_list, "tc": tc_list}


def _parse_br(df: pd.DataFrame) -> list[dict]:
    c = BR_COLS
    records = []
    for _, row in df.iterrows():
        records.append({
            "id":          str(row.get(c["id"], "")).strip(),
            "title":       str(row.get(c["title"], "")).strip(),
            "description": str(row.get(c.get("description", "Description"), "")).strip(),
            "priority":    str(row.get(c.get("priority", "Priority"), "Medium")).strip(),
            "category":    str(row.get(c.get("category", "Category"), "")).strip(),
            "source":      str(row.get(c.get("source", "Source"), "")).strip(),
            "status":      str(row.get(c.get("status", "Status"), "Active")).strip(),
        })
    return [r for r in records if r["id"]]


def _parse_fr(df: pd.DataFrame) -> list[dict]:
    c = FR_COLS
    records = []
    for _, row in df.iterrows():
        records.append({
            "id":          str(row.get(c["id"], "")).strip(),
            "title":       str(row.get(c["title"], "")).strip(),
            "description": str(row.get(c.get("description", "Description"), "")).strip(),
            "br_refs":     _norm_refs(row.get(c["br_refs"], "")),
            "type":        str(row.get(c.get("type", "Type"), "Functional")).strip(),
            "component":   str(row.get(c.get("component", "Component"), "")).strip(),
            "status":      str(row.get(c.get("status", "Status"), "Active")).strip(),
        })
    return [r for r in records if r["id"]]


def _parse_tc(df: pd.DataFrame) -> list[dict]:
    c = TC_COLS
    records = []
    for _, row in df.iterrows():
        records.append({
            "id":       str(row.get(c["id"], "")).strip(),
            "title":    str(row.get(c["title"], "")).strip(),
            "fr_refs":  _norm_refs(row.get(c["fr_refs"], "")),
            "type":     str(row.get(c.get("type", "Type"), "Manual")).strip(),
            "result":   str(row.get(c["result"], "Not Run")).strip(),
            "priority": str(row.get(c.get("priority", "Priority"), "Medium")).strip(),
        })
    return [r for r in records if r["id"]]
