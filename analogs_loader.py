import os
import re
from typing import Dict, List

import pandas as pd

RULES_DEFAULT_PATH = "analogs_priority.xlsx"

def extract_base_code(s: str) -> str:
    """Return first 4-6 digit sequence from string ``s``."""
    if not s:
        return ""
    m = re.search(r"\d{4,6}", str(s))
    return m.group(0) if m else ""

def load_analog_rules(path: str = RULES_DEFAULT_PATH) -> Dict[str, List[str]]:
    """Load analog replacement rules from Excel or CSV file.

    Returns mapping of base code -> list of alternative base codes.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    _, ext = os.path.splitext(path)
    if ext.lower() in (".xls", ".xlsx"):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)
    cols = ["Товар", "Аналог №1", "Аналог №2", "Аналог №3", "Аналог №4"]
    df = df[[c for c in cols if c in df.columns]].fillna("")
    rules: Dict[str, List[str]] = {}
    for _, row in df.iterrows():
        base = extract_base_code(str(row.get("Товар", "")))
        if not base:
            continue
        alts: List[str] = []
        for col in cols[1:]:
            code = extract_base_code(str(row.get(col, "")))
            if code and code not in alts:
                alts.append(code)
        if not alts:
            continue
        if base not in rules:
            rules[base] = alts
        else:
            for code in alts:
                if code not in rules[base]:
                    rules[base].append(code)
    return rules
