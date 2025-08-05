from pathlib import Path
import pandas as pd

CATALOG_PATH = Path("flexy_catalog_clean.xlsx")  # готовый каталог


def load_catalog() -> pd.DataFrame:
    """Load Flexy catalog from Excel.

    Returns
    -------
    pandas.DataFrame
        Columns: ``code``, ``family``, ``length_m``, ``color``, ``price_rub``.
    """

    if not CATALOG_PATH.exists():
        raise FileNotFoundError("flexy_catalog_clean.xlsx not found")

    df = pd.read_excel(CATALOG_PATH, sheet_name=0)

    required = ["code", "family", "length_m", "color", "price_rub"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Flexy catalog missing columns: {missing}")

    df["code"] = df["code"].astype(str).str.strip()
    df["family"] = df["family"].astype(str).str.strip()
    df["color"] = df["color"].astype(str).str.strip()
    df["length_m"] = pd.to_numeric(df["length_m"], errors="coerce").astype(float)
    df["price_rub"] = pd.to_numeric(df["price_rub"], errors="coerce").astype(float)

    return df

################################################################
# Пояснения:
# • Файл ``flexy_catalog_clean.xlsx`` уже приведён к нужному формату,
#   поэтому дополнительная очистка не требуется.
################################################################
