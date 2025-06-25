from pathlib import Path
import pandas as pd

CATALOG_PATH = Path("Flexy.xlsx")
CACHE_PATH = Path(".flexy_cache.parquet")

def load_catalog() -> pd.DataFrame:
    if CACHE_PATH.exists() and CACHE_PATH.stat().st_mtime >= CATALOG_PATH.stat().st_mtime:
        return pd.read_parquet(CACHE_PATH)
    df = pd.read_excel(CATALOG_PATH, dtype=str)
    df = df.rename(columns={
        "Код": "code",
        "Семейство": "family",
        "Длина, м": "length_m",
        "Цвет": "color",
        "Цена": "price_rub",
    })
    df["code"] = df["code"].astype(str).str.strip()
    df["length_m"] = pd.to_numeric(df["length_m"], errors="coerce")
    df["price_rub"] = pd.to_numeric(df["price_rub"], errors="coerce")
    df.to_parquet(CACHE_PATH, index=False)
    return df
