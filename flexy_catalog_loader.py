from pathlib import Path
import pandas as pd

CATALOG_PATH = Path("flexy_catalog_clean.xlsx")   # готовый каталог
CACHE_PATH = Path(".flexy_cache.parquet")       # кэш для скорости

def load_catalog() -> pd.DataFrame:
    """Возвращает таблицу code|family|length_m|color|price_rub."""
    # 1. используем кэш, если файл не изменился
    if (
        CACHE_PATH.exists()
        and CATALOG_PATH.exists()
        and CACHE_PATH.stat().st_mtime >= CATALOG_PATH.stat().st_mtime
    ):
        return pd.read_parquet(CACHE_PATH)

    # 2. читаем готовый Excel
    if not CATALOG_PATH.exists():
        raise FileNotFoundError("flexy_catalog_clean.xlsx not found")

    df = pd.read_excel(CATALOG_PATH, dtype=str)

    # 3. валидация обязательных колонок
    required = ["code", "family", "length_m", "color", "price_rub"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Flexy catalog missing columns: {missing}")

    # 4. приведение типов
    df["code"] = df["code"].astype(str).str.strip()
    df["family"] = df["family"].str.strip()
    df["length_m"] = pd.to_numeric(df["length_m"], errors="coerce")
    df["price_rub"] = pd.to_numeric(df["price_rub"], errors="coerce")
    df["color"] = df["color"].str.strip()

    # 5. сохранить кэш и вернуть
    df.to_parquet(CACHE_PATH, index=False)
    return df

################################################################
# Пояснения:
# •  Больше никакого rename/регэкспа – файл уже чистый.
# •  Если Flexy.xlsx изменят > перегенерируют clean-каталог,
#    кэш автоматически сбросится.
################################################################
