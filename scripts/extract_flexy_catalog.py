# extract_flexy_catalog.py  (лежит рядом с invoice_app.py)
from pathlib import Path
import warnings, re, camelot, pandas as pd

warnings.simplefilter("ignore", FutureWarning)          # глушим лавину предупреждений

PDF = Path("Flexy.pdf")                # имя каталога-PDF
OUT = Path("profiles_catalog.xlsx")    # куда сохраняем итог

# карта «№ страницы → категория»
CAT = {
    1: "карнизные", 2: "световые", 3: "теневые", 4: "парящие+контур",
    5: "izi", 6: "бесшедевенничество", 7: "классические",
    8: "трековые", 9: "многоуровневые", 10: "закладные",
    11: "другие", 12: "экраны", 13: "заглушки", 14: "комплектующие"
}

rows, rx = [], re.compile(r"\b(\d{4,6})\b")             # ищем 4-6-значные коды

for p in range(1, 15):                                   # 1-14 страницы
    tables = camelot.read_pdf(str(PDF), pages=str(p), flavor="lattice", strip_text=" \n")
    if not tables:                                       # fallback, если нет сетки
        tables = camelot.read_pdf(str(PDF), pages=str(p), flavor="stream",  strip_text=" \n")

    for t in tables:
        df = t.df.replace("", pd.NA).dropna(how="all", axis=1)

        # --- ищем заголовок профиля ---
        for i, cell in enumerate(df.iloc[:, 0]):
            if pd.isna(cell):
                continue
            m = rx.search(str(cell))
            if not m:
                continue

            code  = m.group(1)
            name  = cell.split("(", 1)[0].strip()
            fam   = name.split()[0]

            # сканируем 3-4 нижних строки на цвет / длину / цену
            blk = df.iloc[i+1:i+5].fillna(method="ffill", axis=1)

            for _, r in blk.iterrows():
                raw_color = r.iloc[0]
                raw_color = "" if pd.isna(raw_color) else str(raw_color)
                color     = raw_color.split()[0].lower() if raw_color else "unknown"

                txt     = " ".join(r.fillna("").astype(str))
                lens    = re.findall(r"(\d[.,]?\d*)\s*м", txt)
                price_m = re.search(r"(\d+)\s*руб", txt)
                if not lens or not price_m:
                    continue

                price = int(price_m.group(1))
                for ln in lens:
                    rows.append({
                        "code":     code,
                        "name":     name,
                        "family":   fam,
                        "category": CAT.get(p, "прочее"),
                        "color":    color,
                        "length_m": float(ln.replace(",", ".")),
                        "price_rub": price
                    })

pd.DataFrame(rows).drop_duplicates().to_excel(OUT, index=False)
print(f"[OK] catalog saved: {OUT}")
