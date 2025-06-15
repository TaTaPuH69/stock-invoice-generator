# extract_flexy_catalog.py  (положи рядом с invoice_app.py)
from pathlib import Path
import camelot, re, pandas as pd

PDF = Path("Flexy.pdf")                  # имя, которое реально лежит в dir
OUT = Path("profiles_catalog.xlsx")      # чтобы Excel лёг в корень

# карта «№ стра­ни­цы → ка­те­го­рия»  (можешь переименовать под себя)
CAT = {
  1: "карнизные", 2: "световые", 3: "теневые", 4: "парящие+контур",
  5: "izi", 6: "бесщелевые+нишевые", 7: "классические",
  8: "трековые", 9: "многоуровневые", 10: "закладные",
  11: "другие", 12: "экраны", 13: "заглушки", 14: "комплектующие"
}

rows, rx = [], re.compile(r"\(ПФ\s*([0-9\-\/]+)\)")
for p in range(1, 15):
    tables = camelot.read_pdf(str(PDF), pages=str(p), flavor="stream")
    for t in tables:
        df = t.df.replace("", pd.NA).dropna(how="all", axis=1)
               # --- ищем заголовок профиля ------------------------
        for i, cell in enumerate(df.iloc[:, 0]):
            if pd.isna(cell):
                continue
            m = rx.search(str(cell))
            if not m:
                continue
            code = m.group(1)
            name = cell.split("(")[0].strip()
            fam  = name.split()[0]

            # сканируем 3-4 нижних строки на цвет / длину / цену
            blk = df.iloc[i+1:i+5].fillna(method="ffill", axis=1)

            for _, r in blk.iterrows():
                raw_color = r.iloc[0]
                raw_color = "" if pd.isna(raw_color) else str(raw_color)
                color = raw_color.split()[0].lower() if raw_color else "unknown"

                txt = " ".join(r.fillna("").astype(str))
                lens   = re.findall(r"(\d+[.,]?\d*)\s*м", txt)
                price_m = re.search(r"(\d+)\s*руб", txt)
                if not lens or not price_m:
                    continue
                price = int(price_m.group(1))

                for ln in lens:
                    rows.append({
                        "code": code,
                        "name": name,
                        "family": fam,
                        "category": CAT.get(p, "прочее"),
                        "color": color,
                        "length_m": float(ln.replace(",", ".")),
                        "price_rub": price
                    })
pd.DataFrame(rows).drop_duplicates().to_excel(OUT, index=False)
print(f"✅ catalog saved: {OUT}")
