import pandas as pd
import logging
import os
from dataclasses import dataclass, field
from typing import List, Optional
from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END
import pandas as pd
import logging
import os
# (остальные ваши импорты)

# ─────── КАТАЛОГ ───────
from pathlib import Path
import pandas as pd          # ← если уже выше, эту строку можно убрать

CATALOG_PATH = Path("profiles_catalog.xlsx")
_catalog = pd.read_excel(CATALOG_PATH)

def find_analog(code: str, length: float) -> str | None:
    fam = _catalog.loc[_catalog["code"] == code, "family"]
    if fam.empty:
        return None
    fam = fam.iat[0]
    subset = _catalog.query("family == @fam and length_m == @length")
    if subset.empty:
        return None
    return subset.sort_values("price_rub").iloc[0]["code"]
# ───────────────────────

logging.basicConfig(filename="app.log", level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")

VAT_RATE = 0.2

def read_table(path: str) -> pd.DataFrame:
    _, ext = os.path.splitext(path)
    if ext.lower() in [".xls", ".xlsx"]:
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str, sep=";")
    df = df.applymap(lambda x: str(x).replace(",", ".") if isinstance(x, str) else x)
    df = df.replace({"": pd.NA}).dropna(how="all")
    return df

@dataclass
class StockManager:
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    stock_column: str = "Остаток"

    def _detect_stock_column(self) -> Optional[str]:
        """Return the column name that contains stock values."""
        for col in self.df.columns:
            name = col.strip().lower()
            if "остаток" in name or "остатки" in name:
                return col
        return None

    def load(self, path: str):
        self.df = read_table(path)
        col = self._detect_stock_column()
        if not col:
            cols = ", ".join(self.df.columns)
            raise ValueError(
                f"Не найдена колонка с остатками. Ожидались названия вроде 'Остаток'. Найдены: {cols}"
            )
        self.stock_column = col
        self.df[self.stock_column] = self.df[self.stock_column].astype(float)
        self.df["Цена"] = self.df["Цена"].astype(float)
        duplicates = self.df[self.df.duplicated("Артикул")]
        if not duplicates.empty:
            logging.warning(f"Дубликаты в остатках: {duplicates['Артикул'].tolist()}")
        logging.info(f"Загружено {len(self.df)} позиций остатков")

    def allocate(self, article: str, qty: float) -> Optional[pd.Series]:
        rows = self.df[self.df["Артикул"] == article]
        if not rows.empty:
            row = rows.iloc[0]
            if row[self.stock_column] >= qty:
                idx = row.name
                self.df.at[idx, self.stock_column] -= qty
                return row
        return None

    def find_analog(self, category: str, color: str, coating: str, width: float, used: List[str]) -> Optional[pd.Series]:
        candidates = self.df[
            (self.df["Категория"] == category) &
            (self.df["Цвет"] == color) &
            (self.df["Покрытие"] == coating) &
            (self.df["Артикул"].isin(used) == False) &
            (self.df[self.stock_column] > 0)
        ]
        if "Ширина" in candidates.columns:
            candidates = candidates[abs(candidates["Ширина"].astype(float) - width) <= 10]
        if not candidates.empty:
            return candidates.iloc[0]
        return None

@dataclass
class InvoiceProcessor:
    stock: StockManager
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    original_sum: float = 0.0
    used_analogs: List[str] = field(default_factory=list)
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)

    def load(self, path: str):
        self.df = read_table(path)
        self.df["Количество"] = self.df["Количество"].astype(float)
        self.df["Цена"] = self.df["Цена"].astype(float)
        duplicates = self.df[self.df.duplicated("Артикул")]
        if not duplicates.empty:
            logging.warning(f"Дубликаты в счете: {duplicates['Артикул'].tolist()}")
        self.original_sum = (self.df["Количество"] * self.df["Цена"]).sum()
        logging.info(f"Загружен счет на сумму {self.original_sum:.2f}")

    def process(self):
        # --------------------------------------------------
        #  перебираем строки счёта и резервируем позиции
        # --------------------------------------------------
        self.used_analogs: list[str] = []          # список уже-использованных

        for _, row in self.df.iterrows():          # ← 8 пробелов
            art      = row["Артикул"]
            length_m = row.get("Длина, м", 0)

            # 1) пробуем найти артикул-аналог по каталогу
            analog_code = find_analog(art, length_m)

            if analog_code:                        # ← 12 пробелов
                art_to_use = analog_code
                comment    = f"замена на {analog_code}"
            else:
                art_to_use = art
                comment    = ""

            qty   = row["Количество"]
            price = row["Цена"]

            # 2) резервируем выбранный артикул
            stock_row = self.stock.allocate(art_to_use, qty)

            # 3) добавляем строку в результат
            self.result_rows.append({
                "Артикул":    art_to_use,
                "Количество": qty,
                "Цена":       price,
                "Замена":     comment,
            })

            # 4) если товара нет — ищем *физический* аналог
            if stock_row is None:                  # ← 12 пробелов
                analog = self.stock.find_analog(   # 16 пробелов
                    row.get("Категория",  ""),     # категория
                    row.get("Цвет",       ""),     # цвет
                    row.get("Покрытие",   ""),     # покрытие
                    row.get("Ширина",     0),      # ширина/длина
                    self.used_analogs,             # уже использованные
                )

                if analog is not None and analog[self.stock.stock_column] >= qty:
                    idx = analog.name
                    # списываем остаток
                    self.stock.df.at[idx, self.stock.stock_column] -= qty
                    self.used_analogs.append(art)

                    # правим ПОСЛЕДНЮЮ добавленную строку
                    self.result_rows[-1]["Артикул"] = analog["Артикул"]
                    self.result_rows[-1]["Замена"]  = f"замена на {analog['Артикул']}"
                    continue                        # ← 20 пробелов (внутри if analog)

            # --- поиск аналога НЕ понадобился ---
            analog = self.stock.find_analog(
                row.get("Категория",  ""),
                row.get("Цвет",       ""),
                row.get("Покрытие",   ""),
                row.get("Ширина",     0),
                self.used_analogs,
            )
            if analog is not None and analog[self.stock.stock_column] >= qty:
                idx = analog.name
                self.stock.df.at[idx, self.stock.stock_column] -= qty
                self.used_analogs.append(analog["Артикул"])
                self.result_rows.append({
                    "Артикул":    analog["Артикул"],
                    "Количество": qty,
                    "Цена":       analog["Цена"],
                    "Замена":     art,             # что заменили
                })
                self.log.append(f"{art} заменен на {analog['Артикул']}")
                self.log.append(f"Не удалось найти {art} в нужном количестве")
                logging.error(f"Не удалось найти {art} в нужном количестве")
                continue                            # ← 16 пробелов (внутри for-цикла)

    def to_dataframe(self) -> pd.DataFrame:
        df = pd.DataFrame(self.result_rows)
        df["Сумма"] = df["Количество"] * df["Цена"]
        df["НДС"] = df["Сумма"] - df["Сумма"] / (1 + VAT_RATE)
        return df

    def save(self, path: str):
        df = self.to_dataframe()
        total = df["Сумма"].sum()
        vat = df["НДС"].sum()
        df.loc[len(df.index)] = ["Итого", "", "", total, vat]
        df.to_excel(path, index=False)
        logging.info(f"Счет сохранен в {path}")

class App:
    def __init__(self):
        self.root = Tk()
        self.root.title("Invoice Builder")
        self.log_text = Text(self.root, height=20, width=80)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = Scrollbar(self.root, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

        Button(self.root, text="Загрузить остатки", command=self.load_stock).pack()
        Button(self.root, text="Загрузить счет", command=self.load_invoice).pack()
        Button(self.root, text="Собрать счет", command=self.build_invoice).pack()

        self.stock = StockManager()
        self.processor = InvoiceProcessor(self.stock)
        self.stock_file = None
        self.invoice_file = None

    def log(self, msg: str):
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        logging.info(msg)

    def load_stock(self):
        path = filedialog.askopenfilename()
        if path:
            try:
                self.stock.load(path)
                self.stock_file = path
                self.log(f"Остатки загружены: {len(self.stock.df)} строк")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def load_invoice(self):
        path = filedialog.askopenfilename()
        if path:
            try:
                self.processor.load(path)
                self.invoice_file = path
                self.log(f"Счет загружен: {len(self.processor.df)} строк")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def build_invoice(self):
        if self.stock.df.empty or self.processor.df.empty:
            messagebox.showwarning("Внимание", "Загрузите остатки и счет")
            return
        self.processor.process()
        base, ext = os.path.splitext(os.path.basename(self.invoice_file))
        out_path = f"{base}_processed.xlsx"
        self.processor.save(out_path)
        self.log("\n".join(self.processor.log))
        messagebox.showinfo("Готово", f"Новый счет сохранен: {out_path}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    App().run()
