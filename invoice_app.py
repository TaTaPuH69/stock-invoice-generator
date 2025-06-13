import pandas as pd
import logging
import os
from dataclasses import dataclass, field
from typing import List, Optional
from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END

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

    def load(self, path: str):
        self.df = read_table(path)
        self.df["Остаток"] = self.df["Остаток"].astype(float)
        self.df["Цена"] = self.df["Цена"].astype(float)
        duplicates = self.df[self.df.duplicated("Артикул")]
        if not duplicates.empty:
            logging.warning(f"Дубликаты в остатках: {duplicates['Артикул'].tolist()}")
        logging.info(f"Загружено {len(self.df)} позиций остатков")

    def allocate(self, article: str, qty: float) -> Optional[pd.Series]:
        rows = self.df[self.df["Артикул"] == article]
        if not rows.empty:
            row = rows.iloc[0]
            if row["Остаток"] >= qty:
                idx = row.name
                self.df.at[idx, "Остаток"] -= qty
                return row
        return None

    def find_analog(self, category: str, color: str, coating: str, width: float, used: List[str]) -> Optional[pd.Series]:
        candidates = self.df[
            (self.df["Категория"] == category) &
            (self.df["Цвет"] == color) &
            (self.df["Покрытие"] == coating) &
            (self.df["Артикул"].isin(used) == False) &
            (self.df["Остаток"] > 0)
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
        for _, row in self.df.iterrows():
            art = row["Артикул"]
            qty = row["Количество"]
            price = row["Цена"]
            stock_row = self.stock.allocate(art, qty)
            if stock_row is not None:
                self.result_rows.append({
                    "Артикул": art,
                    "Количество": qty,
                    "Цена": price,
                    "Замена": ""
                })
                continue
            # search analog
            analog = self.stock.find_analog(row.get("Категория", ""), row.get("Цвет", ""), row.get("Покрытие", ""), row.get("Ширина", 0), self.used_analogs)
            if analog is not None and analog["Остаток"] >= qty:
                idx = analog.name
                self.stock.df.at[idx, "Остаток"] -= qty
                self.used_analogs.append(analog["Артикул"])
                self.result_rows.append({
                    "Артикул": analog["Артикул"],
                    "Количество": qty,
                    "Цена": analog["Цена"],
                    "Замена": art
                })
                self.log.append(f"{art} заменен на {analog['Артикул']}")
                continue
            self.log.append(f"Не удалось найти {art} в нужном количестве")
            logging.error(f"Не удалось найти {art} в нужном количестве")

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
