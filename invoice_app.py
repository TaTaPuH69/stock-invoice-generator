#!/usr/bin/env python
# coding: utf-8
"""
Invoice Builder GUI
© 2025 — Stock Invoice Generator (o3 edition)
"""

# ──────────────────────────── imports ────────────────────────────
from __future__ import annotations

import os
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional

import pandas as pd
from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END

# ──────────────────────── глобальная настройка ───────────────────
logging.basicConfig(
    filename="app.log",
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s: %(message)s",
    encoding="utf-8",
)

VAT_RATE = 0.20
CATALOG_PATH = Path("profiles_catalog.xlsx")
_catalog = pd.read_excel(CATALOG_PATH)

# ──────────────────────────── helpers ────────────────────────────
def read_table(path: str) -> pd.DataFrame:
    _, ext = os.path.splitext(path)
    if ext.lower() in (".xls", ".xlsx"):
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str, sep=";")

    # запятая → точка
    df = df.applymap(
        lambda x: str(x).replace(",", ".") if isinstance(x, str) else x
    )
    return df.replace({"": pd.NA}).dropna(how="all")


def find_analog(code: str, length: float) -> Optional[str]:
    fam = _catalog.loc[_catalog["code"] == code, "family"]
    if fam.empty:
        return None
    fam = fam.iat[0]
    subset = _catalog.query("family == @fam and length_m == @length")
    if subset.empty:
        return None
    return subset.sort_values("price_rub").iloc[0]["code"]


# ───────────────────────── StockManager ──────────────────────────
@dataclass
class StockManager:
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    stock_column: str = "Остаток"

    # ── service ───────────────────────────────────────────────────
def _detect_stock_column(self) -> Optional[str]:
    """
    Возвращает имя колонки с количеством на складе.
    Принимаются варианты:
        • остаток / остатки
        • saldo / остаток_колво
        • сальдо ... количество
    """
    for col in self.df.columns:
        name = col.strip().lower()

        if (
            "остаток" in name or
            ("сальдо" in name and "колич" in name) or
            ("debet" in name and "колич" in name) or
            ("quantity" in name and "end" in name)          # на всякий случай
        ):
            return col
    return None


    # ── public API ────────────────────────────────────────────────
    def load(self, path: str) -> None:
        self.df = read_table(path)
        col = self._detect_stock_column()
        if not col:
            raise ValueError(
                "Не найдена колонка с остатками – ожидается столбец «Остаток»"
            )
        self.stock_column = col
        self.df[self.stock_column] = self.df[self.stock_column].astype(float)
        self.df["Цена"] = self.df["Цена"].astype(float)

        dups = self.df[self.df.duplicated("Артикул")]
        if not dups.empty:
            logging.warning(f"Дубликаты в остатках: {dups['Артикул'].tolist()}")

        logging.info(f"Загружено позиций на складе: {len(self.df)}")

    def allocate(self, article: str, qty: float) -> Optional[pd.Series]:
        rows = self.df[self.df["Артикул"] == article]
        if not rows.empty:
            row = rows.iloc[0]
            if row[self.stock_column] >= qty:
                self.df.at[row.name, self.stock_column] -= qty
                return row
        return None

    def find_analog(
        self,
        category: str,
        color: str,
        coating: str,
        width: float,
        used: List[str],
    ) -> Optional[pd.Series]:
        cand = self.df[
            (self.df["Категория"] == category)
            & (self.df["Цвет"] == color)
            & (self.df["Покрытие"] == coating)
            & (~self.df["Артикул"].isin(used))
            & (self.df[self.stock_column] > 0)
        ]
        if "Ширина" in cand.columns:
            cand = cand[
                abs(cand["Ширина"].astype(float) - width) <= 10
            ]
        return None if cand.empty else cand.iloc[0]


# ─────────────────────── InvoiceProcessor ────────────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    df: pd.DataFrame = field(default_factory=pd.DataFrame)

    original_sum: float = 0.0
    used_analogs: List[str] = field(default_factory=list)
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)

    # ── загрузка счёта ────────────────────────────────────────────
    def load(self, path: str) -> None:
        self.df = read_table(path)
        self.df["Количество"] = self.df["Количество"].astype(float)
        self.df["Цена"] = self.df["Цена"].astype(float)

        dups = self.df[self.df.duplicated("Артикул")]
        if not dups.empty:
            logging.warning(f"Дубликаты в счёте: {dups['Артикул'].tolist()}")

        self.original_sum = (self.df["Количество"] * self.df["Цена"]).sum()
        logging.info(f"Загружен счёт на {self.original_sum:,.2f} ₽")

    # ── основная логика ───────────────────────────────────────────
    def process(self) -> None:
        self.used_analogs.clear()
        self.result_rows.clear()
        self.log.clear()

        for _, row in self.df.iterrows():
            art = row["Артикул"]
            length_m = row.get("Длина, м", 0)

            analog_code = find_analog(art, length_m)
            art_to_use = analog_code or art
            comment = f"замена на {analog_code}" if analog_code else ""

            qty = row["Количество"]
            price = row["Цена"]

            stock_row = self.stock.allocate(art_to_use, qty)
            self.result_rows.append(
                dict(Артикул=art_to_use, Количество=qty, Цена=price, Замена=comment)
            )
            if stock_row is not None:
                continue  # всё зарезервировали

            analog = self.stock.find_analog(
                row.get("Категория", ""),
                row.get("Цвет", ""),
                row.get("Покрытие", ""),
                row.get("Ширина", 0),
                self.used_analogs,
            )
            if analog is None or analog[self.stock.stock_column] < qty:
                msg = f"Не удалось найти {art} в нужном количестве"
                self.log.append(msg)
                logging.error(msg)
                continue

            # списываем аналог
            self.stock.df.at[analog.name, self.stock.stock_column] -= qty
            self.used_analogs.append(art)

            last = self.result_rows[-1]
            last["Артикул"] = analog["Артикул"]
            last["Замена"] = f"замена на {analog['Артикул']}"
            self.log.append(f"{art} заменён на {analog['Артикул']}")

    # ── вывод ─────────────────────────────────────────────────────
    def to_dataframe(self) -> pd.DataFrame:
        df = pd.DataFrame(self.result_rows)
        df["Сумма"] = df["Количество"] * df["Цена"]
        df["НДС"] = df["Сумма"] - df["Сумма"] / (1 + VAT_RATE)
        return df

    def save(self, path: str) -> None:
        df = self.to_dataframe()
        total = df["Сумма"].sum()
        vat = df["НДС"].sum()
        df.loc[len(df.index)] = ["Итого", "", "", total, vat]
        df.to_excel(path, index=False)
        logging.info(f"Счёт сохранён в {path}")


# ───────────────────────────── GUI ───────────────────────────────
class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("Invoice Builder")

        self.log_text = Text(self.root, height=20, width=90, font=("Consolas", 10))
        self.log_text.pack(side="left", fill="both", expand=True)
        Scrollbar(self.root, command=self.log_text.yview).pack(
            side="right", fill="y"
        )
        self.log_text.configure(yscrollcommand=self.log_text.yview)

        Button(self.root, text="Загрузить остатки", command=self.load_stock).pack()
        Button(self.root, text="Загрузить счёт", command=self.load_invoice).pack()
        Button(self.root, text="Собрать счёт", command=self.build_invoice).pack()

        self.stock = StockManager()
        self.processor = InvoiceProcessor(stock=self.stock)
        self.stock_file: Optional[str] = None
        self.invoice_file: Optional[str] = None

    # ── helpers ──────────────────────────────────────────────────
    def gui_log(self, msg: str) -> None:
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        logging.info(msg)

    # ── callbacks ────────────────────────────────────────────────
    def load_stock(self) -> None:
        path = filedialog.askopenfilename()
        if not path:
            return
        try:
            self.stock.load(path)
            self.stock_file = path
            self.gui_log(f"Остатки загружены: {len(self.stock.df)} строк")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def load_invoice(self) -> None:
        path = filedialog.askopenfilename()
        if not path:
            return
        try:
            self.processor.load(path)
            self.invoice_file = path
            self.gui_log(f"Счёт загружен: {len(self.processor.df)} строк")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def build_invoice(self) -> None:
        if self.stock.df.empty or self.processor.df.empty:
            messagebox.showwarning("Внимание", "Загрузите остатки и счёт")
            return

        self.processor.process()
        base, _ = os.path.splitext(os.path.basename(self.invoice_file))
        out_path = f"{base}_processed.xlsx"
        self.processor.save(out_path)

        self.gui_log("\n".join(self.processor.log))
        messagebox.showinfo("Готово", f"Новый счёт сохранён: {out_path}")

    # ── run ──────────────────────────────────────────────────────
    def run(self) -> None:
        self.root.mainloop()


# ────────────────────────── entry point ──────────────────────────
if __name__ == "__main__":
    App().run()
