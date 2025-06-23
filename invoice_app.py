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
import re
import unicodedata
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional

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

# ─── util: нормализуем имя колонки ───
def _normalize(col: str) -> str:
    """
    Приводит заголовок столбца к унифицированному виду:
    • lower()         – без регистра
    • удаляем пробелы, «-», табы и переводы строк
    • ё → е
    """
    return (
        str(col)
        .lower()
        .replace("\n", "")      # NEW: убираем перевод строки
        .replace("\r", "")      #         "
        .replace("\t", "")      #         "
        .replace(" ", "")       # было
        .replace("-", "")       # было
        .replace("ё", "е")      # было
    )

# ── поиск строки с заголовками в счёте ───────────────────────────
def _find_header_row(path: str, max_row: int = 40) -> int:
    """Возвращает индекс строки с заголовками таблицы счёта."""
    for i in range(max_row):
        row = pd.read_excel(path, skiprows=i, nrows=1, header=None).fillna("")
        cells = [_normalize(str(c)) for c in row.values.ravel()]
        has_code = any(c.startswith(("код", "артикул")) for c in cells)
        has_qty = any(c.startswith("количест") or c.startswith("колво") or c.startswith("qty") for c in cells)
        if has_code and has_qty:
            return i
    raise ValueError("Header row not found")

# ─── настройка «жёстких» координат ───
FIXED_STOCK_ROW = 9   # B10 → 10-я строка  ➜  index 9
FIXED_STOCK_COL = 1   # B  → второй столбец ➜  index 1
# ──────────────────────────────────────


# ─── read_table (берём 2-й столбец с 10-й строки) ───
def read_table(path: str) -> pd.DataFrame:
    """
    Читает Excel / CSV-файл и возвращает DataFrame
    ▸ Excel: пропускаем первые 9 строк (0-based => строка 10),
      берём все данные без заголовка.
    ▸ CSV: то же самое (skiprows=9, без header).
    Оставляем два столбца: первый (артикул / наименование)
    и второй — количество (переименуем в 'Остаток').
    """
    _, ext = os.path.splitext(path)

    kw_args = dict(dtype=str, header=None, skiprows=9)

    if ext.lower() in (".xls", ".xlsx"):
        df = pd.read_excel(path, **kw_args)
    else:
        df = pd.read_csv(path, sep=";", **kw_args)

    # оставляем только первые два столбца
    df = df.iloc[:, :2]
    df.columns = ["Артикул", "Остаток"]      # как угодно, главное второй - количество
    df["Остаток"] = df["Остаток"].astype(float)

    # заменяем запятую на точку в числах и убираем пустые строки
    df.replace({",": "."}, regex=True, inplace=True)
    df.dropna(how="all", inplace=True)

    return df
# ─────────────────────────────────────────────────────



# ─── StockManager.load (оставляем как есть) ───
# в self.stock_column у вас уже будет строка "Остаток",
# потому что read_table переименовал нужный столбец.



# ---------- StockManager._detect_stock_column ----------
def _detect_stock_column(self) -> str | None:
    """Возвращает название колонки, содержащей остатки/кол-во."""
    kw = {"остаток", "остатки", "колво", "количество", "qty"}

    for col in self.df.columns:
        name = _normalize(col)
        if any(k in name for k in kw):
            return col          # нашли подходящий столбец

    return None                 # ничего не подошло


# ---------- вспомогательная ----------
def _norm_cell(text: str) -> str:
    """
    • приводит строку к NFC-форме (убирает скрытые акценты в кириллице)
    • удаляет все символы категории «Zs» (прочие пробелы) и «Cc» (управляющие)
    • убирает дефисы, подчёркивания, точки.
    """
    text = unicodedata.normalize("NFC", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) not in {"Zs", "Cc"})
    text = re.sub(r"[-_.\s]", "", text)   # ещё раз на всякий
    return text.lower()



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

    # ────────────────────────────────────────────────────────────
    def load(self, path: str) -> None:
        """Загружает остатки без поиска заголовков.

        Значения берутся из столбца B начиная с десятой строки.
        """
        raw = pd.read_excel(path, header=None)
        qty = raw.iloc[FIXED_STOCK_ROW:, FIXED_STOCK_COL]
        articles = raw.iloc[FIXED_STOCK_ROW:, 0]

        self.df = pd.DataFrame({"Артикул": articles, "Остаток": qty})
        self.df.dropna(how="all", inplace=True)
        self.df.reset_index(drop=True, inplace=True)

        self.stock_column = "Остаток"
        logging.info(f"Загружено {len(self.df)} строк остатков")

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
        """Загружает счёт, автоматически определяя строку заголовка."""
        hdr = _find_header_row(path)
        df = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)

        rename_map: dict[str, str] = {}
        for col in df.columns:
            norm = _normalize(col)
            if norm.startswith(("код", "артикул")):
                rename_map[col] = "Артикул"
            elif norm.startswith(("количест", "колво", "qty")):
                rename_map[col] = "Количество"
            elif norm.startswith(("цена", "стоимость", "price")):
                rename_map[col] = "Цена"

        df.rename(columns=rename_map, inplace=True)

        if "Цена" not in df.columns:
            df["Цена"] = pd.NA

        df = df.loc[:, [c for c in ["Артикул", "Количество", "Цена"] if c in df.columns]]
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        df.dropna(how="all", inplace=True)

        df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce")
        df["Цена"] = pd.to_numeric(df["Цена"], errors="coerce")
        df.dropna(subset=["Количество"], inplace=True)

        self.df = df
        # ↓↓↓ дальнейший (старый) код оставляем без изменений ↓↓↓

        dups = self.df[self.df.duplicated("Артикул")]
        if not dups.empty:
            logging.warning(f"Дубликаты в счёте: {dups['Артикул'].tolist()}")

        if self.df["Цена"].notna().any():
            self.original_sum = (
                self.df["Количество"] * self.df["Цена"]
            ).sum()
            logging.info(
                f"Загружен счёт на {self.original_sum:,.2f} ₽"
            )
        else:
            self.original_sum = 0.0
            logging.info("Загружен счёт без цен")

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
        scroll_bar = Scrollbar(self.root, command=self.log_text.yview)
        scroll_bar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll_bar.set)
        self.log_text.pack(side="left", fill="both", expand=True)

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
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

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
