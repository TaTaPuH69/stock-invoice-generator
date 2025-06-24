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

        # --- ENRICH -------------------------------------------------
        # приводим код к str и мержим с _catalog
        self.df["Артикул"] = self.df["Артикул"].astype(str).str.strip()
        cat = _catalog.copy()
        cat["code"] = cat["code"].astype(str).str.strip()

        enrich = (
            cat[["code", "family", "length_m", "color", "price_rub"]]
            .rename(
                columns={
                    "code": "Артикул",
                    "family": "Семейство",
                    "length_m": "Длина, м",
                    "color": "Цвет",
                }
            )
        )
        self.df = self.df.merge(enrich, on="Артикул", how="left")

        # устраняем возможные дубли «Семейство_x / Семейство_y» и т.п.
        norm2orig = {}
        for col in list(self.df.columns):
            key = _normalize(col)
            if key in norm2orig:
                primary = norm2orig[key]
                self.df[primary] = self.df[primary].fillna(self.df[col])
                self.df.drop(columns=[col], inplace=True)
            else:
                norm2orig[key] = col

        # гарантируем обязательные поля
        for col in ["Семейство", "Длина, м", "Цвет", "price_rub"]:
            if col not in self.df.columns:
                self.df[col] = pd.NA

        self.stock_column = "Остаток"
        logging.info(f"Загружено {len(self.df)} строк остатков")

        for col in ["Категория", "Цвет", "Покрытие", "Ширина"]:
            if col not in self.df.columns:
                self.df[col] = pd.NA

    def allocate(self, article: str, qty: float) -> Optional[pd.Series]:
        rows = self.df[self.df["Артикул"] == article]
        if not rows.empty:
            row = rows.iloc[0]
            if row[self.stock_column] >= qty:
                self.df.at[row.name, self.stock_column] -= qty
                return row
        return None

    def allocate_partial(self, article: str, qty: float) -> float:
        """Списывает доступное количество и возвращает остаток."""
        rows = self.df[self.df["Артикул"] == article]
        if rows.empty:
            return qty
        row = rows.iloc[0]
        avail = row[self.stock_column]
        take = min(avail, qty)
        if take > 0:
            self.df.at[row.name, self.stock_column] -= take
        return qty - take

    def find_analog(
        self,
        family: str,
        length: float,
        color: str,
        used: list[str],
        target_price: float,
    ) -> Optional[pd.Series]:
        """
        Возвращает строку-аналог:
        • то же семейство
        • длина ±0.05 м
        • приоритет точно совпавшему цвету
        • сортировка по минимальной разнице цены
        Если ничего нет — None.
        """
        df = self.df

        fam_col = next((c for c in df.columns if _normalize(c) == "семейство"), None)
        family_mask = df[fam_col] == family if fam_col else pd.Series(True, index=df.index)

        # базовый фильтр
        mask = (
            family_mask
            & (df[self.stock_column] > 0)
            & (~df["Артикул"].isin(used))
        )
        if "Длина, м" in df.columns:
            mask &= (df["Длина, м"].astype(float) - length).abs() <= 0.05

        cand = df[mask].copy()
        if cand.empty:
            return None

        # совпадающий цвет, если задан
        if color and "Цвет" in cand.columns:
            same = cand[cand["Цвет"] == color]
            cand = same if not same.empty else cand

        # ближе всех по цене
        if pd.notna(target_price) and "price_rub" in cand.columns:
            cand["__diff__"] = (cand["price_rub"] - target_price).abs()
            cand = cand.sort_values("__diff__")

        return cand.iloc[0]


# ─────────────────────── InvoiceProcessor ────────────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    df: pd.DataFrame = field(default_factory=pd.DataFrame)

    original_sum: float = 0.0
    used_analogs: List[str] = field(default_factory=list)
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None
    output_columns: List[str] = field(default_factory=list)

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

        self.output_columns = list(df.columns)
        if "Комментарий" not in self.output_columns:
            self.output_columns.append("Комментарий")
        self.invoice_path = path
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
            need = row["Количество"]
            art = row["Артикул"]
            # --- LOOKUP ---------------------------------------------------
            if art in _catalog["code"].values:
                cat_row = _catalog[_catalog["code"] == art].iloc[0]
                family = row.get("Семейство") or cat_row["family"]
                length = row.get("Длина, м") or cat_row["length_m"]
                color = row.get("Цвет") or cat_row["color"]
                price = row.get("Цена", pd.NA)
                if pd.isna(price):
                    price = cat_row["price_rub"]
            else:
                family = row.get("Семейство", "")
                length = row.get("Длина, м", 0.0)
                color = row.get("Цвет", "")
                price = row.get("Цена", pd.NA)
            # ----------------------------------------------------------------

            left = self.stock.allocate_partial(art, need)
            shipped = need - left

            if shipped:
                base = {c: row.get(c, "") for c in self.output_columns}
                base["Количество"] = int(shipped) if shipped.is_integer() else shipped
                base.setdefault("Комментарий", "")
                self.result_rows.append(base)

            if left == 0:
                continue

            analog = self.stock.find_analog(
                family=family,
                length=length,
                color=color,
                used=self.used_analogs,
                target_price=price,
            )
            if analog is None or analog[self.stock.stock_column] < left:
                self.log.append(f"Не хватило {art}; аналогов нет")
                continue

            self.stock.allocate_partial(analog["Артикул"], left)
            self.used_analogs.append(analog["Артикул"])

            add = {c: "" for c in self.output_columns}
            add.update({
                "Артикул": analog["Артикул"],
                "Количество": int(left) if left.is_integer() else left,
                "Цена": round(analog["price_rub"], 2),
                "Комментарий": f"аналог для {art}",
            })
            self.result_rows.append(add)
            self.log.append(f"{art}: {left} шт → {analog['Артикул']}")

    # ── вывод ─────────────────────────────────────────────────────
    def to_dataframe(self) -> pd.DataFrame:
        df = pd.DataFrame(self.result_rows)
        if "Комментарий" not in df.columns:
            df["Комментарий"] = ""
        for col in ["Количество", "Цена"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
        df["Сумма"] = (df["Количество"] * df["Цена"]).round(2)
        df["НДС"] = (df["Сумма"] - df["Сумма"] / (1 + VAT_RATE)).round(2)
        return df

    def save(self, path: str) -> None:
        """
        Берём оригинальный счёт (self.invoice_path), дописываем новые строки
        result_rows в том же формате и порядке колонок, сохраняем в path.
        """
        base_df = pd.read_excel(
            self.invoice_path,
            skiprows=_find_header_row(self.invoice_path),
            header=0,
            dtype=str,
        )
        if "Комментарий" not in base_df.columns:
            base_df["Комментарий"] = ""
        for r in self.result_rows[len(self.df):]:
            new = {c: "" for c in base_df.columns}
            for k, v in r.items():
                new[k] = v
            base_df.loc[len(base_df)] = new
        base_df.to_excel(path, index=False)
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
