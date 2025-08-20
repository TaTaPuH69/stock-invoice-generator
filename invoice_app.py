#!/usr/bin/env python
# coding: utf-8
"""
Invoice Builder GUI
© 2025 — Stock Invoice Generator (o3 edition)
"""

# ──────────────────────────── imports ────────────────────────────
from __future__ import annotations

import sys, logging
from logging.handlers import RotatingFileHandler
logger = logging.getLogger("invoice")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(name)s: %(message)s")
    _fh = RotatingFileHandler("app.log", maxBytes=1_048_576, backupCount=5, encoding="utf-8")
    _fh.setFormatter(_fmt)
    logger.addHandler(_fh)
    logger.propagate = False

def _attach_cli_stream_logger():
    """Attach stdout stream handler once (for --cli / --show-log)."""
    import logging, sys
    _fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(name)s: %(message)s")
    for h in logger.handlers:
        if isinstance(h, logging.StreamHandler) and getattr(h, "_cli", False):
            return
    sh = logging.StreamHandler(sys.stdout)
    sh._cli = True  # mark
    sh.setFormatter(_fmt)
    logger.addHandler(sh)

import os
import shutil
import re
import unicodedata
import argparse
import pandas as pd

def _is_filled(val) -> bool:
    """True, если val не None, не pd.NA и не пустая строка."""
    return pd.notna(val) and str(val).strip() != ""

from flexy_catalog_loader import load_catalog
from dataclasses import dataclass, field
from typing import List, Optional

from analogs_loader import (
    extract_base_code,
    load_analog_rules,
    RULES_DEFAULT_PATH,
)

try:  # GUI components are optional for CLI mode
    from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END, Toplevel
except Exception:  # noqa: BLE001 - fine for environments without tkinter
    Tk = filedialog = messagebox = Text = Scrollbar = Button = END = Toplevel = None

VAT_RATE = 0.20
_catalog = load_catalog()

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
# ───────────────────────── StockManager ──────────────────────────
@dataclass
class StockManager:
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    stock_column: str = "Остаток"

    # ────────────────────────────────────────────────────────────
    def load(self, path: str) -> None:
        raw = pd.read_excel(path, header=None)
        qty = raw.iloc[FIXED_STOCK_ROW:, FIXED_STOCK_COL]
        articles = raw.iloc[FIXED_STOCK_ROW:, 0]

        self.df = pd.DataFrame({"Артикул": articles, "Остаток": qty})
        self.df.dropna(how="all", inplace=True)
        self.df.reset_index(drop=True, inplace=True)

        self.df["Артикул"] = self.df["Артикул"].astype(str).str.strip()
        self.df["Остаток"] = pd.to_numeric(self.df["Остаток"], errors="coerce").fillna(0)
        self.df["BaseCode"] = self.df["Артикул"].map(extract_base_code)

        if isinstance(_catalog, pd.DataFrame) and {"code", "price_rub"} <= set(_catalog.columns):
            cat = _catalog.copy()
            cat["code"] = cat["code"].astype(str).str.strip()
            enrich = cat[["code", "price_rub"]].rename(columns={"code": "Артикул"})
            enrich["price_rub"] = pd.to_numeric(enrich["price_rub"], errors="coerce")
            self.df = self.df.merge(enrich, on="Артикул", how="left")
        else:
            self.df["price_rub"] = pd.NA

        self.stock_column = "Остаток"
        logger.info(f"Загружено {len(self.df)} строк остатков")

    def allocate(self, article: str, qty: float) -> Optional[pd.Series]:
        """Reserve ``qty`` items of ``article`` if available."""

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

    def allocate_by_basecode(self, base_code: str, qty: float) -> list[tuple[pd.Series, float]]:
        """
        Списывает qty по позициям, где BaseCode == base_code.
        Цвет/длина игнорируются. Берём строки с наибольшим остатком.
        Возвращает список (row, taken). Остаток уменьшает в self.df.
        """
        df = self.df
        if "BaseCode" not in df.columns:
            return []
        pool = df[(df["BaseCode"] == base_code) & (df[self.stock_column] > 0)].copy()
        if pool.empty:
            return []
        pool = pool.sort_values(self.stock_column, ascending=False)
        left = float(qty)
        out: list[tuple[pd.Series, float]] = []
        for idx, row in pool.iterrows():
            if left <= 0:
                break
            avail = float(row[self.stock_column])
            take = min(avail, left)
            if take > 0:
                self.df.at[idx, self.stock_column] = avail - take
                out.append((row, take))
                left -= take
        return out


# ─────────────────────── InvoiceProcessor ────────────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    rules: dict[str, List[str]] = field(default_factory=dict)
    df: pd.DataFrame = field(default_factory=pd.DataFrame)

    original_sum: float = 0.0
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None
    base_columns: List[str] = field(default_factory=list)
    disp_code: str = ""
    disp_qty: str = ""
    disp_price: str = ""
    col_code: str = "Артикул"
    col_qty: str = "Количество"
    col_price: str = "Цена"

    # ── загрузка счёта ────────────────────────────────────────────
    def load(self, path: str) -> None:
        """Загружает счёт, автоматически определяя строку заголовка."""
        hdr = _find_header_row(path)
        base = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)
        self.invoice_path = path
        self.base_columns = list(base.columns)
        self.disp_code = next(
            (c for c in self.base_columns if _normalize(c).startswith(("код", "артикул"))),
            self.base_columns[0] if self.base_columns else "",
        )
        self.disp_qty = next(
            (c for c in self.base_columns if _normalize(c).startswith("кол")),
            self.base_columns[1] if len(self.base_columns) > 1 else self.base_columns[0],
        )
        self.disp_price = next(
            (
                c
                for c in self.base_columns
                if _normalize(c).startswith(("цена", "стоимость", "price"))
            ),
            "",
        )
        rename_map = {self.disp_code: self.col_code, self.disp_qty: self.col_qty}
        if self.disp_price:
            rename_map[self.disp_price] = self.col_price
        df = base.rename(columns=rename_map)
        df = df.loc[:, [c for c in [self.col_code, self.col_qty, self.col_price] if c in df.columns]]
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        df.dropna(how="all", inplace=True)
        df[self.col_qty] = pd.to_numeric(df[self.col_qty], errors="coerce")
        if self.col_price in df.columns:
            df[self.col_price] = pd.to_numeric(df.get(self.col_price), errors="coerce")
        df.dropna(subset=[self.col_qty], inplace=True)
        self.df = df
        dups = self.df[self.df.duplicated(self.col_code)]
        if not dups.empty:
            logger.warning(f"Дубликаты в счёте: {dups[self.col_code].tolist()}")
        if self.col_price in self.df.columns and self.df[self.col_price].notna().any():
            self.original_sum = (self.df[self.col_qty] * self.df[self.col_price]).sum()
            logger.info(f"Загружен счёт на {self.original_sum:,.2f} ₽")
        else:
            self.original_sum = 0.0
            logger.info("Загружен счёт без цен")

    # ── основная логика ───────────────────────────────────────────
    def process(self) -> None:
        """Process invoice rows using available stock and catalog."""

        self.result_rows.clear()
        self.log.clear()

        # --- VALIDATE INPUT -------------------------------------------------
        required_cols = {"Артикул", "Количество"}
        missing = required_cols - set(self.df.columns)
        if missing:
            msg = f"В счёте нет колонок: {', '.join(missing)}"
            self.log.append(msg)
            logger.error(msg)
            return
        # --------------------------------------------------------------------

        for _, row in self.df.iterrows():
            art = row[self.col_code]
            need = row[self.col_qty]
            price = row.get(self.col_price, pd.NA)

            cat_row = _catalog[_catalog["code"] == art]
            if not cat_row.empty and pd.isna(price):
                price = cat_row.iloc[0]["price_rub"]

            left = self.stock.allocate_partial(art, need)
            shipped = need - left

            if shipped:
                row_out = {c: "" for c in self.base_columns}
                row_out[self.disp_code] = art
                row_out[self.disp_qty] = shipped
                if self.disp_price and pd.notna(price):
                    row_out[self.disp_price] = price
                self.result_rows.append(row_out)

            if left > 0:
                base = extract_base_code(art)
                cand_bases = self.rules.get(base, [])
                if cand_bases:
                    for base_cand in cand_bases:
                        if left <= 0:
                            break
                        allocs = self.stock.allocate_by_basecode(base_cand, left)
                        for row2, taken in allocs:
                            add = {c: "" for c in self.base_columns}
                            add[self.disp_code] = row2["Артикул"]
                            add[self.disp_qty] = taken
                            if self.disp_price:
                                pr = pd.to_numeric(row2.get("price_rub"), errors="coerce")
                                if pd.notna(pr):
                                    add[self.disp_price] = float(pr)
                                elif pd.notna(price):
                                    add[self.disp_price] = float(price)
                            add["Комментарий"] = "по правилам"
                            add["_is_analog"] = True
                            self.result_rows.append(add)
                            left -= taken
                            self.log.append(
                                f"{art}: {taken} шт → {row2['Артикул']} (по правилам)"
                            )
                    if left > 0:
                        self.log.append(
                            f"{art}: не хватило {left} шт — аналоги по правилам закончились"
                        )
                else:
                    self.log.append(f"{art}: аналогов нет по правилам")

        if not self.result_rows:
            msg = "аналогов не найдено, счёт не изменён"
            self.log.append(msg)
            logger.info(msg)

    # ── вывод ─────────────────────────────────────────────────────
    def to_dataframe(self) -> pd.DataFrame:
        cols = self.base_columns.copy()
        if "Комментарий" not in cols:
            cols.append("Комментарий")
        df = pd.DataFrame(self.result_rows)
        return df.reindex(columns=cols)

    def save(self, path: str) -> None:
        """Save processed invoice to ``path``.

        Новый файл создаётся всегда. Колонка «Комментарий» добавляется,
        если её не было в исходном счёте.
        """

        base = pd.read_excel(
            self.invoice_path,
            skiprows=_find_header_row(self.invoice_path),
            header=0,
            dtype=str,
        )

        if "Комментарий" not in base.columns:
            base["Комментарий"] = ""

        analog_rows = [r for r in self.result_rows if r.get("_is_analog")]
        add = [{c: r.get(c, "") for c in base.columns} for r in analog_rows]
        if add:
            base = pd.concat([base, pd.DataFrame(add)], ignore_index=True)

        base.to_excel(path, index=False)
        logger.info(f"Счёт сохранён в {path}")



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
        Button(self.root, text="Загрузить правила", command=self.load_rules).pack()
        Button(self.root, text="Посмотреть логи", command=self.view_logs).pack()
        Button(self.root, text="Скачать логи", command=self.save_logs).pack()

        self.stock = StockManager()
        import os
        self.rules: dict[str, List[str]] = {}
        try:
            if os.path.exists(RULES_DEFAULT_PATH):
                self.rules = load_analog_rules(RULES_DEFAULT_PATH)
                self.gui_log(f"Правила загружены: {len(self.rules)} баз (авто)")
        except Exception:  # noqa: BLE001
            logger.exception("Автозагрузка правил не удалась")

        self.processor = InvoiceProcessor(stock=self.stock, rules=self.rules)
        self.stock_file: Optional[str] = None
        self.invoice_file: Optional[str] = None

    # ── helpers ──────────────────────────────────────────────────
    def gui_log(self, msg: str) -> None:
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        logger.info(msg)

    def view_logs(self) -> None:
        top = Toplevel(self.root)
        top.title("Логи приложения")
        txt = Text(top, width=100, height=40, font=("Consolas", 10))
        txt.pack(fill="both", expand=True)
        try:
            with open("app.log", "r", encoding="utf-8") as f:
                txt.insert("end", f.read())
        except FileNotFoundError:
            txt.insert("end", "Лог-файл не найден")
        txt.config(state="disabled")

    def save_logs(self) -> None:
        dst = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("All files", "*.*")],
        )
        if dst:
            shutil.copyfile("app.log", dst)
            messagebox.showinfo("Логи сохранены", f"Файл сохранён: {dst}")

    # ── callbacks ────────────────────────────────────────────────
    def load_stock(self) -> None:
        path = filedialog.askopenfilename()
        if not path:
            return
        try:
            self.stock.load(path)
            self.stock_file = path
            self.gui_log(f"Остатки загружены: {len(self.stock.df)} строк")
        except Exception as exc:  # noqa: BLE001
            logger.exception("Ошибка при загрузке остатков")
            messagebox.showerror("Ошибка", f"{type(exc).__name__}: {exc}")

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

    def load_rules(self) -> None:
        path = filedialog.askopenfilename()
        if not path:
            return
        try:
            self.rules = load_analog_rules(path)
            self.gui_log(f"Правила загружены: {len(self.rules)} баз")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Ошибка", str(exc))
        self.processor.rules = self.rules

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
    parser = argparse.ArgumentParser(description="Invoice Builder")
    parser.add_argument("--cli", nargs=2, metavar=("STOCK", "INVOICE"), help="run without GUI")
    parser.add_argument("--show-log", action="store_true", help="print last 100 log lines")
    parser.add_argument("--rules", help="path to analog rules file")
    args = parser.parse_args()

    if args.cli or args.show_log:
        _attach_cli_stream_logger()

    if args.show_log:
        if os.path.exists("app.log"):
            with open("app.log", "r", encoding="utf-8") as f:
                lines = f.readlines()[-100:]
            for line in lines:
                print(line, end="")
        sys.exit(0)

    if args.cli:
        stock_path, invoice_path = args.cli
        stock = StockManager()
        rules_path = args.rules or RULES_DEFAULT_PATH
        try:
            rules = load_analog_rules(rules_path)
            logger.info(f"Правила загружены: {len(rules)} баз")
        except FileNotFoundError:
            rules = {}
            logger.info("Правила не найдены")
        proc = InvoiceProcessor(stock=stock, rules=rules)
        stock.load(stock_path)
        proc.load(invoice_path)
        proc.process()
        base, _ = os.path.splitext(os.path.basename(invoice_path))
        out_path = f"{base}_processed.xlsx"
        proc.save(out_path)
        for line in proc.log:
            print(line)
    else:
        app = App()
        if args.rules:
            try:
                app.rules = load_analog_rules(args.rules)
                logger.info(f"Правила загружены: {len(app.rules)} баз")
            except FileNotFoundError:
                logger.info("Правила не найдены")
            app.processor.rules = app.rules
        app.run()
