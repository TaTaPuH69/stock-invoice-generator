#!/usr/bin/env python
# coding: utf-8
"""
Invoice Builder GUI — fixed visual + full-line replacement by rules
"""

from __future__ import annotations

import os, sys, re, unicodedata, argparse, logging, shutil
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field
from typing import List, Optional, Dict

import pandas as pd
from flexy_catalog_loader import load_catalog

# ───────────────────────── logging ─────────────────────────
logger = logging.getLogger("invoice")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(name)s: %(message)s")
    fh = RotatingFileHandler("app.log", maxBytes=1_048_576, backupCount=5, encoding="utf-8")
    fh.setFormatter(_fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(_fmt)
    logger.addHandler(fh)
    logger.addHandler(sh)
    logger.propagate = False

# ──────────────────────── constants ───────────────────────
VAT_RATE = 0.20
RULES_DEFAULT_PATH = "analogs_priority.xlsx"
CATALOG: pd.DataFrame = load_catalog()

# ──────────────────────── tk (optional) ───────────────────
try:
    from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END, Toplevel
except Exception:
    Tk = filedialog = messagebox = Text = Scrollbar = Button = END = Toplevel = None  # type: ignore

# ────────────────────── helpers / normalize ───────────────
def _normalize(s: str) -> str:
    return (
        str(s).lower()
        .replace("\n","").replace("\r","").replace("\t","")
        .replace(" ", "").replace("-", "").replace("ё","е")
    )

def _find_header_row(path: str, max_row: int = 40) -> int:
    for i in range(max_row):
        row = pd.read_excel(path, skiprows=i, nrows=1, header=None).fillna("")
        cells = [_normalize(str(c)) for c in row.values.ravel()]
        has_code = any(c.startswith(("код", "артикул")) for c in cells)
        has_qty  = any(c.startswith("кол") or c.startswith("qty") for c in cells)
        if has_code and has_qty:
            return i
    raise ValueError("Header row not found")

def extract_base_code(s: str) -> str:
    m = re.search(r"\d{4,6}", str(s))
    return m.group(0) if m else ""

def load_analog_rules(path: str = RULES_DEFAULT_PATH) -> Dict[str, List[str]]:
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")
    def norm(x: str) -> str: return str(x).strip().lower().replace(" ", "")
    base_col = None
    for c in df.columns:
        if norm(c) in {"товар","база","код","артикул","base"}:
            base_col = c; break
    if base_col is None:
        base_col = df.columns[0]
    analog_cols = [c for c in df.columns if ("аналог" in norm(c) or "analog" in norm(c))]
    if not analog_cols:
        analog_cols = [c for c in df.columns if c != base_col]

    rules: Dict[str, List[str]] = {}
    for _, r in df.iterrows():
        base = extract_base_code(r.get(base_col, ""))
        if not base:
            continue
        lst: List[str] = []
        for c in analog_cols:
            cand = extract_base_code(r.get(c, ""))
            if cand and cand != base and cand not in lst:
                lst.append(cand)
        if lst:
            rules[base] = lst
    return rules

# имена для «Товар»
def _load_name_map() -> Dict[str, str]:
    try:
        df = pd.read_excel("profiles_catalog.xlsx", dtype=str).fillna("")
    except Exception:
        return {}
    def norm(x: str) -> str: return _normalize(x)
    code_col = next((c for c in df.columns if norm(c).startswith(("код","артикул","sku","code"))), None)
    name_col = next((c for c in df.columns if norm(c).startswith(("товар","наимен","product","опис"))), None)
    if not code_col or not name_col:
        return {}
    m = dict(
        zip(df[code_col].astype(str).str.strip(), df[name_col].astype(str).str.strip())
    )
    return m

NAME_MAP: Dict[str, str] = _load_name_map()

def name_for(code: str) -> str:
    return NAME_MAP.get(str(code), str(code))

# ────────────────────── stock manager ─────────────────────
FIXED_STOCK_ROW = 9
FIXED_STOCK_COL = 1

@dataclass
class StockManager:
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    stock_column: str = "Остаток"

    def load(self, path: str) -> None:
        raw = pd.read_excel(path, header=None)
        qty = raw.iloc[FIXED_STOCK_ROW:, FIXED_STOCK_COL]
        articles = raw.iloc[FIXED_STOCK_ROW:, 0]
        df = pd.DataFrame({"Артикул": articles, "Остаток": qty})
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)

        df["Артикул"] = df["Артикул"].astype(str).str.strip()
        df["Остаток"] = pd.to_numeric(df["Остаток"], errors="coerce").fillna(0.0)
        df["BaseCode"] = df["Артикул"].map(extract_base_code)

        # цены из каталога, если есть
        if {"code", "price_rub"} <= set(CATALOG.columns):
            cat = CATALOG[["code","price_rub"]].copy()
            cat["code"] = cat["code"].astype(str).str.strip()
            cat["price_rub"] = pd.to_numeric(cat["price_rub"], errors="coerce")
            df = df.merge(cat.rename(columns={"code":"Артикул"}), on="Артикул", how="left")

        self.df = df
        self.stock_column = "Остаток"
        logger.info(f"Загружено {len(self.df)} строк остатков")

    # доступно сейчас
    def available(self, article: str) -> float:
        rows = self.df[self.df["Артикул"] == article]
        if rows.empty:
            return 0.0
        return float(rows.iloc[0][self.stock_column])

    # уменьшить остаток (если хватает)
    def consume_exact(self, article: str, qty: float) -> bool:
        rows = self.df[self.df["Артикул"] == article]
        if rows.empty:
            return False
        idx = rows.index[0]
        avail = float(rows.iloc[0][self.stock_column])
        if avail >= qty:
            self.df.at[idx, self.stock_column] = avail - qty
            return True
        return False

    # раздать по всем артикулам с нужным base_code (по убыванию остатков)
    def allocate_by_basecode(self, base_code: str, qty: float) -> List[tuple[pd.Series, float]]:
        if not base_code:
            return []
        pool = self.df[(self.df["BaseCode"] == base_code) & (self.df[self.stock_column] > 0.0)].copy()
        if pool.empty:
            return []
        pool = pool.sort_values(self.stock_column, ascending=False)
        left = float(qty)
        out: List[tuple[pd.Series, float]] = []
        for idx, r in pool.iterrows():
            if left <= 0:
                break
            avail = float(r[self.stock_column])
            take = min(avail, left)
            if take > 0:
                self.df.at[idx, self.stock_column] = avail - take
                out.append((r, take))
                left -= take
        return out

# ───────────────────── invoice processor ─────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    df: pd.DataFrame = field(default_factory=pd.DataFrame)

    # исходный шаблон и имена колонок
    base_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    output_columns: List[str] = field(default_factory=list)
    col_code: Optional[str] = None
    col_name: Optional[str] = None
    col_qty: Optional[str] = None
    col_price: Optional[str] = None
    col_total: Optional[str] = None
    col_vat_incl: Optional[str] = None
    col_unit: Optional[str] = None

    # результаты и служебные поля
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None

    rules: Dict[str, List[str]] = field(default_factory=dict)

    def _pick_col(self, cols: List[str], *prefixes: str) -> Optional[str]:
        for c in cols:
            n = _normalize(c)
            if any(n.startswith(p) for p in prefixes):
                return c
        return None

    # загрузка исходного счёта (и фиксация визуала)
    def load(self, path: str) -> None:
        hdr = _find_header_row(path)
        base = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)
        self.invoice_path = path
        self.base_df = base.copy()
        self.output_columns = list(base.columns)

        # определяем ключевые столбцы по заголовкам
        cols = list(base.columns)
        self.col_name  = self._pick_col(cols, "товар", "наимен", "product")
        self.col_code  = self._pick_col(cols, "код", "артикул", "sku", "code")
        self.col_qty   = self._pick_col(cols, "кол", "qty")
        self.col_price = self._pick_col(cols, "цена", "стоим", "price")
        self.col_total = self._pick_col(cols, "всего", "сумма", "итог")
        self.col_vat_incl = self._pick_col(cols, "втчндс", "втомчислендс")
        self.col_unit  = self._pick_col(cols, "ед", "единиц")

        # соберём рабочую df с минимальными полями
        work = base.copy()
        # если нет "Цена" в файле — добавим пустую (чтобы не падать)
        if self.col_price and self.col_price not in work.columns:
            work[self.col_price] = pd.NA
        self.df = work[[c for c in [self.col_code, self.col_qty, self.col_price] if c]].copy()
        # приведение типов
        if self.col_qty:
            self.df[self.col_qty] = pd.to_numeric(self.df[self.col_qty], errors="coerce")
        if self.col_price:
            self.df[self.col_price] = pd.to_numeric(self.df[self.col_price], errors="coerce")

        logger.info(f"Счёт загружен: {len(self.df)} строк")

    # собрать строку по шаблону исходной строки
    def _clone_row(self, base_row: pd.Series) -> dict:
        out = {c: base_row.get(c, "") for c in self.output_columns}
        return out

    # сформировать строку для аналога
    def _make_analog_row(self, base_row: pd.Series, art: str, qty: float, price: Optional[float]) -> dict:
        row = self._clone_row(base_row)
        if self.col_name:
            row[self.col_name] = name_for(art)
        if self.col_code:
            row[self.col_code] = art
        if self.col_qty is not None:
            row[self.col_qty] = qty
        # цена: приоритет — переданная, потом по каталогу, потом исходная
        pr = None
        if price is not None and pd.notna(price):
            pr = float(price)
        elif {"code","price_rub"} <= set(CATALOG.columns):
            m = CATALOG[CATALOG["code"].astype(str) == str(art)]
            if not m.empty:
                pr = float(pd.to_numeric(m.iloc[0]["price_rub"], errors="coerce"))
        if pr is None and self.col_price:
            try:
                pr = float(base_row.get(self.col_price))
            except Exception:
                pr = None
        if pr is not None and self.col_price:
            row[self.col_price] = pr
        # итоги
        if pr is not None and self.col_total and self.col_qty:
            total = round(float(qty) * float(pr), 2)
            row[self.col_total] = total
            if self.col_vat_incl:
                row[self.col_vat_incl] = round(total * VAT_RATE, 2)
        return row

    def process(self) -> None:
        self.result_rows.clear()
        self.log.clear()

        if not (self.col_code and self.col_qty):
            self.log.append("Не найдены колонки кода/количества в счёте")
            logger.error(self.log[-1])
            return

        # правила
        # автозагрузка из файла рядом с приложением, если явно не передали
        if not self.rules and os.path.exists(RULES_DEFAULT_PATH):
            try:
                self.rules = load_analog_rules(RULES_DEFAULT_PATH)
                logger.info(f"Правила загружены: {len(self.rules)} баз (авто)")
            except Exception:
                logger.exception("Автозагрузка правил не удалась")

        # идём по исходным строкам (той же последовательности)
        for idx, base_row in self.base_df.iterrows():
            art = str(base_row.get(self.col_code, "")).strip() if self.col_code else ""
            need = float(pd.to_numeric(base_row.get(self.col_qty, 0), errors="coerce") or 0)
            price0 = None
            if self.col_price:
                p = pd.to_numeric(base_row.get(self.col_price), errors="coerce")
                price0 = float(p) if pd.notna(p) else None

            if not art or need <= 0:
                # просто копируем строку как есть (пустая/сервисная)
                self.result_rows.append(self._clone_row(base_row))
                continue

            avail = self.stock.available(art)

            if avail >= need:
                # оригинала хватает — списываем и копируем строку как есть
                self.stock.consume_exact(art, need)
                self.result_rows.append(self._clone_row(base_row))
                continue

            # иначе — ПОЛНАЯ замена строки аналогами по правилам
            base_code = extract_base_code(art)
            cand_bases = self.rules.get(base_code, [])

            if not cand_bases:
                self.log.append(f"{art}: аналогов нет по правилам — строка не закрыта")
                # исходную строку НЕ копируем (согласно ТЗ — при нехватке замещаем, а не смешиваем)
                continue

            left = need
            closed_any = False
            for b in cand_bases:
                if left <= 0:
                    break
                allocs = self.stock.allocate_by_basecode(b, left)
                for r, taken in allocs:
                    analog_code = str(r["Артикул"])
                    analog_price = pd.to_numeric(r.get("price_rub"), errors="coerce")
                    out_row = self._make_analog_row(base_row, analog_code, float(taken),
                                                    float(analog_price) if pd.notna(analog_price) else None)
                    self.result_rows.append(out_row)
                    self.log.append(f"{art}: {taken} шт → {name_for(analog_code)} (по правилам)")
                    left -= float(taken)
                    closed_any = True

            if left > 0:
                self.log.append(f"{art}: не хватило {left} шт — строка не закрыта")
                # строку-оригинал не возвращаем (замена частичная недопустима по ТЗ)

            if not closed_any:
                # совсем ничего не подобралось — строку пропускаем
                pass

    def to_dataframe(self) -> pd.DataFrame:
        if not self.result_rows:
            return pd.DataFrame(columns=self.output_columns)
        df = pd.DataFrame(self.result_rows)
        # выровнять порядок и наличие колонок под исходный шаблон
        df = df.reindex(columns=self.output_columns, fill_value="")
        return df

    def save(self, path: str) -> None:
        out = self.to_dataframe()
        # Пишем ровно таблицу результата, без “верхней” исходной части
        out.to_excel(path, index=False)
        logger.info(f"Счёт сохранён в {path}")

# ───────────────────────────── GUI ─────────────────────────────
class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("Invoice Builder")

        self.log_text = Text(self.root, height=20, width=90, font=("Consolas", 10))
        sb = Scrollbar(self.root, command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)

        Button(self.root, text="Загрузить остатки", command=self.load_stock).pack()
        Button(self.root, text="Загрузить счёт", command=self.load_invoice).pack()
        Button(self.root, text="Собрать счёт", command=self.build_invoice).pack()
        Button(self.root, text="Посмотреть логи", command=self.view_logs).pack()
        Button(self.root, text="Скачать логи", command=self.save_logs).pack()

        self.stock = StockManager()
        self.processor = InvoiceProcessor(stock=self.stock)

        # автоправила, если файл лежит рядом
        if os.path.exists(RULES_DEFAULT_PATH):
            try:
                self.processor.rules = load_analog_rules(RULES_DEFAULT_PATH)
                self.gui_log(f"Правила загружены: {len(self.processor.rules)} баз (авто)")
            except Exception:
                logger.exception("Автозагрузка правил не удалась")

        self.stock_file: Optional[str] = None
        self.invoice_file: Optional[str] = None

    def gui_log(self, msg: str) -> None:
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        logger.info(msg)

    def view_logs(self) -> None:
        top = Toplevel(self.root); top.title("Логи приложения")
        txt = Text(top, width=100, height=40, font=("Consolas", 10)); txt.pack(fill="both", expand=True)
        try:
            with open("app.log", "r", encoding="utf-8") as f:
                txt.insert("end", f.read())
        except FileNotFoundError:
            txt.insert("end", "Лог-файл не найден")
        txt.config(state="disabled")

    def save_logs(self) -> None:
        dst = filedialog.asksaveasfilename(defaultextension=".log",
                                           filetypes=[("Log files", "*.log"), ("All files", "*.*")])
        if dst:
            shutil.copyfile("app.log", dst)
            messagebox.showinfo("Логи сохранены", f"Файл сохранён: {dst}")

    def load_stock(self) -> None:
        path = filedialog.askopenfilename()
        if not path: return
        try:
            self.stock.load(path)
            self.stock_file = path
            self.gui_log(f"Остатки загружены: {len(self.stock.df)} строк")
        except Exception as exc:
            logger.exception("Ошибка при загрузке остатков")
            messagebox.showerror("Ошибка", f"{type(exc).__name__}: {exc}")

    def load_invoice(self) -> None:
        path = filedialog.askopenfilename()
        if not path: return
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
        base_name, _ = os.path.splitext(os.path.basename(self.invoice_file))
        out_path = f"{base_name}_processed.xlsx"
        self.processor.save(out_path)
        if self.processor.log:
            self.gui_log("\n".join(self.processor.log))
        messagebox.showinfo("Готово", f"Новый счёт сохранён: {out_path}")

    def run(self) -> None:
        self.root.mainloop()

# ────────────────────────── entry point ──────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Invoice Builder")
    parser.add_argument("--cli", nargs=2, metavar=("STOCK", "INVOICE"), help="run without GUI")
    parser.add_argument("--rules", help="path to analog rules file")
    args = parser.parse_args()

    if args.cli:
        stock_path, invoice_path = args.cli
        stock = StockManager(); stock.load(stock_path)
        proc = InvoiceProcessor(stock=stock)
        proc.load(invoice_path)
        proc.rules = load_analog_rules(args.rules or RULES_DEFAULT_PATH) if (args.rules or os.path.exists(RULES_DEFAULT_PATH)) else {}
        proc.process()
        base_name, _ = os.path.splitext(os.path.basename(invoice_path))
        out_path = f"{base_name}_processed.xlsx"
        proc.save(out_path)
        for line in proc.log: print(line)
    else:
        app = App()
        if args.rules:
            app.processor.rules = load_analog_rules(args.rules)
        app.run()
