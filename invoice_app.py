#!/usr/bin/env python
# coding: utf-8
"""
Invoice Builder GUI
© 2025 — Stock Invoice Generator (o3 edition)
"""

from __future__ import annotations

import os, re, sys, math, shutil, logging, unicodedata, argparse
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from logging.handlers import RotatingFileHandler

import pandas as pd
from flexy_catalog_loader import load_catalog

# ────────────────────────── logging ──────────────────────────────
logger = logging.getLogger("invoice")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(name)s: %(message)s")
    _fh = RotatingFileHandler("app.log", maxBytes=1_048_576, backupCount=5, encoding="utf-8")
    _fh.setFormatter(_fmt)
    _sh = logging.StreamHandler(sys.stdout)
    _sh.setFormatter(_fmt)
    logger.addHandler(_fh)
    logger.addHandler(_sh)
    logger.propagate = False

try:
    from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END, Toplevel
except Exception:
    Tk = filedialog = messagebox = Text = Scrollbar = Button = END = Toplevel = None

# ────────────────────────── constants ────────────────────────────
VAT_RATE = 0.20
RULES_DEFAULT_PATH = "analogs_priority.xlsx"
CATALOG = load_catalog()

# Допуски по сумме и шаг округления метров
AMOUNT_TOL = 0.03        # ±3% от целевой суммы строки
AMOUNT_ABS_TOL = 5.0     # или 5 ₽ абсолют
QTY_STEP = 0.1           # шаг округления по метрам (0.1 м)

# Жёсткий формат складской выгрузки (A — артикул, B — количество, начиная с 10-й строки)
FIXED_STOCK_ROW = 9
FIXED_STOCK_COL = 1

# ───────────────────────── util ──────────────────────────────────
def _normalize(col: str) -> str:
    return (str(col).lower()
            .replace("\n","").replace("\r","").replace("\t","")
            .replace(" ","").replace("-","").replace("ё","е"))

def _find_header_row(path: str, max_row: int = 40) -> int:
    for i in range(max_row):
        row = pd.read_excel(path, skiprows=i, nrows=1, header=None).fillna("")
        cells = [_normalize(str(c)) for c in row.values.ravel()]
        has_code = any(c.startswith(("код","артикул")) for c in cells)
        has_qty  = any(c.startswith("количест") or c.startswith("колво") or c.startswith("qty") for c in cells)
        if has_code and has_qty:
            return i
    raise ValueError("Header row not found")

def extract_base_code(s: str) -> str:
    m = re.search(r"\d{4,6}", str(s))
    return m.group(0) if m else ""

def round_step(x: float, step: float = QTY_STEP) -> float:
    if step <= 0: return x
    return round(round(x/step)*step, 3)

def load_analog_rules(path: str = RULES_DEFAULT_PATH) -> dict[str, list[str]]:
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")

    def norm(x: str) -> str: return str(x).strip().lower().replace(" ","")
    base_candidates = {"товар","база","код","артикул"}
    base_col = next((c for c in df.columns if norm(c) in base_candidates), df.columns[0])

    analog_cols = [c for c in df.columns if ("аналог" in norm(c) or "analog" in norm(c))]
    if not analog_cols:
        analog_cols = [c for c in df.columns if c != base_col]

    rules: dict[str, list[str]] = {}
    for _, row in df.iterrows():
        base = extract_base_code(row.get(base_col, ""))
        if not base: continue
        lst: list[str] = []
        for c in analog_cols:
            cand = extract_base_code(row.get(c, ""))
            if cand and cand != base and cand not in lst:
                lst.append(cand)
        if lst:
            rules[base] = lst
    return rules

def load_analog_prices(path: str = RULES_DEFAULT_PATH) -> dict[str, dict[str, float]]:
    """
    Ищем пары колонок «Аналог N»/«Цена N» (или Analog/Price).
    Возвращаем dict[base_code][analog_base] = price_override (float|None).
    """
    prices: dict[str, dict[str, float]] = {}
    if not os.path.exists(path): return prices
    df = pd.read_excel(path, dtype=str).fillna("")

    def norm(s: str) -> str: return str(s).strip().lower()
    base_candidates = {"товар","база","код","артикул"}
    base_col = next((c for c in df.columns if norm(c).replace(" ","") in base_candidates), df.columns[0])

    analog_cols: list[tuple[str,str]] = []  # (col_name, numkey)
    price_by_num: dict[str,str] = {}
    for c in df.columns:
        cn = norm(c)
        m = re.search(r"(\d+)", cn or "")
        num = m.group(1) if m else ""
        if "аналог" in cn or "analog" in cn:
            analog_cols.append((c, num))
        elif "цена" in cn or "price" in cn:
            price_by_num[num] = c

    for _, r in df.iterrows():
        base = extract_base_code(r.get(base_col, ""))
        if not base: continue
        for col, num in analog_cols:
            ab = extract_base_code(r.get(col, ""))
            if not ab: continue
            price_val = None
            pcol = price_by_num.get(num)
            if pcol:
                pv = str(r.get(pcol,"")).replace(",", ".")
                pv = pd.to_numeric(pv, errors="coerce")
                if pd.notna(pv): price_val = float(pv)
            prices.setdefault(base, {})[ab] = price_val
    return prices

# ─────────────────────── StockManager ────────────────────────────
@dataclass
class StockManager:
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    stock_column: str = "Остаток"

    def load(self, path: str) -> None:
        raw = pd.read_excel(path, header=None)
        qty = raw.iloc[FIXED_STOCK_ROW:, FIXED_STOCK_COL]
        articles = raw.iloc[FIXED_STOCK_ROW:, 0]

        self.df = pd.DataFrame({"Артикул": articles, "Остаток": qty})
        self.df.dropna(how="all", inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.df["Артикул"] = self.df["Артикул"].astype(str).str.strip()
        self.df["Остаток"] = pd.to_numeric(self.df["Остаток"], errors="coerce").fillna(0.0)
        self.df["BaseCode"] = self.df["Артикул"].map(extract_base_code)

        # Подтянем цены из каталога
        if isinstance(CATALOG, pd.DataFrame) and {"code","price_rub"} <= set(CATALOG.columns):
            cat = CATALOG.copy()
            cat["code"] = cat["code"].astype(str).str.strip()
            enrich = cat[["code","price_rub"]].rename(columns={"code":"Артикул"})
            enrich["price_rub"] = pd.to_numeric(enrich["price_rub"], errors="coerce")
            self.df = self.df.merge(enrich, on="Артикул", how="left")

        self.stock_column = "Остаток"
        logger.info(f"Загружено {len(self.df)} строк остатков")

    def available(self, article: str) -> float:
        rows = self.df[self.df["Артикул"] == article]
        if rows.empty: return 0.0
        return float(rows.iloc[0][self.stock_column])

    def allocate_full(self, article: str, qty: float) -> bool:
        """Списать qty полностью, только если хватает. Иначе — ничего не списывать."""
        rows = self.df[self.df["Артикул"] == article]
        if rows.empty: return False
        row = rows.iloc[0]
        avail = float(row[self.stock_column])
        if avail >= float(qty):
            self.df.at[row.name, self.stock_column] = avail - float(qty)
            return True
        return False

# ─────────────────────── InvoiceProcessor ────────────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    app: Optional["App"] = None
    df: pd.DataFrame = field(default_factory=pd.DataFrame)

    original_sum: float = 0.0
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None
    base_columns: List[str] = field(default_factory=list)

    col_code: str = "Артикул"
    col_qty: str  = "Количество"
    col_price: str = "Цена"

    def load(self, path: str) -> None:
        hdr = _find_header_row(path)
        df = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)
        self.invoice_path = path
        self.base_columns = list(df.columns)

        # выровняем ключевые колонки
        rename_map: dict[str,str] = {}
        for col in df.columns:
            norm = _normalize(col)
            if norm.startswith(("код","артикул")):         rename_map[col] = "Артикул"
            elif norm.startswith(("количест","колво","qty")): rename_map[col] = "Количество"
            elif norm.startswith(("цена","стоимость","price")): rename_map[col] = "Цена"
        if rename_map:
            df = df.rename(columns=rename_map)

        if "Цена" not in df.columns:
            df["Цена"] = pd.NA

        df = df.loc[:, [c for c in ["Артикул","Количество","Цена"] if c in df.columns] + [c for c in df.columns if c not in {"Артикул","Количество","Цена"}]]
        df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
        df.dropna(how="all", inplace=True)
        df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce")
        df["Цена"] = pd.to_numeric(df.get("Цена"), errors="coerce")
        df.dropna(subset=["Количество"], inplace=True)

        self.col_code = "Артикул"; self.col_qty = "Количество"; self.col_price = "Цена"
        self.df = df

        if self.df["Цена"].notna().any():
            self.original_sum = (self.df["Количество"] * self.df["Цена"]).sum()
            logger.info(f"Загружен счёт на {self.original_sum:,.2f} ₽")
        else:
            self.original_sum = 0.0
            logger.info("Загружен счёт без цен")

    def _append_row_like_source(self, row_src: pd.Series, code: str, qty: float, price: float, comment: str = "") -> None:
        add: Dict[str, Any] = {c: row_src.get(c, "") for c in self.base_columns}

        # код в «Товар»/«Код»/«Артикул»
        for name in self.base_columns:
            n = _normalize(name)
            if n in {"товар","код","артикул"}:
                add[name] = code

        if self.col_qty in add:
            add[self.col_qty] = qty

        # цена
        set_price = False
        for name in self.base_columns:
            if _normalize(name).startswith(("цена","стоимость","price")):
                add[name] = price
                set_price = True
                break
        if not set_price and "Цена" in add:
            add["Цена"] = price

        total = float(qty) * float(price)
        for name in self.base_columns:
            n = _normalize(name)
            if n.startswith("всего"):
                add[name] = total
            elif ("ндс" in n) or n.startswith("вт.чндс") or n.startswith("втчндс"):
                add[name] = round(total * VAT_RATE / (1.0 + VAT_RATE), 2)

        for name in self.base_columns:
            if _normalize(name).startswith("комментар"):
                add[name] = str(comment)

        self.result_rows.append(add)

    def process(self) -> None:
        """Если исходного не хватает — полная замена аналогами; иначе — 100% исходный. Метры аналогов под сумму."""
        self.result_rows.clear()
        self.log.clear()

        required_cols = {"Артикул","Количество"}
        missing = required_cols - set(self.df.columns)
        if missing:
            msg = f"В счёте нет колонок: {', '.join(missing)}"
            self.log.append(msg); logger.error(msg); return

        rules = self.app.rules if self.app and getattr(self.app, "rules", None) else {}
        rule_prices = getattr(self.app, "rules_price", {}) if self.app else {}

        for _, row in self.df.iterrows():
            art  = str(row[self.col_code]).strip()
            need = float(row[self.col_qty])
            price_main = row.get(self.col_price, pd.NA)

            # цена исходника — из счёта, иначе из каталога
            if pd.isna(price_main):
                cat_row = CATALOG[CATALOG["code"] == art]
                if not cat_row.empty:
                    price_main = pd.to_numeric(cat_row.iloc[0].get("price_rub"), errors="coerce")
            price_main = float(price_main) if pd.notna(price_main) else None

            target_amount = (need * price_main) if (price_main is not None) else None

            avail = self.stock.available(art)

            # 1) Хватает исходного → отгружаем целиком исходный (без аналогов)
            if avail >= need:
                ok = self.stock.allocate_full(art, need)
                if ok:
                    self._append_row_like_source(row, art, need, price_main if price_main is not None else 0.0, comment="основной товар")
                else:
                    self.log.append(f"{art}: не удалось списать исходный, переходим к аналогам")
                continue

            # 2) Не хватает исходника → ПОЛНАЯ замена аналогами (исходный вообще не отгружаем)
            base_code = extract_base_code(art)
            cand_bases = rules.get(base_code, [])

            if not cand_bases:
                self.log.append(f"{art}: аналогов нет по правилам — строка не закрыта")
                continue

            amount_left = target_amount if target_amount is not None else None
            meters_left = need  # если цены нет, будем закрывать по метрам

            for b in cand_bases:
                # Пул остатков по баз-коду аналога
                pool = self.stock.df[
                    (self.stock.df["BaseCode"] == b) & (self.stock.df[self.stock.stock_column] > 0)
                ].copy()
                if pool.empty:
                    continue
                pool = pool.sort_values(self.stock.stock_column, ascending=False)

                # Цена для этой базы по приоритету: файл правил → ценник склада (из каталога) → цена исходника
                price_override_for_b = None
                if base_code in rule_prices and b in rule_prices[base_code]:
                    price_override_for_b = rule_prices[base_code][b]

                for idx, r in pool.iterrows():
                    # условие выхода
                    if amount_left is not None:
                        if amount_left <= max(AMOUNT_ABS_TOL, (target_amount or 0.0) * AMOUNT_TOL):
                            amount_left = 0.0
                            break
                    else:
                        if meters_left <= 0:
                            break

                    avail_i = float(r[self.stock.stock_column])
                    if avail_i <= 0:
                        continue

                    pr = price_override_for_b
                    if pr is None:
                        pr = pd.to_numeric(r.get("price_rub"), errors="coerce")
                        if pd.isna(pr) and price_main is not None:
                            pr = price_main
                    if pr is None or (isinstance(pr, float) and (pr <= 0 or math.isnan(pr))):
                        continue
                    pr = float(pr)

                    if amount_left is not None:
                        qty_target = amount_left / pr if pr > 0 else 0.0
                        qty_take = max(QTY_STEP, round_step(qty_target, QTY_STEP))
                        qty_take = min(qty_take, avail_i)
                        if qty_take <= 0:
                            continue
                        # списываем
                        self.stock.df.at[idx, self.stock.stock_column] = avail_i - qty_take
                        amount_left = max(0.0, amount_left - qty_take * pr)
                        self._append_row_like_source(row, str(r["Артикул"]), qty_take, pr, comment=f"аналог {b}")
                    else:
                        # цены исходника нет — закрываем метры
                        qty_take = min(avail_i, meters_left)
                        if qty_take <= 0:
                            continue
                        self.stock.df.at[idx, self.stock.stock_column] = avail_i - qty_take
                        meters_left = max(0.0, meters_left - qty_take)
                        self._append_row_like_source(row, str(r["Артикул"]), qty_take, pr, comment=f"аналог {b}")

                # контроль выхода
                if amount_left is not None and amount_left <= max(AMOUNT_ABS_TOL, (target_amount or 0.0) * AMOUNT_TOL):
                    amount_left = 0.0
                    break
                if amount_left is None and meters_left <= 0:
                    break

            # итоги по строке
            if amount_left is not None:
                if amount_left > max(AMOUNT_ABS_TOL, (target_amount or 0.0) * AMOUNT_TOL):
                    self.log.append(f"{art}: не удалось добрать по сумме {amount_left:.2f} ₽ (нехватка складских остатков)")
            else:
                if meters_left > 0:
                    self.log.append(f"{art}: аналогов не хватило, осталось {meters_left:.2f} м")

        if not self.result_rows:
            msg = "аналогов не найдено, счёт не изменён"
            self.log.append(msg); logger.info(msg)

    def to_dataframe(self) -> pd.DataFrame:
        if not self.result_rows:
            return pd.DataFrame(columns=self.base_columns if self.base_columns else ["Артикул","Количество","Цена","Комментарий"])
        df = pd.DataFrame(self.result_rows)
        for c in self.base_columns:
            if c not in df.columns:
                df[c] = ""
        return df[self.base_columns]

    def save(self, path: str) -> None:
        out = self.to_dataframe()
        out.to_excel(path, index=False)
        logger.info(f"Счёт сохранён в {path}")

# ───────────────────────────── GUI ───────────────────────────────
class App:
    def __init__(self) -> None:
        self.root = Tk(); self.root.title("Invoice Builder")

        self.log_text = Text(self.root, height=20, width=90, font=("Consolas", 10))
        scroll_bar = Scrollbar(self.root, command=self.log_text.yview)
        scroll_bar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll_bar.set)
        self.log_text.pack(side="left", fill="both", expand=True)

        Button(self.root, text="Загрузить остатки", command=self.load_stock).pack()
        Button(self.root, text="Загрузить счёт", command=self.load_invoice).pack()
        Button(self.root, text="Загрузить правила", command=self.load_rules).pack()
        Button(self.root, text="Собрать счёт", command=self.build_invoice).pack()
        Button(self.root, text="Посмотреть логи", command=self.view_logs).pack()
        Button(self.root, text="Скачать логи", command=self.save_logs).pack()

        self.stock = StockManager()
        self.rules: dict[str, list[str]] = {}
        self.rules_price: dict[str, dict[str, float]] = {}

        try:
            if os.path.exists(RULES_DEFAULT_PATH):
                self.rules = load_analog_rules(RULES_DEFAULT_PATH)
                self.rules_price = load_analog_prices(RULES_DEFAULT_PATH)
                logger.info(f"Правила загружены: {len(self.rules)} баз (авто)")
            else:
                logger.info("Правила не найдены")
        except Exception:
            logger.exception("Автозагрузка правил не удалась")

        self.processor = InvoiceProcessor(stock=self.stock, app=self)
        self.stock_file: Optional[str] = None
        self.invoice_file: Optional[str] = None

    def gui_log(self, msg: str) -> None:
        self.log_text.insert(END, msg + "\n"); self.log_text.see(END); logger.info(msg)

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
        dst = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log files","*.log"),("All files","*.*")])
        if dst:
            shutil.copyfile("app.log", dst)
            messagebox.showinfo("Логи сохранены", f"Файл сохранён: {dst}")

    def load_stock(self) -> None:
        path = filedialog.askopenfilename()
        if not path: return
        try:
            self.stock.load(path); self.stock_file = path
            self.gui_log(f"Остатки загружены: {len(self.stock.df)} строк")
        except Exception as exc:
            logger.exception("Ошибка при загрузке остатков")
            messagebox.showerror("Ошибка", f"{type(exc).__name__}: {exc}")

    def load_invoice(self) -> None:
        path = filedialog.askopenfilename()
        if not path: return
        try:
            self.processor.load(path); self.invoice_file = path
            self.gui_log(f"Счёт загружен: {len(self.processor.df)} строк")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def load_rules(self) -> None:
        path = filedialog.askopenfilename()
        if not path: return
        try:
            self.rules = load_analog_rules(path)
            self.rules_price = load_analog_prices(path)
            self.gui_log(f"Правила загружены: {len(self.rules)} баз")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
        self.processor.app = self

    def build_invoice(self) -> None:
        if self.stock.df.empty or self.processor.df.empty:
            messagebox.showwarning("Внимание", "Загрузите остатки и счёт"); return

        self.processor.process()
        base, _ = os.path.splitext(os.path.basename(self.invoice_file))
        out_path = f"{base}_processed.xlsx"
        self.processor.save(out_path)

        self.gui_log("\n".join(self.processor.log))
        messagebox.showinfo("Готово", f"Новый счёт сохранён: {out_path}")

    def run(self) -> None:
        self.root.mainloop()

# ────────────────────────── entry point ──────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Invoice Builder")
    parser.add_argument("--cli", nargs=2, metavar=("STOCK","INVOICE"), help="run without GUI")
    parser.add_argument("--show-log", action="store_true", help="print last 100 log lines")
    parser.add_argument("--rules", help="path to analog rules file")
    args = parser.parse_args()

    if args.show_log:
        if os.path.exists("app.log"):
            with open("app.log", "r", encoding="utf-8") as f:
                lines = f.readlines()[-100:]
            for line in lines: print(line, end="")
        sys.exit(0)

    if args.cli:
        stock_path, invoice_path = args.cli
        stock = StockManager()
        rules_path = args.rules or RULES_DEFAULT_PATH
        try:
            rules = load_analog_rules(rules_path)
            rules_price = load_analog_prices(rules_path)
            logger.info(f"Правила загружены: {len(rules)} баз")
        except FileNotFoundError:
            rules = {}; rules_price = {}
            logger.info("Правила не найдены")
        dummy_app = type("Dummy", (), {"rules": rules, "rules_price": rules_price})()
        proc = InvoiceProcessor(stock=stock, app=dummy_app)
        stock.load(stock_path)
        proc.load(invoice_path)
        proc.process()
        base, _ = os.path.splitext(os.path.basename(invoice_path))
        out_path = f"{base}_processed.xlsx"
        proc.save(out_path)
        for line in proc.log: print(line)
    else:
        app = App()
        if args.rules:
            try:
                app.rules = load_analog_rules(args.rules)
                app.rules_price = load_analog_prices(args.rules)
                logger.info(f"Правила загружены: {len(app.rules)} баз")
            except FileNotFoundError:
                logger.info("Правила не найдены")
        app.processor.app = app
        app.run()
