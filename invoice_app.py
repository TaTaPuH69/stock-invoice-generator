#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations

import os, sys, re, logging, shutil, unicodedata, argparse
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple

import pandas as pd
from openpyxl import load_workbook

from flexy_catalog_loader import load_catalog

# ── logging ──────────────────────────────────────────────────────
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

# ── constants ────────────────────────────────────────────────────
VAT_RATE = 0.20
RULES_DEFAULT_PATH = "analogs_priority.xlsx"
FIXED_STOCK_ROW = 9
FIXED_STOCK_COL = 1

# если понадобится чуть «мягче» сравнивать суммы — можно подкрутить
AMOUNT_ABS_TOL = 0.0     # абсолютная толерантность по сумме, ₽
AMOUNT_REL_TOL = 0.00    # относительная, доля (0.05 = 5%)

# подбор аналогов по стоимости допускает отклонение 3%
VALUE_TOL = 0.03

_catalog = load_catalog()


@dataclass
class AnalogCandidate:
    base: str
    price: Optional[float] = None

# ── helpers ──────────────────────────────────────────────────────
def _normalize(col: str) -> str:
    return (
        str(col).lower()
        .replace("\n", "").replace("\r", "").replace("\t", "")
        .replace(" ", "").replace("-", "").replace("ё", "е")
    )

def _find_header_row(path: str, max_row: int = 40) -> int:
    for i in range(max_row):
        row = pd.read_excel(path, skiprows=i, nrows=1, header=None).fillna("")
        cells = [_normalize(str(c)) for c in row.values.ravel()]
        has_code = any(c.startswith(("код", "артикул")) for c in cells)
        has_qty  = any(c.startswith(("количест", "колво", "qty")) for c in cells)
        if has_code and has_qty:
            return i
    raise ValueError("Header row not found")

def extract_base_code(s: str) -> str:
    m = re.search(r"\d{4,6}", str(s))
    return m.group(0) if m else ""

def swap_base_in_code(code_str: str, new_base: str) -> str:
    """Заменяет первую группу 4–6 цифр на новую базу, остальное сохраняет."""
    code_str = str(code_str)
    m = re.search(r"(\d{4,6})", code_str)
    if not m:
        return str(new_base)
    s, e = m.span(1)
    return code_str[:s] + str(new_base) + code_str[e:]

def load_analog_rules(path: str = RULES_DEFAULT_PATH) -> Dict[str, List[AnalogCandidate]]:
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")

    def norm(x: str) -> str:
        return str(x).strip().lower().replace(" ", "")

    base_candidates = {"товар", "база", "код", "артикул"}
    base_col = next((c for c in df.columns if norm(c) in base_candidates), df.columns[0])

    analog_cols = [c for c in df.columns if ("аналог" in norm(c) or "analog" in norm(c))]
    price_cols = [c for c in df.columns if ("цена" in norm(c) or "price" in norm(c))]

    if not analog_cols:
        analog_cols = [c for c in df.columns if c not in {base_col} | set(price_cols)]

    def suffix(name: str) -> str:
        m = re.search(r"(\d+)$", norm(name))
        return m.group(1) if m else ""

    price_by_suffix = {suffix(c): c for c in price_cols if suffix(c)}
    default_price_col = next((c for c in price_cols if suffix(c) == ""), None)

    rules: Dict[str, List[AnalogCandidate]] = {}
    for _, row in df.iterrows():
        base = extract_base_code(row.get(base_col, ""))
        if not base:
            continue
        candidates: List[AnalogCandidate] = []
        for c in analog_cols:
            cand_base = extract_base_code(row.get(c, ""))
            if not cand_base or cand_base == base:
                continue
            suf = suffix(c)
            price_val = None
            price_col = price_by_suffix.get(suf)
            if price_col:
                price_val = pd.to_numeric(row.get(price_col, ""), errors="coerce")
            elif default_price_col:
                price_val = pd.to_numeric(row.get(default_price_col, ""), errors="coerce")
            price_val = float(price_val) if pd.notna(price_val) else None
            if cand_base not in {cand.base for cand in candidates}:
                candidates.append(AnalogCandidate(base=cand_base, price=price_val))
        if candidates:
            rules[base] = candidates
    return rules

def _is_filled(val) -> bool:
    return pd.notna(val) and str(val).strip() != ""

# ── stock ────────────────────────────────────────────────────────
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

        # подцепим цены (если есть в каталоге)
        if isinstance(_catalog, pd.DataFrame) and {"code", "price_rub"} <= set(_catalog.columns):
            cat = _catalog.copy()
            cat["code"] = cat["code"].astype(str).str.strip()
            enrich = cat[["code", "price_rub"]].rename(columns={"code": "Артикул"})
            enrich["price_rub"] = pd.to_numeric(enrich["price_rub"], errors="coerce")
            self.df = self.df.merge(enrich, on="Артикул", how="left")

        self.stock_column = "Остаток"
        logger.info(f"Загружено {len(self.df)} строк остатков")

    def total_for_article(self, article: str) -> float:
        rows = self.df[self.df["Артикул"] == article]
        return float(rows[self.stock_column].sum()) if not rows.empty else 0.0

    def total_for_base(self, base_code: str) -> float:
        rows = self.df[self.df["BaseCode"] == base_code]
        return float(rows[self.stock_column].sum()) if not rows.empty else 0.0

    def avg_price_for_base(self, base_code: str) -> Optional[float]:
        pool = self.df[self.df["BaseCode"] == base_code]
        if pool.empty:
            return None
        pr = pd.to_numeric(pool.get("price_rub"), errors="coerce")
        pr = pr[pr.notna()]
        if pr.empty:
            return None
        return float(pr.mean())

    def allocate_partial(self, article: str, qty: float) -> float:
        """Снимает со склада до qty метров по точному артикулу и возвращает отпущенное количество."""
        pool = self.df[(self.df["Артикул"] == article) & (self.df[self.stock_column] > 0)]
        left = float(qty)
        shipped = 0.0
        for idx, r in pool.iterrows():
            if left <= 0:
                break
            avail = float(r[self.stock_column])
            take = min(avail, left)
            if take > 0:
                self.df.at[idx, self.stock_column] = avail - take
                left -= take
                shipped += take
        return shipped

    def allocate_value_by_base(
        self,
        base_code: str,
        target_value: float,
        price_override: Optional[float],
        tol_value: float,
    ) -> Tuple[List[Tuple[pd.Series, float]], float]:
        pool = self.df[(self.df["BaseCode"] == base_code) & (self.df[self.stock_column] > 0)].copy()
        if pool.empty or target_value <= tol_value:
            return [], target_value

        if price_override is not None:
            pool["unit_price"] = float(price_override)
        else:
            pool["unit_price"] = pd.to_numeric(pool.get("price_rub"), errors="coerce")

        missing = pool[pool["unit_price"].isna()]
        for _, r in missing.iterrows():
            logger.warning(f"{r['Артикул']}: отсутствует цена; строка пропущена")
        pool.dropna(subset=["unit_price"], inplace=True)
        pool = pool[pool["unit_price"] > 0]
        if pool.empty:
            return [], target_value

        if price_override is not None:
            pool = pool.sort_values(self.stock_column, ascending=False)
        else:
            pool = pool.sort_values(["unit_price", self.stock_column], ascending=[True, False])

        remaining = float(target_value)
        allocations: List[Tuple[pd.Series, float]] = []
        for idx, r in pool.iterrows():
            if remaining <= tol_value:
                break
            avail = float(r[self.stock_column])
            unit_price = float(r["unit_price"])
            qty = min(avail, remaining / unit_price)
            qty = round(qty, 2)
            if qty <= 0:
                continue
            self.df.at[idx, self.stock_column] = avail - qty
            row_data = self.df.loc[idx].copy()
            allocations.append((row_data, qty))
            remaining -= qty * unit_price

        return allocations, remaining

    def display_name_for_base(self, base_code: str) -> str:
        rows = self.df[self.df["BaseCode"] == base_code]
        if not rows.empty:
            return str(rows.iloc[0]["Артикул"])
        return f"Профиль {base_code}"

# ── invoice ──────────────────────────────────────────────────────
@dataclass
class InvoiceProcessor:
    stock: StockManager
    app: Optional["App"] = None
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    raw_df: pd.DataFrame = field(default_factory=pd.DataFrame)

    original_sum: float = 0.0
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None
    header_row_idx: int = 0
    output_columns: List[str] = field(default_factory=list)

    col_code: str = "Артикул"
    col_qty: str = "Количество"
    col_price: str = "Цена"
    col_code_orig: Optional[str] = None
    col_qty_orig: Optional[str] = None
    col_price_orig: Optional[str] = None
    col_name_orig: Optional[str] = None

    def load(self, path: str) -> None:
        hdr = _find_header_row(path)
        base = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)
        self.invoice_path = path
        self.header_row_idx = hdr

        self.raw_df = base.copy()
        self.output_columns = list(base.columns)
        if "Комментарий" not in self.output_columns:
            self.output_columns.append("Комментарий")

        df = base.copy()
        rename_map: dict[str, str] = {}
        for col in df.columns:
            norm = _normalize(col)
            if norm.startswith(("код", "артикул")):
                rename_map[col] = "Артикул"
            elif norm.startswith(("количест", "колво", "qty")):
                rename_map[col] = "Количество"
            elif norm.startswith(("цена", "стоимость", "price")):
                rename_map[col] = "Цена"
            elif norm.startswith(("товар", "наимен", "product", "item")):
                rename_map[col] = "Товар"
        df.rename(columns=rename_map, inplace=True)

        if "Цена" not in df.columns:
            df["Цена"] = pd.NA

        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        df.dropna(how="all", inplace=True)

        df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce")
        df["Цена"] = pd.to_numeric(df.get("Цена"), errors="coerce")
        df.dropna(subset=["Количество"], inplace=True)

        self.raw_df = self.raw_df.loc[df.index].copy()
        df.reset_index(drop=True, inplace=True)
        self.raw_df.reset_index(drop=True, inplace=True)

        self.col_code = "Артикул"
        self.col_qty = "Количество"
        self.col_price = "Цена"

        inv_map = {v: k for k, v in rename_map.items()}
        self.col_code_orig = inv_map.get("Артикул")
        self.col_qty_orig = inv_map.get("Количество")
        self.col_price_orig = inv_map.get("Цена")
        self.col_name_orig = inv_map.get("Товар")

        if self.col_code_orig is None and "Артикул" in self.output_columns:
            self.col_code_orig = "Артикул"
        if self.col_qty_orig is None and "Количество" in self.output_columns:
            self.col_qty_orig = "Количество"
        if self.col_price_orig is None and "Цена" in self.output_columns:
            self.col_price_orig = "Цена"
        if self.col_name_orig is None and "Товар" in self.output_columns:
            self.col_name_orig = "Товар"

        self.df = df

        if self.df["Цена"].notna().any():
            self.original_sum = (self.df["Количество"] * self.df["Цена"]).sum()
            logger.info(f"Загружен счёт на {self.original_sum:,.2f} ₽")
        else:
            self.original_sum = 0.0
            logger.info("Загружен счёт без цен")

    def process(self) -> None:
        self.result_rows.clear()
        self.log.clear()

        rules = self.app.rules if self.app and getattr(self.app, "rules", None) else {}

        for (idx, row), (_, orig_row) in zip(self.df.iterrows(), self.raw_df.iterrows()):
            art = str(row.get(self.col_code, "")).strip()
            if not art:
                continue
            need = float(row.get(self.col_qty, 0) or 0)

            price = row.get(self.col_price, pd.NA)
            if pd.isna(price):
                cat_row = _catalog[_catalog["code"] == art]
                if not cat_row.empty:
                    price = pd.to_numeric(cat_row.iloc[0].get("price_rub"), errors="coerce")
            orig_price = float(price) if pd.notna(price) else 0.0

            target_sum = need * orig_price
            tol_value = max(10.0, VALUE_TOL * target_sum)

            shipped = self.stock.allocate_partial(art, need)
            shipped_sum = shipped * orig_price
            self.log.append(
                f"{art}: отгружено {shipped:.2f} м исходного кода по {orig_price:.2f} → {shipped_sum:.2f} ₽"
            )
            if shipped > 0:
                rec = {c: orig_row.get(c, "") for c in self.output_columns}
                if self.col_code_orig:
                    rec[self.col_code_orig] = art
                if self.col_qty_orig:
                    rec[self.col_qty_orig] = shipped
                if self.col_price_orig:
                    rec[self.col_price_orig] = orig_price
                self.result_rows.append(rec)

            value_left = max(target_sum - shipped_sum, 0.0)
            if value_left <= tol_value:
                continue

            self.log.append(f"{art}: добираем {value_left:.2f} ₽ аналогами")
            base_code = extract_base_code(art)
            candidates = rules.get(base_code, [])
            for cand in candidates:
                allocs, value_left = self.stock.allocate_value_by_base(
                    cand.base, value_left, cand.price, tol_value
                )
                for stock_row, qty in allocs:
                    analog_code = str(stock_row.get("Артикул"))
                    unit_price = cand.price if cand.price is not None else float(stock_row.get("price_rub", 0) or 0)
                    analog_sum = qty * unit_price
                    rec = {c: orig_row.get(c, "") for c in self.output_columns}
                    if self.col_code_orig:
                        rec[self.col_code_orig] = analog_code
                    if self.col_name_orig:
                        rec[self.col_name_orig] = analog_code
                    if self.col_qty_orig:
                        rec[self.col_qty_orig] = qty
                    if self.col_price_orig:
                        rec[self.col_price_orig] = unit_price
                    self.result_rows.append(rec)
                    self.log.append(
                        f"{art}: {qty:.2f} м → {analog_code} по {unit_price:.2f} ({analog_sum:.2f} ₽)"
                    )
                if value_left <= tol_value:
                    break

            if value_left > tol_value:
                self.log.append(
                    f"{art}: не хватило {value_left:.2f} ₽ — аналоги закончились"
                )

    def to_dataframe(self) -> pd.DataFrame:
        if not self.result_rows:
            return pd.DataFrame(columns=self.output_columns)
        return pd.DataFrame(self.result_rows, columns=self.output_columns)

    def save(self, path: str) -> None:
        """Перезаписываем ТОЛЬКО табличную часть исходного файла, сохраняя шапку/визуал."""
        if not self.invoice_path:
            raise RuntimeError("invoice_path is not set")

        out_df = self.to_dataframe()
        rows = out_df.to_dict(orient="records")

        wb = load_workbook(self.invoice_path)
        ws = wb.active

        header_row_1 = self.header_row_idx + 1
        headers = [cell.value for cell in ws[header_row_1]]
        norm_headers = [_normalize(h) if h is not None else "" for h in headers]

        def find_index(prefixes: tuple[str, ...]) -> Optional[int]:
            for i, h in enumerate(norm_headers):
                if any(h.startswith(p) for p in prefixes):
                    return i
            return None

        idx_item  = find_index(("товар", "наимен", "product", "item"))
        idx_code  = find_index(("код", "артикул"))
        idx_qty   = find_index(("количест", "колво", "qty"))
        idx_price = find_index(("цена", "стоимость", "price"))
        idx_total = find_index(("всего", "итого", "сумма"))
        idx_vat   = find_index(("вт.чндс", "ндс"))
        idx_comm  = next((i for i, h in enumerate(norm_headers) if "комментар" in h), None)

        # очистить старые строки
        start = header_row_1 + 1
        if start <= ws.max_row:
            ws.delete_rows(start, ws.max_row - header_row_1)

        # записать новые строки
        r = start
        for rec in rows:
            values = [rec.get(h, "") for h in headers]

            q = 0.0
            p = 0.0
            try:
                if self.col_qty_orig:
                    q = float(rec.get(self.col_qty_orig, 0) or 0)
                if self.col_price_orig:
                    p = float(rec.get(self.col_price_orig, 0) or 0)
            except Exception:
                q = p = 0.0
            if idx_total is not None:
                values[idx_total] = round(q * p, 2)
            if idx_vat is not None:
                values[idx_vat] = round(q * p * VAT_RATE, 2)
            if idx_comm is not None and idx_comm < len(values):
                values[idx_comm] = rec.get("Комментарий", "")

            for c, val in enumerate(values, start=1):
                ws.cell(row=r, column=c, value=val)
            r += 1

        wb.save(path)
        logger.info(f"Счёт сохранён в {path}")

# ── GUI ──────────────────────────────────────────────────────────
try:
    from tkinter import Tk, filedialog, messagebox, Text, Scrollbar, Button, END, Toplevel
except Exception:
    Tk = filedialog = messagebox = Text = Scrollbar = Button = END = Toplevel = None

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
        self.rules: Dict[str, List[AnalogCandidate]] = {}
        try:
            if os.path.exists(RULES_DEFAULT_PATH):
                self.rules = load_analog_rules(RULES_DEFAULT_PATH)
                logger.info(f"Правила загружены: {len(self.rules)} баз (авто)")
            else:
                logger.info("Правила не найдены")
        except Exception:
            logger.exception("Автозагрузка правил не удалась")

        self.processor = InvoiceProcessor(stock=self.stock, app=self)
        self.stock_file: Optional[str] = None
        self.invoice_file: Optional[str] = None

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

    def load_stock(self) -> None:
        path = filedialog.askopenfilename()
        if not path:
            return
        try:
            self.stock.load(path)
            self.stock_file = path
            self.gui_log(f"Остатки загружены: {len(self.stock.df)} строк")
        except Exception as exc:
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
            self.processor.app = self
            self.gui_log(f"Правила загружены: {len(self.rules)} баз")
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

    def run(self) -> None:
        self.root.mainloop()

# ── entrypoint ───────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Invoice Builder")
    parser.add_argument("--cli", nargs=2, metavar=("STOCK", "INVOICE"), help="run without GUI")
    parser.add_argument("--show-log", action="store_true", help="print last 100 log lines")
    parser.add_argument("--rules", help="path to analog rules file")
    args = parser.parse_args()

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
        dummy_app = type("Dummy", (), {"rules": rules})()
        proc = InvoiceProcessor(stock=stock, app=dummy_app)
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
        app.processor.app = app
        app.run()
