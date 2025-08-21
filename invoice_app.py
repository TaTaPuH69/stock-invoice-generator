#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations

import os, sys, re, logging, shutil, unicodedata, argparse
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field
from typing import List, Optional

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

_catalog = load_catalog()

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

def load_analog_rules(path: str = RULES_DEFAULT_PATH) -> dict[str, list[str]]:
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")

    def norm(x: str) -> str:
        return str(x).strip().lower().replace(" ", "")

    base_candidates = {"товар", "база", "код", "артикул"}
    base_col = next((c for c in df.columns if norm(c) in base_candidates), None)
    if base_col is None:
        base_col = df.columns[0]

    analog_cols = [c for c in df.columns if ("аналог" in norm(c) or "analog" in norm(c))]
    if not analog_cols:
        analog_cols = [c for c in df.columns if c != base_col]

    rules: dict[str, list[str]] = {}
    for _, row in df.iterrows():
        base = extract_base_code(row.get(base_col, ""))
        if not base:
            continue
        lst: list[str] = []
        for c in analog_cols:
            cand = extract_base_code(row.get(c, ""))
            if cand and cand != base and cand not in lst:
                lst.append(cand)
        if lst:
            rules[base] = lst
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

    def allocate_exact(self, article: str, qty: float) -> bool:
        """Резервирует qty по точному артикулу, если хватает. Возвращает True/False."""
        if self.total_for_article(article) < qty:
            return False
        left = qty
        idxs = self.df.index[self.df["Артикул"] == article]
        for idx in idxs:
            if left <= 0:
                break
            avail = float(self.df.at[idx, self.stock_column])
            take = min(avail, left)
            if take > 0:
                self.df.at[idx, self.stock_column] = avail - take
                left -= take
        return True

    def allocate_by_basecode(self, base_code: str, qty: float) -> bool:
        """Резервирует qty суммарно по позиции с нужной базой. Возвращает True/False."""
        if self.total_for_base(base_code) < qty or qty <= 0:
            return False
        left = qty
        pool = self.df[(self.df["BaseCode"] == base_code) & (self.df[self.stock_column] > 0)].copy()
        pool = pool.sort_values(self.stock_column, ascending=False)
        for idx, r in pool.iterrows():
            if left <= 0:
                break
            avail = float(r[self.stock_column])
            take = min(avail, left)
            if take > 0:
                self.df.at[idx, self.stock_column] = avail - take
                left -= take
        return left <= 1e-9

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

    original_sum: float = 0.0
    result_rows: List[dict] = field(default_factory=list)
    log: List[str] = field(default_factory=list)
    invoice_path: Optional[str] = None
    header_row_idx: int = 0
    headers: List[str] = field(default_factory=list)

    col_code: str = "Артикул"
    col_qty: str = "Количество"
    col_price: str = "Цена"

    def load(self, path: str) -> None:
        hdr = _find_header_row(path)
        df = pd.read_excel(path, skiprows=hdr, header=0, dtype=str)
        self.invoice_path = path
        self.header_row_idx = hdr

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

        self.col_code = "Артикул"
        self.col_qty = "Количество"
        self.col_price = "Цена"

        self.df = df
        self.headers = list(df.columns)
        if "Комментарий" not in self.headers:
            self.headers.append("Комментарий")

        if self.df["Цена"].notna().any():
            self.original_sum = (self.df["Количество"] * self.df["Цена"]).sum()
            logger.info(f"Загружен счёт на {self.original_sum:,.2f} ₽")
        else:
            self.original_sum = 0.0
            logger.info("Загружен счёт без цен")

    def process(self) -> None:
        """При замене подбираем КОЛИЧЕСТВО аналога так, чтобы сумма была максимально близкой к исходной."""
        self.result_rows.clear()
        self.log.clear()

        rules = self.app.rules if self.app and getattr(self.app, "rules", None) else {}

        for _, row in self.df.iterrows():
            art = row.get(self.col_code, "")
            need = float(row.get(self.col_qty, 0) or 0)
            price = row.get(self.col_price, pd.NA)  # цена исходного
            item_name = row.get("Товар", "")

            # если нет цены — попробуем из каталога
            if pd.isna(price):
                cat_row = _catalog[_catalog["code"] == art]
                if not cat_row.empty:
                    price = pd.to_numeric(cat_row.iloc[0].get("price_rub"), errors="coerce")

            # если и сейчас цены нет — не сможем подгонять по сумме, оставим попытку «как есть»
            target_amount = None if pd.isna(price) else float(price) * need

            # 1) пробуем исходный товар целиком
            if self.stock.allocate_exact(art, need):
                rec = {c: row.get(c, "") for c in self.headers}
                rec["Артикул"] = art
                rec["Количество"] = need
                if pd.notna(price):
                    rec["Цена"] = float(price)
                self.result_rows.append(rec)
                continue

            # 2) выбираем ОДИН лучший аналог ─ тот, где возможная сумма ближе к целевой
            base_code = extract_base_code(art)
            cand_bases = rules.get(base_code, [])

            best = None  # (diff, base_cand, qty_use, pr_unit)
            for base_cand in cand_bases:
                avail_total = int(self.stock.total_for_base(base_cand))
                if avail_total <= 0:
                    continue

                pr_unit = self.stock.avg_price_for_base(base_cand)
                if pr_unit is None or pr_unit <= 0:
                    # fallback: если не знаем цену аналога, смысла подгонять нет
                    pr_unit = float(price) if not pd.isna(price) else None
                if pr_unit is None or pr_unit <= 0:
                    continue

                if target_amount is None:
                    # нет целевой суммы — будем стремиться оставить исходное количество,
                    # но не больше доступного
                    qty_target = min(avail_total, int(round(need)))
                else:
                    qty_target = int(round(target_amount / pr_unit))
                    qty_target = max(1, qty_target)
                    qty_target = min(avail_total, qty_target)

                # пересчёт суммы и отклонения
                total_here = qty_target * pr_unit
                diff = abs((target_amount or (need * pr_unit)) - total_here)

                # если равные diff — отдаём приоритет более раннему в списке правил
                if (best is None) or (diff < best[0]):
                    best = (diff, base_cand, qty_target, pr_unit)

            if best is not None:
                _, base_cand, qty_use, pr_unit = best

                # финальная попытка аллокации (на момент выбора склад мог измениться)
                avail_now = int(self.stock.total_for_base(base_cand))
                qty_use = max(1, min(avail_now, int(qty_use)))
                if qty_use > 0 and self.stock.allocate_by_basecode(base_cand, qty_use):
                    new_code = swap_base_in_code(art, base_cand)
                    new_item = self.stock.display_name_for_base(base_cand)

                    rec = {c: row.get(c, "") for c in self.headers}
                    rec["Артикул"] = new_code                # колонка «Код»
                    rec["Товар"]   = new_item                # колонка «Товар»
                    rec["Количество"] = float(qty_use)       # подобранное кол-во
                    rec["Цена"] = float(pr_unit)             # цена аналога за единицу

                    # лог — показываем расхождение по сумме
                    target = target_amount if target_amount is not None else need * pr_unit
                    actual = qty_use * pr_unit
                    delta = actual - target
                    self.log.append(
                        f"{art} → {new_code}: цена {pr_unit:.2f}, кол-во {qty_use} "
                        f"(цель {target:.2f}, факт {actual:.2f}, Δ {delta:+.2f})"
                    )

                    self.result_rows.append(rec)
                    continue

            # если сюда дошли — аналог не найден/не подобрался
            self.log.append(f"{art}: аналог не подобран — строка без изменений")
            rec = {c: row.get(c, "") for c in self.headers}
            self.result_rows.append(rec)

    def to_dataframe(self) -> pd.DataFrame:
        if not self.result_rows:
            return pd.DataFrame(columns=self.headers)
        df = pd.DataFrame(self.result_rows)
        keep = [c for c in self.headers if c in df.columns]
        if "Комментарий" in df.columns and "Комментарий" not in keep:
            keep.append("Комментарий")
        return df[keep]

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
            values = [None] * len(headers)

            if idx_item  is not None: values[idx_item]  = rec.get("Товар", rec.get("Наименование", ""))
            if idx_code  is not None: values[idx_code]  = rec.get("Артикул", rec.get("Код", ""))
            if idx_qty   is not None: values[idx_qty]   = rec.get("Количество", "")
            if idx_price is not None: values[idx_price] = rec.get("Цена", "")

            if idx_comm is not None:
                values[idx_comm] = rec.get("Комментарий", "")

            # Всего / НДС
            try:
                q = float(rec.get("Количество", 0) or 0)
                p = float(rec.get("Цена", 0) or 0)
            except Exception:
                q = p = 0.0
            if idx_total is not None:
                values[idx_total] = round(q * p, 2)
            if idx_vat is not None:
                values[idx_vat] = round(q * p * VAT_RATE, 2)

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
        self.rules: dict[str, list[str]] = {}
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
