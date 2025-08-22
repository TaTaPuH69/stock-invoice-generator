# coding: utf-8
"""Stock invoice generator with GUI and CLI.
Single file implementation following specification.
"""
from __future__ import annotations

import os
import sys
import re
import argparse
import logging
import unicodedata
import fnmatch
import io
from typing import Dict, List, Optional, Tuple

import pandas as pd
from pandas import DataFrame
from openpyxl import Workbook  # ensure openpyxl available
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import flexy_catalog_loader

# ----------------------------------------------------------------------------
# logging setup
log_stream = io.StringIO()
logger = logging.getLogger("invoice")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    fh = logging.FileHandler("invoice_app.log", encoding="utf-8")
    fh.setFormatter(fmt)
    stream_h = logging.StreamHandler(log_stream)
    stream_h.setFormatter(fmt)
    logger.addHandler(sh)
    logger.addHandler(fh)
    logger.addHandler(stream_h)
    logger.propagate = False

# ----------------------------------------------------------------------------
# constants
VALUE_EPS = 0.01
VAT_RATE = 20.0 / 120.0

# ----------------------------------------------------------------------------
# helpers

def resolve_existing_path(path: str) -> str:
    """Resolve path with normalization and search.
    Logs resolution info. Returns found path or original expanded path.
    """
    orig = path
    path = os.path.expanduser(path)
    def exists(p: str) -> Optional[str]:
        if os.path.exists(p):
            return os.path.abspath(p)
        return None
    for candidate in [path, os.path.abspath(path)]:
        res = exists(candidate)
        if res:
            logger.info("[%s] resolved to: %s", orig, res)
            return res
    for form in ["NFC", "NFD", "NFKC", "NFKD"]:
        normed = unicodedata.normalize(form, path)
        res = exists(normed)
        if res:
            logger.info("[%s] resolved to: %s", orig, res)
            return res
    search_dirs: List[str] = []
    cwd = os.getcwd()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    search_dirs.extend([cwd, script_dir, os.path.dirname(script_dir)])
    home = os.path.expanduser("~")
    search_dirs.extend([home, os.path.join(home, "Downloads"), os.path.join(home, "Desktop")])
    search_dirs = [d for d in search_dirs if os.path.isdir(d)]
    base = os.path.basename(path)
    pattern = base if any(ch in base for ch in "*?") else f"*{base}*"
    found: List[str] = []
    for d in search_dirs:
        for root, dirs, files in os.walk(d):
            if root.startswith("/System") or root.startswith("/Library") or root.startswith("/usr"):
                dirs[:] = []
                continue
            for name in files:
                if fnmatch.fnmatch(name.lower(), pattern.lower()):
                    full = os.path.join(root, name)
                    found.append(full)
                    if name.lower() == base.lower():
                        logger.info("[%s] resolved to: %s", orig, full)
                        return full
    if found:
        logger.info("[%s] not found. Candidates:", orig)
        for f in found[:10]:
            logger.info("  %s", f)
    else:
        logger.info("[%s] not found", orig)
    return path

def to_str_cell(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_float_cell(x) -> Optional[float]:
    if pd.isna(x):
        return None
    try:
        return float(str(x).replace(" ", "").replace(",", "."))
    except Exception:
        return None

def round_qty(x: float) -> float:
    return round(float(x), 2)

def round_money(x: float) -> float:
    return round(float(x), 2)

def looks_like_code_scalar(s: str) -> bool:
    if not s:
        return False
    if re.search(r"[А-Яа-я]", s):
        return False
    return bool(re.search(r"\d{4,6}", s))

def extract_base_code(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"\d{6}", s)
    if m:
        return m.group(0)
    m = re.search(r"\d{5}", s)
    if m:
        return m.group(0)
    m = re.search(r"\d{4}", s)
    if m:
        return m.group(0)
    return ""

# ----------------------------------------------------------------------------
# loading

def load_analog_rules(path: str) -> Dict[str, List[Tuple[str, Optional[float]]]]:
    path = resolve_existing_path(path)
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")
    columns = list(df.columns)
    base_col = 0
    for i, c in enumerate(columns):
        s = to_str_cell(c).lower()
        if any(k in s for k in ["товар", "база", "код", "артикул", "base"]):
            base_col = i
            break
    analog_cols: List[int] = []
    price_cols: List[int] = []
    for i, c in enumerate(columns):
        s = to_str_cell(c).lower()
        if "аналог" in s or "analog" in s:
            analog_cols.append(i)
        if "цена" in s or "price" in s:
            price_cols.append(i)
    if not analog_cols:
        analog_cols = [i for i in range(len(columns)) if i != base_col and i not in price_cols]
    pairs: Dict[int, Optional[int]] = {}
    for ac in analog_cols:
        suf = re.findall(r"\d+", to_str_cell(columns[ac]))
        price_match = None
        if suf:
            for pc in price_cols:
                if re.findall(r"\d+", to_str_cell(columns[pc])) == suf:
                    price_match = pc
                    break
        if price_match is None:
            for offs in range(1,4):
                if ac+offs in price_cols:
                    price_match = ac+offs
                    break
        pairs[ac] = price_match
    rules: Dict[str, List[Tuple[str, Optional[float]]]] = {}
    for _, row in df.iterrows():
        base_code = extract_base_code(to_str_cell(row.iat[base_col]))
        if not base_code:
            continue
        lst: List[Tuple[str, Optional[float]]] = []
        for ac, pc in pairs.items():
            analog_code = extract_base_code(to_str_cell(row.iat[ac]))
            if not analog_code:
                continue
            price_val = None
            if pc is not None:
                price_val = to_float_cell(row.iat[pc])
            lst.append((analog_code, price_val))
        if lst:
            rules.setdefault(base_code, []).extend(lst)
    return rules

def load_flexy_catalog() -> DataFrame:
    cat = flexy_catalog_loader.load_catalog()
    if isinstance(cat, pd.DataFrame):
        cat = cat.copy()
        cat["BaseCode"] = cat["code"].map(extract_base_code)
    else:
        cat = pd.DataFrame(columns=["code", "price_rub", "BaseCode"])
    return cat

def load_profiles_mapping(path: str) -> Dict[str, str]:
    path = resolve_existing_path(path)
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, dtype=str).fillna("")
    cols = list(df.columns)
    code_col = 0
    name_col = 1 if len(cols) > 1 else 0
    for i, c in enumerate(cols):
        s = to_str_cell(c).lower()
        if any(k in s for k in ["код", "артикул", "sku", "code"]):
            code_col = i
        if any(k in s for k in ["товар", "наимен", "product", "name"]):
            name_col = i
    mapping: Dict[str, str] = {}
    for _, row in df.iterrows():
        code = to_str_cell(row.iat[code_col])
        name = to_str_cell(row.iat[name_col])
        if code and name:
            mapping[code] = name
    return mapping

def load_stock(path: str, catalog: DataFrame) -> DataFrame:
    path = resolve_existing_path(path)
    df = pd.read_excel(path, header=None, skiprows=9, usecols=[0,1])
    df.columns = ["code", "qty"]
    df["code"] = df["code"].apply(to_str_cell)
    df["qty"] = df["qty"].apply(to_float_cell).fillna(0)
    df = df.groupby("code", as_index=False)["qty"].sum()
    df["BaseCode"] = df["code"].map(extract_base_code)
    catalog = catalog[["code", "price_rub", "BaseCode"]]
    df = df.merge(catalog, on="code", how="left")
    return df

def read_invoice_table(path: str) -> Tuple[List[str], DataFrame, Dict[str,int]]:
    path = resolve_existing_path(path)
    df = pd.read_excel(path, header=None)
    header_idx = 0
    code_keys = ["код","артикул","sku","code","код/артикул","артикул/код"]
    qty_keys = ["кол-во","количество","кол.","qty","quantity","кол"]
    for i in range(min(80, len(df))):
        row = [to_str_cell(x).lower() for x in df.iloc[i]]
        has_code = any(any(k in cell for k in code_keys) for cell in row)
        has_qty = any(any(k in cell for k in qty_keys) for cell in row)
        if has_code and has_qty:
            header_idx = i
            break
    headers = [to_str_cell(x) for x in df.iloc[header_idx]]
    data = df.iloc[header_idx+1:].reset_index(drop=True)
    roles: Dict[str,int] = {}
    for idx, h in enumerate(headers):
        hl = h.lower()
        if any(k in hl for k in ["№","номер","no"]):
            roles.setdefault("num", idx)
        if any(k in hl for k in ["код","артикул","sku","code","код/артикул","артикул/код"]):
            roles.setdefault("code", idx)
        if any(k in hl for k in ["товар","наименование","product","name","наим"]):
            roles.setdefault("name", idx)
        if any(k in hl for k in ["кол-во","количество","кол.","qty","quantity","кол"]):
            roles.setdefault("qty", idx)
        if any(k in hl for k in ["ед","unit"]):
            roles.setdefault("unit", idx)
        if any(k in hl for k in ["цена","price"]):
            roles.setdefault("price", idx)
        if any(k in hl for k in ["в т.ч. ндс","ндс"]):
            roles.setdefault("nds", idx)
        if any(k in hl for k in ["сумма","итого","total","стоимость","всего"]):
            roles.setdefault("total", idx)
    return headers, data, roles

# ----------------------------------------------------------------------------
# processing

def process_invoice(headers: List[str], data: DataFrame, roles: Dict[str,int], stock_df: DataFrame,
                     rules: Dict[str,List[Tuple[str,Optional[float]]]], profiles: Dict[str,str],
                     catalog: DataFrame) -> DataFrame:
    stock_dict: Dict[str, Dict[str, Optional[float]]] = {}
    for _, r in stock_df.iterrows():
        stock_dict[r["code"]] = {
            "qty": float(r["qty"]),
            "price": to_float_cell(r.get("price_rub")),
            "base": to_str_cell(r.get("BaseCode"))
        }
    stock_by_base: Dict[str, List[str]] = {}
    for code, info in stock_dict.items():
        stock_by_base.setdefault(info["base"], []).append(code)

    result_rows: List[List] = []
    for row_idx in range(len(data)):
        row = data.iloc[row_idx]
        qty = to_float_cell(row.iat[roles.get("qty", -1)]) if "qty" in roles else None
        if qty is None or qty <= 0:
            continue
        code_cell = to_str_cell(row.iat[roles.get("code", -1)]) if "code" in roles else ""
        name_cell = to_str_cell(row.iat[roles.get("name", -1)]) if "name" in roles else ""
        price_src = to_float_cell(row.iat[roles.get("price", -1)]) if "price" in roles else None
        total_src = to_float_cell(row.iat[roles.get("total", -1)]) if "total" in roles else None
        code = code_cell
        if not looks_like_code_scalar(code):
            code = extract_base_code(name_cell)
        if not code:
            code = extract_base_code(code_cell)
        if not code:
            text = " ".join([to_str_cell(x) for x in row.tolist()])
            code = extract_base_code(text)
        base_code = extract_base_code(code)
        name = profiles.get(code, code)
        if not code or qty is None:
            continue
        stock_info = stock_dict.get(code)
        if stock_info and stock_info["qty"] >= qty:
            chosen_price = price_src if price_src is not None else stock_info["price"]
            if chosen_price is None:
                cat_price = catalog.loc[catalog["code"]==code, "price_rub"].dropna()
                if not cat_price.empty:
                    chosen_price = float(cat_price.iloc[0])
            if chosen_price is None:
                logger.info("%s: не удалось вычислить цену", code)
                continue
            chosen_price = round_money(chosen_price)
            total = round_money(chosen_price * qty)
            nds = round_money(total * VAT_RATE)
            new_row = list(row)
            if "code" in roles:
                new_row[roles["code"]] = code
            if "name" in roles:
                new_row[roles["name"]] = name
            if "qty" in roles:
                new_row[roles["qty"]] = round_qty(qty)
            if "price" in roles:
                new_row[roles["price"]] = chosen_price
            if "total" in roles:
                new_row[roles["total"]] = total
            if "nds" in roles:
                new_row[roles["nds"]] = nds
            result_rows.append(new_row)
            stock_info["qty"] -= qty
            logger.info("%s: %.2f → без замены по %.2f (со склада)", code, qty, chosen_price)
            continue
        if not base_code or base_code not in rules:
            logger.info("%s: строка не закрыта (нет правил)", code)
            continue
        target_total = price_src * qty if price_src is not None else None
        remaining_qty = qty
        acc_value = 0.0
        analog_rows: List[List] = []
        for analog_base, rule_price in rules.get(base_code, []):
            for stock_code in stock_by_base.get(analog_base, []):
                info = stock_dict.get(stock_code)
                if info is None or info["qty"] <= 0:
                    continue
                price_choice = rule_price
                price_source = "по правилам" if rule_price is not None else None
                if price_choice is None:
                    price_choice = info["price"]
                    price_source = "из каталога" if price_choice is not None else None
                if price_choice is None:
                    price_choice = price_src
                    price_source = "исходная цена" if price_choice is not None else None
                if price_choice is None:
                    continue
                take_qty = min(info["qty"], remaining_qty)
                if target_total is not None:
                    if acc_value + take_qty*price_choice > target_total*(1+VALUE_EPS):
                        take_qty = max(0, (target_total - acc_value)/price_choice)
                take_qty = round_qty(take_qty)
                if take_qty <= 0:
                    continue
                line_total = round_money(price_choice * take_qty)
                nds = round_money(line_total * VAT_RATE)
                new_row = list(row)
                if "code" in roles:
                    new_row[roles["code"]] = stock_code
                if "name" in roles:
                    new_row[roles["name"]] = profiles.get(stock_code, stock_code)
                if "qty" in roles:
                    new_row[roles["qty"]] = take_qty
                if "price" in roles:
                    new_row[roles["price"]] = round_money(price_choice)
                if "total" in roles:
                    new_row[roles["total"]] = line_total
                if "nds" in roles:
                    new_row[roles["nds"]] = nds
                analog_rows.append(new_row)
                info["qty"] -= take_qty
                remaining_qty = round_qty(remaining_qty - take_qty)
                acc_value = round_money(acc_value + line_total)
                logger.info("%s: %.2f → %s по %.2f (%s)", code, take_qty, stock_code, price_choice, price_source)
                if remaining_qty <= 0:
                    break
            if remaining_qty <= 0:
                break
        if remaining_qty > 0:
            logger.info("%s: строка не закрыта (нехватка аналога)", code)
            for nr in analog_rows:
                scode = to_str_cell(nr[roles["code"]]) if "code" in roles else ""
                qty_back = to_float_cell(nr[roles["qty"]]) if "qty" in roles else 0
                if scode in stock_dict:
                    stock_dict[scode]["qty"] += qty_back
            continue
        if target_total is not None and abs(acc_value - target_total) > target_total*VALUE_EPS:
            logger.info("итог исходной: %.2f ₽; собрано: %.2f ₽ (Δ=%+.2f ₽) — отклонение", target_total, acc_value, acc_value-target_total)
            for nr in analog_rows:
                scode = to_str_cell(nr[roles["code"]]) if "code" in roles else ""
                qty_back = to_float_cell(nr[roles["qty"]]) if "qty" in roles else 0
                if scode in stock_dict:
                    stock_dict[scode]["qty"] += qty_back
            continue
        if target_total is not None:
            logger.info("итог исходной: %.2f ₽; собрано: %.2f ₽ (Δ=%+.2f ₽)", target_total, acc_value, acc_value-target_total)
        result_rows.extend(analog_rows)
    cleaned: List[List] = []
    for r in result_rows:
        if any(to_str_cell(c) != "" for c in r):
            cleaned.append(r)
    result_df = pd.DataFrame(cleaned, columns=headers)
    if "num" in roles:
        for i in range(len(result_df)):
            result_df.iat[i, roles["num"]] = i + 1
    return result_df

# ----------------------------------------------------------------------------
# saving

def save_result(headers: List[str], df: DataFrame, original_invoice_path: str) -> str:
    out_path = os.path.splitext(original_invoice_path)[0] + "_processed.xlsx"
    df_out = pd.DataFrame([headers] + df.values.tolist())
    df_out.to_excel(out_path, header=False, index=False)
    logger.info("Сохранён файл: %s", out_path)
    return out_path

# ----------------------------------------------------------------------------
# GUI

class InvoiceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Invoice App")
        self.stock_path = ""
        self.invoice_path = ""
        self.rules_path = "analogs_priority.xlsx"
        self.profiles_path = "profiles_catalog.xlsx"

        frm = tk.Frame(self.root)
        frm.pack(padx=10, pady=10)
        self.stock_label = tk.Label(frm, text="Остатки: -")
        self.stock_label.pack(anchor="w")
        self.invoice_label = tk.Label(frm, text="Счёт: -")
        self.invoice_label.pack(anchor="w")
        tk.Button(frm, text="Загрузить остатки", command=self.load_stock).pack(fill="x")
        tk.Button(frm, text="Загрузить счёт", command=self.load_invoice).pack(fill="x")
        tk.Button(frm, text="Собрать счёт", command=self.build_invoice).pack(fill="x")
        tk.Button(frm, text="Посмотреть логи", command=self.show_logs).pack(fill="x")
        tk.Button(frm, text="Скачать логи", command=self.save_logs).pack(fill="x")

    def load_stock(self):
        p = filedialog.askopenfilename()
        if p:
            self.stock_path = resolve_existing_path(p)
            self.stock_label.config(text=f"Остатки: {self.stock_path}")

    def load_invoice(self):
        p = filedialog.askopenfilename()
        if p:
            self.invoice_path = resolve_existing_path(p)
            self.invoice_label.config(text=f"Счёт: {self.invoice_path}")

    def build_invoice(self):
        if not self.stock_path or not self.invoice_path:
            messagebox.showwarning("Недостаточно данных", "Загрузите остатки и счёт")
            return
        catalog = load_flexy_catalog()
        profiles = load_profiles_mapping(self.profiles_path)
        rules = load_analog_rules(self.rules_path)
        stock = load_stock(self.stock_path, catalog)
        headers, data, roles = read_invoice_table(self.invoice_path)
        result_df = process_invoice(headers, data, roles, stock, rules, profiles, catalog)
        save_result(headers, result_df, self.invoice_path)
        messagebox.showinfo("Готово", "Файл собран")

    def show_logs(self):
        win = tk.Toplevel(self.root)
        win.title("Логи")
        text = scrolledtext.ScrolledText(win, width=100, height=30)
        text.pack(fill="both", expand=True)
        text.insert("1.0", log_stream.getvalue())
        text.config(state="disabled")

    def save_logs(self):
        p = filedialog.asksaveasfilename(defaultextension=".log")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(log_stream.getvalue())
            messagebox.showinfo("Сохранено", p)

    def run(self):
        self.root.mainloop()

# ----------------------------------------------------------------------------
# CLI

def run_cli(stock_path: str, invoice_path: str, rules_path: str, profiles_path: str) -> None:
    catalog = load_flexy_catalog()
    profiles = load_profiles_mapping(profiles_path)
    rules = load_analog_rules(rules_path)
    stock = load_stock(stock_path, catalog)
    headers, data, roles = read_invoice_table(invoice_path)
    result_df = process_invoice(headers, data, roles, stock, rules, profiles, catalog)
    save_result(headers, result_df, resolve_existing_path(invoice_path))

# ----------------------------------------------------------------------------
# main

def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("stock", nargs="?")
    parser.add_argument("invoice", nargs="?")
    parser.add_argument("--cli", action="store_true")
    parser.add_argument("--rules", default="analogs_priority.xlsx")
    parser.add_argument("--profiles", default="profiles_catalog.xlsx")
    args = parser.parse_args(argv)
    if args.cli:
        if not args.stock or not args.invoice:
            print("Stock and invoice paths required in CLI mode")
            return
        run_cli(args.stock, args.invoice, args.rules, args.profiles)
    else:
        app = InvoiceGUI()
        app.run()

if __name__ == "__main__":
    main()
