#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Invoice building utility (GUI/CLI)."""

from __future__ import annotations

import os
import re
import sys
import glob
import argparse
import unicodedata
from pathlib import Path
from typing import List, Dict, Optional, Tuple

import pandas as pd
from flexy_catalog_loader import load_catalog

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:  # pragma: no cover - tkinter may be absent
    tk = None  # type: ignore

# ---------------------------------------------------------------------------
# constants
VALUE_EPS = 0.01
QTY_STEP = 0.1
VAT_RATE = 0.20
RULES_DEFAULT_FILE = "analogs_priority.xlsx"
PROFILES_FILE = "profiles_catalog.xlsx"
SCAN_HEADER_ROWS = 80

CODE_NO_CYR = re.compile(r"^[^\u0400-\u04FF]+$")
BASE_DIGITS = re.compile(r"\d{4,6}")

# global log
GLOBAL_LOG: List[str] = []

# ---------------------------------------------------------------------------
# logging helper

def log_message(msg: str, log: Optional[List[str]] = None) -> None:
    if log is None:
        log = GLOBAL_LOG
    log.append(msg)
    print(msg)

# ---------------------------------------------------------------------------
# helpers

def resolve_existing_path(path: str) -> str:
    if not path:
        return path
    orig = path
    candidates = [path, os.path.abspath(path)]
    for norm in ("NFC", "NFD", "NFKC", "NFKD"):
        np = unicodedata.normalize(norm, path)
        candidates.append(np)
        candidates.append(os.path.abspath(np))
    for c in candidates:
        if os.path.exists(c):
            return os.path.abspath(c)
    basename = os.path.basename(path)
    search_dirs = []
    cwd = os.getcwd()
    script_dir = os.path.dirname(__file__)
    search_dirs.extend([cwd, script_dir, os.path.dirname(cwd), os.path.dirname(script_dir)])
    home = str(Path.home())
    search_dirs.extend([os.path.join(home, "Downloads"), os.path.join(home, "Desktop"), home])
    seen: set[str] = set()
    for base in search_dirs:
        base = os.path.abspath(base)
        if not os.path.isdir(base) or base in seen:
            continue
        seen.add(base)
        for root, dirs, files in os.walk(base):
            if root.startswith("/System") or root.startswith("/Library"):
                dirs[:] = []
                continue
            for f in files:
                if f.lower() == basename.lower():
                    return os.path.join(root, f)
            pattern = os.path.join(root, f"*{basename}*")
            for p in glob.glob(pattern):
                if os.path.isfile(p):
                    return p
    return orig

def to_str_cell(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_float_cell(x) -> Optional[float]:
    if pd.isna(x) or x in ("", None):
        return None
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return None

def round_qty(x: float) -> float:
    if x is None:
        return 0.0
    q = max(0.0, x)
    return round(round(q / QTY_STEP) * QTY_STEP, 2)

def round_money(x: float) -> float:
    return round(float(x or 0.0), 2)

def looks_like_code_scalar(s: str) -> bool:
    s = to_str_cell(s)
    if not s:
        return False
    return bool(CODE_NO_CYR.match(s) and BASE_DIGITS.search(s))

def extract_base_code(s: str) -> Optional[str]:
    s = to_str_cell(s)
    if not s:
        return None
    for n in (6, 5, 4):
        m = re.search(fr"\d{{{n}}}", s)
        if m:
            return m.group(0)
    return None

# ---------------------------------------------------------------------------
# loading helpers

def load_analog_rules(rules_path: str) -> Dict[str, List[Tuple[str, Optional[float]]]]:
    path = resolve_existing_path(rules_path)
    if os.path.exists(path):
        log_message(f"[{rules_path}] resolved to: {path}")
    else:
        log_message(f"[{rules_path}] not found")
        return {}
    df = pd.read_excel(path, dtype=str)
    if df.empty:
        return {}
    cols = list(df.columns)
    def norm(s: str) -> str:
        return s.lower().strip()
    base_candidates = ["товар", "база", "код", "артикул", "base"]
    base_col = cols[0]
    for c in cols:
        if any(k in norm(c) for k in base_candidates):
            base_col = c
            break
    analog_cols: List[str] = []
    price_cols: List[str] = []
    for c in cols:
        n = norm(c)
        if "аналог" in n or "analog" in n:
            analog_cols.append(c)
        if "цена" in n or "price" in n:
            price_cols.append(c)
    def col_num(s: str) -> Optional[str]:
        m = re.search(r"(\d+)", s)
        return m.group(1) if m else None
    price_by_num = {col_num(c): c for c in price_cols}
    info: List[Tuple[str, Optional[str]]] = []
    for ac in analog_cols:
        num = col_num(ac)
        pcol = None
        if num and num in price_by_num:
            pcol = price_by_num[num]
        else:
            idx = cols.index(ac)
            for j in range(1, 4):
                if idx + j < len(cols):
                    cand = cols[idx + j]
                    nn = norm(cand)
                    if "цена" in nn or "price" in nn:
                        pcol = cand
                        break
        info.append((ac, pcol))
    rules: Dict[str, List[Tuple[str, Optional[float]]]] = {}
    for _, r in df.iterrows():
        base = extract_base_code(r.get(base_col, ""))
        if not base:
            continue
        lst: List[Tuple[str, Optional[float]]] = []
        for ac, pcol in info:
            analog = extract_base_code(r.get(ac, ""))
            if not analog:
                continue
            price = None
            if pcol:
                price = to_float_cell(r.get(pcol))
            lst.append((analog, price))
        if lst:
            rules.setdefault(base, []).extend(lst)
    return rules

def load_flexy_catalog() -> pd.DataFrame:
    cat = load_catalog()
    cat = cat.copy()
    cat["code"] = cat["code"].astype(str).str.strip()
    cat["BaseCode"] = cat["code"].map(extract_base_code)
    return cat

def load_profiles_mapping() -> Dict[str, str]:
    path = resolve_existing_path(PROFILES_FILE)
    if os.path.exists(path):
        log_message(f"[{PROFILES_FILE}] resolved to: {path}")
    else:
        log_message(f"[{PROFILES_FILE}] not found")
        return {}
    df = pd.read_excel(path, dtype=str)
    def norm(s: str) -> str:
        return s.lower().strip()
    code_names = ["код", "артикул", "sku", "code"]
    name_names = ["товар", "наименование", "наим", "product", "name"]
    code_col = df.columns[0]
    name_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    for c in df.columns:
        if any(n == norm(c) for n in code_names):
            code_col = c
        if any(n == norm(c) for n in name_names):
            name_col = c
    mapping: Dict[str, str] = {}
    for _, row in df.iterrows():
        code = to_str_cell(row.get(code_col))
        name = to_str_cell(row.get(name_col))
        if code:
            mapping[code] = name or code
    return mapping

def load_stock(stock_path: str, catalog: pd.DataFrame, log: List[str]) -> pd.DataFrame:
    df = pd.read_excel(stock_path, header=None, skiprows=9, usecols=[0, 1])
    df.columns = ["Code", "Stock"]
    df["Code"] = df["Code"].astype(str).str.strip()
    df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0.0)
    df = df.groupby("Code", as_index=False).agg({"Stock": "sum"})
    df["BaseCode"] = df["Code"].map(extract_base_code)
    cat_prices = catalog[["code", "price_rub"]].rename(columns={"code": "Code"})
    df = df.merge(cat_prices, on="Code", how="left")
    log_message(f"stock rows: {len(df)}", log)
    return df

# ---------------------------------------------------------------------------
# invoice reading
CODE_KEYS = ["код", "артикул", "sku", "code", "код/артикул", "артикул/код"]
QTY_KEYS = ["кол-во", "количество", "кол.", "qty", "quantity", "кол"]
PRICE_KEYS = ["цена", "price", "ед.цена", "цена за"]
TOTAL_KEYS = ["всего", "сумма", "итого", "amount", "стоимость"]
NDS_KEYS = ["в т.ч. ндс", "ндс", "vat"]
NAME_KEYS = ["товар", "наименование", "номенклатура", "product", "item", "наим."]
UNIT_KEYS = ["ед.", "ед", "единица", "ед. изм", "ед.изм", "unit"]
NUM_KEYS = ["№", "номер", "no"]


def find_header_row(df_raw: pd.DataFrame, log: List[str]) -> int:
    max_row = min(SCAN_HEADER_ROWS, len(df_raw))
    for i in range(max_row):
        row = df_raw.iloc[i].astype(str).str.lower()
        has_code = any(any(k in to_str_cell(c) for k in CODE_KEYS) for c in row)
        has_qty = any(any(k in to_str_cell(c) for k in QTY_KEYS) for c in row)
        if has_code and has_qty:
            log_message(f"header row: {i}", log)
            return i
    log_message("header row not found, using 0", log)
    return 0

def read_invoice_table(invoice_path: str, log: List[str]) -> Tuple[pd.DataFrame, List[str], Dict[str, int]]:
    df_raw = pd.read_excel(invoice_path, header=None, dtype=object)
    hrow = find_header_row(df_raw, log)
    headers = df_raw.iloc[hrow].tolist()
    df = df_raw.iloc[hrow + 1 :].copy()
    df.columns = headers
    roles: Dict[str, int] = {}
    for idx, h in enumerate(headers):
        n = to_str_cell(h).lower()
        if any(k == n or k in n for k in CODE_KEYS):
            roles["code"] = idx
        if any(k == n or k in n for k in QTY_KEYS):
            roles["qty"] = idx
        if any(k == n or k in n for k in PRICE_KEYS):
            roles["price"] = idx
        if any(k == n or k in n for k in TOTAL_KEYS):
            roles["total"] = idx
        if any(k == n or k in n for k in NDS_KEYS):
            roles["nds"] = idx
        if any(k == n or k in n for k in NAME_KEYS):
            roles["name"] = idx
        if any(k == n or k in n for k in UNIT_KEYS):
            roles["unit"] = idx
        if any(k == n or k in n for k in NUM_KEYS):
            roles["num"] = idx
    return df.reset_index(drop=True), headers, roles

# ---------------------------------------------------------------------------
# inventory
class Inventory:
    def __init__(self, df: pd.DataFrame):
        self.df = df.set_index("Code")

    def get_avail(self, code: str) -> float:
        if code in self.df.index:
            return float(self.df.at[code, "Stock"])
        return 0.0

    def get_price(self, code: str) -> Optional[float]:
        if code in self.df.index:
            return to_float_cell(self.df.at[code, "price_rub"])
        return None

    def consume(self, code: str, qty: float) -> None:
        if code in self.df.index:
            cur = float(self.df.at[code, "Stock"])
            self.df.at[code, "Stock"] = max(0.0, cur - qty)

    def get_by_base(self, base: str) -> List[Dict[str, float]]:
        rows = self.df[self.df["BaseCode"] == base]
        res: List[Dict[str, float]] = []
        for code, r in rows.iterrows():
            res.append({"code": code, "stock": float(r["Stock"]), "price": to_float_cell(r.get("price_rub"))})
        return res

# ---------------------------------------------------------------------------
# processing

def process_invoice(df: pd.DataFrame, headers: List[str], roles: Dict[str, int], inventory: Inventory,
                    catalog: pd.DataFrame, profiles: Dict[str, str], rules: Dict[str, List[Tuple[str, Optional[float]]]],
                    log: List[str]) -> pd.DataFrame:
    result_rows: List[List] = []
    for idx in range(len(df)):
        row = df.iloc[idx]
        art_idx = roles.get("code")
        qty_idx = roles.get("qty")
        price_idx = roles.get("price")
        name_idx = roles.get("name")
        total_idx = roles.get("total")
        nds_idx = roles.get("nds")
        art = to_str_cell(row.iat[art_idx]) if art_idx is not None else ""
        qty_src = to_float_cell(row.iat[qty_idx]) if qty_idx is not None else None
        price_src = to_float_cell(row.iat[price_idx]) if price_idx is not None else None
        name_val = to_str_cell(row.iat[name_idx]) if name_idx is not None else ""
        if qty_src is None or qty_src <= 0:
            log_message("строка пропущена: нет количества", log)
            continue
        if not looks_like_code_scalar(art):
            m = re.search(r"[A-Za-z0-9]*\d{4,6}[A-Za-z0-9]*", name_val)
            cand = m.group(0) if m else ""
            if looks_like_code_scalar(cand):
                art = cand
            else:
                log_message("строка пропущена: не распознан код", log)
                continue
        base_code = extract_base_code(art)
        if not base_code:
            log_message(f"{art}: нет базового кода", log)
            continue
        if price_src is None:
            price_src = to_float_cell(catalog.loc[catalog["code"] == art, "price_rub"].head(1))
        avail = inventory.get_avail(art)
        if avail >= qty_src and price_src is not None:
            q = round_qty(qty_src)
            price = round_money(price_src)
            total = round_money(q * price)
            nds = round_money(total * VAT_RATE / (1 + VAT_RATE))
            inventory.consume(art, q)
            row_new = row.tolist()
            if name_idx is not None:
                row_new[name_idx] = profiles.get(art, profiles.get(base_code, art))
            if art_idx is not None:
                row_new[art_idx] = art
            if qty_idx is not None:
                row_new[qty_idx] = q
            if price_idx is not None:
                row_new[price_idx] = price
            if total_idx is not None:
                row_new[total_idx] = total
            if nds_idx is not None:
                row_new[nds_idx] = nds
            result_rows.append(row_new)
            continue
        rule_list = rules.get(base_code)
        if not rule_list:
            log_message(f"{art}: нет правил аналогов", log)
            continue
        lots: List[Dict[str, object]] = []
        for analog_base, rule_price in rule_list:
            for lot in inventory.get_by_base(analog_base):
                if lot["stock"] <= 0:
                    continue
                price = rule_price if rule_price is not None else lot.get("price")
                source = (
                    "по правилам" if rule_price is not None else
                    "из каталога" if lot.get("price") is not None else
                    "исходная цена"
                )
                if price is None:
                    if price_src is not None:
                        price = price_src
                        source = "исходная цена"
                    else:
                        continue
                lots.append({"code": lot["code"], "avail": lot["stock"], "price": price, "src": source})
        if not lots:
            log_message(f"{art}: нет доступных аналогов", log)
            continue
        local_rows: List[List] = []
        if price_src is not None:
            target = qty_src * price_src
            acc_value = 0.0
            for i, lot in enumerate(lots):
                p = lot["price"]
                if i < len(lots) - 1:
                    q = max(0.0, (target - acc_value) / p)
                    q = min(q, lot["avail"])
                    q = round_qty(q)
                    acc_value += round_money(q * p)
                else:
                    q = max(0.0, (target - acc_value) / p)
                    q = min(q, lot["avail"])
                    q = round_qty(q)
                    acc_value += round_money(q * p)
                    diff = target - acc_value
                    dq = diff / p
                    q = max(0.0, min(lot["avail"], q + dq))
                    q = round_qty(q)
                    acc_value = round_money(acc_value + round_money(dq * p))
                if q > 0:
                    inventory.consume(lot["code"], q)
                    name = profiles.get(lot["code"], profiles.get(extract_base_code(lot["code"]) or "", lot["code"]))
                    new_row = ["" for _ in headers]
                    if art_idx is not None:
                        new_row[art_idx] = lot["code"]
                    if name_idx is not None:
                        new_row[name_idx] = name
                    if qty_idx is not None:
                        new_row[qty_idx] = q
                    if price_idx is not None:
                        new_row[price_idx] = round_money(p)
                    total = round_money(q * p)
                    if total_idx is not None:
                        new_row[total_idx] = total
                    nds = round_money(total * VAT_RATE / (1 + VAT_RATE))
                    if nds_idx is not None:
                        new_row[nds_idx] = nds
                    local_rows.append(new_row)
                    log_message(f"{art}: {q:.2f} → {lot['code']} по {p:.2f} ₽ ({lot['src']})", log)
            diff = target - sum(round_money(r[price_idx] * r[qty_idx]) if price_idx is not None and qty_idx is not None else 0 for r in local_rows)
            tol = max(target * 0.01, VALUE_EPS)
            if abs(diff) > tol:
                log_message(f"{art}: не попали в допуск суммы", log)
                continue
            log_message(f"итог исходной: {target:.2f} ₽; собрано: {target - diff:.2f} ₽ (Δ={-diff:+.2f} ₽)", log)
            result_rows.extend(local_rows)
        else:
            remaining = qty_src
            for lot in lots:
                if remaining <= 0:
                    break
                q = min(lot["avail"], remaining)
                q = round_qty(q)
                if q <= 0:
                    continue
                remaining -= q
                inventory.consume(lot["code"], q)
                name = profiles.get(lot["code"], profiles.get(extract_base_code(lot["code"]) or "", lot["code"]))
                new_row = ["" for _ in headers]
                if art_idx is not None:
                    new_row[art_idx] = lot["code"]
                if name_idx is not None:
                    new_row[name_idx] = name
                if qty_idx is not None:
                    new_row[qty_idx] = q
                price = round_money(lot["price"])
                if price_idx is not None:
                    new_row[price_idx] = price
                total = round_money(q * price)
                if total_idx is not None:
                    new_row[total_idx] = total
                nds = round_money(total * VAT_RATE / (1 + VAT_RATE))
                if nds_idx is not None:
                    new_row[nds_idx] = nds
                result_rows.append(new_row)
                log_message(f"{art}: {q:.2f} → {lot['code']} по {price:.2f} ₽ ({lot['src']})", log)
            if remaining > 0:
                log_message(f"{art}: не хватило количества", log)
                continue
    result = pd.DataFrame(result_rows, columns=headers)
    num_idx = roles.get("num")
    if num_idx is not None:
        result.iloc[:, num_idx] = range(1, len(result) + 1)
    return result

# ---------------------------------------------------------------------------
# saving

def save_result(df: pd.DataFrame, input_invoice_path: str, log: List[str]) -> str:
    out_path = Path(input_invoice_path).with_name(Path(input_invoice_path).stem + "_result.xlsx")
    df.to_excel(out_path, index=False)
    log_message(f"result saved: {out_path}", log)
    return str(out_path)

# ---------------------------------------------------------------------------
# CLI

def run_cli(stock_path: str, invoice_path: str, rules_path: str, log: Optional[List[str]] = None) -> Optional[str]:
    if log is None:
        log = GLOBAL_LOG
    stock_path_res = resolve_existing_path(stock_path)
    log_message(f"[{stock_path}] resolved to: {stock_path_res}" if os.path.exists(stock_path_res) else f"[{stock_path}] not found", log)
    invoice_path_res = resolve_existing_path(invoice_path)
    log_message(f"[{invoice_path}] resolved to: {invoice_path_res}" if os.path.exists(invoice_path_res) else f"[{invoice_path}] not found", log)
    rules_path_res = resolve_existing_path(rules_path)
    log_message(f"[{rules_path}] resolved to: {rules_path_res}" if os.path.exists(rules_path_res) else f"[{rules_path}] not found", log)
    catalog = load_flexy_catalog()
    stock_df = load_stock(stock_path_res, catalog, log)
    inventory = Inventory(stock_df)
    rules = load_analog_rules(rules_path_res)
    profiles = load_profiles_mapping()
    inv_df, headers, roles = read_invoice_table(invoice_path_res, log)
    result_df = process_invoice(inv_df, headers, roles, inventory, catalog, profiles, rules, log)
    return save_result(result_df, invoice_path_res, log)

# ---------------------------------------------------------------------------
# GUI
class InvoiceGUI:
    def __init__(self) -> None:
        if tk is None:
            raise RuntimeError("tkinter not available")
        self.root = tk.Tk()
        self.root.title("Invoice Builder")
        self.log: List[str] = GLOBAL_LOG
        self.stock_path: Optional[str] = None
        self.invoice_path: Optional[str] = None
        self.rules_path: str = RULES_DEFAULT_FILE
        frm = tk.Frame(self.root)
        frm.pack(padx=10, pady=10)
        tk.Button(frm, text="Загрузить остатки", command=self.load_stock).grid(row=0, column=0, sticky="ew")
        tk.Button(frm, text="Загрузить счёт", command=self.load_invoice).grid(row=0, column=1, sticky="ew")
        tk.Button(frm, text="Собрать счёт", command=self.process).grid(row=0, column=2, sticky="ew")
        tk.Button(frm, text="Посмотреть логи", command=self.show_logs).grid(row=0, column=3, sticky="ew")
        tk.Button(frm, text="Скачать логи", command=self.save_logs).grid(row=0, column=4, sticky="ew")
        self.text = tk.Text(self.root, width=100, height=20)
        self.text.pack(fill="both", expand=True)

    def load_stock(self) -> None:
        path = filedialog.askopenfilename(title="Остатки")
        if path:
            res = resolve_existing_path(path)
            self.stock_path = res
            log_message(f"[{path}] resolved to: {res}" if os.path.exists(res) else f"[{path}] not found", self.log)
            self.update_log_widget()

    def load_invoice(self) -> None:
        path = filedialog.askopenfilename(title="Счёт")
        if path:
            res = resolve_existing_path(path)
            self.invoice_path = res
            log_message(f"[{path}] resolved to: {res}" if os.path.exists(res) else f"[{path}] not found", self.log)
            self.update_log_widget()

    def process(self) -> None:
        if not self.stock_path or not self.invoice_path:
            messagebox.showerror("Ошибка", "Загрузите файлы остатков и счёта")
            return
        out = run_cli(self.stock_path, self.invoice_path, self.rules_path, self.log)
        self.update_log_widget()
        if out:
            messagebox.showinfo("Готово", f"Файл сохранён: {out}")

    def show_logs(self) -> None:
        self.update_log_widget()

    def save_logs(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".txt", title="Сохранить логи")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self.log))

    def update_log_widget(self) -> None:
        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, "\n".join(self.log))

    def run(self) -> None:
        self.root.mainloop()

# ---------------------------------------------------------------------------
# main

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cli", nargs=2, metavar=("STOCK", "INVOICE"))
    parser.add_argument("--rules", default=RULES_DEFAULT_FILE)
    args = parser.parse_args()
    if args.cli:
        run_cli(args.cli[0], args.cli[1], args.rules, GLOBAL_LOG)
    else:
        if tk is None:
            print("tkinter is not available", file=sys.stderr)
            return
        gui = InvoiceGUI()
        gui.run()

if __name__ == "__main__":
    main()
