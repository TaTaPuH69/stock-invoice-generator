#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Invoice Builder — единый файл с поддержкой:
- GUI через системный Python 2.7 (Tkinter есть) → вызывает python3.11 для расчётов
- CLI через python3.11 (pandas/openpyxl)

Функционал расчёта:
- Структура счёта сохраняется 1:1 (те же имена и порядок колонок; дубликаты ок)
- "Код" всегда артикул (нет кириллицы + содержит \d{4,6}); "Товар" — имя из profiles_catalog.xlsx или сам код
- Если исходного хватает → переносим строку как есть (нормализуем код/товар, пересчитываем Всего/НДС)
- Если не хватает → ПОЛНАЯ замена аналогами по правилам, без смешивания с оригиналом
  Цена: правила → каталог Flexy → исходная цена.
  Подгонка суммы до целевой (±1%) с коррекцией на последнем лоте; шаг по метрам 0.1
- "в т.ч. НДС" = Всего * 20/120
- "№/Номер/No" — автонумерация 1…N (если колонка есть)
- Результат: ~/Downloads/<имя_счёта>_processed.xlsx
"""

import os
import re
import sys
import math
import glob
import logging
import unicodedata
import subprocess

# ----------------------- общие константы -----------------------
VAT_RATE          = 0.20
VALUE_EPS         = 0.01   # ±1%
AMOUNT_ABS_EPS    = 0.01
QTY_STEP          = 0.1
SCAN_HEADER_ROWS  = 80
RULES_DEFAULT_FILE = "analogs_priority.xlsx"
PROFILES_FILE      = "profiles_catalog.xlsx"

PY2 = (sys.version_info[0] == 2)

# ----------------------- логгер -----------------------
logger = logging.getLogger("invoice_app")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s")
    try:
        fh = logging.FileHandler("app.log")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    logger.propagate = False

# ----------------------- общие утилиты (Py2/Py3) -----------------------
try:
    text_type = unicode  # Py2
except NameError:
    text_type = str      # Py3

CODE_NO_CYR  = re.compile(u'^[^\u0400-\u04FF]+$')  # нет кириллицы
BASE_DIGITS6 = re.compile(r'\d{6}')
BASE_DIGITS5 = re.compile(r'\d{5}')
BASE_DIGITS4 = re.compile(r'\d{4}')

def norm_header(h):
    s = u'' if h is None else (h if isinstance(h, text_type) else text_type(h))
    return s.lower().replace(u'\n', u' ').replace(u'\r', u' ').strip()

def to_str_cell(x):
    if x is None:
        return u""
    try:
        s = text_type(x)
    except Exception:
        s = u""
    s = s.strip()
    if s.lower() in (u"nan", u"none", u"null"):
        return u""
    return s

def to_float_cell(x):
    s = to_str_cell(x)
    if s == u"":
        return None
    s = s.replace(u' ', u'').replace(u'\u00A0', u'').replace(u',', u'.')
    try:
        return float(s)
    except Exception:
        return None

def round_money(x):
    return round(float(x), 2)

def round_qty(x):
    if QTY_STEP > 0:
        steps = round(float(x) / QTY_STEP, 4)
        x = steps * QTY_STEP
    return round(float(x), 2)

def looks_like_code_scalar(s):
    txt = to_str_cell(s)
    if txt == u"":
        return False
    if not CODE_NO_CYR.match(txt):
        return False
    return bool(BASE_DIGITS6.search(txt) or BASE_DIGITS5.search(txt) or BASE_DIGITS4.search(txt))

def extract_base_code(s):
    txt = to_str_cell(s)
    m = BASE_DIGITS6.search(txt)
    if m: return m.group(0)
    m = BASE_DIGITS5.search(txt)
    if m: return m.group(0)
    m = BASE_DIGITS4.search(txt)
    if m: return m.group(0)
    return None

def resolve_existing_path(path):
    """Нормализация Unicode + поиск по имени в cwd, папке скрипта, ~/Downloads, ~/Desktop, ~ (неглубокий рекурс)."""
    if not path:
        return path
    cand = [path, os.path.abspath(path)]
    for form in ('NFC', 'NFD', 'NFKC', 'NFKD'):
        try:
            cand.append(unicodedata.normalize(form, path))
        except Exception:
            pass
    for p in cand:
        if os.path.exists(p):
            logger.info('[path] Resolved as-is: {0}'.format(p))
            return p
    base = os.path.basename(path)
    base_lower = base.lower()
    roots = []
    try: roots.append(os.getcwd())
    except Exception: pass
    try: roots.append(os.path.dirname(os.path.abspath(__file__)))
    except Exception: pass
    roots.append(os.path.expanduser('~/Downloads'))
    roots.append(os.path.expanduser('~/Desktop'))
    roots.append(os.path.expanduser('~'))

    for r in roots:
        try:
            for g in glob.glob(os.path.join(r, '*'+base+'*')):
                if os.path.isfile(g):
                    logger.info('[path] Found by glob in {0}: {1}'.format(r, g))
                    return g
        except Exception:
            pass

    for r in roots:
        try:
            for root, dirs, files in os.walk(r):
                if root[len(r):].count(os.sep) > 2:
                    continue
                for f in files:
                    if f.lower() == base_lower:
                        p = os.path.join(root, f)
                        logger.info('[path] Found by walk in {0}: {1}'.format(r, p))
                        return p
        except Exception:
            pass
    logger.warning('[path] Not found: {0}'.format(path))
    return path

# ----------------------- Блок расчёта (используется в Python 3) -----------------------
def run_pipeline_py3(stock_path, invoice_path, rules_path):
    """
    Полный расчёт (pandas/openpyxl). Выполняется под Python 3.x.
    Возвращает финальный путь к результату.
    """
    import pandas as pd
    # flexy loader (по ТЗ — не трогаем)
    import flexy_catalog_loader

    log = []

    # ----- парсер правил -----
    def load_analog_rules(rules_path_in):
        path = resolve_existing_path(rules_path_in or RULES_DEFAULT_FILE)
        if not os.path.exists(path):
            log.append('[rules] Файл не найден: ' + path)
            return {}
        try:
            df = pd.read_excel(path, header=0, dtype=object)
        except Exception as e:
            log.append('[rules] Ошибка чтения: ' + str(e))
            return {}

        cols = list(df.columns)
        cols_low = [norm_header(c) for c in cols]

        base_cands = [u'товар', u'база', u'код', u'артикул', u'base']
        base_idx = None
        for i,n in enumerate(cols_low):
            if any(k in n for k in base_cands):
                base_idx = i; break
        if base_idx is None:
            base_idx = 0

        analog_idxs = [i for i,n in enumerate(cols_low) if (u'аналог' in n or u'analog' in n)]
        price_idxs  = [i for i,n in enumerate(cols_low) if (u'цена' in n or u'price' in n)]

        def num_from(text):
            m = re.search(r'(\d+)', text or u'')
            return m.group(1) if m else None

        analog_num = {i: num_from(cols_low[i]) for i in analog_idxs}
        price_num  = {i: num_from(cols_low[i]) for i in price_idxs}

        price_for = {}
        for ai in analog_idxs:
            n = analog_num.get(ai)
            chosen = None
            if n is not None:
                for pi in price_idxs:
                    if price_num.get(pi) == n:
                        chosen = pi; break
            if chosen is None:
                for off in (1,2,3):
                    pi = ai + off
                    if pi in price_idxs:
                        chosen = pi; break
            price_for[ai] = chosen

        rules = {}
        for _, row in df.iterrows():
            base_code = extract_base_code(row.iat[base_idx] if base_idx < len(row) else None)
            if not base_code:
                continue
            pairs = []
            for ai in analog_idxs:
                aval = row.iat[ai] if ai < len(row) else None
                analog_base = extract_base_code(aval)
                if not analog_base:
                    continue
                pi = price_for.get(ai)
                price_val = None
                if pi is not None and pi < len(row):
                    price_val = to_float_cell(row.iat[pi])
                pairs.append((analog_base, price_val))
            if pairs:
                if base_code not in rules:
                    rules[base_code] = []
                # merge unique
                seen = set()
                merged = []
                for p in rules[base_code] + pairs:
                    key = (p[0], '' if p[1] is None else str(p[1]))
                    if key not in seen:
                        seen.add(key); merged.append(p)
                rules[base_code] = merged
        log.append('[rules] Загружено базовых кодов: {0} ({1})'.format(len(rules), os.path.basename(path)))
        return rules

    # ----- каталог/профили -----
    def load_flexy_catalog():
        try:
            cat = flexy_catalog_loader.load_catalog()
            if 'code' not in cat.columns or 'price_rub' not in cat.columns:
                raise Exception('В каталоге нет code/price_rub')
            cat = cat.copy()
            cat['code'] = cat['code'].astype(str)
            cat['BaseCode'] = cat['code'].map(extract_base_code)
            log.append('[catalog] Каталог Flexy: {0} позиций'.format(len(cat)))
            return cat
        except Exception as e:
            log.append('[catalog] Ошибка загрузки: ' + str(e))
            return pd.DataFrame(columns=['code','price_rub','BaseCode'])

    def load_profiles_mapping():
        path = resolve_existing_path(PROFILES_FILE)
        if not os.path.exists(path):
            log.append('[profiles] profiles_catalog.xlsx не найден — имена = код')
            return {}
        try:
            df = pd.read_excel(path, dtype=object)
            cols = list(df.columns)
            low = [norm_header(c) for c in cols]
            code_idx = None; name_idx = None
            for i,n in enumerate(low):
                if code_idx is None and (u'код' in n or u'артикул' in n or u'sku' in n or u'code' in n):
                    code_idx = i
                if name_idx is None and (u'товар' in n or u'наименование' in n or u'наим' in n or u'product' in n or u'name' in n):
                    name_idx = i
            if code_idx is None or name_idx is None:
                raise Exception('Не найдены колонки кода/наименования')
            mp = {}
            for _, r in df.iterrows():
                c = to_str_cell(r.iat[code_idx])
                nm = to_str_cell(r.iat[name_idx])
                if c:
                    mp[c] = nm or c
            log.append('[profiles] Наименований: {0}'.format(len(mp)))
            return mp
        except Exception as e:
            log.append('[profiles] Ошибка загрузки: ' + str(e) + ' — имена = код')
            return {}

    # ----- остатки -----
    def load_stock(stock_path_in, catalog_df):
        stock_path_in = resolve_existing_path(stock_path_in)
        if not os.path.exists(stock_path_in):
            raise IOError('Файл остатков не найден: ' + stock_path_in)
        raw = pd.read_excel(stock_path_in, header=None)
        articles = raw.iloc[9:, 0]
        qtys     = raw.iloc[9:, 1]
        df = pd.DataFrame({'Артикул': articles, 'Остаток': qtys})
        df = df.dropna(how='all')
        df['Артикул'] = df['Артикул'].astype(str).str.strip()
        df['Остаток'] = pd.to_numeric(df['Остаток'], errors='coerce').fillna(0.0)
        df = df[df['Артикул'] != '']
        df = df.groupby('Артикул', as_index=False)['Остаток'].sum()
        df['BaseCode'] = df['Артикул'].map(extract_base_code)
        if not catalog_df.empty:
            price_map = dict(zip(catalog_df['code'], catalog_df['price_rub']))
            df['price_rub'] = df['Артикул'].map(price_map)
        else:
            df['price_rub'] = pd.NA
        log.append('[stock] Остатков артикулов: {0}'.format(len(df)))
        return df

    # ----- чтение счёта -----
    def is_code_header(h):  n = norm_header(h); return (u'код' in n) or (u'артикул' in n) or (u'sku' in n) or (u'code' in n)
    def is_qty_header(h):   n = norm_header(h); return (u'кол-во' in n) or (u'количество' in n) or (u'qty' in n) or (u'кол.' in n) or (u'кол' in n)
    def is_price_header(h): n = norm_header(h); return (u'цена' in n) or (u'price' in n) or (u'ед.цена' in n) or (u'цена за' in n)
    def is_total_header(h): n = norm_header(h); return (u'всего' in n) or (u'сумма' in n) or (u'итого' in n) or (u'amount' in n) or (u'стоимость' in n)
    def is_nds_header(h):   n = norm_header(h); return (u'в т.ч. ндс' in n) or (u'ндс' in n) or (u'vat' in n)
    def is_name_header(h):  n = norm_header(h); return (u'товар' in n) or (u'наименование' in n) or (u'номенклатура' in n) or (u'product' in n) or (u'item' in n) or (u'наим.' in n)
    def is_unit_header(h):  n = norm_header(h); return (u'ед.' in n) or (u'ед ' in n) or (u'единица' in n) or (u'ед. изм' in n) or (u'ед.изм' in n) or (u'unit' in n)
    def is_num_header(h):   n = norm_header(h); return (u'№' in n) or (u'номер' in n) or (u'no' in n)

    def find_invoice_header_row(raw_df):
        max_rows = min(SCAN_HEADER_ROWS, len(raw_df))
        for i in range(max_rows):
            row = raw_df.iloc[i].tolist()
            names = [to_str_cell(x) for x in row]
            if any(is_code_header(n) for n in names) and any(is_qty_header(n) for n in names):
                log.append('[invoice] Заголовок на строке (1‑based): {0}'.format(i+1))
                return i
        log.append('[invoice] Заголовок не найден — берём первую строку')
        return 0

    def read_invoice_table(invoice_path_in):
        invoice_path_in = resolve_existing_path(invoice_path_in)
        if not os.path.exists(invoice_path_in):
            raise IOError('Счёт не найден: ' + invoice_path_in)
        raw = pd.read_excel(invoice_path_in, header=None, dtype=object)
        header_row = find_invoice_header_row(raw)
        headers = [to_str_cell(x) for x in raw.iloc[header_row].tolist()]
        df = raw.iloc[header_row+1:].copy()
        df.columns = headers
        df = df.dropna(how='all')

        def is_total_row(series):
            for v in series.values:
                s = to_str_cell(v).lower()
                if u'итого' in s and len(s) <= 20:
                    return True
            return False

        table = df[~df.apply(is_total_row, axis=1)].copy()
        columns_order = headers[:]

        roles = {}
        for i, col in enumerate(headers):
            if 'num'   not in roles and is_num_header(col):    roles['num'] = i
            if 'name'  not in roles and is_name_header(col):   roles['name'] = i
            if 'code'  not in roles and is_code_header(col):   roles['code'] = i
            if 'qty'   not in roles and is_qty_header(col):    roles['qty']  = i
            if 'unit'  not in roles and is_unit_header(col):   roles['unit'] = i
            if 'price' not in roles and is_price_header(col):  roles['price']= i
            if 'nds'   not in roles and is_nds_header(col):    roles['nds']  = i
            if 'total' not in roles and is_total_header(col):  roles['total']= i

        log.append('[invoice] Табличных строк: {0}; колонок: {1}'.format(len(table), len(columns_order)))
        return table, columns_order, roles

    # ----- инвентарь -----
    class Inventory(object):
        def __init__(self, stock_df):
            self.by_code = dict(zip(stock_df['Артикул'], stock_df['Остаток']))
            self.price_by_code = dict(zip(stock_df['Артикул'], stock_df['price_rub']))
            self.base_to_codes = {}
            for _, r in stock_df.iterrows():
                bc = r.get('BaseCode')
                art = r.get('Артикул')
                q   = r.get('Остаток')
                if bc:
                    self.base_to_codes.setdefault(bc, []).append((art, float(q or 0.0)))
            for bc in self.base_to_codes:
                self.base_to_codes[bc].sort(key=lambda t: (-t[1], t[0]))

        def get_avail(self, code):
            return float(self.by_code.get(code, 0.0))

        def consume(self, code, qty):
            cur = float(self.by_code.get(code, 0.0))
            self.by_code[code] = max(0.0, cur - float(qty))

        def lots_for_base(self, base_code):
            return list(self.base_to_codes.get(base_code, []))

    # ----- нормализация полей -----
    def normalize_code_and_name_values(vals, roles, name_map):
        code_i = roles.get('code'); name_i = roles.get('name')
        raw_code = to_str_cell(vals[code_i]) if code_i is not None and code_i < len(vals) else u""
        raw_name = to_str_cell(vals[name_i]) if name_i is not None and name_i < len(vals) else u""
        code_val = raw_code if looks_like_code_scalar(raw_code) else u""
        if not code_val and raw_name and looks_like_code_scalar(raw_name):
            code_val = raw_name
        if not code_val:
            for candidate in (raw_code, raw_name):
                if candidate and CODE_NO_CYR.match(candidate) and (BASE_DIGITS6.search(candidate) or BASE_DIGITS5.search(candidate) or BASE_DIGITS4.search(candidate)):
                    code_val = candidate.strip(); break
        if code_i is not None and code_i < len(vals):
            vals[code_i] = code_val
        if name_i is not None and name_i < len(vals):
            vals[name_i] = name_map.get(code_val, code_val)

    def recalc_totals_values(vals, roles):
        qty_i   = roles.get('qty');   price_i = roles.get('price')
        total_i = roles.get('total'); nds_i   = roles.get('nds')
        if qty_i is not None and price_i is not None and total_i is not None:
            q = to_float_cell(vals[qty_i]); p = to_float_cell(vals[price_i])
            if q is not None and p is not None:
                total = round_money(q * p); vals[total_i] = total
                if nds_i is not None:
                    vals[nds_i] = round_money(total * VAT_RATE / (1.0 + VAT_RATE))
        elif nds_i is not None and total_i is not None:
            t = to_float_cell(vals[total_i])
            if t is not None:
                vals[nds_i] = round_money(t * VAT_RATE / (1.0 + VAT_RATE))

    def price_for_code(code, price_from_rule, price_src_row, catalog_price_by_code):
        if price_from_rule is not None:
            return float(price_from_rule)
        p = catalog_price_by_code.get(code)
        if p is not None and not (isinstance(p, float) and pd.isna(p)):
            try: return float(p)
            except Exception: pass
        if price_src_row is not None:
            return float(price_src_row)
        return None

    # ----- обработка счёта -----
    def process_invoice(stock_df, invoice_table, columns_order, roles, rules, catalog_df, name_map):
        inv = Inventory(stock_df)
        result_rows_values = []
        n_cols = len(columns_order)
        catalog_price_by_code = dict(zip(catalog_df['code'], catalog_df['price_rub'])) if not catalog_df.empty else {}

        def base_row_values(src_row):
            vals = []
            for i in range(n_cols):
                col = columns_order[i]
                try:
                    vals.append(src_row.get(col))
                except Exception:
                    vals.append(None)
            return vals

        def append_if_not_empty(vals):
            for v in vals:
                if to_str_cell(v) != u"":
                    result_rows_values.append(vals); return

        def add_row_from_source(src_row):
            vals = base_row_values(src_row)
            normalize_code_and_name_values(vals, roles, name_map)
            recalc_totals_values(vals, roles)
            append_if_not_empty(vals)

        def add_row_for_analog(src_row, code, qty, price):
            vals = base_row_values(src_row)
            code_i  = roles.get('code'); name_i  = roles.get('name')
            qty_i   = roles.get('qty');  price_i = roles.get('price')
            total_i = roles.get('total'); nds_i  = roles.get('nds')

            if code_i is not None and code_i < n_cols:  vals[code_i] = code
            if name_i is not None and name_i < n_cols:  vals[name_i] = name_map.get(code, code)
            if qty_i  is not None and qty_i  < n_cols:  vals[qty_i]  = round_qty(qty)
            if price_i is not None and price_i < n_cols: vals[price_i] = round_money(price)
            if total_i is not None and price_i is not None and qty_i is not None:
                total = round_money(round_qty(qty) * round_money(price))
                vals[total_i] = total
                if nds_i is not None and nds_i < n_cols:
                    vals[nds_i] = round_money(total * VAT_RATE / (1.0 + VAT_RATE))
            elif nds_i is not None and total_i is not None:
                t = to_float_cell(vals[total_i])
                if t is not None:
                    vals[nds_i] = round_money(t * VAT_RATE / (1.0 + VAT_RATE))
            append_if_not_empty(vals)

        code_i  = roles.get('code'); name_i = roles.get('name')
        qty_i   = roles.get('qty');  price_i= roles.get('price')

        for _, src_row in invoice_table.iterrows():
            if code_i is None or qty_i is None:
                continue

            art = to_str_cell(src_row.iat[code_i]) if code_i < len(src_row) else u""
            if not looks_like_code_scalar(art):
                alt = to_str_cell(src_row.iat[name_i]) if name_i is not None else u""
                if looks_like_code_scalar(alt):
                    art = alt

            qty_src   = to_float_cell(src_row.iat[qty_i]) if qty_i < len(src_row) else None
            price_src = to_float_cell(src_row.iat[price_i]) if (price_i is not None and price_i < len(src_row)) else None
            if not art or qty_src is None or qty_src <= 0:
                continue

            if inv.get_avail(art) >= qty_src:
                add_row_from_source(src_row)
                inv.consume(art, qty_src)
                continue

            base_code = extract_base_code(art)
            if not base_code or base_code not in rules:
                log.append(u'{0}: аналогов нет по правилам — строка не закрыта'.format(art)); continue

            lots = []  # (код, avail, price, origin)
            for analog_base, price_from_rule in rules[base_code]:
                base_lots = inv.lots_for_base(analog_base)
                if not base_lots: continue
                for lot_code, lot_avail in base_lots:
                    if lot_avail <= 0: continue
                    p = price_for_code(lot_code, price_from_rule, price_src, inv.price_by_code if inv.price_by_code else catalog_price_by_code)
                    if p is None:
                        p = price_for_code(lot_code, price_from_rule, price_src, catalog_price_by_code)
                    origin = u'по правилам' if price_from_rule is not None else (u'из каталога' if (lot_code in (inv.price_by_code or {}) or lot_code in catalog_price_by_code) else u'исходная цена')
                    lots.append((lot_code, float(lot_avail), (None if p is None else float(p)), origin))

            if not lots:
                log.append(u'{0}: аналогов нет на складе — строка не закрыта'.format(art)); continue
            if price_src is None and all(l[2] is None for l in lots):
                log.append(u'{0}: не удалось вычислить цену — строка не закрыта'.format(art)); continue

            created_any = False

            if price_src is not None:
                target_total = round_money(qty_src * price_src)
                allowed_delta = max(AMOUNT_ABS_EPS, abs(target_total) * VALUE_EPS)
                acc_value = 0.0
                used = []  # (i, q)

                for i, (c, avail_i, p_i, origin_i) in enumerate(lots):
                    if p_i is None or p_i <= 0: continue
                    remaining_value = target_total - acc_value
                    if remaining_value <= allowed_delta: break
                    q_desired = 0.0 if p_i == 0 else (remaining_value / p_i)
                    q_take = min(avail_i, max(0.0, q_desired))
                    q_take = round_qty(q_take)
                    if q_take <= 0: continue
                    acc_value = round_money(acc_value + q_take * p_i)
                    used.append((i, q_take))

                if used:
                    last_i, last_q = used[-1]
                    c_last, avail_last, p_last, origin_last = lots[last_i]
                    diff = target_total - acc_value
                    if abs(diff) > allowed_delta and p_last and p_last > 0:
                        dq = diff / p_last
                        new_q = min(avail_last, max(0.0, last_q + dq))
                        new_q = round_qty(new_q)
                        acc_value = round_money(acc_value - last_q * p_last + new_q * p_last)
                        used[-1] = (last_i, new_q)

                if used and abs(target_total - acc_value) <= allowed_delta:
                    for i, q in used:
                        c_i, avail_i, p_i, origin_i = lots[i]
                        add_row_for_analog(src_row, c_i, q, p_i)
                        inv.consume(c_i, q)
                        log.append(u'{0}: {1:.2f} → {2} по {3:.2f} ₽ ({4})'.format(art, q, c_i, p_i, origin_i))
                        created_any = True
                    log.append(u'итог исходной: {0:.2f} ₽; собрано: {1:.2f} ₽ (Δ={2:+.2f} ₽)'.format(target_total, acc_value, target_total - acc_value))
                else:
                    log.append(u'{0}: не удалось уложиться в допуск по сумме — строка не закрыта'.format(art))

            else:
                remaining_qty = round_qty(qty_src)
                used_q = []
                for i, (c, avail_i, p_i, origin_i) in enumerate(lots):
                    if remaining_qty <= 0: break
                    if p_i is None:
                        used_q = []; log.append(u'{0}: нет цены у аналога {1} — строка не закрыта'.format(art, c)); break
                    q_take = min(avail_i, remaining_qty)
                    q_take = round_qty(q_take)
                    if q_take <= 0: continue
                    used_q.append((i, q_take))
                    remaining_qty = round_qty(remaining_qty - q_take)

                if used_q and remaining_qty <= 0:
                    for i, q in used_q:
                        c_i, avail_i, p_i, origin_i = lots[i]
                        add_row_for_analog(src_row, c_i, q, p_i)
                        inv.consume(c_i, q)
                        log.append(u'{0}: {1:.2f} → {2} по {3:.2f} ₽ ({4})'.format(art, q, c_i, p_i, origin_i))
                        created_any = True
                else:
                    log.append(u'{0}: аналогов недостаточно — строка не закрыта'.format(art))

            if not created_any:
                continue

        result_df = pd.DataFrame(result_rows_values, columns=columns_order)
        if not result_df.empty:
            mask = result_df.apply(lambda r: any(to_str_cell(v) != u'' for v in r), axis=1)
            result_df = result_df[mask]

        num_i = roles.get('num')
        if num_i is not None and not result_df.empty:
            result_df.iloc[:, num_i] = list(range(1, len(result_df) + 1))

        return result_df, log

    # ------ запуск пайплайна ------
    rules   = load_analog_rules(rules_path)
    catalog = load_flexy_catalog()
    profiles= load_profiles_mapping()
    stock   = load_stock(stock_path, catalog)
    inv_tbl, cols_order, roles = read_invoice_table(invoice_path)
    result_df, log = process_invoice(stock, inv_tbl, cols_order, roles, rules, catalog, profiles)

    base = os.path.splitext(os.path.basename(invoice_path))[0]
    out_name = base + "_processed.xlsx"
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    try:
        if not os.path.exists(downloads):
            os.makedirs(downloads)
    except Exception:
        pass
    out_path = os.path.join(downloads, out_name)
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        result_df.to_excel(w, index=False)

    for line in log:
        try:
            sys.stdout.write((line + "\n").encode('utf-8') if PY2 else (line + "\n"))
        except Exception:
            pass
    sys.stdout.write(("\n[save] Результат сохранён: {0}\n".format(out_path)).encode('utf-8') if PY2 else "\n[save] Результат сохранён: {0}\n".format(out_path))
    return out_path

# ----------------------- GUI-обёртка под Python 2.7 -----------------------
if PY2:
    try:
        from Tkinter import Tk, Text, Scrollbar, Button, END, Toplevel
        import tkFileDialog as filedialog
        import tkMessageBox as messagebox
    except Exception:
        Tk = None

    class Py2GUI(object):
        def __init__(self):
            if Tk is None:
                sys.stderr.write("GUI недоступен (Tkinter отсутствует). Запустите в CLI через python3.11.\n")
                sys.exit(2)
            self.root = Tk()
            self.root.title("Invoice Builder (GUI via Py2 → calc via Py3)")
            self.stock_path = None
            self.invoice_path = None
            self.rules_path = RULES_DEFAULT_FILE

            self.txt = Text(self.root, height=22, width=100)
            sc = Scrollbar(self.root, command=self.txt.yview)
            self.txt.configure(yscrollcommand=sc.set)
            self.txt.pack(side='left', fill='both', expand=True)
            sc.pack(side='right', fill='y')

            Button(self.root, text="Загрузить остатки", command=self.pick_stock).pack()
            Button(self.root, text="Загрузить счёт",   command=self.pick_invoice).pack()
            Button(self.root, text="Загрузить правила",command=self.pick_rules).pack()
            Button(self.root, text="Собрать счёт",     command=self.build).pack()

        def log(self, msg):
            try:
                self.txt.insert(END, msg + u"\n"); self.txt.see(END)
            except Exception:
                pass
            try:
                logger.info(msg)
            except Exception:
                pass

        def pick_stock(self):
            p = filedialog.askopenfilename(title="Выберите файл остатков", filetypes=[("Excel","*.xlsx *.xls")])
            if p:
                self.stock_path = p; self.log(u"[ui] Остатки: " + p)

        def pick_invoice(self):
            p = filedialog.askopenfilename(title="Выберите файл счёта", filetypes=[("Excel","*.xlsx *.xls")])
            if p:
                self.invoice_path = p; self.log(u"[ui] Счёт: " + p)

        def pick_rules(self):
            p = filedialog.askopenfilename(title="Правила analogs_priority.xlsx", filetypes=[("Excel","*.xlsx *.xls")])
            if p:
                self.rules_path = p; self.log(u"[ui] Правила: " + p)

        def find_py3(self):
            # попробуем python3.11, затем python3
            for exe in ("python3.11","python3"):
                for path in os.environ.get("PATH","").split(os.pathsep):
                    full = os.path.join(path, exe)
                    if os.path.isfile(full) and os.access(full, os.X_OK):
                        return full
            return None

        def build(self):
            if not self.stock_path or not self.invoice_path:
                messagebox.showerror("Ошибка", "Загрузите остатки и счёт"); return
            py3 = self.find_py3()
            if not py3:
                messagebox.showerror("Ошибка", "Не найден python3.11/python3 в PATH"); return
            cmd = [
                py3, os.path.abspath(__file__),
                "--cli", self.stock_path, self.invoice_path,
                "--rules", self.rules_path
            ]
            self.log(u"[run] " + u" ".join(cmd))
            try:
                p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                out = p.communicate()[0]
                try:
                    out = out.decode('utf-8', 'ignore')
                except Exception:
                    pass
                self.log(out)
                if p.returncode == 0:
                    messagebox.showinfo("Готово", "Проверьте ~/Downloads — файл *_processed.xlsx")
                else:
                    messagebox.showerror("Ошибка", "Команда завершилась с кодом {0}".format(p.returncode))
            except Exception as e:
                messagebox.showerror("Ошибка запуска", unicode(e))

        def run(self):
            self.root.mainloop()

# ----------------------- main -----------------------
def main():
    if PY2:
        # GUI‑обёртка (Py2 → Py3)
        app = Py2GUI()
        app.run()
        return

    # Python 3.x → CLI режим
    import argparse
    parser = argparse.ArgumentParser(description='Invoice Builder (CLI via Python 3)')
    parser.add_argument('--cli', nargs=2, metavar=('STOCK_XLSX','INVOICE_XLSX'), help='запуск расчёта')
    parser.add_argument('--rules', default=RULES_DEFAULT_FILE, help='путь к analogs_priority.xlsx')
    args = parser.parse_args()

    if not args.cli:
        sys.stderr.write("GUI недоступен под этим интерпретатором. Запустите: python invoice_app.py (Py2 GUI) или используйте CLI:\n"
                         "python3.11 invoice_app.py --cli \"/path/остатки.xlsx\" \"/path/счет.xlsx\" --rules analogs_priority.xlsx\n")
        sys.exit(2)

    stock_path, invoice_path = args.cli
    rules_path = args.rules or RULES_DEFAULT_FILE
    outp = run_pipeline_py3(stock_path, invoice_path, rules_path)
    print("\n✅ Готово: {0}".format(outp))

if __name__ == '__main__':
    if PY2:
        try:
            # для корректного вывода юникода в Py2
            reload(sys)
            sys.setdefaultencoding('utf-8')
        except Exception:
            pass
    main()
