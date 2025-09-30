#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Invoice Builder GUI / CLI — Python 2.7 compatible
© 2025 — Stock Invoice Generator

Функционал:
- Сохраняет структуру счёта 1:1 (те же колонки и порядок; дубликаты заголовков допустимы)
- "Код" всегда артикул (без кириллицы + содержит \d{4,6}); "Товар" — имя из profiles_catalog.xlsx (если есть), иначе сам код
- Если исходного артикула хватает — переносим строку как есть (нормализуем код/товар, пересчитываем итоги/НДС)
- Если не хватает — ПОЛНАЯ замена аналогами (без смешивания с оригиналом) по rules:
  * Цена: правила → каталог Flexy → исходная цена строки
  * Подгонка суммы до целевой (±1%) с корректировкой на последнем лоте; шаг по метрам 0.1
- "в т.ч. НДС" = Всего * 20/120
- "№/Номер/No" — автонумерация 1…N (если колонка есть)
- Результат сохраняется в ~/Downloads/<имя_счёта>_processed.xlsx
- Резолвер путей: нормализация Unicode и поиск по базовому имени в cwd, папке скрипта, ~/Downloads, ~/Desktop, ~ (рекурсивно 1 уровень)
"""

import os
import re
import sys
import math
import glob
import logging
import unicodedata

import pandas as pd

# Tkinter для Py2
try:
    from Tkinter import Tk, Text, Scrollbar, Button, END, Toplevel
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
except Exception:
    Tk = None
    Text = Scrollbar = Button = END = Toplevel = None
    filedialog = None
    messagebox = None

# Каталог Flexy (не менять по ТЗ)
try:
    import flexy_catalog_loader
except Exception as _e:
    flexy_catalog_loader = None

# ----------------------- Константы -----------------------
VAT_RATE       = 0.20      # 20%
VALUE_EPS      = 0.01      # ±1% допуск по сумме
AMOUNT_ABS_TOL = 0.01      # абсолютный минимум по сумме на допуск
QTY_STEP       = 0.1       # шаг по метрам (0.1 м)
SCAN_HEADER_ROWS = 80
RULES_DEFAULT_FILE = "analogs_priority.xlsx"
PROFILES_FILE     = "profiles_catalog.xlsx"

# ----------------------- Логирование -----------------------
logger = logging.getLogger("invoice_app")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s")
    _fh = logging.FileHandler("app.log")
    _fh.setFormatter(_fmt)
    _sh = logging.StreamHandler(sys.stdout)
    _sh.setFormatter(_fmt)
    logger.addHandler(_fh); logger.addHandler(_sh)
    logger.propagate = False

# ----------------------- Утилиты -----------------------
CODE_NO_CYR  = re.compile(u'^[^\u0400-\u04FF]+$')  # нет кириллицы
BASE_DIGITS6 = re.compile(r'\d{6}')
BASE_DIGITS5 = re.compile(r'\d{5}')
BASE_DIGITS4 = re.compile(r'\d{4}')

def normalize_header_name(h):
    s = '' if h is None else unicode(h) if not isinstance(h, unicode) else h
    s = s.lower().replace(u'\n', u' ').replace(u'\r', u' ').strip()
    return s

def to_str_cell(x):
    try:
        if pd.isna(x):
            return u""
    except Exception:
        pass
    try:
        s = unicode(x)
    except Exception:
        s = u""
    return s.strip()

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
    # округление к шагу QTY_STEP и до 2 знаков
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
    """Надёжный резолвер файлов: нормализация Unicode + поиск по имени в cwd, папке скрипта, ~/Downloads, ~/Desktop, ~."""
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
    try:
        roots.append(os.getcwd())
    except Exception:
        pass
    try:
        roots.append(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        pass
    roots.append(os.path.expanduser('~/Downloads'))
    roots.append(os.path.expanduser('~/Desktop'))
    roots.append(os.path.expanduser('~'))

    # прямой поиск по шаблону в roots
    for r in roots:
        try:
            for g in glob.glob(os.path.join(r, '*'+base+'*')):
                if os.path.exists(g) and os.path.isfile(g):
                    logger.info('[path] Found by glob in {0}: {1}'.format(r, g))
                    return g
        except Exception:
            pass

    # неглубокий рекурс
    for r in roots:
        try:
            for root, dirs, files in os.walk(r):
                # ограничим глубину до 2 уровней
                depth = root[len(r):].count(os.sep)
                if depth > 2:
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

# ----------------------- Парсинг правил аналогов -----------------------
def load_analog_rules(rules_path, log):
    """Возвращает dict: base_code -> list of (analog_base_code, price_or_None)."""
    rules_path = resolve_existing_path(rules_path or RULES_DEFAULT_FILE)
    if not os.path.exists(rules_path):
        log.append('[rules] Файл не найден: ' + rules_path)
        return {}

    try:
        df = pd.read_excel(rules_path, dtype=object)
    except Exception as e:
        log.append('[rules] Ошибка чтения: ' + unicode(e))
        return {}

    cols = list(df.columns)
    cols_low = [normalize_header_name(c) for c in cols]

    # базовая колонка
    base_cands = [u'товар', u'база', u'код', u'артикул', u'base']
    base_idx = None
    for i, n in enumerate(cols_low):
        for k in base_cands:
            if k in n:
                base_idx = i; break
        if base_idx is not None:
            break
    if base_idx is None:
        base_idx = 0

    # аналоговые и ценовые колонки
    analog_idxs = []
    price_idxs  = []
    for i, n in enumerate(cols_low):
        if (u'аналог' in n) or (u'analog' in n):
            analog_idxs.append(i)
        if (u'цена' in n) or (u'price' in n):
            price_idxs.append(i)

    # номера для пар
    def num_from(text):
        m = re.search(r'(\d+)', text or u'')
        return m.group(1) if m else None

    analog_num = {}
    price_num  = {}
    for i in analog_idxs:
        analog_num[i] = num_from(cols_low[i])
    for i in price_idxs:
        price_num[i] = num_from(cols_low[i])

    # привязка цен к аналогам
    price_for = {}
    for ai in analog_idxs:
        n = analog_num.get(ai)
        chosen = None
        if n is not None:
            for pi in price_idxs:
                if price_num.get(pi) == n:
                    chosen = pi
                    break
        if chosen is None:
            # ближайшая цена справа в пределах 3 столбцов
            for off in (1,2,3):
                pi = ai + off
                if pi in price_idxs:
                    chosen = pi
                    break
        price_for[ai] = chosen

    rules = {}
    for ridx, row in df.iterrows():
        base_code = extract_base_code(row.iloc[base_idx] if base_idx < len(row) else None)
        if not base_code:
            continue
        pairs = []
        for ai in analog_idxs:
            aval = row.iloc[ai] if ai < len(row) else None
            analog_base = extract_base_code(aval)
            if not analog_base:
                continue
            pi = price_for.get(ai)
            price_val = None
            if pi is not None and pi < len(row):
                price_val = to_float_cell(row.iloc[pi])
            pairs.append((analog_base, price_val))
        if pairs:
            if base_code not in rules:
                rules[base_code] = []
            # уберём дубликаты (по 2‑кортежам)
            seen = set()
            new_list = []
            for p in rules[base_code] + pairs:
                key = (p[0], p[1] if p[1] is not None else u'')
                if key not in seen:
                    seen.add(key)
                    new_list.append(p)
            rules[base_code] = new_list

    log.append('[rules] Загружено базовых кодов: {0} ({1})'.format(len(rules), os.path.basename(rules_path)))
    return rules

# ----------------------- Каталог Flexy и профили -----------------------
def load_flexy_catalog(log):
    if flexy_catalog_loader is None:
        log.append('[catalog] Нет модуля flexy_catalog_loader')
        return pd.DataFrame(columns=['code','price_rub','BaseCode'])
    try:
        cat = flexy_catalog_loader.load_catalog()
        if 'code' not in cat.columns or 'price_rub' not in cat.columns:
            raise Exception('В каталоге нет колонок code/price_rub')
        cat = cat.copy()
        cat['code'] = cat['code'].astype(unicode)
        cat['BaseCode'] = cat['code'].map(extract_base_code)
        log.append('[catalog] Каталог Flexy: {0} позиций'.format(len(cat)))
        return cat
    except Exception as e:
        log.append('[catalog] Ошибка загрузки: ' + unicode(e))
        return pd.DataFrame(columns=['code','price_rub','BaseCode'])

def load_profiles_mapping(log):
    path = resolve_existing_path(PROFILES_FILE)
    if not os.path.exists(path):
        log.append('[profiles] profiles_catalog.xlsx не найден — имена = код')
        return {}
    try:
        df = pd.read_excel(path, dtype=object)
        cols = list(df.columns)
        low = [normalize_header_name(c) for c in cols]
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
            c = to_str_cell(r.iloc[code_idx])
            nm = to_str_cell(r.iloc[name_idx])
            if c:
                mp[c] = nm or c
        log.append('[profiles] Наименований: {0}'.format(len(mp)))
        return mp
    except Exception as e:
        log.append('[profiles] Ошибка загрузки: ' + unicode(e) + ' — имена = код')
        return {}

# ----------------------- Остатки -----------------------
def load_stock(stock_path, catalog_df, log):
    stock_path = resolve_existing_path(stock_path)
    if not os.path.exists(stock_path):
        raise IOError('Файл остатков не найден: ' + stock_path)

    raw = pd.read_excel(stock_path, header=None)
    articles = raw.iloc[9:, 0]
    qtys     = raw.iloc[9:, 1]

    df = pd.DataFrame({'Артикул': articles, 'Остаток': qtys})
    df = df.dropna(how='all')
    df['Артикул'] = df['Артикул'].astype(unicode).str.strip()
    df['Остаток'] = pd.to_numeric(df['Остаток'], errors='coerce').fillna(0.0)
    df = df[df['Артикул'] != u'']
    # агрег по дубликатам
    df = df.groupby('Артикул', as_index=False)['Остаток'].sum()

    df['BaseCode'] = df['Артикул'].map(extract_base_code)

    if not catalog_df.empty:
        price_map = dict(zip(catalog_df['code'], catalog_df['price_rub']))
        df['price_rub'] = df['Артикул'].map(price_map)
    else:
        df['price_rub'] = pd.np.nan

    log.append('[stock] Остатков артикулов: {0}'.format(len(df)))
    return df

# ----------------------- Чтение счёта -----------------------
def is_code_header(h):   n = normalize_header_name(h); return (u'код' in n) or (u'артикул' in n) or (u'sku' in n) or (u'code' in n)
def is_qty_header(h):    n = normalize_header_name(h); return (u'кол-во' in n) or (u'количество' in n) or (u'qty' in n) or (u'кол.' in n) or (u'кол' in n)
def is_price_header(h):  n = normalize_header_name(h); return (u'цена' in n) or (u'price' in n) or (u'ед.цена' in n) or (u'цена за' in n)
def is_total_header(h):  n = normalize_header_name(h); return (u'всего' in n) or (u'сумма' in n) or (u'итого' in n) or (u'amount' in n) or (u'стоимость' in n)
def is_nds_header(h):    n = normalize_header_name(h); return (u'в т.ч. ндс' in n) or (u'ндс' in n) or (u'vat' in n)
def is_name_header(h):   n = normalize_header_name(h); return (u'товар' in n) or (u'наименование' in n) or (u'номенклатура' in n) or (u'product' in n) or (u'item' in n) or (u'наим.' in n)
def is_unit_header(h):   n = normalize_header_name(h); return (u'ед.' in n) or (u'ед ' in n) or (u'единица' in n) or (u'ед. изм' in n) or (u'ед.изм' in n) or (u'unit' in n)
def is_num_header(h):    n = normalize_header_name(h); return (u'№' in n) or (u'номер' in n) or (u'no' in n)

def find_invoice_header_row(raw_df, log):
    max_rows = min(SCAN_HEADER_ROWS, len(raw_df))
    for i in range(max_rows):
        row = raw_df.iloc[i].tolist()
        names = [to_str_cell(x) for x in row]
        if any([is_code_header(n) for n in names]) and any([is_qty_header(n) for n in names]):
            log.append('[invoice] Заголовок найден на строке (1‑based): {0}'.format(i+1))
            return i
    log.append('[invoice] Заголовок эвристикой не найден — берём первую строку')
    return 0

def read_invoice_table(invoice_path, log):
    invoice_path = resolve_existing_path(invoice_path)
    if not os.path.exists(invoice_path):
        raise IOError('Счёт не найден: ' + invoice_path)

    raw = pd.read_excel(invoice_path, header=None, dtype=object)
    header_row = find_invoice_header_row(raw, log)
    headers = [to_str_cell(x) for x in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row+1:].copy()
    df.columns = headers
    df = df.dropna(how='all')

    # выкинем строковые "итоги"
    def is_total_row(series):
        for v in series.values:
            s = to_str_cell(v).lower()
            if u'итого' in s and len(s) <= 20:
                return True
        return False

    table = df[~df.apply(is_total_row, axis=1)].copy()
    columns_order = headers[:]

    # роли — по ИНДЕКСАМ (позициям), чтобы дубли заголовков не ломали логику
    roles = {}
    for i, col in enumerate(headers):
        if 'num'   not in roles and is_num_header(col):    roles['num'] = i
        if 'name'  not in roles and is_name_header(col):   roles['name'] = i
        if 'code'  not in roles and is_code_header(col):   roles['code'] = i
        if 'qty'   not in roles and is_qty_header(col):    roles['qty'] = i
        if 'unit'  not in roles and is_unit_header(col):   roles['unit'] = i
        if 'price' not in roles and is_price_header(col):  roles['price'] = i
        if 'nds'   not in roles and is_nds_header(col):    roles['nds'] = i
        if 'total' not in roles and is_total_header(col):  roles['total'] = i

    log.append('[invoice] Табличных строк: {0}; колонок: {1}'.format(len(table), len(columns_order)))
    return table, columns_order, roles

# ----------------------- Инвентарь -----------------------
class Inventory(object):
    def __init__(self, stock_df):
        # актуальные остатки по коду
        self.by_code = dict(zip(stock_df['Артикул'], stock_df['Остаток']))
        # мап цены из stock_df (из каталога подтянуты)
        self.price_by_code = dict(zip(stock_df['Артикул'], stock_df['price_rub']))
        # индексация по базовому коду
        self.base_to_codes = {}
        for _, r in stock_df.iterrows():
            bc = r.get('BaseCode')
            art = r.get('Артикул')
            q   = r.get('Остаток')
            if bc:
                self.base_to_codes.setdefault(bc, []).append((art, float(q or 0.0)))
        # сортируем по доступу
        for bc in self.base_to_codes:
            self.base_to_codes[bc].sort(key=lambda t: (-t[1], t[0]))

    def get_avail(self, code):
        return float(self.by_code.get(code, 0.0))

    def consume(self, code, qty):
        cur = float(self.by_code.get(code, 0.0))
        newv = max(0.0, cur - float(qty))
        self.by_code[code] = newv

    def lots_for_base(self, base_code):
        # возвращает список (код, доступ)
        return list(self.base_to_codes.get(base_code, []))

# ----------------------- Нормализация кода/имени -----------------------
def normalize_code_and_name_values(vals, roles, name_map):
    code_i = roles.get('code')
    name_i = roles.get('name')

    raw_code = to_str_cell(vals[code_i]) if code_i is not None and code_i < len(vals) else u""
    raw_name = to_str_cell(vals[name_i]) if name_i is not None and name_i < len(vals) else u""

    code_val = raw_code if looks_like_code_scalar(raw_code) else u""
    if not code_val and raw_name and looks_like_code_scalar(raw_name):
        code_val = raw_name

    if not code_val:
        # попробуем выдернуть кусок без кириллицы, содержащий базовые цифры
        for candidate in (raw_code, raw_name):
            if candidate and CODE_NO_CYR.match(candidate) and (BASE_DIGITS6.search(candidate) or BASE_DIGITS5.search(candidate) or BASE_DIGITS4.search(candidate)):
                code_val = candidate.strip()
                break

    if code_i is not None and code_i < len(vals):
        vals[code_i] = code_val
    if name_i is not None and name_i < len(vals):
        vals[name_i] = name_map.get(code_val, code_val)

def recalc_totals_values(vals, roles):
    qty_i   = roles.get('qty')
    price_i = roles.get('price')
    total_i = roles.get('total')
    nds_i   = roles.get('nds')

    if qty_i is not None and price_i is not None and total_i is not None:
        q = to_float_cell(vals[qty_i])
        p = to_float_cell(vals[price_i])
        if q is not None and p is not None:
            total = round_money(q * p)
            vals[total_i] = total
            if nds_i is not None:
                vals[nds_i] = round_money(total * VAT_RATE / (1.0 + VAT_RATE))
    elif nds_i is not None and total_i is not None:
        t = to_float_cell(vals[total_i])
        if t is not None:
            vals[nds_i] = round_money(t * VAT_RATE / (1.0 + VAT_RATE))

# ----------------------- Цена коду -----------------------
def price_for_code(code, price_from_rule, price_src_row, catalog_price_by_code):
    if price_from_rule is not None:
        return float(price_from_rule)
    p = catalog_price_by_code.get(code)
    if p is not None and not (isinstance(p, float) and pd.isna(p)):
        try:
            return float(p)
        except Exception:
            pass
    if price_src_row is not None:
        return float(price_src_row)
    return None

# ----------------------- Основная обработка -----------------------
def process_invoice(stock_df, invoice_table, columns_order, roles, rules, catalog_df, name_map, log):
    inv = Inventory(stock_df)
    result_rows_values = []   # список списков по позициям колонок
    n_cols = len(columns_order)
    catalog_price_by_code = {}
    if not catalog_df.empty:
        catalog_price_by_code = dict(zip(catalog_df['code'], catalog_df['price_rub']))

    def base_row_values(src_row):
        vals = []
        # src_row — Series; берем в порядке столбцов
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
                result_rows_values.append(vals)
                return

    def add_row_from_source(src_row):
        vals = base_row_values(src_row)
        normalize_code_and_name_values(vals, roles, name_map)
        recalc_totals_values(vals, roles)
        append_if_not_empty(vals)

    def add_row_for_analog(src_row, code, qty, price):
        vals = base_row_values(src_row)
        code_i  = roles.get('code')
        name_i  = roles.get('name')
        qty_i   = roles.get('qty')
        price_i = roles.get('price')
        total_i = roles.get('total')
        nds_i   = roles.get('nds')

        if code_i is not None and code_i < n_cols:
            vals[code_i] = code
        if name_i is not None and name_i < n_cols:
            vals[name_i] = name_map.get(code, code)
        if qty_i is not None and qty_i < n_cols:
            vals[qty_i] = round_qty(qty)
        if price_i is not None and price_i < n_cols:
            vals[price_i] = round_money(price)
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

    code_i  = roles.get('code')
    name_i  = roles.get('name')
    qty_i   = roles.get('qty')
    price_i = roles.get('price')

    # обход строк счёта
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

        # если хватает исходного артикула — переносим как есть
        if inv.get_avail(art) >= qty_src:
            add_row_from_source(src_row)
            inv.consume(art, qty_src)
            continue

        base_code = extract_base_code(art)
        if not base_code or base_code not in rules:
            log.append(u'{0}: аналогов нет по правилам — строка не закрыта'.format(art))
            continue

        # подготовка лотов по приоритету правил
        lots = []   # (code, avail, price, origin)
        for analog_base, price_from_rule in rules[base_code]:
            base_lots = inv.lots_for_base(analog_base)
            if not base_lots:
                continue
            for lot_code, lot_avail in base_lots:
                if lot_avail <= 0:
                    continue
                p = price_for_code(lot_code, price_from_rule, price_src, inv.price_by_code if inv.price_by_code else catalog_price_by_code)
                if p is None:
                    # попробуем из общего каталога
                    p = price_for_code(lot_code, price_from_rule, price_src, catalog_price_by_code)
                origin = u'по правилам' if price_from_rule is not None else (u'из каталога' if (lot_code in (inv.price_by_code or {}) or lot_code in catalog_price_by_code) else u'исходная цена')
                lots.append((lot_code, float(lot_avail), float(p) if p is not None else None, origin))

        if not lots:
            log.append(u'{0}: аналогов нет на складе — строка не закрыта'.format(art))
            continue

        if price_src is None and all([l[2] is None for l in lots]):
            log.append(u'{0}: не удалось вычислить цену — строка не закрыта'.format(art))
            continue

        created_any = False

        if price_src is not None:
            # Подгонка по сумме
            target_total = round_money(qty_src * price_src)
            allowed_delta = max(AMOUNT_ABS_TOL, abs(target_total) * VALUE_EPS)

            acc_value = 0.0
            used = []  # (i, q)

            for i, (c, avail_i, p_i, origin_i) in enumerate(lots):
                if p_i is None or p_i <= 0:
                    continue
                remaining_value = target_total - acc_value
                if remaining_value <= allowed_delta:
                    break
                q_desired = 0.0 if p_i == 0 else (remaining_value / p_i)
                q_take = min(avail_i, max(0.0, q_desired))
                q_take = round_qty(q_take)
                if q_take <= 0:
                    continue
                acc_value = round_money(acc_value + q_take * p_i)
                used.append((i, q_take))

            if used:
                # корректировка на последнем лоте
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
            # Подгонка по количеству (если исходная цена неизвестна)
            remaining_qty = round_qty(qty_src)
            used_q = []  # (i, q)
            for i, (c, avail_i, p_i, origin_i) in enumerate(lots):
                if remaining_qty <= 0:
                    break
                if p_i is None:
                    used_q = []
                    log.append(u'{0}: нет цены у аналога {1} — строка не закрыта'.format(art, c))
                    break
                q_take = min(avail_i, remaining_qty)
                q_take = round_qty(q_take)
                if q_take <= 0:
                    continue
                used_q.append((i, q_take))
                remaining_qty = round_qty(remaining_qty - q_take)

            if used_q and remaining_qty <= 0:
                for i, q in used_q:
                    c_i, avail_i, p_i, origin_i = lots[i]
                    add_row_for_analog(src_row, c_i, q, p_i)
                    inv.consume(c_i, q)
                    log.append(u'{0}: {1:.2f} → {2} по {3:.2f} ₽ ({4})'.format(art, q, c_i, p_i, origin_i))
                    created_any = True
                # инфо‑лог по сумме
                total_i = roles.get('total')
                if total_i is not None:
                    sum_analog = 0.0
                    cnt = len(used_q)
                    for vals in result_rows_values[-cnt:]:
                        t = to_float_cell(vals[total_i])
                        if t is not None:
                            sum_analog += t
                    log.append(u'итог исходной: цена отсутствовала; сумма по аналогам ≈ {0:.2f} ₽'.format(round_money(sum_analog)))
            else:
                log.append(u'{0}: аналогов недостаточно — строка не закрыта'.format(art))

        if not created_any:
            continue

    # Сборка DataFrame в точном порядке колонок
    result_df = pd.DataFrame(result_rows_values, columns=columns_order)

    # удалить полностью пустые строки (на всякий)
    if not result_df.empty:
        mask = result_df.apply(lambda r: any([to_str_cell(v) != u'' for v in r]), axis=1)
        result_df = result_df[mask]

    # автонумерация
    num_i = roles.get('num')
    if num_i is not None and not result_df.empty:
        # присвоим 1…N в столбце с индексом num_i
        nums = list(range(1, len(result_df) + 1))
        # если дубликаты заголовков — установим по позиции
        colname = columns_order[num_i]
        result_df.iloc[:, num_i] = nums

    return result_df

# ----------------------- Сохранение -----------------------
def save_result(df, input_invoice_path, log):
    base = os.path.splitext(os.path.basename(input_invoice_path))[0]
    out_name = base + u"_processed.xlsx"
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    try:
        if not os.path.exists(downloads):
            os.makedirs(downloads)
    except Exception:
        pass
    out_path = os.path.join(downloads, out_name)
    with pd.ExcelWriter(out_path) as writer:
        df.to_excel(writer, index=False)
    log.append(u'[save] Результат сохранён: {0}'.format(out_path))
    return out_path

# ----------------------- GUI -----------------------
class InvoiceGUI(object):
    def __init__(self, root):
        self.root = root
        root.title('Stock → Invoice Builder (Py2.7)')

        self.log = []
        self.stock_path = None
        self.invoice_path = None
        self.rules_path = RULES_DEFAULT_FILE

        # UI
        self.txt = Text(root, height=20, width=100)
        sc = Scrollbar(root, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sc.set)
        self.txt.pack(side='left', fill='both', expand=True)
        sc.pack(side='right', fill='y')

        Button(root, text='Загрузить остатки', command=self.pick_stock).pack()
        Button(root, text='Загрузить счёт', command=self.pick_invoice).pack()
        Button(root, text='Собрать счёт', command=self.build).pack()
        Button(root, text='Посмотреть логи', command=self.show_logs).pack()

    def log_add(self, msg):
        self.log.append(msg)
        try:
            self.txt.insert(END, msg + u'\n'); self.txt.see(END)
        except Exception:
            pass
        logger.info(msg)

    def pick_stock(self):
        if filedialog is None:
            self.log_add(u'GUI: нет filedialog')
            return
        p = filedialog.askopenfilename(title='Выберите файл остатков', filetypes=[('Excel', '*.xlsx *.xls')])
        if p:
            self.stock_path = p
            self.log_add(u'[ui] Остатки: ' + p)

    def pick_invoice(self):
        if filedialog is None:
            self.log_add(u'GUI: нет filedialog')
            return
        p = filedialog.askopenfilename(title='Выберите файл счёта', filetypes=[('Excel', '*.xlsx *.xls')])
        if p:
            self.invoice_path = p
            self.log_add(u'[ui] Счёт: ' + p)

    def show_logs(self):
        top = Toplevel(self.root); top.title('Логи')
        t = Text(top, height=25, width=100)
        t.pack(fill='both', expand=True)
        try:
            for line in self.log:
                t.insert(END, line + u'\n')
            t.config(state='disabled')
        except Exception:
            pass

    def build(self):
        if not self.stock_path or not self.invoice_path:
            if messagebox:
                messagebox.showerror('Ошибка', 'Загрузите остатки и счёт')
            return
        try:
            outp = run_pipeline(self.stock_path, self.invoice_path, self.rules_path, self.log)
            self.log_add(u'✅ Готово: ' + outp)
            if messagebox:
                messagebox.showinfo('Готово', u'Результат: ' + outp)
        except Exception as e:
            self.log_add(u'❌ Ошибка: ' + unicode(e))
            if messagebox:
                messagebox.showerror('Ошибка', unicode(e))

# ----------------------- Пайплайн -----------------------
def run_pipeline(stock_path, invoice_path, rules_path, log):
    rules   = load_analog_rules(rules_path, log)
    catalog = load_flexy_catalog(log)
    name_mp = load_profiles_mapping(log)
    stock   = load_stock(stock_path, catalog, log)
    inv_tbl, cols_order, roles = read_invoice_table(invoice_path, log)
    result  = process_invoice(stock, inv_tbl, cols_order, roles, rules, catalog, name_mp, log)
    outp    = save_result(result, invoice_path, log)
    # печать логов в stdout для CLI
    try:
        for line in log:
            sys.stdout.write((line + u'\n').encode('utf-8'))
    except Exception:
        pass
    return outp

# ----------------------- main -----------------------
def main():
    import argparse
    parser = argparse.ArgumentParser(description='Invoice Builder (Py2.7)')
    parser.add_argument('--cli', nargs=2, metavar=('STOCK_XLSX','INVOICE_XLSX'), help='CLI-режим')
    parser.add_argument('--rules', default=RULES_DEFAULT_FILE, help='Путь к analogs_priority.xlsx')
    args = parser.parse_args()

    if args.cli:
        stock_path   = resolve_existing_path(args.cli[0])
        invoice_path = resolve_existing_path(args.cli[1])
        rules_path   = resolve_existing_path(args.rules)
        run_pipeline(stock_path, invoice_path, rules_path, [])
    else:
        if Tk is None:
            sys.stderr.write("GUI недоступен в этой среде. Используйте --cli.\n")
            sys.exit(2)
        root = Tk()
        app = InvoiceGUI(root)
        root.mainloop()

if __name__ == '__main__':
    # Для Py2: включим unicode‑литералы в stdout
    try:
        reload(sys)
        sys.setdefaultencoding('utf-8')
    except Exception:
        pass
    main()
