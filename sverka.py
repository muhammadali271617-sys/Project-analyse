"""
Сверка Excel (анализ цен) с PDF счёт-фактурами Didox.uz.

Берёт «цветные» строки рабочего листа (<месяц>_проверка / <месяц>_разницалар),
где Разница (AM) вышла за пределы (AO / AP), находит для каждой строки фактуру
по номеру (R) и договору (Q), находит в фактуре нужную позицию и сравнивает
Кол-во (H), Цену (I), Сумму (J). Результат — HTML-отчёт по кодам товаров
с проблемами.

Запуск:
    python sverka.py "D:\\папка с Excel и фактурами"
    python sverka.py "D:\\папка" --period 2026-10          (без вопроса о периоде)
    python sverka.py "D:\\папка" --period 2026-09,2026-10
    python sverka.py "D:\\папка" --period все

Импортные строки (AA = «имп») не проверяются.
Колонки AT/AU (комментарии бухгалтера/экономиста) не читаются.
"""

import argparse
import collections
import datetime as dt
import html
import logging
import re
import sys
from pathlib import Path

import openpyxl
import pdfplumber

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# Меняется при каждом обновлении программы — видно на странице, в отчёте и в Excel
VERSION = "2026-11-18 · v9"

# ---------------------------------------------------------------- допуски
TOL_QTY = 0.001
TOL_SUM = 1.0
SWAP_MIN_SIM = 0.15    # ниже — у позиции с теми же числами «совсем другое» название
SWAP_NAME_SIM = 0.6    # выше — позиция по названию считается уверенно найденной


def tol_price(p):
    return max(0.01, abs(p) * 1e-4)


# ---------------------------------------------------------------- утилиты
def norm_num(s):
    """Номер фактуры/договора для сравнения: без пробелов, регистра, / = ."""
    return re.sub(r"\s+", "", str(s or "")).upper().replace("/", ".")


def strip_brackets(stem):
    return re.sub(r"\s+", " ", re.sub(r"\([^)]*\)", "", stem)).strip()


def to_float(s):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).replace("\n", "").replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def eq(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


def fmt(x, nd=2):
    if x is None:
        return "—"
    if isinstance(x, float) and abs(x - round(x)) < 1e-9 and abs(x) >= 1:
        nd = 0
    s = f"{x:,.{nd}f}".replace(",", " ").replace(".", ",")
    if "," in s:
        s = s.rstrip("0").rstrip(",")
    return s


def esc(s):
    return html.escape(str(s if s is not None else ""))


# ---------------------------------------------------------------- единицы
def unit_info(u):
    """Возвращает (базовая единица, сколько базовых в 1 единице, исходный текст)."""
    raw = str(u or "").strip()
    s = raw.lower().replace("ё", "е")
    m = re.search(r"(paqir|пакир|ведро|qop|мешок|упак\w*|коробк\w*|quti)\s*=\s*([\d.,]+)\s*([a-zа-я]+)", s)
    if m:
        base, k, _ = unit_info(m.group(3))
        return base, (to_float(m.group(2)) or 1) * k, raw
    table = [
        (r"^(шт|штук|штука|дона|dona|pcs|комплект|компл|к-т)", "шт", 1),
        (r"^(тонна|тн|т|tonna|ton)$", "кг", 1000),
        (r"^(кг|килограмм|kilogramm|kg|килограм)", "кг", 1),
        (r"^(г|гр|грамм|gramm|gram)$", "кг", 0.001),
        (r"^(м2|кв\.?\s*м|квадрат|kv\.?\s*m|m2)", "м2", 1),
        (r"^(м|метр|metr|m)$", "м", 1),
        (r"^(л|литр|litr|l)$", "л", 1),
    ]
    s2 = s.rstrip(".").strip()
    for rx, base, k in table:
        if re.search(rx, s2):
            return base, k, raw
    return s2, 1, raw


# ---------------------------------------------------------------- похожесть названий
STOP = {"для", "и", "в", "на", "с", "из", "мм", "mm", "шт", "см", "по", "без", "марка", "артикул", "модель"}


# Сокращения, которые встречаются в фактурах и в 1С
ABBREV = [
    (r"(?<![a-zа-я])х\s*/\s*к(?![a-zа-я])", " холоднокатаный "),
    (r"(?<![a-zа-я])г\s*/\s*к(?![a-zа-я])", " горячекатаный "),
    (r"(?<![a-zа-я])оцинк\.?(?![a-zа-я])", " оцинкованный "),
    (r"(?<![a-zа-я])нерж\.?(?![a-zа-я])", " нержавеющий "),
    (r"(?<![a-zа-я])п/э(?![a-zа-я])", " полиэтиленовый "),
    (r"\bметал(?!л)", "металл"),
]


def tokens(name):
    s = str(name or "").lower().replace("ё", "е").replace(",", ".")
    for rx, full in ABBREV:
        s = re.sub(rx, full, s)
    out = {}
    for t in re.findall(r"[0-9]+(?:\.[0-9]+)?|[a-zа-я][a-zа-я0-9\-]*", s):
        t = t.strip("-")
        if not t or t in STOP:
            continue
        if re.fullmatch(r"[0-9.]+", t):
            out[t] = 2.0                      # числа и размеры весят больше
        elif re.search(r"[0-9]", t):
            out[t] = 2.0                      # артикулы типа RST-534A
        elif len(t) >= 3:
            out[t[:4]] = 1.0                  # основа слова: первые 4 буквы (стрейч = стретч)
    return out


def similarity(a, b):
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    common = sum(w for t, w in ta.items() if t in tb)
    return common / min(sum(ta.values()), sum(tb.values()))


# ---------------------------------------------------------------- чтение фактур
HEADER_RX = re.compile(r"Ҳисоб-фактура[^:\n]*:\s*(\S+)\s+(\d{2}\.\d{2}\.\d{4})")
BODY_RX = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s+даги\s+(\S+)-сонли(?:(?!даги)[\s\S]){0,80}?Ҳисобварақ-фактура")
CONTRACT_RX = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s+даги\s+(.+?)-сонли\s+шартномага")
SUPPLIER_RX = re.compile(r"Етказиб берувчи:\s*(.+?)\s+Сотиб олувчи")
INN_RX = re.compile(r"Етказиб берувчининг СТИР\s+(\d{9})")

# Официальный PDF-бланк Didox (кнопка «Скачать», файлы «Счет_фактура_без_акта_…»): шапка на русском
RU_NUM_RX = re.compile(r"Сч[её]т-фактура\s*№\s*(\S+)\s+от\s+(\d{2}\.\d{2}\.\d{4})", re.I)
RU_CONTRACT_RX = re.compile(r"к договору\s*№\s*(.+?)\s+от\s+(\d{2}\.\d{2}\.\d{4})", re.I)
RU_SUPPLIER_RX = re.compile(r"Поставщик:\s*(.+?)\s+Покупатель:")
RU_INN_RX = re.compile(r"Идентификационный номер\s+(\d{9})")
CORRECTION_RX = re.compile(r"№\s*(\S+)\s+ЭҲФ\s*га\s+Тузатувчи", re.I)  # исправленная фактура


def parse_invoice(path):
    inv = {"file": path.name, "items": [], "error": None}
    try:
        with pdfplumber.open(path) as pdf:
            t0 = pdf.pages[0].extract_text() or ""
            m = HEADER_RX.search(t0)
            mb = BODY_RX.search(t0)
            mr = RU_NUM_RX.search(t0)
            inv["num"] = m.group(1) if m else (mb.group(2) if mb else (mr.group(1) if mr else None))
            inv["date"] = m.group(2) if m else (mb.group(1) if mb else (mr.group(2) if mr else None))
            m = CONTRACT_RX.search(t0)
            mr = RU_CONTRACT_RX.search(t0)
            inv["contract"], inv["contract_date"] = ((m.group(2), m.group(1)) if m else
                                                     (mr.group(1), mr.group(2)) if mr else (None, None))
            m = SUPPLIER_RX.search(t0) or RU_SUPPLIER_RX.search(t0)
            inv["supplier"] = m.group(1).strip() if m else None
            m = CORRECTION_RX.search(t0)
            inv["correction"] = m.group(1) if m else None
            m = INN_RX.search(t0)
            if not m:
                # в бланке поставщик слева, покупатель справа — их строки перемешиваются в тексте,
                # поэтому ИНН поставщика ищем только в левой половине страницы
                p0 = pdf.pages[0]
                left = p0.crop((0, 0, p0.width / 2, p0.height)).extract_text() or ""
                m = RU_INN_RX.search(left)
            inv["inn"] = m.group(1) if m else None
            for page in pdf.pages:
                for tb in page.extract_tables():
                    hdr = next((i for i, r in enumerate(tb)
                                if [c for c in r if c][:3] == ["1", "2", "3"]), None)
                    if hdr is None:
                        continue
                    c1 = tb[hdr].index("1")
                    if c1 == 0 or c1 + 8 >= len(tb[hdr]):
                        continue
                    for r in tb[hdr + 1:]:
                        cell = lambda k: (r[c1 + k] or "").replace("\n", " ").strip()
                        no = (r[c1 - 1] or "").strip()
                        if no.isdigit():
                            inv["items"].append(dict(
                                no=int(no), name=cell(0), ikpu=cell(1), unit=cell(2),
                                qty=to_float(cell(3)), price=to_float(cell(4)), sum=to_float(cell(5)),
                                rate=cell(6), vat=to_float(cell(7)), total=to_float(cell(8))))
                        elif not no and cell(0) and inv["items"] and not cell(3):
                            inv["items"][-1]["name"] += " " + cell(0)     # перенос на новую страницу
    except Exception as e:  # битый PDF
        inv["error"] = f"Файл не читается: {e}"
        return inv
    for it in inv["items"]:
        # перенос строки внутри артикула: «RST- 534A», «0204- 008355» → «RST-534A», «0204-008355»
        it["name"] = re.sub(r"(?<=\w)-\s+(?=\w)", "-", it["name"])
    if not inv["num"]:
        inv["error"] = "Не найден номер счёт-фактуры — документ не соответствует стандарту Didox"
    elif not inv["items"]:
        inv["error"] = "Не найдена таблица товаров — документ не соответствует стандарту Didox"
    return inv


# ---------------------------------------------------------------- Excel
def col(letter):
    return openpyxl.utils.column_index_from_string(letter) - 1


COLS = {k: col(k) for k in ["B", "C", "F", "G", "H", "I", "J", "O", "Q", "R", "Z", "AA", "AL",
                            "AM", "AO", "AP", "AH"]}


def work_sheets(sheetnames):
    """Все листы вида «<месяц>_проверка» или «<месяц>_разницалар» (регистр не важен)."""
    return [n for n in sheetnames
            if n.lower().strip().endswith("_проверка") or n.lower().strip().endswith("_разницалар")]


def choose_work_sheet(sheetnames, arg=None, ask=input):
    found = work_sheets(sheetnames)
    if arg:
        if arg not in sheetnames:
            sys.exit(f"Лист «{arg}» не найден. Листы: {', '.join(sheetnames)}")
        return arg
    if not found:
        sys.exit("Не найден лист «<месяц>_проверка» или «<месяц>_разницалар». Листы: " + ", ".join(sheetnames))
    if len(found) == 1:
        return found[0]
    print("\nВ книге несколько листов для проверки:")
    for i, n in enumerate(found, 1):
        print(f"  {i}. {n}")
    while True:
        a = ask("Какой лист проверять? Введите номер: ").strip()
        if a.isdigit() and 1 <= int(a) <= len(found):
            return found[int(a) - 1]
        print(f"Введите число от 1 до {len(found)}")


def find_values_sheet(wb):
    for name in wb.sheetnames:
        if "значение" in name.lower():
            return name
    return None


def is_colored(r):
    am, ao, ap = r[COLS["AM"]], r[COLS["AO"]], r[COLS["AP"]]
    if not isinstance(am, (int, float)):
        return False
    return (isinstance(ap, (int, float)) and am > ap) or (isinstance(ao, (int, float)) and am < ao)


def month_key(d):
    if isinstance(d, dt.datetime):
        return f"{d.year}-{d.month:02d}"
    if isinstance(d, str):
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", d)
        if m:
            return f"{m.group(3)}-{m.group(2)}"
    return "без даты"


MONTHS_RU = ["", "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
             "сентябрь", "октябрь", "ноябрь", "декабрь"]


def month_name(k):
    m = re.fullmatch(r"(\d{4})-(\d{2})", k)
    return f"{MONTHS_RU[int(m.group(2))]} {m.group(1)}" if m else k


def ask_period(available, arg):
    keys = sorted(available)
    if arg:
        if arg.strip().lower() in ("все", "all", "*"):
            return set(keys)
        chosen = {p.strip() for p in arg.split(",") if p.strip()}
        bad = chosen - set(keys)
        if bad:
            sys.exit(f"Нет строк за период: {', '.join(sorted(bad))}. Есть: {', '.join(keys)}")
        return chosen
    print("\nЗа какой период проверять? Цветные строки есть за месяцы:")
    for i, k in enumerate(keys, 1):
        print(f"  {i}. {month_name(k):<16} ({available[k]} строк)")
    print("Введите один номер (один месяц), несколько через запятую (например 1,2) или 'все'.")
    while True:
        a = input("Период: ").strip().lower()
        if a in ("все", "all", "*"):
            return set(keys)
        try:
            idx = {int(x) for x in a.split(",") if x.strip()}
            if idx and all(1 <= i <= len(keys) for i in idx):
                return {keys[i - 1] for i in idx}
        except ValueError:
            pass
        print("Не понял. Пример: 1  или  1,2  или  все")


# ---------------------------------------------------------------- сверка строки
def compare_numbers(H, I, J, it, ex_unit):
    """Сравнивает числа строки Excel с позицией фактуры.
    Возвращает (совпало_полей, список_заметок, флаги)."""
    eb, ek, _ = unit_info(ex_unit)
    ib, ik, iraw = unit_info(it["unit"])
    qty, price, summ, total = it["qty"], it["price"], it["sum"], it["total"]
    price_vat = total / qty if total and qty else None
    flags = set()
    notes = []

    # возможный пересчёт единиц (тонна→кг, г→кг, ведро 29,7 кг → кг)
    factor = ik / ek if eb == ib and ek and ik else 1.0
    q_ok = eq(H, qty, TOL_QTY)
    q_conv = factor != 1 and qty is not None and eq(H, qty * factor, max(TOL_QTY, abs(qty * factor) * 1e-6))
    p_ok = price is not None and I is not None and (eq(I, price, tol_price(price)) or
                                                   (summ and H and eq(I, summ / H, tol_price(price))))
    p_conv = factor != 1 and price is not None and I is not None and eq(I, price / factor, tol_price(price / factor))
    p_vat = price_vat is not None and I is not None and eq(I, price_vat, tol_price(price_vat))
    s_ok = eq(J, summ, TOL_SUM)
    s_vat = eq(J, total, TOL_SUM)

    score = (q_ok or q_conv) + (p_ok or p_conv or p_vat) + (s_ok or s_vat)

    if q_conv and (p_conv or s_ok):
        flags.add("unit_conv_ok")
        notes.append(f"Единицы: в Excel «{ex_unit}», в фактуре «{iraw}». Пересчёт верный "
                     f"({fmt(qty, 3)} {iraw} = {fmt(H, 3)} {ex_unit}).")
    elif eb != ib or (factor != 1 and q_ok):
        flags.add("unit")
        if factor != 1 and q_ok:
            notes.append(f"Единицы: в Excel {fmt(H, 3)} «{ex_unit}», в фактуре {fmt(qty, 3)} «{iraw}». "
                         f"Число совпадает, но единица другая: {fmt(qty, 3)} {iraw} = "
                         f"{fmt(qty * factor, 3)} {ex_unit}. Количество не пересчитано.")
        else:
            notes.append(f"Единицы: в Excel «{ex_unit}», в фактуре «{iraw}».")

    if not (q_ok or q_conv):
        flags.add("qty")
        notes.append(f"Кол-во: в Excel {fmt(H, 3)}, в фактуре {fmt(qty, 3)}.")
    if not (p_ok or p_conv):
        if p_vat:
            flags.add("vat")
            notes.append(f"Цена: в Excel {fmt(I)} — это цена С НДС; в фактуре без НДС {fmt(price)}.")
        else:
            flags.add("price")
            notes.append(f"Цена: в Excel {fmt(I)}, в фактуре {fmt(price)}"
                         + (f" (с НДС {fmt(price_vat)})" if price_vat else "") + ".")
    if not s_ok:
        if s_vat:
            flags.add("vat")
            notes.append(f"Сумма: в Excel {fmt(J)} — это сумма С НДС; в фактуре без НДС {fmt(summ)}.")
        else:
            flags.add("sum")
            notes.append(f"Сумма: в Excel {fmt(J)}, в фактуре без НДС {fmt(summ)}, с НДС {fmt(total)}.")
    return score, notes, flags


def grouped_items(items):
    """Один и тот же товар (то же название и единица) в нескольких строках фактуры —
    дополнительная «сложенная» позиция: кол-во, суммы и НДС складываются."""
    groups = collections.OrderedDict()
    for it in items:
        key = (re.sub(r"\s+", " ", it["name"].lower()).strip(), unit_info(it["unit"])[0])
        groups.setdefault(key, []).append(it)
    out = []
    for members in groups.values():
        if len(members) < 2:
            continue
        add = lambda k: (sum(m[k] for m in members) if all(m[k] is not None for m in members) else None)
        qty, summ = add("qty"), add("sum")
        out.append({
            "no": "+".join(str(m["no"]) for m in members), "name": members[0]["name"],
            "ikpu": members[0].get("ikpu"), "unit": members[0]["unit"], "qty": qty,
            "price": round(summ / qty, 2) if qty and summ is not None else members[0]["price"],
            "sum": summ, "rate": members[0].get("rate"), "vat": add("vat"), "total": add("total"),
            "members": members})
    return out


def match_row(row, inv):
    """Находит позицию в фактуре и формирует вывод."""
    H, I, J, F, G = row["H"], row["I"], row["J"], row["F"], row["G"]
    items = inv["items"]
    scored = []
    for it in items + grouped_items(items):
        sc, notes, flags = compare_numbers(H, I, J, it, G)
        scored.append((it, sc, notes, flags, similarity(F, it["name"])))

    by_name = sorted(scored, key=lambda x: -x[4])
    name_best = by_name[0] if by_name and by_name[0][4] >= 0.5 else None
    if name_best and len(by_name) > 1 and by_name[1][4] >= name_best[4] - 0.1:
        # несколько похожих названий — выбираем среди них по числам
        tied = [x for x in by_name if x[4] >= name_best[4] - 0.1]
        # при равенстве — сложенная позиция (весь товар фактуры), а не одна из её строк
        name_best = max(tied, key=lambda x: (x[1], x[4], bool(x[0].get("members"))))
    num_best = max(scored, key=lambda x: (x[1], x[4], bool(x[0].get("members")))) if scored else None

    res = {"item": None, "status": "", "notes": [], "flags": set(), "sim": 0.0, "other": None}
    if not scored:
        res["status"] = "ОШИБКА"
        res["notes"].append("В фактуре нет позиций.")
        return res

    perfect = [x for x in scored if x[1] == 3]          # кол-во, цена и сумма совпали
    perf_best = max(perfect, key=lambda x: x[4]) if perfect else None

    if perf_best and (perf_best[4] >= SWAP_MIN_SIM or not name_best or name_best[4] < SWAP_NAME_SIM):
        # числа совпали и название хоть немного похоже (общее слово) — это та позиция
        chosen = perf_best
    elif perf_best:
        # числа из Excel принадлежат товару с совсем другим названием,
        # а по названию явно подходит другая позиция — строки перепутаны
        it_n, it_x = name_best[0], perf_best[0]
        res.update(item=it_n, sim=name_best[4], other=it_x, status="ПЕРЕПУТАНЫ СТРОКИ")
        res["flags"] = {"swap"}
        res["notes"].append(
            f"По названию это поз. {it_n['no']} «{it_n['name']}»: {fmt(it_n['qty'], 3)} × {fmt(it_n['price'])} "
            f"= {fmt(it_n['sum'])}. А числа из Excel ({fmt(H, 3)} × {fmt(I)} = {fmt(J)}) принадлежат "
            f"поз. {it_x['no']} «{it_x['name']}». Похоже, в 1С перепутаны строки.")
        return res
    elif name_best:
        chosen = name_best
    elif num_best[1] >= 2:
        chosen = num_best
    else:
        res["status"] = "ПОЗИЦИЯ НЕ НАЙДЕНА"
        res["flags"] = {"notfound"}
        res["item"] = num_best[0] if num_best[1] >= 1 else None
        res["notes"].append(f"В фактуре нет позиции, похожей на «{F}» ни по названию, ни по числам "
                            f"({fmt(H, 3)} × {fmt(I)} = {fmt(J)}).")
        return res

    it, sc, notes, flags, sim = chosen
    res.update(item=it, sim=sim, notes=list(notes), flags=set(flags))
    if it.get("members"):
        ms = it["members"]
        prices = {m["price"] for m in ms}
        res["flags"].add("grouped")
        res["notes"].insert(0, f"ℹ В фактуре этот товар в {len(ms)} строках (поз. {', '.join(str(m['no']) for m in ms)}): "
                               f"{' + '.join(fmt(m['qty'], 3) for m in ms)} = {fmt(it['qty'], 3)}, "
                               f"сумма {' + '.join(fmt(m['sum']) for m in ms)} = {fmt(it['sum'])} — сложено для сверки"
                               + ("." if len(prices) == 1 else f"; цены в строках разные ({', '.join(fmt(p) for p in sorted(prices))}), "
                                                               f"взята средняя {fmt(it['price'])}."))
    if sim < SWAP_MIN_SIM:
        res["flags"].add("name")
        if sc == 3:
            res["notes"].append(f"Названия без общих слов: в Excel «{F}», в фактуре «{it['name']}». "
                                f"Кол-во, цена и сумма совпали — проверьте глазами, тот ли это товар.")
        else:
            res["flags"].add("other_item")
            res["notes"].insert(0, f"Товара «{F}» в фактуре не найдено. Часть чисел совпадает с поз. "
                                   f"{it['no']} «{it['name']}» — это другой товар. Возможно, в 1С "
                                   f"перепутаны строки.")
    hard = res["flags"] & {"qty", "price", "sum", "vat"}
    if hard:
        # чьи это числа? ищем другую позицию с той же суммой или ценой
        for o in items:
            if o is it or o in it.get("members", []):
                continue
            if eq(J, o["sum"], TOL_SUM) or eq(J, o["total"], TOL_SUM):
                res["notes"].append(f"Сумма из Excel ({fmt(J)}) совпадает с поз. {o['no']} «{o['name']}» — "
                                    f"похоже, в 1С перепутаны строки.")
                res["other"] = o
                res["flags"].add("sum_other")
                break
            if o["price"] and eq(I, o["price"], tol_price(o["price"])) and "price" in res["flags"]:
                res["notes"].append(f"Цена из Excel ({fmt(I)}) совпадает с ценой поз. {o['no']} «{o['name']}».")
                res["other"] = o
                res["flags"].add("price_other")
                break
        res["status"] = "ОШИБКА"
    elif res["flags"] & {"unit", "name"}:
        res["status"] = "ПРОВЕРИТЬ"
    else:
        res["status"] = "OK"
    return res


# ---------------------------------------------------------------- основной процесс
class CheckError(Exception):
    """Понятная пользователю ошибка (нет Excel, нет листа и т.п.)."""


class NeedSheet(CheckError):
    """В книге несколько листов для проверки — нужно выбрать один."""
    def __init__(self, sheets):
        super().__init__("В книге несколько листов для проверки — выберите лист.")
        self.sheets = sheets


def _say(progress, text, done=None, total=None):
    if progress:
        progress(text, done, total)


def find_excels(folder):
    """Excel-файлы в папке (без временных ~$, проверенных копий и книг «Ошибки по бухгалтерам»)."""
    return sorted(p for p in Path(folder).iterdir()
                  if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                  and not p.name.startswith("~$") and "проверено" not in p.name.lower()
                  and not p.name.lower().startswith("ошибки_по_бухгалтерам"))


def is_open_in_excel(path):
    """Excel создаёт рядом файл «~$имя» пока книга открыта."""
    path = Path(path)
    return (path.parent / ("~$" + path.name[2:])).exists() or (path.parent / ("~$" + path.name)).exists()


def load_excel(xlsx, sheet=None, progress=None, ask=input):
    """Читает рабочий лист (цветные местные строки) и историю цен. Excel после этого закрыт."""
    _say(progress, f"Открываю Excel «{Path(xlsx).name}»…")
    try:
        wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    except PermissionError:
        raise CheckError(f"Не удаётся открыть «{Path(xlsx).name}». Закройте его в Excel и попробуйте снова.")
    try:
        found = work_sheets(wb.sheetnames)
        if sheet is None and len(found) > 1 and ask is None:
            raise NeedSheet(found)
        try:
            ws_name = choose_work_sheet(wb.sheetnames, sheet, ask=ask)
        except SystemExit as e:
            raise CheckError(str(e))
        _say(progress, f"Читаю лист «{ws_name}»…")
        rows = []
        for rn, r in enumerate(wb[ws_name].iter_rows(min_row=5, values_only=True), start=5):
            if len(r) <= COLS["AP"] or r[COLS["C"]] is None:
                continue
            if not is_colored(r) or str(r[COLS["AA"]]).strip().lower() == "имп":
                continue
            rows.append({
                "rn": rn, "code": r[COLS["C"]], "F": r[COLS["F"]], "G": r[COLS["G"]],
                "H": to_float(r[COLS["H"]]), "I": to_float(r[COLS["I"]]), "J": to_float(r[COLS["J"]]),
                "O": r[COLS["O"]], "Q": r[COLS["Q"]], "R": str(r[COLS["R"]] or "").strip(),
                "Z": r[COLS["Z"]], "AL": to_float(r[COLS["AL"]]), "AM": r[COLS["AM"]],
                "month": month_key(r[COLS["Z"]])})
        # история цен по коду — лист «… (значение)»
        hist = collections.defaultdict(list)
        vs = find_values_sheet(wb)
        if vs:
            _say(progress, f"Читаю историю цен (лист «{vs}»)…")
            for rn, r in enumerate(wb[vs].iter_rows(min_row=5, values_only=True), start=5):
                if len(r) <= COLS["AL"] or r[COLS["C"]] is None:
                    continue
                hist[r[COLS["C"]]].append({"date": r[COLS["Z"]], "supplier": r[COLS["O"]], "R": r[COLS["R"]],
                                           "qty": to_float(r[COLS["H"]]), "unit": r[COLS["G"]],
                                           "price": to_float(r[COLS["AL"]]), "imp": r[COLS["AA"]]})
        return {"xlsx": Path(xlsx), "sheets": found, "ws_name": ws_name, "rows": rows, "hist": hist,
                "months": collections.Counter(x["month"] for x in rows)}
    finally:
        wb.close()


def run_check(folder, data, period, progress=None, report_path=None, excel_out=None, write_excel=True):
    """Сверка цветных строк за период с фактурами из папки. Возвращает всё для отчётов и интерфейса."""
    folder = Path(folder)
    xlsx, ws_name, hist = data["xlsx"], data["ws_name"], data["hist"]
    rows = [x for x in data["rows"] if x["month"] in period]

    # фактуры (только PDF в самой папке)
    pdfs = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")
    invoices, bad_pdfs = [], []
    for i, p in enumerate(pdfs, 1):
        _say(progress, f"Читаю фактуры {i} из {len(pdfs)}…", i - 1, len(pdfs))
        inv = parse_invoice(p)
        (bad_pdfs if inv["error"] else invoices).append(inv)
    _say(progress, "Сверяю строки с фактурами…", len(pdfs), len(pdfs))
    by_num = collections.defaultdict(list)
    for inv in invoices:
        by_num[norm_num(inv["num"])].append(inv)

    # дубли строк
    sig_rows = collections.defaultdict(list)
    for x in rows:
        sig_rows[(x["code"], x["R"], norm_num(x["Q"]), x["H"], x["I"], x["J"])].append(x["rn"])

    results, missing = [], collections.defaultdict(list)
    for x in rows:
        res = {"row": x, "inv": None, "match": None, "warn": []}
        dups = sig_rows[(x["code"], x["R"], norm_num(x["Q"]), x["H"], x["I"], x["J"])]
        if len(dups) > 1:
            res["warn"].append("Дубль строки: такие же строки " + ", ".join(str(n) for n in dups if n != x["rn"]))
        eb, _, _ = unit_info(x["G"])
        if eb == "шт" and x["H"] is not None and abs(x["H"] - round(x["H"])) > 1e-9:
            res["warn"].append(f"Дробное количество в штуках: {fmt(x['H'], 3)} {x['G']}")
        cands = by_num.get(norm_num(strip_brackets(x["R"])), [])
        qn = norm_num(x["Q"])
        same_contract = [i for i in cands if i["contract"] and norm_num(i["contract"]) in qn]
        if same_contract:
            res["inv"] = same_contract[0]
        elif cands:
            res["status"] = "НЕТ ФАКТУРЫ"
            res["other_contract"] = True
            res["note"] = (f"Есть фактура № {x['R']}, но с другим договором "
                           f"({cands[0]['contract']}, {cands[0]['supplier']}). Нужна фактура по договору {x['Q']}.")
        else:
            res["status"] = "НЕТ ФАКТУРЫ"
            res["note"] = "Фактура не найдена в папке."
        if res["inv"]:
            res["match"] = match_row(x, res["inv"])
            res["status"] = res["match"]["status"]
            if res["inv"].get("correction"):
                res["match"]["notes"].append(f"ℹ Фактура исправленная (тузатувчи) к ЭҲФ № {res['inv']['correction']} — "
                                             f"сверка по исправленным цифрам.")
        else:
            missing[(x["O"], x["Q"], x["R"])].append(x["rn"])
        res["short"] = short_problem(res)
        results.append(res)

    out = Path(report_path) if report_path else folder / f"отчет_сверки_{ws_name}.html"
    _say(progress, "Сохраняю отчёт…")
    write_html(out, xlsx.name, ws_name, period, results, missing, bad_pdfs, invoices, hist)

    dst, excel_error = None, None
    if write_excel:
        import excel_writer
        dst = Path(excel_out) if excel_out else xlsx.with_name(f"{xlsx.stem}_проверено{xlsx.suffix}")
        _say(progress, "Записываю результаты в копию Excel…")
        try:
            excel_writer.write_copy(xlsx, dst, ws_name, results, lambda code: hist_stats(code, hist), VERSION)
        except PermissionError:
            excel_error = (f"Не удалось сохранить «{dst.name}» — этот файл открыт в Excel. "
                           f"Закройте его и запустите проверку снова.")
            dst = None
    return {"results": results, "missing": missing, "bad_pdfs": bad_pdfs, "invoices": invoices,
            "report": out, "excel": dst, "excel_error": excel_error, "period": period, "ws_name": ws_name,
            "counts": collections.Counter(r["status"] for r in results)}


def short_problem(r):
    """Коротко, в чём проблема (колонка «Проблема»)."""
    m, f = r.get("match"), set()
    if m:
        f = m["flags"]
    parts = []
    if r["status"] == "НЕТ ФАКТУРЫ":
        parts.append("Фактура другого договора" if r.get("other_contract") else "Нет фактуры")
    if "swap" in f:
        parts.append("Перепутаны строки")
    if "notfound" in f:
        parts.append("Нет такой позиции в фактуре")
    if "other_item" in f:
        parts.append("Другой товар")
    for flag, label in (("qty", "Кол-во"), ("price", "Цена"), ("sum", "Сумма"), ("vat", "Взято с НДС"),
                        ("unit", "Единица измерения"), ("name", "Название не похоже")):
        if flag in f and not (flag == "name" and "other_item" in f):
            parts.append(label)
    if "grouped" in f and m and m.get("item"):
        parts.append(f"Сложено {len(m['item']['members'])} строки фактуры")
    if "sum_other" in f:
        parts.append("Сумма от другой позиции")
    if "price_other" in f:
        parts.append("Цена от другой позиции")
    for w in r["warn"]:
        parts.append("Дубль строки" if w.startswith("Дубль") else "Дробные штуки")
    return ", ".join(parts)


def main():
    ap = argparse.ArgumentParser(description="Сверка Excel с PDF счёт-фактурами")
    ap.add_argument("folder", nargs="?", default=".", help="папка, где лежат Excel и PDF-фактуры")
    ap.add_argument("--excel", help="путь к Excel (если в папке их несколько)")
    ap.add_argument("--period", help="месяцы: 2026-10  или  2026-09,2026-10  или  все")
    ap.add_argument("--sheet", help="имя листа, если в книге их несколько")
    ap.add_argument("-o", "--output", help="куда сохранить HTML-отчёт")
    ap.add_argument("--excel-out", help="куда сохранить копию Excel с результатами")
    ap.add_argument("--no-excel", action="store_true", help="не создавать копию Excel")
    args = ap.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"Папка не найдена: {folder}")
    if args.excel:
        xlsx = Path(args.excel)
    else:
        cands = find_excels(folder)
        if len(cands) != 1:
            sys.exit("В папке должен быть ровно один Excel-файл (или укажите --excel). Найдено: "
                     + (", ".join(p.name for p in cands) or "ни одного"))
        xlsx = cands[0]

    def progress(text, done=None, total=None):
        if done is not None and total:
            print(f"\r{text}", end="", flush=True)
        else:
            print(text)

    try:
        data = load_excel(xlsx, args.sheet, progress)
    except CheckError as e:
        sys.exit(str(e))
    print(f"Рабочий лист: {data['ws_name']}")
    period = ask_period(data["months"], args.period)
    print(f"Цветных строк за период ({', '.join(month_name(p) for p in sorted(period))}): "
          f"{sum(data['months'][p] for p in period)}")
    res = run_check(folder, data, period, progress, args.output, args.excel_out, not args.no_excel)
    print("\nИтог: " + ", ".join(f"{k}: {v}" for k, v in res["counts"].most_common()))
    print(f"Отчёт: {res['report']}")
    if res["excel"]:
        print(f"Excel: {res['excel']}")
    if res["excel_error"]:
        print(res["excel_error"])


# ---------------------------------------------------------------- HTML-отчёт
STATUS_CLS = {"OK": "ok", "ПРОВЕРИТЬ": "warn", "ОШИБКА": "bad", "ПЕРЕПУТАНЫ СТРОКИ": "bad",
              "ПОЗИЦИЯ НЕ НАЙДЕНА": "bad", "НЕТ ФАКТУРЫ": "none"}
SEVERITY = {"ПЕРЕПУТАНЫ СТРОКИ": 0, "ОШИБКА": 1, "ПОЗИЦИЯ НЕ НАЙДЕНА": 2, "ПРОВЕРИТЬ": 3, "OK": 5, "НЕТ ФАКТУРЫ": 6}


def is_problem(r):
    return r["status"] in ("ПЕРЕПУТАНЫ СТРОКИ", "ОШИБКА", "ПОЗИЦИЯ НЕ НАЙДЕНА", "ПРОВЕРИТЬ") or r["warn"]


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if isinstance(d, dt.datetime) else esc(d)


def hist_stats(code, hist):
    """Мин / макс / средняя / медиана цены кода по всем приходам листа «… (значение)»."""
    prices = [x["price"] for x in hist.get(code, []) if x["price"]]
    if not prices:
        return None
    srt = sorted(prices)
    n = len(srt)
    med = srt[n // 2] if n % 2 else (srt[n // 2 - 1] + srt[n // 2]) / 2
    return {"n": n, "min": srt[0], "max": srt[-1], "avg": sum(srt) / n, "med": med}


def price_history_html(code, hist, cur_rn_price):
    h = [x for x in hist.get(code, []) if x["price"]]
    if not h:
        return "<p class='muted'>Нет других приходов этого кода на листе «значение».</p>"
    st = hist_stats(code, hist)
    prices = [x["price"] for x in h]
    avg, med = st["avg"], st["med"]
    head = (f"<div class='hist-stats'><span>Приходов: <b>{len(h)}</b></span>"
            f"<span>Мин: <b>{fmt(min(prices))}</b></span><span>Макс: <b>{fmt(max(prices))}</b></span>"
            f"<span>Средняя: <b>{fmt(avg)}</b></span><span>Медиана: <b>{fmt(med)}</b></span>")
    if cur_rn_price:
        dev = cur_rn_price / med - 1 if med else 0
        head += f"<span>Цена в строке от медианы: <b class='{ 'neg' if dev < 0 else 'pos'}'>{dev:+.0%}</b></span>"
    head += "</div>"
    h = sorted(h, key=lambda x: x["date"] if isinstance(x["date"], dt.datetime) else dt.datetime.min)[-12:]
    rows = "".join(
        f"<tr><td>{fmt_date(x['date'])}</td><td>{esc(x['supplier'])}</td><td>{esc(x['R'])}</td>"
        f"<td class='n'>{fmt(x['qty'], 3)} {esc(x['unit'])}</td><td class='n'>{fmt(x['price'])}</td></tr>"
        for x in reversed(h))
    return (head + "<details><summary>Последние приходы (до 12)</summary><table class='hist'><tr><th>Дата</th>"
            "<th>Поставщик</th><th>Номер вх.</th><th>Кол-во</th><th>Цена без НДС, сум</th></tr>"
            + rows + "</table></details>")


def write_html(out, xlsx_name, ws_name, period, results, missing, bad_pdfs, invoices, hist):
    c = collections.Counter(r["status"] for r in results)
    problems = [r for r in results if is_problem(r)]
    by_code = collections.OrderedDict()
    for r in sorted(problems, key=lambda r: (min(SEVERITY.get(r["status"], 9), 4 if r["warn"] else 9),
                                             str(r["row"]["code"]))):
        by_code.setdefault(r["row"]["code"], []).append(r)

    def row_block(r):
        x, m = r["row"], r["match"]
        cls = STATUS_CLS.get(r["status"], "warn")
        inv = r["inv"]
        parts = []
        if m:
            parts += [f"<li>{esc(n)}</li>" for n in m["notes"]]
        if r.get("note"):
            parts.append(f"<li>{esc(r['note'])}</li>")
        parts += [f"<li class='w'>⚠ {esc(w)}</li>" for w in r["warn"]]
        inv_line = ""
        if inv and m and m["item"]:
            it = m["item"]
            inv_line = (f"<div class='inv'>Фактура <b>{esc(inv['num'])}</b> от {esc(inv['date'])} · "
                        f"{esc(inv['supplier'])} · договор {esc(inv['contract'])} · файл {esc(inv['file'])}<br>"
                        f"Поз. {it['no']}: «{esc(it['name'])}» — {fmt(it['qty'], 3)} {esc(it['unit'])} × "
                        f"{fmt(it['price'])} = {fmt(it['sum'])} (с НДС {fmt(it['total'])})"
                        f" · похожесть названия {m['sim']:.0%}</div>")
        status = r["status"] if r["status"] not in ("OK",) else "OK, но есть замечания"
        return (f"<div class='row {cls}'><div class='rh'><span class='st'>{esc(status)}</span>"
                f"<span>Строка <b>{x['rn']}</b></span><span>{esc(x['O'])}</span><span>{esc(x['Q'])}</span>"
                f"<span>№ вх. <b>{esc(x['R'])}</b></span><span>{fmt_date(x['Z'])}</span></div>"
                f"<div class='ex'>Excel: {fmt(x['H'], 3)} {esc(x['G'])} × {fmt(x['I'])} = {fmt(x['J'])}</div>"
                f"{inv_line}<ul>{''.join(parts)}</ul></div>")

    code_html = []
    for code, rs in by_code.items():
        x0 = rs[0]["row"]
        worst = min(rs, key=lambda r: SEVERITY.get(r["status"], 9))["status"]
        code_html.append(
            f"<section class='code {STATUS_CLS.get(worst, 'warn')}' data-search='{esc(str(code).lower())} "
            f"{esc(str(x0['F']).lower())}'><h3><span class='cd'>{esc(code)}</span> {esc(x0['F'])}</h3>"
            + "".join(row_block(r) for r in rs)
            + "<div class='histbox'><h4>Цены этого кода в других приходах</h4>"
            + price_history_html(code, hist, x0["AL"]) + "</div></section>")

    miss_rows = "".join(
        f"<tr><td>{esc(o)}</td><td>{esc(q)}</td><td><b>{esc(rr)}</b></td><td>{len(v)}</td>"
        f"<td class='small'>{', '.join(map(str, v[:15]))}{' …' if len(v) > 15 else ''}</td></tr>"
        for (o, q, rr), v in sorted(missing.items(), key=lambda kv: (str(kv[0][0]), str(kv[0][2]))))
    bad_rows = "".join(f"<tr><td>{esc(b['file'])}</td><td>{esc(b['error'])}</td></tr>" for b in bad_pdfs)

    stat = lambda k, cls: f"<div class='stat {cls}'><b>{c.get(k, 0)}</b><span>{k.lower()}</span></div>"
    now = dt.datetime.now().strftime("%d.%m.%Y %H:%M")
    doc = f"""<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Проблемные коды</title>
<style>
:root{{--bg:#05060f;--card:rgba(18,20,48,.72);--line:rgba(160,150,255,.2);--tx:#e9ecff;--mu:#9aa3c7;
--ok:#34d399;--bad:#fb7185;--warn:#fbbf24;--none:#94a3b8;--cy:#a5f3fc}}
*{{box-sizing:border-box}}
body{{margin:0;color:var(--tx);font:15px/1.5 "Segoe UI",system-ui,Arial,sans-serif;
background:radial-gradient(1100px 600px at 10% -10%,rgba(139,92,246,.35),transparent 60%),
radial-gradient(900px 500px at 95% 10%,rgba(236,72,153,.2),transparent 60%),
radial-gradient(1000px 700px at 60% 110%,rgba(59,130,246,.25),transparent 60%),var(--bg);background-attachment:fixed}}
main{{max-width:1150px;margin:0 auto;padding:32px 16px 60px}}
h1{{margin:0;font-size:30px;background:linear-gradient(90deg,#a5f3fc,#c4b5fd 45%,#f9a8d4);
-webkit-background-clip:text;background-clip:text;color:transparent}}
.sub{{color:var(--mu);margin:4px 0 22px}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:20px;margin-bottom:20px}}
.stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}}
.stat{{border:1px solid var(--line);border-radius:12px;padding:12px 14px;background:rgba(255,255,255,.03)}}
.stat b{{display:block;font-size:28px}} .stat span{{color:var(--mu)}}
.stat.ok b{{color:var(--ok)}} .stat.bad b{{color:var(--bad)}} .stat.warn b{{color:var(--warn)}} .stat.none b{{color:var(--none)}}
h2{{margin:0 0 12px;font-size:21px}}
.code{{border:1px solid var(--line);border-left:4px solid var(--warn);border-radius:14px;padding:16px 18px;
margin-bottom:16px;background:var(--card)}}
.code.bad{{border-left-color:var(--bad)}} .code.ok{{border-left-color:var(--ok)}}
.code h3{{margin:0 0 10px;font-size:17px}} .cd{{font-family:Consolas,monospace;color:var(--cy);margin-right:6px}}
.row{{border-top:1px dashed var(--line);padding:10px 0}}
.rh{{display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center;color:var(--mu)}} .rh b{{color:var(--tx)}}
.st{{font-weight:700;padding:2px 10px;border-radius:999px;font-size:13px;letter-spacing:.3px}}
.row.bad .st{{background:rgba(251,113,133,.16);color:var(--bad)}}
.row.warn .st{{background:rgba(251,191,36,.14);color:var(--warn)}}
.row.ok .st{{background:rgba(52,211,153,.14);color:var(--ok)}}
.row.none .st{{background:rgba(148,163,184,.16);color:var(--none)}}
.ex{{margin-top:6px;font-family:Consolas,monospace}}
.inv{{margin-top:4px;color:#c7d2fe;font-size:14px}}
ul{{margin:6px 0 0 18px;padding:0}} li{{margin:2px 0}} li.w{{color:var(--warn)}}
.histbox{{margin-top:10px;padding:10px 12px;border-radius:10px;background:rgba(255,255,255,.03)}}
.histbox h4{{margin:0 0 6px;font-size:14px;color:var(--mu)}}
.hist-stats{{display:flex;flex-wrap:wrap;gap:6px 16px}} .neg{{color:var(--bad)}} .pos{{color:var(--warn)}}
table{{width:100%;border-collapse:collapse;font-size:14px;margin-top:8px}}
th,td{{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}}
th{{color:#ddd6fe;background:rgba(139,92,246,.14)}} td.n{{text-align:right;font-family:Consolas,monospace}}
.small{{font-size:12px;color:var(--mu)}} .muted{{color:var(--mu)}}
summary{{cursor:pointer;color:var(--cy);margin-top:6px}}
input{{width:100%;padding:12px 14px;border-radius:10px;border:1px solid var(--line);background:rgba(0,0,0,.3);
color:var(--tx);font:inherit;margin-bottom:14px}}
.legend li{{margin:3px 0}}
</style></head><body><main>
<h1>Проблемные коды</h1>
<p class="sub">{esc(xlsx_name)} · лист «{esc(ws_name)}» · период: {esc(', '.join(month_name(p) for p in sorted(period)))}
 · фактур в папке: {len(invoices)} · сформировано {now} · версия программы {VERSION}</p>
<div class="card"><div class="stats">
<div class="stat"><b>{len(results)}</b><span>цветных строк</span></div>
{stat('OK', 'ok')}{stat('ПЕРЕПУТАНЫ СТРОКИ', 'bad')}{stat('ОШИБКА', 'bad')}{stat('ПОЗИЦИЯ НЕ НАЙДЕНА', 'bad')}
{stat('ПРОВЕРИТЬ', 'warn')}{stat('НЕТ ФАКТУРЫ', 'none')}
<div class="stat warn"><b>{len(by_code)}</b><span>кодов с проблемами</span></div>
</div>
<ul class="legend small" style="margin-top:14px">
<li><b>ПЕРЕПУТАНЫ СТРОКИ</b> — числа из Excel принадлежат в фактуре другому товару.</li>
<li><b>ОШИБКА</b> — товар найден, но кол-во / цена / сумма отличаются (или взято с НДС).</li>
<li><b>ПОЗИЦИЯ НЕ НАЙДЕНА</b> — в фактуре нет такого товара ни по названию, ни по числам.</li>
<li><b>ПРОВЕРИТЬ</b> — числа совпадают, но другая единица измерения или непохожее название.</li>
<li>⚠ — дубли строк и дробные штуки.</li></ul></div>

<div class="card"><h2>Коды с проблемами ({len(by_code)})</h2>
<input id="q" placeholder="Поиск по коду или названию…" oninput="
var v=this.value.toLowerCase();document.querySelectorAll('.code').forEach(function(s){{
s.style.display=s.dataset.search.indexOf(v)>=0?'':'none'}})">
{''.join(code_html) or "<p>Проблем не найдено.</p>"}
</div>

<div class="card"><h2>Нет фактуры — скачать из Didox ({len(missing)})</h2>
<table><tr><th>Поставщик</th><th>Договор</th><th>Номер</th><th>Строк</th><th>Строки Excel</th></tr>{miss_rows}</table></div>

<div class="card"><h2>PDF, не соответствующие стандарту ({len(bad_pdfs)})</h2>
{('<table><tr><th>Файл</th><th>Причина</th></tr>' + bad_rows + '</table>') if bad_pdfs else '<p class="muted">Все PDF прочитаны.</p>'}
</div>
</main></body></html>"""
    out.write_text(doc, encoding="utf-8")


if __name__ == "__main__":
    main()
