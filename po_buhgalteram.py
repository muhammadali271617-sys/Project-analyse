"""
Ошибки по бухгалтерам: из проверенной книги (…_проверено.xlsx) делает НОВУЮ книгу,
где у каждого бухгалтера (колонка X «Ответственный») свой лист с его строками.

Отбор: дата в Z — в выбранном месяце, колонка AX («Проблема (коротко)») не пустая.
Каждый лист — точная копия листа проверки: тот же XML листа (оформление, ширины,
скрытые/сгруппированные колонки, высоты строк, условное форматирование, автофильтр),
в нём оставлены только строки бухгалтера, перенумерованные подряд с 5-й строки.
Исходный файл не меняется.

Запуск:
    py po_buhgalteram.py "D:\\...\\Анализ_Цен_AEZ_октябрь__разницалар_проверено.xlsx" --month 2026-10
Результат: «Ошибки_по_бухгалтерам_<месяц>_<год>.xlsx» рядом с исходником.
"""

import argparse
import collections
import datetime as dt
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, unescape

import openpyxl
from openpyxl.formula.translate import Translator

FIRST_DATA_ROW = 5
COL_X, COL_Z, COL_AX = 24, 26, 50          # 1-based: X, Z, AX
MONTHS_RU = ["", "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
             "сентябрь", "октябрь", "ноябрь", "декабрь"]

ROW_RX = re.compile(r'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RX = re.compile(r'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', re.S)
F_RX = re.compile(r"<f\b([^>]*?)(?:/>|>([^<]*)</f>)")
ENT = {"&quot;": '"', "&apos;": "'"}


# ---------------------------------------------------------------- пакет xlsx
def sheet_parts(z):
    """[(имя листа, путь к XML, r:id)] в порядке книги."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'\bId="([^"]+)"', m.group(0)).group(1)
        t = re.search(r'\bTarget="([^"]+)"', m.group(0)).group(1)
        targets[rid] = t.lstrip("/") if t.startswith("/xl/") else "xl/" + t
    out = []
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        name = unescape(re.search(r'\bname="([^"]*)"', m.group(0)).group(1), ENT)
        rid = re.search(r'\br:id="([^"]+)"', m.group(0)).group(1)
        out.append((name, targets[rid], rid))
    return out


def find_check_sheet(names, wanted=None):
    if wanted:
        if wanted not in names:
            sys.exit(f"Лист «{wanted}» не найден. Листы: {', '.join(names)}")
        return wanted
    found = [n for n in names if n.lower().strip().endswith(("_проверка", "_разницалар"))]
    if len(found) != 1:
        sys.exit("Укажите лист проверки (--sheet). Подходящие: " + (", ".join(found) or "нет"))
    return found[0]


# ---------------------------------------------------------------- отбор строк
def month_of(v):
    if isinstance(v, dt.datetime):
        return f"{v.year}-{v.month:02d}"
    if isinstance(v, dt.date):
        return f"{v.year}-{v.month:02d}"
    if isinstance(v, str):
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", v)
        if m:
            return f"{m.group(3)}-{m.group(2)}"
    return None


def read_candidates(path, sheet):
    """Строки с заполненной «Проблемой» (AX): [(номер строки, логин, месяц ГГГГ-ММ)]."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    try:
        for rn, r in enumerate(wb[sheet].iter_rows(min_row=FIRST_DATA_ROW, values_only=True), start=FIRST_DATA_ROW):
            if len(r) < COL_AX:
                continue
            problem = r[COL_AX - 1]
            if problem is None or not str(problem).strip():
                continue
            month = month_of(r[COL_Z - 1])
            if month:
                out.append((rn, str(r[COL_X - 1] or "").strip() or "(без ответственного)", month))
    finally:
        wb.close()
    return out


def month_name(k):
    y, m = k.split("-")
    return f"{MONTHS_RU[int(m)]} {y}"


def available_months(cands):
    """{месяц: число строк} по возрастанию."""
    c = collections.Counter(m for _, _, m in cands)
    return dict(sorted(c.items()))


def group_rows(cands, months):
    groups = collections.defaultdict(list)
    for rn, login, m in cands:
        if m in months:
            groups[login].append(rn)
    return groups


def output_name(src, months):
    ms = sorted(months)
    years = {m[:4] for m in ms}
    if len(years) == 1:
        label = "_".join(MONTHS_RU[int(m[5:])] for m in ms) + "_" + ms[0][:4]
    else:
        label = "_".join(f"{MONTHS_RU[int(m[5:])]}_{m[:4]}" for m in ms)
    return Path(src).with_name(f"Ошибки_по_бухгалтерам_{label}.xlsx")


def ask_months(avail):
    keys = list(avail)
    print("\nЗа какой период? Строки с проблемами есть за месяцы:")
    for i, k in enumerate(keys, 1):
        print(f"  {i}. {month_name(k):<16} ({avail[k]} строк)")
    print("Введите номер (один месяц), несколько через запятую (например 1,2) или «все».")
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


def parse_months(arg, avail):
    if arg.strip().lower() in ("все", "all", "*"):
        return set(avail)
    months = {m.strip() for m in arg.split(",") if m.strip()}
    bad = [m for m in months if not re.fullmatch(r"\d{4}-\d{2}", m)]
    if bad:
        sys.exit("Месяц нужно указать как ГГГГ-ММ, например 2026-10 (несколько — через запятую)")
    return months


def make_book(src, sheet=None, months=None, dst=None, ask=None):
    """Для интерфейса и командной строки: создаёт книгу и проверяет её.
    Возвращает (путь, [(логин, число строк)], замечания проверки)."""
    src = Path(src)
    with zipfile.ZipFile(src) as z:
        names = [n for n, _, _ in sheet_parts(z)]
    sheet = find_check_sheet(names, sheet)
    cands = read_candidates(src, sheet)
    groups = group_rows(cands, months)
    if not groups:
        raise ValueError("За выбранный период нет строк с заполненной «Проблемой» (AX).")
    dst = Path(dst) if dst else output_name(src, months)
    order = build_book(src, dst, sheet, groups)
    return dst, [(login, len(rows)) for login, rows in order], verify(dst, order, months)


# ---------------------------------------------------------------- лист
def col_idx(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def unshare_all(sheet):
    """Общие (протянутые) формулы → обычные, чтобы строки можно было переставлять."""
    masters = {}
    for cm in CELL_RX.finditer(sheet):
        fm = F_RX.search(cm.group(0))
        if fm and 't="shared"' in fm.group(1) and "ref=" in fm.group(1):
            si = re.search(r'\bsi="(\d+)"', fm.group(1)).group(1)
            masters[si] = (cm.group(1) + cm.group(2), unescape(fm.group(2) or "", ENT))

    def fix(cm):
        cell = cm.group(0)
        fm = F_RX.search(cell)
        if not fm or 't="shared"' not in fm.group(1):
            return cell
        si = re.search(r'\bsi="(\d+)"', fm.group(1)).group(1)
        if si not in masters:
            return cell
        origin, formula = masters[si]
        text = Translator("=" + formula, origin=origin).translate_formula(cm.group(1) + cm.group(2))[1:]
        return cell[:fm.start()] + f"<f>{escape(text)}</f>" + cell[fm.end():]

    return CELL_RX.sub(fix, sheet) if masters else sheet


def move_row(row_xml, old, new):
    """Строка old → new: номер строки, адреса ячеек, формулы (сдвиг относительных ссылок)."""
    row_xml = re.sub(r'(<row\b[^>]*?\br=")\d+(")', rf"\g<1>{new}\2", row_xml, count=1)
    row_xml = re.sub(r'(<row\b[^>]*?)\s+hidden="(?:1|true)"', r"\1", row_xml, count=1)   # фильтр не наследуем

    def cell(cm):
        c = cm.group(0)
        letters = cm.group(1)
        c = re.sub(rf'(\br="){letters}{old}"', rf'\g<1>{letters}{new}"', c, count=1)
        fm = F_RX.search(c)
        if fm and fm.group(2) and old != new:
            text = unescape(fm.group(2), ENT)
            moved = Translator("=" + text, origin=f"{letters}{old}").translate_formula(f"{letters}{new}")[1:]
            c = c[:fm.start()] + f"<f{fm.group(1)}>{escape(moved)}</f>" + c[fm.end():]
        return c

    return CELL_RX.sub(cell, row_xml)


def build_sheet(sheet, rows_keep, first_tab):
    """XML листа, в котором остались строки 1–4 и rows_keep (перенумерованные с 5-й)."""
    a = sheet.index("<sheetData")
    a_end = sheet.index(">", a) + 1
    if sheet[a_end - 2] == "/":                    # <sheetData/>
        body, b = "", a_end
    else:
        b = sheet.index("</sheetData>")
        body = sheet[a_end:b]
    head, tail = sheet[:a_end if sheet[a_end - 2] != "/" else a], sheet[b:]
    rows = {int(m.group(1)): m.group(0) for m in ROW_RX.finditer(body)}

    out = [move_row(rows[r], r, r) for r in sorted(rows) if r < FIRST_DATA_ROW]
    new = FIRST_DATA_ROW
    for r in rows_keep:
        out.append(move_row(rows[r], r, new))
        new += 1
    last = new - 1
    max_col = 1
    for m in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"', "".join(out)):
        max_col = max(max_col, col_idx(m.group(1)))
    last_col = col_letters(max_col)

    if sheet[a_end - 2] == "/":
        head = head + "<sheetData>"
    xml = head + "".join(out) + tail

    # шапка листа
    xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{last_col}{last}"/>', xml, count=1)
    xml = re.sub(r'(<sheetPr\b[^>]*?)\s+filterMode="[^"]*"', r"\1", xml, count=1)
    xml = re.sub(r'(<sheetView\b[^>]*?)\s+tabSelected="[^"]*"', r"\1", xml)
    xml = re.sub(r'(<sheetView\b[^>]*?)\s+topLeftCell="[^"]*"', r"\1", xml)
    xml = re.sub(r"<selection\b[^>]*/>", "", xml)
    if first_tab:
        xml = re.sub(r"<sheetView\b", '<sheetView tabSelected="1"', xml, count=1)

    # автофильтр A4:<последняя колонка><последняя строка>, без условий — видны все строки
    af = re.search(r'<autoFilter ref="([A-Z]+)(\d+):[A-Z]+\d+"[^>]*?(?:/>|>.*?</autoFilter>)', xml, re.S)
    ref = f"{af.group(1)}{af.group(2)}:{last_col}{max(last, int(af.group(2)))}" if af else f"A4:{last_col}{last}"
    if af:
        xml = xml[:af.start()] + f'<autoFilter ref="{ref}"/>' + xml[af.end():]
    # объединения ячеек внутри данных теряют смысл — оставляем только в шапке
    def merges(m):
        keep = [x for x in re.findall(r'<mergeCell ref="([^"]+)"/>', m.group(0))
                if max(int(n) for n in re.findall(r"\d+", x)) < FIRST_DATA_ROW]
        return (f'<mergeCells count="{len(keep)}">' + "".join(f'<mergeCell ref="{k}"/>' for k in keep)
                + "</mergeCells>") if keep else ""
    xml = re.sub(r"<mergeCells\b.*?</mergeCells>", merges, xml, flags=re.S)
    # ссылки на чужие части (принтер, рисунки, гиперссылки) в новой книге не нужны
    xml = re.sub(r'(<pageSetup\b[^>]*?)\s+r:id="[^"]*"', r"\1", xml)
    xml = re.sub(r"<(drawing|legacyDrawing|legacyDrawingHF|picture)\b[^>]*/>", "", xml)
    xml = re.sub(r"<hyperlinks>.*?</hyperlinks>", "", xml, flags=re.S)
    return xml, ref, last


# ---------------------------------------------------------------- новая книга
def build_book(src, dst, sheet_name, groups):
    with zipfile.ZipFile(src) as z:
        parts = sheet_parts(z)
        sp = next(p for n, p, _ in parts if n == sheet_name)
        template = unshare_all(z.read(sp).decode("utf-8"))

        order = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
        sheets_xml, sheet_entries, rel_entries, ct_entries, names = {}, [], [], [], []
        used = set()
        for i, (login, rows) in enumerate(order):
            title = re.sub(r"[\[\]:*?/\\]", "_", login)[:31] or f"Лист{i + 1}"
            base, k = title, 2
            while title.lower() in used:
                title = f"{base[:28]}_{k}"
                k += 1
            used.add(title.lower())
            xml, ref, _ = build_sheet(template, rows, i == 0)
            path = f"xl/worksheets/sheet{i + 1}.xml"
            sheets_xml[path] = xml
            rid = f"rIdS{i + 1}"
            sheet_entries.append(f'<sheet name="{escape(title, ENT_REV)}" sheetId="{i + 1}" r:id="{rid}"/>')
            rel_entries.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
                               f'2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>')
            ct_entries.append(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-'
                              f'officedocument.spreadsheetml.worksheet+xml"/>')
            c1, r1, c2, r2 = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref).groups()
            q = "'" + title.replace("'", "''") + "'"
            names.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
                         f'{escape(q)}!${c1}${r1}:${c2}${r2}</definedName>')

        # workbook.xml: только новые листы и их фильтры; связи с удалёнными частями убираем
        wb = z.read("xl/workbook.xml").decode("utf-8")
        wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(sheet_entries) + "</sheets>", wb, flags=re.S)
        wb = re.sub(r"<definedNames>.*?</definedNames>", "", wb, flags=re.S)
        wb = re.sub(r"(</sheets>)", r"\1<definedNames>" + "".join(names).replace("\\", "\\\\") + "</definedNames>", wb, count=1)
        wb = re.sub(r"<(externalReferences|pivotCaches)>.*?</\1>", "", wb, flags=re.S)
        wb = re.sub(r'(<workbookView\b[^>]*?)\s+(activeTab|firstSheet)="[^"]*"', r"\1", wb)
        wb = re.sub(r'(<workbookView\b[^>]*?)\s+(activeTab|firstSheet)="[^"]*"', r"\1", wb)

        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        keep_types = ("styles", "theme", "sharedStrings")
        kept = [m.group(0) for m in re.finditer(r"<Relationship\b[^>]*/>", rels)
                if any(f"/relationships/{t}\"" in m.group(0) for t in keep_types)]
        rels = re.sub(r"<Relationship\b[^>]*/>", "", rels)
        rels = rels.replace("</Relationships>", "".join(kept + rel_entries) + "</Relationships>")

        drop_prefix = ("xl/worksheets/", "xl/calcChain.xml", "xl/drawings/", "xl/media/", "xl/printerSettings/",
                       "xl/pivotTables/", "xl/pivotCache/", "xl/externalLinks/", "xl/tables/", "xl/comments",
                       "xl/threadedComments/", "xl/persons/", "docProps/app.xml")
        ct = z.read("[Content_Types].xml").decode("utf-8")
        ct = re.sub(r'<Override\b[^>]*PartName="/([^"]+)"[^>]*/>',
                    lambda m: "" if m.group(1).startswith(drop_prefix) else m.group(0), ct)
        ct = ct.replace("</Types>", "".join(ct_entries) + "</Types>")
        root_rels = z.read("_rels/.rels").decode("utf-8")
        root_rels = re.sub(r'<Relationship\b[^>]*Target="/?docProps/app\.xml"[^>]*/>', "", root_rels)

        replace = {"xl/workbook.xml": wb, "xl/_rels/workbook.xml.rels": rels, "[Content_Types].xml": ct,
                   "_rels/.rels": root_rels}
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
            for info in z.infolist():
                n = info.filename
                if n.startswith(drop_prefix):
                    continue
                data = replace[n].encode("utf-8") if n in replace else z.read(n)
                out.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
            for path, xml in sheets_xml.items():
                out.writestr(path, xml.encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
    return order


ENT_REV = {'"': "&quot;"}


# ---------------------------------------------------------------- проверка результата
def verify(dst, order, months):
    wb = openpyxl.load_workbook(dst, read_only=False)
    problems = []
    for (login, rows), ws in zip(order, wb.worksheets):
        data = list(ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=False))
        if len(data) != len(rows):
            problems.append(f"{ws.title}: строк {len(data)}, ожидалось {len(rows)}")
        for i, row in enumerate(data, start=FIRST_DATA_ROW):
            x, z, ax = row[COL_X - 1].value, row[COL_Z - 1].value, row[COL_AX - 1].value
            if str(x or "").strip() != login and login != "(без ответственного)":
                problems.append(f"{ws.title}!X{i}: {x}")
            if month_of(z) not in months:
                problems.append(f"{ws.title}!Z{i}: {z}")
            if ax is None or not str(ax).strip():
                problems.append(f"{ws.title}!AX{i}: пусто")
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    if "#REF" in v:
                        problems.append(f"{ws.title}!{c.coordinate}: {v}")
                    refs = {int(n) for n in re.findall(r"[A-Z]+\$?(\d+)", v)}
                    if refs and refs != {i}:
                        problems.append(f"{ws.title}!{c.coordinate}: ссылки на строки {sorted(refs)}")
        if ws.auto_filter.ref is None:
            problems.append(f"{ws.title}: нет автофильтра")
        if not ws.conditional_formatting:
            problems.append(f"{ws.title}: нет условного форматирования")
    return problems


def main():
    ap = argparse.ArgumentParser(description="Ошибки по бухгалтерам — отдельный лист на каждого")
    ap.add_argument("excel", help="проверенная книга …_проверено.xlsx")
    ap.add_argument("--month", help="ГГГГ-ММ, несколько через запятую или «все» (по умолчанию — спросит)")
    ap.add_argument("--sheet", help="лист проверки (по умолчанию — «<месяц>_проверка/_разницалар»)")
    ap.add_argument("-o", "--output", help="куда сохранить новую книгу")
    args = ap.parse_args()

    src = Path(args.excel)
    if not src.exists():
        sys.exit(f"Файл не найден: {src}")
    with zipfile.ZipFile(src) as z:
        names = [n for n, _, _ in sheet_parts(z)]
    sheet = find_check_sheet(names, args.sheet)
    print(f"Читаю «{src.name}», лист «{sheet}»…")
    cands = read_candidates(src, sheet)
    avail = available_months(cands)
    if not avail:
        sys.exit("В книге нет строк с заполненной «Проблемой» (AX). Сначала выполните сверку.")
    months = parse_months(args.month, avail) if args.month else ask_months(avail)

    try:
        dst, counts, problems = make_book(src, sheet, months, args.output)
    except ValueError as e:
        sys.exit(str(e))
    except PermissionError:
        sys.exit("Не удалось сохранить книгу — закройте её в Excel и запустите снова.")

    print(f"\nГотово: {dst}")
    for login, n in counts:
        print(f"  {login:<28} {n:>4}")
    print(f"  {'всего':<28} {sum(n for _, n in counts):>4}")
    print("\nПроверка: " + ("всё в порядке." if not problems else f"найдено {len(problems)} замечаний:"))
    for p in problems[:30]:
        print("  •", p)


if __name__ == "__main__":
    main()
