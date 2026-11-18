"""
Запись результатов сверки в КОПИЮ Excel-файла.

Исходный файл не меняется. В копии меняются только:
  * лист проверки — добавляются колонки AV…BM в проверенных строках и заголовки в строке 2;
  * styles.xml — добавляются несколько стилей (цвета статусов, проценты, числа).
Всё остальное (формулы, условное форматирование, фильтры, другие листы, картинки)
копируется байт в байт — книгу не пересохраняет никакая библиотека.
"""

import re
import zipfile
from xml.sax.saxutils import escape

FIRST_COL = 48  # AV

HEADERS = [
    ("AV", "Сверка: статус", 18),
    ("AW", "Сверка: описание", 90),
    ("AX", "Проблема (коротко)", 26),
    ("AY", "Название в фактуре", 45),
    ("AZ", "Название совпадает", 12),
    ("BA", "Похожесть названия", 11),
    ("BB", "Приходов кода (все)", 11),
    ("BC", "Мин. цена кода", 15),
    ("BD", "Макс. цена кода", 15),
    ("BE", "Средняя цена кода", 15),
    ("BF", "Медиана цены кода", 15),
    ("BG", "Цена строки от медианы", 12),
    ("BH", "Фактура (файл, позиция)", 22),
    ("BI", "Ед. изм. в фактуре", 14),
    ("BJ", "Кол-во в фактуре", 13),
    ("BK", "Цена в фактуре (без НДС)", 15),
    ("BL", "Сумма в фактуре (без НДС)", 17),
    ("BM", "Проверка (H=BJ, I=BK, J=BL)", 13),
]
LAST_COL = FIRST_COL + len(HEADERS) - 1  # BJ = 62

SIM_TRUE = 0.5  # с какой похожести названия считаем «ИСТИНА»


def col_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def xml_text(s):
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(s))
    return escape(s)


# ------------------------------------------------------------------ styles.xml
def add_styles(styles):
    """Добавляет шрифт, заливки и форматы ячеек. Возвращает (новый xml, словарь стилей)."""
    def count(tag):
        return int(re.search(rf"<{tag}\b[^>]*\bcount=\"(\d+)\"", styles).group(1))

    def append(xml, tag, items):
        n = count(tag)
        xml = re.sub(rf"(<{tag}\b[^>]*\bcount=\")(\d+)(\")", lambda m: f"{m.group(1)}{n + len(items)}{m.group(3)}",
                     xml, count=1)
        xml = xml.replace(f"</{tag}>", "".join(items) + f"</{tag}>", 1)
        return xml, n

    bold = ('<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/>'
            '<charset val="204"/><scheme val="minor"/></font>')
    styles, f_bold = append(styles, "fonts", [bold])

    colors = {"ok": "FFC6EFCE", "bad": "FFFFC7CE", "warn": "FFFFEB9C", "none": "FFE7E6E6", "head": "FFE4DFEC"}
    fills = [f'<fill><patternFill patternType="solid"><fgColor rgb="{c}"/><bgColor indexed="64"/>'
             f'</patternFill></fill>' for c in colors.values()]
    styles, f0 = append(styles, "fills", fills)
    fill_id = {k: f0 + i for i, k in enumerate(colors)}

    xfs = {
        "ok":   f'<xf numFmtId="0" fontId="{f_bold}" fillId="{fill_id["ok"]}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
        "bad":  f'<xf numFmtId="0" fontId="{f_bold}" fillId="{fill_id["bad"]}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
        "warn": f'<xf numFmtId="0" fontId="{f_bold}" fillId="{fill_id["warn"]}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
        "none": f'<xf numFmtId="0" fontId="{f_bold}" fillId="{fill_id["none"]}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
        "text": '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        "pct":  '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        "num":  '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        "int":  '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        "head": (f'<xf numFmtId="0" fontId="{f_bold}" fillId="{fill_id["head"]}" borderId="0" xfId="0" '
                 f'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" '
                 f'vertical="center" wrapText="1"/></xf>'),
    }
    styles, x0 = append(styles, "cellXfs", list(xfs.values()))
    return styles, {k: x0 + i for i, k in enumerate(xfs)}


# ------------------------------------------------------------------ ячейки
def c_str(ref, s, st):
    return f'<c r="{ref}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{xml_text(s)}</t></is></c>'


def c_num(ref, v, st):
    return f'<c r="{ref}" s="{st}"><v>{repr(float(v))}</v></c>'


def c_bool(ref, v, st):
    return f'<c r="{ref}" s="{st}" t="b"><v>{1 if v else 0}</v></c>'


CELL_RX = re.compile(r'<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', re.S)
ROW_RX = re.compile(r'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)


def patch_row(row_xml, new_cells):
    """new_cells: {номер колонки: xml}. Старые ячейки в этих колонках заменяются."""
    m = re.match(r"<row\b([^>]*?)(/?)>", row_xml)
    attrs = re.sub(r'\s+spans="[^"]*"', "", m.group(1))
    inner = "" if m.group(2) else row_xml[m.end():-len("</row>")]
    cells = [(col_idx(cm.group(1)), cm.group(0)) for cm in CELL_RX.finditer(inner)]
    cells = [c for c in cells if c[0] not in new_cells] + list(new_cells.items())
    cells.sort(key=lambda c: c[0])
    return f"<row{attrs}>" + "".join(x for _, x in cells) + "</row>"


def patch_cols(sheet):
    """Ширины новых колонок. Существующий <col>, перекрывающий AV…BJ, разрезается."""
    m = re.search(r"<cols>(.*?)</cols>", sheet, re.S)
    widths = {col_idx(c): w for c, _, w in HEADERS}
    new_col = lambda a, b, w=None, attrs="": (
        f'<col min="{a}" max="{b}"' + (f' width="{w}" customWidth="1"' if w else "") + attrs + "/>")
    if not m:
        cols = "".join(new_col(i, i, w) for i, w in widths.items())
        return sheet.replace("<sheetData", f"<cols>{cols}</cols><sheetData", 1)
    out = []
    for cm in re.finditer(r"<col\b[^>]*/>", m.group(1)):
        el = cm.group(0)
        a = int(re.search(r'\bmin="(\d+)"', el).group(1))
        b = int(re.search(r'\bmax="(\d+)"', el).group(1))
        if b < FIRST_COL or a > LAST_COL:
            out.append(el)
            continue
        rest = re.sub(r'\s+(min|max|width|customWidth|bestFit)="[^"]*"', "", el[4:-2])
        if a < FIRST_COL:
            out.append(el.replace(f'max="{b}"', f'max="{FIRST_COL - 1}"'))
        if b > LAST_COL:
            out.append(re.sub(r'\bmin="\d+"', f'min="{LAST_COL + 1}"', el))
    for i, w in widths.items():
        out.append(new_col(i, i, w))
    out.sort(key=lambda e: int(re.search(r'\bmin="(\d+)"', e).group(1)))
    return sheet[:m.start(1)] + "".join(out) + sheet[m.end(1):]


# ------------------------------------------------------------------ строка результата
STATUS_STYLE = {"OK": "ok", "ОШИБКА": "bad", "ПЕРЕПУТАНЫ СТРОКИ": "bad", "ПОЗИЦИЯ НЕ НАЙДЕНА": "bad",
                "ПРОВЕРИТЬ": "warn", "НЕТ ФАКТУРЫ": "none"}


def xl_round(x, nd=2):
    """Округление как ОКРУГЛ в Excel (0,5 — от нуля)."""
    from decimal import Decimal, ROUND_HALF_UP
    return Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP)


def c_check(rn, x, it, st):
    """Формула =И(ОКРУГЛ(H;2)=ОКРУГЛ(BJ;2);ОКРУГЛ(I;2)=ОКРУГЛ(BK;2);ОКРУГЛ(J;2)=ОКРУГЛ(BL;2))
    с уже посчитанным значением, чтобы ИСТИНА/ЛОЖЬ было видно сразу."""
    f = (f"AND(ROUND(H{rn},2)=ROUND(BJ{rn},2),ROUND(I{rn},2)=ROUND(BK{rn},2),"
         f"ROUND(J{rn},2)=ROUND(BL{rn},2))")
    pairs = [(x.get("H"), it.get("qty")), (x.get("I"), it.get("price")), (x.get("J"), it.get("sum"))]
    val = all(xl_round(a or 0) == xl_round(b or 0) for a, b in pairs)
    return f'<c r="BM{rn}" s="{st}" t="b"><f>{f}</f><v>{1 if val else 0}</v></c>'


def row_cells(rn, r, stats, S):
    m, x = r["match"], r["row"]
    cells = {}
    put = lambda col, xml_fn, v, st: cells.__setitem__(col_idx(col), xml_fn(f"{col}{rn}", v, S[st]))

    status = r["status"]
    style = STATUS_STYLE.get(status, "warn")
    if status == "OK" and r["warn"]:
        style = "warn"
    put("AV", c_str, "Нет фактуры" if status == "НЕТ ФАКТУРЫ" else status, style)

    notes = list(m["notes"]) if m else []
    if r.get("note"):
        notes.append(r["note"])
    notes += [f"⚠ {w}" for w in r["warn"]]
    if notes:
        put("AW", c_str, " | ".join(notes), "text")
    if r.get("short"):
        put("AX", c_str, r["short"], "text")

    it = m["item"] if m else None
    if it:
        put("AY", c_str, it["name"], "text")
        put("AZ", c_bool, m["sim"] >= SIM_TRUE, "text")
        put("BA", c_num, m["sim"], "pct")
        put("BH", c_str, f"{r['inv']['file']}, поз. {it['no']}", "text")
        if it.get("unit"):
            put("BI", c_str, it["unit"], "text")
        if it["qty"] is not None:
            put("BJ", c_num, it["qty"], "num")
        if it["price"] is not None:
            put("BK", c_num, it["price"], "num")
        if it["sum"] is not None:
            put("BL", c_num, it["sum"], "num")
        cells[col_idx("BM")] = c_check(rn, x, it, S["text"])

    if stats:
        put("BB", c_num, stats["n"], "int")
        put("BC", c_num, stats["min"], "num")
        put("BD", c_num, stats["max"], "num")
        put("BE", c_num, stats["avg"], "num")
        put("BF", c_num, stats["med"], "num")
        if x.get("AL") and stats["med"]:
            put("BG", c_num, x["AL"] / stats["med"] - 1, "pct")
    return cells


# ------------------------------------------------------------------ главное
def sheet_path(z, sheet_name):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        name = re.search(r'\bname="([^"]*)"', m.group(0)).group(1)
        name = name.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        if name == sheet_name:
            rid = re.search(r'\br:id="([^"]+)"', m.group(0)).group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        if f'Id="{rid}"' in m.group(0):
            target = re.search(r'\bTarget="([^"]+)"', m.group(0)).group(1)
            return target.lstrip("/") if target.startswith("/xl/") else "xl/" + target
    raise ValueError(f"Лист «{sheet_name}» не найден внутри файла")


def write_copy(src, dst, sheet_name, results, hist_stats_fn):
    with zipfile.ZipFile(src) as z:
        sp = sheet_path(z, sheet_name)
        sheet = z.read(sp).decode("utf-8")
        styles, S = add_styles(z.read("xl/styles.xml").decode("utf-8"))

        new = {2: {col_idx(c): c_str(f"{c}2", title, S["head"]) for c, title, _ in HEADERS}}
        for r in results:
            rn = r["row"]["rn"]
            new[rn] = row_cells(rn, r, hist_stats_fn(r["row"]["code"]), S)

        found = set()

        def repl(m):
            rn = int(m.group(1))
            if rn in new:
                found.add(rn)
                return patch_row(m.group(0), new[rn])
            return m.group(0)

        sheet = ROW_RX.sub(repl, sheet)
        lost = set(new) - found
        if lost:
            raise ValueError(f"Не найдены строки в листе: {sorted(lost)[:10]}")
        sheet = patch_cols(sheet)
        sheet = re.sub(r'(<dimension ref="[A-Z]+\d+:)([A-Z]+)(\d+)"',
                       lambda m: f'{m.group(1)}{col_letters(max(col_idx(m.group(2)), LAST_COL))}{m.group(3)}"',
                       sheet, count=1)
        # фильтр расширяется на новые колонки — можно фильтровать по статусу.
        # Excel хранит диапазон фильтра ещё и в скрытом имени _xlnm._FilterDatabase —
        # оба места должны совпадать, иначе Excel пишет «ошибка в части содержимого».
        new_ref = None
        m = re.search(r'<autoFilter ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', sheet)
        if m:
            last = col_letters(max(col_idx(m.group(3)), LAST_COL))
            new_ref = (m.group(1), m.group(2), last, m.group(4))
            sheet = sheet[:m.start()] + f'<autoFilter ref="{m.group(1)}{m.group(2)}:{last}{m.group(4)}"' + sheet[m.end():]

        workbook = z.read("xl/workbook.xml").decode("utf-8")
        if new_ref:
            workbook = fix_filter_name(workbook, sheet_name, new_ref)

        # calcChain.xml — служебный список формул. Новых формул (колонка BM) в нём нет,
        # поэтому убираем его: Excel сам построит его заново при открытии книги.
        drop = {"xl/calcChain.xml"}
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rels = re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", rels)
        ctypes = z.read("[Content_Types].xml").decode("utf-8")
        ctypes = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", ctypes)

        replace = {sp: sheet, "xl/styles.xml": styles, "xl/workbook.xml": workbook,
                   "xl/_rels/workbook.xml.rels": rels, "[Content_Types].xml": ctypes}
        with zipfile.ZipFile(dst, "w") as out:
            for info in z.infolist():
                if info.filename in drop:
                    continue
                data = replace[info.filename].encode("utf-8") if info.filename in replace else z.read(info.filename)
                out.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    return len(results)


def fix_filter_name(workbook, sheet_name, ref):
    """Синхронизирует скрытое имя _xlnm._FilterDatabase листа с новым диапазоном фильтра."""
    names = [re.search(r'\bname="([^"]*)"', m.group(0)).group(1)
             for m in re.finditer(r"<sheet\b[^>]*/>", workbook)]
    names = [n.replace("&amp;", "&").replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">") for n in names]
    if sheet_name not in names:
        return workbook
    idx = names.index(sheet_name)
    c1, r1, c2, r2 = ref
    quoted = "'" + sheet_name.replace("'", "''") + "'" if re.search(r"[^\w.]", sheet_name) else sheet_name
    value = xml_text(f"{quoted}!${c1}${r1}:${c2}${r2}")
    rx = re.compile(rf'(<definedName\b(?=[^>]*\bname="_xlnm\._FilterDatabase")(?=[^>]*\blocalSheetId="{idx}")[^>]*>)'
                    r"[^<]*(</definedName>)")
    return rx.sub(lambda m: m.group(1) + value + m.group(2), workbook, count=1)
