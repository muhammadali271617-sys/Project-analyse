"""
Проверка счёт-фактур (PDF из Didox.uz).

Программа открывает каждый PDF в папке, достаёт из него номер счёт-фактуры
и сравнивает его с именем файла.

Правила сравнения:
  * большие и маленькие буквы не различаются (S2991 = s2991);
  * "/" и "." взаимозаменяемы (543/2 = 543.2);
  * лишние пробелы по краям игнорируются.

Запуск:
    python check_invoices.py "D:\\all factures for price analyses"

Результат печатается в консоль и сохраняется в Excel-отчёт
(по умолчанию "отчет_проверки.xlsx" в той же папке, что и PDF).
"""

import argparse
import logging
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# pdfplumber иногда шумит предупреждениями о шрифтах — они не мешают
logging.getLogger("pdfminer").setLevel(logging.ERROR)

DATE = r"\d{2}\.\d{2}\.\d{4}"

# Шапка Didox: "Ҳисоб-фактура актсиз: 52 02.10.2026 дан"
HEADER_PATTERNS = [
    re.compile(rf"Ҳисоб-фактура[^:\n]*:\s*(\S+)\s+{DATE}\s+дан"),
    re.compile(rf"Hisob-faktura[^:\n]*:\s*(\S+)\s+{DATE}\s+dan", re.IGNORECASE),
    re.compile(rf"Сч[её]т-фактура[^:\n]*:\s*(\S+)\s+(?:от\s+)?{DATE}", re.IGNORECASE),
]

# Заголовок документа: "02.10.2026 даги 52-сонли ... Ҳисобварақ-фактура"
BODY_PATTERNS = [
    re.compile(rf"{DATE}\s+даги\s+(\S+)-сонли\s*\n?\s*Ҳисобварақ-фактура"),
    re.compile(rf"{DATE}\s+dagi\s+(\S+)-sonli\s*\n?\s*Hisobvaraq-faktura", re.IGNORECASE),
    re.compile(rf"Сч[её]т-фактура\s*№\s*(\S+)\s+от\s+{DATE}", re.IGNORECASE),
]

STATUS_OK = "OK"
STATUS_MISMATCH = "НЕ СОВПАДАЕТ"
STATUS_NOT_FOUND = "НОМЕР НЕ НАЙДЕН"
STATUS_ERROR = "ОШИБКА ЧТЕНИЯ"


def normalize(number: str) -> str:
    """Приводит номер к единому виду для сравнения.
    Разделители не важны: «472/п» = «472п» = «472.п», «543/2» = «543.2»."""
    return re.sub(r"[\s/._\-]+", "", number.strip().upper())


def find_first(patterns, text):
    for pattern in patterns:
        match = pattern.search(text)
        if match:
            return match.group(1)
    return None


def read_invoice_number(pdf_path: Path):
    """Возвращает (номер из шапки, номер из заголовка документа)."""
    with pdfplumber.open(pdf_path) as pdf:
        # номер всегда на первой странице
        text = pdf.pages[0].extract_text() or ""
    return find_first(HEADER_PATTERNS, text), find_first(BODY_PATTERNS, text)


def file_number_from_name(stem: str) -> str:
    """Номер из имени файла. Текст в скобках игнорируется:
    "543 (Next Glass)" → "543", "52 (1)" → "52"."""
    stem = re.sub(r"\s+", " ", re.sub(r"\([^)]*\)", "", stem)).strip()
    # имя, которое даёт Didox при скачивании: «Счет_фактура_без_акта_6790_от_12_10_2026_»
    m = re.match(r"(?i)^сч[её]т[_ -]фактура(?:[_ ]без[_ ]акта)?[_ ](.+?)[_ ]от[_ ]\d{2}[_.]\d{2}[_.]\d{4}_?$", stem)
    return m.group(1) if m else stem


def check_file(pdf_path: Path) -> dict:
    file_number = file_number_from_name(pdf_path.stem)
    row = {
        "file": pdf_path.name,
        "file_number": file_number,
        "pdf_number": "",
        "status": "",
        "comment": "",
    }
    try:
        header_num, body_num = read_invoice_number(pdf_path)
    except Exception as exc:  # битый / нечитаемый PDF
        row["status"] = STATUS_ERROR
        row["comment"] = str(exc)
        return row

    pdf_number = header_num or body_num
    if not pdf_number:
        row["status"] = STATUS_NOT_FOUND
        row["comment"] = "Не удалось найти номер счёт-фактуры в тексте PDF"
        return row

    row["pdf_number"] = pdf_number
    comments = []
    if header_num and body_num and normalize(header_num) != normalize(body_num):
        comments.append(f"В шапке номер {header_num}, в документе {body_num}")

    # Совпадение — только если ВСЕ найденные номера равны имени файла
    # (если в шапке и заголовке номера разные — это тоже ошибка)
    others = [n for n in (header_num, body_num) if n and normalize(n) != normalize(file_number)]
    if not others and not comments:
        row["status"] = STATUS_OK
    else:
        row["status"] = STATUS_MISMATCH
        if others:
            comments.insert(0, f"Файл назван {file_number}, а в фактуре номер {others[0]}")
    row["comment"] = "; ".join(comments)
    return row


def sort_key(path: Path):
    """Сортировка как в проводнике: 52, 442, 543, 6668 ... затем буквенные."""
    digits = re.match(r"\d+", path.stem)
    return (0, int(digits.group()), path.stem) if digits else (1, 0, path.stem.upper())


def save_report(rows, report_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Проверка"
    headers = ["Файл", "Номер по имени файла", "Номер в фактуре", "Статус", "Комментарий"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    yellow = PatternFill("solid", fgColor="FFEB9C")

    for r in rows:
        ws.append([r["file"], r["file_number"], r["pdf_number"], r["status"], r["comment"]])
        fill = green if r["status"] == STATUS_OK else red if r["status"] == STATUS_MISMATCH else yellow
        ws.cell(row=ws.max_row, column=4).fill = fill

    for i, width in enumerate([18, 22, 18, 20, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    wb.save(report_path)


def main():
    parser = argparse.ArgumentParser(description="Проверка: номер счёт-фактуры в PDF = имя файла")
    parser.add_argument("folder", nargs="?", default=".", help="папка с PDF (по умолчанию текущая)")
    parser.add_argument("-o", "--output", help="путь к Excel-отчёту")
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"Папка не найдена: {folder}")

    pdfs = sorted((p for p in folder.iterdir() if p.suffix.lower() == ".pdf"), key=sort_key)
    if not pdfs:
        sys.exit(f"В папке нет PDF-файлов: {folder}")

    rows = []
    for pdf_path in pdfs:
        row = check_file(pdf_path)
        rows.append(row)
        mark = "✔" if row["status"] == STATUS_OK else "✘"
        print(f"{mark} {row['file']:<20} в фактуре: {row['pdf_number'] or '-':<12} {row['status']}"
              + (f"  ({row['comment']})" if row["comment"] else ""))

    ok = sum(r["status"] == STATUS_OK for r in rows)
    print(f"\nВсего файлов: {len(rows)}   совпало: {ok}   проблем: {len(rows) - ok}")

    report_path = Path(args.output) if args.output else folder / "отчет_проверки.xlsx"
    try:
        save_report(rows, report_path)
        print(f"Отчёт сохранён: {report_path}")
    except PermissionError:
        print(f"Не удалось сохранить отчёт — закройте файл {report_path} в Excel и запустите снова")


if __name__ == "__main__":
    main()
