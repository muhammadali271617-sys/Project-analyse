"""
Сборка proverka_faktur.html: вставляет библиотеки pdf.js и ExcelJS прямо в HTML,
чтобы программа работала офлайн (открытие файла двойным щелчком).

Нужны файлы библиотек (один раз):
    npm i pdfjs-dist@3.11.174 exceljs@4.4.0

Запуск:
    python html/build_html.py --libs <папка с node_modules>
"""

import argparse
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent


def make_inline_safe(code: str) -> str:
    # Внутри <script> нельзя встречать "</script" и "<!--" — экранируем их.
    # "<\/script" и "<\!--" в JS-строках/регулярках означают то же самое.
    code = re.sub(r"</(script)", r"<\\/\1", code, flags=re.IGNORECASE)
    code = code.replace("<!--", "<\\!--")
    # Удаляем ссылки на source map — их файлы не встраиваются
    code = re.sub(r"^//# sourceMappingURL=.*$", "", code, flags=re.MULTILINE)
    return code


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--libs", required=True, help="папка, где лежит node_modules")
    p.add_argument("-o", "--output", default=str(HERE.parent / "proverka_faktur.html"))
    args = p.parse_args()

    nm = Path(args.libs) / "node_modules"
    parts = {
        "/*__PDFJS__*/": nm / "pdfjs-dist/build/pdf.min.js",
        "/*__PDFJS_WORKER__*/": nm / "pdfjs-dist/build/pdf.worker.min.js",
        "/*__EXCELJS__*/": nm / "exceljs/dist/exceljs.min.js",
    }
    html = (HERE / "template.html").read_text(encoding="utf-8")
    for marker, path in parts.items():
        assert html.count(marker) == 1, marker
        code = make_inline_safe(path.read_text(encoding="utf-8"))
        html = html.replace(marker, code)

    out = Path(args.output)
    out.write_text(html, encoding="utf-8", newline="\n")
    print(f"Готово: {out} ({out.stat().st_size / 1024 / 1024:.1f} МБ)")


if __name__ == "__main__":
    main()
