"""
Сверка с фактурами — запуск через страницу в браузере.

Двойной щелчок по start.pyw (без чёрного окна) запускает маленький локальный
веб-сервер на 127.0.0.1 и открывает страницу в Microsoft Edge.
Используется только стандартная библиотека Python; сама сверка — sverka.py.

Сервер доступен только с этого компьютера. Чтобы его остановить, нажмите
«Завершить» на странице (или он сам закроется через 15 минут без открытой страницы).
"""

import json
import os
import secrets
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
PAGE = HERE / "start_page.html"
LOG = HERE / "start.log"
PREFERRED_PORT = 8770
IDLE_EXIT_SEC = 15 * 60
APP_ID = "sverka-faktur"

# pythonw.exe запускается без консоли: sys.stdout/stderr = None, любой print упал бы
if sys.stdout is None or sys.stderr is None:
    _log = open(LOG, "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stdout or _log
    sys.stderr = sys.stderr or _log
sys.path.insert(0, str(HERE))


def log(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
    except OSError:
        pass


# ---------------------------------------------------------------- библиотеки
def check_libs():
    missing = []
    for mod in ("openpyxl", "pdfplumber"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    if missing:
        return ("Не установлены библиотеки: " + ", ".join(missing) + ". Откройте командную строку и "
                "выполните:  py -m pip install pdfplumber openpyxl  — затем запустите программу снова.")
    for mod in ("sverka", "excel_writer", "po_buhgalteram"):
        if not (HERE / f"{mod}.py").exists():
            return f"Рядом с start.pyw нет файла {mod}.py. Положите все файлы программы в одну папку."
    return None


LIBS_ERROR = check_libs()


class _Stub:  # чтобы обработчики ошибок работали, даже если sverka.py не загрузился
    class CheckError(Exception):
        pass

    class NeedSheet(CheckError):
        sheets = []


sverka = _Stub
_loaded_mtimes = {}


def _mtimes():
    return {n: (HERE / f"{n}.py").stat().st_mtime for n in ("sverka", "excel_writer", "po_buhgalteram")
            if (HERE / f"{n}.py").exists()}


def reload_if_changed():
    """Если sverka.py / excel_writer.py заменили, пока страница открыта, — подгружаем новые.
    Иначе продолжала бы работать старая версия из памяти."""
    global sverka
    if LIBS_ERROR or sverka is _Stub:
        return
    now = _mtimes()
    if now != _loaded_mtimes:
        import importlib
        for mod in ("excel_writer", "po_buhgalteram"):
            if mod in sys.modules:
                importlib.reload(sys.modules[mod])
        sverka = importlib.reload(sverka)
        _loaded_mtimes.clear()
        _loaded_mtimes.update(now)
        ST.data = ST.data_key = None   # данные Excel перечитаем уже новой версией
        log(f"Обновлены файлы программы, версия {getattr(sverka, 'VERSION', '?')}")


if not LIBS_ERROR:
    try:
        import sverka  # noqa: E402
        _loaded_mtimes.update(_mtimes())
    except Exception as e:  # синтаксическая ошибка и т.п.
        LIBS_ERROR = f"Не удалось загрузить sverka.py: {e}"
        log(traceback.format_exc())


# ---------------------------------------------------------------- состояние
class State:
    def __init__(self):
        self.lock = threading.Lock()
        self.folder = str(HERE)
        self.data = None          # результат load_excel (кэш для выбранного Excel)
        self.data_key = None
        self.last = None          # результат run_check
        self.acc_src = None       # проверенная копия, из которой делаем книгу по бухгалтерам
        self.acc_sheet = None
        self.acc_path = None      # созданная книга «Ошибки по бухгалтерам»
        self.job = {"state": "idle"}
        self.last_seen = time.time()


ST = State()


def set_job(**kw):
    with ST.lock:
        ST.job.update(kw)


def start_job(kind, fn):
    with ST.lock:
        if ST.job.get("state") == "running":
            return False
        ST.job = {"state": "running", "kind": kind, "text": "Начинаю…", "done": None, "total": None}

    def progress(text, done=None, total=None):
        set_job(text=text, done=done, total=total)

    def worker():
        try:
            result = fn(progress)
            set_job(state="done", result=result)
        except sverka.NeedSheet as e:
            set_job(state="done", result={"need_sheet": e.sheets})
        except sverka.CheckError as e:
            set_job(state="error", error=str(e))
        except PermissionError as e:
            set_job(state="error", error=f"Нет доступа к файлу: {e.filename}. Возможно, он открыт в Excel — "
                                         f"закройте его и попробуйте снова.")
        except SystemExit as e:  # sys.exit() внутри вспомогательных модулей — это понятное сообщение
            set_job(state="error", error=str(e) or "Операция прервана.")
        except BaseException as e:  # что бы ни случилось — задание не должно «зависнуть»
            log(traceback.format_exc())
            set_job(state="error", error=f"Непредвиденная ошибка: {e!r}. Подробности в файле {LOG.name}.")

    threading.Thread(target=worker, daemon=True).start()
    return True


# ---------------------------------------------------------------- действия
def do_scan(folder, excel=None, sheet=None):
    reload_if_changed()
    folder = Path(folder)
    if not folder.is_dir():
        raise sverka.CheckError(f"Папка не найдена: {folder}")
    excels = sverka.find_excels(folder)
    pdfs = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    if not excels:
        raise sverka.CheckError("В этой папке нет Excel-файла (.xlsx). Положите в одну папку Excel "
                                "с листом проверки и PDF-фактуры.")
    if excel:
        xlsx = folder / excel
        if not xlsx.exists():
            raise sverka.CheckError(f"Файл не найден: {excel}")
    elif len(excels) > 1:
        return lambda progress: {"need_excel": [p.name for p in excels]}
    else:
        xlsx = excels[0]

    def job(progress):
        key = (str(xlsx), xlsx.stat().st_mtime, sheet)
        if ST.data_key != key:
            ST.data = sverka.load_excel(xlsx, sheet, progress, ask=None)
            ST.data_key = key
        d = ST.data
        months = [{"key": k, "name": sverka.month_name(k), "count": d["months"][k]} for k in sorted(d["months"])]
        return {"version": getattr(sverka, "VERSION", None),
                "excel": xlsx.name, "excels": [p.name for p in excels], "sheet": d["ws_name"],
                "sheets": d["sheets"], "months": months, "pdfs": len(pdfs),
                "open_warning": (f"Файл «{xlsx.name}» сейчас открыт в Excel. Проверка его только читает, "
                                 f"это не мешает." if sverka.is_open_in_excel(xlsx) else None)}
    return job


def row_json(r):
    x, m, inv = r["row"], r["match"], r["inv"]
    notes = list(m["notes"]) if m else []
    if r.get("note"):
        notes.append(r["note"])
    notes += ["⚠ " + w for w in r["warn"]]
    it = m["item"] if m else None
    return {"status": r["status"], "rn": x["rn"], "code": str(x["code"]), "name": str(x["F"] or ""),
            "supplier": str(x["O"] or ""), "num": x["R"], "short": r.get("short", ""),
            "notes": notes, "warn": bool(r["warn"]),
            "invoice": (f"{inv['file']}, поз. {it['no']}" if inv and it else (inv["file"] if inv else "")),
            "excel": [x["H"], x["I"], x["J"], str(x["G"] or "")],
            "fact": [it["qty"], it["price"], it["sum"], it["unit"]] if it else None}


def do_run(period):
    if _mtimes() != _loaded_mtimes:
        raise sverka.CheckError("Файлы программы обновились — нажмите «↻ Прочитать», чтобы загрузить новую версию, "
                                "и затем «Проверить».")
    data = ST.data
    if not data:
        raise sverka.CheckError("Сначала выберите папку с Excel.")
    period = set(period or [])
    if not period:
        raise sverka.CheckError("Выберите хотя бы один месяц.")
    folder = data["xlsx"].parent
    dst = data["xlsx"].with_name(f"{data['xlsx'].stem}_проверено{data['xlsx'].suffix}")

    def job(progress):
        if dst.exists() and sverka.is_open_in_excel(dst):
            raise sverka.CheckError(f"Файл «{dst.name}» открыт в Excel. Закройте его — программа "
                                    f"запишет в него новые результаты.")
        res = sverka.run_check(folder, data, period, progress)
        ST.last = res
        rows = [row_json(r) for r in res["results"] if r["status"] != "OK" or r["warn"]]
        rows.sort(key=lambda r: (sverka.SEVERITY.get(r["status"], 9), r["rn"]))
        return {"counts": dict(res["counts"]), "total": len(res["results"]), "rows": rows,
                "report": str(res["report"]), "excel": str(res["excel"]) if res["excel"] else None,
                "excel_error": res["excel_error"], "missing": len(res["missing"]),
                "bad_pdfs": [{"file": b["file"], "error": b["error"]} for b in res["bad_pdfs"]],
                "invoices": len(res["invoices"]),
                "period": [sverka.month_name(p) for p in sorted(period)]}
    return job


def checked_copy():
    """Проверенная копия Excel: после проверки — её путь, иначе ищем «…_проверено.xlsx» в папке."""
    if ST.last and ST.last.get("excel") and Path(ST.last["excel"]).exists():
        return Path(ST.last["excel"])
    folder = Path(ST.folder)
    found = sorted((p for p in folder.glob("*_проверено.xlsx") if not p.name.startswith("~$")),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    if not found:
        raise sverka.CheckError("Нет проверенной копии «…_проверено.xlsx» — сначала нажмите «Проверить».")
    return found[0]


def do_acc_scan():
    src = checked_copy()

    def job(progress):
        import po_buhgalteram as pb
        log(f"Шаг 4: читаю {src}")
        progress(f"Читаю «{src.name}»…")
        with pb.zipfile.ZipFile(src) as z:
            names = [n for n, _, _ in pb.sheet_parts(z)]
        try:
            sheet = pb.find_check_sheet(names, ST.data["ws_name"] if ST.data else None)
        except SystemExit as e:
            raise sverka.CheckError(str(e))
        avail = pb.available_months(pb.read_candidates(src, sheet))
        ST.acc_src, ST.acc_sheet = src, sheet
        return {"file": src.name, "sheet": sheet,
                "months": [{"key": k, "name": pb.month_name(k), "count": n} for k, n in avail.items()]}
    return job


def do_acc_build(months):
    months = set(months or [])
    if not months:
        raise sverka.CheckError("Выберите хотя бы один месяц.")
    if not ST.acc_src:
        raise sverka.CheckError("Сначала откройте шаг «Ошибки по бухгалтерам» ещё раз.")
    src, sheet = ST.acc_src, ST.acc_sheet

    def job(progress):
        import po_buhgalteram as pb
        dst = pb.output_name(src, months)
        if dst.exists() and sverka.is_open_in_excel(dst):
            raise sverka.CheckError(f"Файл «{dst.name}» открыт в Excel. Закройте его и попробуйте снова.")
        progress("Собираю листы по бухгалтерам…")
        try:
            dst, counts, problems = pb.make_book(src, sheet, months, dst)
        except ValueError as e:
            raise sverka.CheckError(str(e))
        ST.acc_path = dst
        return {"file": dst.name, "counts": counts, "total": sum(n for _, n in counts), "problems": problems[:20]}
    return job


def open_file(path):
    path = str(path)
    if os.name == "nt":
        os.startfile(path)  # открывает в программе по умолчанию (Excel / браузер)
    elif os.environ.get("SVERKA_NO_BROWSER"):
        log(f"open {path}")
    else:
        subprocess.Popen(["xdg-open", path])


def pick_folder(initial):
    """Окно выбора папки Windows (tkinter из стандартной библиотеки)."""
    import tkinter
    from tkinter import filedialog
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return filedialog.askdirectory(initialdir=initial, title="Папка с Excel и фактурами",
                                       mustexist=True, parent=root)
    finally:
        root.destroy()


# ---------------------------------------------------------------- HTTP
TOKEN = secrets.token_urlsafe(16)


class Handler(BaseHTTPRequestHandler):
    server_version = "Sverka/1.0"

    def log_message(self, fmt, *args):  # без шума в консоли/логе
        pass

    def send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}") if n else {}

    def authorized(self):
        if self.headers.get("X-Token") != TOKEN:
            self.send_json({"error": "Нет доступа"}, 403)
            return False
        return True

    def do_GET(self):
        if self.path == "/api/ping":
            return self.send_json({"app": APP_ID})
        if self.path in ("/", "/index.html"):
            html = PAGE.read_text(encoding="utf-8")
            html = html.replace("{{TOKEN}}", TOKEN)
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        if not self.authorized():
            return
        ST.last_seen = time.time()
        if self.path == "/api/state":
            return self.send_json({"folder": ST.folder, "libs_error": LIBS_ERROR, "job": ST.job,
                                   "version": getattr(sverka, "VERSION", None)})
        if self.path == "/api/job":
            return self.send_json(ST.job)
        self.send_json({"error": "Не найдено"}, 404)

    def do_POST(self):
        if not self.authorized():
            return
        ST.last_seen = time.time()
        try:
            data = self.body()
        except ValueError:
            return self.send_json({"error": "Неверный запрос"}, 400)
        try:
            if LIBS_ERROR and self.path not in ("/api/quit", "/api/heartbeat"):
                return self.send_json({"error": LIBS_ERROR}, 400)
            if self.path == "/api/heartbeat":
                return self.send_json({"ok": True})
            if self.path == "/api/pick-folder":
                chosen = pick_folder(ST.folder)
                if chosen:
                    ST.folder = str(Path(chosen))
                return self.send_json({"folder": ST.folder, "cancelled": not chosen})
            if self.path == "/api/scan":
                folder = data.get("folder") or ST.folder
                ST.folder = str(Path(folder))
                fn = do_scan(folder, data.get("excel"), data.get("sheet"))
                ok = start_job("scan", fn)
                return self.send_json({"started": ok} if ok else {"error": "Уже выполняется другая операция."})
            if self.path == "/api/run":
                ok = start_job("run", do_run(data.get("period")))
                return self.send_json({"started": ok} if ok else {"error": "Уже выполняется другая операция."})
            if self.path == "/api/acc-scan":
                ok = start_job("acc-scan", do_acc_scan())
                return self.send_json({"started": ok} if ok else {"error": "Уже выполняется другая операция."})
            if self.path == "/api/acc-build":
                ok = start_job("acc-build", do_acc_build(data.get("months")))
                return self.send_json({"started": ok} if ok else {"error": "Уже выполняется другая операция."})
            if self.path == "/api/open":
                what = data.get("what")
                last = ST.last
                if what == "acc":
                    path = ST.acc_path
                    if not path or not Path(path).exists():
                        return self.send_json({"error": "Книга по бухгалтерам ещё не создана."}, 400)
                    open_file(path)
                    return self.send_json({"ok": True})
                if not last:
                    return self.send_json({"error": "Сначала выполните проверку."}, 400)
                path = last["excel"] if what == "excel" else last["report"]
                if not path or not Path(path).exists():
                    return self.send_json({"error": "Файл не найден — выполните проверку ещё раз."}, 400)
                open_file(path)
                return self.send_json({"ok": True})
            if self.path == "/api/open-folder":
                open_file(ST.folder)
                return self.send_json({"ok": True})
            if self.path == "/api/quit":
                self.send_json({"ok": True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return
            self.send_json({"error": "Не найдено"}, 404)
        except sverka.CheckError as e:
            self.send_json({"error": str(e)}, 400)
        except Exception as e:
            log(traceback.format_exc())
            self.send_json({"error": f"Непредвиденная ошибка: {e}"}, 500)


# ---------------------------------------------------------------- запуск
def already_running(port):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=1) as r:
            return json.load(r).get("app") == APP_ID
    except Exception:
        return False


def open_browser(url):
    if os.environ.get("SVERKA_NO_BROWSER"):
        return
    if os.name == "nt":
        try:
            os.startfile("microsoft-edge:" + url)
            return
        except OSError:
            pass
    webbrowser.open(url)


def idle_watch(server):
    while True:
        time.sleep(30)
        if ST.job.get("state") != "running" and time.time() - ST.last_seen > IDLE_EXIT_SEC:
            log("Нет открытой страницы — сервер остановлен.")
            server.shutdown()
            return


def main():
    port = int(os.environ.get("SVERKA_PORT", PREFERRED_PORT))
    try:
        server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError:
        if already_running(port):
            # программа уже запущена — просто открываем её страницу
            # (адрес со старым ключом откроется в той же вкладке после обновления)
            open_browser(f"http://127.0.0.1:{port}/")
            return
        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    log(f"Сервер запущен: {url}")
    threading.Thread(target=idle_watch, args=(server,), daemon=True).start()
    open_browser(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()
        log("Сервер остановлен.")


if __name__ == "__main__":
    main()
