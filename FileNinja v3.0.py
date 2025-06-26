import os, shutil, time, re, winreg
from pathlib import Path
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Кэшируем регулярные выражения, чтобы не создавать их при каждом вызове
DATE_RE = re.compile(r"(\d{4}[._-]?\d{2}[._-]?\d{2})")
NUMBER_RE = re.compile(r"(№\s*\d+|\d{3,6})")
LONG_NUM_RE = re.compile(r"\b\d{10,12}\b")
ORG_RE = re.compile(r"(ООО\s*\w+|ИП\s*\w+|\b[A-ZА-Я]{2,}\b)", re.I)

def get_download_folder():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                           r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
            return winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
    except: return str(Path.home() / "Downloads")

FILE_TYPES = {
    "Документы": [".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".xls", ".xlsx", ".csv", ".ods",
                  ".ppt", ".pptx", ".odp", ".xml", ".p7s", ".xml.p7s", ".xps"],
    "Картинки": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp", ".tiff", ".heic"],
    "Видео": [".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".flv", ".mts", ".3gp"],
    "Музыка": [".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a"],
    "Архивы": [".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso"],
    "Исполняемые": [".exe", ".msi", ".bat", ".cmd", ".ps1"],
    "Шрифты": [".ttf", ".otf", ".woff", ".woff2"],
    "ЭЦП/подписи": [".sig", ".p7m", ".cer", ".pem", ".der", ".crt", ".key"],
}
MONTHS = {1:"Январь",2:"Февраль",3:"Март",4:"Апрель",5:"Май",6:"Июнь",
          7:"Июль",8:"Август",9:"Сентябрь",10:"Октябрь",11:"Ноябрь",12:"Декабрь"}
KEYWORD_FOLDERS = {
    "отчет":"Отчёты","итог":"Отчёты","результат":"Отчёты","report":"Отчёты","баланс":"Отчёты","аналитика":"Отчёты",
    "счет":"Счета","invoice":"Счета","оплата":"Счета","платеж":"Счета",
    "накладная":"Накладные","накл":"Накладные","waybill":"Накладные","товар":"Накладные",
    "договор":"Договоры","contract":"Договоры","оферта":"Договоры",
    "акт":"Акты","acceptance":"Акты",
    "фото":"Фото","photo":"Фото","img":"Фото","image":"Фото","скан":"Фото","scan":"Фото","скрин":"Фото","screenshot":"Фото",
    "протокол":"Протоколы","minutes":"Протоколы","совещание":"Протоколы",
    "презентация":"Презентации","presentation":"Презентации","slides":"Презентации",
    "инструкция":"Справки","manual":"Справки","справка":"Справки","регламент":"Справки",
    "проект":"Проекты","task":"Проекты","план":"Планы","roadmap":"Планы",
    "тз":"Техдок","техзадание":"Техдок","смета":"Техдок","спецификация":"Техдок","техдок":"Техдок","чертеж":"Техдок","проектирование":"Техдок",
    "подпись":"Подписи","signature":"Подписи","ecp":"Подписи","сертификат":"Подписи","key":"Подписи",
    "предложение":"Коммерческие","quotation":"Коммерческие","offer":"Коммерческие","кп":"Коммерческие",
    "паспорт":"Личное","анкета":"Личное","cv":"Личное","resume":"Личное","диплом":"Личное",
    "log":"Логи","debug":"Логи","trace":"Логи","ошибка":"Логи","crash":"Логи"
}
JUNK_EXTS = {".tmp", ".part", ".crdownload", ".ds_store"}  # множества — быстрее проверки
JUNK_PREFIXES = ("~$", ".~lock")

SOURCE = get_download_folder()
LOG_DIR = Path(SOURCE) / "Логи"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "ninja_log.txt"

def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{now}] {msg}\n")

def get_ext(filename):
    base, ext = os.path.splitext(filename.lower())
    return ".xml.p7s" if ext == ".p7s" and base.endswith(".xml") else ext

def get_category(ext):
    for cat, exts in FILE_TYPES.items():
        if ext in exts: return cat
    return "Прочее"

def get_folder_by_date(path):
    dt = datetime.fromtimestamp(os.path.getctime(path))
    return f"{dt.year}/{dt.month:02d}.{MONTHS[dt.month]}"

def get_keyword_folder(name):
    low = name.lower()
    for key, folder in KEYWORD_FOLDERS.items():
        if key in low:
            return folder
    return None

def should_ignore(name):
    low = name.lower()
    return low.endswith(tuple(JUNK_EXTS)) or low.startswith(JUNK_PREFIXES)

def extract_requisites(name):
    ext = os.path.splitext(name)[1]
    parts = []
    # Используем заранее скомпилированные регексы
    m = DATE_RE.search(name)
    if m: parts.append(m.group(1).replace("_","").replace(".","").replace("-",""))
    m = NUMBER_RE.search(name)
    if m: parts.append(m.group(1).replace("№","").strip())
    m = LONG_NUM_RE.search(name)
    if m: parts.append(m.group(0))
    m = ORG_RE.search(name)
    if m: parts.append(m.group(1).strip())
    return "_".join(parts) + ext if parts else None

def rename_file(path):
    dir_, name = os.path.dirname(path), os.path.basename(path)
    new = extract_requisites(name)
    if not new: return None
    new_path = os.path.join(dir_, new)
    if new_path != path and not os.path.exists(new_path):
        try:
            os.rename(path, new_path)
            log(f"Переименован: {name} → {new}")
            return new_path
        except Exception as e:
            log(f"Ошибка переименования {name}: {e}")
    return None

def move_file(path):
    if not os.path.isfile(path): return
    new = rename_file(path)
    if new: path = new
    name = os.path.basename(path)
    if should_ignore(name): return
    ext = get_ext(name)
    cat = get_category(ext)
    date = get_folder_by_date(path)
    sub = get_keyword_folder(name)
    # Собираем путь в один проход без лишних операций
    parts = [SOURCE, cat]
    if sub: parts.append(sub)
    parts.append(date)
    target = Path(*parts)
    target.mkdir(parents=True, exist_ok=True)
    dst = target / name
    i = 1
    while dst.exists():
        stem, ext_ = os.path.splitext(name)
        dst = target / f"{stem} ({i}){ext_}"
        i += 1
    try:
        shutil.move(path, dst)
        log(f"Перемещён: {name} → {dst}")
        print(f"[✔] {name} → {cat}/{date}")
    except Exception as e:
        log(f"Ошибка перемещения {name}: {e}")

def sort_existing():
    # Используем генератор с проверкой файла, избегая лишних вызовов
    for entry in os.scandir(SOURCE):
        if entry.is_file():
            move_file(entry.path)

def setup_autostart():
    try:
        exe = shutil.which("pythonw")
        if not exe: return log("❌ Не найден pythonw.exe")
        bat = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup" / "start_fileninja.bat"
        with open(bat, "w", encoding="utf-8") as f:
            f.write(f'@echo off\ncd /d "{os.path.dirname(__file__)}"\nstart "" "{exe}" "{__file__}"\n')
        log(f"✅ Автозапуск настроен: {bat}")
    except Exception as e:
        log(f"❌ Ошибка автозапуска: {e}")

class Handler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory: return
        time.sleep(1.5)  # Ждём, пока файл запишется полностью
        move_file(event.src_path)

def main():
    setup_autostart()
    log("Запуск FileNinja")
    print(f"🥷 FileNinja следит за: {SOURCE}")
    sort_existing()
    observer = Observer()
    observer.schedule(Handler(), SOURCE, recursive=False)
    observer.start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log("FileNinja остановлен вручную")
    observer.join()

if __name__ == "__main__":
    main()
