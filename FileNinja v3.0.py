import os
import shutil
import time
import re
import winreg
from pathlib import Path
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

def get_download_folder():
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        ) as key:
            return winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
    except Exception:
        return str(Path.home() / "Downloads")

FILE_TYPES = {
    "Документы": {".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".xls", ".xlsx", ".csv", ".ods",
                  ".ppt", ".pptx", ".odp", ".xml", ".p7s", ".xml.p7s", ".xps"},
    "Картинки": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp", ".tiff", ".heic"},
    "Видео": {".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".flv", ".mts", ".3gp"},
    "Музыка": {".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a"},
    "Архивы": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso"},
    "Исполняемые": {".exe", ".msi", ".bat", ".cmd", ".ps1"},
    "Шрифты": {".ttf", ".otf", ".woff", ".woff2"},
    "ЭЦП/подписи": {".sig", ".p7m", ".cer", ".pem", ".der", ".crt", ".key"},
}

MONTHS = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь",
    7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

KEYWORD_FOLDERS = {
    "отчет": "Отчёты", "итог": "Отчёты", "результат": "Отчёты", "report": "Отчёты", "баланс": "Отчёты", "аналитика": "Отчёты",
    "счет": "Счета", "invoice": "Счета", "оплата": "Счета", "платеж": "Счета",
    "накладная": "Накладные", "накл": "Накладные", "waybill": "Накладные", "товар": "Накладные",
    "договор": "Договоры", "contract": "Договоры", "оферта": "Договоры",
    "акт": "Акты", "acceptance": "Акты",
    "фото": "Фото", "photo": "Фото", "img": "Фото", "image": "Фото", "скан": "Фото", "scan": "Фото", "скрин": "Фото", "screenshot": "Фото",
    "протокол": "Протоколы", "minutes": "Протоколы", "совещание": "Протоколы",
    "презентация": "Презентации", "presentation": "Презентации", "slides": "Презентации",
    "инструкция": "Справки", "manual": "Справки", "справка": "Справки", "регламент": "Справки",
    "проект": "Проекты", "task": "Проекты", "план": "Планы", "roadmap": "Планы",
    "тз": "Техдок", "техзадание": "Техдок", "смета": "Техдок", "спецификация": "Техдок", "техдок": "Техдок", "чертеж": "Техдок", "проектирование": "Техдок",
    "подпись": "Подписи", "signature": "Подписи", "ecp": "Подписи", "сертификат": "Подписи", "key": "Подписи",
    "предложение": "Коммерческие", "quotation": "Коммерческие", "offer": "Коммерческие", "кп": "Коммерческие",
    "паспорт": "Личное", "анкета": "Личное", "cv": "Личное", "resume": "Личное", "диплом": "Личное",
    "log": "Логи", "debug": "Логи", "trace": "Логи", "ошибка": "Логи", "crash": "Логи"
}

JUNK_EXTS = {".tmp", ".part", ".crdownload", ".ds_store"}
JUNK_PREFIXES = ("~$", ".~lock")

SOURCE = Path(get_download_folder())
LOG_DIR = SOURCE / "Логи"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "ninja_log.txt"

def log(msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")

def get_ext(filename: str) -> str:
    base, ext = os.path.splitext(filename.lower())
    return ".xml.p7s" if ext == ".p7s" and base.endswith(".xml") else ext

def get_category(ext: str) -> str:
    for cat, exts in FILE_TYPES.items():
        if ext in exts:
            return cat
    return "Прочее"

def get_folder_by_date(path: Path) -> str:
    dt = datetime.fromtimestamp(path.stat().st_ctime)
    return f"{dt.year}/{dt.month:02d}.{MONTHS[dt.month]}"

def get_keyword_folder(name: str) -> str | None:
    name_lower = name.lower()
    for key, folder in KEYWORD_FOLDERS.items():
        if key in name_lower:
            return folder
    return None

def should_ignore(name: str) -> bool:
    name_lower = name.lower()
    return name_lower.endswith(tuple(JUNK_EXTS)) or name_lower.startswith(JUNK_PREFIXES)

def move_file(path: Path):
    if not path.is_file():
        return
    name = path.name
    if should_ignore(name):
        return

    ext = get_ext(name)
    category = get_category(ext)
    date_folder = get_folder_by_date(path)
    keyword_folder = get_keyword_folder(name)

    target_parts = [SOURCE, category]
    if keyword_folder:
        target_parts.append(keyword_folder)
    target_parts.append(date_folder)
    target_dir = Path(*target_parts)
    target_dir.mkdir(parents=True, exist_ok=True)

    dst = target_dir / name
    i = 1
    while dst.exists():
        stem, ext_ = os.path.splitext(name)
        dst = target_dir / f"{stem} ({i}){ext_}"
        i += 1

    try:
        shutil.move(str(path), str(dst))
        log(f"Перемещён: {name} → {dst}")
        print(f"[✔] {name} → {category}/{date_folder}")
    except Exception as e:
        log(f"Ошибка перемещения {name}: {e}")

def sort_existing():
    for entry in SOURCE.iterdir():
        if entry.is_file():
            move_file(entry)

def setup_autostart():
    try:
        pythonw = shutil.which("pythonw")
        if not pythonw:
            log("❌ Не найден pythonw.exe, автозапуск не настроен")
            return
        startup_path = Path(os.getenv("APPDATA")) / \
            "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        startup_path.mkdir(parents=True, exist_ok=True)
        bat_path = startup_path / "start_fileninja.bat"
        script_path = Path(__file__).resolve()
        with bat_path.open("w", encoding="utf-8") as f:
            f.write(f'@echo off\ncd /d "{script_path.parent}"\nstart "" "{pythonw}" "{script_path}"\n')
        log(f"✅ Автозапуск настроен: {bat_path}")
    except Exception as e:
        log(f"❌ Ошибка автозапуска: {e}")

class Handler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        time.sleep(1.5)
        move_file(Path(event.src_path))

def main():
    setup_autostart()
    log("Запуск FileNinja")
    print(f"🥷 FileNinja следит за: {SOURCE}")
    sort_existing()
    observer = Observer()
    observer.schedule(Handler(), str(SOURCE), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log("FileNinja остановлен вручную")
    observer.join()

if __name__ == "__main__":
    main()
