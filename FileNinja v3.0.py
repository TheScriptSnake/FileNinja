import os
import shutil
from pathlib import Path
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import winreg
import re

# Получение пути к папке "Загрузки"
def get_download_folder():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders") as key:
            downloads = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
            return downloads
    except Exception:
        return str(Path.home() / "Downloads")

source_folder = get_download_folder()
log_dir = Path(source_folder) / "Логи"
log_dir.mkdir(parents=True , exist_ok=True)
log_file = log_dir / "ninja_log.txt"

file_types = {
    "Документы": [".pdf", ".docx", ".doc", ".txt", ".xls", ".xlsx", ".pptx", ".xml", ".p7s", ".xml.p7s"],
    "Картинки": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg"],
    "Видео": [".mp4", ".mov", ".avi", ".mkv"],
    "Музыка": [".mp3", ".wav", ".ogg", ".flac"],
    "Архивы": [".zip", ".rar", ".7z", ".tar", ".gz"],
}

month_names = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

keyword_folders = {
    "отчет": "Отчёты", "отчёт": "Отчёты", "report": "Отчёты", "summary": "Отчёты",
    "результат": "Отчёты", "photo": "Фото", "фото": "Фото", "img": "Фото",
    "скрин": "Фото", "накладная": "Накладные", "invoice": "Счета", "счет": "Счета",
    "log": "Логи", "debug": "Логи", "trace": "Логи", "договор": "Договоры",
    "contract": "Договоры", "план": "Планы", "project": "Проекты",
    "task": "Проекты", "презентация": "Презентации", "minutes": "Протоколы"
}

ignore_prefixes = ["~$", ".~lock"]
ignore_extensions = [".tmp", ".part", ".crdownload", ".ds_store"]

def log(message):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{now}] {message}\n")

def should_ignore(filename):
    lower = filename.lower()
    return any(lower.startswith(p) for p in ignore_prefixes) or any(lower.endswith(e) for e in ignore_extensions)

def get_extension(filename):
    base, ext1 = os.path.splitext(filename)
    ext1 = ext1.lower()
    if ext1 == ".p7s":
        base2, ext2 = os.path.splitext(base)
        ext2 = ext2.lower()
        if ext2 == ".xml":
            return ".xml.p7s"
    return ext1

def get_category(ext):
    ext = ext.lower()
    for category, ext_list in file_types.items():
        if ext in ext_list:
            return category
    return "Прочее"

def get_date_folder(path):
    dt = datetime.fromtimestamp(os.path.getctime(path))
    return f"{dt.year}/{dt.month:02d}.{month_names[dt.month]}"

def get_keyword_folder(filename):
    lower = filename.lower()
    for keyword, folder in keyword_folders.items():
        if keyword in lower:
            return folder
    return None

def extract_requisites(filename, file_path=None):
    name, ext = os.path.splitext(filename)
    number_match = re.search(r"(№\s*\d+|\d{3,6})", name)
    date_match = re.search(r"(\d{4}[._-]?\d{2}[._-]?\d{2})", name)
    inn_match = re.search(r"\b\d{10,12}\b", name)
    company_match = re.search(r"(ООО\s*\w+|ИП\s*\w+|\b[A-ZА-Я]{2,}\b)", name, re.IGNORECASE)

    parts = []
    if date_match:
        date_str = date_match.group(1).replace("_","").replace(".","").replace("-","")
        parts.append(date_str)
    if number_match:
        parts.append(number_match.group(1).replace("№", "").strip())
    if inn_match:
        parts.append(inn_match.group(0))
    if company_match:
        parts.append(company_match.group(1).strip())

    if parts:
        new_name = "_".join(parts) + ext
        return new_name
    return None

def rename_file_with_requisites(file_path):
    dirname = os.path.dirname(file_path)
    filename = os.path.basename(file_path)

    new_name = extract_requisites(filename, file_path)
    if not new_name:
        return None

    new_path = os.path.join(dirname, new_name)
    if new_path != file_path and not os.path.exists(new_path):
        try:
            os.rename(file_path, new_path)
            log(f"Переименован по реквизитам: {filename} → {new_name}")
            return new_path
        except Exception as e:
            log(f"Ошибка переименования {filename}: {e}")
    return None

def delete_junk_files():
    junk_exts = [".tmp", ".crdownload", ".part", ".ds_store"]
    junk_prefixes = ["~$", ".~lock"]

    deleted_count = 0
    for filename in os.listdir(source_folder):
        lower = filename.lower()
        if any(lower.endswith(ext) for ext in junk_exts) or any(lower.startswith(pref) for pref in junk_prefixes):
            path = os.path.join(source_folder, filename)
            try:
                if os.path.isfile(path):
                    os.remove(path)
                    log(f"Удалён мусор: {filename}")
                    deleted_count += 1
            except Exception as e:
                log(f"Ошибка удаления мусора {filename}: {e}")
    print(f"Удалено мусорных файлов: {deleted_count}")

def archive_old_files(days=30):
    archive_root = os.path.join(source_folder, "Архив")
    now = time.time()
    cutoff = now - days * 86400

    for filename in os.listdir(source_folder):
        path = os.path.join(source_folder, filename)
        if os.path.isfile(path):
            ctime = os.path.getctime(path)
            if ctime < cutoff:
                dt = datetime.fromtimestamp(ctime)
                folder = os.path.join(archive_root, f"{dt.year}", f"{dt.month:02d}.{month_names[dt.month]}")
                os.makedirs(folder, exist_ok=True)
                target_path = os.path.join(folder, filename)
                count = 1
                while os.path.exists(target_path):
                    name, ext = os.path.splitext(filename)
                    target_path = os.path.join(folder, f"{name} ({count}){ext}")
                    count += 1
                try:
                    shutil.move(path, target_path)
                    log(f"Заархивирован: {filename} → {folder}")
                except Exception as e:
                    log(f"Ошибка архивации {filename}: {e}")

def move_file(file_path):
    if not os.path.isfile(file_path):
        return

    filename = os.path.basename(file_path)

    if should_ignore(filename):
        log(f"Игнорирован: {filename}")
        return

    ext = get_extension(filename)
    category = get_category(ext)
    date_folder = get_date_folder(file_path)

    # Убираем keyword_folder, кладём всегда в category/date_folder
    target_dir = os.path.join(source_folder, category, date_folder)
    os.makedirs(target_dir, exist_ok=True)

    new_path = os.path.join(target_dir, filename)
    count = 1
    while os.path.exists(new_path):
        name, ext = os.path.splitext(filename)
        new_path = os.path.join(target_dir, f"{name} ({count}){ext}")
        count += 1

    try:
        shutil.move(file_path, new_path)
        log(f"Перемещён: {filename} → {target_dir}")
        print(f"[✔] {filename} → {category}/{date_folder}")
    except Exception as e:
        log(f"Ошибка перемещения {filename}: {str(e)}")


def sort_existing_files():
    for filename in os.listdir(source_folder):
        path = os.path.join(source_folder, filename)
        move_file(path)

def setup_autostart():
    try:
        pythonw = shutil.which("pythonw")
        if not pythonw:
            log("❌ Не найден pythonw.exe")
            return

        script_path = os.path.realpath(__file__)
        startup_dir = os.path.join(os.getenv("APPDATA"), "Microsoft\\Windows\\Start Menu\\Programs\\Startup")
        bat_path = os.path.join(startup_dir, "start_fileninja.bat")

        content = f'@echo off\ncd /d "{os.path.dirname(script_path)}"\nstart "" "{pythonw}" "{script_path}"\n'

        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(content)

        log(f"✅ Автозапуск настроен: {bat_path}")
    except Exception as e:
        log(f"❌ Ошибка автозапуска: {e}")

class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        time.sleep(1.5)
        move_file(event.src_path)

if __name__ == "__main__":
    setup_autostart()
    log("Запуск FileNinja")
    print(f"🥷 FileNinja следит за: {source_folder}")

    #delete_junk_files()     # Очистка мусора при старте (раскомментировать при необходимости)
    #archive_old_files()      # Автоархивация старых файлов при старте
    sort_existing_files()    # Сортировка уже существующих файлов

    observer = Observer()
    observer.schedule(FileHandler(), source_folder, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log("FileNinja остановлен вручную")

    observer.join()