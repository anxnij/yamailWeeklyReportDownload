import imaplib
import email
from email.header import decode_header
import os
import json
import threading
import time
from pathlib import Path
from datetime import datetime, time as dtime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import requests
import urllib3
import subprocess
import sys
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_public_config_path() -> str:
    return os.path.join(get_base_dir(), "config.json")


def load_public_config() -> dict:
    default_config = {
        APP_NAME = ""
        IMAP_HOST = ""
        IMAP_PORT = 111
    }
    path = get_public_config_path()

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                user_config = json.load(f)
            default_config.update(user_config)
        except Exception:
            pass

    return default_config


APP_CONFIG = load_public_config()

NAME = APP_CONFIG["NAME"]
HOST = APP_CONFIG["HOST"]
PORT = int(APP_CONFIG["PORT"])

# Яндекс может складывать письма не только во Входящие
# (названия папок могут отличаться — если нужно, добавим ещё)
MAILBOXES = ["INBOX", "Рассылки", "Spam", "Спам"]


# ---------------- utils ----------------

def decode_mime(s: str) -> str:
    if not s:
        return ""
    parts = decode_header(s)
    out = ""
    for part, enc in parts:
        if isinstance(part, bytes):
            out += part.decode(enc or "utf-8", errors="replace")
        else:
            out += part
    return out


def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)


def atomic_replace(path: Path, data: bytes):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_bytes(data)
    tmp.replace(path)


def now() -> datetime:
    return datetime.now()


def parse_hhmm(s: str) -> dtime:
    return dtime.fromisoformat(s.strip())


def norm_sender(s: str) -> str:
    return (s or "").strip()


# ---------------- IMAP logic ----------------

def imap_login(email_addr: str, password: str):
    m = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    m.login(email_addr, password)
    return m


def search_latest_id(m, mailbox: str, sender: str):
    st, _ = m.select(mailbox)
    if st != "OK":
        return None

    # sender может быть:
    # - точный email: info@brandquad.ru
    # - домен: brandquad.ru
    # IMAP FROM ищет "подстроку", поэтому домен тоже работает.
    st, data = m.search(None, f'FROM "{sender}"')
    if st != "OK" or not data or not data[0]:
        return None

    ids = data[0].split()
    if not ids:
        return None
    return ids[-1]


def get_latest_message(m, sender: str):
    """
    Берём самое последнее письмо от sender в любой из MAILBOXES
    """
    latest_msg = None
    latest_dt = None

    for box in MAILBOXES:
        try:
            msg_id = search_latest_id(m, box, sender)
            if not msg_id:
                continue

            st, msg_data = m.fetch(msg_id, "(RFC822)")
            if st != "OK":
                continue

            msg = email.message_from_bytes(msg_data[0][1])

            # пробуем достать дату письма
            try:
                dt = email.utils.parsedate_to_datetime(msg.get("Date"))
            except Exception:
                dt = None

            if latest_dt is None:
                latest_msg = msg
                latest_dt = dt
            else:
                # если дату не смогли распарсить — оставляем первое найденное
                if dt and latest_dt and dt > latest_dt:
                    latest_msg = msg
                    latest_dt = dt
                elif dt and not latest_dt:
                    latest_msg = msg
                    latest_dt = dt

        except Exception:
            continue

    return latest_msg


def extract_links(msg) -> list[str]:
    links = []

    def grab(text: str):
        return re.findall(r'https?://[^\s"<>\']+', text, flags=re.IGNORECASE)

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype not in ("text/plain", "text/html"):
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                text = payload.decode(charset, errors="ignore")
            except Exception:
                text = payload.decode(errors="ignore")
            links += grab(text)
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            try:
                text = payload.decode(msg.get_content_charset() or "utf-8", errors="ignore")
            except Exception:
                text = payload.decode(errors="ignore")
            links += grab(text)

    # uniq
    uniq = []
    seen = set()
    for u in links:
        if u not in seen:
            uniq.append(u)
            seen.add(u)
    return uniq


def download_attachment_excel(msg, save_path: Path) -> bool:
    """
    Если в письме есть Excel-вложение — скачиваем первое.
    """
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue

        filename = part.get_filename()
        if not filename:
            continue

        filename_low = decode_mime(filename).lower()
        if not (filename_low.endswith(".xlsx") or filename_low.endswith(".xls")):
            continue

        content = part.get_payload(decode=True)
        if not content:
            continue

        atomic_replace(save_path, content)
        return True

    return False


def download_by_link(msg, save_path: Path, allowed_domain: str | None) -> bool:
    """
    Берём первую ссылку (или первую, которая содержит allowed_domain) и качаем.
    SSL verify выключен для корпоративных цепочек.
    """
    links = extract_links(msg)
    if not links:
        return False

    picked = None
    if allowed_domain:
        ad = allowed_domain.strip().lower()
        for lnk in links:
            if ad in lnk.lower():
                picked = lnk
                break
    else:
        picked = links[0]

    if not picked:
        return False

    r = requests.get(picked, timeout=60, allow_redirects=True, verify=False)
    r.raise_for_status()

    ctype = (r.headers.get("Content-Type") or "").lower()
    if "text/html" in ctype and len(r.content) < 2_000_000:
        return False

    atomic_replace(save_path, r.content)
    return True


# ---------------- App ----------------

class App(tk.Tk):
    def __init__(self, silent: bool = False):
        super().__init__()
        self.silent = silent

        self.title("Автовыгрузка (Яндекс.Почта)")
        self.geometry("860x600")

        self.cfg_path = Path(os.getenv("APPDATA", str(Path.home()))) / APP_NAME / "config.json"
        ensure_dir(self.cfg_path.parent)

        self.stop_event = threading.Event()
        self.worker = None

        self._build_ui()
        self._load_cfg()

        if self.silent:
            self.withdraw()
            self._start_bg()

    # ---------- UI ----------

    def _build_ui(self):
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)

        def row(label, var, r, show=None):
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", pady=2)
            e = ttk.Entry(f, textvariable=var, width=60)
            if show:
                e.configure(show=show)
            e.grid(row=r, column=1, sticky="we", pady=2)

        self.v_email = tk.StringVar()
        self.v_pass = tk.StringVar()
        self.v_sender = tk.StringVar()
        self.v_domain = tk.StringVar()
        self.v_folder = tk.StringVar()
        self.v_filename = tk.StringVar(value="выгрузка.xlsx")

        row("Почта (логин):", self.v_email, 0)
        row("Пароль приложения:", self.v_pass, 1, show="•")
        row("Отправитель (email или домен):", self.v_sender, 2)
        row("Домен ссылки (опционально):", self.v_domain, 3)

        ttk.Label(f, text="Папка сохранения:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(f, textvariable=self.v_folder, width=60).grid(row=4, column=1, sticky="we", pady=2)
        ttk.Button(f, text="Выбрать…", command=self._pick_folder).grid(row=4, column=2, padx=6)

        row("Имя файла (будет перезаписываться):", self.v_filename, 5)

        ttk.Label(f, text="Автозапуск ежедневно (окно времени):").grid(row=6, column=0, sticky="w", pady=2)
        self.v_from = tk.StringVar(value="11:00")
        self.v_to = tk.StringVar(value="12:00")

        tb = ttk.Frame(f)
        tb.grid(row=6, column=1, sticky="w", pady=2)
        ttk.Label(tb, text="с").pack(side="left")
        ttk.Entry(tb, textvariable=self.v_from, width=7).pack(side="left", padx=(6, 10))
        ttk.Label(tb, text="до").pack(side="left")
        ttk.Entry(tb, textvariable=self.v_to, width=7).pack(side="left", padx=6)

        btns = ttk.Frame(f)
        btns.grid(row=7, column=0, columnspan=3, pady=(12, 8), sticky="w")

        ttk.Button(btns, text="Сохранить настройки", command=self._save_cfg).pack(side="left", padx=4)
        ttk.Button(btns, text="Проверить сейчас", command=self._run_once_async).pack(side="left", padx=4)

        self.btn_bg_start = ttk.Button(btns, text="Запустить в фоне", command=self._start_bg)
        self.btn_bg_start.pack(side="left", padx=4)

        self.btn_bg_stop = ttk.Button(btns, text="Остановить", command=self._stop_bg)
        self.btn_bg_stop.pack(side="left", padx=4)

        ttk.Button(btns, text="Активировать ежедневный автозапуск", command=self._enable_autostart).pack(side="left", padx=10)
        ttk.Button(btns, text="Удалить автозапуск", command=self._disable_autostart).pack(side="left", padx=4)

        ttk.Label(f, text="Лог:").grid(row=8, column=0, sticky="nw")
        self.log = tk.Text(f, height=16, wrap="word")
        self.log.grid(row=8, column=1, columnspan=2, sticky="nsew")

        f.columnconfigure(1, weight=1)
        f.rowconfigure(8, weight=1)

        self._log("Готово. Включите автозапуск.")

    # ---------- cfg ----------

    def _pick_folder(self):
        p = filedialog.askdirectory()
        if p:
            self.v_folder.set(p)

    def _save_cfg(self):
        cfg = {
            "email": self.v_email.get().strip(),
            "password": self.v_pass.get().strip(),
            "sender": self.v_sender.get().strip(),
            "domain": self.v_domain.get().strip(),
            "folder": self.v_folder.get().strip(),
            "filename": self.v_filename.get().strip() or "выгрузка.xlsx",
            "from": self.v_from.get().strip() or "11:00",
            "to": self.v_to.get().strip() or "12:00",
        }
        self.cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        self._log(f"Настройки сохранены → {self.cfg_path}")
        if not self.silent:
            messagebox.showinfo("Ок", "Настройки сохранены.")

    def _load_cfg(self):
        if not self.cfg_path.exists():
            return
        cfg = json.loads(self.cfg_path.read_text(encoding="utf-8"))
        self.v_email.set(cfg.get("email", ""))
        self.v_pass.set(cfg.get("password", ""))
        self.v_sender.set(cfg.get("sender", ""))
        self.v_domain.set(cfg.get("domain", ""))
        self.v_folder.set(cfg.get("folder", ""))
        self.v_filename.set(cfg.get("filename", "выгрузка.xlsx"))
        self.v_from.set(cfg.get("from", "11:00"))
        self.v_to.set(cfg.get("to", "12:00"))

    def _read_config(self):
        if not self.cfg_path.exists():
            raise RuntimeError("Нет config.json. Нажмите «Сохранить настройки».")
        cfg = json.loads(self.cfg_path.read_text(encoding="utf-8"))
        if not cfg.get("email") or not cfg.get("password") or not cfg.get("sender"):
            raise RuntimeError("Заполните логин/пароль/отправителя и нажмите «Сохранить настройки».")
        # проверка времени
        parse_hhmm(cfg.get("from", "11:00"))
        parse_hhmm(cfg.get("to", "12:00"))
        return cfg

    # ---------- log ----------

    def _log(self, s: str):
        try:
            self.log.insert("end", f"[{now().strftime('%Y-%m-%d %H:%M:%S')}] {s}\n")
            self.log.see("end")
        except Exception:
            pass

    # ---------- core ----------

    def _run_once_async(self):
        def work():
            self.run_once()
        threading.Thread(target=work, daemon=True).start()

    def run_once(self):
        try:
            cfg = self._read_config()

            email_addr = cfg["email"]
            password = cfg["password"]
            sender = norm_sender(cfg["sender"])
            allowed_domain = (cfg.get("domain") or "").strip() or None

            folder = Path(cfg.get("folder") or "")
            if not str(folder):
                raise RuntimeError("Не указана папка сохранения.")
            ensure_dir(folder)
            save_path = folder / (cfg.get("filename") or "выгрузка.xlsx")

            self._log(f"Проверяю почту {email_addr}…")

            m = imap_login(email_addr, password)
            try:
                msg = get_latest_message(m, sender)
                if not msg:
                    self._log("Писем от отправителя не найдено ни в одной папке.")
                    return

                subj = decode_mime(msg.get("Subject", ""))
                self._log(f"Найдено последнее письмо. Тема: {subj}")

                if download_attachment_excel(msg, save_path):
                    self._log(f"Скачано из вложения → {save_path}")
                    return

                if download_by_link(msg, save_path, allowed_domain):
                    self._log(f"Скачано по ссылке → {save_path}")
                    return

                self._log("В письме нет Excel-вложения и подходящей ссылки.")
            finally:
                try:
                    m.logout()
                except Exception:
                    pass

        except Exception as e:
            self._log(f"Ошибка: {e}")
            if not self.silent:
                messagebox.showerror("Ошибка", str(e))

    # ---------- time window + background ----------

    def _in_time_window(self, start_str: str, end_str: str) -> bool:
        t1 = parse_hhmm(start_str)
        t2 = parse_hhmm(end_str)
        cur = now().time()
        return t1 <= cur <= t2

    def _passed_window_today(self, end_str: str) -> bool:
        t2 = parse_hhmm(end_str)
        return now().time() > t2

    def _start_bg(self):
        if self.worker and self.worker.is_alive():
            return

        self.stop_event.clear()
        self._log("Фоновый режим включён.")

        def loop():
            while not self.stop_event.is_set():
                try:
                    cfg = self._read_config()
                    start_s = cfg.get("from", "11:00")
                    end_s = cfg.get("to", "12:00")

                    if self._in_time_window(start_s, end_s):
                        self.run_once()
                    else:
                        self._log("Сейчас вне окна автозапуска — ничего не происходит.")

                    if self.silent and self._passed_window_today(end_s):
                        self._log("Окно времени прошло. Завершение работы (silent).")
                        try:
                            self.stop_event.set()
                            self.destroy()
                        except Exception:
                            os._exit(0)
                        return

                except Exception as e:
                    self._log(f"Ошибка фона: {e}")

                for _ in range(300):  # 5 минут
                    if self.stop_event.is_set():
                        break
                    time.sleep(1)

        self.worker = threading.Thread(target=loop, daemon=True)
        self.worker.start()

    def _stop_bg(self):
        self.stop_event.set()
        self._log("Фоновый режим остановлен.")

    # ---------- Windows Task Scheduler ----------

    def _enable_autostart(self):
        """
        Создаёт задачу планировщика:
        - ежедневно в время start (v_from)
        - запускает этот EXE (или python+скрипт в dev) с --silent
        """
        try:
            cfg = self._read_config()
            start_time = cfg.get("from", "11:00").strip()
            task_name = "YandexAutoDownloaderDaily"

            exe_path = sys.executable
            script_path = os.path.abspath(sys.argv[0])

            # Если это собранный exe — запускаем его
            # Если это dev режим — запускаем python + main.py
            if exe_path.lower().endswith(".exe") and "python.exe" not in exe_path.lower():
                tr = f'"{exe_path}" --silent'
            else:
                tr = f'"{exe_path}" "{script_path}" --silent'

            cmd = [
                "schtasks",
                "/Create",
                "/F",
                "/SC", "DAILY",
                "/ST", start_time,
                "/TN", task_name,
                "/TR", tr
            ]

            subprocess.run(cmd, check=True)
            self._log(f"Автозапуск активирован: ежедневно в {start_time}")
            if not self.silent:
                messagebox.showinfo(
                    "Готово",
                    f"Автозапуск включён.\nКаждый день в {start_time} программа будет запускаться сама."
                )
        except Exception as e:
            self._log(f"Ошибка автозапуска: {e}")
            if not self.silent:
                messagebox.showerror("Ошибка", str(e))

    def _disable_autostart(self):
        try:
            task_name = "YandexAutoDownloaderDaily"
            cmd = ["schtasks", "/Delete", "/F", "/TN", task_name]
            subprocess.run(cmd, check=True)
            self._log("Автозапуск удалён.")
            if not self.silent:
                messagebox.showinfo("Готово", "Автозапуск удалён.")
        except Exception as e:
            self._log(f"Ошибка удаления автозапуска: {e}")
            if not self.silent:
                messagebox.showerror("Ошибка", str(e))


def main():
    silent = ("--silent" in sys.argv)
    app = App(silent=silent)
    app.mainloop()


if __name__ == "__main__":
    main()