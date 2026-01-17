import os
import sys
import csv
import json
import threading
import socket
import queue
import urllib.parse
import urllib.request
from datetime import datetime, date, timedelta

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkinter import font as tkfont

import keyboard  # global hotkey + RFID/klavye
import pystray
from PIL import Image, ImageDraw
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import winsound  # SES

# ---------- GLOBAL PATHS ----------

# FIX: EXE (onedir/onefile) çalışınca __file__ _internal'ı gösterebiliyor.
# Bu yüzden frozen modda sys.executable (main.exe) konumunu baz alıyoruz.
if getattr(sys, "frozen", False) and hasattr(sys, "executable"):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOGS_DIR = os.path.join(BASE_DIR, "logs")
KISILER_PATH = os.path.join(LOGS_DIR, "Kisiler.csv")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")

KISILER_FIELDNAMES = [
    "Kart Numarası",
    "Ad Soyad",
    "TC Numarası",
    "Cep Telefonu Numarası",
    "İşe Başlama Tarihi",
    "Durum",
    "Adres",
    "Kan Grubu",
    "Acil Durum Kişi Adı",
    "Acil Durum Numarası",
    "Görev / Pozisyon",
    "Notlar",
    "Kayıt Tarihi",
    "Senkronize"
]

FIREBASE_API_KEY = "AIzaSyDBxfIpSQbZ93tIvrnf6xB2b4OQMtzrWVI"
FIREBASE_EMAIL = "uploader@eposta.com"
FIREBASE_PASSWORD = "A9@xF3!qW6"
FIREBASE_BUCKET = "mesaitakip-4d2d8.firebasestorage.app"

UPLOAD_DEBOUNCE_SECONDS = 10
_debounce_lock = threading.Lock()
_debounce_timers = {}
_pending_upload_paths = {}


def _firebase_login():
    payload = json.dumps({
        "email": FIREBASE_EMAIL,
        "password": FIREBASE_PASSWORD,
        "returnSecureToken": True
    }).encode("utf-8")
    url = (
        "https://identitytoolkit.googleapis.com/v1/accounts:"
        f"signInWithPassword?key={FIREBASE_API_KEY}"
    )
    request = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data.get("idToken")


def _firebase_upload(local_path: str, remote_name: str, token: str):
    if not os.path.exists(local_path):
        return
    with open(local_path, "rb") as f:
        body = f.read()
    encoded_name = urllib.parse.quote(remote_name, safe="")
    url = (
        "https://firebasestorage.googleapis.com/v0/b/"
        f"{FIREBASE_BUCKET}/o?uploadType=media&name={encoded_name}"
    )
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "text/csv"
        },
        method="POST"
    )
    with urllib.request.urlopen(request, timeout=30):
        pass


def _upload_with_login(local_path: str, remote_name: str):
    token = None
    try:
        token = _firebase_login()
        if not token:
            return
        _firebase_upload(local_path, remote_name, token)
    except Exception:
        pass
    finally:
        token = None


def _start_upload_thread(local_path: str, remote_name: str):
    thread = threading.Thread(
        target=_upload_with_login,
        args=(local_path, remote_name),
        daemon=True
    )
    thread.start()


def _debounce_fire(kind: str):
    with _debounce_lock:
        local_path = _pending_upload_paths.get(kind)
    if not local_path:
        return
    if kind == "KISILER":
        remote_name = "logs/Kisiler.csv"
    else:
        remote_name = f"logs/{os.path.basename(local_path)}"
    _start_upload_thread(local_path, remote_name)


def on_csv_written(kind: str, local_path: str):
    with _debounce_lock:
        existing = _debounce_timers.get(kind)
        if existing:
            existing.cancel()
        _pending_upload_paths[kind] = local_path
        timer = threading.Timer(UPLOAD_DEBOUNCE_SECONDS, _debounce_fire, args=(kind,))
        timer.daemon = True
        _debounce_timers[kind] = timer
        timer.start()

def load_settings():
    if not os.path.exists(SETTINGS_PATH):
        return {"management_password": "5353", "popup_x": None, "popup_y": None, "panel_mode": "main"}
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            data = {}
    except Exception:
        data = {}
    if "management_password" not in data or not isinstance(data.get("management_password"), str):
        data["management_password"] = "5353"
    if "popup_x" not in data:
        data["popup_x"] = None
    if "popup_y" not in data:
        data["popup_y"] = None
    if "panel_mode" not in data or data.get("panel_mode") not in ("main", "management"):
        data["panel_mode"] = "main"
    return data

def save_settings(data: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_management_password():
    s = load_settings()
    return str(s.get("management_password", "5353"))

def set_management_password(new_pwd: str):
    s = load_settings()
    s["management_password"] = new_pwd
    save_settings(s)

def get_saved_popup_position():
    s = load_settings()
    return s.get("popup_x"), s.get("popup_y")

def set_saved_popup_position(x: int, y: int):
    s = load_settings()
    s["popup_x"] = int(x)
    s["popup_y"] = int(y)
    save_settings(s)

def get_panel_mode():
    s = load_settings()
    return s.get("panel_mode", "main")

def set_panel_mode(mode: str):
    s = load_settings()
    s["panel_mode"] = mode
    save_settings(s)

def restart_app():
    try:
        os.execl(sys.executable, sys.executable, *sys.argv)
    except Exception:
        try:
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception:
            pass

def normalize_kisi_row(row: dict) -> dict:
    out = {k: "0" for k in KISILER_FIELDNAMES}
    for k in KISILER_FIELDNAMES:
        if k in row and row.get(k) not in (None, ""):
            out[k] = str(row.get(k))
    # defaults
    if out["Senkronize"] in (None, ""):
        out["Senkronize"] = "0"
    return out


# ---------- SINGLE INSTANCE (SOCKET LOCK) ----------

INSTANCE_SOCKET_PORT = 65432
_instance_socket = None


def ensure_single_instance():
    """Aynı anda sadece 1 exe çalışsın."""
    global _instance_socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", INSTANCE_SOCKET_PORT))
    except OSError:
        s.close()
        sys.exit(0)
    s.listen(1)
    _instance_socket = s


# ---------- UTILITIES ----------

def ensure_directories_and_files():
    if not os.path.isdir(LOGS_DIR):
        os.makedirs(LOGS_DIR, exist_ok=True)

    if not os.path.exists(KISILER_PATH):
        with open(KISILER_PATH, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(KISILER_FIELDNAMES)

def turkish_month_name(m):
    months = [
        "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
        "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
    ]
    return months[m - 1]


def format_turkish_date_long(d: date) -> str:
    return f"{d.day} {turkish_month_name(d.month)} {d.year}"


def format_date_dotted(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def parse_dotted_date(s: str) -> date:
    return datetime.strptime(s, "%d.%m.%Y").date()


def get_month_log_path(d: date) -> str:
    return os.path.join(LOGS_DIR, d.strftime("%Y-%m") + ".csv")


def ensure_month_log_header(path: str):
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Tarih",
                "Saat",
                "Kart Numarası",
                "Ad Soyad",
                "Kaynak",
                "Senkron"
            ])


def sanitize_turkish(text: str) -> str:
    """
    PDF'te Türkçe karakterler bozulmasın diye
    ASCII'ye yakın dönüştürme (ş->s, ğ->g, vb.).
    """
    mapping = {
        "ş": "s", "Ş": "S",
        "ğ": "g", "Ğ": "G",
        "ı": "i", "İ": "I",
        "ö": "o", "Ö": "O",
        "ü": "u", "Ü": "U",
        "ç": "c", "Ç": "C",
    }
    return "".join(mapping.get(ch, ch) for ch in text)


# ---------- GLOBAL STATE ----------

active_cards = {}  # card_id -> person dict
root = None
date_label = None
header_label = None
tree = None
total_label = None
current_ui_date = date.today()

card_queue = queue.Queue()

tray_icon = None
tray_thread = None
tray_started = False  # tray bir kere çalışsın

_buffer = []  # RFID/klavye buffer

_management_buttons_visible = False
_btn_new = None
_btn_manual = None
_btn_report = None
_btn_back = None
_btn_personel = None
_btn_ayarlar = None
_btn_yonetim = None

# TEK DEĞİŞİKLİK: Yönetim menüsü penceresi duplike açılmasın diye referans
_management_menu_win = None


# ---------- LOAD / SAVE KISILER ----------

def load_kisiler():
    global active_cards
    active_cards = {}
    if not os.path.exists(KISILER_PATH):
        return
    with open(KISILER_PATH, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            row_n = normalize_kisi_row(row)
            if row_n.get("Durum") == "Kart Aktif":
                card_id = row_n.get("Kart Numarası", "").strip()
                if card_id:
                    active_cards[card_id] = row_n


def save_kisiler_rows(rows):
    with open(KISILER_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=KISILER_FIELDNAMES)
        writer.writeheader()
        for r in rows:
            writer.writerow(normalize_kisi_row(r))
    on_csv_written("KISILER", KISILER_PATH)


# ---------- DUPLICATE CHECK ----------

def find_last_log_dt(ad_soyad: str, card_id: str, target_month_date: date):
    log_path = get_month_log_path(target_month_date)
    if not os.path.exists(log_path):
        return None
    last_dt = None
    with open(log_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("Ad Soyad") != ad_soyad:
                continue
            if row.get("Kart Numarası") != card_id:
                continue
            try:
                d = datetime.strptime(row["Tarih"], "%Y-%m-%d").date()
                t = datetime.strptime(row["Saat"], "%H:%M").time()
                row_dt = datetime.combine(d, t)
            except Exception:
                continue
            if last_dt is None or row_dt > last_dt:
                last_dt = row_dt
    return last_dt


def is_duplicate_log(ad_soyad: str, card_id: str, new_dt: datetime):
    last_dt = find_last_log_dt(ad_soyad, card_id, new_dt.date())
    if last_dt is None:
        return False, None
    diff = new_dt - last_dt
    # SADECE İLERİ YÖNDE 30 DAKİKA
    if timedelta(0) <= diff < timedelta(minutes=30):
        return True, last_dt
    return False, last_dt


# ---------- UI HELPERS ----------

def ensure_today_ui():
    global current_ui_date
    today = date.today()
    if today != current_ui_date:
        current_ui_date = today
        refresh_date_labels()
        reload_today_table()


def refresh_date_labels():
    if date_label is not None:
        d = current_ui_date
        date_label.config(text=format_turkish_date_long(d))


def reload_today_table():
    for item in tree.get_children():
        tree.delete(item)
    count = 0
    today = current_ui_date
    log_path = get_month_log_path(today)
    if os.path.exists(log_path):
        with open(log_path, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    row_date = datetime.strptime(row["Tarih"], "%Y-%m-%d").date()
                except Exception:
                    continue
                if row_date != today:
                    continue
                saat = row.get("Saat", "")
                ad = row.get("Ad Soyad", "")
                card = row.get("Kart Numarası", "")
                tree.insert("", "end", values=(saat, ad, card))
                count += 1
    update_total_label(count)


def update_total_label(count: int):
    if total_label is not None:
        total_label.config(text=f"Toplam Okutma Sayısı: {count}")
    tree._lavoni_count = count


def increment_total_label():
    current = getattr(tree, "_lavoni_count", 0)
    current += 1
    update_total_label(current)


# ---------- POPUPS ----------

def show_popup_normal(ad_soyad: str, card_id: str, dt: datetime):
    win = tk.Toplevel(root)
    win.title("KART OKUMA")
    win.overrideredirect(True)
    win.attributes("-topmost", True)

    width, height = 600, 260
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    margin = 20
    saved_x, saved_y = get_saved_popup_position()
    if saved_x is None or saved_y is None:
        x = sw - width - margin
        y = sh - height - margin
    else:
        x = int(saved_x)
        y = int(saved_y)
    win.geometry(f"{width}x{height}+{x}+{y}")

    frame = tk.Frame(win, bg="#222222")
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="KART OKUMA", bg="#444444", fg="white",
                     font=("Segoe UI", 16, "bold"))
    title.pack(fill="x", pady=(0, 10))

    name_label = tk.Label(frame, text=ad_soyad, bg="#222222", fg="white",
                          font=("Segoe UI", 22, "bold"))
    name_label.pack(pady=10)

    card_label = tk.Label(frame, text=f"Kart Numarası: {card_id}", bg="#222222",
                          fg="#bbbbbb", font=("Segoe UI", 14))
    card_label.pack(pady=5)

    dt_str = f"{dt.strftime('%d.%m.%Y')}  {dt.strftime('%H:%M')}"
    time_label = tk.Label(frame, text=dt_str, bg="#222222", fg="#4aa3ff",
                          font=("Segoe UI", 16, "bold"))
    time_label.pack(pady=15)

    win.after(3000, win.destroy)


def show_popup_duplicate(ad_soyad: str, card_id: str, last_dt: datetime):
    win = tk.Toplevel(root)
    win.title("DUPLIKE OKUMA")
    win.overrideredirect(True)
    win.attributes("-topmost", True)

    width, height = 600, 260
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    margin = 20
    saved_x, saved_y = get_saved_popup_position()
    if saved_x is None or saved_y is None:
        x = sw - width - margin
        y = sh - height - margin
    else:
        x = int(saved_x)
        y = int(saved_y)
    win.geometry(f"{width}x{height}+{x}+{y}")

    frame = tk.Frame(win, bg="#331111")
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="DUPLIKE OKUMA", bg="#772222", fg="white",
                     font=("Segoe UI", 16, "bold"))
    title.pack(fill="x", pady=(0, 10))

    name_label = tk.Label(frame, text=ad_soyad, bg="#331111", fg="white",
                          font=("Segoe UI", 22, "bold"))
    name_label.pack(pady=10)

    card_label = tk.Label(frame, text=f"Kart Numarası: {card_id}", bg="#331111",
                          fg="#ffbbbb", font=("Segoe UI", 14))
    card_label.pack(pady=5)

    if last_dt is not None:
        dt_str = f"Son okutma: {last_dt.strftime('%d.%m.%Y')}  {last_dt.strftime('%H:%M')}"
    else:
        dt_str = "Son okutma: -"

    last_label = tk.Label(frame, text=dt_str, bg="#331111", fg="#ffccaa",
                          font=("Segoe UI", 14))
    last_label.pack(pady=5)

    warn_label = tk.Label(frame,
                          text="Bu kart son 30 dakika içinde okutulmuştur.",
                          bg="#331111", fg="#ffaaaa",
                          font=("Segoe UI", 14, "bold"))
    warn_label.pack(pady=15)

    win.after(3000, win.destroy)


# ---------- LOGGING & PROCESSING ----------

def add_log_and_update(ad_soyad: str, card_id: str, dt: datetime,
                       kaynak: str, show_popup: bool):
    load_kisiler()
    ensure_today_ui()

    is_dup, last_dt = is_duplicate_log(ad_soyad, card_id, dt)
    if is_dup:
        try:
            winsound.MessageBeep(winsound.MB_ICONHAND)
        except Exception:
            pass
        show_popup_duplicate(ad_soyad, card_id, last_dt)
        return

    log_path = get_month_log_path(dt.date())
    ensure_month_log_header(log_path)
    with open(log_path, "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            dt.date().strftime("%Y-%m-%d"),
            dt.strftime("%H:%M"),
            card_id,
            ad_soyad,
            kaynak,
            "0"
        ])
    on_csv_written("MONTH", log_path)

    reload_today_table()

    # OLUMLU SES
    try:
        winsound.PlaySound(
            r"C:\Windows\Media\Speech On.wav",
            winsound.SND_FILENAME | winsound.SND_ASYNC
        )
    except Exception:
        pass

    if show_popup:
        show_popup_normal(ad_soyad, card_id, dt)


def process_card_from_queue():
    load_kisiler()
    while not card_queue.empty():
        source, card_id = card_queue.get()
        card_id = card_id.strip()
        if not card_id:
            continue
        person = active_cards.get(card_id)
        if not person:
            continue
        ad_soyad = person.get("Ad Soyad", "").strip()
        now = datetime.now()
        add_log_and_update(ad_soyad, card_id, now, source, show_popup=True)
    root.after(100, process_card_from_queue)


# ---------- KEYBOARD / RFID LISTENER ----------

def keyboard_event(e):
    global _buffer
    if e.event_type != "down":
        return
    name = e.name
    if name == "enter":
        card_id = "".join(_buffer).strip()
        _buffer = []
        if card_id:
            card_queue.put(("RFID", card_id))
    elif len(name) == 1:
        _buffer.append(name)


def start_keyboard_listener():
    keyboard.hook(keyboard_event)
    keyboard.add_hotkey("ctrl+alt+l", lambda: show_main_window())


# ---------- SYSTEM TRAY ----------

def create_tray_image(color=(0, 0, 0, 255)):
    # İconu DOSYADAN OKUMADAN, birebir piksel maskesiyle çiz
    base = Image.new("RGBA", (16, 16), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)

    # (y, x_start, x_end) run-length çizim listesi
    runs = [
        (2, 4, 12),
        (3, 3, 13),
        (4, 2, 3), (4, 13, 14),
        (5, 2, 3), (5, 13, 14),
        (6, 2, 3), (6, 8, 8), (6, 13, 14),
        (7, 2, 3), (7, 7, 9), (7, 13, 14),
        (8, 2, 3), (8, 8, 8), (8, 13, 14),
        (9, 2, 3), (9, 6, 7), (9, 9, 10), (9, 13, 14),
        (10, 2, 3), (10, 6, 10), (10, 13, 14),
        (11, 2, 3), (11, 7, 9), (11, 13, 14),
        (12, 2, 3), (12, 13, 14),
        (13, 3, 13),
        (14, 4, 12),
    ]

    for y, x1, x2 in runs:
        d.rectangle([x1, y, x2, y], fill=color)

    # Tray için büyüt (pixel-art net kalsın)
    return base.resize((64, 64), Image.NEAREST)


def request_app_exit():
    # Tk ana thread üzerinden tam çıkış
    def _exit_now():
        try:
            if tray_icon is not None:
                try:
                    tray_icon.stop()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass
        try:
            os._exit(0)
        except Exception:
            pass

    try:
        root.after(0, _exit_now)
    except Exception:
        try:
            os._exit(0)
        except Exception:
            pass


def tray_run():
    global tray_icon
    image = create_tray_image((255, 0, 0, 255)) if get_panel_mode() == "management" else create_tray_image((0, 0, 0, 255))

    def _tray_show(icon, item):
        # Tray thread içinden Tk arayüzüne güvenli dönüş
        try:
            if get_panel_mode() == "management":
                root.after(0, open_management_menu)
            else:
                root.after(0, show_main_window)
        except Exception:
            pass

    def _tray_quit(icon, item):
        try:
            icon.stop()
        except Exception:
            pass
        try:
            request_app_exit()
        except Exception:
            pass

    # SADECE BURASI DEĞİŞTİ: Ana Panel'de "Çıkış" menüde gösterilmez
    if get_panel_mode() == "management":
        menu = pystray.Menu(
            pystray.MenuItem("Aç", _tray_show, default=True),
            pystray.MenuItem("Çıkış", _tray_quit)
        )
    else:
        menu = pystray.Menu(
            pystray.MenuItem("Aç", _tray_show, default=True)
        )

    tray_icon = pystray.Icon("Lavoni", image, "Lavoni Kayıt Sistemi", menu=menu)
    tray_icon.run()


# ---------- DARK THEME HELPERS ----------

def style_dark_window(win):
    win.configure(bg="#222222")


def create_dark_button(parent, text, command=None):
    return tk.Button(parent, text=text, command=command,
                     bg="#444444", fg="white",
                     activebackground="#555555",
                     activeforeground="white",
                     relief="raised")


# ---------- NEW CARD WINDOW ----------

def open_new_card_window():
    load_kisiler()

    def save():
        card_id = entry_card.get().strip()
        ad = entry_name.get().strip()
        tc = entry_tc.get().strip()
        tel_raw = entry_phone.get().strip()
        ise_baslama = entry_start.get().strip()

        if not card_id or not ad or not tc or not tel_raw or not ise_baslama:
            messagebox.showerror("Hata", "Tüm alanlar zorunludur.")
            return

        if not (tc.isdigit() and len(tc) == 11):
            messagebox.showerror("Hata", "TC Numarası 11 haneli rakamlardan oluşmalıdır.")
            return

        digits = "".join(ch for ch in tel_raw if ch.isdigit())
        if not (len(digits) == 10 and digits.startswith("5")):
            messagebox.showerror("Hata", "Cep telefonu 5XXXXXXXXX formatında 10 haneli olmalıdır.")
            return
        tel = digits

        try:
            d_ise = parse_dotted_date(ise_baslama)
        except Exception:
            messagebox.showerror("Hata", "İşe Başlama Tarihi gg.aa.yyyy formatında olmalıdır.")
            return

        today = date.today()
        if not (d_ise.year == today.year and d_ise.month == today.month):
            messagebox.showerror("Hata", "İşe Başlama Tarihi sadece içinde bulunduğunuz ayda olabilir.")
            return

        rows = []
        if os.path.exists(KISILER_PATH):
            with open(KISILER_PATH, "r", encoding="utf-8", newline="") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    rows.append(r)

        for r in rows:
            if r.get("Kart Numarası") == card_id and r.get("Durum") == "Kart Aktif":
                r["Durum"] = "Kart Devredildi"

        now = datetime.now()
        yeni_satir = {
            "Kart Numarası": card_id,
            "Ad Soyad": ad,
            "TC Numarası": tc,
            "Cep Telefonu Numarası": tel,
            "İşe Başlama Tarihi": ise_baslama,
            "Durum": "Kart Aktif",
            "Kayıt Tarihi": now.strftime("%d.%m.%Y %H:%M"),
            "Senkronize": "0"
        }
        rows.append(yeni_satir)
        save_kisiler_rows(rows)
        load_kisiler()
        messagebox.showinfo("Başarılı", "Kart başarıyla tanımlandı.")
        win.destroy()

    win = tk.Toplevel(root)
    win.title("Yeni Kart Tanımla")
    win.grab_set()
    style_dark_window(win)

    tk.Label(win, text="Kart Numarası:", bg="#222222", fg="white").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    entry_card = tk.Entry(win, width=30)
    entry_card.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(win, text="Ad Soyad:", bg="#222222", fg="white").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_name = tk.Entry(win, width=30)
    entry_name.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(win, text="TC Numarası:", bg="#222222", fg="white").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    entry_tc = tk.Entry(win, width=30)
    entry_tc.grid(row=2, column=1, padx=5, pady=5)

    tk.Label(win, text="Cep Telefonu Numarası:", bg="#222222", fg="white").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    entry_phone = tk.Entry(win, width=30)
    entry_phone.grid(row=3, column=1, padx=5, pady=5)

    tk.Label(win, text="İşe Başlama Tarihi (gg.aa.yyyy):", bg="#222222", fg="white").grid(row=4, column=0, sticky="e", padx=5, pady=5)
    entry_start = tk.Entry(win, width=30)
    entry_start.grid(row=4, column=1, padx=5, pady=5)
    entry_start.insert(0, format_date_dotted(date.today()))

    btn_save = create_dark_button(win, "Kaydet", save)
    btn_save.grid(row=5, column=0, columnspan=2, pady=10)


# ---------- MANUAL LOG WINDOW ----------

def open_manual_log_window():
    load_kisiler()

    if not active_cards:
        messagebox.showerror("Hata", "Aktif kart bulunmuyor.")
        return

    win = tk.Toplevel(root)
    win.title("Manuel Kayıt")
    win.grab_set()
    style_dark_window(win)

    cards_list = sorted(active_cards.items(), key=lambda kv: kv[1].get("Ad Soyad", ""))
    display_values = [f"{v.get('Ad Soyad','')} - {k}" for k, v in cards_list]

    tk.Label(win, text="Kişi (Ad Soyad - Kart):", bg="#222222", fg="white").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    combo_person = ttk.Combobox(win, values=display_values, state="readonly", width=40)
    combo_person.grid(row=0, column=1, padx=5, pady=5)
    if display_values:
        combo_person.current(0)

    tk.Label(win, text="Tarih (gg.aa.yyyy):", bg="#222222", fg="white").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    entry_date = tk.Entry(win, width=20)
    entry_date.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    entry_date.insert(0, format_date_dotted(date.today()))

    tk.Label(win, text="Saat (SS:dd):", bg="#222222", fg="white").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    entry_time = tk.Entry(win, width=20)
    entry_time.grid(row=2, column=1, padx=5, pady=5, sticky="w")
    entry_time.insert(0, datetime.now().strftime("%H:%M"))

    def save_manual():
        sel = combo_person.get().strip()
        if not sel:
            messagebox.showerror("Hata", "Lütfen bir kişi seçin.")
            return
        try:
            ad_part, card_part = sel.rsplit(" - ", 1)
        except ValueError:
            messagebox.showerror("Hata", "Seçim formatı hatalı.")
            return
        ad_soyad = ad_part.strip()
        card_id = card_part.strip()

        d_str = entry_date.get().strip()
        t_str = entry_time.get().strip()

        try:
            d = parse_dotted_date(d_str)
        except Exception:
            messagebox.showerror("Hata", "Tarih gg.aa.yyyy formatında olmalıdır.")
            return

        today = date.today()
        if not (d.year == today.year and d.month == today.month):
            messagebox.showerror("Hata", "Manuel giriş sadece içinde bulunduğunuz ay için yapılabilir.")
            return
        if d > today:
            messagebox.showerror("Hata", "Bugünden ileri tarih için manuel kayıt yapılamaz.")
            return

        try:
            t = datetime.strptime(t_str, "%H:%M").time()
        except Exception:
            messagebox.showerror("Hata", "Saat SS:dd formatında olmalıdır.")
            return

        dt = datetime.combine(d, t)
        add_log_and_update(ad_soyad, card_id, dt, "Manuel", show_popup=True)
        win.destroy()

    btn_save = create_dark_button(win, "Kaydet", save_manual)
    btn_save.grid(row=3, column=0, columnspan=2, pady=10)


# ---------- REPORT GENERATION ----------

def generate_report(for_current_month: bool):
    font_name = "Helvetica"
    try:
        font_path = os.path.join(BASE_DIR, "DejaVuSans.ttf")
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont("DejaVuSans", font_path))
            font_name = "DejaVuSans"
    except Exception:
        font_name = "Helvetica"

    today = date.today()
    if for_current_month:
        target_year = today.year
        target_month = today.month
        report_label_date = today.strftime("%d.%m.%Y")
        filename = f"Güncel Ay Raporu {report_label_date} Raporu.pdf"
    else:
        if today.month == 1:
            target_year = today.year - 1
            target_month = 12
        else:
            target_year = today.year
            target_month = today.month - 1
        report_label_date = f"{target_month:02d}.{target_year}"
        filename = f"Geçmiş Ay Raporu {report_label_date}.pdf"

    dummy_date = date(target_year, target_month, 1)
    log_path = get_month_log_path(dummy_date)
    if not os.path.exists(log_path):
        messagebox.showinfo("Bilgi", "Bu ay için log kaydı bulunamadı.")
        return

    logs = []
    with open(log_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                d = datetime.strptime(row["Tarih"], "%Y-%m-%d").date()
                t = datetime.strptime(row["Saat"], "%H:%M").time()
            except Exception:
                continue
            ad = row.get("Ad Soyad", "")
            card = row.get("Kart Numarası", "")
            logs.append((ad, card, d, t))

    if not logs:
        messagebox.showinfo("Bilgi", "Bu ay için log kaydı bulunamadı.")
        return

    from collections import defaultdict
    import calendar

    day_groups = defaultdict(list)
    for ad, card, d, t in logs:
        day_groups[(ad, card, d)].append(t)

    person_days = defaultdict(dict)

    for (ad, card, d), times in day_groups.items():
        times_sorted = sorted(times)
        if len(times_sorted) == 0:
            start_t, end_t, minutes = None, None, None
        elif len(times_sorted) == 1:
            start_t = times_sorted[0]
            end_t = None
            minutes = None
        else:
            start_t = times_sorted[0]
            end_t = times_sorted[-1]
            dt_start = datetime.combine(d, start_t)
            dt_end = datetime.combine(d, end_t)
            diff_min = int((dt_end - dt_start).total_seconds() // 60)
            minutes = max(diff_min, 0)
        person_days[(ad, card)][d] = (start_t, end_t, minutes)

    person_totals = {}
    for key, days in person_days.items():
        total_min = 0
        for _, (_, _, mins) in days.items():
            if mins is not None:
                total_min += mins
        person_totals[key] = total_min

    sorted_persons = sorted(person_totals.items(),
                            key=lambda kv: (kv[0][0], kv[0][1]))

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = BASE_DIR
    pdf_path = os.path.join(desktop, filename)

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    # ÖZET SAYFA
    c.setFont(font_name, 16)
    title = f"{'Güncel Ay' if for_current_month else 'Geçmiş Ay'} Raporu - {report_label_date}"
    c.drawString(40, height - 50, sanitize_turkish(title))

    x_ad = 40
    x_kart = 260
    x_toplam = 430

    y_header = height - 90
    c.setFont(font_name, 12)
    c.drawString(x_ad, y_header, sanitize_turkish("Ad Soyad"))
    c.drawString(x_kart, y_header, sanitize_turkish("Kart Numarasi"))
    c.drawString(x_toplam, y_header, sanitize_turkish("Toplam Mesai (saat:dakika)"))

    c.setLineWidth(0.5)
    c.line(40, y_header - 2, width - 40, y_header - 2)

    y = y_header - 20
    c.setFont(font_name, 11)

    for (ad, card), total_min in sorted_persons:
        hours = total_min // 60
        mins = total_min % 60
        total_str = f"{hours:02d}:{mins:02d}"

        if y < 80:
            c.showPage()
            c.setFont(font_name, 16)
            c.drawString(40, height - 50, sanitize_turkish(title))
            c.setFont(font_name, 12)
            y_header = height - 90
            c.drawString(x_ad, y_header, sanitize_turkish("Ad Soyad"))
            c.drawString(x_kart, y_header, sanitize_turkish("Kart Numarasi"))
            c.drawString(x_toplam, y_header, sanitize_turkish("Toplam Mesai (saat:dakika)"))
            c.setLineWidth(0.5)
            c.line(40, y_header - 2, width - 40, y_header - 2)
            y = y_header - 20
            c.setFont(font_name, 11)

        c.drawString(x_ad, y, sanitize_turkish(ad))
        c.drawString(x_kart, y, sanitize_turkish(card))
        c.drawString(x_toplam, y, total_str)
        # satır altına çizgi
        c.line(40, y - 2, width - 40, y - 2)
        y -= 18

    c.showPage()

    # DETAY SAYFALAR
    _, last_day = calendar.monthrange(target_year, target_month)
    month_dates = [date(target_year, target_month, d)
                   for d in range(1, last_day + 1)]

    x_tarih = 40
    x_bas = 170
    x_bit = 270
    x_mes = 380

    for (ad, card), _total_min in sorted_persons:
        c.setFont(font_name, 16)
        c.drawString(40, height - 50, sanitize_turkish(f"{ad} - {card}"))

        c.setFont(font_name, 12)
        y_header = height - 90
        c.drawString(x_tarih, y_header, sanitize_turkish("Tarih"))
        c.drawString(x_bas, y_header, sanitize_turkish("Baslangic"))
        c.drawString(x_bit, y_header, sanitize_turkish("Bitis"))
        c.drawString(x_mes, y_header, sanitize_turkish("Gunluk Mesai"))

        c.setLineWidth(0.5)
        c.line(40, y_header - 2, width - 40, y_header - 2)

        y = y_header - 20
        c.setFont(font_name, 11)

        days = person_days.get((ad, card), {})
        for d in month_dates:
            if y < 80:
                c.showPage()
                c.setFont(font_name, 16)
                c.drawString(40, height - 50, sanitize_turkish(f"{ad} - {card}"))
                c.setFont(font_name, 12)
                y_header = height - 90
                c.drawString(x_tarih, y_header, sanitize_turkish("Tarih"))
                c.drawString(x_bas, y_header, sanitize_turkish("Baslangic"))
                c.drawString(x_bit, y_header, sanitize_turkish("Bitis"))
                c.drawString(x_mes, y_header, sanitize_turkish("Gunluk Mesai"))
                c.setLineWidth(0.5)
                c.line(40, y_header - 2, width - 40, y_header - 2)
                y = y_header - 20
                c.setFont(font_name, 11)

            date_str = d.strftime("%d.%m.%Y")

            if d in days:
                st, et, mins = days[d]
                if st is None:
                    bas = "OFF"
                    bit = "OFF"
                    mes = "-"
                else:
                    bas = st.strftime("%H:%M")
                    if et is None:
                        bit = ""
                        mes = "-"
                    else:
                        bit = et.strftime("%H:%M")
                        if mins is None:
                            mes = "-"
                        else:
                            h = mins // 60
                            m = mins % 60
                            mes = f"{h:02d}:{m:02d}"
            else:
                bas = "OFF"
                bit = "OFF"
                mes = "-"

            c.drawString(x_tarih, y, date_str)
            c.drawString(x_bas, y, bas)
            c.drawString(x_bit, y, bit)
            c.drawString(x_mes, y, mes)
            c.line(40, y - 2, width - 40, y - 2)
            y -= 18

        c.showPage()

    c.save()
    messagebox.showinfo("Rapor Oluşturuldu",
                        f"PDF masaüstüne kaydedildi:\n{pdf_path}")



# ---------- AYARLAR / POPUP KONUMU ----------

def center_window(win, width, height):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - width) // 2
    y = (sh - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")

def open_popup_position_window():
    win = tk.Toplevel(root)
    win.title("BİLDİRİM KONUMU")
    win.configure(bg="#111111")

    width, height = 650, 380
    center_window(win, width, height)

    info = tk.Label(win, text="Aşağıdaki çerçeveyi sürükleyip konumlandırın ve Kaydet'e basın.",
                    bg="#111111", fg="white", font=("Segoe UI", 11, "bold"))
    info.pack(pady=(12, 8))

    # Örnek popup çerçevesi (gerçek popup ile aynı boyut/renk)
    sample = tk.Toplevel(win)
    sample.title("ÖRNEK BİLDİRİM")
    sample.overrideredirect(True)
    sample.attributes("-topmost", True)

    p_w, p_h = 600, 260
    sw = sample.winfo_screenwidth()
    sh = sample.winfo_screenheight()
    x = (sw - p_w) // 2
    y = (sh - p_h) // 2
    sample.geometry(f"{p_w}x{p_h}+{x}+{y}")

    frame = tk.Frame(sample, bg="#222222")
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="KART OKUMA", bg="#444444", fg="white",
                     font=("Segoe UI", 18, "bold"))
    title.pack(fill="x")

    body = tk.Label(frame, text="Bu bir örnek bildirimin çerçevesidir.", bg="#222222",
                    fg="white", font=("Segoe UI", 14))
    body.pack(expand=True)

    # Drag logic
    drag_data = {"x": 0, "y": 0}

    def on_press(event):
        drag_data["x"] = event.x
        drag_data["y"] = event.y

    def on_motion(event):
        nx = sample.winfo_x() + (event.x - drag_data["x"])
        ny = sample.winfo_y() + (event.y - drag_data["y"])
        sample.geometry(f"+{nx}+{ny}")

    frame.bind("<ButtonPress-1>", on_press)
    frame.bind("<B1-Motion>", on_motion)
    title.bind("<ButtonPress-1>", on_press)
    title.bind("<B1-Motion>", on_motion)

    btn_frame = tk.Frame(win, bg="#111111")
    btn_frame.pack(pady=12)

    def save_pos():
        set_saved_popup_position(sample.winfo_x(), sample.winfo_y())
        try:
            sample.destroy()
        except Exception:
            pass
        win.destroy()
        messagebox.showinfo("Kaydedildi", "Bildirim konumu kaydedildi.")

    def cancel():
        try:
            sample.destroy()
        except Exception:
            pass
        win.destroy()

    btn_save = create_dark_button(btn_frame, "Kaydet", save_pos)
    btn_save.grid(row=0, column=0, padx=6)
    btn_cancel = create_dark_button(btn_frame, "Vazgeç", cancel)
    btn_cancel.grid(row=0, column=1, padx=6)


def change_management_password():
    old_pwd = simpledialog.askstring("Şifre Değiştir", "Mevcut şifre (4 hane):", show="*")
    if old_pwd is None:
        return
    if old_pwd != get_management_password():
        messagebox.showerror("Hata", "Mevcut şifre yanlış.")
        return

    new_pwd = simpledialog.askstring("Şifre Değiştir", "Yeni şifre (4 hane):", show="*")
    if new_pwd is None:
        return
    new_pwd2 = simpledialog.askstring("Şifre Değiştir", "Yeni şifre (tekrar):", show="*")
    if new_pwd2 is None:
        return
    if new_pwd != new_pwd2:
        messagebox.showerror("Hata", "Yeni şifreler eşleşmiyor.")
        return
    if not (new_pwd.isdigit() and len(new_pwd) == 4):
        messagebox.showerror("Hata", "Şifre 4 haneli rakam olmalıdır.")
        return
    set_management_password(new_pwd)
    messagebox.showinfo("Başarılı", "Şifre güncellendi.")

def open_panel_mode_window():
    win = tk.Toplevel(root)
    win.title("PANEL MODU")
    win.configure(bg="#111111")
    center_window(win, 420, 240)
    win.grab_set()

    tk.Label(
        win,
        text="Panel modunu seçin:",
        bg="#111111",
        fg="white",
        font=("Segoe UI", 11, "bold")
    ).pack(pady=(16, 10))

    def set_main():
        set_panel_mode("main")
        win.destroy()
        restart_app()

    def set_management():
        set_panel_mode("management")
        win.destroy()
        restart_app()

    btn_main = create_dark_button(win, "Ana Panel", set_main)
    btn_main.pack(padx=12, pady=8, fill="x")

    btn_mgmt = create_dark_button(win, "Yönetim Panel", set_management)
    btn_mgmt.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", win.destroy)
    btn_close.pack(padx=12, pady=(8, 16), fill="x")


def open_ayarlar_menu():
    win = tk.Toplevel(root)
    win.title("AYARLAR")
    win.configure(bg="#111111")
    center_window(win, 420, 280)
    win.grab_set()

    btn_panel = create_dark_button(win, "Panel Modu", lambda: [open_panel_mode_window(), win.destroy()])
    btn_panel.pack(padx=12, pady=(18, 8), fill="x")

    btn_pos = create_dark_button(win, "Bildirim Konumu", lambda: [open_popup_position_window(), win.destroy()])
    btn_pos.pack(padx=12, pady=8, fill="x")

    btn_pwd = create_dark_button(win, "Şifre Değiştir", lambda: [change_management_password(), win.destroy()])
    btn_pwd.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", win.destroy)
    btn_close.pack(padx=12, pady=(8, 18), fill="x")


# ---------- PERSONEL ----------

KAN_GRUPLARI = ["Bilinmiyor", "A Rh+", "A Rh-", "B Rh+", "B Rh-", "AB Rh+", "AB Rh-", "0 Rh+", "0 Rh-"]
DURUM_SECENEKLERI = ["Kart Aktif", "Kart Devredildi"]

def read_all_kisiler_rows():
    rows = []
    if not os.path.exists(KISILER_PATH):
        return rows
    with open(KISILER_PATH, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(normalize_kisi_row(r))
    return rows

def write_all_kisiler_rows(rows):
    save_kisiler_rows(rows)
    load_kisiler()

def open_personel_menu():
    win = tk.Toplevel(root)
    win.title("PERSONEL")
    win.configure(bg="#111111")
    center_window(win, 460, 280)
    win.grab_set()

    btn_new = create_dark_button(win, "Yeni Personel Tanımla", lambda: [open_personel_new_window(), win.destroy()])
    btn_new.pack(padx=12, pady=(18, 8), fill="x")

    btn_edit = create_dark_button(win, "Mevcutları Düzenle", lambda: [open_personel_edit_menu(), win.destroy()])
    btn_edit.pack(padx=12, pady=8, fill="x")

    btn_list = create_dark_button(win, "Liste Al", lambda: [open_personel_list_menu(), win.destroy()])
    btn_list.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", win.destroy)
    btn_close.pack(padx=12, pady=(8, 18), fill="x")


def open_personel_new_window():
    win = tk.Toplevel(root)
    win.title("YENİ PERSONEL TANIMLA")
    win.configure(bg="#111111")
    center_window(win, 560, 620)
    win.grab_set()

    frm = tk.Frame(win, bg="#111111")
    frm.pack(fill="both", expand=True, padx=12, pady=12)

    def add_row(r, label, widget):
        tk.Label(frm, text=label, bg="#111111", fg="white", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, sticky="w", pady=4)
        widget.grid(row=r, column=1, sticky="ew", pady=4)
        frm.grid_columnconfigure(1, weight=1)

    e_card = tk.Entry(frm, font=("Segoe UI", 11))
    e_name = tk.Entry(frm, font=("Segoe UI", 11))
    e_tc = tk.Entry(frm, font=("Segoe UI", 11))
    e_cep = tk.Entry(frm, font=("Segoe UI", 11))
    e_ise = tk.Entry(frm, font=("Segoe UI", 11))

    durum_var = tk.StringVar(value="Kart Aktif")
    cb_durum = ttk.Combobox(frm, textvariable=durum_var, values=DURUM_SECENEKLERI, state="readonly", font=("Segoe UI", 11))

    e_adres = tk.Entry(frm, font=("Segoe UI", 11))
    kan_var = tk.StringVar(value="0")
    cb_kan = ttk.Combobox(frm, textvariable=kan_var, values=KAN_GRUPLARI, state="readonly", font=("Segoe UI", 11))
    e_acil_ad = tk.Entry(frm, font=("Segoe UI", 11))
    e_acil_no = tk.Entry(frm, font=("Segoe UI", 11))
    e_gorev = tk.Entry(frm, font=("Segoe UI", 11))
    e_not = tk.Entry(frm, font=("Segoe UI", 11))

    add_row(0, "Kart Numarası", e_card)
    add_row(1, "Ad Soyad", e_name)
    add_row(2, "TC Numarası", e_tc)
    add_row(3, "Cep Telefonu Numarası", e_cep)
    add_row(4, "İşe Başlama Tarihi (gg.aa.yyyy)", e_ise)
    add_row(5, "Durum", cb_durum)
    add_row(6, "Adres", e_adres)
    add_row(7, "Kan Grubu", cb_kan)
    add_row(8, "Acil Durum Kişi Adı", e_acil_ad)
    add_row(9, "Acil Durum Numarası", e_acil_no)
    add_row(10, "Görev / Pozisyon", e_gorev)
    add_row(11, "Notlar", e_not)

    def save_new():
        card_id = e_card.get().strip()
        ad = e_name.get().strip()
        tc = e_tc.get().strip()
        cep = e_cep.get().strip()
        ise = e_ise.get().strip()

        if not card_id or not ad or not tc or not cep or not ise:
            messagebox.showerror("Hata", "Zorunlu alanlar: Kart, Ad Soyad, TC, Cep, İşe Başlama.")
            return
        # tarih format kontrol
        try:
            parse_dotted_date(ise)
        except Exception:
            messagebox.showerror("Hata", "İşe Başlama Tarihi gg.aa.yyyy formatında olmalı.")
            return

        rows = read_all_kisiler_rows()
        for r in rows:
            if r.get("Kart Numarası") == card_id:
                messagebox.showerror("Hata", "Bu kart numarası zaten kayıtlı.")
                return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = {
            "Kart Numarası": card_id,
            "Ad Soyad": ad,
            "TC Numarası": tc,
            "Cep Telefonu Numarası": cep,
            "İşe Başlama Tarihi": ise,
            "Durum": durum_var.get(),
            "Adres": e_adres.get().strip() or "0",
            "Kan Grubu": kan_var.get(),
            "Acil Durum Kişi Adı": e_acil_ad.get().strip() or "0",
            "Acil Durum Numarası": e_acil_no.get().strip() or "0",
            "Görev / Pozisyon": e_gorev.get().strip() or "0",
            "Notlar": e_not.get().strip() or "0",
            "Kayıt Tarihi": now_str,
            "Senkronize": "0"
        }
        rows.append(normalize_kisi_row(new_row))
        write_all_kisiler_rows(rows)
        win.destroy()
        messagebox.showinfo("Başarılı", "Personel kaydedildi.")

    btns = tk.Frame(win, bg="#111111")
    btns.pack(pady=(0, 12))
    create_dark_button(btns, "Kaydet", save_new).grid(row=0, column=0, padx=6)
    create_dark_button(btns, "Vazgeç", win.destroy).grid(row=0, column=1, padx=6)


def open_personel_edit_menu():
    win = tk.Toplevel(root)
    win.title("MEVCUTLARI DÜZENLE")
    win.configure(bg="#111111")
    center_window(win, 460, 240)
    win.grab_set()

    btn_a = create_dark_button(win, "Aktif Personel", lambda: [open_personel_edit_list(True), win.destroy()])
    btn_a.pack(padx=12, pady=(18, 8), fill="x")

    btn_p = create_dark_button(win, "Pasif Personel", lambda: [open_personel_edit_list(False), win.destroy()])
    btn_p.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", win.destroy)
    btn_close.pack(padx=12, pady=(8, 18), fill="x")


def open_personel_edit_list(is_active: bool):
    win = tk.Toplevel(root)
    win.title("AKTİF PERSONEL" if is_active else "PASİF PERSONEL")
    win.configure(bg="#111111")
    center_window(win, 560, 220)
    win.grab_set()

    rows = read_all_kisiler_rows()
    filtered = []
    for r in rows:
        if is_active and r.get("Durum") == "Kart Aktif":
            filtered.append(r)
        if (not is_active) and r.get("Durum") == "Kart Devredildi":
            filtered.append(r)

    filtered = sorted(filtered, key=lambda x: x.get("Ad Soyad",""))

    tk.Label(win, text="Personel Seç:", bg="#111111", fg="white", font=("Segoe UI", 11, "bold")).pack(pady=(14,6))
    choices = [f'{r.get("Ad Soyad","")} | {r.get("Kart Numarası","")}' for r in filtered]
    sel = tk.StringVar(value=choices[0] if choices else "")
    cb = ttk.Combobox(win, textvariable=sel, values=choices, state="readonly", font=("Segoe UI", 11))
    cb.pack(padx=12, fill="x")

    def open_selected():
        if not sel.get():
            return
        idx = choices.index(sel.get())
        open_personel_edit_form(filtered[idx], is_active)
        win.destroy()

    btns = tk.Frame(win, bg="#111111")
    btns.pack(pady=14)
    create_dark_button(btns, "Aç", open_selected).grid(row=0, column=0, padx=6)
    create_dark_button(btns, "Kapat", win.destroy).grid(row=0, column=1, padx=6)


def open_personel_edit_form(row_data: dict, is_active: bool):
    win = tk.Toplevel(root)
    win.title("PERSONEL DÜZENLE")
    win.configure(bg="#111111")
    center_window(win, 560, 620)
    win.grab_set()

    frm = tk.Frame(win, bg="#111111")
    frm.pack(fill="both", expand=True, padx=12, pady=12)

    def add_row(r, label, widget, disabled=False):
        tk.Label(frm, text=label, bg="#111111", fg="white", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, sticky="w", pady=4)
        widget.grid(row=r, column=1, sticky="ew", pady=4)
        if disabled:
            try:
                widget.configure(state="disabled")
            except Exception:
                pass
        frm.grid_columnconfigure(1, weight=1)

    e_card = tk.Entry(frm, font=("Segoe UI", 11))
    e_card.insert(0, row_data.get("Kart Numarası",""))
    e_card.configure(state="disabled")

    e_name = tk.Entry(frm, font=("Segoe UI", 11)); e_name.insert(0, row_data.get("Ad Soyad",""))
    e_tc = tk.Entry(frm, font=("Segoe UI", 11)); e_tc.insert(0, row_data.get("TC Numarası",""))
    e_cep = tk.Entry(frm, font=("Segoe UI", 11)); e_cep.insert(0, row_data.get("Cep Telefonu Numarası",""))
    e_ise = tk.Entry(frm, font=("Segoe UI", 11)); e_ise.insert(0, row_data.get("İşe Başlama Tarihi",""))

    durum_var = tk.StringVar(value=row_data.get("Durum","Kart Aktif"))
    cb_durum = ttk.Combobox(frm, textvariable=durum_var, values=DURUM_SECENEKLERI, state="readonly", font=("Segoe UI", 11))

    e_adres = tk.Entry(frm, font=("Segoe UI", 11)); e_adres.insert(0, row_data.get("Adres","0"))
    kan_var = tk.StringVar(value="0")
    cb_kan = ttk.Combobox(frm, textvariable=kan_var, values=KAN_GRUPLARI, state="readonly", font=("Segoe UI", 11))
    e_acil_ad = tk.Entry(frm, font=("Segoe UI", 11)); e_acil_ad.insert(0, row_data.get("Acil Durum Kişi Adı","0"))
    e_acil_no = tk.Entry(frm, font=("Segoe UI", 11)); e_acil_no.insert(0, row_data.get("Acil Durum Numarası","0"))
    e_gorev = tk.Entry(frm, font=("Segoe UI", 11)); e_gorev.insert(0, row_data.get("Görev / Pozisyon","0"))
    e_not = tk.Entry(frm, font=("Segoe UI", 11)); e_not.insert(0, row_data.get("Notlar","0"))

    add_row(0, "Kart Numarası", e_card, disabled=True)
    add_row(1, "Ad Soyad", e_name)
    add_row(2, "TC Numarası", e_tc)
    add_row(3, "Cep Telefonu Numarası", e_cep)
    add_row(4, "İşe Başlama Tarihi (gg.aa.yyyy)", e_ise)
    add_row(5, "Durum", cb_durum, disabled=(not is_active))
    add_row(6, "Adres", e_adres)
    add_row(7, "Kan Grubu", cb_kan)
    add_row(8, "Acil Durum Kişi Adı", e_acil_ad)
    add_row(9, "Acil Durum Numarası", e_acil_no)
    add_row(10, "Görev / Pozisyon", e_gorev)
    add_row(11, "Notlar", e_not)

    def save_edit():
        # pasif personelde durum kilitli: Kart Devredildi kalır
        if is_active:
            # aktif -> pasif yapılabilir
            pass
        else:
            durum_var.set("Kart Devredildi")

        try:
            parse_dotted_date(e_ise.get().strip())
        except Exception:
            messagebox.showerror("Hata", "İşe Başlama Tarihi gg.aa.yyyy formatında olmalı.")
            return

        rows = read_all_kisiler_rows()
        for r in rows:
            if r.get("Kart Numarası") == row_data.get("Kart Numarası"):
                r["Ad Soyad"] = e_name.get().strip() or r["Ad Soyad"]
                r["TC Numarası"] = e_tc.get().strip() or r["TC Numarası"]
                r["Cep Telefonu Numarası"] = e_cep.get().strip() or r["Cep Telefonu Numarası"]
                r["İşe Başlama Tarihi"] = e_ise.get().strip() or r["İşe Başlama Tarihi"]
                r["Durum"] = durum_var.get()
                r["Adres"] = e_adres.get().strip() or "0"
                r["Kan Grubu"] = kan_var.get()
                r["Acil Durum Kişi Adı"] = e_acil_ad.get().strip() or "0"
                r["Acil Durum Numarası"] = e_acil_no.get().strip() or "0"
                r["Görev / Pozisyon"] = e_gorev.get().strip() or "0"
                r["Notlar"] = e_not.get().strip() or "0"
                break
        write_all_kisiler_rows(rows)
        win.destroy()
        messagebox.showinfo("Başarılı", "Personel güncellendi.")

    btns = tk.Frame(win, bg="#111111")
    btns.pack(pady=(0, 12))
    create_dark_button(btns, "Kaydet", save_edit).grid(row=0, column=0, padx=6)
    create_dark_button(btns, "Vazgeç", win.destroy).grid(row=0, column=1, padx=6)


# ---------- PERSONEL LİSTE AL (PDF) ----------

def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")

def generate_personel_list_pdf(is_active: bool):
    font_name = "Helvetica"
    try:
        font_path = os.path.join(BASE_DIR, "DejaVuSans.ttf")
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont("DejaVuSans", font_path))
            font_name = "DejaVuSans"
    except Exception:
        font_name = "Helvetica"

    today = date.today()
    today_str = today.strftime("%Y-%m-%d")

    suffix = "Aktif Personel" if is_active else "Pasif Personel"
    filename = f"{today_str} Personel Listesi {suffix}.pdf"
    out_path = os.path.join(get_desktop_path(), filename)

    rows = read_all_kisiler_rows()
    filtered = [r for r in rows if (r.get("Durum") == ("Kart Aktif" if is_active else "Kart Devredildi"))]
    filtered = sorted(filtered, key=lambda x: x.get("Ad Soyad",""))

    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4

    # First page summary
    c.setFont(font_name, 18)
    c.drawString(50, height - 60, sanitize_turkish(f"{suffix} - Ozet Liste ({today_str})"))

    # (2) Başlıklarda Türkçe karakterleri düzelt (sanitize)
    headers = [
        sanitize_turkish("Ad Soyad"),
        sanitize_turkish("TC"),
        sanitize_turkish("Cep"),
        sanitize_turkish("Gorev / Pozisyon"),
        sanitize_turkish("Ise Baslama"),
    ]
    col_x = [50, 210, 300, 390, 520]
    y = height - 95
    c.setFont(font_name, 10)
    for i, h in enumerate(headers):
        c.drawString(col_x[i], y, h)
    y -= 10
    c.line(50, y, width - 50, y)
    y -= 15

    c.setFont(font_name, 9)
    for r in filtered:
        if y < 80:
            c.showPage()
            y = height - 60
        c.drawString(col_x[0], y, sanitize_turkish(str(r.get("Ad Soyad",""))[:30]))
        c.drawString(col_x[1], y, sanitize_turkish(str(r.get("TC Numarası",""))[:11]))
        c.drawString(col_x[2], y, sanitize_turkish(str(r.get("Cep Telefonu Numarası",""))[:15]))
        c.drawString(col_x[3], y, sanitize_turkish(str(r.get("Görev / Pozisyon",""))[:18]))
        c.drawString(col_x[4], y, sanitize_turkish(str(r.get("İşe Başlama Tarihi",""))[:10]))
        y -= 14

    # Person pages
    for r in filtered:
        c.showPage()
        c.setFont(font_name, 16)
        c.drawString(50, height - 60, sanitize_turkish(r.get("Ad Soyad","")))

        c.setFont(font_name, 11)
        y = height - 90

        # 13 satır (Senkronize hariç)
        keys_13 = [k for k in KISILER_FIELDNAMES if k != "Senkronize"]
        for k in keys_13:
            v = str(r.get(k,"0"))
            c.drawString(50, y, f"{sanitize_turkish(k)}: {sanitize_turkish(v)}")
            y -= 18
            c.line(50, y, width - 50, y)  # satırlar arası çizgi
            y -= 12
            if y < 80:
                c.showPage()
                c.setFont(font_name, 11)
                y = height - 60

    c.save()
    return out_path


def open_personel_list_menu():
    win = tk.Toplevel(root)
    win.title("LİSTE AL")
    win.configure(bg="#111111")
    center_window(win, 460, 240)
    win.grab_set()

    def do_active():
        path = generate_personel_list_pdf(True)
        messagebox.showinfo("Oluşturuldu", f"PDF oluşturuldu:\n{path}")
        win.destroy()

    def do_passive():
        path = generate_personel_list_pdf(False)
        messagebox.showinfo("Oluşturuldu", f"PDF oluşturuldu:\n{path}")
        win.destroy()

    btn_a = create_dark_button(win, "Aktif Personel", do_active)
    btn_a.pack(padx=12, pady=(18, 8), fill="x")

    btn_p = create_dark_button(win, "Pasif Personel", do_passive)
    btn_p.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", win.destroy)
    btn_close.pack(padx=12, pady=(8, 18), fill="x")


# ---------- REPORT DIALOG ----------

def open_report_dialog():
    win = tk.Toplevel(root)
    win.title("Rapor Al")
    win.grab_set()
    style_dark_window(win)

    lbl = tk.Label(win, text="Hangi raporu almak istiyorsunuz?",
                   bg="#222222", fg="white",
                   font=("Segoe UI", 12, "bold"))
    lbl.pack(padx=10, pady=10)

    btn_current = create_dark_button(win, "Güncel Ay Raporu",
                                     lambda: [generate_report(True), win.destroy()])
    btn_current.pack(padx=10, pady=5, fill="x")

    btn_prev = create_dark_button(win, "Geçmiş Ay Raporu",
                                  lambda: [generate_report(False), win.destroy()])
    btn_prev.pack(padx=10, pady=5, fill="x")


# ---------- MANAGEMENT MENU (YENİ AÇILIŞ MEKANİZMASI) ----------

def open_management_menu():
    global _management_menu_win  # TEK DEĞİŞİKLİK

    # TEK DEĞİŞİKLİK: Zaten açıksa yeni pencere açma, mevcutu öne getir.
    try:
        if _management_menu_win is not None and _management_menu_win.winfo_exists():
            _management_menu_win.deiconify()
            _management_menu_win.lift()
            _management_menu_win.focus_force()
            return
    except Exception:
        pass

    win = tk.Toplevel(root)
    _management_menu_win = win  # TEK DEĞİŞİKLİK

    win.title("YÖNETİM")
    win.configure(bg="#111111")
    center_window(win, 460, 360)
    win.grab_set()

    btn_manual = create_dark_button(win, "Manuel Kayıt", lambda: [open_manual_log_window(), win.destroy()])
    btn_manual.pack(padx=12, pady=(18, 8), fill="x")

    btn_new = create_dark_button(win, "Yeni Kart Tanımla", lambda: [open_new_card_window(), win.destroy()])
    btn_new.pack(padx=12, pady=8, fill="x")

    btn_report = create_dark_button(win, "Rapor Al", lambda: [open_report_dialog(), win.destroy()])
    btn_report.pack(padx=12, pady=8, fill="x")

    btn_personel = create_dark_button(win, "Personel", lambda: [open_personel_menu(), win.destroy()])
    btn_personel.pack(padx=12, pady=8, fill="x")

    btn_ayarlar = create_dark_button(win, "Ayarlar", lambda: [open_ayarlar_menu(), win.destroy()])
    btn_ayarlar.pack(padx=12, pady=8, fill="x")

    btn_close = create_dark_button(win, "Kapat", request_app_exit if get_panel_mode()=="management" else win.destroy)
    btn_close.pack(padx=12, pady=(8, 18), fill="x")


def open_management_menu_password():
    try:
        pwd = simpledialog.askstring("Şifre", "Yönetim şifresini girin:", show="*", parent=root)
    except Exception:
        pwd = None
    if pwd != get_management_password():
        return
    open_management_menu()


# ---------- MAIN WINDOW ----------

def show_main_window():
    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.after(100, lambda: root.attributes("-topmost", False))


def on_close_window():
    root.withdraw()


def configure_tk_turkish_support(app: tk.Tk):
    try:
        app.tk.call("encoding", "system", "utf-8")
    except tk.TclError:
        pass

    default_family = "Segoe UI"
    app.option_add("*Font", (default_family, 10))
    for font_name in (
        "TkDefaultFont",
        "TkTextFont",
        "TkFixedFont",
        "TkMenuFont",
        "TkHeadingFont",
        "TkCaptionFont",
        "TkSmallCaptionFont",
        "TkIconFont",
        "TkTooltipFont",
    ):
        try:
            tkfont.nametofont(font_name).configure(family=default_family)
        except tk.TclError:
            pass


def build_main_window():
    global root, header_label, date_label, tree, total_label
    global _btn_new, _btn_manual, _btn_report, _btn_personel, _btn_ayarlar, _btn_back, _btn_yonetim

    root.title("Lavoni Kayıt Sistemi – Günlük Kayıtlar")
    root.geometry("1100x600")
    root.configure(bg="#111111")
    root.protocol("WM_DELETE_WINDOW", on_close_window)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview",
                    background="#222222",
                    foreground="white",
                    fieldbackground="#222222",
                    rowheight=24)
    style.configure("Treeview.Heading",
                    background="#333333",
                    foreground="white",
                    font=("Segoe UI", 10, "bold"))
    style.map("Treeview", background=[("selected", "#444466")])

    main_frame = tk.Frame(root, bg="#111111")
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    header_label_local = tk.Label(main_frame, text="Lavoni - Günlük Kayıtlar",
                                  font=("Segoe UI", 18, "bold"),
                                  bg="#111111", fg="white")
    header_label_local.pack()
    header_label = header_label_local

    date_label_local = tk.Label(main_frame,
                                text=format_turkish_date_long(current_ui_date),
                                font=("Segoe UI", 20, "bold"),
                                bg="#111111", fg="#4aa3ff")
    date_label_local.pack(pady=(0, 10))
    global date_label
    date_label = date_label_local

    columns = ("Saat", "Ad Soyad", "Kart Numarası")
    tree_local = ttk.Treeview(main_frame, columns=columns,
                              show="headings", height=15, style="Treeview")
    for col in columns:
        tree_local.heading(col, text=col)
        if col == "Saat":
            tree_local.column(col, width=80, anchor="center")
        else:
            tree_local.column(col, width=260, anchor="center")
    tree_local.pack(fill="both", expand=True)
    tree_local._lavoni_count = 0
    global tree
    tree = tree_local

    total_label_local = tk.Label(main_frame,
                                 text="Toplam Okutma Sayısı: 0",
                                 font=("Segoe UI", 12, "bold"),
                                 bg="#111111", fg="white")
    total_label_local.pack(pady=5)
    global total_label
    total_label = total_label_local

    button_frame = tk.Frame(main_frame, bg="#111111")
    button_frame.pack(pady=10)

    # Yönetim artık menü penceresi açar (şifreli)
    _btn_yonetim = create_dark_button(button_frame, "Yönetim", open_management_menu_password)
    _btn_yonetim.grid(row=0, column=0, padx=5)

    # Aşağıdakiler içerik olarak korunuyor; sadece ilk açılışta gizli kalacak
    _btn_manual = create_dark_button(button_frame, "Manuel Kayıt", open_manual_log_window)
    _btn_manual.grid(row=0, column=1, padx=5)

    _btn_new = create_dark_button(button_frame, "Yeni Kart Tanımla", open_new_card_window)
    _btn_new.grid(row=0, column=2, padx=5)

    _btn_report = create_dark_button(button_frame, "Rapor Al", open_report_dialog)
    _btn_report.grid(row=0, column=3, padx=5)

    _btn_personel = create_dark_button(button_frame, "Personel", open_personel_menu)
    _btn_personel.grid(row=0, column=4, padx=5)

    _btn_ayarlar = create_dark_button(button_frame, "Ayarlar", open_ayarlar_menu)
    _btn_ayarlar.grid(row=0, column=5, padx=5)

    _btn_back = create_dark_button(button_frame, "Geri", lambda: None)
    _btn_back.grid(row=0, column=6, padx=5)

    # ilk açılışta SADECE Yönetim görünsün
    _btn_manual.grid_remove()
    _btn_new.grid_remove()
    _btn_report.grid_remove()
    _btn_personel.grid_remove()
    _btn_ayarlar.grid_remove()
    _btn_back.grid_remove()

    reload_today_table()


# ---------- MAIN ----------

def main():
    global tray_thread, tray_started, root

    ensure_single_instance()
    ensure_directories_and_files()
    load_kisiler()

    root = tk.Tk()
    configure_tk_turkish_support(root)
    build_main_window()

    # Tray sadece bir kez başlasın
    if not tray_started:
        tray_thread = threading.Thread(target=tray_run, daemon=True)
        tray_thread.start()
        tray_started = True

    mode = get_panel_mode()

    if mode == "management":
        # Yönetim Panel: sadece yönetim paneli açık kalsın, kart okuma fonksiyonu iptal
        root.withdraw()
        root.after(0, open_management_menu)
    else:
        root.withdraw()
        start_keyboard_listener()
        root.after(100, process_card_from_queue)

    root.mainloop()


if __name__ == "__main__":
    main()
