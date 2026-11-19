# -*- coding: utf-8 -*-
"""
Mobile Legends Meta Analiz Araci
================================
mobilelegends.com/rank sayfasindaki kahraman istatistiklerini (secilme, kazanma
ve yasaklanma oranlari) Selenium ile ceker, her kahraman icin bir "Hero Power"
puani hesaplar ve kahramanlari tier'lara (SSS - F) ayirir.

Sekmeler:
  1) "Veri Cek"            -> Zaman araligi ve rank secilerek veri cekme
  2) "Meta Tier List"      -> Rol / lane / tier filtreleme, siralama, Excel'e aktarma
  3) "Lane x Tier Tablosu" -> Satirlar = tier, sutunlar = lane; her hucrede o tier
                              ve lane'e giren kahramanlarin portreleri

Hero Power formulu:  (1.4 * pick_rate * win_rate) + (ban_rate * 3)
Tier'lar Hero Power siralamasindaki yuzdelik dilime gore atanir (bkz. TIER_CUTOFFS).

Kurulum:
    pip install -r requirements.txt

Calistirma:
    python mobile_legends_tier_list.py
"""

import sys
import os
import re
import json
import time
import urllib.request

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QTextEdit, QTableWidget, QTableWidgetItem,
    QLineEdit, QHeaderView, QMessageBox, QFileDialog, QStyledItemDelegate,
    QAbstractItemView, QProgressBar, QGroupBox, QFormLayout, QDialog,
    QDialogButtonBox, QScrollArea, QGridLayout,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import (
    QColor, QFont, QStandardItemModel, QStandardItem, QPixmap, QPainter,
    QPainterPath, QIcon,
)

# ------------------------------------------------------------------
# Selenium (sadece "Veri Cek" sekmesinde kullanilir)
# ------------------------------------------------------------------
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ------------------------------------------------------------------
# Excel export
# ------------------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ====================================================================
# SABITLER
# ====================================================================
URL = "https://www.mobilelegends.com/rank"
LOAD_WAIT = 30
WINDOW_SIZE = (900, 650)

X_PARENT = "//div[@data-node='2684827']"
X_ROWS = "//div[@data-node='2684827']//div[@data-node='2684916']"
X_ROWS_CLASS = "//div[@data-node='2684827']//div[contains(@class, 'mt-list-item')]"
X_ROWS_GENERIC = "//div[@data-node='2684827']//div[.//div[@data-node='2693412']]"

X_NAME = ".//div[@data-node='2693412']//span"
X_PICK = ".//div[@data-node='2684925']//span"
X_WIN = ".//div[@data-node='2684926']//span"
X_BAN = ".//div[@data-node='2687632']//span"
# Satirdaki kahraman portresinin <img> etiketi
X_IMG = ".//div[@data-node='2684911']//img"

X_TIME_DROPDOWN = "//div[@data-node='2684829']"
X_RANK_DROPDOWN = "//div[@data-node='2684880']"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
HERO_ROLES_FILE = os.path.join(OUTPUT_DIR, "hero_roles.json")
RAW_DATA_FILE = os.path.join(OUTPUT_DIR, "mobile_legends_heroes.txt")
TIERLIST_FILE = os.path.join(OUTPUT_DIR, "heroes_tier_list.txt")
LAST_QUERY_FILE = os.path.join(OUTPUT_DIR, "last_query.json")
# Kahraman portreleri ve rank logolari script ile ayni klasordeki
# IMAGES/HEROES ve IMAGES/RANKS altinda tutulur. Bu klasorlere elle resim
# eklenebilir; program klasorleri otomatik tarar.
IMAGES_DIR = os.path.join(SCRIPT_DIR, "IMAGES", "HEROES")
RANKS_DIR = os.path.join(SCRIPT_DIR, "IMAGES", "RANKS")

TIME_OPTIONS = ["Past 1 day", "Past 3 days", "Past 7 days", "Past 15 days", "Past 30 days"]
RANK_OPTIONS = ["ALL", "Epic", "Legend", "Mythic", "Mythical Honor", "Mythical Glory+"]

ROLES_LIST = ["TANK", "FIGHTER", "ASSASSIN", "MAGE", "MARKSMAN", "SUPPORT"]
LANES_LIST = ["GOLD", "EXP", "MID", "ROAM", "JUNGLE"]

TIER_ORDER = ["SSS", "SS", "S", "A", "B", "C", "F"]
TIER_COLORS = {
    "SSS": "#e74c3c",
    "SS": "#e67e22",
    "S": "#f1c40f",
    "A": "#2ecc71",
    "B": "#3498db",
    "C": "#95a5a6",
    "F": "#7f8c8d",
}
TIER_RANK = {t: i for i, t in enumerate(TIER_ORDER)}
TIER_CUTOFFS = [("SSS", 0.03), ("SS", 0.13), ("S", 0.33), ("A", 0.53), ("B", 0.73), ("C", 0.90), ("F", 1.01)]

LANE_LABELS = {"GOLD": "GOLD LANE", "EXP": "EXP LANE", "MID": "MID LANE", "ROAM": "ROAMER", "JUNGLE": "JUNGLER"}

# ------------------------------------------------------------------
# Bilinen kahramanlarin varsayilan Rol / Lane atamalari (COKLU destekli).
# Format: "ISIM": (["ROL1", "ROL2", ...], ["LANE1", "LANE2", ...])
# Yeni cikan / eslesmeyen kahramanlar veri cekilince otomatik tespit edilip
# "Yeni Kahraman Bulundu" penceresiyle kullaniciya sorulur; buraya elle eklemek gerekmez.
# ------------------------------------------------------------------
_SINGLE = {
    "MIYA": ("MARKSMAN", "GOLD"), "BALMOND": ("FIGHTER", "JUNGLE"), "SABER": ("ASSASSIN", "JUNGLE"),
    "ALICE": ("MAGE", "MID"), "NANA": ("MAGE", "MID"), "TIGREAL": ("TANK", "ROAM"),
    "ALUCARD": ("FIGHTER", "JUNGLE"), "KARINA": ("ASSASSIN", "JUNGLE"), "AKAI": ("TANK", "ROAM"),
    "FRANCO": ("TANK", "ROAM"), "BRUNO": ("MARKSMAN", "GOLD"), "CLINT": ("MARKSMAN", "GOLD"),
    "RAFAELA": ("SUPPORT", "ROAM"), "EUDORA": ("MAGE", "MID"), "ZILONG": ("FIGHTER", "JUNGLE"),
    "FANNY": ("ASSASSIN", "JUNGLE"), "LAYLA": ("MARKSMAN", "GOLD"), "MINOTAUR": ("TANK", "ROAM"),
    "LOLITA": ("SUPPORT", "ROAM"), "HAYABUSA": ("ASSASSIN", "JUNGLE"), "FREYA": ("FIGHTER", "JUNGLE"),
    "GORD": ("MAGE", "MID"), "NATALIA": ("ASSASSIN", "JUNGLE"), "KAGURA": ("MAGE", "MID"),
    "CHOU": ("FIGHTER", "EXP"), "SUN": ("FIGHTER", "JUNGLE"), "ALPHA": ("FIGHTER", "JUNGLE"),
    "RUBY": ("FIGHTER", "JUNGLE"), "YI SUN-SHIN": ("ASSASSIN", "JUNGLE"), "MOSKOV": ("MARKSMAN", "GOLD"),
    "JOHNSON": ("TANK", "ROAM"), "CYCLOPS": ("MAGE", "MID"), "ESTES": ("SUPPORT", "ROAM"),
    "HILDA": ("FIGHTER", "JUNGLE"), "AURORA": ("MAGE", "MID"), "LAPU-LAPU": ("FIGHTER", "EXP"),
    "VEXANA": ("MAGE", "MID"), "KARRIE": ("MARKSMAN", "GOLD"), "GATOTKACA": ("TANK", "ROAM"),
    "IRITHEL": ("MARKSMAN", "GOLD"), "GROCK": ("TANK", "EXP"), "ARGUS": ("FIGHTER", "EXP"),
    "ODETTE": ("MAGE", "MID"), "LANCELOT": ("ASSASSIN", "JUNGLE"), "DIGGIE": ("SUPPORT", "ROAM"),
    "HYLOS": ("TANK", "ROAM"), "ZHASK": ("MAGE", "MID"), "HELCURT": ("ASSASSIN", "JUNGLE"),
    "PHARSA": ("MAGE", "MID"), "LESLEY": ("MARKSMAN", "GOLD"), "JAWHEAD": ("FIGHTER", "EXP"),
    "ANGELA": ("SUPPORT", "ROAM"), "GUSION": ("ASSASSIN", "JUNGLE"), "VALIR": ("MAGE", "MID"),
    "MARTIS": ("FIGHTER", "JUNGLE"), "URANUS": ("TANK", "EXP"), "HANABI": ("MARKSMAN", "GOLD"),
    "CHANG'E": ("MAGE", "MID"), "KAJA": ("SUPPORT", "ROAM"), "ALDOUS": ("FIGHTER", "JUNGLE"),
    "CLAUDE": ("MARKSMAN", "GOLD"), "VALE": ("MAGE", "MID"), "LEOMORD": ("FIGHTER", "JUNGLE"),
    "LUNOX": ("MAGE", "MID"), "HANZO": ("ASSASSIN", "JUNGLE"), "BELERICK": ("TANK", "EXP"),
    "KIMMY": ("MARKSMAN", "GOLD"), "THAMUZ": ("FIGHTER", "JUNGLE"), "HARITH": ("MAGE", "MID"),
    "MINSITTHAR": ("FIGHTER", "EXP"), "FARAMIS": ("SUPPORT", "ROAM"), "BADANG": ("FIGHTER", "EXP"),
    "KHUFRA": ("TANK", "ROAM"), "GRANGER": ("MARKSMAN", "GOLD"), "GUINEVERE": ("FIGHTER", "JUNGLE"),
    "TERIZLA": ("FIGHTER", "EXP"), "X.BORG": ("FIGHTER", "EXP"), "LING": ("ASSASSIN", "JUNGLE"),
    "DYRROTH": ("FIGHTER", "EXP"), "LYLIA": ("MAGE", "MID"), "BAXIA": ("TANK", "JUNGLE"),
    "MASHA": ("FIGHTER", "JUNGLE"), "WANWAN": ("MARKSMAN", "GOLD"), "SILVANNA": ("FIGHTER", "EXP"),
    "CECILION": ("MAGE", "MID"), "CARMILLA": ("SUPPORT", "ROAM"), "ATLAS": ("TANK", "ROAM"),
    "POPOL AND KUPA": ("MARKSMAN", "GOLD"), "YU ZHONG": ("FIGHTER", "JUNGLE"), "LUO YI": ("MAGE", "MID"),
    "BENEDETTA": ("ASSASSIN", "JUNGLE"), "KHALEED": ("FIGHTER", "EXP"), "BARATS": ("TANK", "EXP"),
    "BRODY": ("MARKSMAN", "GOLD"), "YVE": ("MAGE", "MID"), "PAQUITO": ("FIGHTER", "EXP"),
    "GLOO": ("TANK", "EXP"), "BEATRIX": ("MARKSMAN", "GOLD"), "PHOVEUS": ("FIGHTER", "JUNGLE"),
    "NATAN": ("MARKSMAN", "GOLD"), "AULUS": ("FIGHTER", "JUNGLE"), "AAMON": ("ASSASSIN", "JUNGLE"),
    "VALENTINA": ("MAGE", "MID"), "EDITH": ("TANK", "ROAM"), "FLORYN": ("SUPPORT", "ROAM"),
    "YIN": ("FIGHTER", "JUNGLE"), "MELISSA": ("MARKSMAN", "GOLD"), "XAVIER": ("MAGE", "MID"),
    "FREDRINN": ("FIGHTER", "JUNGLE"), "JOY": ("ASSASSIN", "JUNGLE"), "NOVARIA": ("MAGE", "MID"),
    "ARLOTT": ("FIGHTER", "EXP"), "IXIA": ("MARKSMAN", "GOLD"), "NOLAN": ("ASSASSIN", "JUNGLE"),
    "CICI": ("FIGHTER", "EXP"), "CHIP": ("SUPPORT", "ROAM"), "ZHUXIN": ("MAGE", "MID"),
    "SUYOU": ("ASSASSIN", "JUNGLE"), "LUKAS": ("FIGHTER", "EXP"), "ZETIAN": ("MAGE", "MID"),
    "KALEA": ("FIGHTER", "EXP"),
}

# Birden fazla rol / lane'de oynanabilen kahramanlar icin ozel gecersiz kilma listesi
_MULTI_OVERRIDES = {
    "BANE": (["FIGHTER", "MAGE"], ["GOLD"]),
    "ROGER": (["FIGHTER", "MARKSMAN"], ["JUNGLE"]),
    "HARLEY": (["MAGE", "ASSASSIN"], ["MID"]),
    "SELENA": (["ASSASSIN", "MAGE"], ["MID", "ROAM"]),
    "KADITA": (["ASSASSIN", "MAGE"], ["MID"]),
    "ESMERALDA": (["MAGE", "TANK"], ["EXP", "ROAM"]),
    "JULIAN": (["ASSASSIN", "FIGHTER"], ["JUNGLE"]),
    "MATHILDA": (["ASSASSIN", "SUPPORT"], ["JUNGLE", "ROAM"]),
    "YIN": (["FIGHTER", "ASSASSIN"], ["JUNGLE"]),
    "BENEDETTA": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "SUYOU": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "LING": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "KHUFRA": (["TANK"], ["ROAM", "JUNGLE"]),
    "GUSION": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "HANZO": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "AAMON": (["ASSASSIN"], ["JUNGLE", "MID"]),
    "FREDRINN": (["FIGHTER", "TANK"], ["JUNGLE"]),
    "PAQUITO": (["FIGHTER", "ASSASSIN"], ["EXP"]),
    "CARMILLA": (["SUPPORT", "TANK"], ["ROAM"]),
    "GLOO": (["TANK", "FIGHTER"], ["EXP"]),
    "BARATS": (["TANK", "FIGHTER"], ["EXP"]),
    "URANUS": (["TANK", "FIGHTER"], ["EXP"]),
    "MASHA": (["FIGHTER", "ASSASSIN"], ["JUNGLE"]),
}

DEFAULT_HERO_ROLES = {name: ([role], [lane]) for name, (role, lane) in _SINGLE.items()}
DEFAULT_HERO_ROLES.update(_MULTI_OVERRIDES)


# ====================================================================
# YARDIMCI FONKSIYONLAR
# ====================================================================
def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(IMAGES_DIR, exist_ok=True)
    os.makedirs(RANKS_DIR, exist_ok=True)


def normalize_key(s):
    """Karsilastirma icin: kucuk harfe cevir, harf/rakam disindaki her seyi at.
    Boylece 'Mythical Glory+', 'mythical_glory', 'MythicalGlory.png' gibi farkli
    yazimlar ayni kabul edilir."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


IMG_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".bmp")


def sanitize_filename(name):
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", name.strip())
    return safe.strip("_") or "hero"


def local_image_path(hero_name):
    """Yeni indirilecek bir portrenin kaydedilecegi varsayilan (.png) yol."""
    return os.path.join(IMAGES_DIR, sanitize_filename(hero_name) + ".png")


def find_local_hero_image(hero_name):
    """IMAGES/HEROES klasorunde bu kahramana ait bir resim ara (elle
    eklenmis veya adi degistirilmis dosyalar dahil). Once tam isim + bilinen
    uzantilarla, sonra normalize edilmis isim eslesmesiyle dener."""
    if not os.path.isdir(IMAGES_DIR):
        return None
    base = sanitize_filename(hero_name)
    for ext in IMG_EXTENSIONS:
        p = os.path.join(IMAGES_DIR, base + ext)
        if os.path.exists(p):
            return p
    target = normalize_key(hero_name)
    try:
        for fname in os.listdir(IMAGES_DIR):
            b, ext = os.path.splitext(fname)
            if ext.lower() not in IMG_EXTENSIONS:
                continue
            if normalize_key(b) == target:
                return os.path.join(IMAGES_DIR, fname)
    except Exception:
        pass
    return None


def find_rank_logo(rank_name):
    """IMAGES/RANKS klasorunde verilen rank/lig adiyla eslesen bir logo ara."""
    if not rank_name or not os.path.isdir(RANKS_DIR):
        return None
    target = normalize_key(rank_name)
    best_partial = None
    try:
        for fname in os.listdir(RANKS_DIR):
            b, ext = os.path.splitext(fname)
            if ext.lower() not in IMG_EXTENSIONS:
                continue
            nb = normalize_key(b)
            if nb == target:
                return os.path.join(RANKS_DIR, fname)
            if best_partial is None and (target in nb or nb in target) and nb:
                best_partial = os.path.join(RANKS_DIR, fname)
    except Exception:
        pass
    return best_partial


def save_last_query(rank, time_range):
    ensure_output_dir()
    try:
        with open(LAST_QUERY_FILE, "w", encoding="utf-8") as f:
            json.dump({"rank": rank, "time_range": time_range}, f, ensure_ascii=False)
    except Exception:
        pass


def load_last_query():
    if os.path.exists(LAST_QUERY_FILE):
        try:
            with open(LAST_QUERY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"rank": None, "time_range": None}


def download_hero_image(hero_name, url, log=None):
    """Once IMAGES/HEROES klasorunde bu kahramana ait elle eklenmis/tasinmis
    bir resim var mi diye bakar (varsa onu kullanir, tekrar indirmez).
    Yoksa verilen URL'den indirir. Basarili olursa yerel dosya yolunu,
    olmazsa None doner."""
    ensure_output_dir()
    existing = find_local_hero_image(hero_name)
    if existing:
        return existing
    if not url:
        return None
    dest = local_image_path(hero_name)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = resp.read()
        with open(dest, "wb") as f:
            f.write(data)
        return dest
    except Exception as e:
        if log:
            log(f"Resim indirilemedi ({hero_name}): {e}")
        return None


def make_circular_pixmap(path, size=46):
    """Verilen resmi daire seklinde kirpilmis bir QPixmap olarak doner.
    Dosya yoksa/bozuksa None doner (o zaman placeholder daire kullanilir)."""
    if not path or not os.path.exists(path):
        return None
    src = QPixmap(path)
    if src.isNull():
        return None
    src = src.scaled(size, size, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
    result = QPixmap(size, size)
    result.fill(Qt.transparent)
    painter = QPainter(result)
    painter.setRenderHint(QPainter.Antialiasing)
    path_clip = QPainterPath()
    path_clip.addEllipse(0, 0, size, size)
    painter.setClipPath(path_clip)
    x = (size - src.width()) // 2
    y = (size - src.height()) // 2
    painter.drawPixmap(x, y, src)
    painter.end()
    return result


def parse_num(s):
    if s is None:
        return 0.0
    s = str(s).replace("%", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def compute_power(pick, win, ban):
    return round((1.4 * pick * win) + (ban * 3), 1)


def assign_tier(percentile):
    for tier, cut in TIER_CUTOFFS:
        if percentile <= cut:
            return tier
    return "F"


def compute_tiers(hero_list):
    ordered = sorted(hero_list, key=lambda h: h["power"], reverse=True)
    n = len(ordered)
    for idx, h in enumerate(ordered):
        pct = (idx + 1) / n if n else 1
        h["tier"] = assign_tier(pct)
    return ordered


def load_hero_roles():
    ensure_output_dir()
    roles = {k: (list(v[0]), list(v[1])) for k, v in DEFAULT_HERO_ROLES.items()}
    if os.path.exists(HERO_ROLES_FILE):
        try:
            with open(HERO_ROLES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for name, rl in saved.items():
                roles[name] = (list(rl[0]), list(rl[1]))
        except Exception:
            pass
    return roles


def save_hero_roles(roles_dict):
    ensure_output_dir()
    with open(HERO_ROLES_FILE, "w", encoding="utf-8") as f:
        json.dump({k: [v[0], v[1]] for k, v in roles_dict.items()}, f, ensure_ascii=False, indent=2)


def get_time_xpath(option_text):
    return f"//div[@data-node='2684831' and .//span[normalize-space()='{option_text}']]"


def get_rank_xpath(option_text):
    return f"//div[@data-node='2684882' and .//span[normalize-space()='{option_text}']]"


def initials_of(name):
    parts = [p for p in name.replace(".", " ").replace("'", " ").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[1][0]).upper()


# ====================================================================
# MODERN GORUNUM (QSS)
# ====================================================================
APP_STYLE = """
QWidget {
    background-color: #131a2b;
    color: #e8ecf5;
    font-family: "Segoe UI", "Trebuchet MS", sans-serif;
    font-size: 10.5pt;
}
QMainWindow { background-color: #0f1526; }

QTabWidget::pane { border: 1px solid #26304a; border-radius: 10px; top: -1px; background: #171f36; }
QTabBar::tab {
    background: #1a2238; color: #9fabc7; padding: 10px 22px; margin-right: 4px;
    border-top-left-radius: 8px; border-top-right-radius: 8px; font-weight: 600;
}
QTabBar::tab:selected { background: #f2b134; color: #14192b; }
QTabBar::tab:hover:!selected { background: #232d4a; color: #ffffff; }

QGroupBox {
    border: 1px solid #2a3555; border-radius: 10px; margin-top: 14px;
    padding-top: 14px; font-weight: 600; color: #f2b134;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }

QPushButton {
    background-color: #f2b134; color: #14192b; border: none; border-radius: 8px;
    padding: 9px 18px; font-weight: 700;
}
QPushButton:hover { background-color: #ffc75a; }
QPushButton:pressed { background-color: #d99a1f; }
QPushButton:disabled { background-color: #3a4160; color: #7d869e; }

QPushButton#secondary {
    background-color: #232d4a; color: #e8ecf5; border: 1px solid #384a7a;
}
QPushButton#secondary:hover { background-color: #2c3860; }

QLineEdit, QComboBox {
    background-color: #1c2540; border: 1px solid #2f3a5c; border-radius: 6px;
    padding: 6px 10px; color: #e8ecf5; selection-background-color: #f2b134;
}
QComboBox::drop-down { border: none; width: 24px; }
QComboBox QAbstractItemView {
    background-color: #1c2540; color: #e8ecf5; selection-background-color: #f2b134;
    selection-color: #14192b; outline: none; border: 1px solid #2f3a5c;
}

QTextEdit {
    background-color: #0d1220; border: 1px solid #2a3555; border-radius: 8px; color: #9be89b;
}

QTableWidget {
    background-color: #171f36; alternate-background-color: #1b2440;
    gridline-color: #2a3555; border: 1px solid #2a3555; border-radius: 8px;
    selection-background-color: #f2b13455;
}
QHeaderView::section {
    background-color: #232d4a; color: #f2b134; padding: 8px; border: none;
    border-bottom: 2px solid #f2b134; font-weight: 700;
}
QTableWidget::item { padding: 4px; }

QScrollBar:vertical { background: #171f36; width: 12px; border-radius: 6px; }
QScrollBar::handle:vertical { background: #384a7a; border-radius: 6px; min-height: 24px; }
QScrollBar::handle:vertical:hover { background: #4a5f9a; }
QScrollBar:horizontal { background: #171f36; height: 12px; border-radius: 6px; }
QScrollBar::handle:horizontal { background: #384a7a; border-radius: 6px; min-width: 24px; }

QProgressBar {
    background-color: #1c2540; border: 1px solid #2f3a5c; border-radius: 6px; text-align: center; color: #e8ecf5;
}
QProgressBar::chunk { background-color: #f2b134; border-radius: 6px; }

QLabel#pageTitle { font-size: 15pt; font-weight: 800; color: #f2b134; }
QLabel#hint { color: #8b96b5; font-style: italic; }

QDialog { background-color: #131a2b; }
"""


# ====================================================================
# HERO CIRCLE (kahraman portresi; resim yoksa bas harfli placeholder daire)
# ====================================================================
class HeroCircle(QLabel):
    def __init__(self, hero, size=46, parent=None):
        super().__init__(parent)
        self.SIZE = size
        self.setFixedSize(self.SIZE, self.SIZE)
        self.setAlignment(Qt.AlignCenter)
        border_color = TIER_COLORS.get(hero["tier"], "#666666")
        border_w = 2 if self.SIZE >= 28 else 1
        font_pt = max(6, self.SIZE // 4)

        pixmap = make_circular_pixmap(hero.get("image"), max(4, self.SIZE - 2 * border_w))
        if pixmap is not None:
            # Gercek portre bulundu: resmi goster, ince tier renkli cerceve ciz
            canvas = QPixmap(self.SIZE, self.SIZE)
            canvas.fill(Qt.transparent)
            painter = QPainter(canvas)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.drawPixmap(border_w, border_w, pixmap)
            pen_color = QColor(border_color)
            painter.setPen(pen_color)
            painter.setBrush(Qt.NoBrush)
            painter.drawEllipse(border_w - 1, border_w - 1, self.SIZE - border_w, self.SIZE - border_w)
            painter.end()
            self.setPixmap(canvas)
        else:
            # Resim henuz indirilmemis / bulunamamis: bos placeholder daire
            self.setText(initials_of(hero["name"]))
            self.setStyleSheet(f"""
                QLabel {{
                    background-color: #232d4a;
                    color: #e8ecf5;
                    border: {border_w}px solid {border_color};
                    border-radius: {self.SIZE // 2}px;
                    font-weight: 700;
                    font-size: {font_pt}pt;
                }}
            """)

        tip = f"{hero['name']}\nPower: {hero['power']}\nRol: {', '.join(hero['role'])}\nLane: {', '.join(hero['lane'])}\nTier: {hero['tier']}"
        self.setToolTip(tip)


# ====================================================================
# COKLU SECIMLI COMBOBOX (rol / lane icin)
# ====================================================================
class CheckableComboBox(QComboBox):
    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().installEventFilter(self)
        self._model = QStandardItemModel(self)
        for text in items:
            item = QStandardItem(text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setData(Qt.Unchecked, Qt.CheckStateRole)
            self._model.appendRow(item)
        self.setModel(self._model)
        self.view().pressed.connect(self._on_item_pressed)
        self._closing = False

    def _on_item_pressed(self, index):
        item = self._model.itemFromIndex(index)
        new_state = Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked
        item.setCheckState(new_state)
        self._update_text()

    def eventFilter(self, obj, event):
        if obj == self.lineEdit() and event.type() == event.MouseButtonPress:
            self.showPopup()
            return True
        return super().eventFilter(obj, event)

    def _update_text(self):
        checked = self.checked_items()
        self.lineEdit().setText(", ".join(checked) if checked else "")

    def checked_items(self):
        result = []
        for i in range(self._model.rowCount()):
            item = self._model.item(i)
            if item.checkState() == Qt.Checked:
                result.append(item.text())
        return result

    def set_checked_items(self, values):
        values = set(values)
        for i in range(self._model.rowCount()):
            item = self._model.item(i)
            item.setCheckState(Qt.Checked if item.text() in values else Qt.Unchecked)
        self._update_text()


class MultiSelectDelegate(QStyledItemDelegate):
    """Rol / Lane sutunlari icin coklu secimli editor."""

    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.items = items

    def createEditor(self, parent, option, index):
        combo = CheckableComboBox(self.items, parent)
        return combo

    def setEditorData(self, editor, index):
        current = index.data(Qt.DisplayRole) or ""
        values = [v.strip() for v in current.split(",") if v.strip()]
        editor.set_checked_items(values)

    def setModelData(self, editor, model, index):
        model.setData(index, ", ".join(editor.checked_items()), Qt.DisplayRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


# ====================================================================
# YENI KAHRAMAN BULUNDU DIYALOGU
# ====================================================================
class NewHeroDialog(QDialog):
    def __init__(self, new_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Yeni Kahraman(lar) Bulundu")
        self.resize(560, 460)
        self.result_map = {}

        layout = QVBoxLayout(self)
        title = QLabel(f"⚠ {len(new_names)} yeni kahraman tespit edildi. Rol ve Lane sec:")
        title.setObjectName("pageTitle")
        title.setWordWrap(True)
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        form = QFormLayout(inner)
        scroll.setWidget(inner)
        layout.addWidget(scroll)

        self.combos = {}
        for name in new_names:
            role_combo = CheckableComboBox(ROLES_LIST)
            lane_combo = CheckableComboBox(LANES_LIST)
            row = QHBoxLayout()
            row.addWidget(QLabel("Rol:"))
            row.addWidget(role_combo)
            row.addWidget(QLabel("Lane:"))
            row.addWidget(lane_combo)
            container = QWidget()
            container.setLayout(row)
            form.addRow(QLabel(f"<b>{name}</b>"), container)
            self.combos[name] = (role_combo, lane_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(self.on_accept)
        layout.addWidget(buttons)

    def on_accept(self):
        for name, (role_combo, lane_combo) in self.combos.items():
            roles = role_combo.checked_items() or ["UNKNOWN"]
            lanes = lane_combo.checked_items() or ["UNKNOWN"]
            self.result_map[name] = (roles, lanes)
        self.accept()


# ====================================================================
# SCRAPER THREAD
# ====================================================================
class ScraperThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, selected_time, selected_rank, headless=False):
        super().__init__()
        self.selected_time = selected_time
        self.selected_rank = selected_rank
        self.headless = headless

    def log(self, msg):
        self.log_signal.emit(msg)

    def wait_present(self, driver, xpath, timeout=LOAD_WAIT):
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

    def wait_clickable(self, driver, xpath, timeout=LOAD_WAIT):
        return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def current_rows(self, driver):
        rows = driver.find_elements(By.XPATH, X_ROWS)
        if not rows:
            rows = driver.find_elements(By.XPATH, X_ROWS_CLASS)
        if not rows:
            rows = driver.find_elements(By.XPATH, X_ROWS_GENERIC)
        return rows

    def wait_rows_increase(self, driver, previous_count, timeout=15):
        end = time.time() + timeout
        while time.time() < end:
            cnt = len(self.current_rows(driver))
            if cnt > previous_count:
                return cnt
            time.sleep(0.5)
        return previous_count

    def scroll_container_or_window(self, driver):
        # ONCELIK: gercek fare tekerlegi (wheel) davranisini taklit eden
        # ActionChains kullan, AMA hedefi mutlaka listeyi iceren konteynere
        # (X_PARENT) sabitle. ScrollOrigin vermeden yapilan genel bir
        # scroll_by_amount, sayfadaki o anki mouse konumunu (genelde
        # (0,0)/viewport) kaydirir; bu da ic ice gecmis (nested) kaydirilabilir
        # listeyi degil, sayfanin kendisini kaydirip listeyi hic tetiklemeyebilir
        # (bu durumda liste ilk ~20 satirdan sonra yuklenmeyi durdurur).
        try:
            parent = driver.find_element(By.XPATH, X_PARENT)
        except Exception:
            parent = None

        if parent is not None:
            try:
                origin = ScrollOrigin.from_element(parent)
                ActionChains(driver).scroll_from_origin(origin, 0, 700).perform()
                time.sleep(0.8)
                return
            except Exception:
                pass

        try:
            ActionChains(driver).scroll_by_amount(0, 700).perform()
            time.sleep(0.8)
            return
        except Exception:
            pass

        # YEDEK YONTEM (eski selenium/tarayici surumlerinde ActionChains
        # desteklenmezse): eski JS tabanli kaydirma.
        scrolled = False
        if parent is not None:
            try:
                driver.execute_script(
                    "arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight - 80;", parent
                )
                scrolled = True
            except Exception:
                scrolled = False
        if not scrolled:
            driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(0.6)

    def scroll_until_all_loaded(self, driver, max_rounds=300):
        # KURAL: son iki turda gorunen satir sayisi AYNIYSA, tum kahramanlarin
        # yuklendigi kabul edilip kaydirma durur (yeni hero eklense/cikarilsa
        # bile kod degismesine gerek kalmaz).
        self.log("Tum kahramanlar yuklenene kadar kaydiriliyor...")
        prev_cnt = -1
        round_no = 0
        while round_no < max_rounds:
            round_no += 1
            rows = self.current_rows(driver)
            cnt = len(rows)
            self.log(f"Tur {round_no} - gorunen satir: {cnt}")

            if cnt == prev_cnt and cnt > 0:
                self.log(f"Son 2 turda satir sayisi degismedi ({cnt}), tum kahramanlar yuklendi kabul ediliyor.")
                break
            prev_cnt = cnt

            if rows:
                try:
                    ActionChains(driver).scroll_to_element(rows[-1]).perform()
                except Exception:
                    try:
                        driver.execute_script(
                            "arguments[0].scrollIntoView({behavior: 'instant', block: 'end'});", rows[-1]
                        )
                    except Exception:
                        pass
            # Bazi lazy-load listeleri tek kaydirmada tetiklenmiyor, guvenlik
            # icin ust uste iki kaydirma uyguluyoruz.
            self.scroll_container_or_window(driver)
            self.scroll_container_or_window(driver)
            # DOM'a yeni satirlarin gelmesi (network + render) icin bekleme
            time.sleep(1.2)

    def run(self):
        driver = None
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-extensions")
            # Sayfa "load" (tum kaynaklar/reklam/analytics bitene kadar) yerine
            # DOM hazir olunca donsun; bazen agir reklam/analytics scriptleri
            # yuzunden "load" olayi hic tetiklenmeyip driver.get() sonsuza kadar
            # bekleyebiliyor (ekranda Chrome'un bos yeni sekme/Google sayfasi
            # takili kaliyormus gibi gorunmesinin sebebi budur).
            options.page_load_strategy = "eager"
            # Kahraman resimlerini zaten ayrica indiriyoruz (urllib ile), bu yuzden
            # tarayicinin resimleri render etmesine gerek yok - hem hizlanir hem
            # de "img src" ozniteligi resim indirilmeden de DOM'da mevcut olur.
            options.add_experimental_option(
                "prefs", {"profile.managed_default_content_settings.images": 2}
            )
            if self.headless:
                options.add_argument("--headless=new")
            self.log("Chrome baslatiliyor...")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            driver.set_window_size(*WINDOW_SIZE)
            driver.set_window_position(50, 50)
            driver.set_page_load_timeout(35)

            self.log(f"Sayfa aciliyor: {URL}")
            try:
                driver.get(URL)
            except TimeoutException:
                # "eager" stratejisiyle nadiren de olsa hala takilirsa: yuklemeyi
                # zorla durdur ve DOM'da ihtiyacimiz olan elementler var mi diye
                # devam et (genelde bu noktada icerik zaten gelmis olur).
                self.log("Sayfa yuklemesi zaman asimina ugradi, yukleme durduruluyor ve devam ediliyor...")
                try:
                    driver.execute_script("window.stop();")
                except Exception:
                    pass
            self.wait_present(driver, "//body")

            try:
                btn = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.CLASS_NAME, "mt-cb-policy-close")))
                btn.click()
            except Exception:
                pass
            try:
                btn = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, "mt-cb-s")))
                btn.click()
            except Exception:
                pass

            self.wait_present(driver, X_PARENT, timeout=LOAD_WAIT)

            self.log(f"Zaman filtresi ayarlaniyor: {self.selected_time}")
            self.wait_clickable(driver, X_TIME_DROPDOWN).click()
            self.wait_clickable(driver, get_time_xpath(self.selected_time)).click()
            self.wait_rows_increase(driver, 0, timeout=8)

            self.log(f"Rank filtresi ayarlaniyor: {self.selected_rank}")
            self.wait_clickable(driver, X_RANK_DROPDOWN).click()
            self.wait_clickable(driver, get_rank_xpath(self.selected_rank)).click()
            self.wait_rows_increase(driver, 0, timeout=12)

            self.scroll_until_all_loaded(driver)

            rows = self.current_rows(driver)
            self.log(f"Toplam {len(rows)} satir bulundu, veriler okunuyor...")

            results = []
            for row in rows:
                try:
                    name = row.find_element(By.XPATH, X_NAME).text.strip().upper()
                    pick = parse_num(row.find_element(By.XPATH, X_PICK).text)
                    win = parse_num(row.find_element(By.XPATH, X_WIN).text)
                    ban = parse_num(row.find_element(By.XPATH, X_BAN).text)
                    img_url = None
                    try:
                        img_url = row.find_element(By.XPATH, X_IMG).get_attribute("src")
                    except Exception:
                        img_url = None
                    if name:
                        results.append({"name": name, "pick": pick, "win": win, "ban": ban, "image_url": img_url})
                except Exception:
                    continue

            self.log(f"Okuma tamamlandi: {len(results)} kahraman. Portreler indiriliyor...")
            for h in results:
                path = download_hero_image(h["name"], h.get("image_url"), log=self.log)
                h["image"] = path
            downloaded = sum(1 for h in results if h.get("image"))
            self.log(f"Portre indirme tamamlandi: {downloaded}/{len(results)}.")
            self.finished_signal.emit(results)

        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass


# ====================================================================
# TABLO ITEM SINIFI
# ====================================================================
class NumericItem(QTableWidgetItem):
    def __init__(self, value, text=None):
        super().__init__(text if text is not None else str(value))
        self.value = value
        self.setTextAlignment(Qt.AlignCenter)

    def __lt__(self, other):
        if isinstance(other, NumericItem):
            return self.value < other.value
        return super().__lt__(other)


# ====================================================================
# ANA PENCERE
# ====================================================================
COLS = ["TIER", "HERO POWER", "HERO NAME", "PICK RATE", "WIN RATE", "BAN RATE", "HERO ROLE", "HERO LANE"]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mobile Legends Meta Tool")
        self.resize(1250, 760)

        self.hero_roles = load_hero_roles()
        self.hero_data = []
        self.scraper_thread = None
        self.rank_badge_logos = []
        self.rank_badge_texts = []
        self.big_rank_logos = []
        self.big_rank_texts = []

        last_query = load_last_query()
        self.current_rank = last_query.get("rank")
        self.current_time_range = last_query.get("time_range")

        tabs = QTabWidget()
        self.setCentralWidget(tabs)
        tabs.addTab(self.build_scrape_tab(), "⚔  Veri Cek")
        tabs.addTab(self.build_meta_tab(), "📊  Meta Tier List")
        tabs.addTab(self.build_lane_tier_tab(), "🗂  Lane x Tier Tablosu")

        self.try_autoload()
        self.update_rank_badge()

    def make_rank_badge(self):
        """Sag ust kosede gosterilecek kucuk 'RANK · Zaman Araligi' rozeti (logo + yazi)."""
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)
        logo_lbl = QLabel()
        logo_lbl.setFixedSize(28, 28)
        logo_lbl.setScaledContents(True)
        logo_lbl.setVisible(False)
        text_lbl = QLabel("Veri yok")
        text_lbl.setStyleSheet("color:#f2b134; font-weight:700;")
        row.addWidget(logo_lbl)
        row.addWidget(text_lbl)
        self.rank_badge_logos.append(logo_lbl)
        self.rank_badge_texts.append(text_lbl)
        return container

    def make_big_rank_header(self):
        """Lane x Tier sekmesinin ustunde buyuk sekilde gosterilen
        'RANK · Zaman Araligi' basligi (sayfa basligi yerine gecer)."""
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)
        logo_lbl = QLabel()
        logo_lbl.setFixedSize(42, 42)
        logo_lbl.setScaledContents(True)
        logo_lbl.setVisible(False)
        text_lbl = QLabel("Veri yok")
        text_lbl.setObjectName("pageTitle")
        text_lbl.setStyleSheet("font-size: 19pt; font-weight: 800; color: #f2b134;")
        row.addWidget(logo_lbl)
        row.addWidget(text_lbl)
        self.big_rank_logos.append(logo_lbl)
        self.big_rank_texts.append(text_lbl)
        return container

    def update_rank_badge(self):
        if self.current_rank and self.current_time_range:
            text = f"{self.current_rank}  ·  {self.current_time_range}"
        elif self.current_rank or self.current_time_range:
            text = self.current_rank or self.current_time_range
        else:
            text = "Veri yok"

        logo_path = find_rank_logo(self.current_rank) if self.current_rank else None
        raw_pixmap = QPixmap(logo_path) if logo_path else None
        if raw_pixmap is not None and raw_pixmap.isNull():
            raw_pixmap = None

        for lbl in self.rank_badge_texts:
            lbl.setText(text)
        for lbl in self.rank_badge_logos:
            if raw_pixmap is not None:
                lbl.setPixmap(raw_pixmap.scaled(28, 28, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                lbl.setVisible(True)
            else:
                lbl.clear()
                lbl.setVisible(False)

        for lbl in self.big_rank_texts:
            lbl.setText(text)
        for lbl in self.big_rank_logos:
            if raw_pixmap is not None:
                lbl.setPixmap(raw_pixmap.scaled(42, 42, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                lbl.setVisible(True)
            else:
                lbl.clear()
                lbl.setVisible(False)

    # ----------------------------------------------------------------
    # TAB 1: VERI CEK
    # ----------------------------------------------------------------
    def build_scrape_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        title = QLabel("Mobile Legends Rank Verisi Cek")
        title.setObjectName("pageTitle")
        layout.addWidget(title)

        form_box = QGroupBox("Filtre Secenekleri")
        form = QFormLayout(form_box)
        form.setSpacing(10)

        self.time_combo = QComboBox()
        self.time_combo.addItems(TIME_OPTIONS)
        self.time_combo.setCurrentText("Past 30 days")
        form.addRow("Zaman Araligi:", self.time_combo)

        self.rank_combo = QComboBox()
        self.rank_combo.addItems(RANK_OPTIONS)
        self.rank_combo.setCurrentText("Mythical Glory+")
        form.addRow("Rank:", self.rank_combo)

        layout.addWidget(form_box)

        btn_row = QHBoxLayout()
        self.scrape_btn = QPushButton("▶  Veri Cek")
        self.scrape_btn.clicked.connect(self.start_scrape)
        btn_row.addWidget(self.scrape_btn)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        btn_row.addWidget(self.progress)
        layout.addLayout(btn_row)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 9))
        layout.addWidget(self.log_box)

        return w

    def start_scrape(self):
        self.scrape_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.log_box.clear()
        self.log_box.append("Tarayici baslatiliyor, bu pencereyi kapatma...")

        self.scraper_thread = ScraperThread(self.time_combo.currentText(), self.rank_combo.currentText())
        self.scraper_thread.log_signal.connect(self.log_box.append)
        self.scraper_thread.finished_signal.connect(self.on_scrape_finished)
        self.scraper_thread.error_signal.connect(self.on_scrape_error)
        self.scraper_thread.start()

    def on_scrape_error(self, msg):
        self.progress.setVisible(False)
        self.scrape_btn.setEnabled(True)
        self.log_box.append(f"HATA: {msg}")
        QMessageBox.critical(self, "Hata", f"Veri cekme basarisiz oldu:\n{msg}")

    def on_scrape_finished(self, results):
        self.progress.setVisible(False)
        self.scrape_btn.setEnabled(True)
        self.log_box.append(f"Tamamlandi: {len(results)} kahraman verisi alindi.")
        if not results:
            QMessageBox.warning(self, "Uyari", "Hic veri bulunamadi.")
            return
        if self.scraper_thread is not None:
            self.current_rank = self.scraper_thread.selected_rank
            self.current_time_range = self.scraper_thread.selected_time
            save_last_query(self.current_rank, self.current_time_range)
            self.update_rank_badge()
        self.process_and_load(results)
        self.save_raw_data(results)
        QMessageBox.information(self, "Bitti", "Veri cekme tamamlandi. Diger sekmelere gecebilirsin.")

    def save_raw_data(self, results):
        ensure_output_dir()
        with open(RAW_DATA_FILE, "w", encoding="utf-8") as f:
            for h in results:
                f.write(f"{h['name']}\t{h['pick']}\t{h['win']}\t{h['ban']}\n")

    # ----------------------------------------------------------------
    # TAB 2: META TIER LIST
    # ----------------------------------------------------------------
    def build_meta_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title = QLabel("Meta Tier List")
        title.setObjectName("pageTitle")
        title_row.addWidget(title)
        title_row.addStretch()
        title_row.addWidget(self.make_rank_badge())
        layout.addLayout(title_row)

        filter_box = QGroupBox("Filtreler (birlikte calisir - orn: Rol=ASSASSIN + Tier=SSS)")
        f_layout = QHBoxLayout(filter_box)

        f_layout.addWidget(QLabel("Rol:"))
        self.role_filter = QComboBox()
        self.role_filter.addItems(["TUMU"] + ROLES_LIST)
        self.role_filter.currentIndexChanged.connect(self.apply_filters)
        f_layout.addWidget(self.role_filter)

        f_layout.addWidget(QLabel("Lane:"))
        self.lane_filter = QComboBox()
        self.lane_filter.addItems(["TUMU"] + LANES_LIST)
        self.lane_filter.currentIndexChanged.connect(self.apply_filters)
        f_layout.addWidget(self.lane_filter)

        f_layout.addWidget(QLabel("Tier:"))
        self.tier_filter = QComboBox()
        self.tier_filter.addItems(["TUMU"] + TIER_ORDER)
        self.tier_filter.currentIndexChanged.connect(self.apply_filters)
        f_layout.addWidget(self.tier_filter)

        f_layout.addWidget(QLabel("Ara:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Kahraman adi...")
        self.search_box.textChanged.connect(self.apply_filters)
        f_layout.addWidget(self.search_box)

        clear_btn = QPushButton("Filtreleri Temizle")
        clear_btn.setObjectName("secondary")
        clear_btn.clicked.connect(self.clear_filters)
        f_layout.addWidget(clear_btn)

        layout.addWidget(filter_box)

        btn_row = QHBoxLayout()
        reload_btn = QPushButton("⟳  Dosyadan Yenile")
        reload_btn.setObjectName("secondary")
        reload_btn.clicked.connect(self.try_autoload)
        btn_row.addWidget(reload_btn)

        save_roles_btn = QPushButton("💾  Rol/Lane Degisikliklerini Kaydet")
        save_roles_btn.clicked.connect(self.save_role_edits)
        btn_row.addWidget(save_roles_btn)

        export_btn = QPushButton("⇩  Excel'e Aktar")
        export_btn.clicked.connect(self.export_excel)
        btn_row.addWidget(export_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        hint = QLabel("İpucu: Sütun başlığına tıklayınca büyükten küçüğe / küçükten büyüğe sıralanır. "
                      "Rol ve Lane hücrelerine çift tıklayıp birden fazla seçenek işaretleyebilirsin "
                      "(örn. hem ASSASSIN hem MAGE, hem MID hem JUNGLE).")
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.table = QTableWidget()
        self.table.setColumnCount(len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        for i in [0, 1, 3, 4, 5, 6, 7]:
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
        self.table.setItemDelegateForColumn(6, MultiSelectDelegate(ROLES_LIST, self.table))
        self.table.setItemDelegateForColumn(7, MultiSelectDelegate(LANES_LIST, self.table))
        self.table.setIconSize(QSize(28, 28))
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.itemChanged.connect(self.on_cell_edited)
        layout.addWidget(self.table)

        return w

    # ----------------------------------------------------------------
    # TAB 3: LANE x TIER TABLOSU
    # ----------------------------------------------------------------
    def build_lane_tier_tab(self):
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(8)

        header_row = QHBoxLayout()
        header_row.addWidget(self.make_big_rank_header())
        header_row.addStretch()
        refresh_btn = QPushButton("⟳  Yenile")
        refresh_btn.setObjectName("secondary")
        refresh_btn.clicked.connect(self.refresh_lane_tier_table)
        header_row.addWidget(refresh_btn)
        outer.addLayout(header_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.lane_tier_container = QWidget()
        self.lane_tier_grid = QVBoxLayout(self.lane_tier_container)
        self.lane_tier_grid.setSpacing(8)
        scroll.setWidget(self.lane_tier_container)
        outer.addWidget(scroll)

        return w

    def refresh_lane_tier_table(self):
        # eski icerigi temizle
        while self.lane_tier_grid.count():
            item = self.lane_tier_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.hero_data:
            empty = QLabel("Henuz veri yok. Once 'Veri Cek' sekmesinden tarama yap.")
            empty.setObjectName("hint")
            self.lane_tier_grid.addWidget(empty)
            return

        # BASLIK SATIRI
        header_row = QHBoxLayout()
        header_row.addWidget(self._tier_label_cell("TIER", is_header=True))
        for lane in LANES_LIST:
            lbl = QLabel(LANE_LABELS.get(lane, lane))
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("background:#232d4a; color:#f2b134; font-weight:800; "
                               "border-radius:6px; padding:4px; font-size:9pt;")
            header_row.addWidget(lbl, 1)
        header_widget = QWidget()
        header_widget.setLayout(header_row)
        self.lane_tier_grid.addWidget(header_widget)

        # Bu tier+lane kombinasyonlarina dusen kahramanlari once bir kere
        # hesapla (asagidaki boyut denemelerinde tekrar tekrar filtrelemeyelim).
        heroes_by_cell = {}
        for tier in TIER_ORDER:
            for lane in LANES_LIST:
                heroes_by_cell[(tier, lane)] = sorted(
                    [h for h in self.hero_data if h["tier"] == tier and lane in h["lane"]],
                    key=lambda x: x["power"], reverse=True,
                )

        # OTOMATIK OLCEKLEME: 7 tier'in (SSS..F) TAMAMI tek ekranda (kaydirmadan)
        # gorunsun diye, once buyuk daire boyutuyla dener, sigmiyorsa
        # daireyi kucultup sutun sayisini artirarak tekrar dener; ekrana sigan
        # en buyuk boyutu bulunca onunla ciziyoruz. Veri cok fazlaysa (aşırı
        # kalabalik bir tier) en kucuk boyutta bile sigmayabilir; o durumda
        # kaydirma cubugu yine de yedek olarak devrede kalir.
        ROW_SPACING = 4
        self.lane_tier_grid.setSpacing(ROW_SPACING)
        TIER_COL_W = 48 + 8
        lane_area_width = max(90, (self.width() - 40 - TIER_COL_W) // max(1, len(LANES_LIST)))
        available_height = max(220, self.height() - 250)

        candidate_sizes = [34, 30, 26, 22, 18, 16, 14, 12]
        chosen_size, chosen_cols = candidate_sizes[-1], 1
        for size in candidate_sizes:
            spacing = 3
            item_size = size + spacing
            cols = max(1, (lane_area_width - 10) // item_size)
            total_h = 30  # baslik satiri yaklasik yuksekligi
            for tier in TIER_ORDER:
                max_lines = 1
                for lane in LANES_LIST:
                    cnt = len(heroes_by_cell[(tier, lane)])
                    lines = max(1, -(-cnt // cols)) if cnt else 1
                    max_lines = max(max_lines, lines)
                row_h = max(30, max_lines * item_size + 10)
                total_h += row_h + ROW_SPACING
            chosen_size, chosen_cols = size, cols
            if total_h <= available_height:
                break

        item_size = chosen_size + 3
        for tier in TIER_ORDER:
            row_layout = QHBoxLayout()
            row_layout.addWidget(self._tier_label_cell(tier))

            for lane in LANES_LIST:
                heroes = heroes_by_cell[(tier, lane)]
                rows_needed = max(1, -(-len(heroes) // chosen_cols)) if heroes else 1
                cell_height = max(30, rows_needed * item_size + 10)

                cell = QWidget()
                cell.setStyleSheet("background:#171f36; border:1px solid #26304a; border-radius:6px;")
                cell.setMinimumHeight(cell_height)
                grid = QGridLayout(cell)
                grid.setContentsMargins(4, 4, 4, 4)
                grid.setSpacing(3)
                grid.setAlignment(Qt.AlignLeft | Qt.AlignTop)
                for idx, hero in enumerate(heroes):
                    r, c = divmod(idx, chosen_cols)
                    grid.addWidget(HeroCircle(hero, size=chosen_size), r, c)
                row_layout.addWidget(cell, 1)

            row_widget = QWidget()
            row_widget.setLayout(row_layout)
            self.lane_tier_grid.addWidget(row_widget)

        self.lane_tier_grid.addStretch()

    def _tier_label_cell(self, text, is_header=False):
        lbl = QLabel(text)
        lbl.setFixedWidth(48)
        lbl.setAlignment(Qt.AlignCenter)
        if is_header:
            lbl.setStyleSheet("background:#232d4a; color:#f2b134; font-weight:800; "
                               "border-radius:6px; padding:4px; font-size:9pt;")
        else:
            color = TIER_COLORS.get(text, "#666666")
            lbl.setStyleSheet(f"background:{color}; color:#14192b; font-weight:900; "
                               f"border-radius:6px; padding:4px; font-size:10pt;")
        return lbl

    # ----------------------------------------------------------------
    # VERI ISLEME
    # ----------------------------------------------------------------
    def process_and_load(self, raw_results):
        new_names = [r["name"] for r in raw_results if r["name"] not in self.hero_roles]
        if new_names:
            dialog = NewHeroDialog(new_names, self)
            if dialog.exec_() == QDialog.Accepted:
                self.hero_roles.update(dialog.result_map)
            else:
                for name in new_names:
                    self.hero_roles[name] = (["UNKNOWN"], ["UNKNOWN"])
            save_hero_roles(self.hero_roles)

        hero_list = []
        for r in raw_results:
            power = compute_power(r["pick"], r["win"], r["ban"])
            role, lane = self.hero_roles.get(r["name"], (["UNKNOWN"], ["UNKNOWN"]))
            image_path = r.get("image") or find_local_hero_image(r["name"])
            hero_list.append({
                "name": r["name"], "pick": r["pick"], "win": r["win"], "ban": r["ban"],
                "power": power, "role": list(role), "lane": list(lane),
                "image": image_path,
            })
        hero_list = compute_tiers(hero_list)
        self.hero_data = hero_list
        self.save_tierlist(hero_list)
        self.apply_filters()
        self.refresh_lane_tier_table()

    def save_tierlist(self, hero_list):
        ensure_output_dir()
        with open(TIERLIST_FILE, "w", encoding="utf-8") as f:
            for h in sorted(hero_list, key=lambda x: x["power"], reverse=True):
                role_s = ",".join(h["role"])
                lane_s = ",".join(h["lane"])
                image_s = h.get("image") or ""
                f.write(f"{h['tier']}\t{h['power']}\t{h['name']}\t{h['pick']}\t{h['win']}\t{h['ban']}\t{role_s}\t{lane_s}\t{image_s}\n")

    def try_autoload(self):
        if not os.path.exists(TIERLIST_FILE):
            return
        hero_list = []
        try:
            with open(TIERLIST_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    parts = line.rstrip("\n").split("\t")
                    if len(parts) == 9:
                        tier, power, name, pick, win, ban, role_s, lane_s, image_s = parts
                    elif len(parts) == 8:
                        tier, power, name, pick, win, ban, role_s, lane_s = parts
                        image_s = ""
                    else:
                        continue
                    # dosyada yol bos/gecersizse, isme gore yerel resmi tekrar dene
                    image_path = image_s if (image_s and os.path.exists(image_s)) else None
                    if image_path is None:
                        image_path = find_local_hero_image(name)
                    hero_list.append({
                        "name": name, "pick": float(pick), "win": float(win), "ban": float(ban),
                        "power": float(power),
                        "role": [x for x in role_s.split(",") if x],
                        "lane": [x for x in lane_s.split(",") if x],
                        "tier": tier,
                        "image": image_path,
                    })
        except Exception:
            return
        if hero_list:
            self.hero_data = hero_list
            self.apply_filters()
            self.refresh_lane_tier_table()

    # ----------------------------------------------------------------
    # FILTRE / SIRALAMA / TABLO DOLDURMA
    # ----------------------------------------------------------------
    def clear_filters(self):
        self.role_filter.setCurrentIndex(0)
        self.lane_filter.setCurrentIndex(0)
        self.tier_filter.setCurrentIndex(0)
        self.search_box.clear()

    def apply_filters(self):
        role = self.role_filter.currentText()
        lane = self.lane_filter.currentText()
        tier = self.tier_filter.currentText()
        search = self.search_box.text().strip().upper()

        filtered = [
            h for h in self.hero_data
            if (role == "TUMU" or role in h["role"])
            and (lane == "TUMU" or lane in h["lane"])
            and (tier == "TUMU" or h["tier"] == tier)
            and (search in h["name"])
        ]
        self.populate_table(filtered)

    def populate_table(self, data_list):
        self.table.itemChanged.disconnect(self.on_cell_edited)
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.table.setRowCount(len(data_list))

        for row_idx, h in enumerate(sorted(data_list, key=lambda x: x["power"], reverse=True)):
            color = QColor(TIER_COLORS.get(h["tier"], "#666666"))
            color.setAlpha(70)

            tier_item = NumericItem(TIER_RANK.get(h["tier"], 99), text=h["tier"])
            power_item = NumericItem(h["power"], text=str(h["power"]))
            name_item = QTableWidgetItem(h["name"])
            icon_pixmap = make_circular_pixmap(h.get("image"), 28)
            if icon_pixmap is not None:
                name_item.setIcon(QIcon(icon_pixmap))
            pick_item = NumericItem(h["pick"], text=str(h["pick"]))
            win_item = NumericItem(h["win"], text=str(h["win"]))
            ban_item = NumericItem(h["ban"], text=str(h["ban"]))
            role_item = QTableWidgetItem(", ".join(h["role"]))
            lane_item = QTableWidgetItem(", ".join(h["lane"]))

            row_items = [tier_item, power_item, name_item, pick_item, win_item, ban_item, role_item, lane_item]
            for col, item in enumerate(row_items):
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(color)
                if col not in (6, 7):
                    flags = item.flags() & ~Qt.ItemIsEditable
                    item.setFlags(flags)
                self.table.setItem(row_idx, col, item)
            name_item.setData(Qt.UserRole, h["name"])

        self.table.setSortingEnabled(True)
        self.table.itemChanged.connect(self.on_cell_edited)

    def on_cell_edited(self, item):
        if item.column() not in (6, 7):
            return
        row = item.row()
        name_item = self.table.item(row, 2)
        if not name_item:
            return
        hero_name = name_item.text()
        new_values = [v.strip() for v in item.text().split(",") if v.strip()]
        for h in self.hero_data:
            if h["name"] == hero_name:
                if item.column() == 6:
                    h["role"] = new_values or ["UNKNOWN"]
                else:
                    h["lane"] = new_values or ["UNKNOWN"]
                break

    def save_role_edits(self):
        for h in self.hero_data:
            self.hero_roles[h["name"]] = (list(h["role"]), list(h["lane"]))
        save_hero_roles(self.hero_roles)
        self.save_tierlist(self.hero_data)
        self.refresh_lane_tier_table()
        QMessageBox.information(self, "Kaydedildi", "Rol/Lane degisiklikleri kaydedildi.")

    # ----------------------------------------------------------------
    # EXCEL EXPORT
    # ----------------------------------------------------------------
    def export_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "Uyari", "Tabloda aktarilacak veri yok.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Excel Kaydet", os.path.join(OUTPUT_DIR, "meta_export.xlsx"),
                                               "Excel Files (*.xlsx)")
        if not path:
            return

        ensure_output_dir()
        wb = Workbook()
        ws = wb.active
        ws.title = "Meta Tier List"

        header_fill = PatternFill(start_color="B7B7B7", end_color="B7B7B7", fill_type="solid")
        header_font = Font(bold=True)
        thin = Side(style="thin", color="999999")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center")

        for col_idx, header in enumerate(COLS, start=1):
            c = ws.cell(row=1, column=col_idx, value=header)
            c.font = header_font
            c.fill = header_fill
            c.alignment = center
            c.border = border

        for row_idx in range(self.table.rowCount()):
            tier_text = self.table.item(row_idx, 0).text()
            hexcolor = TIER_COLORS.get(tier_text, "#FFFFFF").lstrip("#")
            fill = PatternFill(start_color=hexcolor, end_color=hexcolor, fill_type="solid")
            for col_idx in range(len(COLS)):
                item = self.table.item(row_idx, col_idx)
                val = item.text() if item else ""
                c = ws.cell(row=row_idx + 2, column=col_idx + 1, value=val)
                c.fill = fill
                c.alignment = center
                c.border = border

        widths = [8, 12, 22, 12, 10, 10, 20, 18]
        for i, wdt in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = wdt

        wb.save(path)
        QMessageBox.information(self, "Aktarildi", f"Excel dosyasi kaydedildi:\n{path}")


def main():
    ensure_output_dir()
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    app.setFont(QFont("Segoe UI", 10))
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()