# timetray.py  –  TimeTray v2
# Polished Windows system-tray multi-timezone viewer + converter (PySide6)
#
# Install:  pip install pyside6 tzdata
# Run:      python timetray.py

import sys, os, json, re, math, subprocess
from datetime import datetime, timezone, timedelta
from typing import List, Optional

try:
    from zoneinfo import ZoneInfo, available_timezones
except ImportError:
    print("Need Python 3.9+ and tzdata: pip install tzdata")
    raise

from PySide6.QtCore import Qt, QTimer, QDateTime, QPoint, QSize, QPropertyAnimation, QEasingCurve, QTime
from PySide6.QtGui import (
    QIcon, QAction, QPainter, QPixmap, QColor, QFont, QFontDatabase,
    QCursor, QGuiApplication, QLinearGradient, QPen, QBrush, QPainterPath,
    QKeySequence
)
from PySide6.QtWidgets import (
    QApplication, QSystemTrayIcon, QMenu, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QDateTimeEdit, QTimeEdit, QComboBox, QTableWidget, QTableWidgetItem, QSizePolicy,
    QPushButton, QDialog, QListWidget, QListWidgetItem, QLineEdit,
    QGridLayout, QMessageBox, QFrame, QScrollArea, QSizePolicy,
    QStackedWidget, QCompleter, QAbstractItemView
)

# ──────────────────────────────────────────────
#  Constants
# ──────────────────────────────────────────────
APP_NAME    = "TimeTray"
DEFAULT_TZS = ["Europe/Vilnius", "America/New_York", "Europe/London", "Asia/Dubai"]
DEFAULT_24H = True

# Friendly alias map  →  IANA id
TZ_ALIASES: dict[str, str] = {
    "new york":       "America/New_York",
    "nyc":            "America/New_York",
    "est":            "America/New_York",
    "edt":            "America/New_York",
    "eastern":        "America/New_York",
    "chicago":        "America/Chicago",
    "cst":            "America/Chicago",
    "cdt":            "America/Chicago",
    "central":        "America/Chicago",
    "denver":         "America/Denver",
    "mst":            "America/Denver",
    "mountain":       "America/Denver",
    "los angeles":    "America/Los_Angeles",
    "la":             "America/Los_Angeles",
    "pst":            "America/Los_Angeles",
    "pdt":            "America/Los_Angeles",
    "pacific":        "America/Los_Angeles",
    "toronto":        "America/Toronto",
    "sao paulo":      "America/Sao_Paulo",
    "london":         "Europe/London",
    "gmt":            "Etc/UTC",
    "utc":            "Etc/UTC",
    "paris":          "Europe/Paris",
    "cet":            "Europe/Paris",
    "berlin":         "Europe/Berlin",
    "amsterdam":      "Europe/Amsterdam",
    "madrid":         "Europe/Madrid",
    "rome":           "Europe/Rome",
    "milan":          "Europe/Rome",
    "vilnius":        "Europe/Vilnius",
    "riga":           "Europe/Riga",
    "tallinn":        "Europe/Tallinn",
    "warsaw":         "Europe/Warsaw",
    "stockholm":      "Europe/Stockholm",
    "helsinki":       "Europe/Helsinki",
    "athens":         "Europe/Athens",
    "istanbul":       "Europe/Istanbul",
    "moscow":         "Europe/Moscow",
    "msk":            "Europe/Moscow",
    "dubai":          "Asia/Dubai",
    "uae":            "Asia/Dubai",
    "gulf":           "Asia/Dubai",
    "riyadh":         "Asia/Riyadh",
    "ksa":            "Asia/Riyadh",
    "tel aviv":       "Asia/Jerusalem",
    "cairo":          "Africa/Cairo",
    "nairobi":        "Africa/Nairobi",
    "johannesburg":   "Africa/Johannesburg",
    "lagos":          "Africa/Lagos",
    "karachi":        "Asia/Karachi",
    "pkt":            "Asia/Karachi",
    "mumbai":         "Asia/Kolkata",
    "delhi":          "Asia/Kolkata",
    "india":          "Asia/Kolkata",
    "ist":            "Asia/Kolkata",
    "kolkata":        "Asia/Kolkata",
    "dhaka":          "Asia/Dhaka",
    "bangkok":        "Asia/Bangkok",
    "jakarta":        "Asia/Jakarta",
    "singapore":      "Asia/Singapore",
    "sgt":            "Asia/Singapore",
    "kuala lumpur":   "Asia/Kuala_Lumpur",
    "hong kong":      "Asia/Hong_Kong",
    "hkt":            "Asia/Hong_Kong",
    "beijing":        "Asia/Shanghai",
    "shanghai":       "Asia/Shanghai",
    "cst china":      "Asia/Shanghai",
    "taipei":         "Asia/Taipei",
    "seoul":          "Asia/Seoul",
    "kst":            "Asia/Seoul",
    "tokyo":          "Asia/Tokyo",
    "jst":            "Asia/Tokyo",
    "sydney":         "Australia/Sydney",
    "aest":           "Australia/Sydney",
    "melbourne":      "Australia/Melbourne",
    "auckland":       "Pacific/Auckland",
    "nzst":           "Pacific/Auckland",
    "honolulu":       "Pacific/Honolulu",
    "hst":            "Pacific/Honolulu",
}

TZ_DISPLAY_NAMES: dict[str, str] = {
    "America/New_York":      "New York",
    "America/Chicago":       "Chicago",
    "America/Denver":        "Denver",
    "America/Los_Angeles":   "Los Angeles",
    "America/Toronto":       "Toronto",
    "America/Sao_Paulo":     "São Paulo",
    "Europe/London":         "London",
    "Europe/Paris":          "Paris",
    "Europe/Berlin":         "Berlin",
    "Europe/Amsterdam":      "Amsterdam",
    "Europe/Madrid":         "Madrid",
    "Europe/Rome":           "Rome",
    "Europe/Vilnius":        "Vilnius",
    "Europe/Riga":           "Riga",
    "Europe/Tallinn":        "Tallinn",
    "Europe/Warsaw":         "Warsaw",
    "Europe/Stockholm":      "Stockholm",
    "Europe/Helsinki":       "Helsinki",
    "Europe/Athens":         "Athens",
    "Europe/Istanbul":       "Istanbul",
    "Europe/Moscow":         "Moscow",
    "Asia/Dubai":            "Dubai",
    "Asia/Riyadh":           "Riyadh",
    "Asia/Jerusalem":        "Tel Aviv",
    "Africa/Cairo":          "Cairo",
    "Africa/Nairobi":        "Nairobi",
    "Africa/Johannesburg":   "Johannesburg",
    "Africa/Lagos":          "Lagos",
    "Asia/Karachi":          "Karachi",
    "Asia/Kolkata":          "Mumbai",
    "Asia/Dhaka":            "Dhaka",
    "Asia/Bangkok":          "Bangkok",
    "Asia/Jakarta":          "Jakarta",
    "Asia/Singapore":        "Singapore",
    "Asia/Kuala_Lumpur":     "Kuala Lumpur",
    "Asia/Hong_Kong":        "Hong Kong",
    "Asia/Shanghai":         "Shanghai",
    "Asia/Taipei":           "Taipei",
    "Asia/Seoul":            "Seoul",
    "Asia/Tokyo":            "Tokyo",
    "Australia/Sydney":      "Sydney",
    "Australia/Melbourne":   "Melbourne",
    "Pacific/Auckland":      "Auckland",
    "Pacific/Honolulu":      "Honolulu",
    "Etc/UTC":               "UTC",
}

STYLESHEET = """
QWidget {
    font-family: "Segoe UI", "SF Pro Display", "Helvetica Neue", sans-serif;
}

#popup_root { background: transparent; }
#popup_card {
    background: #16181d;
    border: 1px solid #2a2d35;
    border-radius: 14px;
}

#header { background: transparent; }
#app_title {
    color: #e8eaf0;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 2px;
}
#mode_pill {
    color: #4ade80;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    background: #14532d;
    border-radius: 8px;
    padding: 2px 8px;
}
#mode_pill_frozen {
    color: #fb923c;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    background: #431407;
    border-radius: 8px;
    padding: 2px 8px;
}

#tz_card {
    background: #1e2028;
    border: 1px solid #2a2d35;
    border-radius: 10px;
    padding: 0px;
}
#tz_card:hover {
    background: #22252f;
    border: 1px solid #3d4150;
}
#tz_city {
    color: #8b8fa8;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.8px;
}
#tz_time {
    color: #e8eaf0;
    font-size: 22px;
    font-weight: 300;
    letter-spacing: -0.5px;
}
#tz_date {
    color: #4a4e62;
    font-size: 10px;
}
#tz_offset {
    color: #3d4150;
    font-size: 10px;
    font-weight: 600;
}
#tz_remove_btn {
    color: #3a3d4a;
    font-size: 14px;
    font-weight: 400;
    background: transparent;
    border: none;
    padding: 0px 4px;
    margin: 0px;
    min-width: 20px;
    max-width: 20px;
}
#tz_remove_btn:hover { color: #ef4444; }

#add_section {
    background: #1a1c22;
    border: 1px solid #22252f;
    border-radius: 10px;
}
#search_input {
    background: #22252f;
    color: #e8eaf0;
    border: 1px solid #2a2d35;
    border-radius: 7px;
    padding: 7px 12px;
    font-size: 12px;
    selection-background-color: #3b4fd8;
}
#search_input:focus {
    border: 1px solid #3b4fd8;
    outline: none;
}
#search_input::placeholder { color: #3a3d4a; }
#add_btn {
    background: #2a3080;
    color: #a5b4fc;
    border: none;
    border-radius: 7px;
    padding: 7px 14px;
    font-size: 12px;
    font-weight: 600;
}
#add_btn:hover {
    background: #3b4fd8;
    color: #ffffff;
}
#add_btn:disabled {
    background: #1e2028;
    color: #3a3d4a;
}
#search_hint {
    color: #3a3d4a;
    font-size: 10px;
}

#divider {
    background: #22252f;
    max-height: 1px;
    min-height: 1px;
}

#converter_label {
    color: #4a4e62;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
}
#conv_dt_edit {
    background: #1e2028;
    color: #e8eaf0;
    border: 1px solid #2a2d35;
    border-radius: 7px;
    padding: 5px 10px;
    font-size: 12px;
}
#conv_dt_edit:focus { border: 1px solid #3b4fd8; }
#conv_src_combo {
    background: #1e2028;
    color: #e8eaf0;
    border: 1px solid #2a2d35;
    border-radius: 7px;
    padding: 5px 10px;
    font-size: 12px;
}
#conv_src_combo QAbstractItemView {
    background: #1e2028;
    color: #e8eaf0;
    selection-background-color: #3b4fd8;
    border: 1px solid #2a2d35;
}
#now_btn {
    background: transparent;
    color: #4a4e62;
    border: 1px solid #2a2d35;
    border-radius: 7px;
    padding: 5px 12px;
    font-size: 11px;
    font-weight: 600;
}
#now_btn:hover {
    color: #4ade80;
    border-color: #14532d;
    background: #0f2e1a;
}

QScrollArea { background: transparent; border: none; }
QScrollBar:vertical {
    background: transparent;
    width: 4px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #2a2d35;
    border-radius: 2px;
    min-height: 20px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
"""

# ──────────────────────────────────────────────
#  Helpers
# ──────────────────────────────────────────────

def clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(v, hi))

def appdata_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path

CONFIG_PATH = os.path.join(appdata_dir(), "config.json")

def display_name(iana: str) -> str:
    if iana in TZ_DISPLAY_NAMES:
        return TZ_DISPLAY_NAMES[iana]
    return iana.split("/")[-1].replace("_", " ")

def utc_offset_str(dt: datetime, tz: ZoneInfo) -> str:
    off = dt.astimezone(tz).utcoffset()
    if off is None:
        return "UTC"
    total_min = int(off.total_seconds() // 60)
    sign = "+" if total_min >= 0 else "−"
    total_min = abs(total_min)
    hh, mm = divmod(total_min, 60)
    return f"UTC{sign}{hh}:{mm:02d}" if mm else f"UTC{sign}{hh}"

def fmt_time(dt: datetime, use_24h: bool) -> str:
    return dt.strftime("%H:%M" if use_24h else "%I:%M %p").lstrip("0") if not use_24h else dt.strftime("%H:%M")

def fmt_date(dt: datetime) -> str:
    return dt.strftime("%a, %b %-d") if sys.platform != "win32" else dt.strftime("%a, %b %d").replace(" 0", " ")

def resolve_tz_query(text: str) -> Optional[str]:
    t = text.strip().lower()
    if not t: return None
    if t in TZ_ALIASES: return TZ_ALIASES[t]
    all_tz = available_timezones()
    if text.strip() in all_tz: return text.strip()
    for tz in all_tz:
        if tz.lower() == t: return tz
    for tz in sorted(all_tz):
        if t in tz.lower(): return tz
    return None

def build_autocomplete_list() -> List[str]:
    items = []
    for alias in TZ_ALIASES: items.append(alias.title())
    for iana, friendly in TZ_DISPLAY_NAMES.items():
        if friendly not in items: items.append(friendly)
    return sorted(set(items))

# ──────────────────────────────────────────────
#  Config
# ──────────────────────────────────────────────

class Config:
    def __init__(self):
        self.timezones: List[str] = list(DEFAULT_TZS)
        self.use_24h: bool = DEFAULT_24H

    @classmethod
    def load(cls):
        cfg = cls()
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                tzs = data.get("timezones") or DEFAULT_TZS
                cfg.timezones = tzs
                cfg.use_24h = bool(data.get("use_24h", DEFAULT_24H))
            except Exception:
                pass
        return cfg

    def save(self):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"timezones": self.timezones, "use_24h": self.use_24h}, f, indent=2)

# ──────────────────────────────────────────────
#  Tray icon
# ──────────────────────────────────────────────

def make_tray_icon() -> QIcon:
    size = 128
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)

    p.setBrush(QColor("#1e2028"))
    p.setPen(QPen(QColor("#3b4fd8"), 6))
    p.drawEllipse(4, 4, size - 8, size - 8)

    p.setPen(QPen(QColor("#6366f1"), 4, Qt.SolidLine, Qt.RoundCap))
    cx, cy, r = size // 2, size // 2, size // 2 - 16
    for angle_deg, length_ratio in [(300, 0.45), (60, 0.62)]:
        rad = math.radians(angle_deg - 90)
        ex = cx + int(r * length_ratio * math.cos(rad))
        ey = cy + int(r * length_ratio * math.sin(rad))
        p.drawLine(cx, cy, ex, ey)

    p.setBrush(QColor("#6366f1"))
    p.setPen(Qt.NoPen)
    p.drawEllipse(cx - 5, cy - 5, 10, 10)
    p.end()
    return QIcon(pm)

# ──────────────────────────────────────────────
#  Single timezone card widget
# ──────────────────────────────────────────────

class TZCard(QWidget):
    def __init__(self, iana: str, parent=None):
        super().__init__(parent)
        self.iana = iana
        self.setObjectName("tz_card")
        
        # FIX: Replaced setFixedHeight(58) with a flexible minimum 
        # to ensure text fits regardless of screen scaling.
        self.setMinimumHeight(72)
        self.setCursor(Qt.ArrowCursor)

        self.city_lbl = QLabel(display_name(iana).upper())
        self.city_lbl.setObjectName("tz_city")
        self.time_lbl = QLabel("--:--")
        self.time_lbl.setObjectName("tz_time")
        self.date_lbl = QLabel("")
        self.date_lbl.setObjectName("tz_date")
        self.offset_lbl = QLabel("")
        self.offset_lbl.setObjectName("tz_offset")

        self.remove_btn = QPushButton("×")
        self.remove_btn.setObjectName("tz_remove_btn")
        self.remove_btn.setFixedSize(20, 20)
        self.remove_btn.setToolTip(f"Remove {display_name(iana)}")
        self.remove_btn.hide()
        self.remove_btn.setCursor(Qt.PointingHandCursor)

        left = QVBoxLayout()
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(0)
        left.addWidget(self.city_lbl)
        left.addWidget(self.time_lbl)
        left.addWidget(self.date_lbl)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)
        right.addWidget(self.remove_btn, 0, Qt.AlignTop | Qt.AlignRight)
        right.addStretch(1)
        right.addWidget(self.offset_lbl, 0, Qt.AlignBottom | Qt.AlignRight)

        row = QHBoxLayout(self)
        row.setContentsMargins(14, 8, 10, 8) # slightly increased vertical margins
        row.addLayout(left, 1)
        row.addLayout(right)

    def enterEvent(self, e):
        self.remove_btn.show()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.remove_btn.hide()
        super().leaveEvent(e)

    def update_time(self, src_dt: datetime, use_24h: bool):
        try:
            tz = ZoneInfo(self.iana)
            local = src_dt.astimezone(tz)
            self.time_lbl.setText(fmt_time(local, use_24h))
            self.date_lbl.setText(fmt_date(local))
            self.offset_lbl.setText(utc_offset_str(src_dt, tz))
        except Exception:
            self.time_lbl.setText("--:--")

# ──────────────────────────────────────────────
#  Popup window
# ──────────────────────────────────────────────

class _ClickSelectLineEdit(QLineEdit):
    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        self.selectAll()

class PopupWindow(QWidget):
    def __init__(self, cfg: Config):
        super().__init__()
        self.cfg = cfg
        self._live_mode = True
        self._cards: List[TZCard] = []

        self.setObjectName("popup_root")
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumWidth(380)
        self.setMaximumWidth(420)

        self._build_ui()
        self._rebuild_cards()

        self._tick = QTimer(self)
        self._tick.setInterval(1000)
        self._tick.timeout.connect(self._on_tick)
        self._tick.start()

    def _build_ui(self):
        card = QWidget()
        card.setObjectName("popup_card")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.addWidget(card)

        main = QVBoxLayout(card)
        main.setContentsMargins(0, 0, 0, 12)
        main.setSpacing(0)

        hdr = QHBoxLayout()
        hdr.setContentsMargins(16, 12, 12, 8)
        title = QLabel("TIMETRAY")
        title.setObjectName("app_title")
        self._mode_pill = QLabel("● LIVE")
        self._mode_pill.setObjectName("mode_pill")

        hdr.addWidget(title)
        hdr.addStretch(1)
        hdr.addWidget(self._mode_pill)
        main.addLayout(hdr)

        div0 = QFrame(); div0.setObjectName("divider"); div0.setFixedHeight(1)
        main.addWidget(div0)
        main.addSpacing(6)

        self._cards_container = QWidget()
        self._cards_layout = QVBoxLayout(self._cards_container)
        self._cards_layout.setContentsMargins(10, 0, 10, 0)
        self._cards_layout.setSpacing(4)

        scroll = QScrollArea()
        scroll.setWidget(self._cards_container)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main.addWidget(scroll, 1)

        main.addSpacing(8)

        add_section = QWidget()
        add_section.setObjectName("add_section")
        add_lay = QVBoxLayout(add_section)
        add_lay.setContentsMargins(10, 8, 10, 8)
        add_lay.setSpacing(6)

        search_row = QHBoxLayout()
        self._search = QLineEdit()
        self._search.setObjectName("search_input")
        self._search.setPlaceholderText("Add city or timezone…  e.g. Tokyo, EST")

        completer = QCompleter(build_autocomplete_list(), self._search)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self._search.setCompleter(completer)

        self._add_btn = QPushButton("Add")
        self._add_btn.setObjectName("add_btn")
        self._add_btn.setFixedWidth(52)
        self._add_btn.setCursor(Qt.PointingHandCursor)

        search_row.addWidget(self._search, 1)
        search_row.addWidget(self._add_btn)
        add_lay.addLayout(search_row)

        hint = QLabel("Type a city name, timezone abbrev, or IANA id")
        hint.setObjectName("search_hint")
        add_lay.addWidget(hint)
        main.addWidget(add_section, 0)
        main.addSpacing(8)

        div1 = QFrame(); div1.setObjectName("divider"); div1.setFixedHeight(1)
        main.addWidget(div1)
        main.addSpacing(8)

        conv_lbl = QLabel("TIME CONVERTER")
        conv_lbl.setObjectName("converter_label")
        wrap = QHBoxLayout(); wrap.setContentsMargins(14, 0, 14, 0)
        wrap.addWidget(conv_lbl)
        main.addLayout(wrap)
        main.addSpacing(6)

        conv_row = QHBoxLayout()
        conv_row.setContentsMargins(10, 0, 10, 0)
        conv_row.setSpacing(6)

        self._dt_edit = _ClickSelectLineEdit()
        self._dt_edit.setObjectName("conv_dt_edit")
        self._dt_edit.setPlaceholderText("e.g. 2330")
        self._dt_edit.setFixedWidth(90)
        self._dt_edit.setAlignment(Qt.AlignCenter)
        self._conv_time = QTime.currentTime() 
        self._conv_date_offset = 0

        self._src_combo = QComboBox()
        self._src_combo.setObjectName("conv_src_combo")
        self._src_combo.setEditable(True)
        for iana in sorted(available_timezones()):
            self._src_combo.addItem(iana)
        idx = self._src_combo.findText("Europe/Vilnius")
        if idx >= 0:
            self._src_combo.setCurrentIndex(idx)

        self._now_btn = QPushButton("Now")
        self._now_btn.setObjectName("now_btn")
        self._now_btn.setFixedWidth(48)
        self._now_btn.setCursor(Qt.PointingHandCursor)
        self._now_btn.setToolTip("Reset to current time")

        conv_row.addWidget(self._dt_edit, 3)
        conv_row.addWidget(self._src_combo, 2)
        conv_row.addWidget(self._now_btn)
        main.addLayout(conv_row)

        self._add_btn.clicked.connect(self._on_add)
        self._search.returnPressed.connect(self._on_add)
        self._dt_edit.editingFinished.connect(self._on_time_input)
        self._src_combo.currentTextChanged.connect(self._refresh_all)
        self._now_btn.clicked.connect(self._reset_to_now)

    def _rebuild_cards(self):
        while self._cards_layout.count():
            item = self._cards_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._cards.clear()

        for iana in self.cfg.timezones:
            self._add_card(iana)

        self._cards_layout.addStretch(1)
        self._fit_scroll()
        self._refresh_all()

    def _fit_scroll(self):
        n = len(self._cards)
        card_h = 72 # FIX: Updated scroll area logic to match new card height
        spacing = 4
        padding = 8
        h = n * card_h + max(0, n - 1) * spacing + padding
        h = max(80, min(h, 360))
        scroll = self._cards_container.parent()
        if scroll:
            scroll.setFixedHeight(h)

    def _add_card(self, iana: str):
        card = TZCard(iana)
        card.remove_btn.clicked.connect(lambda _checked=False, i=iana: self._remove_zone(i))
        insert_idx = len(self._cards)
        self._cards_layout.insertWidget(insert_idx, card)
        self._cards.append(card)

    def _remove_zone(self, iana: str):
        if len(self.cfg.timezones) <= 1:
            return
        self.cfg.timezones.remove(iana)
        self.cfg.save()
        self._rebuild_cards()

    def _on_add(self):
        text = self._search.text().strip()
        if not text: return
        iana = resolve_tz_query(text)
        if not iana:
            self._search.setStyleSheet("border: 1px solid #ef4444;")
            QTimer.singleShot(1200, lambda: self._search.setStyleSheet(""))
            return
        if iana in self.cfg.timezones:
            self._search.clear()
            return
        self.cfg.timezones.append(iana)
        self.cfg.save()
        self._add_card(iana)
        self._fit_scroll()
        self._refresh_all()
        self._search.clear()

    def _on_tick(self):
        if self._live_mode:
            self._conv_time = QTime.currentTime()
            self._conv_date_offset = 0
            self._refresh_all()

    def _on_time_input(self):
        raw = self._dt_edit.text().strip().replace(" ", "")
        if not raw:
            self._reset_to_now()
            return
        qt = self._parse_time_input(raw)
        if qt is None or not qt.isValid():
            self._dt_edit.clear()
            self._reset_to_now()
            return
        self._conv_time = qt
        self._dt_edit.setText(qt.toString("HH:mm"))
        
        now = QTime.currentTime()
        diff = abs(now.secsTo(qt))
        if diff > 120:
            self._live_mode = False
            self._mode_pill.setObjectName("mode_pill_frozen")
            self._mode_pill.setText("● FROZEN")
            self._mode_pill.setStyleSheet("")
        self._refresh_all()

    @staticmethod
    def _parse_time_input(raw: str):
        if not raw: return None
        raw = raw.lower().strip()
        pm = raw.endswith("pm")
        am = raw.endswith("am")
        if pm or am: raw = raw[:-2].strip().rstrip(":")
        digits = raw.replace(":", "").replace(".", "")

        if re.fullmatch(r'\d{1,4}', digits):
            n = int(digits)
            if len(digits) <= 2:
                h, m = n, 0
            elif len(digits) == 3:
                h, m = int(digits[0]), int(digits[1:])
            else:
                h, m = int(digits[:2]), int(digits[2:])
            if pm and h != 12: h += 12
            if am and h == 12: h = 0
            return QTime(h % 24, m % 60)
        return None

    def _reset_to_now(self):
        self._live_mode = True
        self._conv_time = QTime.currentTime()
        self._conv_date_offset = 0
        self._mode_pill.setObjectName("mode_pill")
        self._mode_pill.setText("● LIVE")
        self._mode_pill.setStyleSheet("")
        self._dt_edit.blockSignals(True)
        self._dt_edit.clear()
        self._dt_edit.blockSignals(False)
        self._refresh_all()

    def _refresh_all(self):
        if self._live_mode:
            # LIVE MODE: Ignore the converter box. Use the true, current universal time.
            src_dt = datetime.now(timezone.utc)
        else:
            # CONVERTER MODE: Build the custom time based on the textbox and dropdown.
            src_tz_id = self._src_combo.currentText()
            
            # Resolve aliases like "EST" or "nyc" to proper IANA zones ("America/New_York")
            real_iana = resolve_tz_query(src_tz_id) or src_tz_id
            
            try:
                src_tz = ZoneInfo(real_iana)
            except Exception:
                return

            today = datetime.now().date()
            base_date = today + timedelta(days=self._conv_date_offset)
            naive = datetime(base_date.year, base_date.month, base_date.day,
                             self._conv_time.hour(), self._conv_time.minute(), 0)
            src_dt = naive.replace(tzinfo=src_tz)

        # Update all cards with the calculated time
        for card in self._cards:
            card.update_time(src_dt, self.cfg.use_24h)

    def set_time_format(self, use_24h: bool):
        self.cfg.use_24h = use_24h
        self._refresh_all()

# ──────────────────────────────────────────────
#  Tray application
# ──────────────────────────────────────────────

class TrayApp:
    def __init__(self, app: QApplication):
        self.app = app
        self.cfg = Config.load()

        icon = make_tray_icon()
        self.tray = QSystemTrayIcon(icon, parent=app)
        self.tray.setToolTip("TimeTray — timezone viewer")

        menu = QMenu()
        self._open_act  = QAction("Open TimeTray")
        self._fmt_act   = QAction("Use 12-hour format" if self.cfg.use_24h else "Use 24-hour format")
        self._exit_act  = QAction("Exit")
        
        menu.addAction(self._open_act)
        menu.addSeparator()
        menu.addAction(self._fmt_act)
        menu.addSeparator()
        menu.addAction(self._exit_act)
        self.tray.setContextMenu(menu)

        self.popup = PopupWindow(self.cfg)
        self.popup.setStyleSheet(STYLESHEET)

        self._open_act.triggered.connect(self.show_popup)
        self._fmt_act.triggered.connect(self._toggle_fmt)
        self._exit_act.triggered.connect(self._quit)
        self.tray.activated.connect(self._tray_activated)
        self.tray.show()

    def _tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_popup()

    def show_popup(self):
        if self.popup.isVisible():
            self.popup.hide()
            return
            
        anchor = QCursor.pos()
        tr = self.tray.geometry()
        if tr.isValid():
            anchor = tr.center()

        screen = QGuiApplication.screenAt(anchor) or QGuiApplication.primaryScreen()
        geo = screen.availableGeometry()

        popup_h = min(560, geo.height() - 16)
        popup_w = 420
        self.popup.setFixedSize(popup_w, popup_h)

        x = clamp(anchor.x() - popup_w // 2, geo.left(), geo.right() - popup_w)
        y_above = anchor.y() - popup_h - 8
        y_below = anchor.y() + 8
        
        if y_above >= geo.top():
            y = y_above
        else:
            y = clamp(y_below, geo.top(), geo.bottom() - popup_h)

        self.popup.move(x, y)
        self.popup._refresh_all()
        self.popup.show()
        self.popup.raise_()
        self.popup.activateWindow()

    def _toggle_fmt(self):
        self.cfg.use_24h = not self.cfg.use_24h
        self.cfg.save()
        self._fmt_act.setText("Use 12-hour format" if self.cfg.use_24h else "Use 24-hour format")
        self.popup.set_time_format(self.cfg.use_24h)

    def _quit(self):
        self.tray.hide()
        self.app.quit()

# ──────────────────────────────────────────────
#  Global hotkey  (Windows only, Alt+T)
# ──────────────────────────────────────────────

_HOTKEY_ID = 1
_MOD_ALT   = 0x0001
_VK_T      = 0x54

def _register_hotkey(hwnd: int) -> bool:
    try:
        import ctypes
        return bool(ctypes.windll.user32.RegisterHotKey(hwnd, _HOTKEY_ID, _MOD_ALT, _VK_T))
    except Exception:
        return False

def _unregister_hotkey(hwnd: int):
    try:
        import ctypes
        ctypes.windll.user32.UnregisterHotKey(hwnd, _HOTKEY_ID)
    except Exception:
        pass

class HotkeyReceiver(QWidget):
    def __init__(self, callback):
        super().__init__()
        self._cb = callback
        self.setFixedSize(0, 0)
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)
        QTimer.singleShot(200, self._register)

    def _register(self):
        try:
            ok = _register_hotkey(int(self.winId()))
            if not ok: print("TimeTray: Alt+T hotkey unavailable (already in use?)")
        except Exception as e:
            print(f"TimeTray: hotkey skipped ({e})")

    def nativeEvent(self, event_type, message):
        try:
            import ctypes, ctypes.wintypes
            msg = ctypes.wintypes.MSG.from_address(int(message))
            if msg.message == 0x0312 and msg.wParam == _HOTKEY_ID:
                self._cb()
                return True, 0
        except Exception:
            pass
        return False, 0

    def closeEvent(self, e):
        try:
            _unregister_hotkey(int(self.winId()))
        except Exception:
            pass
        super().closeEvent(e)

# ──────────────────────────────────────────────
#  Windows startup registration
# ──────────────────────────────────────────────

def _startup_vbs_path() -> str:
    startup_dir = os.path.join(
        os.environ.get("APPDATA", os.path.expanduser("~")),
        r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    return os.path.join(startup_dir, "TimeTray.vbs")

def _ensure_startup():
    vbs_path = _startup_vbs_path()
    if os.path.exists(vbs_path): return
    
    script_path = os.path.abspath(__file__)
    python_exe = sys.executable
    pythonw = python_exe.replace("python.exe", "pythonw.exe")
    if os.path.exists(pythonw): python_exe = pythonw
        
    vbs = (
        'Set WshShell = CreateObject("WScript.Shell")\n'
        f'WshShell.Run Chr(34) & "{python_exe}" & Chr(34) & " " & Chr(34) & "{script_path}" & Chr(34), 0, False\n'
    )
    try:
        os.makedirs(os.path.dirname(vbs_path), exist_ok=True)
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs)
        print(f"TimeTray: startup entry created → {vbs_path}")
    except Exception as e:
        print(f"TimeTray: could not create startup entry ({e})")

# ──────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setApplicationName(APP_NAME)

    _ensure_startup()
    tray_app = TrayApp(app)

    if sys.platform == "win32":
        _hotkey_win = HotkeyReceiver(tray_app.show_popup)
        _hotkey_win.show()
        _hotkey_win.hide()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()