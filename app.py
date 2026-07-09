#!/usr/bin/env python3
"""
Window Rescue — Visual Multi-Monitor Window Manager
====================================================
See ALL your windows on a mini-map. Rescue the ones stuck off-screen.

  pip install PyQt5 pywin32 Pillow
  python window_rescue.py

Features:
  • Mini-map with real window thumbnails (not just colored boxes)
  • Green border = on-screen  |  Red border = off-screen
  • Drag any window thumbnail → real window moves with it (live!)
  • Double-click a red window → snaps it to Monitor 1
  • Right-click → context menu (rescue to specific monitor, focus, info)
  • "Rescue All" button → brings every off-screen window back at once
  • Hover tooltip → full title, exact coordinates, size
  • Auto-refreshes every 3 seconds
  • System tray icon → minimize to tray, stays running in background
"""

import sys
import platform
import ctypes
import ctypes.wintypes
from typing import List, Optional, Dict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QGraphicsView, QGraphicsScene, QGraphicsObject,
    QGraphicsItem, QSystemTrayIcon, QMenu, QAction, QMessageBox,
    QFrame, QStatusBar, QToolTip, QSizePolicy,
)
from PyQt5.QtCore import Qt, QRectF, QPointF, QTimer, QThread, pyqtSignal, QSize
from PyQt5.QtGui import (
    QPainter, QColor, QPen, QBrush, QFont, QFontMetrics,
    QPixmap, QImage, QIcon, QCursor, QPainterPath, QLinearGradient,
)

import win32gui
import win32ui
import win32api
import win32con
from PIL import Image

# ── Win32 constants ───────────────────────────────────────────────────────────
SW_RESTORE       = 9
SWP_NOSIZE       = 0x0001
SWP_NOZORDER     = 0x0004
SWP_NOACTIVATE   = 0x0010
SWP_SHOWWINDOW   = 0x0040
GWL_EXSTYLE      = win32con.GWL_EXSTYLE
WS_EX_TOOLWINDOW = win32con.WS_EX_TOOLWINDOW
WS_EX_APPWINDOW  = win32con.WS_EX_APPWINDOW

_u32 = ctypes.windll.user32

# ── Colors ────────────────────────────────────────────────────────────────────
BG          = "#1e1e2e"
SURFACE     = "#181825"
OVERLAY     = "#313244"
MUTED       = "#45475a"
TEXT        = "#cdd6f4"
SUBTEXT     = "#a6adc8"
MON_FILL    = "#1c1c2e"
MON_BORDER  = "#585b70"
GREEN       = "#40a02b"
GREEN_L     = "#a6e3a1"
RED         = "#d20f39"
RED_L       = "#f38ba8"
ORANGE      = "#fe640b"
BLUE        = "#89b4fa"


# ── Data models ───────────────────────────────────────────────────────────────
class Monitor:
    def __init__(self, x: int, y: int, w: int, h: int):
        self.x, self.y, self.w, self.h = x, y, w, h

    @property
    def right(self):  return self.x + self.w
    @property
    def bottom(self): return self.y + self.h


class WinInfo:
    def __init__(self, hwnd: int, title: str, l: int, t: int, r: int, b: int,
                 minimized: bool = False, z_order: int = 0):
        self.hwnd  = hwnd
        self.title = title
        self.left, self.top, self.right, self.bottom = l, t, r, b
        self.is_off       = False
        self.is_minimized = minimized
        self.z_order      = z_order   # 0 = topmost (foreground)
        self.thumbnail: Optional[QPixmap] = None

    @property
    def w(self): return self.right - self.left
    @property
    def h(self): return self.bottom - self.top


# ── Win32 helpers ─────────────────────────────────────────────────────────────
def get_monitors() -> List[Monitor]:
    """Use win32api for reliable multi-monitor enumeration."""
    result: List[Monitor] = []
    for _hmon, _hdc, rect in win32api.EnumDisplayMonitors():
        l, t, r, b = rect
        result.append(Monitor(l, t, r - l, b - t))
    return result or [Monitor(0, 0, 1920, 1080)]


# Window classes that are always system UI — never show these
_SKIP_CLASSES = {
    "Shell_TrayWnd",          # taskbar
    "Shell_SecondaryTrayWnd", # secondary taskbar
    "DV2ControlHost",         # magnifier
    "MsgrIMEWindowClass",
    "SysShadow",
    "progman",                # desktop
    "WorkerW",                # desktop worker
    "Windows.UI.Core.CoreWindow",  # system UWP shell chrome
    "TaskManagerWindow",
}


def _accept_window(hwnd: int) -> bool:
    """Quick visibility / class check shared by both normal and minimized paths."""
    if not win32gui.IsWindowVisible(hwnd):
        return False
    title = win32gui.GetWindowText(hwnd).strip()
    if not title:
        return False
    try:
        cls = win32gui.GetClassName(hwnd)
        if cls.lower() in {c.lower() for c in _SKIP_CLASSES}:
            return False
    except Exception:
        pass
    return True


def get_windows() -> List[WinInfo]:
    """
    Walk windows in true Z-order (front → back) using GetTopWindow /
    GW_HWNDNEXT so the minimap renders them in the correct stacking order.
    z_order=0 is the topmost (foreground) window.
    """
    wins: List[WinInfo] = []
    _skip_cls_lower = {c.lower() for c in _SKIP_CLASSES}
    z = 0

    try:
        hwnd = win32gui.GetTopWindow(None)   # foreground-most window
    except Exception:
        hwnd = 0

    while hwnd:
        try:
            if _accept_window(hwnd):
                title = win32gui.GetWindowText(hwnd).strip()

                if win32gui.IsIconic(hwnd):
                    # Minimized — use restored rect from GetWindowPlacement
                    try:
                        pl = win32gui.GetWindowPlacement(hwnd)
                        l, t, r, b = pl[4]
                        if (r - l) >= 100 and (b - t) >= 50:
                            wins.append(WinInfo(hwnd, title, l, t, r, b,
                                                minimized=True, z_order=z))
                            z += 1
                    except Exception:
                        pass
                else:
                    l, t, r, b = win32gui.GetWindowRect(hwnd)
                    if (r - l) >= 50 and (b - t) >= 30:
                        wins.append(WinInfo(hwnd, title, l, t, r, b,
                                            z_order=z))
                        z += 1
        except Exception:
            pass

        try:
            hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
        except Exception:
            break

    return wins


def find_offscreen_deep(monitors: List[Monitor]) -> List[WinInfo]:
    """
    Deep scan: enumerate ALL top-level windows regardless of visibility,
    and return only those positioned entirely outside every monitor.
    Catches Citrix/RDP seamless dialogs that have WS_VISIBLE=False locally
    but are rendered off-screen on the remote session.
    """
    found: List[WinInfo] = []
    seen_hwnds: set = set()

    def cb(hwnd, _):
        try:
            if hwnd in seen_hwnds:
                return True
            title = win32gui.GetWindowText(hwnd).strip()
            if not title:
                return True
            cls = ""
            try:
                cls = win32gui.GetClassName(hwnd)
            except Exception:
                pass
            if cls.lower() in {c.lower() for c in _SKIP_CLASSES}:
                return True
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            if (r - l) < 50 or (b - t) < 30:
                return True
            # Only care about windows that are fully off every monitor
            dummy = WinInfo(hwnd, title, l, t, r, b)
            if not overlaps_monitor(dummy, monitors):
                seen_hwnds.add(hwnd)
                found.append(dummy)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return found


def overlaps_monitor(w: WinInfo, monitors: List[Monitor]) -> bool:
    for m in monitors:
        if w.left < m.right and w.right > m.x and w.top < m.bottom and w.bottom > m.y:
            return True
    return False


def capture_thumbnail(hwnd: int) -> Optional[QPixmap]:
    """Capture a window's content as a QPixmap using PrintWindow."""
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w, h = r - l, b - t
        if w <= 0 or h <= 0 or w > 9000 or h > 9000:
            return None

        hwnd_dc  = win32gui.GetWindowDC(hwnd)
        mfc_dc   = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc  = mfc_dc.CreateCompatibleDC()
        bmp      = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)

        # PW_RENDERFULLCONTENT (2) works for most modern apps incl. hardware-accel
        ok = _u32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

        info   = bmp.GetInfo()
        bits   = bmp.GetBitmapBits(True)

        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

        if not ok:
            return None

        img  = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                                bits, "raw", "BGRX", 0, 1)
        data = img.tobytes("raw", "RGB")
        qi   = QImage(data, img.width, img.height, img.width * 3, QImage.Format_RGB888)
        return QPixmap.fromImage(qi)

    except Exception:
        return None


def do_rescue(hwnd: int, m: Monitor, offset: int = 0):
    _u32.ShowWindow(hwnd, SW_RESTORE)
    _u32.SetWindowPos(hwnd, 0,
                      m.x + 50 + offset, m.y + 50 + offset, 0, 0,
                      SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW)


def do_move(hwnd: int, x: int, y: int):
    _u32.SetWindowPos(hwnd, 0, x, y, 0, 0,
                      SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)


def do_focus(hwnd: int):
    _u32.ShowWindow(hwnd, SW_RESTORE)
    _u32.SetForegroundWindow(hwnd)


# ── Background thumbnail capture thread ───────────────────────────────────────
class ThumbThread(QThread):
    ready = pyqtSignal(int, QPixmap)   # hwnd, pixmap

    def __init__(self, hwnds: List[int]):
        super().__init__()
        self.hwnds = hwnds

    def run(self):
        for hwnd in self.hwnds:
            px = capture_thumbnail(hwnd)
            if px:
                self.ready.emit(hwnd, px)


# ── Window graphics item ───────────────────────────────────────────────────────
class WinItem(QGraphicsObject):
    """Draggable, thumbnail-showing item in the scene (scene coords = real pixels)."""

    rescueRequested = pyqtSignal(int)        # hwnd
    focusRequested  = pyqtSignal(int)        # hwnd
    moveFinished    = pyqtSignal(int, int, int)  # hwnd, x, y

    def __init__(self, win: WinInfo):
        super().__init__()
        self.win = win
        self.setPos(win.left, win.top)
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)
        self.setAcceptHoverEvents(True)
        self.setCursor(Qt.SizeAllCursor)
        self.setZValue(2)          # overridden in _rebuild_scene with real Z-order
        self._hovered       = False
        self._is_foreground = False

    # ── Geometry ──────────────────────────────────────────────────────────────
    def boundingRect(self) -> QRectF:
        return QRectF(0, 0, self.win.w, self.win.h)

    def update_thumbnail(self, px: QPixmap):
        self.win.thumbnail = px
        self.update()

    # ── Paint ─────────────────────────────────────────────────────────────────
    def paint(self, painter: QPainter, option, widget):
        w, h   = self.win.w, self.win.h
        rect   = QRectF(0, 0, w, h)
        is_off = self.win.is_off

        border_col = QColor(RED   if is_off else GREEN)
        fill_col   = QColor(RED_L if is_off else GREEN_L)
        fill_col.setAlpha(60 if self.win.thumbnail else 180)

        if self._hovered:
            border_col = border_col.lighter(150)

        # Thumbnail or fallback fill
        if self.win.thumbnail:
            painter.drawPixmap(rect.toRect(), self.win.thumbnail)
            # Tinted overlay so color-coding stays visible
            overlay = QColor(RED if is_off else GREEN)
            overlay.setAlpha(55)
            painter.fillRect(rect, overlay)
        else:
            painter.fillRect(rect, fill_col)

        # Border
        bw = 3 if self._hovered else 2
        painter.setPen(QPen(border_col, bw))
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(QRectF(bw / 2, bw / 2, w - bw, h - bw))

        # Title bar strip
        bar_h = min(26, h // 5)
        bar_color = QColor(RED if is_off else GREEN)
        bar_color.setAlpha(210)
        painter.fillRect(QRectF(0, 0, w, bar_h), bar_color)

        font = QFont("Segoe UI", max(7, min(9, int(bar_h * 0.6))), QFont.Bold)
        painter.setFont(font)
        painter.setPen(QColor("white"))
        fm    = QFontMetrics(font)
        title = fm.elidedText(self.win.title, Qt.ElideRight, int(w - 10))
        painter.drawText(QRectF(5, 0, w - 10, bar_h), Qt.AlignVCenter | Qt.AlignLeft, title)

        # Foreground window — gold glow border
        if self._is_foreground and not is_off:
            glow = QPen(QColor("#f9e2af"), 4)
            painter.setPen(glow)
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(QRectF(2, 2, w - 4, h - 4))

        # Off-screen "!" badge
        if is_off:
            bs = min(20, w // 5, h // 5)
            bx, by = w - bs - 4, bar_h + 4
            painter.setBrush(QColor(RED))
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(QRectF(bx, by, bs, bs))
            f2 = QFont("Segoe UI", max(7, bs // 2), QFont.Bold)
            painter.setFont(f2)
            painter.setPen(QColor("white"))
            painter.drawText(QRectF(bx, by, bs, bs), Qt.AlignCenter, "!")

    # ── Hover tooltip ─────────────────────────────────────────────────────────
    def hoverEnterEvent(self, e):
        self._hovered = True
        self.update()
        w = self.win
        status = f"<font color='{RED_L}'>OFF-SCREEN ⚠</font>" if w.is_off else f"<font color='{GREEN_L}'>On-screen ✓</font>"
        QToolTip.showText(
            QCursor.pos(),
            f"<b>{w.title}</b><br>"
            f"Status&nbsp;&nbsp;&nbsp;: {status}<br>"
            f"Position : ({w.left}, {w.top})<br>"
            f"Size&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: {w.w} × {w.h} px",
        )

    def hoverLeaveEvent(self, e):
        self._hovered = False
        self.update()
        QToolTip.hideText()

    # ── Drag — move real window live ──────────────────────────────────────────
    def mouseMoveEvent(self, e):
        super().mouseMoveEvent(e)   # QGraphicsItem handles position update
        pos = self.scenePos()
        do_move(self.win.hwnd, int(pos.x()), int(pos.y()))

    def mouseReleaseEvent(self, e):
        super().mouseReleaseEvent(e)
        pos = self.scenePos()
        nx, ny = int(pos.x()), int(pos.y())
        # Sync win info
        ww, wh       = self.win.w, self.win.h
        self.win.left, self.win.top     = nx, ny
        self.win.right, self.win.bottom = nx + ww, ny + wh
        do_move(self.win.hwnd, nx, ny)
        self.moveFinished.emit(self.win.hwnd, nx, ny)

    def mouseDoubleClickEvent(self, e):
        if self.win.is_off:
            self.rescueRequested.emit(self.win.hwnd)
        else:
            self.focusRequested.emit(self.win.hwnd)


# ── Monitor graphics item ──────────────────────────────────────────────────────
class MonItem(QGraphicsObject):
    def __init__(self, mon: Monitor, index: int):
        super().__init__()
        self.mon   = mon
        self.index = index
        self.setPos(mon.x, mon.y)
        self.setZValue(0)

    def boundingRect(self) -> QRectF:
        return QRectF(0, 0, self.mon.w, self.mon.h)

    def paint(self, painter: QPainter, option, widget):
        w, h = self.mon.w, self.mon.h
        rect = QRectF(0, 0, w, h)

        painter.fillRect(rect, QColor(MON_FILL))
        painter.setPen(QPen(QColor(MON_BORDER), 3))
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(rect)

        font = QFont("Segoe UI", 11)
        painter.setFont(font)
        painter.setPen(QColor(SUBTEXT))
        painter.drawText(rect, Qt.AlignCenter,
                         f"Monitor {self.index + 1}\n{self.mon.w} × {self.mon.h}")


# ── Minimized windows tray panel ──────────────────────────────────────────────
class MinimizedBar(QWidget):
    """Horizontal panel showing minimized windows as clickable tiles."""

    restore_requested = pyqtSignal(int)   # hwnd

    # One color per tile (cycles)
    _COLORS = ["#89b4fa", "#a6e3a1", "#fab387", "#f38ba8",
               "#cba6f7", "#94e2d5", "#f9e2af", "#74c7ec"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("minBar")
        self.setFixedHeight(52)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(14, 6, 14, 6)
        outer.setSpacing(0)

        lbl = QLabel("Minimized")
        lbl.setObjectName("minLabel")
        outer.addWidget(lbl)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setObjectName("minSep")
        outer.addWidget(sep)

        # Scrollable tile area
        from PyQt5.QtWidgets import QScrollArea
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setFixedHeight(40)

        self._inner = QWidget()
        self._inner.setObjectName("minInner")
        self._tile_layout = QHBoxLayout(self._inner)
        self._tile_layout.setContentsMargins(4, 0, 4, 0)
        self._tile_layout.setSpacing(6)
        self._tile_layout.addStretch()

        self._scroll.setWidget(self._inner)
        outer.addWidget(self._scroll, stretch=1)

        self._tiles: list = []
        self.hide()

    def update_windows(self, wins: List["WinInfo"]):
        # Clear old tiles
        for t in self._tiles:
            t.deleteLater()
        self._tiles.clear()
        # Remove old stretch
        item = self._tile_layout.takeAt(self._tile_layout.count() - 1)
        del item

        minimized = [w for w in wins if w.is_minimized]

        if not minimized:
            self._tile_layout.addStretch()
            self.hide()
            return

        for i, w in enumerate(minimized):
            color = self._COLORS[i % len(self._COLORS)]
            btn = QPushButton()
            btn.setObjectName("minTile")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip(f"{w.title}\nClick to restore")
            btn.setFixedHeight(30)

            # Icon circle + title
            initial = (w.title[0].upper()) if w.title else "?"
            label   = w.title[:22] + ("…" if len(w.title) > 22 else "")
            btn.setText(f"  {initial}  {label}")
            btn.setStyleSheet(
                f"QPushButton#{btn.objectName()} {{"
                f"  background: {OVERLAY}; color: {color};"
                f"  border: 1px solid {color}; border-radius: 4px;"
                f"  padding: 0 10px; font-size: 9px; font-weight: bold;"
                f"  text-align: left;"
                f"}}"
                f"QPushButton#{btn.objectName()}:hover {{"
                f"  background: {color}; color: {SURFACE};"
                f"}}"
            )
            hwnd = w.hwnd
            btn.clicked.connect(lambda _=False, h=hwnd: self.restore_requested.emit(h))
            self._tile_layout.addWidget(btn)
            self._tiles.append(btn)

        self._tile_layout.addStretch()
        self.show()


# ── Graphics view with background ─────────────────────────────────────────────
class MapView(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.NoDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setBackgroundBrush(QBrush(QColor(BG)))
        self.setFrameStyle(QFrame.NoFrame)

    def wheelEvent(self, e):
        factor = 1.15 if e.angleDelta().y() > 0 else 1 / 1.15
        self.scale(factor, factor)

    def fit_all(self):
        if self.scene() and self.scene().items():
            self.fitInView(self.scene().itemsBoundingRect().adjusted(-60, -60, 60, 60),
                           Qt.KeepAspectRatio)


# ── Main window ───────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Window Rescue")
        self.resize(1000, 620)
        self.setMinimumSize(620, 400)

        self.monitors: List[Monitor] = []
        self.windows:  List[WinInfo] = []
        self._items:   Dict[int, WinItem] = {}   # hwnd → WinItem
        self._deep_extras: Dict[int, WinInfo] = {}  # hwnd → WinInfo from deep scan
        self._thumb_thread: Optional[ThumbThread] = None

        self._build_ui()
        self._setup_tray()
        self._apply_style()

        QTimer.singleShot(150, self.full_refresh)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._auto_refresh)
        self._timer.start(3500)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header
        hdr = QWidget()
        hdr.setObjectName("header")
        hdr_lay = QHBoxLayout(hdr)
        hdr_lay.setContentsMargins(16, 10, 16, 10)

        title = QLabel("Window Rescue")
        title.setObjectName("appTitle")
        sub   = QLabel("visual multi-monitor window manager")
        sub.setObjectName("appSub")
        hdr_lay.addWidget(title)
        hdr_lay.addWidget(sub)
        hdr_lay.addStretch()

        self._rescue_btn = QPushButton("⚠  Rescue All Off-Screen")
        self._rescue_btn.setObjectName("rescueBtn")
        self._rescue_btn.clicked.connect(self.rescue_all)
        self._rescue_btn.setCursor(Qt.PointingHandCursor)

        deep_btn = QPushButton("🔍  Find Off-Screen")
        deep_btn.setObjectName("deepBtn")
        deep_btn.clicked.connect(self.deep_scan)
        deep_btn.setCursor(Qt.PointingHandCursor)
        deep_btn.setToolTip(
            "Deep scan: finds ALL off-screen windows including\n"
            "Citrix / RDP seamless dialogs that are hidden to normal detection."
        )

        refresh_btn = QPushButton("⟳  Refresh")
        refresh_btn.setObjectName("refreshBtn")
        refresh_btn.clicked.connect(self.full_refresh)
        refresh_btn.setCursor(Qt.PointingHandCursor)

        hdr_lay.addWidget(self._rescue_btn)
        hdr_lay.addWidget(deep_btn)
        hdr_lay.addWidget(refresh_btn)
        root.addWidget(hdr)

        # Map view
        self._scene = QGraphicsScene()
        self._view  = MapView()
        self._view.setScene(self._scene)
        root.addWidget(self._view, stretch=1)

        # Minimized windows tray
        self._min_bar = MinimizedBar()
        self._min_bar.restore_requested.connect(self._on_restore_minimized)
        root.addWidget(self._min_bar)

        # Status bar
        self._status = QLabel()
        self._status.setObjectName("statusBar")
        self._status.setContentsMargins(16, 6, 16, 6)
        root.addWidget(self._status)

    def _setup_tray(self):
        self._tray = QSystemTrayIcon(self)
        # Use a simple colored icon
        px = QPixmap(16, 16)
        px.fill(QColor(BLUE))
        self._tray.setIcon(QIcon(px))
        self._tray.setToolTip("Window Rescue")

        menu = QMenu()
        menu.addAction("Show", self.show)
        menu.addAction("Rescue All", self.rescue_all)
        menu.addSeparator()
        menu.addAction("Quit", QApplication.quit)
        self._tray.setContextMenu(menu)
        self._tray.activated.connect(
            lambda reason: self.show() if reason == QSystemTrayIcon.DoubleClick else None
        )
        self._tray.show()

    def _apply_style(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{ background: {BG}; color: {TEXT}; font-family: 'Segoe UI'; }}
            #header {{ background: {SURFACE}; border-bottom: 1px solid {MUTED}; }}
            #appTitle {{ color: {BLUE}; font-size: 14px; font-weight: bold; }}
            #appSub   {{ color: {SUBTEXT}; font-size: 9px; margin-left: 10px; }}
            QPushButton {{
                border: none; border-radius: 5px;
                padding: 6px 16px; font-size: 9px; font-weight: bold;
                color: {TEXT};
            }}
            #rescueBtn       {{ background: {MUTED}; }}
            #rescueBtn:hover {{ background: {RED}; }}
            #deepBtn         {{ background: #313244; color: {BLUE}; border: 1px solid {BLUE}; }}
            #deepBtn:hover   {{ background: {BLUE}; color: {SURFACE}; }}
            #refreshBtn {{ background: {BLUE}; }}
            #refreshBtn:hover {{ background: #74c7ec; }}
            #statusBar {{
                background: {SURFACE}; color: {SUBTEXT}; font-size: 8px;
                border-top: 1px solid {MUTED};
            }}
            QMenu {{ background: {OVERLAY}; color: {TEXT}; border: 1px solid {MUTED}; }}
            QMenu::item:selected {{ background: {MUTED}; }}
            QToolTip {{ background: {OVERLAY}; color: {TEXT}; border: 1px solid {MUTED}; padding: 4px; }}
            #minBar {{
                background: {SURFACE}; border-top: 1px solid {MUTED};
                border-bottom: 1px solid {MUTED};
            }}
            #minLabel {{
                color: {SUBTEXT}; font-size: 8px; font-weight: bold;
                padding: 0 10px 0 0; text-transform: uppercase; letter-spacing: 1px;
            }}
            #minSep {{ color: {MUTED}; margin: 4px 8px; }}
            #minInner {{ background: transparent; }}
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:horizontal {{
                background: {OVERLAY}; height: 4px; border-radius: 2px;
            }}
            QScrollBar::handle:horizontal {{
                background: {MUTED}; border-radius: 2px;
            }}
        """)

    # ── Refresh ───────────────────────────────────────────────────────────────
    def full_refresh(self):
        self.monitors = get_monitors()
        self.windows  = get_windows()
        for w in self.windows:
            w.is_off = not overlaps_monitor(w, self.monitors)

        self._rebuild_scene()
        self._min_bar.update_windows(self.windows)
        self._update_status()
        self._start_thumb_capture()

    def _auto_refresh(self):
        # Lightweight refresh: just re-enumerate windows, keep thumbnails
        old_thumbs = {w.hwnd: w.thumbnail for w in self.windows if w.thumbnail}
        self.monitors = get_monitors()
        self.windows  = get_windows()
        known = {w.hwnd for w in self.windows}
        for w in self.windows:
            w.is_off      = not overlaps_monitor(w, self.monitors)
            w.thumbnail   = old_thumbs.get(w.hwnd)

        # Re-merge deep scan extras that are still off-screen
        for hwnd, extra in list(self._deep_extras.items()):
            if hwnd not in known:
                # Re-check if still off-screen
                try:
                    l, t, r, b = win32gui.GetWindowRect(hwnd)
                    extra.left, extra.top = l, t
                    extra.right, extra.bottom = r, b
                except Exception:
                    del self._deep_extras[hwnd]
                    continue
                if not overlaps_monitor(extra, self.monitors):
                    extra.thumbnail = old_thumbs.get(hwnd)
                    self.windows.append(extra)
                else:
                    del self._deep_extras[hwnd]  # it came back on-screen

        self._rebuild_scene()
        self._min_bar.update_windows(self.windows)
        self._update_status()

        # Refresh thumbnails for new windows only
        new_hwnds = [w.hwnd for w in self.windows if not w.thumbnail]
        if new_hwnds:
            self._start_thumb_capture(new_hwnds)

    def _rebuild_scene(self):
        self._scene.clear()
        self._items.clear()

        # Monitors
        for i, m in enumerate(self.monitors):
            self._scene.addItem(MonItem(m, i))

        # Windows (skip minimized — they go in the tray bar below)
        # z_order=0 is foreground; highest scene Z = most on top
        visible = [w for w in self.windows if not w.is_minimized]
        max_z   = max((w.z_order for w in visible), default=0)

        fg_hwnd = 0
        try:
            fg_hwnd = win32gui.GetForegroundWindow()
        except Exception:
            pass

        for w in visible:
            item = WinItem(w)
            # Scene Z: invert z_order so foreground window renders on top
            item.setZValue(10 + max_z - w.z_order)
            # Gold outline for the active foreground window
            if w.hwnd == fg_hwnd:
                item.setZValue(10 + max_z + 1)   # always on very top
                item._is_foreground = True
            item.rescueRequested.connect(self._on_rescue_request)
            item.focusRequested.connect(self._on_focus_request)
            item.moveFinished.connect(self._on_move_finished)
            self._scene.addItem(item)
            self._items[w.hwnd] = item
            item.setFlag(QGraphicsItem.ItemIsSelectable, True)

        # Defer fit so Qt processes new scene items before we measure bounds
        QTimer.singleShot(50, self._view.fit_all)
        # Connect scene right-click
        self._scene.setContextMenuPolicy = lambda: None  # handled in view

    def _start_thumb_capture(self, hwnds: Optional[List[int]] = None):
        if self._thumb_thread and self._thumb_thread.isRunning():
            return
        target = hwnds or [w.hwnd for w in self.windows]
        self._thumb_thread = ThumbThread(target)
        self._thumb_thread.ready.connect(self._on_thumb_ready)
        self._thumb_thread.start()

    def _on_thumb_ready(self, hwnd: int, px: QPixmap):
        if hwnd in self._items:
            self._items[hwnd].update_thumbnail(px)

    def deep_scan(self):
        """Scan for ALL off-screen windows, including Citrix/RDP ones."""
        self.full_refresh()   # normal refresh first
        extras = find_offscreen_deep(self.monitors)

        # Merge any newly found off-screen windows not already in self.windows
        known = {w.hwnd for w in self.windows}
        added = []
        for w in extras:
            if w.hwnd not in known:
                w.is_off = True
                self.windows.append(w)
                self._deep_extras[w.hwnd] = w   # persist across auto-refresh
                added.append(w)

        if added:
            self._rebuild_scene()
            self._min_bar.update_windows(self.windows)
            self._update_status()
            names = "\n".join(f"  • {w.title[:55]}" for w in added)
            QMessageBox.information(
                self, "Deep Scan — Found Hidden Windows",
                f"Found {len(added)} off-screen window(s) not in normal scan:\n\n{names}\n\n"
                "They now appear in red on the map.\n"
                "Drag them back or use Rescue All."
            )
        else:
            QMessageBox.information(
                self, "Deep Scan — Nothing Extra",
                "No additional off-screen windows found beyond the normal scan.\n\n"
                "If a Citrix/RDP dialog is missing, try clicking Refresh "
                "right after opening it."
            )

    def _on_restore_minimized(self, hwnd: int):
        do_focus(hwnd)
        QTimer.singleShot(300, self.full_refresh)

    def _update_status(self):
        visible   = [w for w in self.windows if not w.is_minimized]
        minimized = [w for w in self.windows if w.is_minimized]
        n_off = sum(1 for w in visible if w.is_off)
        warning = "  ← Rescue All" if n_off else "  ✓ on-screen"
        self._status.setText(
            f"  {len(self.monitors)} monitor(s)  ·  "
            f"{len(visible)} visible  ·  "
            f"{len(minimized)} minimized  ·  "
            f"{n_off} off-screen{warning}"
            f"   |   Scroll to zoom  ·  Drag to move"
        )
        self._rescue_btn.setStyleSheet(
            f"background: {RED}; color: white;" if n_off
            else ""   # fall back to global QSS (light text on MUTED bg)
        )

    # ── Actions ───────────────────────────────────────────────────────────────
    def rescue_all(self):
        off = [w for w in self.windows if w.is_off]
        if not off:
            QMessageBox.information(self, "All Clear", "No off-screen windows found.")
            return
        m = self.monitors[0] if self.monitors else None
        if not m:
            return
        for i, w in enumerate(off):
            do_rescue(w.hwnd, m, offset=i * 28)
        names = "\n".join(f"  • {w.title[:55]}" for w in off)
        QMessageBox.information(self, "Rescued",
                                f"Moved {len(off)} window(s) to Monitor 1:\n\n{names}")
        QTimer.singleShot(300, self.full_refresh)

    def _on_rescue_request(self, hwnd: int):
        m = self.monitors[0] if self.monitors else None
        if m:
            do_rescue(hwnd, m)
            QTimer.singleShot(200, self.full_refresh)

    def _on_focus_request(self, hwnd: int):
        do_focus(hwnd)

    def _on_move_finished(self, hwnd: int, x: int, y: int):
        QTimer.singleShot(200, self.full_refresh)

    # ── Context menu on right-click ───────────────────────────────────────────
    def contextMenuEvent(self, e):
        # Map to scene, find item
        scene_pos = self._view.mapToScene(self._view.mapFromGlobal(e.globalPos()))
        item = self._scene.itemAt(scene_pos, self._view.transform())
        if not isinstance(item, WinItem):
            return

        w   = item.win
        ctx = QMenu(self)
        ctx.setStyleSheet(self.styleSheet())

        title_act = ctx.addAction(w.title[:50])
        title_act.setEnabled(False)
        ctx.addSeparator()

        if w.is_off:
            for i, m in enumerate(self.monitors):
                def _make_rescue(hwnd=w.hwnd, mon=m, idx=i):
                    def _do():
                        do_rescue(hwnd, mon)
                        QTimer.singleShot(200, self.full_refresh)
                    return _do
                ctx.addAction(f"⚠  Rescue → Monitor {i + 1}", _make_rescue())
        else:
            ctx.addAction("▶  Focus / Bring to Front",
                          lambda hwnd=w.hwnd: do_focus(hwnd))

        ctx.addSeparator()
        info = ctx.addAction(f"Position ({w.left}, {w.top})  ·  Size {w.w}×{w.h}")
        info.setEnabled(False)
        ctx.exec_(e.globalPos())

    def closeEvent(self, e):
        # Minimize to tray instead of quitting
        e.ignore()
        self.hide()
        self._tray.showMessage("Window Rescue",
                               "Running in the system tray. Double-click to reopen.",
                               QSystemTrayIcon.Information, 2000)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if platform.system() != "Windows":
        print()
        print("  Window Rescue — Windows only")
        print("  Copy this file to your Windows machine, then:")
        print("    pip install PyQt5 pywin32 Pillow")
        print("    python window_rescue.py")
        print()
        sys.exit(0)

    # DPI awareness — must be called before QApplication
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)   # keep alive in tray

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
