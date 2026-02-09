from dataclasses import dataclass
from pathlib import Path
from win32com.shell import shellcon
import signal
import subprocess
import sys
import time
import win32com.shell.shell as shell

from PySide6.QtCore import Qt, QTimer, QUrl, QSize, QObject, Signal
from PySide6.QtGui import QPixmap, QAction, QIcon, QDrag, QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QHBoxLayout,
    QVBoxLayout,
    QSystemTrayIcon,
    QMenu,
)

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}


def default_screenshots_dir() -> Path:
    return Path(shell.SHGetKnownFolderPath(shellcon.FOLDERID_Screenshots))


def open_explorer_and_select(file_path: str) -> None:
    print(f"Opening explorer for file: {file_path}")
    subprocess.Popen(["explorer.exe", "/select,", file_path], shell=False)


@dataclass
class AppConfig:
    watch_dir: Path
    popup_seconds: int = 5
    max_preview_size: int = 220  # px thumbnail max side


class UiBridge(QObject):
    file_created = Signal(str)


class ScreenshotHandler(FileSystemEventHandler):
    """
    Watchdog runs in a background thread; emit signals into Qt safely via UiBridge.
    """
    def __init__(self, bridge: UiBridge, watch_dir: Path):
        super().__init__()
        self.bridge = bridge
        self.watch_dir = watch_dir

    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() not in IMAGE_EXTS:
            return

        # Some screenshot tools create then write; wait for file to become stable.
        if self._wait_for_stable_file(path):
            self.bridge.file_created.emit(str(path))

    @staticmethod
    def _wait_for_stable_file(path: Path, timeout_s: float = 2.5) -> bool:
        start = time.time()
        last_size = -1
        while time.time() - start < timeout_s:
            try:
                size = path.stat().st_size
                if size > 0 and size == last_size:
                    return True
                last_size = size
            except FileNotFoundError:
                return False
            time.sleep(0.08)
        return True  # best effort


class PreviewPopup(QWidget):
    def __init__(self, config: AppConfig):
        super().__init__()
        self.config = config
        self.current_file: str | None = None

        # Click vs drag handling
        self._press_pos = None
        self._drag_started = False

        self.setWindowFlags(
            Qt.Tool
            | Qt.FramelessWindowHint
            | Qt.WindowStaysOnTopHint
            | Qt.BypassWindowManagerHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self.container = QWidget(self)
        self.container.setObjectName("container")
        self.container.setStyleSheet("""
            QWidget#container {
                background: rgba(25, 25, 25, 235);
                border: 1px solid rgba(255, 255, 255, 40);
                border-radius: 12px;
            }
            QLabel {
                color: white;
            }
        """)

        self.thumb = QLabel()
        self.thumb.setFixedSize(QSize(self.config.max_preview_size, self.config.max_preview_size))
        self.thumb.setAlignment(Qt.AlignCenter)

        self.title = QLabel("Screenshot saved")
        self.title.setStyleSheet("font-weight: 600; font-size: 13px;")

        self.path_label = QLabel("")
        self.path_label.setStyleSheet("font-size: 11px; color: rgba(255,255,255,180);")
        self.path_label.setTextInteractionFlags(Qt.NoTextInteraction)
        self.path_label.setWordWrap(True)

        text_col = QVBoxLayout()
        text_col.addWidget(self.title)
        text_col.addWidget(self.path_label)
        text_col.addStretch(1)

        row = QHBoxLayout(self.container)
        row.setContentsMargins(12, 12, 12, 12)
        row.setSpacing(12)
        row.addWidget(self.thumb)
        row.addLayout(text_col)

        self.hide_timer = QTimer(self)
        self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.hide)

        self.setMouseTracking(True)
        self.container.setMouseTracking(True)
        self.thumb.setMouseTracking(True)
        self.title.setMouseTracking(True)
        self.path_label.setMouseTracking(True)

        self.resize(520, 260)

    def show_preview(self, file_path: str):
        self.current_file = file_path

        p = Path(file_path)
        self.path_label.setText(str(p))

        pix = QPixmap(file_path)
        if pix.isNull():
            # fallback: show filename only
            self.thumb.setText("Preview unavailable")
        else:
            scaled = pix.scaled(
                self.thumb.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.thumb.setPixmap(scaled)

        self._move_to_bottom_right()
        self.show()
        self.raise_()
        self.activateWindow()

        self.hide_timer.start(self.config.popup_seconds * 1000)

    def _move_to_bottom_right(self):
        screen = QGuiApplication.primaryScreen()
        geo = screen.availableGeometry()
        margin = 18
        x = geo.x() + geo.width() - self.width() - margin
        y = geo.y() + geo.height() - self.height() - margin
        self.move(x, y)

    # --- Mouse handling: click opens Explorer; drag drops file ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._press_pos = event.position().toPoint()
            self._drag_started = False
        elif event.button() == Qt.RightButton:
            self.hide()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not self.current_file:
            return super().mouseMoveEvent(event)

        if self._press_pos is None:
            return super().mouseMoveEvent(event)

        if not (event.buttons() & Qt.LeftButton):
            return super().mouseMoveEvent(event)

        move_pos = event.position().toPoint()
        dist = (move_pos - self._press_pos).manhattanLength()
        if dist < 8 or self._drag_started:
            return super().mouseMoveEvent(event)

        self._drag_started = True

        mime = QGuiApplication.clipboard().mimeData()
        # Use QMimeData with file URL for proper file DnD to browser
        from PySide6.QtCore import QMimeData
        md = QMimeData()
        md.setUrls([QUrl.fromLocalFile(self.current_file)])

        drag = QDrag(self)
        drag.setMimeData(md)

        # Use thumbnail as drag pixmap
        pm = self.thumb.pixmap()
        if pm:
            drag.setPixmap(pm)
            drag.setHotSpot(pm.rect().center())

        drag.exec(Qt.CopyAction)

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.current_file:
            # If no drag started, treat as click
            if not self._drag_started:
                open_explorer_and_select(self.current_file)
                self.hide()
        self._press_pos = None
        self._drag_started = False
        super().mouseReleaseEvent(event)


class TrayApp:
    def __init__(self, config: AppConfig):
        self.config = config
        self.bridge = UiBridge()
        self.bridge.file_created.connect(self.on_new_file)

        self.popup = PreviewPopup(config)

        self.tray = QSystemTrayIcon()
        self.tray.setToolTip("Screenshot Preview")
        self.tray.setIcon(QIcon.fromTheme("camera-photo") or QApplication.windowIcon())

        self.menu = QMenu()

        self.pause_action = QAction("Pause monitoring")
        self.pause_action.triggered.connect(self.toggle_pause)
        self.menu.addAction(self.pause_action)

        self.open_folder_action = QAction("Open screenshots folder")
        self.open_folder_action.triggered.connect(self.open_folder)
        self.menu.addAction(self.open_folder_action)

        self.quit_action = QAction("Quit")
        self.quit_action.triggered.connect(QApplication.quit)
        self.menu.addAction(self.quit_action)

        self.tray.setContextMenu(self.menu)
        self.tray.show()

        self._paused = False
        self._observer = None

    def open_folder(self):
        self.config.watch_dir.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer.exe", str(self.config.watch_dir)], shell=False)

    def toggle_pause(self):
        self._paused = not self._paused
        if self._paused:
            self.pause_action.setText("Resume monitoring")
        else:
            self.pause_action.setText("Pause monitoring")

    def on_new_file(self, file_path: str):
        print("Detected:", file_path)
        if self._paused:
            return
        self.popup.show_preview(file_path)

    def start_watching(self):
        self.config.watch_dir.mkdir(parents=True, exist_ok=True)
        handler = ScreenshotHandler(self.bridge, self.config.watch_dir)
        obs = Observer()
        obs.schedule(handler, str(self.config.watch_dir), recursive=False)
        obs.daemon = True
        obs.start()
        self._observer = obs


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Screenshot folder monitor with preview popup")
    parser.add_argument("--watch-dir", type=Path, default=default_screenshots_dir(),
                        help="Directory to watch for new screenshots")
    parser.add_argument("--popup-seconds", type=int, default=5,
                        help="Seconds to show the preview popup")
    parser.add_argument("--max-preview-size", type=int, default=220,
                        help="Max size in pixels for the preview thumbnail's longest side")
    args = parser.parse_args()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    # Handle Ctrl+C to quit the application gracefully
    # On Windows, we need a timer to periodically wake up the interpreter
    # so it can process signals
    def handle_sigint(sig, frame):
        print("Exiting...")
        app.quit()
    signal.signal(signal.SIGINT, handle_sigint)
    
    # Timer to allow Python to process signals in the Qt event loop
    signal_timer = QTimer()
    signal_timer.timeout.connect(lambda: None)
    signal_timer.start(100)  # 100ms interval

    config = AppConfig(
        watch_dir=args.watch_dir,
        popup_seconds=args.popup_seconds,
        max_preview_size=args.max_preview_size
    )
    tray_app = TrayApp(config)
    tray_app.start_watching()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
