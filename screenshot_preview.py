from dataclasses import dataclass
from pathlib import Path
from win32com.shell import shellcon
import signal
import subprocess
import sys
import time
import win32com.shell.shell as shell

from PySide6.QtCore import Qt, QTimer, QUrl, QSize, QObject, Signal, QPropertyAnimation, QEasingCurve, QPoint
from PySide6.QtGui import QPixmap, QAction, QIcon, QDrag, QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QHBoxLayout,
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

        # Animation attributes
        self._slide_in_animation: QPropertyAnimation | None = None
        self._slide_out_animation: QPropertyAnimation | None = None
        self._target_pos: QPoint | None = None
        self._is_animating_out = False

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

        row = QHBoxLayout(self.container)
        row.setContentsMargins(12, 12, 12, 12)
        row.addWidget(self.thumb)

        self.hide_timer = QTimer(self)
        self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.hide)

        self.setMouseTracking(True)
        self.container.setMouseTracking(True)
        self.thumb.setMouseTracking(True)

        # Calculate window size based on thumbnail size + margins (12px each side)
        window_size = self.config.max_preview_size + 24
        self.resize(window_size, window_size)

    def show_preview(self, file_path: str):
        self.current_file = file_path

        pix = QPixmap(file_path)
        if pix.isNull():
            # fallback: show filename only
            self.thumb.setText("Preview unavailable")
        else:
            # Calculate scaled size while respecting max dimension
            orig_width = pix.width()
            orig_height = pix.height()
            max_size = self.config.max_preview_size

            # Scale to fit within max_size, scaling up if smaller
            scale = min(max_size / orig_width, max_size / orig_height)
            new_width = int(orig_width * scale)
            new_height = int(orig_height * scale)

            # Scale the pixmap to the calculated size
            scaled = pix.scaled(new_width, new_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.thumb.setPixmap(scaled)

            # Resize thumbnail label and window to match the image dimensions
            self.thumb.setFixedSize(new_width, new_height)
            self.resize(new_width + 24, new_height + 24)  # add margins

        self._animate_in()

    def _animate_in(self):
        # Cancel any ongoing animations
        if self._slide_out_animation and self._slide_out_animation.state() == QPropertyAnimation.Running:
            self._slide_out_animation.stop()
        if self._slide_in_animation and self._slide_in_animation.state() == QPropertyAnimation.Running:
            self._slide_in_animation.stop()

        # Calculate target position
        screen = QGuiApplication.primaryScreen()
        geo = screen.availableGeometry()
        margin = 24
        target_x = geo.x() + geo.width() - self.width() - margin
        target_y = geo.y() + geo.height() - self.height() - margin
        self._target_pos = QPoint(target_x, target_y)

        # Start off-screen to the right
        start_x = geo.x() + geo.width() + 20
        start_pos = QPoint(start_x, target_y)
        self.move(start_pos)
        self._is_animating_out = False

        # Show the window
        self.show()
        self.raise_()
        self.activateWindow()

        # Create slide-in animation
        self._slide_in_animation = QPropertyAnimation(self, b"pos")
        self._slide_in_animation.setDuration(500)  # 500ms
        self._slide_in_animation.setStartValue(start_pos)
        self._slide_in_animation.setEndValue(self._target_pos)
        self._slide_in_animation.setEasingCurve(QEasingCurve.Type.OutBack)
        self._slide_in_animation.start()

        # Start hide timer
        self.hide_timer.start(self.config.popup_seconds * 1000)

    def _animate_out(self):
        if self._is_animating_out:
            return
        self._is_animating_out = True

        # Cancel any ongoing animations
        if self._slide_in_animation and self._slide_in_animation.state() == QPropertyAnimation.Running:
            self._slide_in_animation.stop()

        # Calculate end position (off-screen to the right)
        screen = QGuiApplication.primaryScreen()
        geo = screen.availableGeometry()
        end_x = geo.x() + geo.width() + self.width()
        end_pos = QPoint(end_x, self._target_pos.y() if self._target_pos else self.y())

        # Create slide-out animation
        self._slide_out_animation = QPropertyAnimation(self, b"pos")
        self._slide_out_animation.setDuration(400)  # 400ms
        self._slide_out_animation.setStartValue(self.pos())
        self._slide_out_animation.setEndValue(end_pos)
        self._slide_out_animation.setEasingCurve(QEasingCurve.Type.InBack)
        self._slide_out_animation.finished.connect(self._hide_after_animation)
        self._slide_out_animation.start()

    def _hide_after_animation(self):
        super(PreviewPopup, self).hide()
        self._is_animating_out = False

    def hide(self):
        self._animate_out()

    def _move_to_bottom_right(self):
        screen = QGuiApplication.primaryScreen()
        geo = screen.availableGeometry()
        margin = 24
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
