import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtWidgets import QSplashScreen
from PySide6.QtGui import QPixmap, QIcon, QPainter
from PySide6.QtCore import Qt
from mcu_compare.ui.main_window import MainWindow
from mcu_compare.data.json_db import JsonDatabase


def ensure_data_dir():
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    if not os.path.exists(data_dir):
        os.makedirs(data_dir, exist_ok=True)
    return data_dir


def main():
    data_dir = ensure_data_dir()
    db_path = os.path.join(data_dir, 'app.json')
    db = JsonDatabase(db_path)
    db.initialize()

    app = QApplication(sys.argv)
    app.setApplicationName('StriveFit')
    # Show splash screen if asset exists
    try:
        root_dir = os.path.dirname(__file__)
        # Prefer project root image if present
        splash_candidates = [
            os.path.join(root_dir, 'fit.png'),
            os.path.join(root_dir, 'mcu_compare', 'assets', 'strivefit.png'),
            os.path.join(root_dir, 'strivefit.jpg'),
        ]
        splash = None
        for splash_path in splash_candidates:
            if os.path.exists(splash_path):
                pix = QPixmap(splash_path)
                if not pix.isNull():
                    # Make a smaller square splash (e.g., 256x256), center the image and keep aspect ratio
                    target_size = 812
                    square = QPixmap(target_size, target_size)
                    square.fill(Qt.transparent)
                    scaled = pix.scaled(target_size, target_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    painter = QPainter(square)
                    try:
                        x = (target_size - scaled.width()) // 2
                        y = (target_size - scaled.height()) // 2
                        painter.drawPixmap(x, y, scaled)
                    finally:
                        painter.end()
                    splash = QSplashScreen(square)
                    splash.setWindowFlag(Qt.WindowStaysOnTopHint, True)
                    splash.show()
                    app.processEvents()
                    break
    except Exception:
        splash = None
    # Load global stylesheet (QSS)
    try:
        base_dir = os.path.join(os.path.dirname(__file__), 'mcu_compare', 'ui')
        qss_path = os.path.join(base_dir, 'styles_dark.qss')
        if not os.path.exists(qss_path):
            qss_path = os.path.join(base_dir, 'styles.qss')
        if os.path.exists(qss_path):
            with open(qss_path, 'r', encoding='utf-8') as f:
                app.setStyleSheet(f.read())
    except Exception:
        pass
    # Set application icon (prefer root strivefit.png)
    try:
        icon_candidates = [
            os.path.join(os.path.dirname(__file__), 'log.png'),
            os.path.join(os.path.dirname(__file__), 'mcu_compare', 'assets', 'strivefit.png'),
        ]
        for icon_path in icon_candidates:
            if os.path.exists(icon_path):
                app.setWindowIcon(QIcon(icon_path))
                break
    except Exception:
        pass

    window = MainWindow(db)
    window.show()
    try:
        if 'splash' in locals() and splash is not None:
            splash.finish(window)
    except Exception:
        pass
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
