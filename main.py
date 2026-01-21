import sys
import os
from PySide6.QtWidgets import QApplication
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
    window = MainWindow(db)
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
