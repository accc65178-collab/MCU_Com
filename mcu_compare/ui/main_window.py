from typing import List, Dict, Any
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import QAbstractItemView

from mcu_compare.engine.similarity import best_match, categorize
from .dialogs import AddMCUDialog, DetailsDialog


CATEGORY_COLORS = {
    'Direct': '#2e7d32',
    'Near': '#1565c0',
    'Partial': '#f9a825',
    'No match': '#c62828'
}


class MainWindow(QMainWindow):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.setWindowTitle('MCU Comparator')
        self.resize(1100, 700)
        self._build_ui()
        self._load_companies()
        self._refresh_table()

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        v = QVBoxLayout(central)

        # Toolbar row
        row = QHBoxLayout()
        row.addWidget(QLabel('Company:'))
        self.company_combo = QComboBox()
        self.company_combo.currentIndexChanged.connect(self._refresh_table)
        row.addWidget(self.company_combo, 2)

        row.addWidget(QLabel('Search company:'))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText('Type to search companies')
        self.search_edit.textChanged.connect(self._load_companies)
        row.addWidget(self.search_edit, 2)

        self.add_btn = QPushButton('Add MCU')
        self.add_btn.clicked.connect(self._open_add_dialog)
        row.addWidget(self.add_btn)

        v.addLayout(row)

        # Table
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(['Company MCU', 'Best OUR MCU', 'Match %', 'Category'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.cellDoubleClicked.connect(self._open_details)
        v.addWidget(self.table)

    def _load_companies(self):
        search = self.search_edit.text() if hasattr(self, 'search_edit') else ''
        companies = self.db.list_companies(search)
        current_id = self.company_combo.currentData() if self.company_combo.count() else None
        self.company_combo.blockSignals(True)
        self.company_combo.clear()
        for c in companies:
            self.company_combo.addItem(c['name'] + (' (Our)' if c['is_ours'] else ''), c['id'])
        self.company_combo.blockSignals(False)
        if current_id is not None:
            idx = self.company_combo.findData(current_id)
            if idx >= 0:
                self.company_combo.setCurrentIndex(idx)

    def _refresh_table(self):
        if self.company_combo.count() == 0:
            return
        company_id = self.company_combo.currentData()
        if company_id is None:
            return
        our_mcus_rows = self.db.list_our_mcus()
        our_mcus = [dict(r) for r in our_mcus_rows]
        feat_cols = self.db.feature_columns()
        mcus = [dict(r) for r in self.db.list_mcus_by_company(company_id)]

        self.table.setRowCount(0)
        for mcu in mcus:
            # Build feature dicts
            target = {k: mcu.get(k) for k in feat_cols}
            best, score, _ = best_match(target, [{k: mm.get(k) for k in feat_cols} | {'id': mm['id'], 'name': mm['name']} for mm in our_mcus])
            category = categorize(score)

            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(mcu['name']))
            self.table.setItem(row, 1, QTableWidgetItem(best['name'] if best else '-'))
            item_score = QTableWidgetItem(f"{score:.1f}")
            item_score.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 2, item_score)
            item_cat = QTableWidgetItem(category)
            item_cat.setTextAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(category, '#444')
            item_cat.setForeground(QBrush(QColor('white')))
            # Background color by category
            self.table.setItem(row, 3, item_cat)
            self.table.item(row, 3).setBackground(QBrush(QColor(color)))
            # Store IDs
            self.table.item(row, 0).setData(Qt.UserRole, mcu['id'])
            self.table.item(row, 1).setData(Qt.UserRole, best['id'] if best else None)

    def _open_add_dialog(self):
        dlg = AddMCUDialog(self.db, self)
        if dlg.exec():
            self._refresh_table()

    def _open_details(self, row: int, column: int):
        comp_id = self.table.item(row, 0).data(Qt.UserRole)
        our_id = self.table.item(row, 1).data(Qt.UserRole)
        if comp_id is None or our_id is None:
            QMessageBox.information(self, 'Details', 'No match available for detailed comparison.')
            return
        dlg = DetailsDialog(self.db, comp_id, our_id, self)
        dlg.exec()
