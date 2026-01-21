from typing import List, Dict, Any
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QMessageBox, QTabWidget
)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
import os
from PySide6.QtGui import QColor, QBrush, QAction, QActionGroup
from PySide6.QtWidgets import QAbstractItemView, QSizePolicy

from mcu_compare.engine.similarity import best_match, categorize
from mcu_compare.engine.similarity import weighted_similarity
from .dialogs import AddMCUDialog, DetailsDialog, AddCompanyDialog, AddNcoEntryDialog, ViewNcoEntriesDialog, EditNcoEntryDialog, EditMCUDialog


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
        self.resize(1400, 900)
        self.setMinimumSize(1200, 800)
        self._build_ui()
        self._load_companies()
        self._refresh_table()
        # Default theme
        self._set_theme('Dark')

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        # Tabs container
        tabs = QTabWidget(self)
        v.addWidget(tabs)
        # Compare tab page/layout
        compare_page = QWidget()
        compare_v = QVBoxLayout(compare_page)

        # Menus
        menubar = self.menuBar()
        data_menu = menubar.addMenu('Data')
        act_add_mcu = QAction('Add MCU', self)
        act_add_mcu.triggered.connect(self._open_add_dialog)
        data_menu.addAction(act_add_mcu)
        act_add_company = QAction('Add Company', self)
        act_add_company.triggered.connect(self._open_add_company)
        data_menu.addAction(act_add_company)

        view_menu = menubar.addMenu('View')
        theme_menu = view_menu.addMenu('Theme')
        theme_group = QActionGroup(self)
        theme_group.setExclusive(True)
        act_dark = QAction('Dark', self, checkable=True)
        act_light = QAction('Light', self, checkable=True)
        act_dark.setChecked(True)
        theme_group.addAction(act_dark)
        theme_group.addAction(act_light)
        theme_menu.addAction(act_dark)
        theme_menu.addAction(act_light)
        act_dark.triggered.connect(lambda: self._set_theme('Dark'))
        act_light.triggered.connect(lambda: self._set_theme('Light'))

        # Toolbar row
        row = QHBoxLayout()
        row.addWidget(QLabel('Company:'))
        self.company_combo = QComboBox()
        self.company_combo.currentIndexChanged.connect(self._refresh_table)
        row.addWidget(self.company_combo, 2)

        # Unified search: mode + query
        row.addWidget(QLabel('Search:'))
        self.search_mode = QComboBox()
        self.search_mode.addItems(['Company', 'MCU'])
        self.search_mode.currentIndexChanged.connect(self._on_search_mode_change)
        row.addWidget(self.search_mode)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText('Type to search...')
        self.search_edit.textChanged.connect(self._on_search_changed)
        row.addWidget(self.search_edit, 3)

        # Compare target selector
        row.addWidget(QLabel('Compare with:'))
        self.compare_combo = QComboBox()
        self.compare_combo.currentIndexChanged.connect(self._refresh_table)
        row.addWidget(self.compare_combo, 2)
        # Edit selected MCU button
        edit_selected_btn = QPushButton('Edit Selected')
        edit_selected_btn.clicked.connect(self._edit_selected_mcu)
        row.addWidget(edit_selected_btn)

        # Spacer ends toolbar

        compare_v.addLayout(row)

        # Table with company column
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['Manufacturer', 'Part NO', 'Compatibility', 'Match %', 'Category'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.cellDoubleClicked.connect(self._open_details)
        compare_v.addWidget(self.table)
        # Add compare tab
        tabs.addTab(compare_page, 'Compare')
        # Build and add NCO tab
        nco_page = self._build_nco_tab()
        tabs.addTab(nco_page, 'NCO/Commission')

    def _load_companies(self):
        # Only filter companies when search mode is Company
        query = self.search_edit.text() if hasattr(self, 'search_edit') else ''
        search = query if self.search_mode.currentText() == 'Company' else ''
        companies = [c for c in self.db.list_companies(search) if not c['is_ours']]
        current_id = self.company_combo.currentData() if self.company_combo.count() else None
        self.company_combo.blockSignals(True)
        self.company_combo.clear()
        for c in companies:
            self.company_combo.addItem(c['name'], c['id'])
        self.company_combo.blockSignals(False)
        # Prefer to retain selection; otherwise select first if available
        idx = -1
        if current_id is not None:
            idx = self.company_combo.findData(current_id)
        if idx >= 0:
            self.company_combo.setCurrentIndex(idx)
        elif self.company_combo.count() > 0:
            self.company_combo.setCurrentIndex(0)
        else:
            # No companies match; clear the table
            self.table.setRowCount(0)
        # Load our MCU choices
        self._load_our_mcu_choices()
        # Ensure table reflects the currently visible/selected company list
        self._refresh_table()

    def _load_our_mcu_choices(self):
        sel_id = self.compare_combo.currentData() if self.compare_combo.count() else None
        our_mcus = self.db.list_our_mcus()
        self.compare_combo.blockSignals(True)
        self.compare_combo.clear()
        self.compare_combo.addItem('Auto (Best Match)', None)
        for r in our_mcus:
            self.compare_combo.addItem(r['name'], r['id'])
        self.compare_combo.blockSignals(False)
        if sel_id is not None:
            idx = self.compare_combo.findData(sel_id)
            if idx >= 0:
                self.compare_combo.setCurrentIndex(idx)

    def _selected_our_mcu_id(self):
        return self.compare_combo.currentData()

    def _refresh_table(self):
        if self.company_combo.count() == 0:
            # No companies to show; clear table
            self.table.setRowCount(0)
            return
        company_id = self.company_combo.currentData()
        query = self.search_edit.text() if hasattr(self, 'search_edit') else ''
        mode = self.search_mode.currentText()
        if company_id is None and not (mode == 'MCU' and query):
            # No selected company and not doing MCU search; clear
            self.table.setRowCount(0)
            return
        our_mcus_rows = self.db.list_our_mcus()
        our_mcus = [dict(r) for r in our_mcus_rows]
        feat_cols = self.db.feature_columns()
        q = query.lower()
        if mode == 'MCU' and q:
            # Search across all competitor companies
            mcus_all = []
            all_companies = self.db.list_companies('')
            companies_map = {c['id']: c['name'] for c in all_companies}
            for c in all_companies:
                if c.get('is_ours'):
                    continue
                mcus_all.extend([dict(r) for r in self.db.list_mcus_by_company(c['id'])])
        elif mode == 'Company':
            mcus_all = [dict(r) for r in self.db.list_mcus_by_company(company_id)]
            companies_map = {company_id: next((c['name'] for c in self.db.list_companies('') if c['id']==company_id), '')}
        else:
            # Fallback: treat as company context
            mcus_all = [dict(r) for r in self.db.list_mcus_by_company(company_id)]
            companies_map = {company_id: next((c['name'] for c in self.db.list_companies('') if c['id']==company_id), '')}

        mcus = [m for m in mcus_all if (q in m.get('name', '').lower())] if (mode == 'MCU' and q) else mcus_all

        self.table.setRowCount(0)
        for mcu in mcus:
            # Build feature dicts
            target = {k: mcu.get(k) for k in feat_cols}
            chosen_id = self._selected_our_mcu_id()
            if chosen_id is None:
                best, score, _ = best_match(target, [{k: mm.get(k) for k in feat_cols} | {'id': mm['id'], 'name': mm['name']} for mm in our_mcus])
                category = categorize(score)
            else:
                # Use selected our MCU only
                mm = next((x for x in our_mcus if x['id'] == chosen_id), None)
                if mm is None:
                    continue
                from mcu_compare.engine.similarity import weighted_similarity
                score, _ = weighted_similarity(target, {k: mm.get(k) for k in feat_cols})
                category = categorize(score)
                best = {'id': mm['id'], 'name': mm['name'], **{k: mm.get(k) for k in feat_cols}}

            row = self.table.rowCount()
            self.table.insertRow(row)
            # Company name
            comp_name = companies_map.get(mcu.get('company_id'), '')
            self.table.setItem(row, 0, QTableWidgetItem(comp_name))
            self.table.setItem(row, 1, QTableWidgetItem(mcu['name']))
            self.table.setItem(row, 2, QTableWidgetItem(best['name'] if best else '-'))
            item_score = QTableWidgetItem(f"{score:.1f}")
            item_score.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, item_score)
            # Category chip as QLabel to avoid overriding row selection highlight
            from PySide6.QtWidgets import QLabel
            chip = QLabel(category)
            chip.setAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(category, '#444')
            chip.setStyleSheet(f"QLabel {{ background-color: {color}; color: white; border-radius: 10px; padding: 2px 6px; margin: 0px; }}")
            chip.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            self.table.setCellWidget(row, 4, chip)
            self.table.setRowHeight(row, 28)
            # Store IDs
            self.table.item(row, 1).setData(Qt.UserRole, mcu['id'])
            self.table.item(row, 2).setData(Qt.UserRole, best['id'] if best else None)

    def _open_add_dialog(self):
        dlg = AddMCUDialog(self.db, self)
        if dlg.exec():
            self._refresh_table()

    def _open_details(self, row: int, column: int):
        comp_id = self.table.item(row, 1).data(Qt.UserRole)
        our_id = self._selected_our_mcu_id() or self.table.item(row, 2).data(Qt.UserRole)
        if comp_id is None or our_id is None:
            QMessageBox.information(self, 'Details', 'No match available for detailed comparison.')
            return
        dlg = DetailsDialog(self.db, comp_id, our_id, self)
        dlg.exec()

    def _edit_selected_mcu(self):
        # Edit the selected competitor MCU (from the Compare tab table)
        if self.table.currentRow() < 0:
            return
        row = self.table.currentRow()
        item = self.table.item(row, 1)
        if item is None:
            return
        mcu_id = item.data(Qt.UserRole)
        if mcu_id is None:
            return
        dlg = EditMCUDialog(self.db, int(mcu_id), self)
        if dlg.exec():
            self._refresh_table()

    def _open_add_company(self):
        dlg = AddCompanyDialog(self.db, self)
        if dlg.exec():
            new_id = dlg.created_company_id
            self._load_companies()
            if new_id is not None:
                idx = self.company_combo.findData(new_id)
                if idx >= 0:
                    self.company_combo.setCurrentIndex(idx)
            self._refresh_table()

    def _open_nco_add(self):
        org_id = self.nco_org_combo.currentData() if hasattr(self, 'nco_org_combo') and self.nco_org_combo.count() else None
        dlg = AddNcoEntryDialog(self.db, self, org_id)
        if dlg.exec():
            pass

    def _open_nco_view(self):
        dlg = ViewNcoEntriesDialog(self.db, self)
        dlg.exec()

    def _build_nco_tab(self) -> QWidget:
        page = QWidget()
        v = QVBoxLayout(page)
        # Actions row
        row = QHBoxLayout()
        row.addWidget(QLabel('NCO:'))
        self.nco_org_combo = QComboBox()
        self.nco_org_combo.currentIndexChanged.connect(self._refresh_nco_table)
        row.addWidget(self.nco_org_combo, 2)
        add_org_btn = QPushButton('Add Org')
        add_org_btn.clicked.connect(self._add_nco_org)
        row.addWidget(add_org_btn)

        add_btn = QPushButton('Add Entry')
        add_btn.clicked.connect(self._open_nco_add)
        row.addWidget(add_btn)
        # Search box
        row.addWidget(QLabel('Search:'))
        self.nco_search = QLineEdit()
        self.nco_search.setPlaceholderText('Type to search NCO/company/MCU')
        self.nco_search.textChanged.connect(self._refresh_nco_table)
        row.addWidget(self.nco_search, 2)
        # Edit button
        edit_btn = QPushButton('Edit Selected')
        edit_btn.clicked.connect(self._edit_selected_nco)
        row.addWidget(edit_btn)
        refresh_btn = QPushButton('Refresh')
        refresh_btn.clicked.connect(self._refresh_nco_table)
        row.addWidget(refresh_btn)
        row.addStretch(1)
        v.addLayout(row)

        # NCO table
        self.nco_table = QTableWidget(0, 7)
        self.nco_table.setHorizontalHeaderLabels(['NCO', 'Company', 'Competitor MCU', 'Quantity', 'Our MCU', 'Match %', 'Category'])
        self.nco_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.nco_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.nco_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.nco_table.setAlternatingRowColors(True)
        self.nco_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.nco_table.cellDoubleClicked.connect(self._open_nco_details)
        v.addWidget(self.nco_table)

        self._load_nco_orgs()
        self._refresh_nco_table()
        return page

    def _refresh_nco_table(self):
        if not hasattr(self, 'nco_table'):
            return
        # Determine selected org
        org_id = self.nco_org_combo.currentData() if hasattr(self, 'nco_org_combo') and self.nco_org_combo.count() else None
        rows = self.db.list_nco_entries(org_id)
        comps = {c['id']: c['name'] for c in self.db.list_companies('')}
        orgs = {o['id']: o['name'] for o in self.db.list_nco_orgs()}
        mcu_name: Dict[int, str] = {}
        for c in self.db.list_companies(''):
            for m in self.db.list_mcus_by_company(c['id']):
                mcu_name[m['id']] = m['name']
        feat_cols = self.db.feature_columns()
        our_mcus = [dict(r) for r in self.db.list_our_mcus()]
        # Apply text filter
        q = self.nco_search.text().lower() if hasattr(self, 'nco_search') else ''
        def matches(row: Dict[str, Any]) -> bool:
            if not q:
                return True
            nco = orgs.get(row.get('org_id'), '')
            comp = comps.get(row.get('company_id'), '')
            cmcu = mcu_name.get(row.get('comp_mcu_id'), '')
            omcu = mcu_name.get(row.get('our_mcu_id'), '')
            return (q in nco.lower()) or (q in comp.lower()) or (q in str(cmcu).lower()) or (q in str(omcu).lower())

        filtered = [r for r in rows if matches(r)]
        self.nco_table.setRowCount(0)
        for r in filtered:
            row = self.nco_table.rowCount()
            self.nco_table.insertRow(row)
            self.nco_table.setItem(row, 0, QTableWidgetItem(orgs.get(r.get('org_id'), '')))
            self.nco_table.setItem(row, 1, QTableWidgetItem(comps.get(r.get('company_id'), '')))
            comp_name = mcu_name.get(r.get('comp_mcu_id'), '')
            self.nco_table.setItem(row, 2, QTableWidgetItem(comp_name))
            self.nco_table.setItem(row, 3, QTableWidgetItem(str(r.get('quantity', 0))))

            # Determine our MCU and compute similarity
            comp = self.db.get_mcu_by_id(int(r.get('comp_mcu_id')))
            target = {k: comp.get(k) for k in feat_cols} if comp else {}
            our_id = r.get('our_mcu_id')
            if our_id is None:
                # compute best match across our MCUs
                candidates = [{k: mm.get(k) for k in feat_cols} | {'id': mm['id'], 'name': mm['name']} for mm in our_mcus]
                best, score, _ = best_match(target, candidates) if target else (None, 0.0, {})
                our_name = best['name'] if best else ''
            else:
                ours = self.db.get_mcu_by_id(int(our_id))
                our_feats = {k: (ours.get(k) if ours else None) for k in feat_cols}
                score, _ = weighted_similarity(target, our_feats) if target else (0.0, {})
                our_name = mcu_name.get(our_id, '')
            self.nco_table.setItem(row, 4, QTableWidgetItem(our_name))
            item_score = QTableWidgetItem(f"{score:.1f}")
            item_score.setTextAlignment(Qt.AlignCenter)
            self.nco_table.setItem(row, 5, item_score)
            cat = categorize(score)
            from PySide6.QtWidgets import QLabel
            chip = QLabel(cat)
            chip.setAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(cat, '#444')
            chip.setStyleSheet(f"QLabel {{ background-color: {color}; color: white; border-radius: 10px; padding: 2px 6px; margin: 0px; }}")
            chip.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            self.nco_table.setCellWidget(row, 6, chip)
            # Store entry id on first column
            if self.nco_table.item(row, 0):
                self.nco_table.item(row, 0).setData(Qt.UserRole, r.get('id'))

    def _load_nco_orgs(self):
        if not hasattr(self, 'nco_org_combo'):
            return
        orgs = self.db.list_nco_orgs()
        current = self.nco_org_combo.currentData() if self.nco_org_combo.count() else None
        self.nco_org_combo.blockSignals(True)
        self.nco_org_combo.clear()
        for o in orgs:
            self.nco_org_combo.addItem(o['name'], o['id'])
        self.nco_org_combo.blockSignals(False)
        if current is not None:
            idx = self.nco_org_combo.findData(current)
            if idx >= 0:
                self.nco_org_combo.setCurrentIndex(idx)

    def _add_nco_org(self):
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, 'Add Organization', 'Organization Name:')
        if ok and name.strip():
            self.db.add_nco_org(name.strip())
            self._load_nco_orgs()
            self._refresh_nco_table()

    def _edit_selected_nco(self):
        if not hasattr(self, 'nco_table') or self.nco_table.currentRow() < 0:
            return
        row = self.nco_table.currentRow()
        entry_id = self.nco_table.item(row, 0).data(Qt.UserRole) if self.nco_table.item(row, 0) else None
        if entry_id is None:
            return
        # Reconstruct entry dict
        entries = self.db.list_nco_entries()
        entry = next((e for e in entries if int(e.get('id', -1)) == int(entry_id)), None)
        if entry is None:
            return
        dlg = EditNcoEntryDialog(self.db, entry, self)
        if dlg.exec():
            self._refresh_nco_table()

    def _open_nco_details(self, row: int, column: int):
        # Open comparison details for the selected NCO entry
        if not hasattr(self, 'nco_table'):
            return
        item = self.nco_table.item(row, 0)
        if item is None:
            return
        entry_id = item.data(Qt.UserRole)
        if entry_id is None:
            return
        entries = self.db.list_nco_entries()
        entry = next((e for e in entries if int(e.get('id', -1)) == int(entry_id)), None)
        if not entry:
            return
        comp_id = int(entry.get('comp_mcu_id'))
        our_id = entry.get('our_mcu_id')
        if our_id is None:
            # Compute best match dynamically
            feat_cols = self.db.feature_columns()
            comp = self.db.get_mcu_by_id(comp_id)
            if not comp:
                return
            target = {k: comp.get(k) for k in feat_cols}
            our_mcus = [dict(r) for r in self.db.list_our_mcus()]
            candidates = [{k: mm.get(k) for k in feat_cols} | {'id': mm['id'], 'name': mm['name']} for mm in our_mcus]
            best, score, _ = best_match(target, candidates)
            our_id = best['id'] if best else None
        if our_id is None:
            return
        dlg = DetailsDialog(self.db, comp_id, int(our_id), self)
        dlg.exec()

    def _set_theme(self, theme: str):
        base_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'ui')
        qss_file = 'styles_dark.qss' if theme == 'Dark' else 'styles_light.qss'
        qss_path = os.path.join(base_dir, qss_file)
        if not os.path.exists(qss_path):
            qss_path = os.path.join(base_dir, 'styles.qss')
        try:
            if os.path.exists(qss_path):
                with open(qss_path, 'r', encoding='utf-8') as f:
                    QApplication.instance().setStyleSheet(f.read())
        except Exception:
            pass

    def _on_search_mode_change(self):
        # Update placeholders and lists when switching mode
        mode = self.search_mode.currentText()
        if mode == 'Company':
            self.search_edit.setPlaceholderText('Type to search companies')
            self._load_companies()
        else:
            self.search_edit.setPlaceholderText('Type to search MCUs')
            # Keep companies dropdown as-is; table will use global MCU search
            self._refresh_table()

    def _on_search_changed(self, *_):
        # When searching companies, update dropdown; for MCUs, refresh table
        if self.search_mode.currentText() == 'Company':
            self._load_companies()
        else:
            self._refresh_table()
