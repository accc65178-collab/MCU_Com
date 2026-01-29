from typing import List, Dict, Any
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QMessageBox, QTabWidget
)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
import webbrowser
import os
from PySide6.QtGui import QColor, QBrush, QAction, QActionGroup, QIcon, QCursor, QKeySequence, QPainter, QPixmap, QPdfWriter, QPageLayout, QPageSize, QImage, QTextDocument
from PySide6.QtCore import QRect, QPoint, QItemSelectionModel
from PySide6.QtGui import QRegion
from PySide6.QtCore import QSize
from PySide6.QtWidgets import QAbstractItemView, QSizePolicy, QFileDialog
from PySide6.QtPrintSupport import QPrinter, QPrintDialog

from mcu_compare.engine.similarity import best_match, categorize
from mcu_compare.engine.similarity import weighted_similarity
from .dialogs import AddMCUDialog, DetailsDialog, AddCompanyDialog, AddNcoEntryDialog, ViewNcoEntriesDialog, EditNcoEntryDialog, EditMCUDialog


CATEGORY_COLORS = {
    'Best Match': '#2e7d32',
    'Partial': '#f9a825',
    'No Match': '#c62828'
}


class MainWindow(QMainWindow):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.setWindowTitle('StriveFit')
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
        self.tabs = QTabWidget(self)
        v.addWidget(self.tabs)
        # Compare tab page/layout
        compare_page = QWidget()
        compare_v = QVBoxLayout(compare_page)

        # Menus
        menubar = self.menuBar()
        # File menu with Print
        file_menu = menubar.addMenu('File')
        self.act_print = QAction('Print', self)
        self.act_print.setShortcut(QKeySequence.Print)
        self.act_print.triggered.connect(self._print_current)
        file_menu.addAction(self.act_print)
        # Export to PDF for reliable table capture
        self.act_export_pdf = QAction('Export Table to PDF…', self)
        self.act_export_pdf.setShortcut('Ctrl+Shift+P')
        self.act_export_pdf.triggered.connect(self._export_current_table_pdf)
        file_menu.addAction(self.act_export_pdf)
        data_menu = menubar.addMenu('Data')
        self.act_add_mcu = QAction('Add MCU', self)
        self.act_add_mcu.triggered.connect(self._open_add_dialog)
        data_menu.addAction(self.act_add_mcu)
        self.act_add_company = QAction('Add Company', self)
        self.act_add_company.triggered.connect(self._open_add_company)
        data_menu.addAction(self.act_add_company)
        self.act_rename_company = QAction('Rename Company', self)
        self.act_rename_company.triggered.connect(self._rename_company)
        data_menu.addAction(self.act_rename_company)
        self.act_delete_company = QAction('Delete Company', self)
        self.act_delete_company.triggered.connect(self._delete_company)
        data_menu.addAction(self.act_delete_company)

        # NCO/Commission submenu
        nco_menu = data_menu.addMenu('NCO/Commission')
        self.act_nco_add_org = QAction('Add Org', self)
        self.act_nco_add_org.triggered.connect(self._add_nco_org)
        nco_menu.addAction(self.act_nco_add_org)
        self.act_nco_rename_org = QAction('Rename Org', self)
        self.act_nco_rename_org.triggered.connect(self._rename_nco_org)
        nco_menu.addAction(self.act_nco_rename_org)
        self.act_nco_delete_org = QAction('Delete Org', self)
        self.act_nco_delete_org.triggered.connect(self._delete_nco_org)
        nco_menu.addAction(self.act_nco_delete_org)
        nco_menu.addSeparator()
        self.act_nco_add_entry = QAction('Add Entry', self)
        self.act_nco_add_entry.triggered.connect(self._open_nco_add)
        nco_menu.addAction(self.act_nco_add_entry)
        self.act_nco_edit_selected = QAction('Edit Selected', self)
        self.act_nco_edit_selected.triggered.connect(self._edit_selected_nco)
        nco_menu.addAction(self.act_nco_edit_selected)
        self.act_nco_delete_entry = QAction('Delete Selected Entry', self)
        self.act_nco_delete_entry.triggered.connect(self._delete_selected_nco)
        nco_menu.addAction(self.act_nco_delete_entry)
        self.act_nco_view = QAction('View Entries', self)
        self.act_nco_view.triggered.connect(self._open_nco_view)
        nco_menu.addAction(self.act_nco_view)

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
        self.company_combo.currentIndexChanged.connect(self._update_compare_actions)
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
        self.btn_edit_selected = QPushButton('Edit Selected')
        self.btn_edit_selected.clicked.connect(self._edit_selected_mcu)
        row.addWidget(self.btn_edit_selected)
        self.btn_delete_selected = QPushButton('Delete Selected')
        self.btn_delete_selected.clicked.connect(self._delete_selected_mcu)
        row.addWidget(self.btn_delete_selected)
        # Category counts summary
        self.cat_counts_label = QLabel('')
        self.cat_counts_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        row.addWidget(self.cat_counts_label, 2)

        # Spacer ends toolbar

        compare_v.addLayout(row)

        # Table with company column
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['Manufacturer', 'Part NO', 'Compatibility', 'Match %', 'Category'])
        # Column sizing: make 'Part NO' wider, others fit content
        t_header = self.table.horizontalHeader()
        t_header.setSectionResizeMode(QHeaderView.ResizeToContents)
        t_header.setSectionResizeMode(1, QHeaderView.Stretch)  # Part NO column
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.cellDoubleClicked.connect(self._open_details)
        # Enable interactive sorting
        self.table.setSortingEnabled(True)
        # Ensure enough row height for icon + text
        try:
            self.table.verticalHeader().setDefaultSectionSize(36)
        except Exception:
            pass
        compare_v.addWidget(self.table)
        # Add compare tab
        self.tabs.addTab(compare_page, 'Compare')
        # Build and add NCO tab
        nco_page = self._build_nco_tab()
        self.tabs.addTab(nco_page, 'NCO/Commission')
        # Toggle action enabled states when switching tabs
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._on_tab_changed(self.tabs.currentIndex())

    def _print_current(self):
        # Determine which widget to print based on current tab
        # If a Details dialog is open and visible, print that instead
        if hasattr(self, '_details_dialog') and getattr(self, '_details_dialog', None) is not None:
            try:
                if self._details_dialog.isVisible():
                    # Call public wrapper on the dialog
                    if hasattr(self._details_dialog, 'print_details'):
                        self._details_dialog.print_details()
                    return
            except Exception:
                pass
        idx = self.tabs.currentIndex() if hasattr(self, 'tabs') else 0
        target = None
        title = 'StriveFit'
        if idx == 0 and hasattr(self, 'table'):
            target = self.table
            title = 'Compare'
            font_pt = 11
            landscape = False
        elif idx == 1 and hasattr(self, 'nco_table'):
            target = self.nco_table
            title = 'NCO/Commission'
            font_pt = 11
            landscape = True
        if target is None:
            QMessageBox.information(self, 'Print', 'No table available to print on this tab.')
            return
        # For tables, open a print dialog and render a text-only table via QTextDocument
        try:
            if isinstance(target, QTableWidget):
                # Defaults if not set above
                font_pt = locals().get('font_pt', 12)
                landscape = locals().get('landscape', False)
                self._print_table_full(target, title, font_pt=font_pt, landscape=landscape)
            else:
                self._print_widget(target, title)
        except Exception as e:
            try:
                QMessageBox.warning(self, 'Print', f'Printing failed: {e}')
            except Exception:
                pass

    def _print_widget(self, widget, title: str = ''):
        # Open system print dialog and render the widget scaled to page
        printer = QPrinter(QPrinter.HighResolution)
        dlg = QPrintDialog(printer, self)
        if title:
            dlg.setWindowTitle(f'Print {title}')
        if dlg.exec() != QPrintDialog.Accepted:
            return
        pix = widget.grab()
        painter = QPainter(printer)
        try:
            page = printer.pageRect()
            # Scale pixmap to fit page
            scaled = pix.scaled(page.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            x = page.x() + (page.width() - scaled.width()) // 2
            y = page.y() + (page.height() - scaled.height()) // 2
            painter.drawPixmap(x, y, scaled)
        finally:
            painter.end()

    def _print_table_full(self, table: QTableWidget, title: str = '', font_pt: int = 12, landscape: bool = False):
        # Build text-only HTML and print via QTextDocument
        html = self._table_to_html(table, font_pt=font_pt)
        if not html:
            QMessageBox.information(self, 'Print', 'No data to print.')
            return
        printer = QPrinter(QPrinter.HighResolution)
        if landscape:
            try:
                # Force via PageLayout to avoid driver overrides
                layout = printer.pageLayout()
                layout.setOrientation(QPageLayout.Landscape)
                printer.setPageLayout(layout)
            except Exception:
                try:
                    printer.setOrientation(QPrinter.Landscape)
                except Exception:
                    pass
        dlg = QPrintDialog(printer, self)
        if title:
            dlg.setWindowTitle(f'Print {title}')
        if dlg.exec() != QPrintDialog.Accepted:
            return
        # Some drivers reset orientation from the dialog; force again after accept
        if landscape:
            try:
                layout = printer.pageLayout()
                layout.setOrientation(QPageLayout.Landscape)
                printer.setPageLayout(layout)
            except Exception:
                try:
                    printer.setOrientation(QPrinter.Landscape)
                except Exception:
                    pass
        doc = QTextDocument()
        doc.setHtml(html)
        try:
            from PySide6.QtCore import QSizeF
            paint = printer.pageLayout().paintRectPixels(printer.resolution())
            doc.setPageSize(QSizeF(paint.size()))
            doc.setTextWidth(paint.width())
        except Exception:
            pass
        doc.print_(printer)

    def _export_current_table_pdf(self):
        # If a Details dialog is open and visible, export that instead
        if hasattr(self, '_details_dialog') and getattr(self, '_details_dialog', None) is not None:
            try:
                if self._details_dialog.isVisible():
                    if hasattr(self._details_dialog, 'export_details_pdf'):
                        self._details_dialog.export_details_pdf()
                    return
            except Exception:
                pass
        idx = self.tabs.currentIndex() if hasattr(self, 'tabs') else 0
        target = None
        title = 'StriveFit'
        if idx == 0 and hasattr(self, 'table'):
            target = self.table
            title = 'Compare'
            font_pt = 11
            landscape = False
        elif idx == 1 and hasattr(self, 'nco_table'):
            target = self.nco_table
            title = 'NCO-Commission'
            font_pt = 11
            landscape = True
        if not isinstance(target, QTableWidget):
            QMessageBox.information(self, 'Export PDF', 'No table available to export on this tab.')
            return
        try:
            # Defaults if not set above
            font_pt = locals().get('font_pt', 12)
            landscape = locals().get('landscape', False)
            self._export_table_pdf(target, f"{title}.pdf", prompt=True, font_pt=font_pt, landscape=landscape)
        except Exception as e:
            try:
                QMessageBox.warning(self, 'Export PDF', f'Export failed: {e}')
            except Exception:
                pass

    def _export_table_pdf(self, table: QTableWidget, default_name: str, prompt: bool = False, *, font_pt: int = 12, landscape: bool = False):
        # Ask for destination path if requested
        if prompt:
            path, _ = QFileDialog.getSaveFileName(self, 'Export Table to PDF', default_name, 'PDF Files (*.pdf)')
            if not path:
                return
            if not path.lower().endswith('.pdf'):
                path += '.pdf'
        else:
            path = default_name
        # Build text-only HTML and export via QPrinter in PdfFormat
        html = self._table_to_html(table, font_pt=font_pt)
        if not html:
            QMessageBox.information(self, 'Export PDF', 'No data to export.')
            return
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
        if landscape:
            try:
                layout = printer.pageLayout()
                layout.setOrientation(QPageLayout.Landscape)
                printer.setPageLayout(layout)
            except Exception:
                try:
                    printer.setOrientation(QPrinter.Landscape)
                except Exception:
                    pass
        doc = QTextDocument()
        doc.setHtml(html)
        try:
            from PySide6.QtCore import QSizeF
            paint = printer.pageLayout().paintRectPixels(printer.resolution())
            doc.setPageSize(QSizeF(paint.size()))
            doc.setTextWidth(paint.width())
        except Exception:
            pass
        doc.print_(printer)

    def _table_to_html(self, table: QTableWidget, *, font_pt: int = 12) -> str:
        # Extract headers
        cols = table.columnCount()
        rows = table.rowCount()
        if cols == 0 or rows == 0:
            return ''
        headers = []
        for c in range(cols):
            hi = table.horizontalHeaderItem(c)
            headers.append((hi.text() if hi else '').strip())
        # Build rows from items/widgets
        from PySide6.QtWidgets import QLabel
        def cell_text(r: int, c: int) -> str:
            # Prefer visual label text (chip/overlay) first to avoid numeric EditRole leakage
            w = table.cellWidget(r, c)
            if w is not None:
                if isinstance(w, QLabel) and w.text():
                    return w.text()
                lbl = w.findChild(QLabel)
                if lbl is not None and lbl.text():
                    return lbl.text()
            it = table.item(r, c)
            if it is not None and it.text():
                return it.text()
            return ''
        # HTML with simple borders, compact font
        html = [
            '<html><head><meta charset="utf-8">',
            '<style>',
            f'table{{border-collapse:collapse;width:100%;table-layout:fixed;font:{font_pt}pt Segoe UI,Arial;}}',
            'th,td{border:1px solid #888;padding:6pt 8pt;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
            'th{background:#eaeaea;font-weight:600;}',
            'body{margin:10mm;}',
            '</style></head><body>',
            '<table>'
        ]
        # Header row
        html.append('<tr>')
        for h in headers:
            html.append(f'<th>{h}</th>')
        html.append('</tr>')
        # Data rows
        for r in range(rows):
            html.append('<tr>')
            for c in range(cols):
                txt = cell_text(r, c)
                # Escape HTML special chars
                txt = (txt or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                html.append(f'<td>{txt}</td>')
            html.append('</tr>')
        html.append('</table></body></html>')
        return ''.join(html)

    def _compose_table_image(self, table: QTableWidget):
        """Build a full-image snapshot by cloning the table into an offscreen widget.
        This avoids viewport-only rendering and ensures all rows/columns are captured."""
        try:
            cols = table.columnCount()
            rows = table.rowCount()
            if cols == 0 or rows == 0:
                return None, 0, 0
            # Create offscreen table
            clone = QTableWidget(rows, cols)
            clone.setStyleSheet(table.styleSheet())
            # Headers
            for c in range(cols):
                hdr = table.horizontalHeaderItem(c)
                clone.setHorizontalHeaderItem(c, QTableWidgetItem(hdr.text() if hdr else ''))
                try:
                    cw = table.columnWidth(c)
                    clone.setColumnWidth(c, cw)
                except Exception:
                    pass
            # Copy content text for each cell (prefer item text; fallback to widget label text)
            from PySide6.QtWidgets import QLabel
            for r in range(rows):
                for c in range(cols):
                    txt = ''
                    it = table.item(r, c)
                    if it is not None:
                        txt = it.text()
                    else:
                        w = table.cellWidget(r, c)
                        if w is not None:
                            # try to find a QLabel child that holds the text
                            lbl = w.findChild(QLabel)
                            if lbl is not None:
                                txt = lbl.text()
                    clone.setItem(r, c, QTableWidgetItem(txt))
                # Preserve row height
                try:
                    clone.setRowHeight(r, table.rowHeight(r))
                except Exception:
                    pass
            # Compute full size (headers + rows)
            hh = table.horizontalHeader()
            vh = table.verticalHeader()
            # Prefer clone metrics; fallback to table metrics
            header_h = max(clone.horizontalHeader().height(), hh.height(), 24)
            vh_w = max(clone.verticalHeader().width(), vh.width(), 24)
            # Sum of row heights (prefer clone values if set)
            try:
                content_h = sum(max(clone.rowHeight(r), table.sizeHintForRow(r) or 24) for r in range(rows))
                if content_h <= 0:
                    raise ValueError()
            except Exception:
                content_h = sum(max(table.rowHeight(r), table.sizeHintForRow(r) or 24) for r in range(rows))
            # Sum of column widths, fallback to viewport width if sum is tiny
            try:
                sum_cols = sum(max(table.columnWidth(c), hh.sectionSize(c)) for c in range(cols))
            except Exception:
                sum_cols = sum(table.columnWidth(c) for c in range(cols))
            if sum_cols <= 0:
                sum_cols = sum(max(clone.columnWidth(c), clone.horizontalHeader().sectionSize(c)) for c in range(cols))
            if sum_cols <= 0:
                sum_cols = max(table.viewport().width(), clone.viewport().width())
            total_w = vh_w + sum_cols + 2
            total_h = header_h + content_h + 2
            if total_w <= 2 or total_h <= 2:
                # Fallback: try rendering the live table directly using computed geometry
                # Compute geometry from the live table
                try:
                    live_h = max(header_h + sum(max(table.rowHeight(r), table.sizeHintForRow(r) or 24) for r in range(rows)) + 2, 50)
                    live_w = max(vh.width() + sum(max(table.columnWidth(c), hh.sectionSize(c)) for c in range(cols)) + 2, 200)
                    # Temporarily resize the live table for full render
                    old_size = table.size()
                    table.resize(live_w, live_h)
                    QApplication.processEvents()
                    img = QImage(live_w, live_h, QImage.Format_ARGB32)
                    img.fill(0xFFFFFFFF)
                    p = QPainter(img)
                    try:
                        table.render(p)
                    finally:
                        p.end()
                    # Restore size
                    table.resize(old_size)
                    return img, live_w, live_h
                except Exception:
                    return None, 0, 0
            # Size the clone appropriately before rendering
            clone.resize(total_w, total_h)
            clone.horizontalHeader().setVisible(True)
            clone.verticalHeader().setVisible(True)
            QApplication.processEvents()
            # Render clone to image
            img = QImage(total_w, total_h, QImage.Format_ARGB32)
            img.fill(0xFFFFFFFF)
            p = QPainter(img)
            try:
                clone.render(p)
            finally:
                p.end()
            clone.deleteLater()
            return img, total_w, total_h
        except Exception:
            return None, 0, 0

    def _load_companies(self):
        # Only filter companies when search mode is Company
        query = self.search_edit.text() if hasattr(self, 'search_edit') else ''
        search = (query or '').lower() if self.search_mode.currentText() == 'Company' else ''
        companies = [c for c in self.db.list_companies(search) if not c['is_ours']]
        current_id = self.company_combo.currentData() if self.company_combo.count() else None
        self.company_combo.blockSignals(True)
        self.company_combo.clear()
        # Add an "All Companies" option to view every competitor's MCUs
        self.company_combo.addItem('All Companies', None)
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
            # Default to All Companies
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
        # Disable sorting during population to avoid None items and blank cells
        prev_sorting = self.table.isSortingEnabled()
        if prev_sorting:
            self.table.setSortingEnabled(False)
        company_id = self.company_combo.currentData()
        query = (self.search_edit.text() or '') if hasattr(self, 'search_edit') else ''
        mode = self.search_mode.currentText()
        our_mcus_rows = self.db.list_our_mcus()
        our_mcus = [dict(r) for r in our_mcus_rows]
        feat_cols = self.db.feature_columns()
        q = (query or '')
        import re
        def _norm(s: str) -> str:
            return re.sub(r"[^a-z0-9]", "", (s or '').lower())
        qn = _norm(q)
        if mode == 'MCU' and q:
            # Global MCU search across all competitor companies
            mcus_all = []
            all_companies = self.db.list_companies('')
            companies_map = {c['id']: c['name'] for c in all_companies}
            for c in all_companies:
                if c.get('is_ours'):
                    continue
                mcus_all.extend([dict(r) for r in self.db.list_mcus_by_company(c['id'])])
        else:
            # Company context (including All Companies when company_id is None)
            all_companies = self.db.list_companies('')
            if company_id is None:
                mcus_all = []
                companies_map = {c['id']: c['name'] for c in all_companies}
                for c in all_companies:
                    if c.get('is_ours'):
                        continue
                    mcus_all.extend([dict(r) for r in self.db.list_mcus_by_company(c['id'])])
            else:
                mcus_all = [dict(r) for r in self.db.list_mcus_by_company(company_id)]
                companies_map = {company_id: next((c['name'] for c in all_companies if c['id']==company_id), '')}

        mcus = [m for m in mcus_all if (qn in _norm(m.get('name', '')))] if (mode == 'MCU' and q) else mcus_all

        self.table.setRowCount(0)
        counts = {'Best Match': 0, 'Partial': 0, 'No Match': 0}
        for mcu in mcus:
            # Build feature dicts
            # Include flags needed by similarity rules (e.g., is_dsp, is_fpga)
            target = ({k: mcu.get(k) for k in feat_cols}
                      | {'name': mcu.get('name', ''), 'is_dsp': mcu.get('is_dsp'), 'is_fpga': mcu.get('is_fpga')})
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
            # Part NO item with ID attached for sorting/selection
            part_item = QTableWidgetItem(mcu['name'])
            part_item.setData(Qt.UserRole, mcu['id'])
            self.table.setItem(row, 1, part_item)
            # Overlay a small PDF button in front of the name
            from PySide6.QtWidgets import QWidget, QHBoxLayout, QToolButton, QLabel
            cell = QWidget()
            h = QHBoxLayout(cell)
            h.setContentsMargins(4, 0, 4, 0)
            h.setSpacing(6)
            pdf_btn = QToolButton()
            pdf_btn.setToolTip('Open datasheet (local)')
            pdf_btn.setCursor(QCursor(Qt.PointingHandCursor))
            pdf_btn.setAutoRaise(True)
            # Use text instead of icon, styled as a pill button
            pdf_btn.setText('PDF')
            pdf_btn.setMinimumWidth(40)
            pdf_btn.setFixedHeight(22)
            pdf_btn.setStyleSheet(
                "QToolButton {"
                "  padding: 2px 8px;"
                "  border: 1px solid #3a6cf4;"
                "  border-radius: 11px;"
                "  background-color: rgba(58,108,244,0.15);"
                "  color: #ffffff;"
                "  font-weight: 600;"
                "}"
                "QToolButton:hover {"
                "  background-color: rgba(58,108,244,0.28);"
                "  border-color: #5e86f7;"
                "}"
                "QToolButton:pressed {"
                "  background-color: rgba(58,108,244,0.40);"
                "  border-color: #86a3fb;"
                "}"
            )
            # Capture current name in lambda default
            pdf_btn.clicked.connect(lambda _, name=mcu['name']: self._download_datasheet(name))
            name_lbl = QLabel(mcu['name'])
            name_lbl.setContentsMargins(0,0,0,0)
            name_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            h.addWidget(pdf_btn)
            h.addWidget(name_lbl)
            h.addStretch(1)
            self.table.setCellWidget(row, 1, cell)
            # Compatibility item with our_mcu_id (or None)
            compat_item = QTableWidgetItem(best['name'] if best else '-')
            compat_item.setData(Qt.UserRole, best['id'] if best else None)
            self.table.setItem(row, 2, compat_item)
            item_score = QTableWidgetItem(f"{score:.1f}")
            item_score.setTextAlignment(Qt.AlignCenter)
            # Ensure numeric sort on Match % using EditRole
            item_score.setData(Qt.EditRole, float(f"{score:.4f}"))
            self.table.setItem(row, 3, item_score)
            # Backing item for sorting by custom order
            cat_item = QTableWidgetItem(category)
            # Higher value sorts first when descending; use 2,1,0 for Best/Partial/No
            order_map = {'Best Match': 2, 'Partial': 1, 'No Match': 0}
            cat_item.setData(Qt.DisplayRole, category)  # ensure visible text is the label
            cat_item.setData(Qt.EditRole, order_map.get(category, -1))  # numeric for sorting
            cat_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 4, cat_item)
            # Category chip as QLabel overlay for visuals
            from PySide6.QtWidgets import QLabel
            chip = QLabel(category)
            chip.setAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(category, '#444')
            chip.setStyleSheet(f"QLabel {{ background-color: {color}; color: white; border-radius: 10px; padding: 2px 6px; margin: 0px; }}")
            chip.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            self.table.setCellWidget(row, 4, chip)
            self.table.setRowHeight(row, 36)
            # Tally counts
            if category in counts:
                counts[category] += 1

        # Update counts label
        self.cat_counts_label.setText(
            f"Best Match: {counts['Best Match']}   |   Partial: {counts['Partial']}   |   No Match: {counts['No Match']}"
        )
        # Restore previous sorting state
        if prev_sorting:
            self.table.setSortingEnabled(True)

    def _download_datasheet(self, mcu_name: str):
        # Open a local PDF from the datasheets folder by matching the name.
        import os, re
        datasheet_dir = r'e:\MCU-com\datasheets'
        try:
            if not os.path.isdir(datasheet_dir):
                os.makedirs(datasheet_dir, exist_ok=True)
            raw = (mcu_name or '').strip()
            # Primary key: substring before '-' (e.g., 'XC7A200T' from 'XC7A200T-2FBG6761 (FPGA)')
            base = raw.split('-', 1)[0].strip()
            # Fallback if no '-' present: strip at first space or '('
            if base == raw:
                base = re.split(r"\s|\(", raw, maxsplit=1)[0]
            norm_base = re.sub(r'[^a-z0-9]+', '', base.lower())
            norm_full = re.sub(r'[^a-z0-9]+', '', raw.lower())
            exact = None
            partial = []
            for fn in os.listdir(datasheet_dir):
                if not fn.lower().endswith('.pdf'):
                    continue
                key = re.sub(r'[^a-z0-9]+', '', os.path.splitext(fn)[0].lower())
                # Try exact base match first
                if norm_base and key == norm_base:
                    exact = os.path.join(datasheet_dir, fn)
                    break
                # Then exact full-name match
                if norm_full and key == norm_full:
                    exact = os.path.join(datasheet_dir, fn)
                    break
                # Then substring matches (prefer base)
                if norm_base and norm_base in key:
                    partial.append(os.path.join(datasheet_dir, fn))
                    continue
                if norm_full and norm_full in key:
                    partial.append(os.path.join(datasheet_dir, fn))
            target = exact or (partial[0] if len(partial) == 1 else None)
            if target:
                os.startfile(target)
                return
            # If multiple candidates, open folder for manual choice without alert
            if len(partial) > 1:
                os.startfile(datasheet_dir)
                return
            # None found: notify only (do not open folder)
            QMessageBox.information(self, 'Datasheet', 'Datasheet is not available ')
        except Exception:
            try:
                # Notify only on failure
                QMessageBox.information(self, 'Datasheet', 'Datasheet is not available ')
            except Exception:
                pass

    def _open_add_dialog(self):
        dlg = AddMCUDialog(self.db, self)
        if dlg.exec():
            self._refresh_table()

    def _open_details(self, row: int, column: int):
        part_item = self.table.item(row, 1)
        compat_item = self.table.item(row, 2)
        comp_id = part_item.data(Qt.UserRole) if part_item else None
        our_id = self._selected_our_mcu_id() or (compat_item.data(Qt.UserRole) if compat_item else None)
        if comp_id is None or our_id is None:
            QMessageBox.information(self, 'Details', 'No match available for detailed comparison.')
            return
        # Resolve potential ID collision: if comp_id points to Our Company due to non-unique IDs,
        # re-resolve competitor MCU by company name and part name from the row.
        try:
            comp_rec = self.db.get_mcu_by_id(int(comp_id))
        except Exception:
            comp_rec = None
        if comp_rec and int(self.db.get_company_by_id(int(comp_rec.get('company_id'))) .get('is_ours', 0)) == 1:
            # Look up competitor company by the table's Manufacturer column text
            company_name = self.table.item(row, 0).text() if self.table.item(row, 0) else ''
            # Find company id by name
            comp_id_fixed = None
            for c in self.db.list_companies(''):
                if not c.get('is_ours') and c.get('name') == company_name:
                    # Find MCU by exact part name within this company
                    part_name = part_item.text() if part_item else ''
                    for m in self.db.list_mcus_by_company(c['id']):
                        if m.get('name') == part_name:
                            comp_id_fixed = m.get('id')
                            break
                    break
            if comp_id_fixed is not None:
                comp_id = comp_id_fixed
        # Prevent comparing the same MCU against itself
        try:
            if int(comp_id) == int(our_id):
                QMessageBox.information(self, 'Details', 'Cannot compare an MCU against itself.')
                return
        except Exception:
            pass
        dlg = DetailsDialog(self.db, comp_id, our_id, self)
        # Track active details dialog so File->Print/Export can target it
        self._details_dialog = dlg
        try:
            dlg.finished.connect(lambda _=None: setattr(self, '_details_dialog', None))
        except Exception:
            pass
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

    def _delete_selected_mcu(self):
        # Delete the selected competitor MCU from the Compare table
        if self.table.currentRow() < 0:
            return
        row = self.table.currentRow()
        item = self.table.item(row, 1)
        if item is None:
            return
        mcu_id = item.data(Qt.UserRole)
        name = item.text() if item else ''
        if mcu_id is None:
            return
        resp = QMessageBox.question(self, 'Delete MCU', f"Delete MCU '{name}'? This will remove related NCO entries.", QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        if self.db.delete_mcu(int(mcu_id)):
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

    def _rename_company(self):
        from PySide6.QtWidgets import QInputDialog
        if self.company_combo.count() == 0:
            return
        company_id = self.company_combo.currentData()
        if company_id is None:
            QMessageBox.information(self, 'Rename Company', 'Select a specific company, not All Companies.')
            return
        current_name = self.company_combo.currentText()
        name, ok = QInputDialog.getText(self, 'Rename Company', 'New name:', text=current_name)
        if ok and name.strip():
            if self.db.update_company_name(int(company_id), name.strip()):
                self._load_companies()
                # keep the same company selected after rename
                idx = self.company_combo.findData(company_id)
                if idx >= 0:
                    self.company_combo.setCurrentIndex(idx)
                self._refresh_table()

    def _delete_company(self):
        if self.company_combo.count() == 0:
            return
        company_id = self.company_combo.currentData()
        if company_id is None:
            QMessageBox.information(self, 'Delete Company', 'Select a specific company, not All Companies.')
            return
        comp = self.db.get_company_by_id(int(company_id))
        if not comp:
            return
        # Safety: do not delete Our Company
        if int(comp.get('is_ours', 0)) == 1:
            QMessageBox.information(self, 'Delete Company', 'Cannot delete Our Company.')
            return
        name = comp.get('name', '')
        resp = QMessageBox.question(self, 'Delete Company', f"Delete company '{name}' and all its MCUs and related NCO entries?", QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        if self.db.delete_company(int(company_id)):
            self._load_companies()
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
        row.addWidget(QLabel('NCO/Commission:'))
        self.nco_org_combo = QComboBox()
        self.nco_org_combo.currentIndexChanged.connect(self._refresh_nco_table)
        self.nco_org_combo.currentIndexChanged.connect(self._update_nco_actions)
        row.addWidget(self.nco_org_combo, 2)
        self.nco_add_btn = QPushButton('Add Entry')
        self.nco_add_btn.clicked.connect(self._open_nco_add)
        row.addWidget(self.nco_add_btn)
        # Search controls (mode + query)
        row.addWidget(QLabel('Search:'))
        self.nco_search_mode = QComboBox()
        self.nco_search_mode.addItems(['All', 'MCU', 'NCO/Commission', 'Company', 'Quantity'])
        self.nco_search_mode.currentIndexChanged.connect(self._on_nco_search_mode_change)
        row.addWidget(self.nco_search_mode)
        self.nco_search = QLineEdit()
        self.nco_search.setPlaceholderText('Type to search...')
        self.nco_search.textChanged.connect(self._refresh_nco_table)
        row.addWidget(self.nco_search, 2)
        # Edit button
        self.nco_edit_btn = QPushButton('Edit Selected')
        self.nco_edit_btn.clicked.connect(self._edit_selected_nco)
        row.addWidget(self.nco_edit_btn)
        refresh_btn = QPushButton('Refresh')
        refresh_btn.clicked.connect(self._refresh_nco_table)
        row.addWidget(refresh_btn)
        row.addStretch(1)
        v.addLayout(row)

        # NCO table
        self.nco_table = QTableWidget(0, 7)
        self.nco_table.setHorizontalHeaderLabels(['NCO/Commission', 'Company', 'Competitor MCU', 'Quantity', 'Our MCU', 'Match %', 'Category'])
        # Column sizing: make 'Competitor MCU' wider, others fit content
        header = self.nco_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Competitor MCU
        self.nco_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.nco_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.nco_table.setAlternatingRowColors(True)
        self.nco_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.nco_table.cellDoubleClicked.connect(self._open_nco_details)
        # Enable header sorting
        self.nco_table.setSortingEnabled(True)
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
        mode = self.nco_search_mode.currentText() if hasattr(self, 'nco_search_mode') else 'All'
        def matches(row: Dict[str, Any]) -> bool:
            if not q:
                return True
            nco = orgs.get(row.get('org_id'), '')
            comp = comps.get(row.get('company_id'), '')
            cmcu = mcu_name.get(row.get('comp_mcu_id'), '')
            omcu = mcu_name.get(row.get('our_mcu_id'), '')
            qty = str(row.get('quantity', ''))
            if mode == 'All':
                return (q in nco.lower()) or (q in comp.lower()) or (q in str(cmcu).lower()) or (q in str(omcu).lower()) or (q in qty.lower())
            if mode == 'MCU':
                return (q in str(cmcu).lower()) or (q in str(omcu).lower())
            if mode == 'NCO/Commission':
                return q in nco.lower()
            if mode == 'Company':
                return q in comp.lower()
            if mode == 'Quantity':
                return q in qty.lower()
            return True

        filtered = [r for r in rows if matches(r)]
        # Temporarily disable sorting during population to avoid items shifting rows mid-insert
        prev_sorting = self.nco_table.isSortingEnabled()
        if prev_sorting:
            self.nco_table.setSortingEnabled(False)
        self.nco_table.setRowCount(0)
        for r in filtered:
            row = self.nco_table.rowCount()
            self.nco_table.insertRow(row)
            self.nco_table.setItem(row, 0, QTableWidgetItem(orgs.get(r.get('org_id'), '')))
            self.nco_table.setItem(row, 1, QTableWidgetItem(comps.get(r.get('company_id'), '')))
            comp_name = mcu_name.get(r.get('comp_mcu_id'), '')
            # Add a small PDF button before the competitor MCU name
            from PySide6.QtWidgets import QWidget, QHBoxLayout, QToolButton, QLabel
            pdf_cell = QWidget()
            ph = QHBoxLayout(pdf_cell)
            ph.setContentsMargins(4, 0, 4, 0)
            ph.setSpacing(6)
            pdf_btn = QToolButton()
            pdf_btn.setText('PDF')
            pdf_btn.setCursor(QCursor(Qt.PointingHandCursor))
            pdf_btn.setAutoRaise(True)
            pdf_btn.setMinimumWidth(40)
            pdf_btn.setFixedHeight(22)
            pdf_btn.setToolTip('Open datasheet (local)')
            pdf_btn.setStyleSheet(
                "QToolButton {"
                "  padding: 2px 8px;"
                "  border: 1px solid #3a6cf4;"
                "  border-radius: 11px;"
                "  background-color: rgba(58,108,244,0.15);"
                "  color: #ffffff;"
                "  font-weight: 600;"
                "}"
                "QToolButton:hover { background-color: rgba(58,108,244,0.28); border-color: #5e86f7; }"
                "QToolButton:pressed { background-color: rgba(58,108,244,0.40); border-color: #86a3fb; }"
            )
            # Keep name label next to button
            name_lbl = QLabel(comp_name)
            name_lbl.setContentsMargins(0,0,0,0)
            name_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            # Connect button click to datasheet open
            pdf_btn.clicked.connect(lambda _, nm=comp_name: self._download_datasheet(nm))
            ph.addWidget(pdf_btn)
            ph.addWidget(name_lbl)
            ph.addStretch(1)
            self.nco_table.setCellWidget(row, 2, pdf_cell)
            self.nco_table.setItem(row, 3, QTableWidgetItem(str(r.get('quantity', 0))))

            # Determine our MCU and compute similarity
            comp = self.db.get_mcu_by_id(int(r.get('comp_mcu_id')))
            target = (({k: comp.get(k) for k in feat_cols}
                      | {'name': comp.get('name', ''), 'is_dsp': comp.get('is_dsp'), 'is_fpga': comp.get('is_fpga')})
                      if comp else {})
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
            # Ensure numeric sort on Match % using EditRole
            item_score.setData(Qt.EditRole, float(f"{score:.4f}"))
            self.nco_table.setItem(row, 5, item_score)
            cat = categorize(score)
            # Backing item for category sorting by custom order
            order_map = {'Best Match': 2, 'Partial': 1, 'No Match': 0}
            cat_item = QTableWidgetItem(cat)
            cat_item.setData(Qt.DisplayRole, cat)  # ensure visible label
            cat_item.setTextAlignment(Qt.AlignCenter)
            cat_item.setData(Qt.EditRole, order_map.get(cat, -1))
            self.nco_table.setItem(row, 6, cat_item)
            from PySide6.QtWidgets import QLabel
            chip = QLabel(cat)
            chip.setAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(cat, '#444')
            chip.setStyleSheet(f"QLabel {{ background-color: {color}; color: white; border-radius: 10px; padding: 2px 6px; margin: 0px; }}")
            chip.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            self.nco_table.setCellWidget(row, 6, chip)
            # Fit row for PDF button height
            try:
                self.nco_table.setRowHeight(row, 36)
            except Exception:
                pass
            # Store entry id on first column
            if self.nco_table.item(row, 0):
                self.nco_table.item(row, 0).setData(Qt.UserRole, r.get('id'))
        # Restore sorting state
        if prev_sorting:
            self.nco_table.setSortingEnabled(True)

    def _on_nco_search_mode_change(self):
        # Update placeholder and refresh table according to selected search mode
        if not hasattr(self, 'nco_search'):
            return
        mode = self.nco_search_mode.currentText()
        placeholders = {
            'All': 'Type to search Org/Company/MCU/Quantity',
            'MCU': 'Type to search by Competitor/Our MCU',
            'NCO/Commission': 'Type to search by Organization',
            'Company': 'Type to search by Company',
            'Quantity': 'Type to search by Quantity',
        }
        self.nco_search.setPlaceholderText(placeholders.get(mode, 'Type to search...'))
        self._refresh_nco_table()

    def _load_nco_orgs(self):
        if not hasattr(self, 'nco_org_combo'):
            return
        orgs = self.db.list_nco_orgs()
        current = self.nco_org_combo.currentData() if self.nco_org_combo.count() else None
        self.nco_org_combo.blockSignals(True)
        self.nco_org_combo.clear()
        # Add an "All Organizations" option to view all entries
        self.nco_org_combo.addItem('All Organizations', None)
        for o in orgs:
            self.nco_org_combo.addItem(o['name'], o['id'])
        self.nco_org_combo.blockSignals(False)
        if current is not None:
            idx = self.nco_org_combo.findData(current)
            if idx >= 0:
                self.nco_org_combo.setCurrentIndex(idx)
        elif self.nco_org_combo.count() > 0:
            # default to "All Organizations"
            self.nco_org_combo.setCurrentIndex(0)
        # Ensure table refresh after loading orgs
        self._refresh_nco_table()

    def _add_nco_org(self):
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, 'Add Organization', 'Organization Name:')
        if ok and name.strip():
            self.db.add_nco_org(name.strip())
            self._load_nco_orgs()
            self._refresh_nco_table()

    def _rename_nco_org(self):
        from PySide6.QtWidgets import QInputDialog
        if not hasattr(self, 'nco_org_combo') or self.nco_org_combo.count() == 0:
            return
        org_id = self.nco_org_combo.currentData()
        if org_id is None:
            QMessageBox.information(self, 'Rename Organization', 'Select a specific organization, not All Organizations.')
            return
        current_name = self.nco_org_combo.currentText()
        name, ok = QInputDialog.getText(self, 'Rename Organization', 'New name:', text=current_name)
        if ok and name.strip():
            if self.db.update_nco_org(int(org_id), name.strip()):
                self._load_nco_orgs()
                self._refresh_nco_table()

    def _delete_nco_org(self):
        if not hasattr(self, 'nco_org_combo') or self.nco_org_combo.count() == 0:
            return
        org_id = self.nco_org_combo.currentData()
        if org_id is None:
            QMessageBox.information(self, 'Delete Organization', 'Select a specific organization, not All Organizations.')
            return
        name = self.nco_org_combo.currentText()
        resp = QMessageBox.question(self, 'Delete Organization', f"Delete organization '{name}' and all its NCO entries?", QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        if self.db.delete_nco_org(int(org_id)):
            self._load_nco_orgs()
            self._refresh_nco_table()

    def _update_compare_actions(self):
        # Disable rename/delete when 'All Companies' is selected
        selected_specific = self.company_combo.currentData() is not None
        for act in [self.act_rename_company, self.act_delete_company]:
            if act:
                act.setEnabled(selected_specific)
                act.setVisible(selected_specific if self.tabs.currentIndex() == 0 else False)

    def _update_nco_actions(self):
        # Disable org edit/delete when 'All Organizations' is selected
        is_specific = self.nco_org_combo.currentData() is not None
        for act in [self.act_nco_rename_org, self.act_nco_delete_org]:
            if act:
                act.setEnabled(is_specific)
                act.setVisible(is_specific if self.tabs.currentIndex() == 1 else False)

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

    def _delete_selected_nco(self):
        if not hasattr(self, 'nco_table') or self.nco_table.currentRow() < 0:
            return
        row = self.nco_table.currentRow()
        entry_id = self.nco_table.item(row, 0).data(Qt.UserRole) if self.nco_table.item(row, 0) else None
        if entry_id is None:
            return
        resp = QMessageBox.question(self, 'Delete Entry', 'Delete selected NCO/Commission entry?', QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        if self.db.delete_nco_entry(int(entry_id)):
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
            # Guard: if best match equals competitor, do not open self-compare
            if best and int(best['id']) == int(comp_id):
                QMessageBox.information(self, 'Details', 'Cannot compare an MCU against itself.')
                return
            our_id = best['id'] if best else None
        if our_id is None:
            return
        # Final guard against same IDs
        try:
            if int(our_id) == int(comp_id):
                QMessageBox.information(self, 'Details', 'Cannot compare an MCU against itself.')
                return
        except Exception:
            pass
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

    def _on_tab_changed(self, index: int):
        # 0 = Compare, 1 = NCO/Commission
        is_compare = (index == 0)
        # Compare actions/buttons enabled only on Compare tab
        for act in [self.act_add_mcu, self.act_add_company, self.act_rename_company, self.act_delete_company]:
            if act:
                act.setEnabled(is_compare)
                act.setVisible(is_compare)
        for w in [self.btn_edit_selected, self.btn_delete_selected, self.company_combo, self.search_mode, self.search_edit, self.compare_combo, self.cat_counts_label]:
            if w:
                w.setEnabled(is_compare)
                w.setVisible(is_compare)
        # NCO actions/buttons enabled only on NCO tab
        is_nco = not is_compare
        for act in [self.act_nco_add_org, self.act_nco_rename_org, self.act_nco_delete_org, self.act_nco_add_entry, self.act_nco_edit_selected, self.act_nco_delete_entry, self.act_nco_view]:
            if act:
                act.setEnabled(is_nco)
                act.setVisible(is_nco)
        for w in [self.nco_org_combo, self.nco_add_btn, self.nco_search_mode, self.nco_search, self.nco_edit_btn]:
            if hasattr(self, 'nco_org_combo') and w:
                w.setEnabled(is_nco)
                w.setVisible(is_nco)
