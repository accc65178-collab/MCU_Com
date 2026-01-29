from typing import Dict, Any
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox, QSpinBox,
    QPushButton, QWidget, QFormLayout, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QProgressBar, QStackedLayout, QScrollArea
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QGuiApplication, QIntValidator, QPalette, QDoubleValidator, QTextDocument, QKeySequence, QShortcut
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtCharts import QChart, QChartView, QPieSeries

from mcu_compare.engine.similarity import weighted_similarity, categorize


FEATURE_FIELDS = [
    ('core', 'Core', 'text'),
    ('core_mark', 'Core mark', 'int'),
    ('dsp_core', 'DSP core', 'bool'),
    ('fpu', 'FPU', 'enum_fpu'),
    ('max_clock_mhz', 'Max clock (MHz)', 'float'),
    ('flash_kb', 'Flash (KB)', 'float'),
    ('sram_kb', 'SRAM (KB)', 'float'),
    ('eeprom', 'EEPROM', 'bool'),
    ('gpios', 'GPIOs', 'int'),
    ('uarts', 'UART/USART', 'int'),
    ('spis', 'SPI', 'int'),
    ('i2cs', 'I2C', 'int'),
    ('pwms', 'PWM', 'int'),
    ('timers', 'Timers', 'int'),
    ('dacs', 'DAC', 'int'),
    ('adcs', 'ADC', 'int'),
    ('cans', 'CAN', 'int'),
    ('power_mgmt', 'Power management', 'bool'),
    ('clock_mgmt', 'Clock management', 'bool'),
    ('qei', 'Quadrature encoder', 'bool'),
    ('internal_osc', 'Internal oscillator', 'bool'),
    ('security_features', 'Security features', 'bool'),
    ('output_compare', 'Output Compare', 'int'),
    ('input_capture', 'Input Capture', 'bool'),
    ('qspi', 'QSPI', 'int'),
    ('ethernet', 'Ethernet', 'bool'),
    ('emif', 'EMIF', 'bool'),
    ('spi_slave', 'SPI Slave', 'bool'),
    ('ext_interrupts', 'External Interrupts', 'int'),
]


# Module-level safe HTML builders for DetailsDialog printing/exporting
def _details_html_safe(dlg) -> str:
    try:
        return _details_html_build(dlg)
    except Exception:
        return ''


def _details_html_build(dlg) -> str:
    # Use dialog header text and table contents to build a simple HTML report
    try:
        title = dlg.header.text() if hasattr(dlg, 'header') else 'MCU Comparison'
    except Exception:
        title = 'MCU Comparison'
    tbl = getattr(dlg, 'table', None)
    if tbl is None:
        return ''
    cols = tbl.columnCount()
    rows = tbl.rowCount()
    headers = [(tbl.horizontalHeaderItem(c).text() if tbl.horizontalHeaderItem(c) else '') for c in range(cols)]
    from PySide6.QtWidgets import QLabel, QProgressBar
    def cell_text(r: int, c: int) -> str:
        w = tbl.cellWidget(r, c)
        if w is not None:
            if isinstance(w, QLabel) and w.text():
                return w.text()
            if isinstance(w, QProgressBar):
                try:
                    return f"{w.value()}%"
                except Exception:
                    return w.text() or ''
            lbl = w.findChild(QLabel)
            if lbl is not None and lbl.text():
                return lbl.text()
        it = tbl.item(r, c)
        if it is not None and it.text():
            return it.text()
        return ''
    html = [
        '<html><head><meta charset="utf-8">',
        '<style>body{font:13pt Segoe UI,Arial;margin:10mm;} .title{font-size:16pt;font-weight:700;margin:6pt 0;} ',
        'table{border-collapse:collapse;width:100%;table-layout:fixed;} th,td{border:1px solid #888;padding:6pt 8pt;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;} ',
        'th{background:#eaeaea;}</style></head><body>',
        f'<div class="title">{title}</div>',
        '<table>'
    ]
    html.append('<tr>')
    for h in headers:
        html.append(f'<th>{h}</th>')
    html.append('</tr>')
    for r in range(rows):
        html.append('<tr>')
        for c in range(cols):
            txt = cell_text(r, c)
            txt = (txt or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            html.append(f'<td>{txt}</td>')
        html.append('</tr>')
    html.append('</table></body></html>')
    return ''.join(html)


class AddMCUDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle('Add MCU')
        self._build_ui()

    def _build_ui(self):
        v = QVBoxLayout(self)

        form = QFormLayout()

        self.company_combo = QComboBox()
        for c in self.db.list_companies():
            self.company_combo.addItem(c['name'], c['id'])
        form.addRow('Company', self.company_combo)

        self.name_edit = QLineEdit()
        form.addRow('MCU Name', self.name_edit)

        self.inputs: Dict[str, QWidget] = {}
        for key, label, typ in FEATURE_FIELDS:
            if typ == 'text':
                w = QLineEdit()
            elif typ == 'int':
                w = QLineEdit()
                w.setValidator(QIntValidator(0, 100000, self))
                w.setText('0')
            elif typ == 'float':
                w = QLineEdit()
                dv = QDoubleValidator(0.0, 1e9, 3, self)
                dv.setNotation(QDoubleValidator.StandardNotation)
                w.setValidator(dv)
                w.setText('0')
            elif typ == 'bool':
                w = QCheckBox()
            elif typ == 'enum_fpu':
                w = QComboBox()
                w.addItems(['None', 'Single precision', 'Double precision'])
            else:
                w = QLineEdit()
            self.inputs[key] = w
            form.addRow(label, w)

        v.addLayout(form)

        btn_row = QHBoxLayout()
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

    def _save(self):
        company_id = self.company_combo.currentData()
        name = self.name_edit.text().strip()
        if not name:
            self.name_edit.setFocus()
            return
        payload = {'name': name}
        for key, label, typ in FEATURE_FIELDS:
            if typ == 'text':
                payload[key] = str(self.inputs[key].text()).strip()
            elif typ == 'int':
                try:
                    payload[key] = int(self.inputs[key].text() or '0')
                except Exception:
                    payload[key] = 0
            elif typ == 'float':
                try:
                    payload[key] = float(self.inputs[key].text() or '0')
                except Exception:
                    payload[key] = 0.0
            elif typ == 'bool':
                payload[key] = 1 if self.inputs[key].isChecked() else 0
            elif typ == 'enum_fpu':
                idx = self.inputs[key].currentIndex()
                payload[key] = idx
            else:
                payload[key] = str(self.inputs[key].text()).strip()
        self.db.insert_mcu(company_id, payload)
        self.accept()


class AddNcoEntryDialog(QDialog):
    def __init__(self, db, parent=None, org_id: int | None = None):
        super().__init__(parent)
        self.db = db
        self._org_id = org_id
        self.setWindowTitle('Add NCO/Commission Entry')
        self._build_ui()

    def _build_ui(self):
        # Wrap full content in a scroll area so the whole dialog scrolls, not the table
        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        v = QVBoxLayout(content)
        scroll.setWidget(content)
        outer.addWidget(scroll)
        form = QFormLayout()

        # Competitor company selector
        self.company_combo = QComboBox()
        self._companies = [c for c in self.db.list_companies('') if not c.get('is_ours')]
        for c in self._companies:
            self.company_combo.addItem(c['name'], c['id'])
        self.company_combo.currentIndexChanged.connect(self._reload_comp_mcUs)
        form.addRow('Company', self.company_combo)

        # Competitor MCU selector
        self.comp_mcu_combo = QComboBox()
        form.addRow('Competitor MCU', self.comp_mcu_combo)

        # Quantity
        self.qty_spin = QSpinBox()
        self.qty_spin.setRange(0, 10_000_000)
        form.addRow('Quantity', self.qty_spin)

        # Our MCU selection removed — will be auto-selected by code

        v.addLayout(form)

        btn_row = QHBoxLayout()
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

        # Initialize MCUs list for the first company
        self._reload_comp_mcUs()

    def _reload_comp_mcUs(self):
        cid = self.company_combo.currentData()
        self.comp_mcu_combo.clear()
        if cid is None:
            return
        for m in self.db.list_mcus_by_company(cid):
            self.comp_mcu_combo.addItem(m['name'], m['id'])

    def _save(self):
        company_id = self.company_combo.currentData()
        comp_mcu_id = self.comp_mcu_combo.currentData()
        quantity = int(self.qty_spin.value())
        if company_id is None or comp_mcu_id is None:
            return
        # Auto-match will be computed by code; save without our_mcu_id
        self.db.add_nco_entry(company_id, comp_mcu_id, quantity, None, '', self._org_id)
        self.accept()


class ViewNcoEntriesDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle('NCO/Commission Entries')
        self.resize(900, 500)
        self._build_ui()

    def _build_ui(self):
        v = QVBoxLayout(self)
        rows = self.db.list_nco_entries()
        table = QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(['Company', 'Competitor MCU', 'Quantity', 'Our MCU'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Build lookup maps
        comps = {c['id']: c['name'] for c in self.db.list_companies('')}
        mcu_name = {}
        for c in self.db.list_companies(''):
            for m in self.db.list_mcus_by_company(c['id']):
                mcu_name[m['id']] = m['name']

        for i, r in enumerate(rows):
            table.setItem(i, 0, QTableWidgetItem(comps.get(r.get('company_id'), '')))
            table.setItem(i, 1, QTableWidgetItem(mcu_name.get(r.get('comp_mcu_id'), '')))
            table.setItem(i, 2, QTableWidgetItem(str(r.get('quantity', 0))))
            table.setItem(i, 3, QTableWidgetItem(mcu_name.get(r.get('our_mcu_id'), '')))

        # Make the table display all rows with no internal scrolling; the dialog scrolls instead
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        try:
            total_h = table.horizontalHeader().height() + sum(table.rowHeight(r) for r in range(table.rowCount())) + 2
            table.setMinimumHeight(total_h)
            table.setMaximumHeight(total_h)
        except Exception:
            pass
        # Make the table display all rows with no internal scrolling; the dialog scrolls instead
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        try:
            total_h = table.horizontalHeader().height() + sum(table.rowHeight(r) for r in range(table.rowCount())) + 2
            table.setMinimumHeight(total_h)
            table.setMaximumHeight(total_h)
        except Exception:
            pass
        # Disable internal scrolling so the outer dialog scrolls instead
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        try:
            t_height = table.horizontalHeader().height() + sum(table.rowHeight(r) for r in range(table.rowCount())) + 2
            table.setMinimumHeight(t_height)
            table.setMaximumHeight(t_height)
        except Exception:
            pass
        v.addWidget(table)


        btn_row = QHBoxLayout()
        print_btn = QPushButton('Print')
        export_btn = QPushButton('Export PDF')
        close_btn = QPushButton('Close')
        print_btn.clicked.connect(self.print_details)
        export_btn.clicked.connect(self.export_details_pdf)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(print_btn)
        btn_row.addWidget(export_btn)
        btn_row.addWidget(close_btn)
        v.addLayout(btn_row)
        # Shortcuts similar to main window
        sc_print = QShortcut(QKeySequence.Print, self)
        sc_print.activated.connect(self.print_details)
        sc_export = QShortcut(QKeySequence('Ctrl+Shift+P'), self)
        sc_export.activated.connect(self.export_details_pdf)

    def _details_html(self) -> str:
        # Header with names and overall percentage
        title = self.header.text()
        # Build table HTML from self.table
        tbl = self.table
        cols = tbl.columnCount()
        rows = tbl.rowCount()
        headers = [(tbl.horizontalHeaderItem(c).text() if tbl.horizontalHeaderItem(c) else '') for c in range(cols)]
        from PySide6.QtWidgets import QLabel, QProgressBar
        def cell_text(r: int, c: int) -> str:
            # Prefer widget contents
            w = tbl.cellWidget(r, c)
            if w is not None:
                if isinstance(w, QLabel) and w.text():
                    return w.text()
                if isinstance(w, QProgressBar):
                    try:
                        return f"{w.value()}%"
                    except Exception:
                        return w.text() or ''
                lbl = w.findChild(QLabel)
                if lbl is not None and lbl.text():
                    return lbl.text()
            it = tbl.item(r, c)
            if it is not None and it.text():
                return it.text()
            return ''
        html = [
            '<html><head><meta charset="utf-8">',
            '<style>body{font:12px Segoe UI,Arial;} .title{font-size:16px;font-weight:700;margin:6px 0;} ',
            'table{border-collapse:collapse;width:100%;} th,td{border:1px solid #888;padding:4px 6px;text-align:left;} ',
            'th{background:#eaeaea;}</style></head><body>',
            f'<div class="title">{title}</div>',
            '<table>'
        ]
        html.append('<tr>')
        for h in headers:
            html.append(f'<th>{h}</th>')
        html.append('</tr>')
        for r in range(rows):
            html.append('<tr>')
            for c in range(cols):
                txt = cell_text(r, c)
                txt = (txt or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                html.append(f'<td>{txt}</td>')
            html.append('</tr>')
        html.append('</table></body></html>')
        return ''.join(html)

    def _print_details(self):
        html = self._details_html()
        if not html:
            return
        printer = QPrinter(QPrinter.HighResolution)
        dlg = QPrintDialog(printer, self)
        dlg.setWindowTitle('Print Details')
        if dlg.exec() != QPrintDialog.Accepted:
            return
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

    def _details_html(self) -> str:
        # Build table HTML from self.table and header
        title = self.header.text() if hasattr(self, 'header') else 'MCU Comparison'
        tbl = self.table if hasattr(self, 'table') else None
        if tbl is None:
            return ''
        cols = tbl.columnCount()
        rows = tbl.rowCount()
        headers = [(tbl.horizontalHeaderItem(c).text() if tbl.horizontalHeaderItem(c) else '') for c in range(cols)]
        from PySide6.QtWidgets import QLabel, QProgressBar
        def cell_text(r: int, c: int) -> str:
            w = tbl.cellWidget(r, c)
            if w is not None:
                if isinstance(w, QLabel) and w.text():
                    return w.text()
                if isinstance(w, QProgressBar):
                    try:
                        return f"{w.value()}%"
                    except Exception:
                        return w.text() or ''
                lbl = w.findChild(QLabel)
                if lbl is not None and lbl.text():
                    return lbl.text()
            it = tbl.item(r, c)
            if it is not None and it.text():
                return it.text()
            return ''
        html = [
            '<html><head><meta charset="utf-8">',
            '<style>body{font:12px Segoe UI,Arial;} .title{font-size:16px;font-weight:700;margin:6px 0;} ',
            'table{border-collapse:collapse;width:100%;} th,td{border:1px solid #888;padding:4px 6px;text-align:left;} ',
            'th{background:#eaeaea;}</style></head><body>',
            f'<div class="title">{title}</div>',
            '<table>'
        ]
        html.append('<tr>')
        for h in headers:
            html.append(f'<th>{h}</th>')
        html.append('</tr>')
        for r in range(rows):
            html.append('<tr>')
            for c in range(cols):
                txt = cell_text(r, c)
                txt = (txt or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                html.append(f'<td>{txt}</td>')
            html.append('</tr>')
        html.append('</table></body></html>')
        return ''.join(html)

    def _details_html_safe(self) -> str:
        try:
            return self._details_html()
        except Exception:
            return ''

    def _details_html_safe(self) -> str:
        # Try existing builder first
        try:
            return self._details_html()
        except Exception:
            pass
        # Fallback: build minimal HTML from header + table
        try:
            title = self.header.text() if hasattr(self, 'header') else 'MCU Comparison'
            tbl = self.table if hasattr(self, 'table') else None
            if tbl is None:
                return ''
            cols = tbl.columnCount()
            rows = tbl.rowCount()
            headers = [(tbl.horizontalHeaderItem(c).text() if tbl.horizontalHeaderItem(c) else '') for c in range(cols)]
            from PySide6.QtWidgets import QLabel, QProgressBar
            def cell_text(r: int, c: int) -> str:
                w = tbl.cellWidget(r, c)
                if w is not None:
                    if isinstance(w, QLabel) and w.text():
                        return w.text()
                    if isinstance(w, QProgressBar):
                        try:
                            return f"{w.value()}%"
                        except Exception:
                            return w.text() or ''
                    lbl = w.findChild(QLabel)
                    if lbl is not None and lbl.text():
                        return lbl.text()
                it = tbl.item(r, c)
                if it is not None and it.text():
                    return it.text()
                return ''
            html = [
                '<html><head><meta charset="utf-8">',
                '<style>body{font:12px Segoe UI,Arial;} .title{font-size:16px;font-weight:700;margin:6px 0;} ',
                'table{border-collapse:collapse;width:100%;} th,td{border:1px solid #888;padding:4px 6px;text-align:left;} ',
                'th{background:#eaeaea;}</style></head><body>',
                f'<div class="title">{title}</div>',
                '<table>'
            ]
            html.append('<tr>')
            for h in headers:
                html.append(f'<th>{h}</th>')
            html.append('</tr>')
            for r in range(rows):
                html.append('<tr>')
                for c in range(cols):
                    txt = cell_text(r, c)
                    txt = (txt or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    html.append(f'<td>{txt}</td>')
                html.append('</tr>')
            html.append('</table></body></html>')
            return ''.join(html)
        except Exception:
            return ''

    def _export_details_pdf(self):
        from PySide6.QtWidgets import QFileDialog
        html = self._details_html()
        if not html:
            return
        path, _ = QFileDialog.getSaveFileName(self, 'Export Details to PDF', 'Details.pdf', 'PDF Files (*.pdf)')
        if not path:
            return
        if not path.lower().endswith('.pdf'):
            path += '.pdf'
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
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


class EditNcoEntryDialog(QDialog):
    def __init__(self, db, entry: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.db = db
        self.entry = dict(entry)
        self.setWindowTitle('Edit NCO/Commission Entry')
        self._build_ui()

    def _build_ui(self):
        v = QVBoxLayout(self)
        form = QFormLayout()

        # Competitor company selector
        self.company_combo = QComboBox()
        self._companies = [c for c in self.db.list_companies('') if not c.get('is_ours')]
        for c in self._companies:
            self.company_combo.addItem(c['name'], c['id'])
        form.addRow('Company', self.company_combo)

        # Competitor MCU selector
        self.comp_mcu_combo = QComboBox()
        form.addRow('Competitor MCU', self.comp_mcu_combo)

        # Quantity
        self.qty_spin = QSpinBox()
        self.qty_spin.setRange(0, 10_000_000)
        form.addRow('Quantity', self.qty_spin)

        # Our MCU (optional)
        self.our_mcu_combo = QComboBox()
        self.our_mcu_combo.addItem('None', None)
        for r in self.db.list_our_mcus():
            self.our_mcu_combo.addItem(r['name'], r['id'])
        form.addRow('Our MCU (replacement)', self.our_mcu_combo)

        v.addLayout(form)

        btn_row = QHBoxLayout()
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

        # Prefill
        # Set company and load MCUs
        comp_id = int(self.entry.get('company_id'))
        idx = self.company_combo.findData(comp_id)
        if idx >= 0:
            self.company_combo.setCurrentIndex(idx)
        self._reload_comp_mcUs()
        # Set competitor MCU
        mm = int(self.entry.get('comp_mcu_id'))
        midx = self.comp_mcu_combo.findData(mm)
        if midx >= 0:
            self.comp_mcu_combo.setCurrentIndex(midx)
        # Quantity
        self.qty_spin.setValue(int(self.entry.get('quantity', 0)))
        # Our MCU
        our_id = self.entry.get('our_mcu_id', None)
        if our_id is not None:
            oidx = self.our_mcu_combo.findData(int(our_id))
            if oidx >= 0:
                self.our_mcu_combo.setCurrentIndex(oidx)

        # Connect change of company to reload MCU list
        self.company_combo.currentIndexChanged.connect(self._reload_comp_mcUs)

    def _reload_comp_mcUs(self):
        cid = self.company_combo.currentData()
        self.comp_mcu_combo.clear()
        if cid is None:
            return
        for m in self.db.list_mcus_by_company(cid):
            self.comp_mcu_combo.addItem(m['name'], m['id'])

    def _save(self):
        entry_id = int(self.entry['id'])
        company_id = self.company_combo.currentData()
        comp_mcu_id = self.comp_mcu_combo.currentData()
        quantity = int(self.qty_spin.value())
        our_mcu_id = self.our_mcu_combo.currentData()
        if company_id is None or comp_mcu_id is None:
            return
        ok = self.db.update_nco_entry(entry_id, company_id=company_id, comp_mcu_id=comp_mcu_id,
                                      quantity=quantity, our_mcu_id=our_mcu_id)
        if ok:
            self.accept()

class EditMCUDialog(QDialog):
    def __init__(self, db, mcu_id: int, parent=None):
        super().__init__(parent)
        self.db = db
        self.mcu_id = mcu_id
        self.setWindowTitle('Edit MCU')
        self._build_ui()

    def _build_ui(self):
        v = QVBoxLayout(self)

        form = QFormLayout()
        mcu = dict(self.db.get_mcu_by_id(self.mcu_id))
        company_id = int(mcu.get('company_id'))

        self.company_label = QLabel(self.db.get_company_by_id(company_id)['name'])
        form.addRow('Company', self.company_label)

        self.name_edit = QLineEdit(mcu.get('name', ''))
        form.addRow('MCU Name', self.name_edit)

        self.inputs: Dict[str, QWidget] = {}
        for key, label, typ in FEATURE_FIELDS:
            val = mcu.get(key)
            if typ == 'text':
                w = QLineEdit(str(val or ''))
            elif typ == 'int':
                w = QLineEdit()
                w.setValidator(QIntValidator(0, 100000, self))
                try:
                    w.setText(str(int(val or 0)))
                except Exception:
                    w.setText('0')
            elif typ == 'bool':
                w = QCheckBox()
                w.setChecked(int(val or 0) == 1)
            elif typ == 'enum_fpu':
                w = QComboBox()
                w.addItems(['None', 'Single precision', 'Double precision'])
                try:
                    w.setCurrentIndex(int(val or 0))
                except Exception:
                    w.setCurrentIndex(0)
            else:
                w = QLineEdit(str(val or ''))
            self.inputs[key] = w
            form.addRow(label, w)

        v.addLayout(form)

        btn_row = QHBoxLayout()
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

    def _save(self):
        name = self.name_edit.text().strip()
        if not name:
            self.name_edit.setFocus()
            return
        payload = {'name': name}
        for key, label, typ in FEATURE_FIELDS:
            if typ == 'text':
                payload[key] = str(self.inputs[key].text()).strip()
            elif typ == 'int':
                try:
                    payload[key] = int(self.inputs[key].text() or '0')
                except Exception:
                    payload[key] = 0
            elif typ == 'bool':
                payload[key] = 1 if self.inputs[key].isChecked() else 0
            elif typ == 'enum_fpu':
                idx = self.inputs[key].currentIndex()
                payload[key] = idx
            else:
                payload[key] = str(self.inputs[key].text()).strip()
        if self.db.update_mcu(self.mcu_id, payload):
            self.accept()


class AddCompanyDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.created_company_id = None
        self.setWindowTitle('Add Company')
        self._build_ui()

    def _build_ui(self):
        v = QVBoxLayout(self)

        form = QFormLayout()
        self.name_edit = QLineEdit()
        form.addRow('Company Name', self.name_edit)
        v.addLayout(form)

        btn_row = QHBoxLayout()
        save_btn = QPushButton('Save')
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

    def _save(self):
        name = self.name_edit.text().strip()
        if not name:
            self.name_edit.setFocus()
            return
        # Our company is pre-added; newly added companies are competitors by default
        cid = self.db.ensure_company(name, 0)
        self.created_company_id = cid
        self.accept()


class DetailsDialog(QDialog):
    def __init__(self, db, comp_mcu_id: int, our_mcu_id: int, parent=None):
        super().__init__(parent)
        self.db = db
        self.comp_mcu_id = comp_mcu_id
        self.our_mcu_id = our_mcu_id
        self.setWindowTitle('MCU Comparison Details')
        # Default size: large, scaled to screen
        screen = QGuiApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
            w = int(geo.width() * 0.9)
            h = int(geo.height() * 0.9)
            self.resize(max(1200, w), max(850, h))
        else:
            self.resize(1400, 950)
        self.setMinimumSize(1200, 850)
        self._build_ui()

    # Public methods used in signal connections
    def print_details(self):
        # Use module-level safe builder to avoid attribute lookup issues
        html = _details_html_safe(self)
        if not html:
            return
        printer = QPrinter(QPrinter.HighResolution)
        try:
            printer.setOrientation(QPrinter.Landscape)
        except Exception:
            pass
        dlg = QPrintDialog(printer, self)
        dlg.setWindowTitle('Print Details')
        if dlg.exec() != QPrintDialog.Accepted:
            return
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

    def export_details_pdf(self):
        from PySide6.QtWidgets import QFileDialog
        html = _details_html_safe(self)
        if not html:
            return
        path, _ = QFileDialog.getSaveFileName(self, 'Export Details to PDF', 'Details.pdf', 'PDF Files (*.pdf)')
        if not path:
            return
        if not path.lower().endswith('.pdf'):
            path += '.pdf'
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
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

    def _build_ui(self):
        v = QVBoxLayout(self)

        comp = dict(self.db.get_mcu_by_id(self.comp_mcu_id))
        ours = dict(self.db.get_mcu_by_id(self.our_mcu_id))

        feat_cols = self.db.feature_columns()
        # Include flags needed by business rules (is_dsp, is_fpga)
        comp_feats = ({k: comp.get(k) for k in feat_cols}
                      | {'name': comp.get('name', ''), 'is_dsp': comp.get('is_dsp'), 'is_fpga': comp.get('is_fpga')})
        our_feats = {k: ours.get(k) for k in feat_cols}

        overall, per_feat = weighted_similarity(comp_feats, our_feats)
        cat = categorize(overall)
        def _cat_color(c: str) -> str:
            return {'Best Match': '#2e7d32', 'Partial': '#f9a825', 'No Match': '#c62828'}.get(c, '#2e7d32')

        self.header = QLabel(f"{comp['name']} vs {ours['name']} — Match {overall:.1f}% ({cat})")
        self.header.setObjectName('headerLabel')
        self.header.setAlignment(Qt.AlignCenter)
        v.addWidget(self.header)

        # Overall similarity progress
        self.overall_bar = QProgressBar()
        self.overall_bar.setRange(0, 100)
        self.overall_bar.setValue(int(round(overall)))
        self.overall_bar.setFormat(f"Overall Match: {overall:.1f}%")
        # Update chunk color by category
        self.overall_bar.setStyleSheet(
            f"QProgressBar {{ font-size: 22px; font-weight: 700; padding: 2px; }}"
            f" QProgressBar::chunk {{ background-color: {_cat_color(cat)}; }}"
        )
        self.overall_bar.setTextVisible(True)
        # Make percentage highly visible
        self.overall_bar.setStyleSheet(
            f"QProgressBar {{ font-size: 22px; font-weight: 700; padding: 2px; }}"
            f" QProgressBar::chunk {{ background-color: {_cat_color(cat)}; }}"
        )
        self.overall_bar.setMinimumHeight(34)
        v.addWidget(self.overall_bar)

        # Top toolbar with Print/Export so it's always visible
        top_tools = QHBoxLayout()
        top_tools.addStretch(1)
        btn_top_print = QPushButton('Print')
        btn_top_export = QPushButton('Export PDF')
        btn_top_print.clicked.connect(self.print_details)
        btn_top_export.clicked.connect(self.export_details_pdf)
        top_tools.addWidget(btn_top_print)
        top_tools.addWidget(btn_top_export)
        v.addLayout(top_tools)

        # Compact donut chart (Match vs Gap)
        series = QPieSeries()
        match_val = max(0.0, min(100.0, overall))
        gap_val = 100.0 - match_val
        series.append('Match', match_val)
        series.append('Gap', gap_val)
        series.setHoleSize(0.55)
        if series.slices():
            s0 = series.slices()[0]
            s0.setLabelVisible(False)
            s0.setBrush(Qt.green)
            s1 = series.slices()[1]
            s1.setLabelVisible(False)
            s1.setBrush(Qt.red)
        chart = QChart()
        chart.addSeries(series)
        chart.legend().setVisible(False)
        chart.setBackgroundVisible(False)
        chart.setTitle("")
        chart_view = QChartView(chart)
        chart_view.setRenderHint(QPainter.Antialiasing)
        self.overlay = QWidget()
        stack = QStackedLayout(self.overlay)
        stack.setContentsMargins(0, 0, 0, 0)
        stack.setStackingMode(QStackedLayout.StackingMode.StackAll)
        stack.addWidget(chart_view)
        self.pct_label = QLabel(f"Overall Match: {overall:.1f}%")
        self.pct_label.setAlignment(Qt.AlignCenter)
        text_color = self.palette().color(QPalette.WindowText).name()
        bg = 'rgba(0,0,0,0.35)' if self.palette().color(QPalette.Window).value() > 128 else 'rgba(255,255,255,0.18)'
        border = '#3b82f6'
        self.pct_label.setStyleSheet(
            f"font-size: 18px; font-weight: 600; color: {text_color}; "
            f"background: {bg}; border: 1px solid {border}; border-radius: 9px; padding: 4px 10px;"
        )
        self.pct_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        stack.addWidget(self.pct_label)
        from PySide6.QtWidgets import QSizePolicy
        self.overlay.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        chart_view.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.pct_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        stack.setAlignment(self.pct_label, Qt.AlignCenter)
        self.overlay.setMinimumHeight(160)
        self.overlay.setMaximumHeight(180)

        # Prepare summary bullets data (will render AFTER the table)
        def numeric(val):
            try:
                return float(val)
            except Exception:
                try:
                    return int(val)
                except Exception:
                    return None

        better: list[str] = []
        worse: list[str] = []
        more_is_better = {
            'max_clock_mhz', 'flash_kb', 'sram_kb', 'gpios', 'uarts', 'spis', 'i2cs', 'pwms',
            'timers', 'dacs', 'adcs', 'cans',
            'output_compare', 'qspi', 'ext_interrupts'
        }
        # Booleans displayed as Yes/No (concise) — do not show 'vs'
        bool_fields = {'eeprom', 'power_mgmt', 'clock_mgmt', 'qei', 'internal_osc', 'security_features'}
        # Directional booleans (display only): treat as booleans for bullet text
        bool_dir_display = {'input_capture', 'ethernet', 'emif', 'spi_slave'}
        for key, label, typ in FEATURE_FIELDS:
            c = comp_feats.get(key)
            o = our_feats.get(key)
            if key in more_is_better:
                cn = numeric(c)
                on = numeric(o)
                if cn is None or on is None:
                    continue
                if on > cn:
                    better.append(f"{label}: {on} vs {cn}")
                elif on < cn:
                    worse.append(f"{label}: {on} vs {cn}")
            elif key in bool_fields or key in bool_dir_display:
                try:
                    cb = 1 if int(c or 0) == 1 else 0
                except Exception:
                    cb = 0
                try:
                    ob = 1 if int(o or 0) == 1 else 0
                except Exception:
                    ob = 0
                if ob > cb:
                    better.append(f"{label}: Yes")
                elif ob < cb:
                    worse.append(f"{label}: No")
            elif key == 'fpu':
                # Higher level is better: 0=None,1=Single,2=Double
                try:
                    cf = int(c or 0)
                except Exception:
                    cf = 0
                try:
                    of = int(o or 0)
                except Exception:
                    of = 0
                names = {0: 'None', 1: 'Single precision', 2: 'Double precision'}
                if of > cf:
                    better.append(f"FPU: {names.get(of, of)} vs {names.get(cf, cf)}")
                elif of < cf:
                    worse.append(f"FPU: {names.get(of, of)} vs {names.get(cf, cf)}")
            # Skip textual cores and optional alt core in bullets

        # Two-column content row: [Main table] [Right sidebar with donut + panels]
        content_row = QHBoxLayout()
        content_row.setContentsMargins(0, 0, 0, 0)
        from PySide6.QtWidgets import QFrame
        # Main area: table only at top-left
        center_col = QVBoxLayout()
        center_col.setContentsMargins(0, 0, 0, 0)
        center_col.setSpacing(6)
        self.table = QTableWidget(len(FEATURE_FIELDS), 4)
        self.table.setHorizontalHeaderLabels(['Feature', 'Competitor', ours.get('name', 'OUR MCU'), 'Similarity'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        center_col.addWidget(self.table)
        # Context menu on table for print/export
        from PySide6.QtWidgets import QMenu
        def _open_table_menu(pos):
            menu = QMenu(self)
            act_print = menu.addAction('Print...')
            act_export = menu.addAction('Export to PDF...')
            action = menu.exec(self.table.mapToGlobal(pos))
            if action == act_print:
                self.print_details()
            elif action == act_export:
                self.export_details_pdf()
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(_open_table_menu)
        # Right sidebar container (donut + panels)
        right_host = QWidget()
        self._right_box = QVBoxLayout(right_host)
        self._right_box.setContentsMargins(0, 0, 0, 0)
        self._right_box.setSpacing(8)
        # Add donut overlay at the top of the sidebar
        self._right_box.addWidget(self.overlay)
        # Helper to build a panel
        def _mk_panel(html: str, ok: bool) -> QFrame:
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            frame.setStyleSheet(
                "QFrame { background: rgba(46, 125, 50, 0.10); border: 1px solid rgba(46,125,50,0.35); border-radius: 6px; padding: 6px; }"
                if ok else
                "QFrame { background: rgba(198, 40, 40, 0.10); border: 1px solid rgba(198,40,40,0.35); border-radius: 6px; padding: 6px; }"
            )
            lbl = QLabel(html)
            lbl.setWordWrap(True)
            lay = QVBoxLayout(frame)
            lay.setContentsMargins(6,6,6,6)
            lay.addWidget(lbl)
            return frame
        # Populate initial panels
        if better:
            better_html = """
            <div style='font-size:14px; line-height:1.35;'>
              <div style='font-weight:700; font-size:15px; color:#2e7d32; margin-bottom:6px;'>Areas OUR MCU is better</div>
              <ul style='list-style:none; margin:0; padding:0;'>
            """ + "".join(
                f"<li><span style='color:#2e7d32; font-weight:700; margin-right:8px;'>✔</span>{item}</li>"
                for item in better
            ) + """
              </ul>
            </div>
            """
            self._frame_better = _mk_panel(better_html, True)
            self._right_box.addWidget(self._frame_better)
        if worse:
            worse_html = """
            <div style='font-size:14px; line-height:1.35;'>
              <div style='font-weight:700; font-size:15px; color:#ffffff; margin-bottom:6px;'>Areas OUR MCU lacks</div>
              <ul style='list-style:none; margin:0; padding:0;'>
            """ + "".join(
                f"<li><span style='color:#ef5350; font-weight:700; margin-right:8px;'>✖</span>{item}</li>"
                for item in worse
            ) + """
              </ul>
            </div>
            """
            self._frame_worse = _mk_panel(worse_html, False)
            self._right_box.addWidget(self._frame_worse)
        # Keep sidebar content aligned to the top
        self._right_box.addStretch(1)
        # Assemble row: give the table most space
        content_row.addLayout(center_col, 3)
        content_row.addWidget(right_host, 1)
        v.addLayout(content_row)

        from mcu_compare.engine.similarity import DEFAULT_WEIGHTS
        # Keep references we need for live updates
        self._comp_feats = comp_feats
        self._our_feats = our_feats
        self._per_feat = per_feat
        self._sim_bars = {}
        for row, (key, label, typ) in enumerate(FEATURE_FIELDS):
            self.table.setItem(row, 0, QTableWidgetItem(label))
            # Boolean formatting as Yes/No
            def fmt(val):
                if typ == 'bool':
                    return 'Yes' if int(val or 0) == 1 else 'No'
                if key == 'fpu':
                    mapping = {0: 'None', 1: 'Single precision', 2: 'Double precision'}
                    try:
                        return mapping.get(int(val), str(val))
                    except Exception:
                        return str(val)
                return str(val)
            item_comp = QTableWidgetItem(fmt(comp.get(key, '')))
            item_comp.setFlags(item_comp.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, item_comp)
            item_our = QTableWidgetItem(fmt(ours.get(key, '')))
            # Make OUR MCU column read-only per requirements
            item_our.setFlags(item_our.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 2, item_our)
            sim = per_feat.get(key, 0.0) * 100.0
            sim_bar = QProgressBar()
            sim_bar.setRange(0, 100)
            sim_bar.setValue(int(round(sim)))
            sim_bar.setFormat(f"{sim:.0f}%")
            sim_bar.setTextVisible(True)
            self.table.setCellWidget(row, 3, sim_bar)
            self._sim_bars[key] = sim_bar
        # Table is read-only; no inline edits

        btn_row = QHBoxLayout()
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        v.addLayout(btn_row)

    def _on_table_item_changed(self, item: QTableWidgetItem):
        # Only handle edits in OUR MCU column
        if item.column() != 2:
            return
        row = item.row()
        if row < 0 or row >= len(FEATURE_FIELDS):
            return
        key, label, typ = FEATURE_FIELDS[row]
        # Only numeric fields are editable
        if typ not in ('int', 'float'):
            return
        text = item.text().strip()
        # Parse value
        try:
            if typ == 'int':
                val = int(float(text or '0'))
            else:
                val = float(text or '0')
        except Exception:
            # Revert display to previous value
            prev = self._our_feats.get(key, 0)
            item.setText(str(prev))
            return
        # Persist to DB and local cache
        if self.db.update_mcu(self.our_mcu_id, {key: val}):
            self._our_feats[key] = val
            self._recompute_and_refresh()

    def _recompute_and_refresh(self):
        # Recompute similarity and update UI pieces
        overall, per_feat = weighted_similarity(self._comp_feats, self._our_feats)
        self._per_feat = per_feat
        cat = categorize(overall)
        def _cat_color(c: str) -> str:
            return {'Best Match': '#2e7d32', 'Partial': '#f9a825', 'No Match': '#c62828'}.get(c, '#2e7d32')
        # Header and bars
        self.header.setText(f"{self._comp_feats.get('name','')} vs {self.db.get_mcu_by_id(self.our_mcu_id)['name']} — Match {overall:.1f}% ({cat})")
        self.overall_bar.setValue(int(round(overall)))
        self.overall_bar.setFormat(f"Overall Match: {overall:.1f}%")
        self.pct_label.setText(f"Overall Match: {overall:.1f}%")
        # Update per-row progress bars
        for (key, _, _typ) in FEATURE_FIELDS:
            bar = self._sim_bars.get(key)
            if bar is not None:
                sim = per_feat.get(key, 0.0) * 100.0
                bar.setValue(int(round(sim)))
                bar.setFormat(f"{sim:.0f}%")
        # Recompute bullets
        better: list[str] = []
        worse: list[str] = []
        def numeric(val):
            try:
                return float(val)
            except Exception:
                try:
                    return int(val)
                except Exception:
                    return None
        more_is_better = {
            'max_clock_mhz', 'flash_kb', 'sram_kb', 'gpios', 'uarts', 'spis', 'i2cs', 'pwms',
            'timers', 'dacs', 'adcs', 'cans',
            'output_compare', 'qspi', 'ext_interrupts'
        }
        bool_fields = {'eeprom', 'power_mgmt', 'clock_mgmt', 'qei', 'internal_osc', 'security_features'}
        bool_dir_display = {'input_capture', 'ethernet', 'emif', 'spi_slave'}
        for key, label, typ in FEATURE_FIELDS:
            c = self._comp_feats.get(key)
            o = self._our_feats.get(key)
            if key in more_is_better:
                cn = numeric(c)
                on = numeric(o)
                if cn is None or on is None:
                    continue
                if on > cn:
                    better.append(f"{label}: {on} vs {cn}")
                elif on < cn:
                    worse.append(f"{label}: {on} vs {cn}")
            elif key in bool_fields or key in bool_dir_display:
                try:
                    cb = 1 if int(c or 0) == 1 else 0
                except Exception:
                    cb = 0
                try:
                    ob = 1 if int(o or 0) == 1 else 0
                except Exception:
                    ob = 0
                if ob > cb:
                    better.append(f"{label}: Yes")
                elif ob < cb:
                    worse.append(f"{label}: No")
            elif key == 'fpu':
                try:
                    cf = int(c or 0)
                except Exception:
                    cf = 0
                try:
                    of = int(o or 0)
                except Exception:
                    of = 0
                names = {0: 'None', 1: 'Single precision', 2: 'Double precision'}
                if of > cf:
                    better.append(f"FPU: {names.get(of, of)} vs {names.get(cf, cf)}")
                elif of < cf:
                    worse.append(f"FPU: {names.get(of, of)} vs {names.get(cf, cf)}")
        # Rebuild sidebar frames
        from PySide6.QtWidgets import QFrame
        # Clear existing items in the sidebar layout
        if hasattr(self, '_side_layout') and self._side_layout is not None:
            while self._side_layout.count():
                w = self._side_layout.takeAt(0).widget()
                if w is not None:
                    w.setParent(None)
        # Add again if non-empty
        if better:
            better_html = """
            <div style='font-size:12px;'>
              <div style='font-weight:600; color:#2e7d32; margin-bottom:4px;'>Areas OUR MCU is better</div>
              <ul style='margin:0 0 0 16px; padding:0;'>
            """ + "".join(f"<li>{item}</li>" for item in better) + """
              </ul>
            </div>
            """
            frame_better = QFrame()
            frame_better.setFrameShape(QFrame.StyledPanel)
            frame_better.setStyleSheet("QFrame { background: rgba(46, 125, 50, 0.10); border: 1px solid rgba(46,125,50,0.35); border-radius: 6px; padding: 6px; }")
            lbl_better = QLabel(better_html)
            lbl_better.setWordWrap(True)
            fb = QVBoxLayout(frame_better)
            fb.setContentsMargins(6,6,6,6)
            fb.addWidget(lbl_better)
            self._side_layout.addWidget(frame_better)
        if worse:
            worse_html = """
            <div style='font-size:12px;'>
              <div style='font-weight:600; color:#c62828; margin-bottom:4px;'>Areas OUR MCU lacks</div>
              <ul style='margin:0 0 0 16px; padding:0;'>
            """ + "".join(f"<li>{item}</li>" for item in worse) + """
              </ul>
            </div>
            """
            frame_worse = QFrame()
            frame_worse.setFrameShape(QFrame.StyledPanel)
            frame_worse.setStyleSheet("QFrame { background: rgba(198, 40, 40, 0.10); border: 1px solid rgba(198,40,40,0.35); border-radius: 6px; padding: 6px; }")
            lbl_worse = QLabel(worse_html)
            lbl_worse.setWordWrap(True)
            fw = QVBoxLayout(frame_worse)
            fw.setContentsMargins(6,6,6,6)
            fw.addWidget(lbl_worse)
            self._side_layout.addWidget(frame_worse)
