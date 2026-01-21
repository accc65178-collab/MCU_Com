from typing import Dict, Any
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox, QSpinBox,
    QPushButton, QWidget, QFormLayout, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QProgressBar, QStackedLayout
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QGuiApplication
from PySide6.QtCharts import QChart, QChartView, QPieSeries

from mcu_compare.engine.similarity import weighted_similarity, categorize


FEATURE_FIELDS = [
    ('core', 'Core', 'text'),
    ('core_alt', 'Core (Alt)', 'text'),
    ('dsp_core', 'DSP core', 'bool'),
    ('fpu', 'FPU', 'enum_fpu'),
    ('max_clock_mhz', 'Max clock (MHz)', 'int'),
    ('flash_kb', 'Flash (KB)', 'int'),
    ('sram_kb', 'SRAM (KB)', 'int'),
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
]


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
                w = QSpinBox()
                w.setRange(0, 100000)
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
                payload[key] = int(self.inputs[key].value())
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
        v = QVBoxLayout(self)
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

        v.addWidget(table)

        btn_row = QHBoxLayout()
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        v.addLayout(btn_row)


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
                w = QSpinBox()
                w.setRange(0, 100000)
                try:
                    w.setValue(int(val or 0))
                except Exception:
                    w.setValue(0)
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
                payload[key] = int(self.inputs[key].value())
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

    def _build_ui(self):
        v = QVBoxLayout(self)

        comp = dict(self.db.get_mcu_by_id(self.comp_mcu_id))
        ours = dict(self.db.get_mcu_by_id(self.our_mcu_id))

        feat_cols = self.db.feature_columns()
        comp_feats = {k: comp.get(k) for k in feat_cols}
        our_feats = {k: ours.get(k) for k in feat_cols}

        overall, per_feat = weighted_similarity(comp_feats, our_feats)
        cat = categorize(overall)

        header = QLabel(f"{comp['name']} vs {ours['name']} — Match {overall:.1f}% ({cat})")
        header.setObjectName('headerLabel')
        header.setAlignment(Qt.AlignCenter)
        v.addWidget(header)

        # Overall similarity progress
        overall_bar = QProgressBar()
        overall_bar.setRange(0, 100)
        overall_bar.setValue(int(round(overall)))
        overall_bar.setFormat(f"Overall Match: {overall:.1f}%")
        overall_bar.setTextVisible(True)
        v.addWidget(overall_bar)

        # Donut chart (Match vs Gap)
        series = QPieSeries()
        match_val = max(0.0, min(100.0, overall))
        gap_val = 100.0 - match_val
        series.append('Match', match_val)
        series.append('Gap', gap_val)
        # Donut hole
        series.setHoleSize(0.55)
        # Optional slice colors align with theme
        if series.slices():
            s0 = series.slices()[0]
            s0.setLabelVisible(True)
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
        # Overlay percentage text centered in donut
        overlay = QWidget()
        stack = QStackedLayout(overlay)
        stack.setContentsMargins(0, 0, 0, 0)
        # Show all stacked widgets so label overlays the chart
        stack.setStackingMode(QStackedLayout.StackingMode.StackAll)
        stack.addWidget(chart_view)
        pct_label = QLabel(f"{overall:.1f}%")
        pct_label.setAlignment(Qt.AlignCenter)
        pct_label.setStyleSheet("font-size: 28px; font-weight: 600;")
        pct_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        stack.addWidget(pct_label)
        # ensure the overlay area is tall enough for chart visibility
        overlay.setMinimumHeight(420)
        v.addWidget(overlay)

        # Summary bullets: where OUR MCU is better and where it lacks (placed above table for visibility)
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
            'timers', 'dacs', 'adcs', 'cans'
        }
        bool_fields = {'eeprom', 'power_mgmt', 'clock_mgmt', 'qei', 'internal_osc', 'security_features'}
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
            elif key in bool_fields:
                try:
                    cb = 1 if int(c or 0) == 1 else 0
                except Exception:
                    cb = 0
                try:
                    ob = 1 if int(o or 0) == 1 else 0
                except Exception:
                    ob = 0
                if ob > cb:
                    better.append(f"{label}: Yes vs No")
                elif ob < cb:
                    worse.append(f"{label}: No vs Yes")
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

        from PySide6.QtWidgets import QFrame
        if better:
            better_html = """
            <div style='font-size:14px;'>
              <div style='font-weight:600; color:#2e7d32; margin-bottom:6px;'>Areas OUR MCU is better</div>
              <ul style='margin:0 0 0 18px; padding:0;'>
            """ + "".join(f"<li>{item}</li>" for item in better) + """
              </ul>
            </div>
            """
            frame_better = QFrame()
            frame_better.setFrameShape(QFrame.StyledPanel)
            frame_better.setStyleSheet("QFrame { background: rgba(46, 125, 50, 0.10); border: 1px solid rgba(46,125,50,0.35); border-radius: 6px; padding: 8px; }")
            lbl_better = QLabel(better_html)
            lbl_better.setWordWrap(True)
            fb = QVBoxLayout(frame_better)
            fb.setContentsMargins(8,8,8,8)
            fb.addWidget(lbl_better)
            v.addWidget(frame_better)
        if worse:
            worse_html = """
            <div style='font-size:14px;'>
              <div style='font-weight:600; color:#c62828; margin-bottom:6px;'>Areas OUR MCU lacks</div>
              <ul style='margin:0 0 0 18px; padding:0;'>
            """ + "".join(f"<li>{item}</li>" for item in worse) + """
              </ul>
            </div>
            """
            frame_worse = QFrame()
            frame_worse.setFrameShape(QFrame.StyledPanel)
            frame_worse.setStyleSheet("QFrame { background: rgba(198, 40, 40, 0.10); border: 1px solid rgba(198,40,40,0.35); border-radius: 6px; padding: 8px; }")
            lbl_worse = QLabel(worse_html)
            lbl_worse.setWordWrap(True)
            fw = QVBoxLayout(frame_worse)
            fw.setContentsMargins(8,8,8,8)
            fw.addWidget(lbl_worse)
            v.addWidget(frame_worse)

        # Feature comparison table
        table = QTableWidget(len(FEATURE_FIELDS), 4)
        table.setHorizontalHeaderLabels(['Feature', 'Competitor', 'OUR MCU', 'Similarity'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        from mcu_compare.engine.similarity import DEFAULT_WEIGHTS
        for row, (key, label, typ) in enumerate(FEATURE_FIELDS):
            table.setItem(row, 0, QTableWidgetItem(label))
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
            table.setItem(row, 1, QTableWidgetItem(fmt(comp.get(key, ''))))
            table.setItem(row, 2, QTableWidgetItem(fmt(ours.get(key, ''))))
            sim = per_feat.get(key, 0.0) * 100.0
            sim_bar = QProgressBar()
            sim_bar.setRange(0, 100)
            sim_bar.setValue(int(round(sim)))
            sim_bar.setFormat(f"{sim:.0f}%")
            sim_bar.setTextVisible(True)
            table.setCellWidget(row, 3, sim_bar)

        v.addWidget(table)

        btn_row = QHBoxLayout()
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        v.addLayout(btn_row)
