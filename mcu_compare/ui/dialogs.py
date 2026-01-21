from typing import Dict, Any
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox, QSpinBox,
    QPushButton, QWidget, QFormLayout, QTableWidget, QTableWidgetItem, QHeaderView
)
from PySide6.QtCore import Qt

from mcu_compare.engine.similarity import weighted_similarity, categorize


FEATURE_FIELDS = [
    ('core', 'Core', 'text'),
    ('dsp_core', 'DSP core', 'int'),
    ('fpu', 'FPU', 'int'),
    ('max_clock_mhz', 'Max clock (MHz)', 'int'),
    ('flash_kb', 'Flash (KB)', 'int'),
    ('sram_kb', 'SRAM (KB)', 'int'),
    ('eeprom_kb', 'EEPROM (KB)', 'int'),
    ('gpios', 'GPIOs', 'int'),
    ('uarts', 'UART/USART', 'int'),
    ('spis', 'SPI', 'int'),
    ('i2cs', 'I2C', 'int'),
    ('pwms', 'PWM', 'int'),
    ('timers', 'Timers', 'int'),
    ('dacs', 'DAC', 'int'),
    ('adcs', 'ADC', 'int'),
    ('cans', 'CAN', 'int'),
    ('power_mgmt', 'Power management', 'int'),
    ('clock_mgmt', 'Clock management', 'int'),
    ('qei', 'Quadrature encoder', 'int'),
    ('internal_osc', 'Internal oscillator', 'int'),
    ('security_features', 'Security features', 'int'),
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
            else:
                w = QSpinBox()
                w.setRange(0, 100000)
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
                payload[key] = self.inputs[key].text().strip()
            else:
                payload[key] = int(self.inputs[key].value())
        self.db.insert_mcu(company_id, payload)
        self.accept()


class DetailsDialog(QDialog):
    def __init__(self, db, comp_mcu_id: int, our_mcu_id: int, parent=None):
        super().__init__(parent)
        self.db = db
        self.comp_mcu_id = comp_mcu_id
        self.our_mcu_id = our_mcu_id
        self.setWindowTitle('MCU Comparison Details')
        self.resize(800, 600)
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
        header.setAlignment(Qt.AlignCenter)
        v.addWidget(header)

        table = QTableWidget(len(FEATURE_FIELDS), 5)
        table.setHorizontalHeaderLabels(['Feature', 'Competitor', 'Our MCU', 'Similarity', 'Weight'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        from mcu_compare.engine.similarity import DEFAULT_WEIGHTS
        for row, (key, label, typ) in enumerate(FEATURE_FIELDS):
            table.setItem(row, 0, QTableWidgetItem(label))
            table.setItem(row, 1, QTableWidgetItem(str(comp.get(key, ''))))
            table.setItem(row, 2, QTableWidgetItem(str(ours.get(key, ''))))
            sim = per_feat.get(key, 0.0) * 100.0
            item_sim = QTableWidgetItem(f"{sim:.0f}%")
            item_sim.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 3, item_sim)
            w = DEFAULT_WEIGHTS.get(key, 0.0)
            table.setItem(row, 4, QTableWidgetItem(f"{w:.2f}"))

        v.addWidget(table)

        btn_row = QHBoxLayout()
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        v.addLayout(btn_row)
