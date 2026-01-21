import os
import json
from typing import List, Dict, Any, Optional
from glob import glob


class JsonDatabase:
    def __init__(self, path: str):
        # "path" is kept for backward compatibility (legacy: data/app.json)
        self.path = path
        self.data_dir = os.path.dirname(self.path)
        os.makedirs(self.data_dir, exist_ok=True)
        # New structure: companies.json and mcus_company_{id}.json files
        self._companies_file = os.path.join(self.data_dir, 'companies.json')

    # ----- File helpers -----
    def _load_companies(self) -> List[Dict[str, Any]]:
        if os.path.exists(self._companies_file):
            with open(self._companies_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []

    def _save_companies(self, companies: List[Dict[str, Any]]):
        with open(self._companies_file, 'w', encoding='utf-8') as f:
            json.dump(companies, f, indent=2)

    def _mcus_file(self, company_id: int) -> str:
        return os.path.join(self.data_dir, f'mcus_company_{company_id}.json')

    def _load_mcus(self, company_id: int) -> List[Dict[str, Any]]:
        fp = self._mcus_file(company_id)
        if os.path.exists(fp):
            with open(fp, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []

    def _save_mcus(self, company_id: int, mcus: List[Dict[str, Any]]):
        with open(self._mcus_file(company_id), 'w', encoding='utf-8') as f:
            json.dump(mcus, f, indent=2)

    # ----- Migration from legacy single-file app.json -----
    def _migrate_legacy_if_needed(self):
        # Legacy single file path (self.path), new companies file doesn't exist
        if os.path.exists(self._companies_file):
            return
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                legacy = json.load(f)
        except Exception:
            return
        companies = legacy.get('companies', [])
        mcus = legacy.get('mcus', [])
        # Save companies
        self._save_companies(companies)
        # Split mcus per company
        by_company: Dict[int, List[Dict[str, Any]]] = {}
        for m in mcus:
            by_company.setdefault(int(m.get('company_id', 0)), []).append(m)
        for cid, items in by_company.items():
            self._save_mcus(cid, items)

    def initialize(self):
        # First migrate if legacy file exists
        self._migrate_legacy_if_needed()
        # If companies still missing, seed
        companies = self._load_companies()
        if not companies:
            companies = [
                {"id": 1, "name": "Our Company", "is_ours": 1},
                {"id": 2, "name": "STMicroelectronics", "is_ours": 0},
                {"id": 3, "name": "NXP", "is_ours": 0},
                {"id": 4, "name": "Microchip", "is_ours": 0},
                {"id": 5, "name": "Renesas", "is_ours": 0},
                {"id": 6, "name": "Infineon", "is_ours": 0},
                {"id": 7, "name": "TI", "is_ours": 0},
                {"id": 8, "name": "Nordic", "is_ours": 0},
            ]
            self._save_companies(companies)
            # Seed MCUs into per-company files
            seeds = [
                (1, {
                    "id": 1, "company_id": 1, "name": "OCM4-120", "core": "ARM Cortex-M4",
                    "dsp_core": 1, "fpu": 1, "max_clock_mhz": 120, "flash_kb": 512, "sram_kb": 128, "eeprom_kb": 4,
                    "gpios": 80, "uarts": 4, "spis": 3, "i2cs": 2, "pwms": 8, "timers": 8, "dacs": 2, "adcs": 3,
                    "cans": 1, "power_mgmt": 1, "clock_mgmt": 1, "qei": 1, "internal_osc": 1, "security_features": 1
                }),
                (1, {
                    "id": 2, "company_id": 1, "name": "OCM0-48", "core": "ARM Cortex-M0+",
                    "dsp_core": 0, "fpu": 0, "max_clock_mhz": 48, "flash_kb": 128, "sram_kb": 32, "eeprom_kb": 2,
                    "gpios": 40, "uarts": 2, "spis": 2, "i2cs": 1, "pwms": 4, "timers": 6, "dacs": 1, "adcs": 1,
                    "cans": 0, "power_mgmt": 1, "clock_mgmt": 1, "qei": 0, "internal_osc": 1, "security_features": 0
                }),
                (1, {
                    "id": 3, "company_id": 1, "name": "OCM7-400", "core": "ARM Cortex-M7",
                    "dsp_core": 1, "fpu": 1, "max_clock_mhz": 400, "flash_kb": 2048, "sram_kb": 512, "eeprom_kb": 8,
                    "gpios": 160, "uarts": 6, "spis": 4, "i2cs": 4, "pwms": 16, "timers": 14, "dacs": 2, "adcs": 4,
                    "cans": 2, "power_mgmt": 1, "clock_mgmt": 1, "qei": 1, "internal_osc": 1, "security_features": 1
                }),
                (2, {
                    "id": 4, "company_id": 2, "name": "STM32F407", "core": "ARM Cortex-M4",
                    "dsp_core": 1, "fpu": 1, "max_clock_mhz": 168, "flash_kb": 1024, "sram_kb": 192, "eeprom_kb": 0,
                    "gpios": 140, "uarts": 6, "spis": 3, "i2cs": 3, "pwms": 12, "timers": 14, "dacs": 2, "adcs": 3,
                    "cans": 2, "power_mgmt": 1, "clock_mgmt": 1, "qei": 1, "internal_osc": 1, "security_features": 1
                }),
                (2, {
                    "id": 5, "company_id": 2, "name": "STM32G0B1", "core": "ARM Cortex-M0+",
                    "dsp_core": 0, "fpu": 0, "max_clock_mhz": 64, "flash_kb": 512, "sram_kb": 144, "eeprom_kb": 0,
                    "gpios": 84, "uarts": 6, "spis": 2, "i2cs": 2, "pwms": 8, "timers": 10, "dacs": 1, "adcs": 1,
                    "cans": 0, "power_mgmt": 1, "clock_mgmt": 1, "qei": 0, "internal_osc": 1, "security_features": 1
                }),
                (4, {
                    "id": 6, "company_id": 4, "name": "ATSAMD21G18", "core": "ARM Cortex-M0+",
                    "dsp_core": 0, "fpu": 0, "max_clock_mhz": 48, "flash_kb": 256, "sram_kb": 32, "eeprom_kb": 0,
                    "gpios": 52, "uarts": 6, "spis": 2, "i2cs": 2, "pwms": 6, "timers": 6, "dacs": 1, "adcs": 1,
                    "cans": 0, "power_mgmt": 1, "clock_mgmt": 1, "qei": 0, "internal_osc": 1, "security_features": 0
                }),
                (3, {
                    "id": 7, "company_id": 3, "name": "LPC55S69", "core": "ARM Cortex-M33",
                    "dsp_core": 1, "fpu": 1, "max_clock_mhz": 150, "flash_kb": 640, "sram_kb": 320, "eeprom_kb": 0,
                    "gpios": 160, "uarts": 10, "spis": 4, "i2cs": 4, "pwms": 16, "timers": 10, "dacs": 2, "adcs": 4,
                    "cans": 2, "power_mgmt": 1, "clock_mgmt": 1, "qei": 1, "internal_osc": 1, "security_features": 1
                }),
            ]
            # Write seeds to per-company files
            per: Dict[int, List[Dict[str, Any]]] = {}
            for cid, m in seeds:
                per.setdefault(cid, []).append(m)
            for cid, arr in per.items():
                self._save_mcus(cid, arr)

    # Utility
    def _next_company_id(self) -> int:
        companies = self._load_companies()
        return (max((c.get('id', 0) for c in companies), default=0) + 1)

    def _next_global_mcu_id(self) -> int:
        return (max((m.get('id', 0) for m in self.all_mcus()), default=0) + 1)

    # API compatible with previous DB layer
    def list_companies(self, search: str = '') -> List[Dict[str, Any]]:
        s = search.lower()
        comps = sorted(self._load_companies(), key=lambda c: (0 if c.get('is_ours') else 1, c['name']))
        if s:
            comps = [c for c in comps if s in c['name'].lower()]
        return comps

    def get_our_company_id(self) -> int:
        for c in self._load_companies():
            if c.get('is_ours') == 1:
                return c['id']
        raise RuntimeError('Our company not found in DB')

    def ensure_company(self, name: str, is_ours: int = 0) -> int:
        companies = self._load_companies()
        for c in companies:
            if c['name'] == name:
                return c['id']
        new_id = self._next_company_id()
        companies.append({"id": new_id, "name": name, "is_ours": int(is_ours)})
        self._save_companies(companies)
        # Ensure a per-company mcus file exists
        if not os.path.exists(self._mcus_file(new_id)):
            self._save_mcus(new_id, [])
        return new_id

    def list_mcus_by_company(self, company_id: int) -> List[Dict[str, Any]]:
        mcus = self._load_mcus(company_id)
        return sorted([m for m in mcus if m.get('company_id') == company_id], key=lambda m: m['name'])

    def get_mcu_by_id(self, mcu_id: int) -> Optional[Dict[str, Any]]:
        # Search across per-company files
        companies = self._load_companies()
        for c in companies:
            for m in self._load_mcus(c['id']):
                if m.get('id') == mcu_id:
                    return m
        return None

    def list_our_mcus(self) -> List[Dict[str, Any]]:
        our_id = self.get_our_company_id()
        return self.list_mcus_by_company(our_id)

    def insert_mcu(self, company_id: int, data: Dict[str, Any]) -> int:
        mcus = self._load_mcus(company_id)
        new_id = self._next_global_mcu_id()
        record: Dict[str, Any] = {
            "id": new_id,
            "company_id": company_id,
        }
        for f in ['name','core','dsp_core','fpu','max_clock_mhz','flash_kb','sram_kb','eeprom_kb','gpios','uarts','spis','i2cs','pwms','timers','dacs','adcs','cans','power_mgmt','clock_mgmt','qei','internal_osc','security_features']:
            record[f] = data.get(f, '' if f in ['name','core'] else 0)
        mcus.append(record)
        self._save_mcus(company_id, mcus)
        return new_id

    def feature_columns(self) -> List[str]:
        return [
            'core','dsp_core','fpu','max_clock_mhz','flash_kb','sram_kb','eeprom_kb','gpios','uarts','spis','i2cs','pwms','timers','dacs','adcs','cans','power_mgmt','clock_mgmt','qei','internal_osc','security_features'
        ]

    def all_mcus(self) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        for c in self._load_companies():
            result.extend(self._load_mcus(c['id']))
        return result
