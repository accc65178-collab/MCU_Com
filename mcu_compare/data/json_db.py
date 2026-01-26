import os
import json
from typing import List, Dict, Any, Optional
from glob import glob
import re


class JsonDatabase:
    def __init__(self, path: str):
        # "path" is kept for backward compatibility (legacy: data/app.json)
        self.path = path
        self.data_dir = os.path.dirname(self.path)
        os.makedirs(self.data_dir, exist_ok=True)
        # New structure: companies.json and mcus_company_{id}.json files
        self._companies_file = os.path.join(self.data_dir, 'companies.json')
        # NCO/Commission store
        self._nco_file = os.path.join(self.data_dir, 'nco.json')
        # NCO Organizations store
        self._nco_orgs_file = os.path.join(self.data_dir, 'nco_orgs.json')

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
        # Prefer name-based slug file. Fallback to ID-based legacy filename.
        comp = self.get_company_by_id(company_id)
        if comp:
            slug = self._slugify(comp.get('name', 'company'), company_id)
            return os.path.join(self.data_dir, f'mcus_{slug}.json')
        return os.path.join(self.data_dir, f'mcus_company_{company_id}.json')

    def _legacy_mcus_file(self, company_id: int) -> str:
        return os.path.join(self.data_dir, f'mcus_company_{company_id}.json')

    def _slugify(self, name: str, company_id: Optional[int] = None) -> str:
        s = name.strip().lower()
        s = re.sub(r'[^a-z0-9]+', '_', s)
        s = re.sub(r'_+', '_', s).strip('_')
        if s:
            return s
        return f'company_{company_id}' if company_id is not None else 'company'

    def _load_mcus(self, company_id: int) -> List[Dict[str, Any]]:
        # Try name-based file first
        fp = self._mcus_file(company_id)
        if os.path.exists(fp):
            with open(fp, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            return [self._normalize_mcu(r) for r in raw]
        # Fallback and migrate from legacy id-based file
        legacy = self._legacy_mcus_file(company_id)
        if os.path.exists(legacy):
            with open(legacy, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data = [self._normalize_mcu(r) for r in data]
            # Save to new path and remove legacy
            self._save_mcus(company_id, data)
            try:
                os.remove(legacy)
            except Exception:
                pass
            return data
        return []

    def _save_mcus(self, company_id: int, mcus: List[Dict[str, Any]]):
        with open(self._mcus_file(company_id), 'w', encoding='utf-8') as f:
            json.dump(mcus, f, indent=2)

    # ----- NCO/Commission helpers -----
    def _load_nco(self) -> List[Dict[str, Any]]:
        if os.path.exists(self._nco_file):
            with open(self._nco_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []

    def _save_nco(self, rows: List[Dict[str, Any]]):
        with open(self._nco_file, 'w', encoding='utf-8') as f:
            json.dump(rows, f, indent=2)

    def _load_nco_orgs(self) -> List[Dict[str, Any]]:
        if os.path.exists(self._nco_orgs_file):
            with open(self._nco_orgs_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []

    def _save_nco_orgs(self, rows: List[Dict[str, Any]]):
        with open(self._nco_orgs_file, 'w', encoding='utf-8') as f:
            json.dump(rows, f, indent=2)

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
            # Initialize NCO/NCO Orgs stores if missing
            if not os.path.exists(self._nco_file):
                self._save_nco([])
            if not os.path.exists(self._nco_orgs_file):
                self._save_nco_orgs([{"id": 1, "name": "Sample Org"}])
            # Migrate any existing NCO rows to include org_id if absent
            rows = self._load_nco()
            changed = False
            for r in rows:
                if 'org_id' not in r:
                    r['org_id'] = 1
                    changed = True
            if changed:
                self._save_nco(rows)

    # Utility
    def _next_company_id(self) -> int:
        companies = self._load_companies()
        return (max((c.get('id', 0) for c in companies), default=0) + 1)

    def _next_global_mcu_id(self) -> int:
        return (max((m.get('id', 0) for m in self.all_mcus()), default=0) + 1)

    def _next_nco_id(self) -> int:
        rows = self._load_nco()
        return (max((r.get('id', 0) for r in rows), default=0) + 1)

    def _next_nco_org_id(self) -> int:
        rows = self._load_nco_orgs()
        return (max((r.get('id', 0) for r in rows), default=0) + 1)

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
        # If marking as our company, clear the flag from others
        if int(is_ours) == 1:
            for c in companies:
                c['is_ours'] = 0
        companies.append({"id": new_id, "name": name, "is_ours": int(is_ours)})
        self._save_companies(companies)
        # Ensure a per-company mcus file exists
        if not os.path.exists(self._mcus_file(new_id)):
            self._save_mcus(new_id, [])
        return new_id

    def delete_mcu(self, mcu_id: int) -> bool:
        # Find the MCU and its company
        target = self.get_mcu_by_id(mcu_id)
        if not target:
            return False
        company_id = int(target.get('company_id'))
        mcus = self._load_mcus(company_id)
        new_mcus = [m for m in mcus if int(m.get('id')) != int(mcu_id)]
        if len(new_mcus) == len(mcus):
            return False
        self._save_mcus(company_id, new_mcus)
        # Clean up NCO entries that reference this competitor MCU
        rows = self._load_nco()
        changed = False
        kept = []
        for r in rows:
            if int(r.get('comp_mcu_id', -1)) == int(mcu_id):
                changed = True
                continue
            # If our_mcu_id matches, keep but null it
            if r.get('our_mcu_id') is not None and int(r.get('our_mcu_id') or -1) == int(mcu_id):
                r['our_mcu_id'] = None
                changed = True
            kept.append(r)
        if changed:
            self._save_nco(kept)
        return True

    def update_company_name(self, company_id: int, new_name: str) -> bool:
        companies = self._load_companies()
        # Prevent duplicate names
        for c in companies:
            if c['name'] == new_name and int(c.get('id')) != int(company_id):
                return False
        # Determine old/new file paths for MCUs to migrate if needed
        comp = self.get_company_by_id(company_id)
        if not comp:
            return False
        old_slug = self._slugify(comp.get('name', ''), company_id)
        new_slug = self._slugify(new_name, company_id)
        old_path = os.path.join(self.data_dir, f'mcus_{old_slug}.json')
        new_path = os.path.join(self.data_dir, f'mcus_{new_slug}.json')
        # Update name in companies list
        for c in companies:
            if int(c.get('id')) == int(company_id):
                c['name'] = new_name
        self._save_companies(companies)
        # If old MCUs file exists and new one doesn't, rename it
        try:
            if os.path.exists(old_path) and (old_path != new_path):
                if not os.path.exists(new_path):
                    os.rename(old_path, new_path)
        except Exception:
            # Ignore file move failures; data will still be accessible via legacy fallback if present
            pass
        return True

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

    def delete_company(self, company_id: int) -> bool:
        # Remove company from list
        companies = self._load_companies()
        before = len(companies)
        companies = [c for c in companies if int(c.get('id')) != int(company_id)]
        if len(companies) == before:
            return False
        self._save_companies(companies)
        # Delete its MCU file if present
        try:
            fp = self._mcus_file(company_id)
            if os.path.exists(fp):
                os.remove(fp)
        except Exception:
            pass
        # Remove related NCO entries
        rows = self._load_nco()
        rows = [r for r in rows if int(r.get('company_id', -1)) != int(company_id)]
        self._save_nco(rows)
        return True

    def insert_mcu(self, company_id: int, data: Dict[str, Any]) -> int:
        mcus = self._load_mcus(company_id)
        new_id = self._next_global_mcu_id()
        record: Dict[str, Any] = {
            "id": new_id,
            "company_id": company_id,
        }
        for f in ['name','core','core_mark','dsp_core','fpu','max_clock_mhz','flash_kb','sram_kb','eeprom','gpios','uarts','spis','i2cs','pwms','timers','dacs','adcs','cans','power_mgmt','clock_mgmt','qei','internal_osc','security_features']:
            record[f] = data.get(f, '' if f in ['name','core'] else 0)
        mcus.append(record)
        self._save_mcus(company_id, mcus)
        return new_id

    def update_mcu(self, mcu_id: int, data: Dict[str, Any]) -> bool:
        # Find the MCU and its company first
        target = self.get_mcu_by_id(mcu_id)
        if not target:
            return False
        company_id = int(target.get('company_id'))
        mcus = self._load_mcus(company_id)
        updated = False
        fields = ['name','core','core_mark','dsp_core','fpu','max_clock_mhz','flash_kb','sram_kb','eeprom','gpios','uarts','spis','i2cs','pwms','timers','dacs','adcs','cans','power_mgmt','clock_mgmt','qei','internal_osc','security_features']
        for idx, rec in enumerate(mcus):
            if rec.get('id') == mcu_id:
                # Update provided fields only
                for f in fields:
                    if f in data:
                        rec[f] = data[f]
                if 'name' in data:
                    rec['name'] = data['name']
                mcus[idx] = rec
                updated = True
                break
        if updated:
            self._save_mcus(company_id, mcus)
        return updated

    def feature_columns(self) -> List[str]:
        return [
            'core','core_mark','dsp_core','fpu','max_clock_mhz','flash_kb','sram_kb','eeprom','gpios','uarts','spis','i2cs','pwms','timers','dacs','adcs','cans','power_mgmt','clock_mgmt','qei','internal_osc','security_features'
        ]

    def all_mcus(self) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        for c in self._load_companies():
            result.extend(self._load_mcus(c['id']))
        return result

    # ----- NCO/Commission public API -----
    def add_nco_entry(self, company_id: int, comp_mcu_id: int, quantity: int, our_mcu_id: Optional[int] = None,
                      notes: str = '', org_id: Optional[int] = None) -> int:
        rows = self._load_nco()
        new_id = self._next_nco_id()
        # default org if not provided
        if org_id is None:
            orgs = self._load_nco_orgs()
            org_id = orgs[0]['id'] if orgs else 1
        rows.append({
            'id': new_id,
            'org_id': org_id,
            'company_id': company_id,
            'comp_mcu_id': comp_mcu_id,
            'quantity': int(quantity),
            'our_mcu_id': our_mcu_id,
            'notes': notes or ''
        })
        self._save_nco(rows)
        return new_id

    def list_nco_entries(self, org_id: Optional[int] = None) -> List[Dict[str, Any]]:
        rows = self._load_nco()
        if org_id is not None:
            rows = [r for r in rows if int(r.get('org_id', 0)) == int(org_id)]
        return rows

    def update_nco_entry(self, entry_id: int, *, company_id: Optional[int] = None,
                         comp_mcu_id: Optional[int] = None, quantity: Optional[int] = None,
                         our_mcu_id: Optional[int] = None, org_id: Optional[int] = None) -> bool:
        rows = self._load_nco()
        updated = False
        for r in rows:
            if int(r.get('id', -1)) == int(entry_id):
                if org_id is not None:
                    r['org_id'] = int(org_id)
                if company_id is not None:
                    r['company_id'] = int(company_id)
                if comp_mcu_id is not None:
                    r['comp_mcu_id'] = int(comp_mcu_id)
                if quantity is not None:
                    r['quantity'] = int(quantity)
                # our_mcu_id can be None
                if our_mcu_id is not None or 'our_mcu_id' in r:
                    r['our_mcu_id'] = our_mcu_id
                updated = True
                break
        if updated:
            self._save_nco(rows)
        return updated

    def delete_nco_entry(self, entry_id: int) -> bool:
        rows = self._load_nco()
        new_rows = [r for r in rows if int(r.get('id', -1)) != int(entry_id)]
        if len(new_rows) == len(rows):
            return False
        self._save_nco(new_rows)
        return True

    # ----- NCO Orgs public API -----
    def list_nco_orgs(self) -> List[Dict[str, Any]]:
        return self._load_nco_orgs()

    def add_nco_org(self, name: str) -> int:
        rows = self._load_nco_orgs()
        # prevent duplicates by name
        for r in rows:
            if r.get('name') == name:
                return r['id']
        new_id = self._next_nco_org_id()
        rows.append({'id': new_id, 'name': name})
        self._save_nco_orgs(rows)
        return new_id

    def update_nco_org(self, org_id: int, name: str) -> bool:
        rows = self._load_nco_orgs()
        # Prevent duplicate names
        for r in rows:
            if r.get('name') == name and int(r.get('id')) != int(org_id):
                return False
        updated = False
        for r in rows:
            if int(r.get('id')) == int(org_id):
                r['name'] = name
                updated = True
                break
        if updated:
            self._save_nco_orgs(rows)
        return updated

    def delete_nco_org(self, org_id: int) -> bool:
        rows = self._load_nco_orgs()
        new_rows = [r for r in rows if int(r.get('id')) != int(org_id)]
        if len(new_rows) == len(rows):
            return False
        self._save_nco_orgs(new_rows)
        # Also delete entries belonging to this org
        nco = self._load_nco()
        nco = [r for r in nco if int(r.get('org_id', -1)) != int(org_id)]
        self._save_nco(nco)
        return True

    # Helpers
    def get_company_by_id(self, company_id: int) -> Optional[Dict[str, Any]]:
        for c in self._load_companies():
            if c.get('id') == company_id:
                return c
        return None

    def _normalize_mcu(self, rec: Dict[str, Any]) -> Dict[str, Any]:
        r = dict(rec)
        # Ensure company_id is an int
        try:
            r['company_id'] = int(r.get('company_id', 0))
        except Exception:
            r['company_id'] = 0
        if 'eeprom' not in r:
            kb = r.get('eeprom_kb', 0)
            try:
                r['eeprom'] = 1 if int(kb) > 0 else 0
            except Exception:
                r['eeprom'] = 0
        if 'core_mark' not in r:
            try:
                r['core_mark'] = int(r.get('core_mark', 0))
            except Exception:
                r['core_mark'] = 0
        return r
