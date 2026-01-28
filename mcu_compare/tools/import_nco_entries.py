#!/usr/bin/env python3
import os
import sys
import csv
import argparse
import re
from typing import Optional, Dict, Any, Tuple

# Optional Excel support
try:
    import openpyxl  # type: ignore
except Exception:
    openpyxl = None

# Allow running from repo root or this tools folder
HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(HERE, os.pardir, os.pardir))
DEFAULT_DATA_APP = os.path.join(REPO_ROOT, 'data', 'app.json')

# Import DB layer
sys.path.insert(0, os.path.abspath(os.path.join(HERE, os.pardir, os.pardir)))
from mcu_compare.data.json_db import JsonDatabase  # noqa: E402


def normalize_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def normalize_key(k: str) -> str:
    """Normalize header keys to compare flexibly (lowercase, strip non-alnum)."""
    return re.sub(r"[^a-z0-9]", "", (k or "").lower())


def load_rows_from_csv(path: str):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            yield row


def load_rows_from_excel(path: str):
    if openpyxl is None:
        raise RuntimeError('openpyxl is not installed; install it or provide a CSV file')
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    # Auto-detect header row by scanning early rows for expected keys
    expected_keys = {normalize_key(k) for k in [
        'org', 'organization', 'nco', 'nco/ commissions', 'nco/commissions',
        'company', 'manufacturer', 'part number', 'competitor_part', 'part',
        'quantity', 'quantity (1y)', 'qty'
    ]}
    rows_iter = ws.iter_rows(min_row=1, values_only=True)
    header = None
    # Scan first 50 rows to find a plausible header
    scanned = []
    for _ in range(50):
        try:
            r = next(rows_iter)
        except StopIteration:
            break
        scanned.append(r)
        cells = [str(x).strip() if x is not None else '' for x in r]
        norms = {normalize_key(c) for c in cells if c}
        if norms & expected_keys:
            header = cells
            break
    if header is None:
        # Fallback: use first scanned row
        if not scanned:
            return
        header = [str(h).strip() if h is not None else '' for h in scanned[0]]
        # remaining rows include from second scanned onward
        remaining_rows = iter(scanned[1:])
    else:
        # remaining rows are whatever left in rows_iter
        remaining_rows = rows_iter
    headers = [str(h).strip() if h is not None else '' for h in header]
    for r in remaining_rows:
        if r is None:
            continue
        row = {headers[i]: (r[i] if i < len(r) else None) for i in range(len(headers))}
        yield row


def require_int(val, field_name: str) -> int:
    try:
        return int(float(str(val).strip()))
    except Exception:
        raise ValueError(f"Invalid integer for {field_name}: {val}")


def resolve_company_id(db: JsonDatabase, company_name: str) -> Optional[int]:
    if not company_name:
        return None
    n = normalize_name(company_name)
    for c in db.list_companies(''):
        if normalize_name(c.get('name', '')) == n:
            return int(c['id'])
    # No implicit creation for companies
    return None


def resolve_competitor_by_name(db: JsonDatabase, part_name: str) -> Tuple[Optional[int], Optional[int]]:
    """Search all non-our companies for a unique part name match.
    Returns (company_id, mcu_id) or (None, None) if not found/ambiguous.
    """
    target = normalize_name(part_name)
    matches: list[Tuple[int, int]] = []
    for c in db.list_companies(''):
        if c.get('is_ours'):
            continue
        for m in db.list_mcus_by_company(c['id']):
            if normalize_name(m.get('name', '')) == target:
                matches.append((c['id'], int(m['id'])))
    if len(matches) == 1:
        return matches[0]
    return (None, None)
    


def resolve_comp_mcu_id(db: JsonDatabase, company_id: int, part_name: str) -> Optional[int]:
    if not part_name:
        return None
    n = normalize_name(part_name)
    for m in db.list_mcus_by_company(company_id):
        if normalize_name(m.get('name', '')) == n:
            return int(m['id'])
    return None


def resolve_our_mcu_id(db: JsonDatabase, part_name: Optional[str]) -> Optional[int]:
    if not part_name:
        return None
    n = normalize_name(part_name)
    our_id = db.get_our_company_id()
    for m in db.list_mcus_by_company(our_id):
        if normalize_name(m.get('name', '')) == n:
            return int(m['id'])
    return None


def ensure_org_id(db: JsonDatabase, org_name: Optional[str]) -> int:
    if not org_name:
        # default to first org
        orgs = db.list_nco_orgs()
        return orgs[0]['id'] if orgs else db.add_nco_org('Default Org')
    # add_nco_org returns existing id if name already present
    return db.add_nco_org(str(org_name).strip())


def import_rows(db: JsonDatabase, rows_iter, *, dry_run: bool = False, verbose: bool = False, preview: int = 0) -> Tuple[int, int, int]:
    """
    Returns: (inserted, skipped, errors)

    Expected columns (case-insensitive):
    - org or organization
    - company or manufacturer
    - competitor_part or competitor or part or comp_part
    - quantity or qty
    - our_part (optional)
    - notes (optional)
    """
    inserted = skipped = errors = 0
    for idx, raw in enumerate(rows_iter, start=1):
        # Flexible, punctuation-insensitive header access
        norm_map: Dict[str, Any] = { normalize_key(k): v for k, v in raw.items() }

        def get(*aliases: str):
            for a in aliases:
                v = norm_map.get(normalize_key(a))
                if v is not None and str(v).strip() != '':
                    return v
            return None

        # Common headers and variants seen in user sheet
        org_name = get('org', 'organization', 'nco_org', 'nco', 'nco/ commissions', 'nco/commissions')
        company_name = get('company', 'manufacturer')
        comp_part = get('competitor_part', 'competitor', 'part', 'comp_part', 'part number', 'part_number')
        qty = get('quantity', 'qty', 'quantity (1y)', 'quantity(1y)', '1y', 'qty (1y)')
        our_part = get('our_part', 'our', 'our_mcu')
        notes = get('notes') or ''
        # Trim text fields
        org_name = str(org_name).strip() if org_name is not None else None
        company_name = str(company_name).strip() if company_name is not None else None
        comp_part = str(comp_part).strip() if comp_part is not None else None
        our_part = str(our_part).strip() if our_part is not None else None

        try:
            # Company may be omitted (we can auto-resolve by part). Require part and quantity only.
            if not comp_part or qty in (None, ''):
                skipped += 1
                if verbose:
                    print(f"Row {idx} skipped: missing part or quantity (part='{comp_part}', qty='{qty}')")
                continue
            company_id: Optional[int]
            comp_mcu_id: Optional[int]
            if company_name:
                company_id = resolve_company_id(db, str(company_name))
                if company_id is None:
                    if verbose:
                        print(f"Row {idx} error: Unknown company '{company_name}'")
                    raise ValueError(f"Unknown company: {company_name}")
                comp_mcu_id = resolve_comp_mcu_id(db, company_id, str(comp_part))
                if comp_mcu_id is None:
                    if verbose:
                        print(f"Row {idx} error: Unknown part under {company_name}: '{comp_part}'")
                    raise ValueError(f"Unknown competitor part under {company_name}: {comp_part}")
            else:
                # Company omitted: find by unique part match across competitors
                company_id, comp_mcu_id = resolve_competitor_by_name(db, str(comp_part))
                if company_id is None or comp_mcu_id is None:
                    if verbose:
                        print(f"Row {idx} error: Competitor part not found or ambiguous: '{comp_part}'")
                    raise ValueError(f"Competitor part not found or ambiguous: {comp_part}")
            quantity = require_int(qty, 'quantity')
            our_mcu_id = resolve_our_mcu_id(db, str(our_part)) if our_part else None
            org_id = ensure_org_id(db, org_name)

            if not dry_run:
                db.add_nco_entry(company_id=company_id,
                                  comp_mcu_id=comp_mcu_id,
                                  quantity=quantity,
                                  our_mcu_id=our_mcu_id,
                                  notes=str(notes or ''),
                                  org_id=org_id)
            inserted += 1
            if preview and inserted <= preview:
                print(f"Row {idx} OK: org='{org_name}', company_id={company_id}, part='{comp_part}', qty={quantity}, our_mcu_id={our_mcu_id}")
        except Exception as ex:
            errors += 1
            print(f"Row {idx} error: {ex}")
    return inserted, skipped, errors


def main():
    p = argparse.ArgumentParser(description='Import NCO entries from CSV/Excel into the JSON DB')
    p.add_argument('input', help='Path to CSV or Excel file (.csv, .xlsx, .xlsm)')
    p.add_argument('--data', default=DEFAULT_DATA_APP, help='Path to data/app.json (used to locate the data directory)')
    p.add_argument('--dry-run', action='store_true', help='Parse and validate only; do not write')
    p.add_argument('--verbose', action='store_true', help='Print detailed reasons for skips and errors')
    p.add_argument('--preview', type=int, default=0, help='Print first N successful parsed rows')
    args = p.parse_args()

    inp = os.path.abspath(args.input)
    if not os.path.exists(inp):
        print(f"Input file not found: {inp}")
        return 2

    db = JsonDatabase(args.data)
    db.initialize()

    ext = os.path.splitext(inp)[1].lower()
    if ext == '.csv':
        rows = load_rows_from_csv(inp)
    elif ext in ('.xlsx', '.xlsm'):
        rows = load_rows_from_excel(inp)
    else:
        print('Unsupported file type. Provide .csv or .xlsx/.xlsm')
        return 2

    inserted, skipped, errors = import_rows(db, rows, dry_run=args.dry_run, verbose=args.verbose, preview=args.preview)
    print(f"Imported: {inserted}, Skipped: {skipped}, Errors: {errors}")
    return 0 if errors == 0 else 1


if __name__ == '__main__':
    raise SystemExit(main())
