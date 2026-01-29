from typing import Dict, Any, List, Tuple
import math

CORE_FAMILIES = {
    'ARM Cortex-M0': 'ARM_Cortex_M',
    'ARM Cortex-M0+': 'ARM_Cortex_M',
    'ARM Cortex-M3': 'ARM_Cortex_M',
    'ARM Cortex-M4': 'ARM_Cortex_M',
    'ARM Cortex-M7': 'ARM_Cortex_M',
    'ARM Cortex-M23': 'ARM_Cortex_M',
    'ARM Cortex-M33': 'ARM_Cortex_M',
    'ARM Cortex-M55': 'ARM_Cortex_M',
    'RISC-V': 'RISC_V',
    'RV32IMC': 'RISC_V',
    'RV32I': 'RISC_V',
    'RV64I': 'RISC_V',
    'FREE-RISC': 'RISC_V',
    'AVR': 'AVR',
    '8051': 'C51',
    'PIC': 'PIC',
}

DEFAULT_WEIGHTS: Dict[str, float] = {
    'core': 0.0,
    'core_mark': 0,
    'dsp_core': 0.02,
    'fpu': 0.15,
    'max_clock_mhz': 0.10,
    'flash_kb': 0.05,
    'sram_kb': 0.05,
    'eeprom': 0.01,
    'gpios': 0.05,
    'uarts': 0.05,
    'spis': 0.05,
    'i2cs': 0.05,
    'pwms': 0.05,
    'timers': 0.05,
    'dacs': 0.02,
    'adcs': 0.03,
    'cans': 0.03,
    'power_mgmt': 0.03,
    'clock_mgmt': 0.03,
    'qei': 0.04,
    'internal_osc': 0.04,
    'security_features': 0.05,
    # New interfaces/peripherals (initial weights; will be normalized overall)
    'output_compare': 0.02,
    'input_capture': 0.02,
    'qspi': 0.03,
    'ethernet': 0.04,
    'emif': 0.03,
    'spi_slave': 0.02,
    'ext_interrupts': 0.03,
}

CATEGORIES = [
    (80.0, 'Best Match'),
    (65.0, 'Partial'),
    (0.0, 'No Match')
]


def categorize(score_pct: float) -> str:
    for threshold, name in CATEGORIES:
        if score_pct >= threshold:
            return name
    return 'No match'


def core_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    fa = CORE_FAMILIES.get(a, a)
    fb = CORE_FAMILIES.get(b, b)
    if fa == fb:
        return 0.8
    return 0.0


def ratio_similarity(x: float, y: float) -> float:
    if x == 0 and y == 0:
        return 1.0
    if x == 0 or y == 0:
        return 0.0
    return min(x, y) / max(x, y)


def coverage_similarity(requirement: float, offer: float) -> float:
    """Directional similarity: how well 'offer' (OUR MCU) covers 'requirement' (competitor).
    - If offer >= requirement, full score 1.0 (we meet or exceed).
    - If offer < requirement, proportional score offer/requirement.
    Handles zeros: requirement==0 and offer>0 => 1.0; both 0 => 1.0.
    """
    try:
        req = float(requirement or 0)
        off = float(offer or 0)
    except Exception:
        return 0.0
    if req <= 0:
        # Competitor requires none; any offer is OK and considered fully covering
        return 1.0
    if off <= 0:
        return 0.0
    return min(1.0, off / req)


def feature_similarity(feature: str, comp_val: Any, our_val: Any) -> float:
    # Core comparison remains categorical/family-based
    if feature == 'core':
        return core_similarity(str(comp_val), str(our_val))
    # Ordinal/numeric fields use directional coverage
    if feature == 'core_mark':
        return coverage_similarity(comp_val, our_val)
    if feature == 'fpu':
        # Higher FPU level is better: 0=None, 1=Single, 2=Double
        try:
            cf = float(comp_val or 0)
            of = float(our_val or 0)
        except Exception:
            return 0.0
        return coverage_similarity(cf, of)
    # Boolean features: directional coverage
    # - If competitor requires it (True) and we lack it => 0
    # - If competitor lacks it (False), we get full credit whether we have it or not; especially if we do, it's a plus
    BOOL_DIR = {
        'input_capture', 'ethernet', 'emif', 'spi_slave'
    }
    if feature in BOOL_DIR:
        def _to01(x: Any) -> float:
            try:
                if isinstance(x, str):
                    xl = x.strip().lower()
                    if xl in ('yes', 'true', '1', 'y', 'on'): return 1.0
                    if xl in ('no', 'false', '0', 'off', ''): return 0.0
                return 1.0 if float(x) > 0 else 0.0
            except Exception:
                return 1.0 if bool(x) else 0.0
        return coverage_similarity(_to01(comp_val), _to01(our_val))
    # Counts and other numerics treated as coverage
    try:
        xa = float(comp_val or 0)
        xb = float(our_val or 0)
    except Exception:
        return 0.0
    return coverage_similarity(xa, xb)


def weighted_similarity(a: Dict[str, Any], b: Dict[str, Any], weights: Dict[str, float] = None) -> Tuple[float, Dict[str, float]]:
    # Start from defaults if no explicit weights passed
    if weights is None:
        # Dynamic adjustment for max_clock_mhz based on competitor requirement (a)
        base = dict(DEFAULT_WEIGHTS)
        try:
            req_clock = float(a.get('max_clock_mhz') or 0)
        except Exception:
            req_clock = 0.0
        target_clock_w = base['max_clock_mhz']
        if req_clock > 300:
            target_clock_w = 0.20
        elif req_clock > 200:
            target_clock_w = 0.15
        # If change is needed, scale other weights so the total stays constant
        if abs(target_clock_w - base['max_clock_mhz']) > 1e-9:
            total_default = sum(base.values())
            remaining_default = total_default - base['max_clock_mhz']
            remaining_target = total_default - target_clock_w
            scale = remaining_target / remaining_default if remaining_default > 0 else 1.0
            for k, v in list(base.items()):
                if k == 'max_clock_mhz':
                    base[k] = target_clock_w
                else:
                    base[k] = v * scale
        # Normalize final weights to sum to 1.0 so percentages are intuitive
        total_after = sum(base.values())
        if total_after > 0:
            base = {k: (v / total_after) for k, v in base.items()}
        weights = base
    # Business rule: if competitor is an FPGA, treat as no match outright
    core_str = str(a.get('core', '') or '')
    is_fpga_flag = False
    try:
        is_fpga_flag = bool(int(a.get('is_fpga', 0)))
    except Exception:
        is_fpga_flag = bool(a.get('is_fpga'))
    if is_fpga_flag or ('fpga' in core_str.lower()):
        return 0.0, {feat: 0.0 for feat in weights.keys()}

    total_w = sum(weights.values())
    score = 0.0
    per_feature: Dict[str, float] = {}
    # Helper to support alias: 'DSP' can be used instead of 'dsp_core' in data
    def _get(d: Dict[str, Any], feat: str):
        if feat == 'dsp_core':
            return d.get('dsp_core', d.get('DSP'))
        return d.get(feat)
    for feat, w in weights.items():
        s = feature_similarity(feat, _get(a, feat), _get(b, feat))
        per_feature[feat] = s
        score += w * s
    if total_w <= 0:
        return 0.0, per_feature
    pct = (score / total_w) * 100.0
    # DSP rule: apply 20% deduction ONLY if competitor JSON has is_dsp set
    is_dsp_flag = False
    try:
        val = a.get('is_dsp', 0)
        # Treat numeric/string '1', 'true' as True
        if isinstance(val, str):
            is_dsp_flag = val.strip().lower() in ('1', 'true', 'yes', 'y')
        else:
            is_dsp_flag = bool(int(val)) if isinstance(val, (int, float)) else bool(val)
    except Exception:
        is_dsp_flag = False
    if is_dsp_flag:
        pct = max(0.0, pct - 20.0)
    return pct, per_feature


def best_match(target: Dict[str, Any], candidates: List[Dict[str, Any]], weights: Dict[str, float] = None) -> Tuple[Dict[str, Any], float, Dict[str, float]]:
    best = None
    best_score = -1.0
    best_pf = {}
    for c in candidates:
        s, pf = weighted_similarity(target, c, weights)
        if s > best_score:
            best = c
            best_score = s
            best_pf = pf
    return best, best_score, best_pf
