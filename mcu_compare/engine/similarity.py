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
    'core': 0.10,
    'core_alt': 0.05,
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
}

CATEGORIES = [
    (90.0, 'Direct'),
    (75.0, 'Near'),
    (60.0, 'Partial'),
    (0.0, 'No match')
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


def feature_similarity(feature: str, a: Any, b: Any) -> float:
    if feature in ('core', 'core_alt'):
        return core_similarity(str(a), str(b))
    try:
        xa = float(a)
        xb = float(b)
    except Exception:
        return 0.0
    return ratio_similarity(xa, xb)


def weighted_similarity(a: Dict[str, Any], b: Dict[str, Any], weights: Dict[str, float] = None) -> Tuple[float, Dict[str, float]]:
    weights = weights or DEFAULT_WEIGHTS
    total_w = sum(weights.values())
    score = 0.0
    per_feature: Dict[str, float] = {}
    for feat, w in weights.items():
        s = feature_similarity(feat, a.get(feat), b.get(feat))
        per_feature[feat] = s
        score += w * s
    if total_w <= 0:
        return 0.0, per_feature
    return (score / total_w) * 100.0, per_feature


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
