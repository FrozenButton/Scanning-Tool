"""Deposit lookup tables, ore tiers, and multiplier code logic."""

import json
import logging
import re
from typing import Any, Dict, List, Optional

from scanning_tool.config import ROCK_TYPE_FILE
from scanning_tool.state import app_state

logger = logging.getLogger(__name__)


# --- Dynamic scan signature table ---
import pandas as pd
from pathlib import Path

# Path to the summary CSV (relative to project root)
SCAN_SIG_CSV = Path(__file__).parent.parent.parent / "csv" / "scansig" / "scan_signatures_summary.csv"

SCAN_SIGNATURES = {}
if SCAN_SIG_CSV.exists():
    try:
        df = pd.read_csv(SCAN_SIG_CSV)
        for _, row in df.iterrows():
            try:
                base_value = int(row["base_value"])
                max_multiplier = int(row["max_multiplier"])
                mineral = row["mineral"]
                category = row["category"]
                SCAN_SIGNATURES[base_value] = {
                    "name": mineral,
                    "category": category,
                    "base_value": base_value,
                    "max_multiplier": max_multiplier,
                }
            except Exception as e:
                logger.warning(f"Bad scan signature row: {row} ({e})")
    except Exception as e:
        logger.warning(f"Failed to load scan signature CSV: {e}")
else:
    logger.warning(f"Scan signature CSV not found: {SCAN_SIG_CSV}")

ORE_TIERS: Dict[str, Dict[str, Any]] = {
    "HIGHEST": {"ores": ["QUANTANIUM", "STILERON", "RICCITE"], "color": "#E88AFF"},
    "HIGH": {"ores": ["TARANITE", "BEXALITE", "GOLD"], "color": "#63E64C"},
    "MEDIUM": {"ores": ["LARANITE", "BORASE", "BERYL", "AGRICIUM", "HEPHAESTANITE"], "color": "#E6E14C"},
    "LOW": {"ores": ["TUNGSTEN", "TITANIUM", "SILICON", "IRON", "QUARTZ", "CORUNDUM", "COPPER", "TIN", "ALUMINUM", "ICE"], "color": "#E69E4C"},
}

ORE_VALUE_MAP: Dict[str, Dict[str, str]] = {}
for _tier, _data in ORE_TIERS.items():
    for _ore in _data["ores"]:
        ORE_VALUE_MAP[_ore.upper()] = {"tier": _tier, "color": _data["color"]}


def build_deposit_tables(rock_data: Dict) -> Dict:
    """Build per-region deposit tables from rock data."""
    deposit_tables: Dict[str, List[Dict[str, str]]] = {}
    for deposit_name, details in rock_data.items():
        ores = details.get("ores", {})
        table = []
        for ore_name, ore_data in ores.items():
            name_up = ore_name.upper()
            value_info = ORE_VALUE_MAP.get(name_up, {"tier": "OTHER", "color": "#888"})
            table.append({
                "name": ore_name.title(),
                "prob": f"{ore_data.get('prob', 0) * 100:.0f}%",
                "min": f"{ore_data.get('minPct', 0) * 100:.0f}%",
                "max": f"{ore_data.get('maxPct', 0) * 100:.0f}%",
                "med": f"{ore_data.get('medPct', 0) * 100:.0f}%",
                "tier": value_info["tier"],
                "color": value_info["color"],
            })
        tier_order = ["HIGHEST", "HIGH", "MEDIUM", "LOW", "OTHER"]
        table.sort(key=lambda x: tier_order.index(x["tier"]))
        deposit_tables[deposit_name.upper()] = table
    return deposit_tables


def load_rock_data() -> None:
    """Load RockType.json and build deposit tables into app_state."""
    with open(ROCK_TYPE_FILE, "r") as f:
        app_state.service_state.rock_data = json.load(f)

    app_state.service_state.deposit_tables = {
        region_name.upper(): build_deposit_tables(region_data)
        for region_name, region_data in app_state.service_state.rock_data.items()
    }



def lookup_deposit(code: Optional[str]) -> Optional[Dict[str, Any]]:
    """Look up a deposit by its numeric code using scraped scan signature data."""
    if not code:
        return None
    try:
        m = re.search(r"(\d+)$", code)
        if not m:
            return None
        num_code = int(m.group(1))
        # Find a matching base_value that divides num_code
        for base_value, info in SCAN_SIGNATURES.items():
            if num_code % base_value == 0:
                deposits = num_code // base_value
                return {
                    "name": info["name"],
                    "base_code": base_value,
                    "deposits": deposits,
                    "category": info["category"],
                    "max_multiplier": info["max_multiplier"],
                }
    except Exception:
        pass
    return None


def extract_code_from_text(raw_text: str):
    """Extract a deposit code from OCR text."""
    if not raw_text:
        return None, None
    matches = app_state.service_state.code_re.findall(raw_text)
    if not matches:
        return None, raw_text
    raw = matches[0].upper()
    if any(ch.isdigit() for ch in raw):
        m = re.match(r"([A-Za-z]?-?)([\d,\.]+)", raw)
        if m:
            prefix, digits = m.groups()
            digits = digits.replace(",", "").replace(".", "")
            candidate = prefix + digits
        else:
            candidate = raw.replace(",", "").replace(".", "")
    else:
        candidate = raw
    return candidate, raw
