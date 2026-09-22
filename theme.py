"""Theme configuration for PLN / IPS Dashboard and Reports.

Contains color palettes, brand constants, and visual configuration.
All source code must be ASCII-only.
"""

from __future__ import annotations
from typing import Dict, List, Tuple

PRIMARY = "#0A3D8F"       # Navy blue
PRIMARY_DARK = "#062A63"
SECONDARY = "#1E88E5"
ACCENT = "#FDB913"        # Yellow
SUCCESS = "#2E7D32"
WARNING = "#F57C00"
DANGER = "#C62828"
NEUTRAL = "#90A4AE"
BG = "#F4F7FB"

PALETTE: List[str] = [
    PRIMARY,
    ACCENT,
    SECONDARY,
    "#26A69A",
    WARNING,
    DANGER,
    NEUTRAL,
    "#7E57C2"
]

STATUS_COLORS: Dict[str, str] = {
    "SELESAI": SUCCESS,
    "PROSES": ACCENT,
    "BELUM ADA STATUS": NEUTRAL,
    "CANCEL": DANGER,
}

PAYMENT_COLORS: Dict[str, str] = {
    "PAYMENT": SUCCESS,
    "PROSES": ACCENT,
    "PROSES PAYMENT": SECONDARY,
    "MENUNGGU": WARNING,
    "CANCEL": DANGER,
    "BELUM ADA STATUS": NEUTRAL,
}

LOGO_PATH = "data/logo.png"
APP_TITLE = "Dashboard Monitoring Pekerjaan"
ORG_NAME = "PT PLN Indonesia Power Services (IPS)"


def hex_to_rgb(hex_str: str) -> Tuple[int, int, int]:
    """Convert hex color string to RGB tuple."""
    h = hex_str.lstrip("#")
    if len(h) == 3:
        h = "".join([c * 2 for c in h])
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
