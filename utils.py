import re
from typing import Tuple, Optional

def xml_escape(text_to_escape: str) -> str:
    return (
        text_to_escape.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;")
    )

def calculate_distance(p1: Tuple[float, float], p2: Tuple[float, float]) -> float:
    # Haversine formula
    from math import radians, sin, cos, sqrt, atan2
    R = 6371e3
    lat1, lon1, lat2, lon2 = map(radians, [p1[0], p1[1], p2[0], p2[1]])
    dlon, dlat = lon2 - lon1, lat2 - lat1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return R * 2 * atan2(sqrt(a), sqrt(1 - a))

def convert_color(input_value: str, output_type: str = "hex", allow_name_lookup: bool = False) -> Optional[str]:
    from colors import COLORS
    if input_value is None:
        return None

    hex_to_name = {v.lower(): k for k, v in COLORS.items()}
    input_hex = ""

    if isinstance(input_value, str):
        value_lower = input_value.lower().strip()
        if value_lower in {k.lower() for k in COLORS}:
            input_hex = COLORS[[k for k in COLORS if k.lower() == value_lower][0]].lower()
        elif value_lower.startswith("#") and len(value_lower) == 7:
            input_hex = value_lower
        else:
            input_hex = "#ffffff"
    elif isinstance(input_value, (list, tuple)) and len(input_value) == 3:
        try:
            input_hex = "#{:02x}{:02x}{:02x}".format(*input_value)
        except (TypeError, ValueError):
            input_hex = "#ffffff"

    if output_type.lower() == 'hex':
        return input_hex
    elif output_type.lower() == 'name':
        return hex_to_name.get(input_hex, "White")
    elif output_type.lower() == 'int_rgb':
        h = input_hex.lstrip('#')
        try:
            return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
        except (TypeError, ValueError):
            return (255, 255, 255)
    elif output_type.lower() == 'str_rgb':
        h = input_hex.lstrip('#')
        try:
            rgb = tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
            return f"{rgb[0]},{rgb[1]},{rgb[2]}"
        except (TypeError, ValueError):
            return "255,255,255"
    return None
