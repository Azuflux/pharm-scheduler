"""Turn a cell's formatting into a portable style (explicit RGB colors), so formatting copied
from one workbook looks identical in another. Excel stores many colors as a palette index or a
theme slot plus tint, which only mean something inside the workbook that defined them."""
from __future__ import annotations

import colorsys
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import COLOR_INDEX

# Excel theme color index -> scheme element
_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4",
                "accent5", "accent6", "hlink", "folHlink"]
# Office default theme, used if a workbook has none
_DEFAULT_THEME = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
                  "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5", "accent4": "FFC000",
                  "accent5": "5B9BD5", "accent6": "70AD47", "hlink": "0563C1", "folHlink": "954F72"}


def theme_colors(wb) -> dict[str, str]:
    xml = getattr(wb, "loaded_theme", None)
    if not xml:
        return dict(_DEFAULT_THEME)
    if isinstance(xml, bytes):
        xml = xml.decode("utf-8", "ignore")
    out = dict(_DEFAULT_THEME)
    for name in _THEME_ORDER:
        m = re.search(rf"<a:{name}>(.*?)</a:{name}>", xml, re.S)
        if not m:
            continue
        inner = m.group(1)
        v = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', inner) or re.search(r'lastClr="([0-9A-Fa-f]{6})"', inner)
        if v:
            out[name] = v.group(1).upper()
    return out


def _apply_tint(hex6: str, tint: float) -> str:
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "".join(f"{round(x * 255):02X}" for x in (r, g, b))


def resolve(color, wb, theme: dict | None = None) -> str | None:
    """RGB hex (6 chars) for an openpyxl Color, or None if it can't be resolved / is automatic."""
    if color is None:
        return None
    t = color.type
    if t == "rgb":
        v = color.rgb
        return v[-6:].upper() if isinstance(v, str) and len(v) >= 6 else None
    if t == "indexed":
        idx = color.indexed
        palette = getattr(wb, "_colors", None) or COLOR_INDEX
        if idx is None or idx >= len(palette) or idx >= 64:
            return None          # 64/65 = system foreground/background: "automatic"
        return palette[idx][-6:].upper()
    if t == "theme":
        theme = theme or theme_colors(wb)
        if color.theme is None or color.theme >= len(_THEME_ORDER):
            return None
        base = theme[_THEME_ORDER[color.theme]]
        return _apply_tint(base, color.tint or 0)
    return None


def portable_style(cell, wb, theme: dict | None = None) -> dict:
    """Style dict (same shape as loader.cell_style) with explicit RGB colors."""
    theme = theme or theme_colors(wb)
    f, n = cell.fill, cell.font
    fill = PatternFill(fill_type=None)
    colored = False
    if f is not None and f.fill_type == "solid":
        rgb = resolve(f.fgColor, wb, theme)
        if rgb:
            fill = PatternFill("solid", fgColor="FF" + rgb)
            colored = True
    font_rgb = resolve(n.color, wb, theme) if n is not None else None
    if font_rgb == "000000":
        font_rgb = None
    font = Font(name=(n.name if n is not None and n.name else "Calibri"),
                sz=(n.sz if n is not None and n.sz else 11),
                b=bool(n.b) if n is not None else False, i=bool(n.i) if n is not None else False,
                color=("FF" + font_rgb) if font_rgb else None)
    colored = colored or bool(font_rgb)
    return {"fill": fill, "font": font, "colored": colored}
