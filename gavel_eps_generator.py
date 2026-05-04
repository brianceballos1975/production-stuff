"""
gavel_eps_generator.py — Generate SVG production layouts for custom gavel band orders.

Pulls pending (not-yet-shipped) gavel orders from ShipStation, downloads Amazon
customization data, and writes a single print-ready SVG layout (24" × 12" pages,
3 columns × 10 rows, 0.25" spacing) with every design on one file.

Usage:
    python gavel_eps_generator.py [--days N] [--output DIR]

Options:
    --days N      Look back N days for pending orders (default: 1)
    --output DIR  Output folder for layout file (default: ./gavel_eps)
"""

import argparse
import base64
import csv
import io
import json
import os
import sys
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path


def _load_env(path: str = ".env") -> None:
    """Load key=value pairs from a .env file into os.environ (no dependencies)."""
    env_file = Path(path)
    if not env_file.exists():
        # Try relative to this script's directory
        env_file = Path(__file__).parent / path
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip())

_load_env()

# Fix Windows cp1252 stdout choking on Unicode customer names/text
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── config ────────────────────────────────────────────────────────────────────

SS_KEY        = os.environ.get("SHIPSTATION_API_KEY", "")
SS_SECRET     = os.environ.get("SHIPSTATION_API_SECRET", "")
SS_AUTH       = "Basic " + base64.b64encode(f"{SS_KEY}:{SS_SECRET}".encode()).decode()
BASE_URL      = "https://ssapi.shipstation.com"
PAGE_SIZE     = 100
TEMPLATES_DIR    = Path(__file__).parent / "templates"
TEMPLATE_PATH    = str(TEMPLATES_DIR / "gavelband_template.svg")
TEMPLATE_PATH_7  = str(TEMPLATES_DIR / "gavelband_template_7.svg")   # 7" walnut/black template
SB_TEMPLATE_PATH = str(TEMPLATES_DIR / "soundblock_template.svg")

# Item name keywords (case-insensitive) that trigger the 7" walnut/black template
BAND_TEMPLATE_7_KEYWORDS = ["walnut", "black"]


def _select_band_template(item_name: str) -> str:
    """Return the appropriate band template path based on the product title."""
    name_lower = (item_name or "").lower()
    if any(kw in name_lower for kw in BAND_TEMPLATE_7_KEYWORDS):
        return TEMPLATE_PATH_7
    return TEMPLATE_PATH

# ── Sound block dimensions (1000 SVG units/inch, same as band template) ────────
SB_W           = 3750.0          # 3.75" × 3.75"
SB_H           = 3750.0
SB_CX          = SB_W / 2        # 1875 — horizontal center
SB_CY          = SB_H / 2        # 1875 — vertical center
SB_TEXT_W      = 2500.0          # 2.50" text area
SB_TEXT_H      = 2500.0
SB_MAX_FONT_SIZE = 500.0         # cap ~36 pt — prevents absurdly large single-line text
SB_COLS        = 6
SB_ROWS        = 3

# ── Trello ────────────────────────────────────────────────────────────────────
TRELLO_API_KEY     = os.environ.get("TRELLO_API_KEY", "")
TRELLO_TOKEN       = os.environ.get("TRELLO_TOKEN", "")
TRELLO_BOARD_NAME  = "customs"
TRELLO_LIST_NAME   = "Customs Ready For Production"

# ── coordinate system: CorelDRAW template at 1000 SVG user units per inch ─────
# Template: 6.875in × 0.95in, viewBox "0 0 6875 950" → 6875/6.875 = 1000 u/in

UNITS_PER_INCH = 1000.0   # SVG user units per inch (CorelDRAW template)

# Band dimensions from template SVG viewBox ("0 0 6875 950")
BAND_W  = 6875.0
BAND_H  = 950.0
BAND_CX = BAND_W / 2   # 3437.5  horizontal center
BAND_CY = BAND_H / 2   # 475.0   vertical center

# Layout grid (all in SVG user units)
PAGE_W = 24 * UNITS_PER_INCH    # 24000
PAGE_H = 12 * UNITS_PER_INCH    # 12000
COLS   = 3
ROWS   = 10
GAP        = 0.25 * UNITS_PER_INCH   # 250  (sound block layout)
BAND_COL_GAP = 0.180 * UNITS_PER_INCH  # 180 — horizontal gap between band columns
BAND_ROW_GAP = 0.100 * UNITS_PER_INCH  # 100 — vertical gap between band rows

# Font sizes in SVG user units (pt × UNITS_PER_INCH/72)
# 1 pt = 1000/72 ≈ 13.889 units
FONT_SIZE_PT = {1: 208.33, 2: 180.56, 3: 152.78, 4: 138.89, 5: 125.00, 6: 111.11}
LEADING = 1.25

# CMYK(0, 0.993347, 1, 0) → RGB #FF0200
BAND_STROKE_COLOR = "#FF0200"
BAND_STROKE_WIDTH = 3.47  # SVG user units (fallback; actual from template attr)

# SKU substrings that identify gavel band products
GAVEL_SKU_PATTERNS = ["GVL", "GF-VLU"]

# Item name keywords (case-insensitive) used as a fallback when the SKU
# doesn't match any pattern — catches products with random Amazon-generated SKUs
GAVEL_NAME_KEYWORDS = ["gavel band", "engraved band", "personalized band", "gavel with"]

# ── Font-to-curves support ────────────────────────────────────────────────────

FONTS_DIR = Path(__file__).parent / "fonts"

# Substitute these font names before lookup (key=lowercase incoming name, value=replacement)
FONT_ALIASES: dict[str, str] = {
    "helvetica":           "Swiss BT",
    "helvetica neue":      "Swiss BT",
    "helvetica-bold":      "Swiss BT Bold",
    "swiss 721 bt":        "Swiss BT",
    "swis721 bt":          "Swiss BT",
}

# Maps lowercase font family name → filename inside FONTS_DIR
FONT_FILE_MAP: dict[str, str] = {
    # Arial family
    "arial":                         "arial.ttf",
    "arial bold":                     "arialbd.ttf",
    "arial italic":                   "ariali.ttf",
    "arial bold italic":              "arialbi.ttf",
    "arial narrow":                   "ARIALN.TTF",
    "arial narrow bold":              "ARIALNB.TTF",
    "arial narrow bold italic":       "ARIALNBI.TTF",
    "arial narrow italic":            "ARIALNI.TTF",
    "arial black":                    "ariblk.ttf",
    # Century
    "century":                        "CENTURY.TTF",
    "century751 bt":                  "Century751 BT Roman.ttf",
    "century751 bt roman":            "Century751 BT Roman.ttf",
    "century751 bt italic":           "Century751 BT Italic.ttf",
    # Clarendon BT
    "clarendon bt":                   "Clarendon BT Roman.ttf",
    "clarendon bt roman":             "Clarendon BT Roman.ttf",
    "clarendon bt bold":              "Clarendon BT Bold.ttf",
    # Copperplate Gothic
    "copperplate gothic":             "COPRGTL.TTF",
    "copperplate gothic light":       "COPRGTL.TTF",
    "copperplate gothic bold":        "COPRGTB.TTF",
    # Georgia
    "georgia":                        "georgia.ttf",
    "georgia bold":                   "georgiab.ttf",
    "georgia italic":                 "georgiai.ttf",
    "georgia bold italic":            "georgiaz.ttf",
    # Goudy Old Style
    "goudy old style":                "GOUDOS.TTF",
    "goudy old style bold":           "GOUDOSB.TTF",
    "goudy old style italic":         "GOUDOSI.TTF",
    # Homemade Apple
    "homemade apple":                 "HomemadeApple-Regular.ttf",
    # Lato
    "lato":                           "Lato-Regular.ttf",
    "lato regular":                   "Lato-Regular.ttf",
    "lato light":                     "Lato-Light.ttf",
    "lato bold":                      "Lato-Bold.ttf",
    "lato italic":                    "Lato-Italic.ttf",
    "lato black":                     "Lato-Black.ttf",
    "lato thin":                      "Lato-Thin.ttf",
    # Playfair Display
    "playfair display":               "PlayfairDisplay-Regular.ttf",
    "playfair display regular":       "PlayfairDisplay-Regular.ttf",
    "playfair display bold":          "PlayfairDisplay-Bold.ttf",
    "playfair display italic":        "PlayfairDisplay-Italic.ttf",
    "playfair display semibold":      "PlayfairDisplay-SemiBold.ttf",
    "playfair display medium":        "PlayfairDisplay-Medium.ttf",
    "playfair display black":         "PlayfairDisplay-Black.ttf",
    # Roboto Slab
    "roboto slab":                    "RobotoSlab-Regular.ttf",
    "roboto slab regular":            "RobotoSlab-Regular.ttf",
    "roboto slab bold":               "RobotoSlab-Bold.ttf",
    "roboto slab light":              "RobotoSlab-Light.ttf",
    "roboto slab medium":             "RobotoSlab-Medium.ttf",
    "roboto slab semibold":           "RobotoSlab-SemiBold.ttf",
    "roboto slab thin":               "RobotoSlab-Thin.ttf",
    # Swiss BT (Helvetica substitute)
    "swiss bt":                       "Swis721 BT Roman.ttf",
    "swiss bt roman":                 "Swis721 BT Roman.ttf",
    "swiss bt bold":                  "Swis721 BT Bold.ttf",
    "swiss bt italic":                "Swis721 BT Italic.ttf",
    "swiss bt bold italic":           "Swis721 BT Bold Italic.ttf",
    # Times New Roman
    "times new roman":                "times.ttf",
    "times":                          "times.ttf",
    "times new roman bold":           "timesbd.ttf",
    "times new roman italic":         "timesi.ttf",
    "times new roman bold italic":    "timesbi.ttf",
    # Verdana
    "verdana":                        "verdana.ttf",
    "verdana bold":                   "verdanab.ttf",
    "verdana italic":                 "verdanai.ttf",
    "verdana bold italic":            "verdanaz.ttf",
    # Comic Sans
    "comic sans":                     "comic.ttf",
    "comic sans ms":                  "comic.ttf",
    "comic sans ms bold":             "comicbd.ttf",
    "comic sans ms italic":           "comici.ttf",
    "comic sans ms bold italic":      "comicz.ttf",
    # Wingdings
    "wingdings":                      "wingding.ttf",
    "wingdings 2":                    "WINGDNG2.TTF",
    "wingdings 3":                    "WINGDNG3.TTF",
    # Arimo
    "arimo":                          "Arimo-Regular.ttf",
    "arimo regular":                  "Arimo-Regular.ttf",
    "arimo bold":                     "Arimo-Bold.ttf",
    "arimo italic":                   "Arimo-Italic.ttf",
    "arimo bold italic":              "Arimo-BoldItalic.ttf",
    "arimo medium":                   "Arimo-Medium.ttf",
    "arimo medium italic":            "Arimo-MediumItalic.ttf",
    "arimo semibold":                 "Arimo-SemiBold.ttf",
    "arimo semibold italic":          "Arimo-SemiBoldItalic.ttf",
    # Rock Salt
    "rock salt":                      "RockSalt-Regular.ttf",
    "rock salt regular":              "RockSalt-Regular.ttf",
    # Cormorant Garamond
    "cormorant garamond":             "CormorantGaramond-Regular.ttf",
    "cormorant garamond regular":     "CormorantGaramond-Regular.ttf",
    "cormorant garamond bold":        "CormorantGaramond-Bold.ttf",
    "cormorant garamond italic":      "CormorantGaramond-Italic.ttf",
    "cormorant garamond bold italic": "CormorantGaramond-BoldItalic.ttf",
    "cormorant garamond light":       "CormorantGaramond-Light.ttf",
    "cormorant garamond light italic":"CormorantGaramond-LightItalic.ttf",
    "cormorant garamond medium":      "CormorantGaramond-Medium.ttf",
    "cormorant garamond medium italic":"CormorantGaramond-MediumItalic.ttf",
    "cormorant garamond semibold":    "CormorantGaramond-SemiBold.ttf",
    "cormorant garamond semibold italic":"CormorantGaramond-SemiBoldItalic.ttf",
    # Alegreya
    "alegreya":                       "Alegreya-Regular.ttf",
    "alegreya regular":               "Alegreya-Regular.ttf",
    "alegreya bold":                  "Alegreya-Bold.ttf",
    "alegreya italic":                "Alegreya-Italic.ttf",
    "alegreya bold italic":           "Alegreya-BoldItalic.ttf",
    "alegreya medium":                "Alegreya-Medium.ttf",
    "alegreya medium italic":         "Alegreya-MediumItalic.ttf",
    "alegreya semibold":              "Alegreya-SemiBold.ttf",
    "alegreya semibold italic":       "Alegreya-SemiBoldItalic.ttf",
    "alegreya extrabold":             "Alegreya-ExtraBold.ttf",
    "alegreya extrabold italic":      "Alegreya-ExtraBoldItalic.ttf",
    "alegreya black":                 "Alegreya-Black.ttf",
    "alegreya black italic":          "Alegreya-BlackItalic.ttf",
    # Syncopate
    "syncopate":                      "Syncopate-Regular.ttf",
    "syncopate regular":              "Syncopate-Regular.ttf",
    "syncopate bold":                 "Syncopate-Bold.ttf",
    # Oswald
    "oswald":                         "Oswald-Regular.ttf",
    "oswald regular":                 "Oswald-Regular.ttf",
    "oswald bold":                    "Oswald-Bold.ttf",
    "oswald light":                   "Oswald-Light.ttf",
    "oswald extralight":              "Oswald-ExtraLight.ttf",
    "oswald extra light":             "Oswald-ExtraLight.ttf",
    "oswald medium":                  "Oswald-Medium.ttf",
    "oswald semibold":                "Oswald-SemiBold.ttf",
    "oswald semi bold":               "Oswald-SemiBold.ttf",
}


def resolve_font_path(family: str) -> tuple[Path, str] | tuple[None, None]:
    """
    Resolve a font family name to a TTF file in FONTS_DIR.
    Applies FONT_ALIASES first (e.g. Helvetica → Swiss BT).
    Returns (path, effective_name) or (None, None) if not found.
    """
    original = family.strip()
    key = original.lower()

    # Apply alias substitution
    if key in FONT_ALIASES:
        family = FONT_ALIASES[key]
        key = family.lower()

    # Exact match in map
    if key in FONT_FILE_MAP:
        p = FONTS_DIR / FONT_FILE_MAP[key]
        if p.exists():
            return p, family

    # Partial match: try progressively shorter prefix of the family name
    parts = key.split()
    for n in range(len(parts) - 1, 0, -1):
        partial = " ".join(parts[:n])
        if partial in FONT_FILE_MAP:
            p = FONTS_DIR / FONT_FILE_MAP[partial]
            if p.exists():
                return p, family

    # Fuzzy scan: find any file in FONTS_DIR whose stem contains the key
    norm_key = key.replace(" ", "").replace("-", "")
    for f in sorted(FONTS_DIR.iterdir()):
        stem_norm = f.stem.lower().replace(" ", "").replace("-", "").replace("_", "")
        if norm_key in stem_norm:
            return f, family

    return None, None


# TTFont cache — fonts are loaded once per process run
_font_cache: dict[str, object] = {}


def _get_ttfont(font_path: str):
    if font_path not in _font_cache:
        from fontTools.ttLib import TTFont as _TTFont
        _font_cache[font_path] = _TTFont(font_path)
    return _font_cache[font_path]


def text_line_to_svg_paths(
    text: str,
    font_path: str,
    font_size: float,
    cx: float,
    baseline_y: float,
) -> list[str]:
    """
    Render one line of text as a list of SVG path d-strings (curves), centered at cx.
    Uses fontTools TransformPen to apply scale + Y-flip + translate in one pass.
    Returns [] for blank text.
    """
    from fontTools.pens.svgPathPen import SVGPathPen as _SVGPathPen
    from fontTools.pens.transformPen import TransformPen as _TransformPen

    text = text.strip()
    if not text:
        return []

    font   = _get_ttfont(font_path)
    gs     = font.getGlyphSet()
    cmap   = font.getBestCmap() or {}
    upem   = font["head"].unitsPerEm
    hmtx   = font["hmtx"].metrics
    scale  = font_size / upem

    # Measure total advance width for horizontal centering
    total_adv = 0
    glyph_seq = []
    for ch in text:
        gname = cmap.get(ord(ch))
        glyph_seq.append(gname)
        adv = hmtx[gname][0] if gname and gname in hmtx else (upem // 3)
        total_adv += adv

    x_cursor = cx - (total_adv * scale) / 2

    # Render each glyph
    # Transform: x' = x_cursor + x*scale,  y' = baseline_y - y*scale
    # Affine matrix (a,b,c,d,e,f): x'=ax+cy+e, y'=bx+dy+f
    #   a=scale, b=0, c=0, d=-scale, e=x_cursor, f=baseline_y
    ntos = lambda n: f"{n:.3f}"  # 3-decimal precision
    paths = []
    for ch, gname in zip(text, glyph_seq):
        adv = hmtx[gname][0] if gname and gname in hmtx else (upem // 3)
        if gname and gname in gs:
            pen   = _SVGPathPen(gs, ntos=ntos)
            t_pen = _TransformPen(pen, (scale, 0, 0, -scale, x_cursor, baseline_y))
            gs[gname].draw(t_pen)
            d = pen.getCommands()
            if d:
                paths.append(d)
        x_cursor += adv * scale

    return paths


# ── Google Fonts families (lowercase)
_GOOGLE_FONTS = {
    "lato", "open sans", "roboto", "montserrat", "oswald", "raleway",
    "source sans pro", "ubuntu", "nunito", "poppins", "merriweather",
    "playfair display", "old standard tt", "homemade apple", "dancing script",
    "great vibes", "pacifico", "lobster", "abril fatface", "noto serif",
    "noto sans", "pt sans", "pt serif", "libre baskerville",
    "cormorant garamond",
}


def is_google_font(family: str) -> bool:
    return family.strip().lower() in _GOOGLE_FONTS


_UNICODE_SUBS = str.maketrans({
    '\u2018': "'", '\u2019': "'",
    '\u201c': '"', '\u201d': '"',
    '\u2013': '-', '\u2014': '--',
    '\u2026': '...', '\u00ae': '(R)', '\u00a9': '(C)',
    '\u00b0': 'deg', '\u2122': '(TM)',
    '\u00e9': 'e', '\u00e8': 'e',
    '\u00e0': 'a', '\u00e2': 'a',
    '\u00f1': 'n', '\u00fc': 'u', '\u00e4': 'a',
})

def sanitize_text(text: str) -> str:
    return text.translate(_UNICODE_SUBS)


def xml_escape_attr(text: str) -> str:
    """Escape text for use inside an XML attribute value (quoted with \")."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )

def xml_escape(text: str) -> str:
    """Escape text for use as XML element content (between tags).
    CorelDRAW renders XML named entities (&amp; &quot; etc.) as literal text
    rather than interpreting them, so we only escape < and > (which would break
    XML structure) and leave & and " as raw characters."""
    return (
        text.replace("<", "&lt;")
            .replace(">", "&gt;")
    )


# ── Band path extraction ──────────────────────────────────────────────────────

import xml.etree.ElementTree as ET
import re as _re

_band_path_cache: dict[str, tuple[str, float, float, float]] = {}
_sb_elements_cache: list[str] | None = None


def _load_soundblock_elements() -> list[str]:
    """
    Read shape elements from the sound block SVG template and return them
    as ready-to-embed SVG strings with fill="white" (CorelDRAW won't lock
    fill="white" paths, unlike fill="none").
    Cached after first call.
    """
    global _sb_elements_cache
    if _sb_elements_cache is not None:
        return _sb_elements_cache

    tree = ET.parse(SB_TEMPLATE_PATH)
    root = tree.getroot()
    SVG_NS = "http://www.w3.org/2000/svg"

    elements = []
    for el in root.iter():
        tag = el.tag.replace(f"{{{SVG_NS}}}", "")
        if tag == "rect":
            x  = el.get("x", "0")
            y  = el.get("y", "0")
            w  = el.get("width", "0")
            h  = el.get("height", "0")
            sw = el.get("stroke-width", str(BAND_STROKE_WIDTH))
            elements.append(
                f'  <rect x="{x}" y="{y}" width="{w}" height="{h}"'
                f' fill="white" stroke="{BAND_STROKE_COLOR}"'
                f' stroke-width="{sw}"/>'
            )
        elif tag == "path":
            d  = el.get("d", "")
            sw = el.get("stroke-width", str(BAND_STROKE_WIDTH))
            if d:
                for sp in _split_subpaths(d):
                    elements.append(
                        f'  <path d="{sp}"'
                        f' fill="white" stroke="{BAND_STROKE_COLOR}"'
                        f' stroke-width="{sw}"/>'
                    )

    _sb_elements_cache = elements
    return elements


def _measure_text_advance(text: str, font_path: str, font_size: float) -> float:
    """Return the total advance width of text in SVG units at the given font size."""
    font  = _get_ttfont(font_path)
    cmap  = font.getBestCmap() or {}
    hmtx  = font["hmtx"].metrics
    upem  = font["head"].unitsPerEm
    scale = font_size / upem
    total = 0.0
    for ch in text:
        gname = cmap.get(ord(ch))
        adv   = hmtx[gname][0] if gname and gname in hmtx else (upem // 3)
        total += adv
    return total * scale

def _load_band_path(template_path: str | None = None) -> tuple[str, float, float, float]:
    """
    Read the band outline path from an SVG template.
    Returns (d_string, stroke_width, band_w, band_h).
    Cached per template path so each file is only parsed once per run.
    """
    if template_path is None:
        template_path = TEMPLATE_PATH
    if template_path in _band_path_cache:
        return _band_path_cache[template_path]

    tree = ET.parse(template_path)
    root = tree.getroot()
    ns   = {"svg": "http://www.w3.org/2000/svg"}

    # Dimensions from viewBox
    vb     = root.get("viewBox", f"0 0 {BAND_W} {BAND_H}").split()
    band_w = float(vb[2])
    band_h = float(vb[3])

    path_el = root.find(".//svg:path", ns)
    if path_el is None:
        path_el = root.find(".//path")
    if path_el is None:
        raise ValueError(f"No <path> element found in {template_path}")

    d = path_el.get("d", "")

    stroke_w = None

    # 1. XML attribute (CorelDRAW)
    xml_sw = path_el.get("stroke-width")
    if xml_sw:
        try:
            stroke_w = float(xml_sw)
        except ValueError:
            pass

    # 2. CSS style attribute / <style> block (Illustrator)
    if stroke_w is None:
        style_el = root.find(".//svg:style", ns)
        if style_el is None:
            style_el = root.find(".//style")
        sources = [path_el.get("style", "")]
        if style_el is not None and style_el.text:
            sources.append(style_el.text)
        for src in sources:
            m = _re.search(r'stroke-width\s*:\s*(\d*\.?\d+)', src)
            if m:
                stroke_w = float(m.group(1))
                break

    if stroke_w is None:
        stroke_w = BAND_STROKE_WIDTH

    _band_path_cache[template_path] = (d, stroke_w, band_w, band_h)
    return _band_path_cache[template_path]


def _split_subpaths(d: str) -> list[str]:
    """
    Split a compound SVG path d-string into individual subpath strings,
    one per M command.  CorelDRAW locks compound paths (multiple M commands
    in a single <path>); splitting into separate <path> elements imports each
    as a normal, unlocked Curve object.

    Supports both:
      - Illustrator format: ...Z M x y...  (uppercase Z, absolute M)
      - CorelDRAW format:   ...z m dx dy...  (lowercase z, RELATIVE m)

    For relative m after z: new subpath start = previous-subpath-start + (dx,dy)
    because after z the current point is the subpath's opening M coordinate.
    This converts each lowercase m to an absolute M so every returned subpath
    begins with an unambiguous absolute M — required for correct offsetting.

    Also strips zero-length lineto artifacts before the closing z/Z:
      - CorelDRAW: l0 0z / L0 0Z  (relative or absolute, comma or space sep)
      - Illustrator: h0Z / H0Z
    """
    _NUM = r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?'
    _NUM_RE = _re.compile(_NUM)

    d = d.strip()

    # Split at every Z/z immediately followed (possibly with whitespace) by M/m.
    raw_parts = _re.split(r'(?<=[Zz])\s*(?=[Mm])', d)

    result = []
    sp_x, sp_y = 0.0, 0.0   # absolute starting point of the CURRENT subpath

    for i, part in enumerate(raw_parts):
        part = part.strip()
        if not part:
            continue

        cmd  = part[0]
        rest = part[1:].strip()

        if cmd in 'Mm':
            # Pull the first two coordinate numbers from the M/m operand
            coord_m = _re.match(
                r'(' + _NUM + r')[\s,]+(' + _NUM + r')[\s,]*',
                rest
            )
            if coord_m:
                dx, dy   = float(coord_m.group(1)), float(coord_m.group(2))
                after_xy = rest[coord_m.end():]
                if cmd == 'm' and i > 0:
                    # Relative m after z → add to previous subpath's starting M
                    new_x = sp_x + dx
                    new_y = sp_y + dy
                else:
                    # Absolute M (or first command which SVG treats as absolute)
                    new_x, new_y = dx, dy
                sp_x, sp_y = new_x, new_y
                part = f'M{new_x:.3f} {new_y:.3f} {after_xy}'.rstrip()
            else:
                # Fallback: can't parse coords — keep as-is, best-effort track M
                if cmd == 'M':
                    nums = _NUM_RE.findall(rest)
                    if len(nums) >= 2:
                        sp_x, sp_y = float(nums[0]), float(nums[1])

        # Strip zero-length artifacts immediately before the closing z/Z:
        #   CorelDRAW: l0 0z or l0,0z
        #   Illustrator: h0Z or H0Z
        part = _re.sub(r'\s*[Ll]0[\s,]+0\s*(?=[Zz])', '', part)
        part = _re.sub(r'\s*[Hh]0\s*(?=[Zz])', '', part)

        part = part.strip()
        if part:
            result.append(part)

    return result if result else [d]


def _offset_path_d(d: str, ox: float, oy: float) -> str:
    """
    Return SVG path d-string with all ABSOLUTE coordinate commands shifted by
    (ox, oy).  Relative commands (lowercase) are left unchanged.

    Handles: M L H V C S Q T A Z  (and their lowercase relatives).
    Using direct coordinate offsets avoids transform="translate()" which
    causes CorelDRAW to wrap the path in a locked group.
    """
    if ox == 0.0 and oy == 0.0:
        return d

    NUM_RE = _re.compile(r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')
    segments = _re.split(r'(?=[MmLlHhVvCcSsQqTtAaZz])', d.strip())
    out = []

    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue
        cmd  = seg[0]
        nums = [float(x) for x in NUM_RE.findall(seg[1:])]

        if cmd == 'M':    # x y [x y …]
            n2 = [v + (ox if i % 2 == 0 else oy) for i, v in enumerate(nums)]
        elif cmd == 'L':  # x y [x y …]
            n2 = [v + (ox if i % 2 == 0 else oy) for i, v in enumerate(nums)]
        elif cmd == 'H':  # x [x …]  — horizontal only
            n2 = [v + ox for v in nums]
        elif cmd == 'V':  # y [y …]  — vertical only
            n2 = [v + oy for v in nums]
        elif cmd in ('C', 'S', 'Q', 'T'):  # all use x,y pairs throughout
            n2 = [v + (ox if i % 2 == 0 else oy) for i, v in enumerate(nums)]
        elif cmd == 'A':  # rx ry x-rot large-arc sweep x y  (groups of 7)
            n2 = list(nums)
            for j in range(0, len(n2), 7):
                if j + 5 < len(n2): n2[j + 5] += ox
                if j + 6 < len(n2): n2[j + 6] += oy
        elif cmd in 'Zz':
            out.append(cmd)
            continue
        else:              # relative command — pass through unchanged
            out.append(seg)
            continue

        out.append(cmd + ' '.join(f'{v:.3f}' for v in n2))

    return ''.join(out)


# ── Individual SVG writer ─────────────────────────────────────────────────────

def write_individual_svg(
    output_path: str,
    text_lines: list[str],
    font_name: str,
    font_path: Path,
    template_path: str | None = None,
) -> None:
    """
    Write a single-band SVG file for one gavel order item.
    Text is converted to curves using fontTools — no font dependency in the SVG.
    Uses the template specified by template_path (defaults to standard template).
    """
    band_d, stroke_w, band_w, band_h = _load_band_path(template_path)
    band_cx    = band_w / 2
    band_cy    = band_h / 2

    n          = len(text_lines)
    fs         = FONT_SIZE_PT.get(n, 10.0)
    line_h     = fs * LEADING
    block_h    = (n - 1) * line_h + fs
    baseline_y = band_cy - block_h / 2 + fs * 0.75
    cx         = band_cx

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{band_w/UNITS_PER_INCH:.6f}in" height="{band_h/UNITS_PER_INCH:.6f}in"')
    out.append(f'     viewBox="0 0 {band_w:.3f} {band_h:.3f}">')

    # Band outline — split compound path so CorelDRAW imports each subpath as
    # a normal unlocked Curve (not a locked group).
    for sp in _split_subpaths(band_d):
        out.append(
            f'  <path d="{sp}"'
            f' fill="white" stroke="{BAND_STROKE_COLOR}"'
            f' stroke-width="{stroke_w:.3f}"/>'
        )

    # Text as curves — no font needed in CorelDRAW
    for i, tline in enumerate(text_lines):
        y_abs = baseline_y + i * line_h
        glyph_paths = text_line_to_svg_paths(sanitize_text(tline), str(font_path), fs, cx, y_abs)
        for gp in glyph_paths:
            out.append(f'  <path d="{gp}" fill="#000000"/>')

    out.append('</svg>')
    Path(output_path).write_text("\n".join(out), encoding="utf-8")


# ── Sound block SVG writer ────────────────────────────────────────────────────

SB_MAX_CHARS_PER_LINE = 35   # word-wrap threshold for sound block text


def _wrap_sb_lines(lines: list[str], max_chars: int = SB_MAX_CHARS_PER_LINE) -> list[str]:
    """
    Word-wrap each line to at most max_chars characters.
    Breaks only at spaces so words are never split.
    Shorter lines → more lines → larger auto-fit font size.
    """
    result = []
    for line in lines:
        words   = line.split()
        current = ""
        for word in words:
            if not current:
                current = word
            elif len(current) + 1 + len(word) <= max_chars:
                current += " " + word
            else:
                result.append(current)
                current = word
        if current:
            result.append(current)
    return result or lines


def write_soundblock_svg(output_path: str, text_lines: list[str], font_name: str, font_path: Path) -> None:
    """
    Write a sound block SVG: template border + text-as-curves centered in the
    2.50" × 2.50" engraving area (SB_TEXT_W × SB_TEXT_H) at the centre of
    the 3.75" × 3.75" artwork.

    Lines are word-wrapped at SB_MAX_CHARS_PER_LINE before layout so the
    auto-fit font size is calculated from the wrapped line count (more lines →
    each line is shorter → font can be larger).

    Font size auto-fits:
      1. Height-constrained: fills the text-area height for the given line count.
      2. Width-constrained: scales down if any line would overflow the text-area width.
      3. Capped at SB_MAX_FONT_SIZE.
    """
    text_lines = _wrap_sb_lines(text_lines)
    n = len(text_lines)
    if n == 0:
        return

    # Height-constrained max font size
    fs = min(
        SB_TEXT_H / (1.0 + (n - 1) * LEADING),
        SB_MAX_FONT_SIZE,
    )

    # Width-constrained: measure advance width of each line and scale down if needed
    max_w = max(
        _measure_text_advance(sanitize_text(ln), str(font_path), fs)
        for ln in text_lines
    )
    if max_w > SB_TEXT_W:
        fs *= SB_TEXT_W / max_w

    line_h     = fs * LEADING
    block_h    = (n - 1) * line_h + fs
    baseline_y = SB_CY - block_h / 2 + fs * 0.75

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{SB_W/UNITS_PER_INCH:.6f}in" height="{SB_H/UNITS_PER_INCH:.6f}in"')
    out.append(f'     viewBox="0 0 {SB_W:.3f} {SB_H:.3f}">')

    # Template border elements (rect/path from soundblock_template.svg)
    for el in _load_soundblock_elements():
        out.append(el)

    # Text as curves centered at SB_CX
    for i, tline in enumerate(text_lines):
        y_abs      = baseline_y + i * line_h
        glyph_paths = text_line_to_svg_paths(sanitize_text(tline), str(font_path), fs, SB_CX, y_abs)
        for gp in glyph_paths:
            out.append(f'  <path d="{gp}" fill="#000000"/>')

    out.append('</svg>')
    Path(output_path).write_text("\n".join(out), encoding="utf-8")


def build_soundblock_layout_svg(items: list[dict]) -> str:
    """
    Arrange sound block items on 24" × 12" pages (6 cols × 3 rows, 0.25" gap).
    Each item dict: {"lines": [...], "font": "...", "font_path": Path, ...}
    """
    items_per_page = SB_COLS * SB_ROWS
    num_pages      = max(1, (len(items) + items_per_page - 1) // items_per_page)
    total_h        = num_pages * PAGE_H

    grid_w   = SB_COLS * SB_W + (SB_COLS - 1) * GAP
    grid_h   = SB_ROWS * SB_H + (SB_ROWS - 1) * GAP
    margin_x = (PAGE_W - grid_w) / 2
    margin_y = (PAGE_H - grid_h) / 2

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{24}in" height="{num_pages * 12}in"')
    out.append(f'     viewBox="0 0 {PAGE_W:.3f} {total_h:.3f}">')

    for p in range(1, num_pages):
        y = p * PAGE_H
        out.append(
            f'  <line x1="0" y1="{y:.3f}" x2="{PAGE_W:.3f}" y2="{y:.3f}"'
            f' stroke="#aaaaaa" stroke-width="1" stroke-dasharray="9 4.5"/>'
        )

    sb_elements = _load_soundblock_elements()

    for idx, item in enumerate(items):
        page = idx // items_per_page
        slot = idx %  items_per_page
        col  = slot %  SB_COLS
        row  = slot // SB_COLS

        bx = margin_x + col * (SB_W + GAP)
        by = page * PAGE_H + margin_y + row * (SB_H + GAP)

        # Template border — offset each element's coordinates
        for el in sb_elements:
            # Inline offset: inject a translate transform for layout positioning.
            # CorelDRAW locks transform= groups only on <path> with inherited attrs;
            # a <g transform="translate"> wrapping borders is acceptable here since
            # these are layout guides, not the final production art.
            out.append(f'  <g transform="translate({bx:.3f},{by:.3f})">')
            out.append(f'  {el.strip()}')
            out.append('  </g>')

        # Text as curves
        n       = len(item["lines"])
        fs      = min(SB_TEXT_H / (1.0 + (n - 1) * LEADING), SB_MAX_FONT_SIZE)
        fp      = item.get("font_path")
        if fp:
            max_w = max(
                _measure_text_advance(sanitize_text(ln), str(fp), fs)
                for ln in item["lines"]
            )
            if max_w > SB_TEXT_W:
                fs *= SB_TEXT_W / max_w

        line_h     = fs * LEADING
        block_h    = (n - 1) * line_h + fs
        cx         = bx + SB_CX
        baseline_y = by + SB_CY - block_h / 2 + fs * 0.75

        if fp:
            for i, tline in enumerate(item["lines"]):
                y_abs = baseline_y + i * line_h
                glyph_paths = text_line_to_svg_paths(
                    sanitize_text(tline), str(fp), fs, cx, y_abs
                )
                for gp in glyph_paths:
                    out.append(f'  <path d="{gp}" fill="#000000"/>')

    out.append('</svg>')
    return "\n".join(out)


# ── Trello helpers ────────────────────────────────────────────────────────────

def _trello_get(path: str, **params) -> object:
    base_params = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    base_params.update(params)
    url = "https://api.trello.com/1" + path + "?" + "&".join(f"{k}={urllib.request.quote(str(v))}" for k, v in base_params.items())
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())


def _trello_post(path: str, data: dict) -> dict:
    base = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    base.update(data)
    body = "&".join(f"{k}={urllib.request.quote(str(v), safe='')}" for k, v in base.items())
    url  = "https://api.trello.com/1" + path
    req  = urllib.request.Request(
        url,
        data=body.encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())


def trello_get_processed_order_numbers() -> set[str]:
    """
    Return every order number already recorded on the Customs board.
    Scans all card titles AND descriptions for strings that look like
    Amazon order numbers (e.g. 112-4450347-1581840).
    """
    import re
    ORDER_RE = re.compile(r'\b\d{3}-\d{7}-\d{7}\b')

    try:
        boards = _trello_get("/members/me/boards", fields="name")
        board  = next((b for b in boards if b["name"].lower() == TRELLO_BOARD_NAME.lower()), None)
        if not board:
            return set()

        # Fetch all cards (open + archived) so we don't reprocess archived orders
        cards = _trello_get(
            f"/boards/{board['id']}/cards/all",
            fields="name,desc",
        )

        found = set()
        for card in cards:
            text = (card.get("name", "") + "\n" + card.get("desc", ""))
            for m in ORDER_RE.finditer(text):
                found.add(m.group())

        return found

    except Exception as e:
        print(f"  ⚠  Could not fetch Trello history (will process all orders): {e}")
        return set()


# Item name keywords that route an order to the Silver Band card
SILVER_BAND_KEYWORDS = ["silver band"]


def _is_silver_band_order(ship: dict) -> bool:
    """Return True if any gavel item in the shipment has 'silver band' in its name."""
    for item in ship.get("items", []):
        if is_gavel(item):
            name = (item.get("name") or "").lower()
            if any(kw in name for kw in SILVER_BAND_KEYWORDS):
                return True
    return False


def trello_create_gavel_card(
    order_numbers: list[str],
    variant: str | None = None,
    rerun: bool = False,
) -> tuple[str, str]:
    """
    Create a Trello card on the 'customs' board under 'Customs Ready For Production'.
    variant : optional label inserted into the title, e.g. "Silver Band"
              → "Custom Gavel Order (Silver Band) {DATE}"
    rerun   : when True the title is just the order number(s) + ' *' suffix
              instead of the standard auto-generated date title.
    Title suffix Set 2, Set 3 … applied when a matching card already exists today.
    Returns: (card_url, card_id)
    """
    if not TRELLO_API_KEY or TRELLO_API_KEY == "YOUR_TRELLO_API_KEY":
        print("  ⚠  Trello credentials not set — skipping card creation.")
        return "", ""

    desc = "\n".join(order_numbers)

    # Find board
    boards = _trello_get("/members/me/boards", fields="name")
    board  = next((b for b in boards if b["name"].lower() == TRELLO_BOARD_NAME.lower()), None)
    if not board:
        raise ValueError(f"Trello board '{TRELLO_BOARD_NAME}' not found. "
                         f"Available: {[b['name'] for b in boards]}")

    # Find list on board
    lists       = _trello_get(f"/boards/{board['id']}/lists", fields="name")
    trello_list = next((l for l in lists if l["name"].lower() == TRELLO_LIST_NAME.lower()), None)
    if not trello_list:
        raise ValueError(f"List '{TRELLO_LIST_NAME}' not found on board '{TRELLO_BOARD_NAME}'. "
                         f"Available: {[l['name'] for l in lists]}")

    if rerun:
        # Title is just the order number(s) with a * suffix to signal it was re-run
        orders_str = ", ".join(order_numbers)
        card_name  = f"{orders_str} *"
    else:
        date_str  = datetime.now().strftime("%m/%d/%Y")
        if variant:
            base_name = f"Custom Gavel Order ({variant}) {date_str}"
        else:
            base_name = f"Custom Gavel Order {date_str}"

        # Check for existing cards with the same date to determine Set suffix.
        import re as _re2
        all_cards  = _trello_get(f"/boards/{board['id']}/cards", fields="name")
        base_lower = base_name.lower()
        suffix_pat = _re2.compile(
            _re2.escape(base_lower) + r'(?:\s+set\s+(\d+))?$'
        )
        existing_nums = []
        for c in all_cards:
            m = suffix_pat.match(c.get("name", "").lower())
            if m:
                existing_nums.append(int(m.group(1)) if m.group(1) else 1)

        if not existing_nums:
            card_name = base_name
        else:
            next_num  = max(existing_nums) + 1
            card_name = f"{base_name} Set {next_num}"

    print(f"  Card title: {card_name}")

    # Create card
    card = _trello_post("/cards", {
        "idList": trello_list["id"],
        "name":   card_name,
        "desc":   desc,
    })
    return card.get("url", ""), card.get("id", "")


def trello_attach_file(card_id: str, file_path: str, mime_type: str = "image/svg+xml") -> None:
    """Upload any file as an attachment to a Trello card."""
    filename = Path(file_path).name
    boundary = b"TrelloBoundaryXx7zA9qPmN"

    with open(file_path, "rb") as f:
        file_data = f.read()

    body = b""
    for field_name, value in [("key", TRELLO_API_KEY), ("token", TRELLO_TOKEN), ("name", filename)]:
        body += b"--" + boundary + b"\r\n"
        body += f'Content-Disposition: form-data; name="{field_name}"\r\n\r\n'.encode()
        body += value.encode() + b"\r\n"

    body += b"--" + boundary + b"\r\n"
    body += f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'.encode()
    body += f"Content-Type: {mime_type}\r\n\r\n".encode()
    body += file_data + b"\r\n"
    body += b"--" + boundary + b"--\r\n"

    url = f"https://api.trello.com/1/cards/{card_id}/attachments"
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary.decode()}"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as r:
        r.read()


def trello_attach_svg(card_id: str, file_path: str) -> None:
    """Upload an SVG file as an attachment to a Trello card."""
    trello_attach_file(card_id, file_path, mime_type="image/svg+xml")


# ── SVG layout builder ────────────────────────────────────────────────────────

def build_layout_svg(items: list[dict]) -> str:
    """
    Arrange all gavel items on 24" × 12" pages (3 cols × 10 rows, 0.25" gap).
    Items may use different band templates (e.g. 7" walnut/black vs standard 6.875").
    Column width = widest template in the batch; each band is centred within its slot.
    """
    # Determine the widest band template used — that sets the column width.
    # All templates share the same band height (950 units).
    col_w = BAND_W   # default
    for item in items:
        tp = item.get("template_path") or TEMPLATE_PATH
        _, _, bw, _ = _load_band_path(tp)
        if bw > col_w:
            col_w = bw
    band_h = BAND_H  # 950, same across all templates

    items_per_page = COLS * ROWS
    num_pages      = max(1, (len(items) + items_per_page - 1) // items_per_page)
    total_h        = num_pages * PAGE_H

    grid_w   = COLS * col_w  + (COLS - 1) * BAND_COL_GAP
    grid_h   = ROWS * band_h + (ROWS - 1) * BAND_ROW_GAP
    margin_x = (PAGE_W - grid_w) / 2
    margin_y = (PAGE_H - grid_h) / 2

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{24}in" height="{num_pages * 12}in"')
    out.append(f'     viewBox="0 0 {PAGE_W:.3f} {total_h:.3f}">')

    for p in range(1, num_pages):
        y = p * PAGE_H
        out.append(
            f'  <line x1="0" y1="{y:.3f}" x2="{PAGE_W:.3f}" y2="{y:.3f}"'
            f' stroke="#aaaaaa" stroke-width="1" stroke-dasharray="9 4.5"/>'
        )

    for idx, item in enumerate(items):
        page = idx // items_per_page
        slot = idx %  items_per_page
        col  = slot %  COLS
        row  = slot // COLS

        # Per-item template dimensions
        tp               = item.get("template_path") or TEMPLATE_PATH
        band_d, stroke_w, band_w, _ = _load_band_path(tp)
        band_cx_local    = band_w / 2

        # Slot top-left; centre narrower bands within the column slot
        slot_x = margin_x + col * (col_w + BAND_COL_GAP)
        bx     = slot_x + (col_w - band_w) / 2
        by     = page * PAGE_H + margin_y + row * (band_h + BAND_ROW_GAP)
        cx     = bx + band_cx_local

        # Font metrics
        text_lines = item["lines"]
        n          = len(text_lines)
        fs         = FONT_SIZE_PT.get(n, 10.0)
        line_h     = fs * LEADING
        block_h    = (n - 1) * line_h + fs
        baseline_y = by + (band_h / 2) - block_h / 2 + fs * 0.75

        # Band outline — one <path> per subpath, coordinates baked in directly
        for sp in _split_subpaths(band_d):
            sp_off = _offset_path_d(sp, bx, by)
            out.append(
                f'  <path d="{sp_off}"'
                f' fill="white" stroke="{BAND_STROKE_COLOR}"'
                f' stroke-width="{stroke_w:.3f}"/>'
            )

        # ── text as curves — no font dependency in the output SVG ──
        font_path_obj = item.get("font_path")
        if font_path_obj:
            for i, tline in enumerate(text_lines):
                y_abs = baseline_y + i * line_h
                glyph_paths = text_line_to_svg_paths(
                    sanitize_text(tline), str(font_path_obj), fs, cx, y_abs
                )
                for gp in glyph_paths:
                    out.append(f'  <path d="{gp}" fill="#000000"/>')


    out.append('</svg>')
    return "\n".join(out)


# ── ShipStation helpers ───────────────────────────────────────────────────────

import time as _time

def ss_get(path: str, params: dict, _retries: int = 3) -> dict:
    """GET from ShipStation API with timeout and retry on rate-limit (429)."""
    url = f"{BASE_URL}{path}?" + "&".join(f"{k}={urllib.request.quote(str(v))}" for k, v in params.items())
    req = urllib.request.Request(url, headers={"Authorization": SS_AUTH})
    for attempt in range(1, _retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait = 60 if attempt == 1 else 120
                print(f"\n  ⚠  ShipStation rate limit hit — waiting {wait}s (attempt {attempt}/{_retries})...")
                _time.sleep(wait)
            else:
                raise
    raise RuntimeError(f"ShipStation API still rate-limiting after {_retries} retries")


def is_gavel(item) -> bool:
    """Return True if the order item is a gavel band product.
    Checks SKU patterns first, then falls back to item name keywords
    to catch products with random Amazon-generated SKUs."""
    if isinstance(item, str):
        sku = item   # backwards-compat: accept bare SKU string
        name = ""
    else:
        sku  = item.get("sku") or ""
        name = item.get("name") or ""
    if not sku and not name:
        return False
    if any(p in sku for p in GAVEL_SKU_PATTERNS):
        return True
    name_lower = name.lower()
    return any(kw in name_lower for kw in GAVEL_NAME_KEYWORDS)


def fetch_orders_by_number(order_numbers: set) -> list[dict]:
    """Fetch specific orders from ShipStation by order number, regardless of status."""
    results = []
    for order_num in order_numbers:
        try:
            data = ss_get("/orders", {"orderNumber": order_num, "pageSize": 10})
            for s in data.get("orders", []):
                if s.get("orderNumber") == order_num:
                    results.append(s)
                    break
        except Exception as e:
            print(f"  ⚠  Could not fetch order {order_num} from ShipStation: {e}")
    return results


def fetch_gavel_shipments(days: int) -> list[dict]:
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
    params = {
        "pageSize":        PAGE_SIZE,
        "sortBy":          "ModifyDate",
        "sortDir":         "DESC",
        "modifyDateStart": cutoff,
        "orderStatus":     "awaiting_shipment",
    }
    page, total_pages, results = 1, 1, []
    while page <= total_pages:
        params["page"] = page
        data = ss_get("/orders", params)
        for s in data.get("orders", []):
            for item in s.get("items", []):
                if is_gavel(item) and any(
                    o.get("name") == "CustomizedURL" for o in item.get("options", [])
                ):
                    results.append(s)
                    break
        total_pages = data.get("pages", 1)
        print(f"  Page {page}/{total_pages} — {len(results)} gavel orders so far")
        page += 1
    return results


# ── Amazon customization helpers ──────────────────────────────────────────────

def fetch_customization(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        content = r.read()
    z = zipfile.ZipFile(io.BytesIO(content))
    for name in z.namelist():
        if name.endswith(".json"):
            return json.loads(z.read(name))
    return {}


# ── Sound block option detection ──────────────────────────────────────────────

# Labels that identify the "which sound block?" dropdown/radio field
_SB_OPTION_LABELS = {
    "add sound block", "sound block", "sound block option",
    "add a sound block", "add sound block?",
}

# Labels that identify the sound block's engraving text field
_SB_TEXT_LABELS = {
    "add your custom text", "custom text", "sound block text",
    "engraving text", "add custom text", "your custom text",
    "block engraving", "custom engraving",
}

# Substrings in the option value that mean "engrave the sound block"
_SB_ENGRAVE_KEYWORDS = (
    "custom engraved", "custom engrave", "add custom", "add engraved",
    "engraved sound block",
)

# Substrings meaning "no engraving" (also covers "Gavel Only")
_SB_NO_ENGRAVE_KEYWORDS = (
    "gavel only", "no engraving", "no engrave", "not engraved",
    "without engraving", "no sound block",
)

# ── Suede bag option detection ────────────────────────────────────────────────

# Labels that identify the suede/velvet bag dropdown
_SUEDE_OPTION_LABELS = {
    "suede gift bag", "suede bag", "gift bag",
    "would you like to include a suede gift bag",
    "would you like to include a suede gift bag?",
}

# Substrings in the option value that mean "yes, add a bag"
_SUEDE_YES_KEYWORDS = (
    "add suede", "add gavel suede", "add both", "suede gavel bag",
    "suede sound block", "gift bag",
)


def _classify_sb_option(value: str) -> str:
    """Return 'custom_engraved', 'no_engraving', or 'unknown'."""
    v = value.strip().lower()
    if any(kw in v for kw in _SB_ENGRAVE_KEYWORDS):
        return "custom_engraved"
    if any(kw in v for kw in _SB_NO_ENGRAVE_KEYWORDS):
        return "no_engraving"
    print(f"\n   ⚠  Unrecognised sound block option value: '{value}' — treating as no engraving")
    return "unknown"


def parse_customization(cust_json: dict) -> dict:
    """
    Parse Amazon customization JSON and return:
    {
        "band_lines" : ["Line1", "Line2", ...],
        "font"       : "Arial",
        "sb_option"  : "custom_engraved" | "no_engraving" | None,
        "sb_lines"   : ["Line1", ...],   # text for sound block
        "sb_font"    : "Arial",
    }
    """
    band_lines: list[str] = []
    sb_lines:   list[str] = []
    font       = "Arial"
    sb_font    = "Arial"
    sb_option: str | None = None
    wants_suede_gavel = False
    wants_suede_sb    = False

    def _classify_suede(val: str) -> None:
        nonlocal wants_suede_gavel, wants_suede_sb
        v = val.strip().lower()
        if not any(kw in v for kw in _SUEDE_YES_KEYWORDS):
            return
        has_sb  = "sound block" in v or "block" in v
        has_gav = "gavel" in v or "both" in v
        if "both" in v:
            wants_suede_gavel = True
            wants_suede_sb    = True
        elif has_sb:
            wants_suede_sb    = True
        else:
            wants_suede_gavel = True

    # ── Version 3.0 structure ──────────────────────────────────────────────
    surfaces = (
        cust_json.get("version3.0", {})
                 .get("customizationInfo", {})
                 .get("surfaces", [])
    )
    for surf in surfaces:
        for area in surf.get("areas", []):
            label = (area.get("label") or area.get("name") or "").strip().lower()
            # option_val: value for dropdown/option areas (never added to band text)
            # text: value for engraving text areas (added to band/sb lines)
            option_val = (area.get("optionValue") or "").strip()
            text       = (area.get("text") or area.get("value") or "").strip()
            ff         = area.get("fontFamily", "").strip()

            # Suede bag option dropdown — use option_val
            if any(kw in label for kw in _SUEDE_OPTION_LABELS):
                if option_val:
                    _classify_suede(option_val)
                continue

            # Sound block option dropdown — use option_val (or text as fallback)
            if any(kw in label for kw in _SB_OPTION_LABELS):
                val = option_val or text
                if val:
                    sb_option = _classify_sb_option(val)
                continue

            # Sound block engraving text
            if any(kw in label for kw in _SB_TEXT_LABELS):
                if ff and sb_font == "Arial":
                    sb_font = ff
                if text:
                    sb_lines.append(text)
                continue

            # Regular band text
            if ff and font == "Arial":
                font = ff
            if text:
                band_lines.append(text)

    # ── Legacy / fallback structure (walk) ────────────────────────────────
    # Run when band lines OR sb lines are still missing.
    # v3.0 surfaces can store sound block text as "ImagePrinting" (no text
    # field), so sb_lines may be empty even when band_lines was filled.
    # We only accumulate each list that is still empty to avoid duplicates.
    _need_band = not band_lines
    _need_sb   = not sb_lines
    if _need_band or _need_sb:
        def walk(node, depth=0):
            nonlocal font, sb_font, sb_option
            if isinstance(node, dict):
                node_type  = node.get("type", "")
                node_label = (node.get("label") or node.get("name") or "").strip().lower()

                if node_type == "TextCustomization":
                    v = (node.get("inputValue") or "").strip()
                    if any(kw in node_label for kw in _SB_TEXT_LABELS):
                        if v and _need_sb:
                            sb_lines.append(v)
                    else:
                        if v and _need_band:
                            band_lines.append(v)

                elif node_type == "FontCustomization":
                    fam = node.get("fontSelection", {}).get("family", "")
                    if fam:
                        if any(kw in node_label for kw in _SB_TEXT_LABELS):
                            if sb_font == "Arial" and _need_sb:
                                sb_font = fam
                        else:
                            if font == "Arial" and _need_band:
                                font = fam

                elif node_type in (
                    "DropdownCustomization", "RadioCustomization",
                    "SelectCustomization", "CheckboxCustomization",
                ):
                    v = (node.get("inputValue") or node.get("selectedValue") or "").strip()
                    if any(kw in node_label for kw in _SB_OPTION_LABELS) and v:
                        sb_option = _classify_sb_option(v)
                    if any(kw in node_label for kw in _SUEDE_OPTION_LABELS) and v:
                        _classify_suede(v)

                for val in node.values():
                    walk(val, depth + 1)

            elif isinstance(node, list):
                for item in node:
                    walk(item, depth + 1)

        walk(cust_json.get("customizationData", {}))

    # If sb_option is still None but we found sb_lines, assume custom_engraved
    if sb_option is None and sb_lines:
        sb_option = "custom_engraved"

    # Split any lines that arrived with embedded newlines (Amazon sometimes
    # stores the full text block as a single field with \n separators)
    def _split_nl(lines: list[str]) -> list[str]:
        result = []
        for ln in lines:
            for part in ln.splitlines():
                part = part.strip()
                if part:
                    result.append(part)
        return result

    return {
        "band_lines":  _split_nl(band_lines),
        "font":        font,
        "sb_option":          sb_option,
        "sb_lines":           _split_nl(sb_lines),
        "sb_font":            sb_font,
        "wants_suede_gavel":  wants_suede_gavel,
        "wants_suede_sb":     wants_suede_sb,
    }


# ── Packing slip ──────────────────────────────────────────────────────────────

def _html_esc(s: str) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def _barcode_svg(order_num: str) -> str:
    """Return an <img> tag with a base64-encoded 300 DPI PNG Code128 barcode.
    PNG prints reliably across all print dialogs; inline SVG does not."""
    try:
        import barcode as _bc
        from barcode.writer import ImageWriter as _IW
        import io as _io, base64 as _b64
        code = _bc.get("code128", order_num, writer=_IW())
        buf = _io.BytesIO()
        code.write(buf, options={
            "module_width":  0.38,  # mm per bar module (narrower)
            "module_height": 6.5,   # mm bar height (slightly taller)
            "quiet_zone":    4.0,   # mm white space each side
            "font_size":     0,
            "text_distance": 0,
            "write_text":    False,
            "dpi":           300,
        })
        data = _b64.b64encode(buf.getvalue()).decode()
        return (f'<img src="data:image/png;base64,{data}" '
                f'style="width:100%;height:auto;display:block">')
    except Exception as e:
        print(f"\n   ⚠  Barcode generation failed for {order_num}: {e}")
        return f'<div style="font-family:monospace;font-size:8px;letter-spacing:2px">{order_num}</div>'


def _wants_suede(ship: dict) -> bool:
    for item in ship.get("items", []):
        name = (item.get("name") or "").lower()
        if "suede" in name or "velvet bag" in name or "pouch" in name:
            return True
    return False


def _sku_has_soundblock(sku: str) -> bool:
    """Return True if the SKU indicates a sound block variant (ends with SB or -SB)."""
    return sku.upper().strip().endswith("SB")


def write_packing_slip(output_path: str, order_num: str, customer: str,
                       ship: dict, slip_items: list[dict]) -> None:
    """Generate a print-ready HTML work order / packing slip for one order."""
    from datetime import date as _date

    # ── Address from ShipStation ───────────────────────────────────────────────
    ship_to   = ship.get("shipTo") or {}
    addr1     = _html_esc(ship_to.get("street1") or "")
    addr2     = _html_esc(ship_to.get("street2") or "")
    city      = _html_esc(ship_to.get("city") or "")
    state     = _html_esc(ship_to.get("state") or "")
    postal    = _html_esc(ship_to.get("postalCode") or "")
    country   = _html_esc(ship_to.get("country") or "")
    city_line = ", ".join(filter(None, [city, f"{state} {postal}".strip(), country]))

    raw_svc    = ship.get("requestedShippingService") or ship.get("serviceCode") or ""
    svc_label  = _html_esc(raw_svc)

    raw_odate = ship.get("orderDate") or ""
    try:
        from datetime import datetime as _dt
        odate_fmt = _dt.fromisoformat(raw_odate[:10]).strftime("%m/%d/%Y")
    except Exception:
        odate_fmt = _date.today().strftime("%m/%d/%Y")

    print_date = _date.today().strftime("%m/%d/%Y")

    # ── Stamp detection ────────────────────────────────────────────────────────
    has_sb_custom  = any(i.get("sb_option") == "custom_engraved" for i in slip_items)
    has_sb_no      = any(
        (i.get("sb_option") in ("no_engraving", "unknown") or _sku_has_soundblock(i.get("sku", "")))
        and not i.get("want_sb")
        for i in slip_items
    )
    has_suede_gavel = (
        any(i.get("wants_suede_gavel") for i in slip_items)
        or _wants_suede(ship)
    )
    has_suede_sb = any(i.get("wants_suede_sb") for i in slip_items)

    stamps = []
    if has_sb_custom:
        stamps.append('<div class="stamp stamp-sb">Sound Block — Custom</div>')
    if has_sb_no:
        stamps.append('<div class="stamp stamp-sb-no">Sound Block — NO</div>')
    if has_suede_gavel:
        stamps.append('<div class="stamp stamp-suede">Gavel Gift Bag</div>')
    if has_suede_sb:
        stamps.append('<div class="stamp stamp-suede">Sound Block Gift Bag</div>')
    stamp_html = "\n".join(stamps) if stamps else '<div class="no-stamp">No Special Instructions</div>'

    # ── Items table rows ───────────────────────────────────────────────────────
    rows_html = ""
    for it in slip_items:
        eng_lines = "".join(
            f"<div class='eng-line'><span class='eng-num'>L{i}:</span> {_html_esc(ln)}</div>"
            for i, ln in enumerate(it["text_lines"], 1)
        )
        sb_block = ""
        if it.get("want_sb") and it.get("sb_lines"):
            sb_text = " / ".join(it["sb_lines"])
            sb_block = f"""<div class="sb-divider">- - - - - - - - - - - - - - - - - - - - - - - - -</div>
            <div class="sb-section">
              <div class="sb-header">Sound Block: <span class="font-note">{_html_esc(it.get('sb_font',''))}</span></div>
              <div class="eng-line"><span class="eng-num">Text:</span> {_html_esc(sb_text)}</div>
            </div>"""
        elif it.get("sb_font_error"):
            sb_block = f'<div class="sb-error">&#9888; Sound block skipped — font "{_html_esc(it["sb_font_error"])}" not found</div>'

        rows_html += f"""<tr>
          <td class="qty-cell">{_html_esc(str(it['qty']))}</td>
          <td class="detail-cell">
            <div class="item-name">{_html_esc(it.get('item_name') or it['sku'])}</div>
            <div class="item-meta"><b>SKU:</b> {_html_esc(it['sku'])}</div>
            <div class="eng-block">
              <div class="eng-section-label">Gavel: <span class="font-note">{_html_esc(it['font'])}</span></div>
              {eng_lines}
            </div>
            {sb_block}
          </td>
        </tr>"""

    # ── Barcode ────────────────────────────────────────────────────────────────
    barcode_svg = _barcode_svg(order_num)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Work Order — {_html_esc(order_num)}</title>
<style>
  @page{{size:4in 6in;margin:0.18in 0.2in}}
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
  html{{background:#fff}}
  body{{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;font-size:8px;line-height:1.35}}

  /* ── TOP HEADER: ship-to left, stamps right (table for xhtml2pdf compat) ── */
  .top-table{{width:100%;border-collapse:collapse;margin-bottom:5px}}
  .ship-to-cell{{vertical-align:top;width:62%}}
  .stamps-cell{{vertical-align:top;width:38%;text-align:right}}
  .ship-to-label{{font-size:7px;margin-bottom:2px}}
  .ship-name{{font-size:10px;font-weight:700;line-height:1.3}}
  .ship-addr{{font-size:9px;font-weight:700;line-height:1.35}}

  /* ── Stamps ── */
  .stamp{{
    display:block;margin-bottom:3px;padding:4px 8px;
    font-size:8px;font-weight:900;text-transform:uppercase;letter-spacing:.05em;
    border:1px solid #000;border-radius:2px;background:#fff;color:#000;
    -webkit-print-color-adjust:exact;print-color-adjust:exact;
  }}
  .stamp-sb{{border-color:#000}}
  .stamp-sb-no{{border-color:#000}}
  .stamp-suede{{border-color:#000}}
  .no-stamp{{font-size:7px;color:#777;font-style:italic}}

  /* ── Divider ── */
  .divider{{border:none;border-top:1px solid #000;margin:5px 0}}

  /* ── Order ID + barcode row ── */
  .order-row{{margin-bottom:5px}}
  .order-id{{font-size:10px;font-weight:700;margin-bottom:3px}}
  .order-sub{{font-size:7.5px;margin-bottom:4px}}
  .barcode-wrap{{width:100%;margin:4px 0}}

  /* ── Info box ── */
  .info-box{{display:flex;border:1px solid #000;margin-bottom:6px}}
  .info-left{{flex:1;padding:4px 5px;border-right:1px solid #000}}
  .info-right{{flex:1;padding:4px 5px}}
  .info-label{{font-weight:700;font-size:8px;margin-bottom:2px}}
  .info-row{{display:flex;gap:3px;margin-bottom:2px;font-size:7.5px}}
  .info-key{{font-weight:700;white-space:nowrap}}

  /* ── Items table ── */
  .items-table{{width:100%;border-collapse:collapse;margin-bottom:6px}}
  .items-table th{{border:1px solid #000;padding:3px 5px;font-weight:700;
                   font-size:8px;text-align:left;background:#f0f0f0;
                   -webkit-print-color-adjust:exact;print-color-adjust:exact}}
  .items-table td{{border:1px solid #000;padding:4px 5px;vertical-align:top}}
  .qty-cell{{width:52px;text-align:center;font-size:9px;font-weight:700}}
  .detail-cell{{font-size:8px}}
  .item-name{{font-weight:700;font-size:8.5px;margin-bottom:2px}}
  .item-meta{{margin-bottom:1px;color:#333}}
  .eng-block{{margin:3px 0 2px 0;padding-left:2px}}
  .eng-section-label{{font-weight:700;font-size:7.5px;margin-bottom:2px}}
  .eng-line{{margin-bottom:1px;font-size:9.5px;font-weight:700}}
  .eng-num{{font-weight:700;color:#555;font-size:8px;display:inline-block;width:28px}}

  /* ── Sound block section ── */
  .sb-divider{{font-size:7px;color:#777;letter-spacing:1px;margin:4px 0 3px 0;border:none;background:none}}
  .sb-section{{margin-top:0;padding-left:0;border:none}}
  .sb-header{{font-weight:700;font-size:7.5px;margin-bottom:2px}}
  .font-note{{font-style:italic;font-weight:400;font-size:7px}}
  .sb-error{{margin-top:4px;border:1.5px dashed #000;padding:3px 5px;font-size:7px}}

  /* ── Footer ── */
  .footer{{margin-top:6px;font-size:7px;line-height:1.4;border-top:1px solid #000;padding-top:5px}}
  .footer-heading{{font-weight:700;margin-bottom:2px}}
  .footer div{{margin-bottom:0;border:none}}

  @media print{{
    .items-table tr{{page-break-inside:avoid;break-inside:avoid}}
  }}
</style>
</head>
<body>

<!-- TOP: Ship To (left) | Stamps (right) -->
<table class="top-table">
<tr>
  <td class="ship-to-cell">
    <div class="ship-to-label">Ship To:</div>
    <div class="ship-name">{_html_esc(customer)}</div>
    <div class="ship-addr">{addr1}</div>
    {"<div class='ship-addr'>" + addr2 + "</div>" if addr2 else ""}
    <div class="ship-addr">{city_line}</div>
  </td>
  <td class="stamps-cell">{stamp_html}</td>
</tr>
</table>

<hr class="divider">

<!-- Order ID + barcode -->
<div class="order-row">
  <div class="order-id">Order ID: {_html_esc(order_num)}</div>
  <div class="order-sub">Thank you for your custom gavel order from All Quality.</div>
  <div class="barcode-wrap">{barcode_svg}</div>
</div>

<!-- Info box -->
<div class="info-box">
  <div class="info-left">
    <div class="info-label">Shipping Address:</div>
    <div>{_html_esc(customer)}</div>
    <div>{addr1}</div>
    {"<div>" + addr2 + "</div>" if addr2 else ""}
    <div>{city_line}</div>
  </div>
  <div class="info-right">
    <div class="info-row"><span class="info-key">Order Date:</span><span>{odate_fmt}</span></div>
    <div class="info-row"><span class="info-key">Print Date:</span><span>{print_date}</span></div>
    {"<div class='info-row'><span class='info-key'>Shipping Service:</span><span>" + svc_label + "</span></div>" if svc_label else ""}
    <div class="info-row"><span class="info-key">Order #:</span><span>{_html_esc(order_num)}</span></div>
  </div>
</div>

<!-- Items table -->
<table class="items-table">
  <thead>
    <tr>
      <th>Qty Ordered</th>
      <th>Product Details</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>

<div class="footer">
  <div class="footer-heading">Returning your item:</div>
  <div>Go to &ldquo;Your Account&rdquo; on Amazon.com, click &ldquo;Your Orders&rdquo; and then click the &ldquo;seller profile&rdquo; link for this order to get information about the return and refund policies that apply.</div>
  <div>Visit https://www.amazon.com/returns to print a return shipping label. Please have your order ID ready.</div>
  <div><b>Thanks for buying on Amazon Marketplace.</b> To provide feedback for the seller please visit www.amazon.com/feedback. To contact the seller, go to Your Orders in Your Account. Click the seller&rsquo;s name under the appropriate product. Then, in the &ldquo;Further Information&rdquo; section, click &ldquo;Contact the Seller.&rdquo;</div>
</div>

</body>
</html>"""

    Path(output_path).write_text(html, encoding="utf-8")


def html_to_pdf(html_path: str, pdf_path: str) -> bool:
    """Convert a work-order HTML file to PDF.
    Tries WeasyPrint first (preferred on Linux/Cloud Run); falls back to
    xhtml2pdf (pure-Python, works on Windows without GTK).
    Returns True on success."""
    # ── Try WeasyPrint (requires GTK / available on Linux) ──────────────────
    try:
        from weasyprint import HTML as _WP_HTML
        _WP_HTML(filename=html_path).write_pdf(pdf_path)
        return True
    except ImportError:
        pass  # not installed — fall through to xhtml2pdf
    except Exception as e:
        # Installed but missing native libs (e.g. Windows without GTK)
        if "libgobject" not in str(e) and "cannot load library" not in str(e):
            print(f"   ⚠  WeasyPrint error for {Path(html_path).name}: {e}")
        # fall through to xhtml2pdf
    # ── Fallback: xhtml2pdf (pure Python) ───────────────────────────────────
    try:
        from xhtml2pdf import pisa
        html_text = Path(html_path).read_text(encoding="utf-8")
        with open(pdf_path, "wb") as fh:
            result = pisa.CreatePDF(html_text, dest=fh, encoding="utf-8")
        if result.err:
            print(f"   ⚠  PDF conversion had errors for {Path(html_path).name}: {result.err}")
            return False
        return True
    except Exception as e:
        print(f"   ⚠  PDF conversion failed for {Path(html_path).name}: {e}")
        return False


def merge_pdfs(pdf_paths: list, output_path: str) -> bool:
    """Merge a list of PDF files into a single PDF (one order per page). Returns True on success."""
    if not pdf_paths:
        return False
    try:
        from pypdf import PdfWriter
        writer = PdfWriter()
        for p in pdf_paths:
            writer.append(str(p))
        with open(output_path, "wb") as fh:
            writer.write(fh)
        return True
    except Exception as e:
        print(f"   ⚠  PDF merge failed: {e}")
        return False


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate SVG gavel band layouts")
    parser.add_argument("--days",   type=int, default=1,  help="Days to look back (default 1)")
    parser.add_argument("--output", default="gavel_eps",  help="Output directory")
    parser.add_argument("--orders", nargs="+", metavar="ORDER_NUM",
                        help="Force-process specific order numbers (bypasses Trello dedup)")
    parser.add_argument("--trello-card", metavar="CARD_ID",
                        help="Upload to this existing Trello card instead of creating a new one")
    args = parser.parse_args()

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    force_orders    = set(args.orders) if args.orders else set()
    target_card_id  = args.trello_card or None

    print(f"\n{'='*60}")
    print("Gavel Band SVG Layout Generator")
    print(f"Look-back: {args.days} day(s)  |  Output: {out_dir.resolve()}")
    print(f"{'='*60}\n")

    if force_orders:
        # Fetch requested orders directly — bypasses status filter and look-back window
        print(f"Fetching {len(force_orders)} specific order(s) from ShipStation...")
        shipments = fetch_orders_by_number(force_orders)
        found = {s.get("orderNumber", "") for s in shipments}
        missing = force_orders - found
        if missing:
            print(f"  ⚠  Order(s) not found in ShipStation: {', '.join(sorted(missing))}")
        if not shipments:
            print("None of the requested orders found — nothing to do.")
            return
        print(f"  Force-processing {len(shipments)} order(s): {', '.join(s.get('orderNumber','') for s in shipments)}\n")
    else:
        print("Fetching gavel orders from ShipStation...")
        shipments = fetch_gavel_shipments(days=args.days)
        print(f"\nFound {len(shipments)} gavel order(s) awaiting shipment")

        if not shipments:
            print("Nothing to do.")
            return
        # Cross-check Trello — skip orders already on the Customs board
        print("\nCross-checking Trello Customs board for already-processed orders...")
        already_in_trello = trello_get_processed_order_numbers()
        print(f"  {len(already_in_trello)} order number(s) found on Trello board")

        new_shipments = [
            s for s in shipments
            if s.get("orderNumber", "") not in already_in_trello
        ]
        skipped = len(shipments) - len(new_shipments)
        if skipped:
            print(f"  Skipping {skipped} already-processed order(s)")
        print(f"  {len(new_shipments)} new order(s) to process\n")

        if not new_shipments:
            print("All orders already on Trello — nothing to do.")
            return

        shipments = new_shipments

    all_items        = []   # gavel band items for bulk layout
    all_sb_items     = []   # sound block items for bulk layout
    standard_svgs     = []   # SVGs for standard orders
    silver_svgs       = []   # SVGs for silver band orders
    standard_pdf_paths = []  # work order PDFs for standard orders
    silver_pdf_paths   = []  # work order PDFs for silver orders
    summary           = []
    order_nums       = []         # standard orders
    silver_order_nums = []        # silver band orders
    ok               = 0
    errors           = 0

    for idx, ship in enumerate(shipments, 1):
        order_num  = ship.get("orderNumber", f"order_{idx}")
        customer   = (ship.get("shipTo") or {}).get("name", "")
        is_silver  = _is_silver_band_order(ship)
        print(f"[{idx}/{len(shipments)}] {order_num} — {customer}")

        gavel_items = [
            item for item in ship.get("items", [])
            if is_gavel(item) and any(
                o.get("name") == "CustomizedURL" for o in item.get("options", [])
            )
        ]

        order_had_success = False
        slip_items        = []

        for item_idx, item in enumerate(gavel_items, 1):
            sku           = item.get("sku", "NOSKU")
            item_name     = item.get("name", "")
            qty           = item.get("quantity", 1)
            band_template = _select_band_template(item_name)
            url = next(
                o["value"] for o in item.get("options", [])
                if o.get("name") == "CustomizedURL"
            )

            safe_order = order_num.replace("/", "-").replace("\\", "-")
            safe_sku   = sku.replace("/", "-").replace("\\", "-")
            svg_name   = f"{safe_order}_{item_idx}_{safe_sku}.svg"
            svg_path_i = str(out_dir / svg_name)

            print(f"   Item {item_idx}: {sku}  (qty {qty})")
            print(f"     Fetching customization... ", end="", flush=True)

            try:
                cust_json = fetch_customization(url)
                parsed    = parse_customization(cust_json)

                text_lines   = parsed["band_lines"]
                font         = parsed["font"]
                sb_option    = parsed["sb_option"]
                sb_lines     = parsed["sb_lines"]
                sb_font      = parsed["sb_font"]
                wants_suede_gavel = parsed["wants_suede_gavel"]
                wants_suede_sb    = parsed["wants_suede_sb"]

                if not text_lines:
                    raise ValueError("No text found in customization data")

                # Resolve band font → curves
                font_path_obj, effective_font = resolve_font_path(font)
                if font_path_obj is None:
                    alias_key = font.strip().lower()
                    if alias_key in FONT_ALIASES:
                        msg = (f"font '{font}' substituted as '{FONT_ALIASES[alias_key]}' "
                               f"but '{FONT_ALIASES[alias_key]}' TTF not found in fonts/ folder")
                    else:
                        msg = f"font '{font}' not found in fonts/ folder — add the TTF file to continue"
                    raise ValueError(msg)

                # Sound block: only proceed when engraving is requested
                want_sb = sb_option == "custom_engraved" and bool(sb_lines)
                # Warn if SKU implies a sound block but customization data has no SB content
                if _sku_has_soundblock(sku) and not want_sb and sb_option not in ("custom_engraved",):
                    print(f"\n     ⚠  SKU '{sku}' indicates a sound block product but no sound block engraving detected in customization (sb_option={sb_option!r})")
                sb_font_error = None
                if want_sb:
                    sb_font_path, effective_sb_font = resolve_font_path(sb_font)
                    if sb_font_path is None:
                        sb_font_error = sb_font
                        print(f"\n     ERROR: order {order_num} — sound block font '{sb_font}' not found in fonts/ folder — sound block skipped")
                        want_sb = False
                        errors += 1
                elif sb_option in ("no_engraving", "gavel_only", None):
                    pass   # intentional: no sound block needed
                elif sb_option == "custom_engraved" and not sb_lines:
                    print(f"\n     ⚠  Sound block engraving requested but no text found")

                # Print summary
                sb_tag = ""
                if sb_option == "custom_engraved":
                    sb_tag = f"  [sound block: {len(sb_lines)} line(s)]" if want_sb else "  [sound block: NO TEXT]"
                elif sb_option == "no_engraving":
                    sb_tag = "  [sound block: no engraving]"
                elif sb_option is not None:
                    sb_tag = f"  [sound block: {sb_option}]"
                print(f"OK  ({len(text_lines)} lines, font={effective_font}){sb_tag}")
                for i, ln in enumerate(text_lines, 1):
                    print(f"       L{i}: {ln}")
                if want_sb:
                    print(f"     Sound block text ({effective_sb_font}):")
                    for i, ln in enumerate(sb_lines, 1):
                        print(f"       SB{i}: {ln}")

                # Individual gavel band SVG
                write_individual_svg(svg_path_i, text_lines, effective_font, font_path_obj, band_template)
                svg_dest = silver_svgs if is_silver else standard_svgs
                svg_dest.append(Path(svg_path_i))
                print(f"     Band SVG saved → {svg_name}")

                # Individual sound block SVG (only when engraving requested)
                sb_svg_name = None
                if want_sb:
                    sb_svg_name = f"{safe_order}_{item_idx}_{safe_sku}_soundblock.svg"
                    sb_svg_path = str(out_dir / sb_svg_name)
                    write_soundblock_svg(sb_svg_path, sb_lines, effective_sb_font, sb_font_path)
                    svg_dest.append(Path(sb_svg_path))
                    print(f"     Sound block SVG saved → {sb_svg_name}")

                ok += 1
                order_had_success = True

                slip_items.append({
                    "sku":           sku,
                    "item_name":     item_name,
                    "qty":           qty,
                    "text_lines":    text_lines,
                    "font":          effective_font,
                    "sb_option":     sb_option,
                    "sb_lines":      sb_lines if want_sb else [],
                    "sb_font":       effective_sb_font if want_sb else (sb_font or ""),
                    "want_sb":            want_sb,
                    "wants_suede_gavel":  wants_suede_gavel,
                    "wants_suede_sb":     wants_suede_sb,
                    "sb_font_error": sb_font_error,
                })

                all_items.append({
                    "order_number":  order_num,
                    "customer":      customer,
                    "sku":           sku,
                    "qty":           qty,
                    "font":          effective_font,
                    "font_path":     font_path_obj,
                    "lines":         text_lines,
                    "template_path": band_template,
                    "is_silver":     is_silver,
                })
                if want_sb:
                    all_sb_items.append({
                        "order_number": order_num,
                        "customer":     customer,
                        "sku":          sku,
                        "qty":          qty,
                        "font":         effective_sb_font,
                        "font_path":    sb_font_path,
                        "lines":        sb_lines,
                    })
                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "qty":          qty,
                    "font":         effective_font,
                    "lines":        " | ".join(text_lines),
                    "svg_file":     svg_name,
                    "sb_option":    sb_option or "",
                    "sb_svg_file":  sb_svg_name or "",
                    "status":       f"ok (sound block skipped: font '{sb_font_error}' not found)" if sb_font_error else "ok",
                })

            except Exception as e:
                print(f"ERROR: {e}")
                errors += 1
                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "qty":          item.get("quantity", 1),
                    "font":         "",
                    "lines":        "",
                    "svg_file":     "",
                    "sb_option":    "",
                    "sb_svg_file":  "",
                    "status":       f"error: {e}",
                })

        if order_had_success:
            # Generate packing slip for this order
            slip_name = f"{safe_order}_workorder.html"
            slip_path = out_dir / slip_name
            try:
                write_packing_slip(str(slip_path), order_num, customer, ship, slip_items)
                print(f"   Work order saved → {slip_name}")
                # Convert HTML → PDF
                pdf_name = f"{safe_order}_workorder.pdf"
                pdf_path = out_dir / pdf_name
                if html_to_pdf(str(slip_path), str(pdf_path)):
                    (silver_pdf_paths if is_silver else standard_pdf_paths).append(pdf_path)
                    print(f"   Work order PDF   → {pdf_name}")
            except Exception as e:
                print(f"   ⚠  Work order generation failed: {e}")

            if is_silver:
                if order_num not in silver_order_nums:
                    silver_order_nums.append(order_num)
            else:
                if order_num not in order_nums:
                    order_nums.append(order_num)

    # ── Build bulk gavel band layouts (standard and silver) ───────────────────
    def _expand_items(items):
        out = []
        for item in items:
            for _ in range(max(item.get("qty", 1), 1)):
                out.append(item)
        return out

    std_items    = _expand_items([i for i in all_items if not i.get("is_silver")])
    silver_items = _expand_items([i for i in all_items if i.get("is_silver")])

    layout_path        = None
    silver_layout_path = None
    sb_layout_path     = None

    if std_items:
        num_pages   = max(1, (len(std_items) + COLS * ROWS - 1) // (COLS * ROWS))
        print(f"\nBuilding bulk band layout: {len(std_items)} band(s) ({len([i for i in all_items if not i.get('is_silver')])} unique) across {num_pages} page(s)...")
        layout_path = out_dir / "combined_gavel_layout.svg"
        layout_path.write_text(build_layout_svg(std_items), encoding="utf-8")
        print(f"Band layout saved → {layout_path.resolve()}")

    if silver_items:
        num_pages_s = max(1, (len(silver_items) + COLS * ROWS - 1) // (COLS * ROWS))
        print(f"\nBuilding silver band layout: {len(silver_items)} band(s) ({len([i for i in all_items if i.get('is_silver')])} unique) across {num_pages_s} page(s)...")
        silver_layout_path = out_dir / "combined_gavel_layout_silver.svg"
        silver_layout_path.write_text(build_layout_svg(silver_items), encoding="utf-8")
        print(f"Silver band layout saved → {silver_layout_path.resolve()}")

    if all_sb_items:
        num_pages_sb = max(1, (len(all_sb_items) + SB_COLS * SB_ROWS - 1) // (SB_COLS * SB_ROWS))
        print(f"\nBuilding sound block layout: {len(all_sb_items)} block(s) across {num_pages_sb} page(s)...")
        sb_layout_path = out_dir / "combined_soundblock_layout.svg"
        sb_layout_path.write_text(build_soundblock_layout_svg(all_sb_items), encoding="utf-8")
        print(f"Sound block layout saved → {sb_layout_path.resolve()}")

    # ── Merge work order PDFs into batch files ────────────────────────────────
    batch_pdf_path        = None
    silver_batch_pdf_path = None

    if standard_pdf_paths:
        batch_pdf_path = out_dir / "workorders_batch.pdf"
        if merge_pdfs([str(p) for p in standard_pdf_paths], str(batch_pdf_path)):
            print(f"\nBatch work order PDF ({len(standard_pdf_paths)} order(s)) → {batch_pdf_path.resolve()}")
        else:
            batch_pdf_path = None

    if silver_pdf_paths:
        silver_batch_pdf_path = out_dir / "workorders_batch_silver.pdf"
        if merge_pdfs([str(p) for p in silver_pdf_paths], str(silver_batch_pdf_path)):
            print(f"Silver batch work order PDF ({len(silver_pdf_paths)} order(s)) → {silver_batch_pdf_path.resolve()}")
        else:
            silver_batch_pdf_path = None

    # ── Write summary CSV ──────────────────────────────────────────────────────
    csv_path = out_dir / "summary.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["order_number", "customer", "sku", "qty", "font",
                        "lines", "svg_file", "sb_option", "sb_svg_file", "status"],
        )
        writer.writeheader()
        writer.writerows(summary)

    # ── Post Trello cards ──────────────────────────────────────────────────────
    if target_card_id:
        # Upload only the batch PDF(s) to the existing card — no new card created
        pdfs_to_upload = []
        if batch_pdf_path and batch_pdf_path.exists():
            pdfs_to_upload.append(batch_pdf_path)
        if silver_batch_pdf_path and silver_batch_pdf_path.exists():
            pdfs_to_upload.append(silver_batch_pdf_path)
        if pdfs_to_upload:
            print(f"\nUploading packing slip PDF(s) to existing Trello card {target_card_id}...")
            try:
                for i, f in enumerate(pdfs_to_upload, 1):
                    print(f"  [{i}/{len(pdfs_to_upload)}] Uploading {f.name}...", end="", flush=True)
                    trello_attach_file(target_card_id, str(f), mime_type="application/pdf")
                    print(" ✓")
            except Exception as e:
                print(f"  Trello error: {e}")
        else:
            print("\n  No batch PDF generated — nothing uploaded to Trello.")
    else:
        if order_nums:
            print(f"\nPosting Trello card ({len(order_nums)} order numbers)...")
            try:
                card_url, card_id = trello_create_gavel_card(order_nums, rerun=bool(force_orders))
                if card_url:
                    print(f"  Card created → {card_url}")
                if card_id:
                    upload_list = ([layout_path] if layout_path else []) + standard_svgs
                    if sb_layout_path and sb_layout_path.exists():
                        upload_list.append(sb_layout_path)
                    if batch_pdf_path and batch_pdf_path.exists():
                        upload_list.append(batch_pdf_path)
                    for i, f in enumerate(upload_list, 1):
                        print(f"  [{i}/{len(upload_list)}] Uploading {f.name}...", end="", flush=True)
                        if f.suffix.lower() == ".pdf":
                            mime = "application/pdf"
                        elif f.suffix.lower() == ".html":
                            mime = "text/html"
                        else:
                            mime = "image/svg+xml"
                        trello_attach_file(card_id, str(f), mime_type=mime)
                        print(" ✓")
            except Exception as e:
                print(f"  Trello error: {e}")

        if silver_order_nums:
            print(f"\nPosting Trello card (Silver Band, {len(silver_order_nums)} order numbers)...")
            try:
                card_url, card_id = trello_create_gavel_card(silver_order_nums, variant="Silver Band", rerun=bool(force_orders))
                if card_url:
                    print(f"  Card created → {card_url}")
                if card_id:
                    upload_list = ([silver_layout_path] if silver_layout_path else []) + silver_svgs
                    if silver_batch_pdf_path and silver_batch_pdf_path.exists():
                        upload_list.append(silver_batch_pdf_path)
                    for i, f in enumerate(upload_list, 1):
                        print(f"  [{i}/{len(upload_list)}] Uploading {f.name}...", end="", flush=True)
                        if f.suffix.lower() == ".pdf":
                            mime = "application/pdf"
                        elif f.suffix.lower() == ".html":
                            mime = "text/html"
                        else:
                            mime = "image/svg+xml"
                        trello_attach_file(card_id, str(f), mime_type=mime)
                        print(" ✓")
            except Exception as e:
                print(f"  Trello error: {e}")

    print(f"\n{'='*60}")
    print(f"Done.  Individual SVGs: {ok}  Errors: {errors}")
    if layout_path:
        print(f"Band layout    : {layout_path.resolve()}")
    if silver_layout_path:
        print(f"Silver layout  : {silver_layout_path.resolve()}")
    if sb_layout_path:
        print(f"SB layout      : {sb_layout_path.resolve()}  ({len(all_sb_items)} block(s))")
    if batch_pdf_path:
        print(f"Batch PDF      : {batch_pdf_path.resolve()}  ({len(standard_pdf_paths)} order(s))")
    if silver_batch_pdf_path:
        print(f"Silver PDF     : {silver_batch_pdf_path.resolve()}  ({len(silver_pdf_paths)} order(s))")
    print(f"Summary CSV    : {csv_path.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
