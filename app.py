import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import io
import os
import re
from PIL import ImageFont, Image, ImageFilter

st.set_page_config(page_title="AMC NTEP - Slide Auto-Designer", layout="wide")
st.title("AMC NTEP — Public Health Actions Deck Builder")
st.caption(
    "Upload your Word doc (પ્રસ્તાવના → હેતુ → ઉદ્દેશ્યો → Module 1..7, sub-modules "
    "1.1, 1.2 ...). This builds a full MNC-style deck: title, auto-numbered index, "
    "intro pages, a modules-at-a-glance overview, module dividers, and fully designed "
    "sub-module slides with equipment photos — measuring your real text so nothing "
    "overflows the cards."
)

# ----------------------------------------------------------------------
# THEME
# ----------------------------------------------------------------------
NAVY = RGBColor(0x14, 0x2C, 0x5C)
NAVY_LIGHT = RGBColor(0x28, 0x49, 0x86)
GREEN = RGBColor(0x1F, 0xA6, 0x8C)
BLUE_ICON = RGBColor(0x1F, 0x5C, 0xC9)
AMBER_ICON = RGBColor(0xC9, 0x7B, 0x1F)
RED_ICON = RGBColor(0xC0, 0x3B, 0x2E)
BG_LIGHT = RGBColor(0xEE, 0xF5, 0xFB)
CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x22, 0x2B, 0x3A)
TEXT_GRAY = RGBColor(0x46, 0x50, 0x60)
HIGHLIGHT_BG = RGBColor(0xC9, 0xF3, 0xEE)
HIGHLIGHT_TEXT = RGBColor(0x0B, 0x6E, 0x5D)
WARN_BG = RGBColor(0xFB, 0xE4, 0xE1)
WARN_TEXT = RGBColor(0x8A, 0x2A, 0x1E)
# Extra accent colors purely for visual variety (module badges, dividers, TOC
# rows) so the deck doesn't read as one flat navy block after 40 slides.
PURPLE_ICON = RGBColor(0x8E, 0x44, 0xC2)
TEAL_ICON = RGBColor(0x17, 0x8C, 0xA6)
MODULE_PALETTE = [GREEN, BLUE_ICON, AMBER_ICON, PURPLE_ICON, RED_ICON, TEAL_ICON]

# India-flag palette, used on the closing "Thank You" slide.
SAFFRON = RGBColor(0xFF, 0x99, 0x33)
INDIA_GREEN = RGBColor(0x13, 0x88, 0x08)
CHAKRA_NAVY = RGBColor(0x00, 0x00, 0x80)


def tint(rgb_color, amount=0.82):
    """Blend a theme color toward white to get a soft halo/background tint of
    the same hue — keeps things colorful without making anything darker."""
    hexs = str(rgb_color)
    r, g, b = int(hexs[0:2], 16), int(hexs[2:4], 16), int(hexs[4:6], 16)
    r = int(r + (255 - r) * amount)
    g = int(g + (255 - g) * amount)
    b = int(b + (255 - b) * amount)
    return RGBColor(r, g, b)


FONT_GUJARATI = "Shruti"
FONT_ENGLISH = "Calibri"
_GUJARATI_CHAR_RE = re.compile(r"[\u0A80-\u0AFF]")


def pick_font(text):
    """Shruti for any text containing Gujarati script, Calibri otherwise."""
    return FONT_GUJARATI if _GUJARATI_CHAR_RE.search(text or "") else FONT_ENGLISH
# Font files must sit next to this script (or set full paths).
ASSET_DIR = os.path.dirname(__file__)
FONT_TTF_REGULAR = os.path.join(ASSET_DIR, "NotoSansGujarati-Regular.ttf")
FONT_TTF_BOLD = os.path.join(ASSET_DIR, "NotoSansGujarati-Bold.ttf")

# Default title-page artwork. These are expected to sit next to app.py in the
# repo (same convention as the font files above). They can be overridden from
# the Streamlit UI without touching this script.
DEFAULT_HERITAGE_BG = os.path.join(ASSET_DIR, "banner_sidi-saiyyad-jali_902.jpg")
DEFAULT_RIVERFRONT = os.path.join(ASSET_DIR, "riverfront.jpg")

MEASURE_DPI = 96
# PIL has no Indic shaping engine, so it under-measures Gujarati conjuncts/matras.
# This factor widens every measured line so wrapping matches real PowerPoint
# rendering instead of overflowing the card. Tune upward if you still see overlap
# with your real document; 1.14 was enough margin against the real NTEP text.
GUJ_WIDTH_CORRECTION = 1.14

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN_X = Inches(0.55)
HEADER_H = Inches(1.25)
BODY_TOP = HEADER_H + Inches(0.32)
BODY_BOTTOM_MARGIN = Inches(0.4)
COL_GAP = Inches(0.28)
CARD_GAP_V = Inches(0.2)
CARD_PAD = Inches(0.16)

# Fixed height (in inches) reserved for the small in-slide equipment-photo card
# when a sub-module has a matched photo. It flows through the same column
# layout as the text cards, so it never overlaps anything.
IMAGE_CARD_H_IN = 2.15

# A single uploaded equipment photo will never be placed into more than this
# many sub-modules across the whole deck, even if its keyword genuinely
# appears in more places in the text. Overridable from the Streamlit UI.
DEFAULT_EQUIP_MAX_REPEATS = 3

# ----------------------------------------------------------------------
# 0.5 IMAGE / PHOTO HELPERS
# ----------------------------------------------------------------------
def prepare_cover_image(source, target_w_px, target_h_px, blur_radius=0):
    """Crop `source` (a path string or file-like) to COVER a target_w x
    target_h box with no distortion (like CSS background-size: cover),
    optionally Gaussian-blur it, and return a JPEG BytesIO ready for
    slide.shapes.add_picture()."""
    if hasattr(source, "seek"):
        source.seek(0)
    img = Image.open(source).convert("RGB")
    src_w, src_h = img.size
    target_ratio = target_w_px / target_h_px
    src_ratio = src_w / src_h
    if src_ratio > target_ratio:
        new_w = max(1, int(src_h * target_ratio))
        offset = max(0, (src_w - new_w) // 2)
        img = img.crop((offset, 0, offset + new_w, src_h))
    else:
        new_h = max(1, int(src_w / target_ratio))
        offset = max(0, (src_h - new_h) // 2)
        img = img.crop((0, offset, src_w, offset + new_h))
    img = img.resize((target_w_px, target_h_px), Image.LANCZOS)
    if blur_radius > 0:
        img = img.filter(ImageFilter.GaussianBlur(blur_radius))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    return buf


def resolve_front_image(uploaded_file, default_path):
    """Prefer a user-uploaded override; else fall back to the bundled
    repo asset if it exists next to app.py; else None (caller degrades
    gracefully to a solid color)."""
    if uploaded_file is not None:
        return uploaded_file
    if default_path and os.path.exists(default_path):
        return default_path
    return None


def set_shape_alpha(shape, alpha_pct):
    """Make a solid-filled shape translucent. alpha_pct: 0-100 (100 = fully
    opaque, 0 = fully transparent). Used to lay a dark veil over a photo
    background so white text stays legible."""
    spPr = shape._element.spPr
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is None:
        return
    srgbClr = solidFill.find(qn('a:srgbClr'))
    if srgbClr is None:
        return
    alpha_el = srgbClr.makeelement(qn('a:alpha'), {'val': str(int(alpha_pct * 1000))})
    srgbClr.append(alpha_el)


# ----------------------------------------------------------------------
# FIELD DEFINITIONS
# ----------------------------------------------------------------------
# Matching is done on the ENGLISH word in parentheses (e.g. "... (Objective):"),
# not on the Gujarati label itself. This is deliberate: Gujarati spelling in the
# source doc can drift ("શું કરવું?" vs "શું કરવું ?" vs a missing "?"), but the
# English parenthetical is far more consistent. If your doc uses a different
# English word for a field, just add it to the matching list below.
FIELD_META = {
    "OBJECTIVE":    {"gj": "ઉદ્દેશ્ય",              "en": "Objective",            "icon": "target", "color": GREEN,      "bullet": False, "group": "process"},
    "TRIGGER":      {"gj": "ટ્રિગર",                "en": "Trigger",              "icon": "bolt",   "color": AMBER_ICON, "bullet": False, "group": "process"},
    "WHAT_TO_DO":   {"gj": "શું કરવું?",            "en": "What to do",           "icon": "gear",   "color": BLUE_ICON,  "bullet": True,  "group": "process"},
    "WHY":          {"gj": "શા માટે",               "en": "Why",                  "icon": "info",   "color": BLUE_ICON,  "bullet": False, "group": "process"},
    "WHOM":         {"gj": "કોના માટે",             "en": "Whom",                 "icon": "people", "color": BLUE_ICON,  "bullet": False, "group": "process"},
    "RESPONSIBLE":  {"gj": "જવાબદાર વ્યક્તિ",       "en": "Responsible Person",   "icon": "people", "color": BLUE_ICON,  "bullet": False, "group": "process"},
    "TIMELINE":     {"gj": "સમયમર્યાદા",            "en": "Timeline",             "icon": "clock",  "color": GREEN,      "bullet": False, "group": "process"},
    "DOCUMENTATION":{"gj": "દસ્તાવેજીકરણ",         "en": "Documentation",        "icon": "doc",    "color": BLUE_ICON,  "bullet": True,  "group": "quality"},
    "MONITORING":   {"gj": "મોનિટરિંગ સૂચકાંકો",    "en": "Monitoring Indicators","icon": "chart",  "color": HIGHLIGHT_TEXT, "bullet": True, "group": "quality", "highlight": True},
    "CHECKLIST":    {"gj": "સુપરવાઇઝર ચેકલિસ્ટ",    "en": "Supervisor Checklist", "icon": "check",  "color": HIGHLIGHT_TEXT, "bullet": True, "group": "quality", "highlight": True},
    "IF_NOT_DONE":  {"gj": "જો ન કરવામાં આવે તો",  "en": "If Not Done",          "icon": "warn",   "color": WARN_TEXT,  "bullet": True,  "group": "quality", "warn": True},
}
# Order within each grouped slide
PROCESS_ORDER = ["OBJECTIVE", "TRIGGER", "WHAT_TO_DO", "WHY", "WHOM", "RESPONSIBLE", "TIMELINE"]
QUALITY_ORDER = ["DOCUMENTATION", "MONITORING", "CHECKLIST", "IF_NOT_DONE"]
FIELD_ORDER = PROCESS_ORDER + QUALITY_ORDER


def field_label(key):
    m = FIELD_META[key]
    return f"{m['gj']} ({m['en']})"


# ----------------------------------------------------------------------
# 1. WORD PARSER
# ----------------------------------------------------------------------
# Build one regex per field that matches a line starting with (optional Gujarati
# label +) the English word in parentheses, colon optional. This is robust to
# Gujarati spelling drift because we anchor on the English term.
def _field_pattern(en_word):
    return re.compile(r"^.*?\(\s*" + re.escape(en_word) + r"\s*\??\s*\)\s*[:：]?", re.IGNORECASE)


FIELD_PATTERNS = [(key, _field_pattern(meta["en"])) for key, meta in FIELD_META.items()]

MODULE_RE = re.compile(r"^(Module|મોડ્યુલ)\s*(\d+)\s*[:：]?\s*(.*)$", re.IGNORECASE)
SUBMODULE_RE = re.compile(r"^(\d+)\.(\d+)\s+(.*)$")

PREFACE_HEADINGS = [
    ("PREFACE", re.compile(r"^પ્રસ્તાવના\s*$")),
    ("PURPOSE", re.compile(r"^હેતુ\s*$")),
    ("OBJECTIVES", re.compile(r"^માર્ગદર્શિકાના\s*મુખ્ય\s*ઉદ્દેશ્યો\s*$")),
]


def strip_label(text, pattern):
    return pattern.sub("", text).strip()


def parse_word(docx_file):
    """Returns (preface_dict, modules_dict).
    preface_dict: {"PREFACE": [...], "PURPOSE": [...], "OBJECTIVES": [...]}
    modules_dict: {mod_num: {"title": str, "subs": {sub_num: {"title":..., FIELD_KEY: [...]}}}}
    """
    doc = docx.Document(docx_file)
    preface = {"PREFACE": [], "PURPOSE": [], "OBJECTIVES": []}
    modules = {}
    cur_mod, cur_sub, cur_field = None, None, None
    cur_preface_section = None
    in_body = False  # becomes True once we hit Module 1
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Preface section headings (only before first Module)
        if not in_body:
            matched_heading = False
            for key, pat in PREFACE_HEADINGS:
                if pat.match(text):
                    cur_preface_section = key
                    matched_heading = True
                    break
            if matched_heading:
                continue
        m = MODULE_RE.match(text)
        if m:
            in_body = True
            cur_preface_section = None
            cur_mod = m.group(2)
            modules.setdefault(cur_mod, {"title": m.group(3).strip(), "subs": {}})
            cur_sub, cur_field = None, None
            continue
        s = SUBMODULE_RE.match(text)
        if s:
            in_body = True
            main_num = s.group(1)
            cur_mod = main_num if cur_mod is None else cur_mod
            modules.setdefault(cur_mod, {"title": "", "subs": {}})
            cur_sub = f"{s.group(1)}.{s.group(2)}"
            modules[cur_mod]["subs"][cur_sub] = {"title": s.group(3).strip()}
            for key in FIELD_META:
                modules[cur_mod]["subs"][cur_sub][key] = []
            cur_field = None
            continue
        if not in_body:
            if cur_preface_section:
                preface[cur_preface_section].append(text)
            continue
        if not cur_sub:
            continue
        matched = False
        for key, pat in FIELD_PATTERNS:
            if pat.match(text):
                value = strip_label(text, pat)
                cur_field = key
                # Guard against a source-doc quirk where the last process step and
                # the next field's label land on the same paragraph with no break
                # (e.g. "...3. Nikshay માં નોંધવું. જવાબદાર વ્યક્તિ (Responsible Person): STS").
                # We only take what follows THIS field's own label match, so the
                # preceding fragment (already appended to the previous field) is safe.
                modules[cur_mod]["subs"][cur_sub][key] = [value] if value else []
                matched = True
                break
        if matched:
            continue
        if cur_field:
            modules[cur_mod]["subs"][cur_sub][cur_field].append(text)
    return preface, modules


# ----------------------------------------------------------------------
# 2. TEXT MEASUREMENT  (drives dynamic sizing + pagination)
# ----------------------------------------------------------------------
_FONT_CACHE = {}


def _pil_font(size_pt, bold=False):
    key = (round(size_pt, 1), bold)
    if key not in _FONT_CACHE:
        path = FONT_TTF_BOLD if bold else FONT_TTF_REGULAR
        try:
            _FONT_CACHE[key] = ImageFont.truetype(path, max(1, int(size_pt / 72 * MEASURE_DPI)))
        except OSError:
            _FONT_CACHE[key] = ImageFont.load_default()
    return _FONT_CACHE[key]


def _text_width_px(text, font):
    try:
        length = font.getlength(text)
    except AttributeError:
        length = font.getsize(text)[0]
    return length * GUJ_WIDTH_CORRECTION


def wrap_text(text, size_pt, max_width_in, bold=False):
    font = _pil_font(size_pt, bold=bold)
    max_width_px = max_width_in * MEASURE_DPI
    words = text.split(" ")
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if _text_width_px(trial, font) <= max_width_px or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def item_height_in(item, size_pt, max_width_in, bullet, line_spacing=1.18, item_gap_in=0.05):
    prefix_w = 0.22 if bullet else 0.0
    lines = wrap_text(item, size_pt, max_width_in - prefix_w)
    return len(lines) * size_pt * line_spacing / 72 + item_gap_in


def chunk_items(items, max_width_in, avail_h_in, bullet):
    """Pick the largest font (12..9pt) that yields the fewest chunks, then
    return that font size and the list of chunks (each chunk = list[str])."""
    if not items:
        return 12, []
    best = None
    for size in (12, 11, 10, 9.5, 9):
        chunks, cur, cur_h = [], [], 0.0
        for it in items:
            h = item_height_in(it, size, max_width_in, bullet)
            if cur and cur_h + h > avail_h_in:
                chunks.append(cur)
                cur, cur_h = [it], h
            else:
                cur.append(it)
                cur_h += h
        if cur:
            chunks.append(cur)
        if best is None or len(chunks) < best[0]:
            best = (len(chunks), size, chunks)
        if len(chunks) == 1:
            break
    return best[1], best[2]


def fit_one_line_size(text, max_width_in, start_size=23, min_size=15, bold=True):
    """Shrink font until text fits on a single line (used for slide headers so a
    long sub-module title never wraps to 2 lines and spills out of the navy band)."""
    size = start_size
    while size > min_size:
        font = _pil_font(size, bold=bold)
        if _text_width_px(text, font) <= max_width_in * MEASURE_DPI:
            return size
        size -= 1
    return min_size


# ----------------------------------------------------------------------
# 3. DRAWING PRIMITIVES
# ----------------------------------------------------------------------
def add_rect(slide, x, y, w, h, fill, radius=None, shadow=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if radius is not None:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    if shadow:
        el = shp._element.spPr
        from pptx.oxml.ns import qn
        eff = el.makeelement(qn("a:effectLst"), {})
        sh = eff.makeelement(qn("a:outerShdw"), {"blurRad": "95000", "dist": "25000", "dir": "5400000", "rotWithShape": "0"})
        clr = sh.makeelement(qn("a:srgbClr"), {"val": "0F1F3D"})
        clr.append(clr.makeelement(qn("a:alpha"), {"val": "28000"}))
        sh.append(clr)
        eff.append(sh)
        el.append(eff)
    return shp


def add_oval(slide, x, y, d, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.18
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = pick_font(line)
        r.font.color.rgb = color
    return tb


def add_items(slide, x, y, w, h, items, size, color, bullet):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.18
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = (f"\u2022  {item}" if bullet else item)
        r.font.size = Pt(size)
        r.font.name = pick_font(item)
        r.font.color.rgb = color
    return tb


# ---- icon glyphs (flat vector shapes, no badge circle) ----
def icon_target(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    m = int(d * 0.62)
    add_oval(slide, x + int((d - m) / 2), y + int((d - m) / 2), m, WHITE)
    s = int(d * 0.30)
    add_oval(slide, x + int((d - s) / 2), y + int((d - s) / 2), s, color)


def icon_gear(slide, x, y, d, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.GEAR_9, x, y, d, d)
    shp.fill.solid(); shp.fill.fore_color.rgb = color; shp.line.fill.background(); shp.shadow.inherit = False


def icon_people(slide, x, y, d, color):
    head_d = int(d * 0.40)
    add_oval(slide, x + int((d - head_d) / 2), y, head_d, color)
    body_w, body_h = int(d * 0.82), int(d * 0.5)
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + int((d - body_w) / 2),
                                   y + int(head_d * 0.78), body_w, body_h)
    try:
        body.adjustments[0] = 0.5
    except Exception:
        pass
    body.fill.solid(); body.fill.fore_color.rgb = color; body.line.fill.background(); body.shadow.inherit = False


def icon_clock(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    cx, cy = x + d // 2, y + d // 2
    hand1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx - int(d * 0.035), cy - int(d * 0.33), int(d * 0.07), int(d * 0.33))
    hand1.fill.solid(); hand1.fill.fore_color.rgb = WHITE; hand1.line.fill.background(); hand1.shadow.inherit = False
    hand2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx - int(d * 0.035), cy - int(d * 0.24), int(d * 0.07), int(d * 0.24))
    hand2.fill.solid(); hand2.fill.fore_color.rgb = WHITE; hand2.line.fill.background(); hand2.shadow.inherit = False
    hand2.rotation = 90


def icon_chart(slide, x, y, d, color):
    bar_w = int(d * 0.20); gap = int(d * 0.12)
    heights = [int(d * 0.42), int(d * 0.68), int(d * 0.95)]
    total_w = 3 * bar_w + 2 * gap
    start_x = x + int((d - total_w) / 2)
    for i, h in enumerate(heights):
        bx = start_x + i * (bar_w + gap)
        by = y + (d - h)
        add_rect(slide, bx, by, bar_w, h, color)


def icon_bolt(slide, x, y, d, color):
    # simple lightning bolt via a freeform-ish stack of two triangles
    t1 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x + int(d * 0.15), y, int(d * 0.6), int(d * 0.62))
    t1.rotation = 200
    t1.fill.solid(); t1.fill.fore_color.rgb = color; t1.line.fill.background(); t1.shadow.inherit = False
    t2 = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x + int(d * 0.22), y + int(d * 0.4), int(d * 0.6), int(d * 0.62))
    t2.rotation = 20
    t2.fill.solid(); t2.fill.fore_color.rgb = color; t2.line.fill.background(); t2.shadow.inherit = False


def icon_info(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    add_rect(slide, x + int(d * 0.44), y + int(d * 0.42), int(d * 0.12), int(d * 0.36), WHITE)
    add_oval(slide, x + int(d * 0.41), y + int(d * 0.18), int(d * 0.18), WHITE)


def icon_doc(slide, x, y, d, color):
    add_rect(slide, x + int(d * 0.15), y, int(d * 0.7), d, color, radius=0.12)
    for i in range(3):
        add_rect(slide, x + int(d * 0.28), y + int(d * (0.28 + i * 0.2)), int(d * 0.44), int(d * 0.06), WHITE)


def icon_check(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    from pptx.oxml.ns import qn
    # simple check via two thin rotated rectangles
    r1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + int(d * 0.26), y + int(d * 0.48), int(d * 0.28), int(d * 0.09))
    r1.rotation = 45
    r1.fill.solid(); r1.fill.fore_color.rgb = WHITE; r1.line.fill.background(); r1.shadow.inherit = False
    r2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + int(d * 0.40), y + int(d * 0.30), int(d * 0.42), int(d * 0.09))
    r2.rotation = -45
    r2.fill.solid(); r2.fill.fore_color.rgb = WHITE; r2.line.fill.background(); r2.shadow.inherit = False


def icon_warn(slide, x, y, d, color):
    tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x, y, d, d)
    tri.fill.solid(); tri.fill.fore_color.rgb = color; tri.line.fill.background(); tri.shadow.inherit = False
    add_rect(slide, x + int(d * 0.46), y + int(d * 0.38), int(d * 0.08), int(d * 0.28), WHITE)
    add_oval(slide, x + int(d * 0.45), y + int(d * 0.74), int(d * 0.10), WHITE)


def draw_ashoka_chakra(slide, x, y, d, color):
    """Stylised 24-spoke Ashoka Chakra used on the India-themed closing
    slide. x, y, d define the bounding box (top-left corner + diameter)."""
    cx = x + d // 2
    cy = y + d // 2
    ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    ring.fill.background()
    ring.line.color.rgb = color
    ring.line.width = Pt(2.4)
    ring.shadow.inherit = False
    thin = max(int(d * 0.028), 12000)
    # 12 full-diameter rectangles rotated 15° apart = 24 spokes, and rotating
    # a shape whose bounding box is centered on (cx, cy) pivots it correctly
    # around the chakra's own center.
    for i in range(12):
        spoke = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx - d // 2, cy - thin // 2, d, thin)
        spoke.rotation = i * 15
        spoke.fill.solid()
        spoke.fill.fore_color.rgb = color
        spoke.line.fill.background()
        spoke.shadow.inherit = False
    hub_d = int(d * 0.16)
    add_oval(slide, cx - hub_d // 2, cy - hub_d // 2, hub_d, color)


ICON_FN = {
    "target": icon_target, "gear": icon_gear, "people": icon_people, "clock": icon_clock,
    "chart": icon_chart, "bolt": icon_bolt, "info": icon_info, "doc": icon_doc,
    "check": icon_check, "warn": icon_warn,
}


# ----------------------------------------------------------------------
# 4. CARD BUILDER
# ----------------------------------------------------------------------
def draw_card(slide, x, y, w, h, field_key, size_pt, items, label_override=None):
    meta = FIELD_META[field_key]
    if meta.get("warn"):
        bg = WARN_BG
        text_color = WARN_TEXT
    elif meta.get("highlight"):
        bg = HIGHLIGHT_BG
        text_color = HIGHLIGHT_TEXT
    else:
        bg = CARD_WHITE
        text_color = TEXT_GRAY
    add_rect(slide, x, y, w, h, bg, radius=0.06, shadow=True)
    if not meta.get("warn") and not meta.get("highlight"):
        # thin colored accent stripe along the top edge, keyed to the field's
        # own color, so each card reads as belonging to a "type" at a glance
        add_rect(slide, x + Inches(0.1), y, w - Inches(0.2), Inches(0.05), meta["color"])
    icon_d = Inches(0.34)
    pad = CARD_PAD
    icon_cx = x + pad + icon_d // 2
    icon_cy = y + pad + icon_d // 2
    halo_d = Inches(0.5)
    add_oval(slide, icon_cx - halo_d // 2, icon_cy - halo_d // 2, halo_d, tint(meta["color"], 0.78))
    ICON_FN[meta["icon"]](slide, x + pad, y + pad, icon_d, meta["color"])
    label_x = x + pad + icon_d + Inches(0.12)
    label_w = w - pad - icon_d - Inches(0.12) - pad
    add_text(slide, label_x, y + pad - Inches(0.02), label_w, icon_d,
             label_override or field_label(field_key), 12, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    body_y = y + pad + icon_d + Inches(0.10)
    body_w = w - 2 * pad
    body_h = h - (body_y - y) - pad
    add_items(slide, x + pad, body_y, body_w, body_h, items, size_pt, text_color, meta["bullet"])


def draw_image_card(slide, x, y, w, h, keyword_file):
    """Small in-slide 'Equipment' photo card. Sits inside the normal card grid
    (same column-flow as the text cards) so it can never overlap other content."""
    keyword, file = keyword_file
    add_rect(slide, x, y, w, h, CARD_WHITE, radius=0.06, shadow=True)
    add_rect(slide, x + Inches(0.1), y, w - Inches(0.2), Inches(0.05), TEAL_ICON)
    pad = CARD_PAD
    label_h = Inches(0.28)
    add_text(slide, x + pad, y + pad - Inches(0.02), w - 2 * pad, label_h,
             "સાધનસામગ્રી (Equipment)", 11.5, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    img_area_x = x + pad
    img_area_y = y + pad + label_h + Inches(0.06)
    img_area_w = w - 2 * pad
    img_area_h = h - (img_area_y - y) - pad
    try:
        file.seek(0)
        pic = slide.shapes.add_picture(file, img_area_x, img_area_y, height=img_area_h)
        if pic.width > img_area_w:
            file.seek(0)
            slide.shapes._spTree.remove(pic._element)
            pic = slide.shapes.add_picture(file, img_area_x, img_area_y, width=img_area_w)
        pic.left = int(img_area_x + (img_area_w - pic.width) / 2)
        pic.top = int(img_area_y + (img_area_h - pic.height) / 2)
    except Exception:
        add_text(slide, img_area_x, img_area_y, img_area_w, img_area_h, f"[{keyword}]", 12, TEXT_GRAY,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def field_card_height_in(field_key, items, size_pt, content_w_in):
    if field_key == "IMAGE":
        return IMAGE_CARD_H_IN
    meta = FIELD_META[field_key]
    total = sum(item_height_in(it, size_pt, content_w_in, meta["bullet"]) for it in items)
    header_h_in = 0.34 + 0.10
    return header_h_in + total + (CARD_PAD / 914400) * 2


# ----------------------------------------------------------------------
# 5. LAYOUT PLANNER  (shared by dry-run TOC pass and real drawing pass)
# ----------------------------------------------------------------------
def plan_field_group_slides(items_by_field, field_keys, content_w_in, body_h_in, image_piece=None):
    """Given a subset of fields (e.g. the 'process' group or 'quality' group),
    return a list of slide layouts. Each layout is a list of
    (col, field_key, font_size, chunk_items, is_continuation, h_in).
    image_piece, if given, is (keyword, file) for a small equipment-photo card
    that gets folded into the same column-flow as the text cards."""
    pieces = []
    if image_piece is not None:
        pieces.append(("IMAGE", None, image_piece, False))
    for key in field_keys:
        items = [x for x in items_by_field.get(key, []) if x.strip()]
        if not items:
            continue
        avail = body_h_in - 0.34 - 0.10 - (CARD_PAD / 914400) * 2
        font_size, chunks = chunk_items(items, content_w_in, avail, FIELD_META[key]["bullet"])
        for idx, chunk in enumerate(chunks):
            pieces.append((key, font_size, chunk, idx > 0))
    if not pieces:
        return []
    slides_layout = []
    col_used = [0.0, 0.0]
    cur_slide_items = []
    for key, font_size, chunk, is_cont in pieces:
        h_in = field_card_height_in(key, chunk, font_size, content_w_in)
        target_col = 0 if col_used[0] <= col_used[1] else 1
        if col_used[target_col] + h_in > body_h_in and cur_slide_items:
            slides_layout.append(cur_slide_items)
            cur_slide_items = []
            col_used = [0.0, 0.0]
            target_col = 0
        cur_slide_items.append((target_col, key, font_size, chunk, is_cont, h_in))
        col_used[target_col] += h_in + (CARD_GAP_V / 914400)
    if cur_slide_items:
        slides_layout.append(cur_slide_items)
    return slides_layout


def plan_submodule(sub, image_match=None):
    """Returns list of (group_label, layout) pairs -> total slide count is len(list).
    image_match, if given, is (keyword, file) for the sub-module's equipment photo —
    it gets embedded as a small card among the process-group slides."""
    col_w_emu = (SLIDE_W - 2 * MARGIN_X - COL_GAP) // 2
    content_w_in = col_w_emu / 914400 - (CARD_PAD / 914400) * 2
    body_h_in = (SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN) / 914400
    result = []
    process_layout = plan_field_group_slides(sub, PROCESS_ORDER, content_w_in, body_h_in, image_piece=image_match)
    quality_layout = plan_field_group_slides(sub, QUALITY_ORDER, content_w_in, body_h_in)
    for layout in process_layout:
        result.append(("process", layout))
    for layout in quality_layout:
        result.append(("quality", layout))
    if not result:
        result = [("process", [(0, "OBJECTIVE", 12, ["\u2014"], False, 1.0)])]
    return result


# ----------------------------------------------------------------------
# 6. SLIDE DRAWING
# ----------------------------------------------------------------------
def module_accent_color(module_num):
    try:
        return MODULE_PALETTE[(int(module_num) - 1) % len(MODULE_PALETTE)]
    except (ValueError, TypeError):
        return GREEN


def draw_header(slide, module_num, module_title, sub_title, group_label, part_no, part_total, logos):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    add_rect(slide, 0, HEADER_H, SLIDE_W, Inches(0.06), module_accent_color(module_num))
    left_logo, right_logo = logos
    text_left = Inches(0.55)
    text_right_pad = Inches(0.55)
    if left_logo is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(left_logo, Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    if right_logo is not None:
        d = Inches(0.85)
        rx = SLIDE_W - Inches(0.25) - d
        add_oval(slide, rx, (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(right_logo, rx + Inches(0.07), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_right_pad = Inches(1.35)
    header_text_w = SLIDE_W - text_left - text_right_pad
    kicker = f"Module {module_num} \u2022 {module_title}".strip(" \u2022")
    if group_label == "quality":
        kicker += "  \u2022  ગુણવત્તા અને દેખરેખ"
    add_text(slide, text_left, Inches(0.14), header_text_w, Inches(0.3), kicker, 13, RGBColor(0xB9, 0xD3, 0xF2))
    title = sub_title
    if part_total > 1:
        title += f"   ({part_no}/{part_total})"
    title_size = fit_one_line_size(title, header_text_w / 914400, start_size=23, min_size=15)
    add_text(slide, text_left, Inches(0.48), header_text_w, Inches(0.68), title, title_size, WHITE,
             bold=True, anchor=MSO_ANCHOR.MIDDLE)


def draw_field_slide(prs, module_num, module_title, sub_title, group_label, part_no, part_total,
                      layout, logos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    draw_header(slide, module_num, module_title, sub_title, group_label, part_no, part_total, logos)
    col_w_emu = (SLIDE_W - 2 * MARGIN_X - COL_GAP) // 2
    full_w = SLIDE_W - 2 * MARGIN_X
    single_col = not any(item[0] == 1 for item in layout)
    col_y = [BODY_TOP, BODY_TOP]
    col_x = [MARGIN_X, MARGIN_X + col_w_emu + COL_GAP]
    for target_col, key, font_size, chunk, is_cont, h_in in layout:
        w = full_w if single_col else col_w_emu
        x = col_x[0] if single_col else col_x[target_col]
        y = col_y[target_col]
        h = Emu(int(h_in * 914400))
        if key == "IMAGE":
            draw_image_card(slide, x, y, w, h, chunk)
        else:
            label_override = field_label(key) + ("  (\u091a\u093e\u0932\u0941)" if is_cont else "")
            draw_card(slide, x, y, w, h, key, font_size, chunk, label_override=label_override)
        col_y[target_col] = y + h + CARD_GAP_V
    return slide


def draw_equipment_slide(prs, module_num, module_title, sub_title, keyword, image_file, logos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    text_left = Inches(0.55)
    if logos[0] is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    add_text(slide, text_left, Inches(0.14), Inches(10), Inches(0.3),
             f"Module {module_num} \u2022 {module_title}".strip(" \u2022"), 13, RGBColor(0xB9, 0xD3, 0xF2))
    header_text_w = SLIDE_W - text_left - Inches(0.55)
    eq_title = f"{sub_title}  \u2014  Equipment Reference"
    eq_size = fit_one_line_size(eq_title, header_text_w / 914400, start_size=21, min_size=14)
    add_text(slide, text_left, Inches(0.48), header_text_w, Inches(0.68),
             eq_title, eq_size, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    # centered photo card
    card_w, card_h = Inches(6.4), Inches(5.4)
    card_x = (SLIDE_W - card_w) // 2
    card_y = BODY_TOP + Inches(0.1)
    add_rect(slide, card_x, card_y, card_w, card_h, CARD_WHITE, radius=0.04, shadow=True)
    try:
        image_file.seek(0)
        pic = slide.shapes.add_picture(image_file, card_x, card_y, width=card_w)
        if pic.height > card_h - Inches(0.3):
            image_file.seek(0)
            slide.shapes._spTree.remove(pic._element)
            pic = slide.shapes.add_picture(image_file, card_x, card_y, height=card_h - Inches(0.3))
            pic.left = card_x + (card_w - pic.width) // 2
        pic.top = card_y + (card_h - pic.height) // 2
    except Exception:
        add_text(slide, card_x, card_y, card_w, card_h, f"[{keyword}]", 16, TEXT_GRAY,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def draw_photo_gallery_slide(prs, items, logos, part_no, part_total):
    """Grid slide for uploaded reference photos whose keyword never matched
    any sub-module text. This guarantees every photo the user uploads in
    Streamlit ends up SOMEWHERE in the final deck, not silently dropped."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    text_left = Inches(0.55)
    if logos[0] is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    title = "સંદર્ભ તસવીરો (Reference Photos)" + (f"  ({part_no}/{part_total})" if part_total > 1 else "")
    add_text(slide, text_left, Inches(0.14), Inches(10), Inches(0.3), "Additional Reference", 13, RGBColor(0xB9, 0xD3, 0xF2))
    add_text(slide, text_left, Inches(0.48), Inches(10), Inches(0.68), title, 22, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    n = len(items)
    cols = max(1, min(4, n))
    rows = (n + cols - 1) // cols
    gap = Inches(0.22)
    grid_w = SLIDE_W - 2 * MARGIN_X
    grid_h = SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN
    card_w = (grid_w - gap * (cols - 1)) // cols
    card_h = (grid_h - gap * (rows - 1)) // rows
    for i, (kw, file) in enumerate(items):
        r, c = divmod(i, cols)
        x = MARGIN_X + c * (card_w + gap)
        y = BODY_TOP + r * (card_h + gap)
        add_rect(slide, x, y, card_w, card_h, CARD_WHITE, radius=0.06, shadow=True)
        pad = Inches(0.12)
        label_h = Inches(0.26)
        add_text(slide, x + pad, y + pad, card_w - 2 * pad, label_h, kw, 10.5, NAVY, bold=True, align=PP_ALIGN.CENTER)
        img_x, img_y = x + pad, y + pad + label_h + Inches(0.05)
        img_w = card_w - 2 * pad
        img_h = card_h - (img_y - y) - pad
        try:
            file.seek(0)
            pic = slide.shapes.add_picture(file, img_x, img_y, height=img_h)
            if pic.width > img_w:
                file.seek(0)
                slide.shapes._spTree.remove(pic._element)
                pic = slide.shapes.add_picture(file, img_x, img_y, width=img_w)
            pic.left = int(img_x + (img_w - pic.width) / 2)
            pic.top = int(img_y + (img_h - pic.height) / 2)
        except Exception:
            add_text(slide, img_x, img_y, img_w, img_h, f"[{kw}]", 11, TEXT_GRAY,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def draw_title_slide(prs, doc_title, logos, heritage_bg=None, riverfront=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # --- full-bleed blurred heritage photo as the background ---
    bg_drawn = False
    if heritage_bg is not None:
        try:
            bg_buf = prepare_cover_image(heritage_bg, 1600, 900, blur_radius=14)
            slide.shapes.add_picture(bg_buf, 0, 0, width=SLIDE_W, height=SLIDE_H)
            bg_drawn = True
        except Exception:
            bg_drawn = False
    if not bg_drawn:
        add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)

    # dark navy veil over the photo so white text/logos stay legible
    veil = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    set_shape_alpha(veil, 58)  # ~58% opaque navy over the photo

    add_rect(slide, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), GREEN)

    # logos, top center, in their own white roundels
    if logos[0] is not None:
        d = Inches(1.05)
        add_oval(slide, (SLIDE_W - d) // 2 - Inches(1.3), Inches(0.5), d, WHITE)
        slide.shapes.add_picture(logos[0], (SLIDE_W - d) // 2 - Inches(1.3) + Inches(0.08), Inches(0.58), height=d - Inches(0.16))
    if logos[1] is not None:
        d = Inches(1.05)
        add_oval(slide, (SLIDE_W - d) // 2 + Inches(1.3), Inches(0.5), d, WHITE)
        slide.shapes.add_picture(logos[1], (SLIDE_W - d) // 2 + Inches(1.3) + Inches(0.08), Inches(0.58), height=d - Inches(0.16))

    add_text(slide, Inches(1), Inches(1.78), SLIDE_W - Inches(2), Inches(0.45),
             "અમદાવાદ મ્યુનિસિપલ કોર્પોરેશન  \u2022  NTEP", 16, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1), Inches(2.25), SLIDE_W - Inches(2), Inches(1.25),
             doc_title, 32, WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1), Inches(3.6), SLIDE_W - Inches(2), Inches(0.4),
             "Public Health Actions \u2014 Standard Operating Procedures", 16, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)

    # --- riverfront photo as a crisp decorative banner near the bottom ---
    if riverfront is not None:
        try:
            band_w_in, band_h_in = 11.4, 1.5
            band_x = int((SLIDE_W - Inches(band_w_in)) // 2)
            band_y = Inches(4.35)
            add_rect(slide, band_x - Inches(0.06), band_y - Inches(0.06),
                     Inches(band_w_in + 0.12), Inches(band_h_in + 0.12), WHITE, radius=0.05, shadow=True)
            riv_buf = prepare_cover_image(riverfront, 1140, 150, blur_radius=0)
            slide.shapes.add_picture(riv_buf, band_x, band_y, width=Inches(band_w_in), height=Inches(band_h_in))
            add_text(slide, band_x, band_y + Inches(band_h_in) + Inches(0.06), Inches(band_w_in), Inches(0.3),
                     "સાબરમતી રિવરફ્રન્ટ, અમદાવાદ", 11, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)
        except Exception:
            pass

    return slide


def draw_closing_slide(prs, logos, closing_bg=None):
    """Thank-you slide. Uses the same "photo background" idea as the title
    slide, but LESS blurred so the heritage photo actually reads, plus a
    small tricolor strip + Ashoka Chakra medallion for the India touch
    instead of covering the whole slide in flat color bands."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg_drawn = False
    if closing_bg is not None:
        try:
            bg_buf = prepare_cover_image(closing_bg, 1600, 900, blur_radius=6)
            slide.shapes.add_picture(bg_buf, 0, 0, width=SLIDE_W, height=SLIDE_H)
            bg_drawn = True
        except Exception:
            bg_drawn = False
    if not bg_drawn:
        add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)

    # Lighter veil than the title slide (58 -> 44) so the photo stays visible.
    veil = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    set_shape_alpha(veil, 44)

    # thin tricolor accent strip along the very bottom
    strip_h = Inches(0.28)
    third = SLIDE_W // 3
    add_rect(slide, 0, SLIDE_H - strip_h, third, strip_h, SAFFRON)
    add_rect(slide, third, SLIDE_H - strip_h, third, strip_h, WHITE)
    add_rect(slide, 2 * third, SLIDE_H - strip_h, SLIDE_W - 2 * third, strip_h, INDIA_GREEN)

    # small Ashoka Chakra medallion, centered, sitting just above the strip
    outer_d = Inches(0.62)
    cx = SLIDE_W // 2
    cy = SLIDE_H - strip_h - outer_d // 2 - Inches(0.06)
    add_oval(slide, cx - outer_d // 2, cy - outer_d // 2, outer_d, WHITE)
    inner_d = int(outer_d * 0.8)
    draw_ashoka_chakra(slide, cx - inner_d // 2, cy - inner_d // 2, inner_d, CHAKRA_NAVY)

    add_text(slide, Inches(1), Inches(2.85), SLIDE_W - Inches(2), Inches(1),
              "આભાર (Thank You)", 34, WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1), Inches(3.8), SLIDE_W - Inches(2), Inches(0.5),
              "NTEP  \u2022  Ahmedabad Municipal Corporation", 15, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)

    if logos[0] is not None:
        d = Inches(0.9)
        add_oval(slide, Inches(0.4), Inches(0.4), d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.4) + Inches(0.07), Inches(0.4) + Inches(0.07), height=d - Inches(0.14))
    if logos[1] is not None:
        d = Inches(0.9)
        rx = SLIDE_W - Inches(0.4) - d
        add_oval(slide, rx, Inches(0.4), d, WHITE)
        slide.shapes.add_picture(logos[1], rx + Inches(0.07), Inches(0.4) + Inches(0.07), height=d - Inches(0.14))

    return slide


def draw_preface_slide(prs, heading_gj, heading_en, paragraphs, icon, logos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    text_left = Inches(0.55)
    if logos[0] is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    add_text(slide, text_left, Inches(0.14), Inches(10), Inches(0.3), heading_en, 13, RGBColor(0xB9, 0xD3, 0xF2))
    add_text(slide, text_left, Inches(0.48), Inches(10), Inches(0.68), heading_gj, 26, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    card_x, card_y = MARGIN_X, BODY_TOP
    card_w = SLIDE_W - 2 * MARGIN_X
    card_h = SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN
    add_rect(slide, card_x, card_y, card_w, card_h, CARD_WHITE, radius=0.03, shadow=True)
    icon_d = Inches(0.5)
    ICON_FN[icon](slide, card_x + Inches(0.4), card_y + Inches(0.4), icon_d, GREEN)
    body_x = card_x + Inches(0.4)
    body_y = card_y + Inches(0.4) + icon_d + Inches(0.2)
    body_w = card_w - Inches(0.8)
    body_h = card_h - (body_y - card_y) - Inches(0.35)
    content_w_in = body_w / 914400
    size, chunks = chunk_items(paragraphs, content_w_in, body_h / 914400, bullet=False)
    flat = [p for c in chunks for p in c]
    add_items(slide, body_x, body_y, body_w, body_h, flat, size, TEXT_GRAY, bullet=False)
    return slide


def draw_module_overview_slide(prs, modules, module_start_slide, logos):
    """'Modules at a Glance' — grid of module cards, each showing its title
    and the number of sub-modules it contains, real page number linked via TOC."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    text_left = Inches(0.55)
    if logos[0] is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    add_text(slide, text_left, Inches(0.14), Inches(10), Inches(0.3), "Overview", 13, RGBColor(0xB9, 0xD3, 0xF2))
    add_text(slide, text_left, Inches(0.48), Inches(10), Inches(0.68), "મોડ્યુલ ઝલક  (Modules at a Glance)", 24, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    mod_nums = sorted(modules.keys(), key=lambda x: int(x))
    n = len(mod_nums)
    cols = 2 if n <= 4 else (3 if n <= 6 else 4)
    rows = (n + cols - 1) // cols
    gap = Inches(0.22)
    grid_w = SLIDE_W - 2 * MARGIN_X
    grid_h = SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN
    card_w = (grid_w - gap * (cols - 1)) // cols
    card_h = (grid_h - gap * (rows - 1)) // rows
    for i, mod_num in enumerate(mod_nums):
        r, c = divmod(i, cols)
        x = MARGIN_X + c * (card_w + gap)
        y = BODY_TOP + r * (card_h + gap)
        m_color = module_accent_color(mod_num)
        add_rect(slide, x, y, card_w, card_h, CARD_WHITE, radius=0.07, shadow=True)
        add_rect(slide, x, y, Inches(0.08), card_h, m_color, radius=0.02)
        badge_d = Inches(0.5)
        add_oval(slide, x + Inches(0.2), y + Inches(0.2), badge_d, m_color)
        add_text(slide, x + Inches(0.2), y + Inches(0.2), badge_d, badge_d, mod_num, 18, WHITE,
                  bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        title_x = x + Inches(0.2) + badge_d + Inches(0.15)
        title_w = card_w - (title_x - x) - Inches(0.2)
        add_text(slide, title_x, y + Inches(0.18), title_w, badge_d, modules[mod_num]["title"], 13, NAVY,
                  bold=True, anchor=MSO_ANCHOR.MIDDLE)
        n_subs = len(modules[mod_num]["subs"])
        page = module_start_slide.get(mod_num, "\u2014")
        add_text(slide, x + Inches(0.2), y + card_h - Inches(0.5), card_w - Inches(0.4), Inches(0.32),
                  f"{n_subs} sub-modules  \u2022  Slide {page}", 11.5, TEXT_GRAY)
    return slide


def draw_module_divider_slide(prs, module_num, module_title, sub_titles, logos):
    m_color = module_accent_color(module_num)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(slide, 0, 0, Inches(0.16), SLIDE_H, m_color)
    badge_d = Inches(0.85)
    add_oval(slide, Inches(0.9), Inches(0.75), badge_d, m_color)
    add_text(slide, Inches(0.9), Inches(0.75), badge_d, badge_d, module_num, 26, WHITE,
              bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(0.9) + badge_d + Inches(0.25), Inches(0.85), Inches(3), Inches(0.6), "Module", 16, RGBColor(0xB9, 0xD3, 0xF2), bold=True)
    add_text(slide, Inches(0.9), Inches(1.85), SLIDE_W - Inches(1.8), Inches(1.1), module_title, 32, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    y = Inches(3.15)
    for i, (sub_num, sub_title) in enumerate(sub_titles):
        add_oval(slide, Inches(0.9), y + Inches(0.03), Inches(0.16), m_color)
        add_text(slide, Inches(1.25), y - Inches(0.05), SLIDE_W - Inches(2.2), Inches(0.4), f"{sub_num}  {sub_title}", 15, WHITE)
        y += Inches(0.44)
        if y > SLIDE_H - Inches(0.6):
            break
    return slide


def draw_toc_slide(prs, entries, page_no_start, logos, part_no, part_total):
    """entries: list of (level, label, page_str). level 0 = module, 1 = sub-module."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY_LIGHT)
    text_left = Inches(0.55)
    if logos[0] is not None:
        d = Inches(0.85)
        add_oval(slide, Inches(0.25), (HEADER_H - d) // 2, d, WHITE)
        slide.shapes.add_picture(logos[0], Inches(0.32), (HEADER_H - d) // 2 + Inches(0.07), height=d - Inches(0.14))
        text_left = Inches(1.35)
    title = "અનુક્રમણિકા (Index)" + (f"  ({part_no}/{part_total})" if part_total > 1 else "")
    add_text(slide, text_left, Inches(0.14), Inches(10), Inches(0.3), "Table of Contents", 13, RGBColor(0xB9, 0xD3, 0xF2))
    add_text(slide, text_left, Inches(0.48), Inches(10), Inches(0.68), title, 24, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    col_w = (SLIDE_W - 2 * MARGIN_X - COL_GAP) // 2
    col_x = [MARGIN_X, MARGIN_X + col_w + COL_GAP]
    n = len(entries)
    half = (n + 1) // 2
    columns = [entries[:half], entries[half:]]
    col_w_in = col_w / 914400
    for ci, col_entries in enumerate(columns):
        y = BODY_TOP
        for level, label, page in col_entries:
            row_h_in = toc_row_height_in(level, label, col_w_in) - 0.04
            row_h = Emu(int(row_h_in * 914400))
            if level == 0:
                add_rect(slide, col_x[ci], y, col_w, row_h, NAVY_LIGHT, radius=0.15)
                add_text(slide, col_x[ci] + Inches(0.15), y, col_w - Inches(1.0), row_h, label, 13.5, WHITE,
                          bold=True, anchor=MSO_ANCHOR.MIDDLE)
                add_text(slide, col_x[ci] + col_w - Inches(0.8), y, Inches(0.7), row_h, str(page), 12.5, WHITE,
                          align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
            else:
                add_text(slide, col_x[ci] + Inches(0.3), y, col_w - Inches(1.1), row_h, label, 12, TEXT_DARK,
                          anchor=MSO_ANCHOR.MIDDLE)
                add_text(slide, col_x[ci] + col_w - Inches(0.8), y, Inches(0.7), row_h, str(page), 11.5, TEXT_GRAY,
                          align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
            y += row_h + Inches(0.04)
    return slide


def toc_row_height_in(level, label, col_w_in):
    """Wrap-aware row height so a long module/sub-module title that wraps to 2
    lines gets a taller row instead of spilling out of its box."""
    size = 13.5 if level == 0 else 12
    pad_left = 0.15 if level == 0 else 0.3
    text_w_in = col_w_in - pad_left - 0.8
    lines = wrap_text(label, size, max(0.5, text_w_in), bold=(level == 0))
    base = 0.30 if level == 0 else 0.24
    per_extra_line = size * 1.18 / 72
    h = base + max(0, len(lines) - 1) * per_extra_line
    return h + 0.04


def paginate_toc(entries, avail_h_in, col_w_in):
    half_budget = avail_h_in  # each column has this much room; 2 cols per slide
    pages = []
    cur = []
    cur_h = [0.0, 0.0]
    col = 0
    for e in entries:
        level, label = e[0], e[1]
        h = toc_row_height_in(level, label, col_w_in)
        if cur_h[col] + h > half_budget:
            col += 1
            if col > 1:
                pages.append(cur)
                cur = []
                cur_h = [0.0, 0.0]
                col = 0
        cur.append(e)
        cur_h[col] += h
    if cur:
        pages.append(cur)
    return pages


# ----------------------------------------------------------------------
# 7. EQUIPMENT PHOTO MATCHING
# ----------------------------------------------------------------------
def normalize_keyword(name):
    base = os.path.splitext(name)[0]
    return re.sub(r"[\s_\-]+", "", base).lower()


def submodule_joined_tokens(sub):
    """Tokenize a sub-module's full text and return the set of every
    1-4-consecutive-word run joined with no spaces (e.g. "chest","x","ray"
    -> "chestxray"). Shared by the automatic keyword matcher AND the
    placement-UI's auto-suggestions, so both agree on what "matches" means."""
    raw = " ".join(" ".join(sub.get(k, [])) for k in FIELD_META)
    tokens = [t.lower() for t in re.findall(r"[^\W_]+", raw, re.UNICODE)]
    joined = set()
    max_n = 4
    for n in range(1, max_n + 1):
        for i in range(len(tokens) - n + 1):
            joined.add("".join(tokens[i:i + n]))
    return joined


def find_equipment_matches(sub, image_map):
    """Scan a sub-module's field text for any uploaded-photo keyword.
    IMPORTANT: this matches on whole WORD boundaries, not raw substrings —
    see submodule_joined_tokens(). This is used only to compute the
    *suggested default* placement shown in the Streamlit UI; the deck itself
    is built from whatever the user confirms/edits there (see
    auto_suggest_subs_for_keyword and the manual-placement UI below), because
    pure keyword auto-matching alone was still too unreliable on real
    documents to trust without a human check.
    """
    joined = submodule_joined_tokens(sub)
    matches = []
    for kw, file in image_map.items():
        if kw and len(kw) >= 3 and kw in joined:
            matches.append((kw, file))
    return matches


def sub_label(modules, mod_num, sub_num):
    return f"{sub_num}  {modules[mod_num]['subs'][sub_num]['title']}".strip()


def auto_suggest_subs_for_keyword(kw, modules):
    """Every (mod_num, sub_num) whose text contains `kw` as a whole
    word/phrase, in document order. Used only to pre-fill the manual
    placement multiselect — the user can add/remove freely."""
    if not kw or len(kw) < 3:
        return []
    out = []
    for mod_num in sorted(modules.keys(), key=lambda x: int(x)):
        for sub_num in sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split(".")))):
            sub = modules[mod_num]["subs"][sub_num]
            if kw in submodule_joined_tokens(sub):
                out.append((mod_num, sub_num))
    return out


# ----------------------------------------------------------------------
# 8. TWO-PASS BUILD: plan everything first (for real TOC page numbers),
#    then draw.
# ----------------------------------------------------------------------
def run_build(preface, modules, logos, image_map, manual_placement,
              equip_max_repeats=DEFAULT_EQUIP_MAX_REPEATS,
              heritage_bg=None, riverfront=None, closing_bg=None, progress_cb=None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    mod_nums = sorted(modules.keys(), key=lambda x: int(x))

    # ---- Pass 1: figure out how many slides everything takes, in order ----
    fixed_pre = []
    fixed_pre.append(("title", None))
    if preface.get("PREFACE"):
        fixed_pre.append(("preface_pref", None))
    if preface.get("PURPOSE"):
        fixed_pre.append(("preface_purpose", None))
    if preface.get("OBJECTIVES"):
        fixed_pre.append(("preface_objectives", None))

    entries = []
    for mod_num in mod_nums:
        entries.append((0, f"Module {mod_num}: {modules[mod_num]['title']}", None, ("module", mod_num)))
        for sub_num in sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split(".")))):
            title = modules[mod_num]["subs"][sub_num]["title"]
            entries.append((1, f"{sub_num}  {title}", None, ("sub", mod_num, sub_num)))

    body_h_in_toc = (SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN) / 914400
    toc_col_w_in = ((SLIDE_W - 2 * MARGIN_X - COL_GAP) // 2) / 914400
    toc_pages = paginate_toc(entries, body_h_in_toc, toc_col_w_in)
    n_toc_slides = max(1, len(toc_pages))
    overview_slides = 1

    submodule_matches = {}
    submodule_plans = {}
    # `manual_placement` is {keyword: [(mod_num, sub_num), ...]} built from the
    # user's explicit choices in the Streamlit UI (pre-filled with keyword
    # auto-suggestions, but fully user-editable) — this replaces blind
    # keyword auto-matching as the source of truth for where each photo goes,
    # which is what actually fixes photos landing in the wrong/no slides.
    # How many sub-modules each keyword has already been placed into. Walking
    # sub-modules in real document order and capping here means the SAME
    # keyword/photo can never end up in more than `equip_max_repeats` places
    # in the whole deck (defensive re-check — the UI already enforces this).
    equip_usage_count = {kw: 0 for kw in image_map}
    for mod_num in mod_nums:
        sub_nums_sorted = sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split("."))))
        for sub_num in sub_nums_sorted:
            sub = modules[mod_num]["subs"][sub_num]
            raw_matches = [(kw, image_map[kw]) for kw, refs in manual_placement.items()
                            if (mod_num, sub_num) in refs]
            matches = []
            for kw, file in raw_matches:
                if equip_usage_count.get(kw, 0) < equip_max_repeats:
                    matches.append((kw, file))
                    equip_usage_count[kw] = equip_usage_count.get(kw, 0) + 1
            submodule_matches[(mod_num, sub_num)] = matches
            image_match = matches[0] if matches else None
            submodule_plans[(mod_num, sub_num)] = plan_submodule(sub, image_match)

    slide_no = 1
    slide_no += (len(fixed_pre) - 1)
    slide_no += n_toc_slides
    slide_no += overview_slides
    module_start_slide = {}
    submodule_start_slide = {}
    for mod_num in mod_nums:
        slide_no += 1
        module_start_slide[mod_num] = slide_no
        for sub_num in sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split(".")))):
            submodule_start_slide[(mod_num, sub_num)] = slide_no
            plan = submodule_plans[(mod_num, sub_num)]
            slide_no += len(plan)
            slide_no += len(submodule_matches[(mod_num, sub_num)])

    filled_entries = []
    for level, label, _, ref in entries:
        if ref[0] == "module":
            page = module_start_slide[ref[1]]
        else:
            page = submodule_start_slide[(ref[1], ref[2])]
        filled_entries.append((level, label, page))

    # ---- Pass 2: actually draw everything in the same order ----
    draw_title_slide(prs, "Public Health Actions of the TB Department", logos,
                      heritage_bg=heritage_bg, riverfront=riverfront)
    preface_icons = {"PREFACE": "info", "PURPOSE": "target", "OBJECTIVES": "check"}
    preface_headings = {"PREFACE": ("પ્રસ્તાવના", "Preface"), "PURPOSE": ("હેતુ", "Purpose"),
                          "OBJECTIVES": ("માર્ગદર્શિકાના મુખ્ય ઉદ્દેશ્યો", "Key Objectives")}
    for key in ["PREFACE", "PURPOSE", "OBJECTIVES"]:
        if preface.get(key):
            gj, en = preface_headings[key]
            draw_preface_slide(prs, gj, en, preface[key], preface_icons[key], logos)

    toc_pages_final = paginate_toc(
        [(lvl, lbl, pg) for lvl, lbl, pg in filled_entries], body_h_in_toc, toc_col_w_in
    )
    for i, page_entries in enumerate(toc_pages_final):
        draw_toc_slide(prs, page_entries, None, logos, i + 1, len(toc_pages_final))

    draw_module_overview_slide(prs, modules, module_start_slide, logos)

    total = len(mod_nums)
    for mi, mod_num in enumerate(mod_nums):
        sub_items = sorted(modules[mod_num]["subs"].items(), key=lambda x: tuple(map(int, x[0].split("."))))
        draw_module_divider_slide(prs, mod_num, modules[mod_num]["title"],
                                    [(k, v["title"]) for k, v in sub_items], logos)
        for sub_num, sub in sub_items:
            display_title = f"{sub_num}  {sub.get('title', '')}".strip()
            plan = submodule_plans[(mod_num, sub_num)]
            total_parts = len(plan)
            for pi, (group_label, layout) in enumerate(plan):
                draw_field_slide(prs, mod_num, modules[mod_num]["title"], display_title,
                                   group_label, pi + 1, total_parts, layout, logos)
            for kw, file in submodule_matches[(mod_num, sub_num)]:
                draw_equipment_slide(prs, mod_num, modules[mod_num]["title"], display_title,
                                       kw, file, logos)
        if progress_cb:
            progress_cb((mi + 1) / total)

    # Any uploaded photo whose keyword never matched a single sub-module still
    # needs to make it into the deck — collect it into a reference gallery so
    # "all photos I upload" really do end up in the pptx.
    leftover_photos = [(kw, image_map[kw]) for kw in image_map if equip_usage_count.get(kw, 0) == 0]
    if leftover_photos:
        PER_GALLERY = 8
        gallery_chunks = [leftover_photos[i:i + PER_GALLERY] for i in range(0, len(leftover_photos), PER_GALLERY)]
        for gi, chunk in enumerate(gallery_chunks):
            draw_photo_gallery_slide(prs, chunk, logos, gi + 1, len(gallery_chunks))

    # ---- closing slide: light-blur heritage photo + tricolor/chakra accent ----
    draw_closing_slide(prs, logos, closing_bg=closing_bg)

    return prs


# ----------------------------------------------------------------------
# 9. STREAMLIT UI
# ----------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    left_logo_file = st.file_uploader("Left logo (e.g. AMC seal) — optional", type=["png", "jpg", "jpeg"])
with col2:
    right_logo_file = st.file_uploader("Right logo (e.g. NTEP logo) — optional", type=["png", "jpg", "jpeg"])

st.markdown(
    "**Front & closing page artwork.** The title slide uses a blurred heritage "
    "photo as its background plus a sharp riverfront banner. These default to "
    f"`{os.path.basename(DEFAULT_HERITAGE_BG)}` and `{os.path.basename(DEFAULT_RIVERFRONT)}` "
    "if those files already sit next to `app.py` in the repo — upload replacements "
    "below only if you want to override them."
)
colA, colB, colC = st.columns(3)
with colA:
    heritage_upload = st.file_uploader("Title background (heritage photo)", type=["png", "jpg", "jpeg"], key="heritage_bg")
with colB:
    riverfront_upload = st.file_uploader("Title banner (riverfront photo)", type=["png", "jpg", "jpeg"], key="riverfront_img")
with colC:
    closing_bg_upload = st.file_uploader("Closing page background (optional)", type=["png", "jpg", "jpeg"], key="closing_bg")

heritage_bg = resolve_front_image(heritage_upload, DEFAULT_HERITAGE_BG)
riverfront = resolve_front_image(riverfront_upload, DEFAULT_RIVERFRONT)
# Closing page defaults to the riverfront shot (so it doesn't look identical
# to the title page); falls back to the heritage photo if riverfront is missing.
closing_bg = resolve_front_image(closing_bg_upload, DEFAULT_RIVERFRONT)
if closing_bg is None:
    closing_bg = resolve_front_image(None, DEFAULT_HERITAGE_BG)
if heritage_bg is None:
    st.info(f"No heritage background found (looked for `{DEFAULT_HERITAGE_BG}`). Title slide will use a solid navy background instead.")
if riverfront is None:
    st.info(f"No riverfront photo found (looked for `{DEFAULT_RIVERFRONT}`). Title slide will skip the decorative banner.")

st.markdown(
    "**Equipment / reference photos (optional).** Upload any number of photos. "
    "After you upload your Word document below, you'll get an explicit picker "
    "for **exactly which sub-module(s) each photo appears on** — pre-filled with "
    "a keyword guess, but yours to correct. This is what actually fixes photos "
    "landing on the wrong slide or not appearing at all: placement is no longer "
    "left purely to automatic keyword matching. **Every uploaded photo is "
    "guaranteed to appear somewhere in the deck** — leave a photo unplaced and "
    "it lands on a dedicated reference-photo slide near the end instead of "
    "being dropped."
)
equipment_files = st.file_uploader(
    "Equipment photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True
)
equip_max_repeats = st.number_input(
    "Max times a single photo can repeat across the whole deck",
    min_value=1, max_value=15, value=DEFAULT_EQUIP_MAX_REPEATS, step=1,
    help="Once a photo has been placed this many times, further keyword matches for it are skipped."
)

uploaded_docx = st.file_uploader("Upload Gujarati Word Document (.docx)", type=["docx"])

if uploaded_docx:
    preface, modules = parse_word(uploaded_docx)
    for mod in modules.values():
        for sub in mod["subs"].values():
            sub.setdefault("title_display", sub["title"])
    total_subs = sum(len(m["subs"]) for m in modules.values())
    with st.expander(f"🔍 Parsed structure — {len(modules)} modules, {total_subs} sub-modules"):
        st.write({"preface": preface, "modules": modules})

    # Build the keyword->file map, but warn (instead of silently losing a photo)
    # if two uploaded filenames normalize to the same keyword.
    image_map = {}
    collisions = []
    for f in (equipment_files or []):
        key = normalize_keyword(f.name)
        if key in image_map and image_map[key].name != f.name:
            collisions.append((f.name, image_map[key].name, key))
        image_map[key] = f
    if image_map:
        st.caption("Equipment keywords detected: " + ", ".join(sorted(image_map.keys())))
    if collisions:
        for new_name, old_name, key in collisions:
            st.warning(
                f"⚠️ '{new_name}' and '{old_name}' both normalize to the same keyword "
                f"'{key}', so only one of them will be used. Rename one of the files "
                f"(e.g. add a number) if they're meant to be different photos."
            )

    if total_subs == 0:
        st.error("No sub-modules (like 1.1, 1.2) were detected. Check your Word doc headings, "
                  "or adjust MODULE_RE / SUBMODULE_RE in the script if your numbering format differs.")
    elif st.button("Generate Presentation", type="primary"):
        progress = st.progress(0.0, text="Designing slides...")
        with st.spinner("Measuring content and designing slides..."):
            prs = run_build(preface, modules, (left_logo_file, right_logo_file), image_map,
                              equip_max_repeats=int(equip_max_repeats),
                              heritage_bg=heritage_bg, riverfront=riverfront,
                              progress_cb=lambda f: progress.progress(f, text="Designing slides..."))
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
        progress.empty()
        st.success(f"✅ Generated {len(prs.slides._sldIdLst)} slides for {total_subs} sub-modules.")
        st.download_button(
            "📄 Download Presentation",
            data=ppt_io,
            file_name="AMC_NTEP_AutoDesigned.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
