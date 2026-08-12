import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import io
import os
import re
from PIL import ImageFont
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
def tint(rgb_color, amount=0.82):
    """Blend a theme color toward white to get a soft halo/background tint of
    the same hue — keeps things colorful without making anything darker."""
    hexs = str(rgb_color)
    r, g, b = int(hexs[0:2], 16), int(hexs[2:4], 16), int(hexs[4:6], 16)
    r = int(r + (255 - r) * amount)
    g = int(g + (255 - g) * amount)
    b = int(b + (255 - b) * amount)
    return RGBColor(r, g, b)
FONT = "Noto Sans Gujarati"
# Font files must sit next to this script (or set full paths).
FONT_TTF_REGULAR = os.path.join(os.path.dirname(__file__), "NotoSansGujarati-Regular.ttf")
FONT_TTF_BOLD = os.path.join(os.path.dirname(__file__), "NotoSansGujarati-Bold.ttf")
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
# appears in more places in the text.
EQUIP_MAX_REPEATS = 3
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
        r.font.name = FONT
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
        r.font.name = FONT
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
def draw_title_slide(prs, doc_title, logos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(slide, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), GREEN)
    if logos[0] is not None:
        d = Inches(1.1)
        add_oval(slide, (SLIDE_W - d) // 2 - Inches(1.3), Inches(0.7), d, WHITE)
        slide.shapes.add_picture(logos[0], (SLIDE_W - d) // 2 - Inches(1.3) + Inches(0.08), Inches(0.78), height=d - Inches(0.16))
    if logos[1] is not None:
        d = Inches(1.1)
        add_oval(slide, (SLIDE_W - d) // 2 + Inches(1.3), Inches(0.7), d, WHITE)
        slide.shapes.add_picture(logos[1], (SLIDE_W - d) // 2 + Inches(1.3) + Inches(0.08), Inches(0.78), height=d - Inches(0.16))
    add_text(slide, Inches(1), Inches(2.5), SLIDE_W - Inches(2), Inches(0.5),
             "અમદાવાદ મ્યુનિસિપલ કોર્પોરેશન  \u2022  NTEP", 16, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1), Inches(3.0), SLIDE_W - Inches(2), Inches(1.6),
             doc_title, 34, WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1), Inches(4.7), SLIDE_W - Inches(2), Inches(0.4),
             "Public Health Actions \u2014 Standard Operating Procedures", 16, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)
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
def find_equipment_matches(sub, image_map):
    """Scan a sub-module's field text for any uploaded-photo keyword.

    IMPORTANT: this matches on whole WORD boundaries, not raw substrings.
    Filenames are normalized by stripping spaces/underscores/hyphens (e.g.
    "Chest X-Ray.jpg" and "chest_x_ray.png" both become "chestxray"). To find
    that keyword in the sub-module's text we tokenize the text into real words
    and join up to 4 *consecutive* words back together (e.g. "chest", "x",
    "ray" -> "chestxray") and only count it as a match if that joined phrase
    is EXACTLY equal to the keyword. This is what prevents a short/generic
    keyword from accidentally matching inside unrelated words elsewhere in the
    document (the old version concatenated the entire text into one long
    string with no spaces and did a plain substring search, which is why a
    photo could show up on unrelated slides while another photo never
    matched at all).
    """
    raw = " ".join(" ".join(sub.get(k, [])) for k in FIELD_META)
    # Extract pure word-character runs only (letters/digits, Unicode-aware),
    # so punctuation glued onto a word — "(GeneXpert)", "CBNAAT,", "NAAT."
    # etc. — no longer prevents it from equaling a clean filename keyword.
    # (An earlier version split only on whitespace/underscore/hyphen, so a
    # trailing bracket or comma silently broke the match — that's why a
    # correctly-named photo still wasn't showing up at all.)
    tokens = [t.lower() for t in re.findall(r"[^\W_]+", raw, re.UNICODE)]
    joined = set()
    max_n = 4
    for n in range(1, max_n + 1):
        for i in range(len(tokens) - n + 1):
            joined.add("".join(tokens[i:i + n]))
    matches = []
    for kw, file in image_map.items():
        if kw and len(kw) >= 3 and kw in joined:
            matches.append((kw, file))
    return matches
# ----------------------------------------------------------------------
# 8. TWO-PASS BUILD: plan everything first (for real TOC page numbers),
#    then draw.
# ----------------------------------------------------------------------
def build_plan(modules, image_map):
    """Returns a flat ordered list of 'slide jobs' plus metadata needed for TOC/overview.
    Each job is a dict describing what draw_* call to make later."""
    jobs = []  # list of dicts
    jobs.append({"kind": "title"})
    # Preface jobs are appended by caller before this (fixed count), so start
    # counting sub-module content pages after title + preface + TOC(placeholder) + overview.
    return jobs  # (kept for clarity; actual orchestration happens in run_build)
def run_build(preface, modules, logos, image_map, progress_cb=None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    mod_nums = sorted(modules.keys(), key=lambda x: int(x))
    # ---- Pass 1: figure out how many slides everything takes, in order ----
    # Fixed-position pieces
    fixed_pre = []
    fixed_pre.append(("title", None))
    if preface.get("PREFACE"):
        fixed_pre.append(("preface_pref", None))
    if preface.get("PURPOSE"):
        fixed_pre.append(("preface_purpose", None))
    if preface.get("OBJECTIVES"):
        fixed_pre.append(("preface_objectives", None))
    # TOC entries (module + sub-module rows) — page numbers filled after we know
    # how many slides precede each item. We don't know TOC's own slide count yet
    # (it depends on entry count), so: build entries first with placeholder pages,
    # estimate TOC slide count from entry count, then compute real offsets, and
    # if the TOC slide count assumption was wrong (rare, off-by-one from column
    # balancing) redo once — in practice one pass is stable because TOC length
    # only depends on structure, not content.
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
    # module divider (1) + per-submodule plan.
    # Compute each sub-module's equipment matches ONCE here and reuse them for
    # both pass 1 (counting) and pass 2 (drawing) so the two passes can never
    # disagree with each other.
    submodule_matches = {}
    submodule_plans = {}  # (mod_num, sub_num) -> plan (list of (group,layout))
    # How many sub-modules each keyword has already been placed into. Walking
    # sub-modules in real document order and capping here means the SAME
    # keyword/photo can never end up in more than EQUIP_MAX_REPEATS places in
    # the whole deck, no matter how many times its word appears in the text.
    equip_usage_count = {kw: 0 for kw in image_map}
    for mod_num in mod_nums:
        sub_nums_sorted = sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split("."))))
        for sub_num in sub_nums_sorted:
            sub = modules[mod_num]["subs"][sub_num]
            raw_matches = find_equipment_matches(sub, image_map)
            matches = []
            for kw, file in raw_matches:
                if equip_usage_count.get(kw, 0) < EQUIP_MAX_REPEATS:
                    matches.append((kw, file))
                    equip_usage_count[kw] = equip_usage_count.get(kw, 0) + 1
            submodule_matches[(mod_num, sub_num)] = matches
            # Only the first match gets embedded as a small in-slide thumbnail;
            # if a sub-module happens to match more than one photo, the rest
            # still get their own full-page "Equipment Reference" slide below.
            image_match = matches[0] if matches else None
            submodule_plans[(mod_num, sub_num)] = plan_submodule(sub, image_match)
    # Now compute absolute slide numbers (1-indexed) for every module & sub-module.
    slide_no = 1 + len(fixed_pre) - 1 + 1  # placeholder, recomputed properly below
    # Recompute cleanly:
    slide_no = 1  # title
    slide_no += (len(fixed_pre) - 1)  # preface slides (title already counted)
    slide_no += n_toc_slides
    slide_no += overview_slides
    module_start_slide = {}
    submodule_start_slide = {}
    for mod_num in mod_nums:
        slide_no += 1  # module divider
        module_start_slide[mod_num] = slide_no
        for sub_num in sorted(modules[mod_num]["subs"].keys(), key=lambda x: tuple(map(int, x.split(".")))):
            submodule_start_slide[(mod_num, sub_num)] = slide_no
            plan = submodule_plans[(mod_num, sub_num)]
            slide_no += len(plan)
            slide_no += len(submodule_matches[(mod_num, sub_num)])
    total_slides = slide_no - 1 + 1  # closing slide follows
    # Fill entries' page numbers
    filled_entries = []
    for level, label, _, ref in entries:
        if ref[0] == "module":
            page = module_start_slide[ref[1]]
        else:
            page = submodule_start_slide[(ref[1], ref[2])]
        filled_entries.append((level, label, page))
    # ---- Pass 2: actually draw everything in the same order ----
    draw_title_slide(prs, "Public Health Actions of the TB Department", logos)
    preface_icons = {"PREFACE": "info", "PURPOSE": "target", "OBJECTIVES": "check"}
    preface_headings = {"PREFACE": ("પ્રસ્તાવના", "Preface"), "PURPOSE": ("હેતુ", "Purpose"),
                          "OBJECTIVES": ("માર્ગદર્શિકાના મુખ્ય ઉદ્દેશ્યો", "Key Objectives")}
    for key in ["PREFACE", "PURPOSE", "OBJECTIVES"]:
        if preface.get(key):
            gj, en = preface_headings[key]
            draw_preface_slide(prs, gj, en, preface[key], preface_icons[key], logos)
    # re-chunk filled_entries onto the same page boundaries computed in pass 1
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
            # Reuse the exact same match list computed in pass 1 (deterministic,
            # but this also guarantees pass 1's slide count and pass 2's actual
            # slides can never drift apart).
            for kw, file in submodule_matches[(mod_num, sub_num)]:
                draw_equipment_slide(prs, mod_num, modules[mod_num]["title"], display_title,
                                       kw, file, logos)
        if progress_cb:
            progress_cb((mi + 1) / total)
    # closing slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_text(slide, Inches(1), Inches(3.2), SLIDE_W - Inches(2), Inches(1), "આભાર (Thank You)", 32, WHITE,
              bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1), Inches(4.1), SLIDE_W - Inches(2), Inches(0.5),
              "NTEP \u2014 Ahmedabad Municipal Corporation", 15, RGBColor(0xB9, 0xD3, 0xF2), align=PP_ALIGN.CENTER)
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
    "**Equipment / reference photos (optional).** Upload any number of photos. "
    "Name each file after the keyword that should trigger it — e.g. `genexpert.png` "
    "or `chest_x_ray.jpg` will auto-insert a reference slide (plus a small photo "
    "card right on the relevant content slide) into any sub-module whose text "
    "mentions that exact word/phrase (case-insensitive, spaces/underscores/hyphens "
    "ignored). Matching now requires the whole word/phrase to appear together in "
    "the text — a short keyword can no longer accidentally match unrelated slides. "
    "Each photo is also capped at appearing in at most 3 sub-modules, even if its "
    "keyword genuinely comes up more often in the document."
)
equipment_files = st.file_uploader(
    "Equipment photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True
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
