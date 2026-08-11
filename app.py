import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import io
import re
from PIL import ImageFont

st.set_page_config(page_title="AMC NTEP - Slide Auto-Designer", layout="wide")
st.title("AMC NTEP - Sub-Module Slide Auto-Designer")
st.caption(
    "Upload your Word doc. This builds fully-designed slide(s) per sub-module "
    "(1.1, 1.2, 5.1 ...) automatically, measuring your real text so nothing "
    "overflows the cards — long sub-modules simply continue onto a second slide."
)

# ----------------------------------------------------------------------
# THEME
# ----------------------------------------------------------------------
NAVY = RGBColor(0x14, 0x2C, 0x5C)
NAVY_LIGHT = RGBColor(0x28, 0x49, 0x86)
GREEN = RGBColor(0x1F, 0xA6, 0x8C)
BLUE_ICON = RGBColor(0x1F, 0x5C, 0xC9)
BG_LIGHT = RGBColor(0xEE, 0xF5, 0xFB)
CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x22, 0x2B, 0x3A)
TEXT_GRAY = RGBColor(0x46, 0x50, 0x60)
HIGHLIGHT_BG = RGBColor(0xC9, 0xF3, 0xEE)
HIGHLIGHT_TEXT = RGBColor(0x0B, 0x6E, 0x5D)

FONT = "Noto Sans Gujarati"
FONT_TTF_REGULAR = "/usr/share/fonts/truetype/noto/NotoSansGujarati-Regular.ttf"
FONT_TTF_BOLD = "/usr/share/fonts/truetype/noto/NotoSansGujarati-Bold.ttf"
MEASURE_DPI = 96

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN_X = Inches(0.55)
HEADER_H = Inches(1.25)
BODY_TOP = HEADER_H + Inches(0.32)
BODY_BOTTOM_MARGIN = Inches(0.4)
COL_GAP = Inches(0.28)
CARD_GAP_V = Inches(0.2)
CARD_PAD = Inches(0.16)

FIELD_META = {
    "OBJ": {"label": "ઉદ્દેશ્ય (Objective)", "icon": "target", "color": GREEN, "bullet": False},
    "STEPS": {"label": "શું કરવું? (What to do)", "icon": "gear", "color": BLUE_ICON, "bullet": True},
    "WHO": {"label": "જવાબદાર વ્યક્તિ (Responsible)", "icon": "people", "color": BLUE_ICON, "bullet": False},
    "TIME": {"label": "સમયમર્યાદા (Timeline)", "icon": "clock", "color": GREEN, "bullet": False},
    "IND": {"label": "મોનિટરિંગ સૂચકાંકો (Monitoring Indicators)", "icon": "chart", "color": HIGHLIGHT_TEXT, "bullet": True},
}
FIELD_ORDER = ["OBJ", "STEPS", "WHO", "TIME", "IND"]

# ----------------------------------------------------------------------
# 1. WORD PARSER
# ----------------------------------------------------------------------
FIELD_PATTERNS = [
    ("OBJ", r"^ઉદ્દેશ્ય"),
    ("STEPS", r"^શું કરવું\??"),
    ("WHO", r"^જવાબદાર વ્યક્તિ"),
    ("TIME", r"^(સમયમર્યાદા|અમલીકરણનો સમય)"),
    ("IND", r"^મોનિટરિંગ સૂચકાંકો"),
]
MODULE_RE = re.compile(r"^(Module|મોડ્યુલ)\s*(\d+)\s*[:：]?\s*(.*)$", re.IGNORECASE)
SUBMODULE_RE = re.compile(r"^(\d+)\.(\d+)\s+(.*)$")


def strip_label(text, pattern):
    return re.sub(pattern + r"[^:：]*[:：]?", "", text, flags=re.IGNORECASE).strip()


def parse_word(docx_file):
    doc = docx.Document(docx_file)
    modules = {}
    cur_mod, cur_sub, cur_field = None, None, None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        m = MODULE_RE.match(text)
        if m:
            cur_mod = m.group(2)
            modules.setdefault(cur_mod, {"title": m.group(3).strip(), "subs": {}})
            cur_sub, cur_field = None, None
            continue

        s = SUBMODULE_RE.match(text)
        if s:
            main_num = s.group(1)
            cur_mod = main_num if cur_mod is None else cur_mod
            modules.setdefault(cur_mod, {"title": "", "subs": {}})
            cur_sub = f"{s.group(1)}.{s.group(2)}"
            modules[cur_mod]["subs"][cur_sub] = {
                "title": text, "OBJ": [], "STEPS": [], "WHO": [], "TIME": [], "IND": []
            }
            cur_field = None
            continue

        if not cur_sub:
            continue

        matched = False
        for key, pat in FIELD_PATTERNS:
            if re.match(pat, text):
                value = strip_label(text, pat)
                cur_field = key
                modules[cur_mod]["subs"][cur_sub][key] = [value] if value else []
                matched = True
                break
        if matched:
            continue

        if cur_field:
            modules[cur_mod]["subs"][cur_sub][cur_field].append(text)

    return modules


# ----------------------------------------------------------------------
# 2. TEXT MEASUREMENT  (drives dynamic sizing + pagination)
# ----------------------------------------------------------------------
_FONT_CACHE = {}


def _pil_font(size_pt, bold=False):
    key = (round(size_pt, 1), bold)
    if key not in _FONT_CACHE:
        path = FONT_TTF_BOLD if bold else FONT_TTF_REGULAR
        _FONT_CACHE[key] = ImageFont.truetype(path, max(1, int(size_pt / 72 * MEASURE_DPI)))
    return _FONT_CACHE[key]


def wrap_text(text, size_pt, max_width_in):
    font = _pil_font(size_pt)
    max_width_px = max_width_in * MEASURE_DPI
    words = text.split(" ")
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if font.getlength(trial) <= max_width_px or not cur:
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


# ---- icon glyphs (flat vector shapes, no badge circle — matches reference) ----
def icon_target(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    m = int(d * 0.62)
    add_oval(slide, x + int((d - m) / 2), y + int((d - m) / 2), m, WHITE)
    s = int(d * 0.30)
    add_oval(slide, x + int((d - s) / 2), y + int((d - s) / 2), s, color)


def icon_gear(slide, x, y, d, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.GEAR_9, x, y, d, d)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False


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
    body.fill.solid()
    body.fill.fore_color.rgb = color
    body.line.fill.background()
    body.shadow.inherit = False


def icon_clock(slide, x, y, d, color):
    add_oval(slide, x, y, d, color)
    cx, cy = x + d // 2, y + d // 2
    hand1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx - int(d * 0.035), cy - int(d * 0.33),
                                    int(d * 0.07), int(d * 0.33))
    hand1.fill.solid(); hand1.fill.fore_color.rgb = WHITE; hand1.line.fill.background(); hand1.shadow.inherit = False
    hand2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx - int(d * 0.035), cy - int(d * 0.24),
                                    int(d * 0.07), int(d * 0.24))
    hand2.fill.solid(); hand2.fill.fore_color.rgb = WHITE; hand2.line.fill.background(); hand2.shadow.inherit = False
    hand2.rotation = 90


def icon_chart(slide, x, y, d, color):
    bar_w = int(d * 0.20)
    gap = int(d * 0.12)
    heights = [int(d * 0.42), int(d * 0.68), int(d * 0.95)]
    total_w = 3 * bar_w + 2 * gap
    start_x = x + int((d - total_w) / 2)
    for i, h in enumerate(heights):
        bx = start_x + i * (bar_w + gap)
        by = y + (d - h)
        add_rect(slide, bx, by, bar_w, h, color)


ICON_FN = {"target": icon_target, "gear": icon_gear, "people": icon_people,
           "clock": icon_clock, "chart": icon_chart}


# ----------------------------------------------------------------------
# 4. CARD BUILDER
# ----------------------------------------------------------------------
def draw_card(slide, x, y, w, h, field_key, size_pt, items, highlight=False, label_override=None):
    meta = FIELD_META[field_key]
    bg = HIGHLIGHT_BG if highlight else CARD_WHITE
    add_rect(slide, x, y, w, h, bg, radius=0.06, shadow=True)

    icon_d = Inches(0.34)
    pad = CARD_PAD
    ICON_FN[meta["icon"]](slide, x + pad, y + pad, icon_d, meta["color"])
    label_x = x + pad + icon_d + Inches(0.12)
    label_w = w - pad - icon_d - Inches(0.12) - pad
    add_text(slide, label_x, y + pad - Inches(0.02), label_w, icon_d,
              label_override or meta["label"], 12, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    body_y = y + pad + icon_d + Inches(0.10)
    body_w = w - 2 * pad
    body_h = h - (body_y - y) - pad
    color = HIGHLIGHT_TEXT if highlight else TEXT_GRAY
    add_items(slide, x + pad, body_y, body_w, body_h, items, size_pt, color, meta["bullet"])


def field_card_height_in(field_key, items, size_pt, content_w_in):
    meta = FIELD_META[field_key]
    total = sum(item_height_in(it, size_pt, content_w_in, meta["bullet"]) for it in items)
    header_h_in = 0.34 + 0.10
    return header_h_in + total + (CARD_PAD / 914400) * 2


# ----------------------------------------------------------------------
# 5. SLIDE BUILDER  (with automatic pagination)
# ----------------------------------------------------------------------
def draw_header(slide, module_num, module_title, sub_title, continued, logos):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_rect(slide, 0, 0, SLIDE_W, HEADER_H, NAVY)

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
    add_text(slide, text_left, Inches(0.14), header_text_w, Inches(0.3),
              f"Module {module_num} \u2022 {module_title}".strip(" \u2022"),
              13, RGBColor(0xB9, 0xD3, 0xF2))
    title = sub_title + ("  \u2014  \u091a\u093e\u0932\u0941" if continued else "")
    add_text(slide, text_left, Inches(0.48), header_text_w, Inches(0.68),
              title, 23, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)


def build_submodule_slides(prs, module_num, module_title, sub_num, sub, logos):
    col_w_emu = (SLIDE_W - 2 * MARGIN_X - COL_GAP) // 2
    content_w_in = col_w_emu / 914400 - (CARD_PAD / 914400) * 2
    body_h_in = (SLIDE_H - BODY_TOP - BODY_BOTTOM_MARGIN) / 914400

    pieces = []  # (field_key, font_size, chunk_items, highlight, is_continuation)
    for key in FIELD_ORDER:
        items = [x for x in sub.get(key, []) if x.strip()]
        if not items:
            continue
        avail = body_h_in - 0.34 - 0.10 - (CARD_PAD / 914400) * 2
        font_size, chunks = chunk_items(items, content_w_in, avail, FIELD_META[key]["bullet"])
        for idx, chunk in enumerate(chunks):
            pieces.append((key, font_size, chunk, key == "IND", idx > 0))

    if not pieces:
        pieces = [("OBJ", 12, ["\u2014"], False, False)]

    slides_layout = []
    col_used = [0.0, 0.0]
    cur_slide_items = []
    for key, font_size, chunk, highlight, is_cont in pieces:
        h_in = field_card_height_in(key, chunk, font_size, content_w_in)
        target_col = 0 if col_used[0] <= col_used[1] else 1
        if col_used[target_col] + h_in > body_h_in and cur_slide_items:
            slides_layout.append(cur_slide_items)
            cur_slide_items = []
            col_used = [0.0, 0.0]
            target_col = 0
        cur_slide_items.append((target_col, key, font_size, chunk, highlight, is_cont, h_in))
        col_used[target_col] += h_in + (CARD_GAP_V / 914400)
    if cur_slide_items:
        slides_layout.append(cur_slide_items)

    for s_idx, layout in enumerate(slides_layout):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        draw_header(slide, module_num, module_title, sub.get("title_display", sub_num),
                    continued=(s_idx > 0), logos=logos)

        single_col = not any(item[0] == 1 for item in layout)
        full_w = SLIDE_W - 2 * MARGIN_X
        col_y = [BODY_TOP, BODY_TOP]
        col_x = [MARGIN_X, MARGIN_X + col_w_emu + COL_GAP]
        for target_col, key, font_size, chunk, highlight, is_cont, h_in in layout:
            w = full_w if single_col else col_w_emu
            x = col_x[0] if single_col else col_x[target_col]
            y = col_y[target_col]
            h = Emu(int(h_in * 914400))
            label_override = FIELD_META[key]["label"] + ("  (\u091a\u093e\u0932\u0941)" if is_cont else "")
            draw_card(slide, x, y, w, h, key, font_size, chunk,
                       highlight=highlight, label_override=label_override)
            col_y[target_col] = y + h + CARD_GAP_V


# ----------------------------------------------------------------------
# 6. STREAMLIT UI
# ----------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    left_logo_file = st.file_uploader("Left logo (e.g. AMC seal) — optional", type=["png", "jpg", "jpeg"])
with col2:
    right_logo_file = st.file_uploader("Right logo (e.g. NTEP logo) — optional", type=["png", "jpg", "jpeg"])

uploaded_docx = st.file_uploader("Upload Gujarati Word Document (.docx)", type=["docx"])

if uploaded_docx:
    modules = parse_word(uploaded_docx)
    for mod in modules.values():
        for sub in mod["subs"].values():
            sub["title_display"] = sub["title"]

    total_subs = sum(len(m["subs"]) for m in modules.values())
    with st.expander(f"🔍 Parsed structure — {len(modules)} modules, {total_subs} sub-modules"):
        st.write(modules)

    if total_subs == 0:
        st.error("No sub-modules (like 1.1, 1.2) were detected. Check your Word doc headings.")
    elif st.button("Generate Presentation", type="primary"):
        with st.spinner("Measuring content and designing slides..."):
            prs = Presentation()
            prs.slide_width = SLIDE_W
            prs.slide_height = SLIDE_H
            logos = (left_logo_file, right_logo_file)
            for mod_num, mod in sorted(modules.items(), key=lambda x: int(x[0])):
                for sub_num, sub in sorted(mod["subs"].items(),
                                            key=lambda x: tuple(map(int, x[0].split(".")))):
                    build_submodule_slides(prs, mod_num, mod["title"], sub_num, sub, logos)

            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)

        st.success(f"✅ Generated {len(prs.slides._sldIdLst)} slides for {total_subs} sub-modules.")
        st.download_button(
            "📄 Download Presentation",
            data=ppt_io,
            file_name="AMC_NTEP_AutoDesigned.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
