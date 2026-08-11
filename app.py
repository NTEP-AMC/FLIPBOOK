import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import io
import re
import copy

st.set_page_config(page_title="AMC NTEP - Slide Auto-Designer", layout="wide")
st.title("AMC NTEP - Sub-Module Slide Auto-Designer")
st.caption(
    "Upload your Word doc. This builds ONE fully-designed slide per sub-module "
    "(1.1, 1.2, 5.1 ...) automatically — no PPTX template or {{tags}} needed."
)

# ----------------------------------------------------------------------
# COLOR PALETTE  (edit these 5 lines to re-theme the whole deck)
# ----------------------------------------------------------------------
NAVY = RGBColor(0x14, 0x2C, 0x5C)
NAVY_LIGHT = RGBColor(0x28, 0x49, 0x86)
TEAL = RGBColor(0x1F, 0xA6, 0x8C)
BG_LIGHT = RGBColor(0xEE, 0xF5, 0xFB)
CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x22, 0x2B, 0x3A)
TEXT_GRAY = RGBColor(0x5B, 0x66, 0x77)
HIGHLIGHT_BG = RGBColor(0xC9, 0xF3, 0xEE)
HIGHLIGHT_TEXT = RGBColor(0x0B, 0x6E, 0x5D)

FONT = "Noto Sans Gujarati"  # swap for your deployed font; falls back gracefully

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FIELD_ICONS = {
    "OBJ": ("\u25CE", TEAL),        # bullseye  – Objective
    "STEPS": ("\u2699", NAVY_LIGHT),  # gear      – What to do
    "WHO": ("\u25A4", NAVY_LIGHT),   # person-ish – Responsible
    "TIME": ("\u23F1", TEAL),        # stopwatch – Timeline
    "IND": ("\u2B06", HIGHLIGHT_TEXT),  # up-arrow – Monitoring indicators
}
FIELD_LABELS = {
    "OBJ": "ઉદ્દેશ્ય (Objective)",
    "STEPS": "શું કરવું? (What to do)",
    "WHO": "જવાબદાર વ્યક્તિ (Responsible)",
    "TIME": "સમયમર્યાદા (Timeline)",
    "IND": "મોનિટરિંગ સૂચકાંકો (Monitoring Indicators)",
}

# ----------------------------------------------------------------------
# 1. WORD PARSER  -> nested {module_num: {"title":..., "subs": {sub_num: {...}}}}
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
                "title": text, "OBJ": "", "STEPS": [], "WHO": "", "TIME": "", "IND": ""
            }
            cur_field = None
            continue

        if not cur_sub:
            continue

        matched_field = None
        for key, pat in FIELD_PATTERNS:
            if re.match(pat, text):
                matched_field = key
                value = strip_label(text, pat)
                cur_field = key
                if key == "STEPS":
                    modules[cur_mod]["subs"][cur_sub]["STEPS"] = [value] if value else []
                else:
                    modules[cur_mod]["subs"][cur_sub][key] = value
                break
        if matched_field:
            continue

        # continuation line under the current field
        if cur_field == "STEPS":
            modules[cur_mod]["subs"][cur_sub]["STEPS"].append(text)
        elif cur_field:
            modules[cur_mod]["subs"][cur_sub][cur_field] += ("\n" + text)

    return modules


# ----------------------------------------------------------------------
# 2. DRAWING HELPERS
# ----------------------------------------------------------------------
def set_shadow(shape, blur=Pt(10), dist=Pt(3), alpha=78):
    """Soft drop shadow so white cards lift off the light background."""
    sp = shape._element.spPr
    existing = sp.find(qn("a:effectLst"))
    if existing is not None:
        sp.remove(existing)
    effect = sp.makeelement(qn("a:effectLst"), {})
    shadow = effect.makeelement(qn("a:outerShdw"), {
        "blurRad": str(blur), "dist": str(dist), "dir": "5400000",
        "rotWithShape": "0",
    })
    clr = shadow.makeelement(qn("a:srgbClr"), {"val": "0F1F3D"})
    alpha_el = clr.makeelement(qn("a:alpha"), {"val": str(alpha * 1000)})
    clr.append(alpha_el)
    shadow.append(clr)
    effect.append(shadow)
    sp.append(effect)


def add_rect(slide, x, y, w, h, fill, line=None, shadow=False, radius=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if radius is not None:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    if shadow:
        set_shadow(shp)
    return shp


def add_text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, font=FONT, line_spacing=1.15, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = color
    return tb


def add_bullets(slide, x, y, w, h, items, size, color, font=FONT, gap_pt=4, line_spacing=1.1):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        p.space_after = Pt(gap_pt)
        r = p.add_run()
        r.text = f"\u2022  {item}"
        r.font.size = Pt(size)
        r.font.name = font
        r.font.color.rgb = color
    return tb


def add_icon_circle(slide, cx, cy, d, glyph, glyph_color, ring_color):
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, d, d)
    circ.fill.solid()
    circ.fill.fore_color.rgb = ring_color
    circ.line.fill.background()
    circ.shadow.inherit = False
    tf = circ.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = glyph
    r.font.size = Pt(int(d / Emu(1) / 914400 * 26))
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.bold = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return circ


def field_card(slide, x, y, w, h, field_key, value, highlight=False):
    bg = HIGHLIGHT_BG if highlight else CARD_WHITE
    card = add_rect(slide, x, y, w, h, bg, radius=0.07, shadow=True)
    pad = Inches(0.18)
    icon_d = Inches(0.42)
    glyph, ring = FIELD_ICONS[field_key]
    add_icon_circle(slide, x + pad, y + pad, icon_d, glyph, RGBColor(0xFF, 0xFF, 0xFF), ring)
    label_x = x + pad + icon_d + Inches(0.12)
    label_w = w - pad - icon_d - Inches(0.12) - pad
    add_text(slide, label_x, y + pad - Inches(0.02), label_w, icon_d,
              FIELD_LABELS[field_key], 12.5, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    body_y = y + pad + icon_d + Inches(0.10)
    body_w = w - 2 * pad
    body_h = h - (body_y - y) - pad
    text_color = HIGHLIGHT_TEXT if highlight else TEXT_GRAY
    if field_key == "STEPS" and isinstance(value, list):
        add_bullets(slide, x + pad, body_y, body_w, body_h, value, 12.5, text_color)
    else:
        add_text(slide, x + pad, body_y, body_w, body_h, value or "—", 12.5,
                  text_color, bold=highlight)
    return card


def add_decorative_circles(slide):
    """Very light background circles for depth (not thin accent stripes)."""
    d1 = Inches(3.2)
    c1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, SLIDE_W - Inches(1.4), SLIDE_H - Inches(1.6), d1, d1)
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(0xE2, 0xEE, 0xF8)
    c1.line.fill.background()
    c1.shadow.inherit = False
    slide.shapes._spTree.remove(c1._element)
    slide.shapes._spTree.insert(2, c1._element)  # push behind everything drawn after


def build_submodule_slide(prs, module_num, module_title, sub_num, sub):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    add_decorative_circles(slide)

    # ---- Header bar ----
    header_h = Inches(1.25)
    add_rect(slide, 0, 0, SLIDE_W, header_h, NAVY)
    add_text(slide, Inches(0.55), Inches(0.14), Inches(8), Inches(0.3),
              f"Module {module_num} \u2022 {module_title}".strip(" \u2022"),
              13, RGBColor(0xB9, 0xD3, 0xF2), bold=False)
    title_text = sub.get("title", f"{sub_num}")
    add_text(slide, Inches(0.55), Inches(0.48), Inches(11.6), Inches(0.68),
              title_text, 24, RGBColor(0xFF, 0xFF, 0xFF), bold=True,
              anchor=MSO_ANCHOR.MIDDLE)

    # ---- Body: left big card (Objective + Steps), right stacked cards ----
    top = header_h + Inches(0.35)
    left_x = Inches(0.55)
    left_w = Inches(7.55)
    gap = Inches(0.3)
    right_x = left_x + left_w + gap
    right_w = SLIDE_W - right_x - Inches(0.55)
    bottom_margin = Inches(0.4)
    body_h = SLIDE_H - top - bottom_margin

    obj_h = Inches(1.7)
    steps_h = body_h - obj_h - Inches(0.25)
    field_card(slide, left_x, top, left_w, obj_h, "OBJ", sub.get("OBJ", ""))
    field_card(slide, left_x, top + obj_h + Inches(0.25), left_w, steps_h,
               "STEPS", sub.get("STEPS", []))

    n_right = 2 + (1 if sub.get("TIME") else 0)
    r_gap = Inches(0.25)
    r_h = (body_h - r_gap * (n_right - 1)) / n_right
    ry = top
    field_card(slide, right_x, ry, right_w, r_h, "WHO", sub.get("WHO", ""))
    ry += r_h + r_gap
    if sub.get("TIME"):
        field_card(slide, right_x, ry, right_w, r_h, "TIME", sub.get("TIME", ""))
        ry += r_h + r_gap
    field_card(slide, right_x, ry, right_w, r_h, "IND", sub.get("IND", ""), highlight=True)

    return slide


# ----------------------------------------------------------------------
# 3. STREAMLIT UI
# ----------------------------------------------------------------------
uploaded_docx = st.file_uploader("Upload Gujarati Word Document (.docx)", type=["docx"])

if uploaded_docx:
    modules = parse_word(uploaded_docx)
    total_subs = sum(len(m["subs"]) for m in modules.values())
    with st.expander(f"🔍 Parsed structure — {len(modules)} modules, {total_subs} sub-modules"):
        st.write(modules)

    if total_subs == 0:
        st.error(
            "No sub-modules (like 1.1, 1.2) were detected. Check that your Word doc "
            "headings match the expected pattern, or share a sample so the parser regex can be adjusted."
        )
    elif st.button("Generate Presentation", type="primary"):
        with st.spinner("Designing one slide per sub-module..."):
            prs = Presentation()
            prs.slide_width = SLIDE_W
            prs.slide_height = SLIDE_H
            for mod_num, mod in sorted(modules.items(), key=lambda x: int(x[0])):
                for sub_num, sub in sorted(mod["subs"].items(),
                                            key=lambda x: tuple(map(int, x[0].split(".")))):
                    build_submodule_slide(prs, mod_num, mod["title"], sub_num, sub)

            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)

        st.success(f"✅ Generated {total_subs} designed slides.")
        st.download_button(
            "📄 Download Presentation",
            data=ppt_io,
            file_name="AMC_NTEP_AutoDesigned.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
