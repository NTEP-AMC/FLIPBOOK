import streamlit as st
import io
import os
import re
import math
from pathlib import Path

import docx
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE


# ============================================================
# AMC NTEP PROFESSIONAL PPT AUTO-DESIGNER
# ============================================================

st.set_page_config(
    page_title="AMC NTEP Professional PPT Designer",
    page_icon="📊",
    layout="wide",
)

# -----------------------------
# Theme
# -----------------------------
NAVY = RGBColor(0x0D, 0x35, 0x68)
NAVY2 = RGBColor(0x16, 0x4A, 0x7D)
BLUE = RGBColor(0x1F, 0x68, 0xB2)
TEAL = RGBColor(0x00, 0x91, 0x87)
GREEN = RGBColor(0x1F, 0xA6, 0x8C)
CYAN = RGBColor(0x00, 0xA6, 0xC7)
ORANGE = RGBColor(0xF2, 0x8C, 0x28)
RED = RGBColor(0xD9, 0x3B, 0x30)
PURPLE = RGBColor(0x6D, 0x4C, 0xB8)
BG = RGBColor(0xF3, 0xF7, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x2D, 0x3D)
GRAY = RGBColor(0x5B, 0x67, 0x73)
LIGHT_BLUE = RGBColor(0xE5, 0xF1, 0xFA)
LIGHT_GREEN = RGBColor(0xE4, 0xF5, 0xF1)
LIGHT_ORANGE = RGBColor(0xFF, 0xF0, 0xDF)
LIGHT_RED = RGBColor(0xFD, 0xE9, 0xE7)
LIGHT_PURPLE = RGBColor(0xEF, 0xE9, 0xFA)

FONT = "Noto Sans Gujarati"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

MARGIN_X = Inches(0.48)
BODY_TOP = Inches(1.40)
BODY_BOTTOM = Inches(0.35)
GAP = Inches(0.22)


# ============================================================
# Word parsing
# ============================================================

MODULE_RE = re.compile(
    r"^(?:Module|MODULE|મોડ્યુલ)\s*(\d+)\s*[:：\-]?\s*(.*)$",
    re.IGNORECASE,
)

SUBMODULE_RE = re.compile(
    r"^(\d+)\.(\d+)\s*[:：\-]?\s*(.*)$"
)

FIELD_PATTERNS = [
    ("OBJ", r"^(?:ઉદ્દેશ્ય|objective)\b"),
    ("STEPS", r"^(?:શું કરવું|what to do)\b"),
    ("WHO", r"^(?:જવાબદાર વ્યક્તિ|જવાબદાર વ્યક્તિઓ|responsible person|responsible)\b"),
    ("TIME", r"^(?:સમયમર્યાદા|અમલીકરણનો સમય|timeline|time)\b"),
    ("IND", r"^(?:મોનિટરિંગ સૂચકાંકો|monitoring indicators|monitoring)\b"),
    ("WHY", r"^(?:શા માટે મહત્વપૂર્ણ|શા માટે|why important|importance)\b"),
    ("RISK", r"^(?:જોખમ|risks|risk if not done)\b"),
    ("DOC", r"^(?:દસ્તાવેજીકરણ|documentation)\b"),
    ("TRIGGER", r"^(?:ટ્રિગર|trigger)\b"),
]


def clean_text(text):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def strip_field_label(text, pattern):
    return re.sub(pattern + r"\s*[:：\-]?\s*", "", text, flags=re.I).strip()


def parse_word(docx_file):
    doc = docx.Document(docx_file)

    modules = {}
    current_module = None
    current_sub = None
    current_field = None

    for p in doc.paragraphs:
        text = clean_text(p.text)
        if not text:
            continue

        m = MODULE_RE.match(text)
        if m:
            current_module = m.group(1)
            modules.setdefault(
                current_module,
                {"title": m.group(2).strip(), "subs": {}},
            )
            current_sub = None
            current_field = None
            continue

        s = SUBMODULE_RE.match(text)
        if s:
            mod_no = s.group(1)
            sub_no = f"{s.group(1)}.{s.group(2)}"
            current_module = mod_no

            modules.setdefault(
                current_module,
                {"title": "", "subs": {}},
            )

            title = s.group(3).strip()
            modules[current_module]["subs"][sub_no] = {
                "number": sub_no,
                "title": title if title else text,
                "OBJ": [],
                "STEPS": [],
                "WHO": [],
                "TIME": [],
                "IND": [],
                "WHY": [],
                "RISK": [],
                "DOC": [],
                "TRIGGER": [],
                "OTHER": [],
            }

            current_sub = sub_no
            current_field = None
            continue

        if current_sub is None:
            continue

        matched = False

        for key, pattern in FIELD_PATTERNS:
            if re.match(pattern, text, flags=re.I):
                value = strip_field_label(text, pattern)
                current_field = key
                if value:
                    modules[current_module]["subs"][current_sub][key].append(value)
                matched = True
                break

        if matched:
            continue

        if current_field:
            modules[current_module]["subs"][current_sub][current_field].append(text)
        else:
            modules[current_module]["subs"][current_sub]["OTHER"].append(text)

    # If Word uses headings without explicit "Module", infer module title
    for mod_no, mod in modules.items():
        if not mod["title"]:
            mod["title"] = f"Module {mod_no}"

    return modules


# ============================================================
# Text utilities
# ============================================================

def all_text(items):
    return " ".join([x for x in items if x]).strip()


def wrap_text(text, max_chars):
    words = text.split()
    lines = []
    current = ""
    for word in words:
        test = f"{current} {word}".strip()
        if len(test) <= max_chars:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def add_text(
    slide, x, y, w, h, text,
    size=18,
    color=DARK,
    bold=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font=FONT,
):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)

    paragraphs = str(text).split("\n")

    for i, line in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.12
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color

    return box


def add_bullets(
    slide, x, y, w, h, items,
    size=15,
    color=GRAY,
    bullet=True,
    gap=4,
):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)

    for i, item in enumerate(items):
        if not item:
            continue
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        p.line_spacing = 1.08

        r = p.add_run()
        r.text = ("• " if bullet else "") + item
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.color.rgb = color

    return box


# ============================================================
# Shapes
# ============================================================

def fill_shape(shape, fill, line=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line


def rounded_rect(slide, x, y, w, h, fill=WHITE, line=None, radius=True):
    typ = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(typ, x, y, w, h)
    fill_shape(shp, fill, line)
    return shp


def circle(slide, x, y, d, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    fill_shape(shp, fill)
    return shp


def line(slide, x1, y1, x2, y2, color=BLUE, width=2):
    shp = slide.shapes.add_connector(
        1, x1, y1, x2, y2
    )
    shp.line.color.rgb = color
    shp.line.width = Pt(width)
    return shp


def add_shadow(shape):
    # Safe/simple PowerPoint shadow effect.
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_icon_badge(slide, x, y, d, symbol, fill=TEAL):
    circle(slide, x, y, d, fill)
    add_text(
        slide, x, y + Inches(0.01), d, d,
        symbol, size=20, color=WHITE,
        bold=True, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        font="Arial",
    )


# ============================================================
# Asset handling
# ============================================================

def save_uploaded_assets(uploaded_files, asset_dir="assets"):
    path = Path(asset_dir)
    path.mkdir(parents=True, exist_ok=True)

    saved = []

    for f in uploaded_files:
        target = path / f.name
        target.write_bytes(f.getbuffer())
        saved.append(str(target))

    return saved


def load_image_bytes(file_obj):
    if file_obj is None:
        return None
    return io.BytesIO(file_obj.getvalue())


def add_picture_contain(slide, image_source, x, y, w, h):
    if image_source is None:
        return False

    try:
        if hasattr(image_source, "seek"):
            image_source.seek(0)
        slide.shapes.add_picture(image_source, x, y, width=w, height=h)
        return True
    except Exception:
        return False


def find_asset(asset_files, keywords):
    """
    Match an uploaded filename to content keywords.
    Example:
      CBNAAT.png -> ["cbnaat", "naat"]
    """
    if not asset_files:
        return None

    for file in asset_files:
        name = Path(file).stem.lower()
        for key in keywords:
            if key.lower() in name:
                return file

    return None


def keyword_group_for_submodule(sub):
    text = (
        sub["number"] + " " +
        sub["title"] + " " +
        all_text(sub.get("OBJ", [])) + " " +
        all_text(sub.get("STEPS", [])) + " " +
        all_text(sub.get("OTHER", []))
    ).lower()

    groups = []

    if any(k in text for k in ["cbnaat", "cb-gaat", "cartridge", "gene xpert"]):
        groups.append(["cbnaat", "gene", "cartridge"])

    if any(k in text for k in ["truenat", "true nat"]):
        groups.append(["truenat"])

    if any(k in text for k in ["microscope", "microscopy", "માઇક્રોસ્કોપ"]):
        groups.append(["microscope", "microscopy"])

    if any(k in text for k in ["migration", "સ્થળાંતર", "transfer out", "transfer in"]):
        groups.append(["migration", "map", "transfer"])

    if any(k in text for k in ["hiv", "diabetes", "diabetes mellitus"]):
        groups.append(["hiv", "diabetes", "screening"])

    if any(k in text for k in ["notification", "nikshay", "નિક્ષય"]):
        groups.append(["nikshay", "notification"])

    return groups


# ============================================================
# Header/footer
# ============================================================

def add_background(slide):
    rounded_rect(
        slide, 0, 0, SLIDE_W, SLIDE_H,
        BG, radius=False
    )

    # Decorative circles
    circle(slide, Inches(11.9), Inches(-0.55), Inches(1.45), LIGHT_BLUE)
    circle(slide, Inches(12.35), Inches(6.6), Inches(1.2), LIGHT_GREEN)


def add_logo(slide, file_obj, x, y, w, h):
    if file_obj is None:
        return
    try:
        file_obj.seek(0)
        slide.shapes.add_picture(file_obj, x, y, width=w, height=h)
    except Exception:
        pass


def add_header(
    slide,
    module_no,
    module_title,
    sub_title=None,
    left_logo=None,
    right_logo=None,
    section="Public Health Actions",
):
    add_background(slide)

    # Header band
    rounded_rect(
        slide, 0, 0, SLIDE_W, Inches(1.08),
        NAVY, radius=False
    )

    if left_logo:
        add_logo(
            slide, left_logo,
            Inches(0.22), Inches(0.13),
            Inches(0.78), Inches(0.78)
        )

    if right_logo:
        add_logo(
            slide, right_logo,
            Inches(12.25), Inches(0.13),
            Inches(0.78), Inches(0.78)
        )

    add_text(
        slide,
        Inches(1.15), Inches(0.10),
        Inches(10.7), Inches(0.30),
        f"{section}  •  Module {module_no}",
        size=12,
        color=RGBColor(0xB9, 0xD8, 0xF5),
        bold=True,
    )

    title = sub_title if sub_title else module_title

    add_text(
        slide,
        Inches(1.15), Inches(0.40),
        Inches(10.7), Inches(0.52),
        title,
        size=22,
        color=WHITE,
        bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def add_footer(slide, slide_no):
    line(
        slide,
        Inches(0.48), Inches(7.12),
        Inches(12.85), Inches(7.12),
        RGBColor(0xD4, 0xDF, 0xE8),
        1,
    )

    add_text(
        slide,
        Inches(0.50), Inches(7.15),
        Inches(9), Inches(0.20),
        "Ahmedabad Municipal Corporation • NTEP • Public Health Actions",
        size=7.5,
        color=GRAY,
    )

    add_text(
        slide,
        Inches(11.7), Inches(7.13),
        Inches(1.1), Inches(0.22),
        str(slide_no),
        size=9,
        color=NAVY,
        bold=True,
        align=PP_ALIGN.RIGHT,
    )


# ============================================================
# Visual cards
# ============================================================

FIELD_STYLE = {
    "OBJ": ("ઉદ્દેશ્ય (Objective)", TEAL, LIGHT_GREEN, "◎"),
    "STEPS": ("શું કરવું? (What to do)", BLUE, LIGHT_BLUE, "⚙"),
    "WHO": ("જવાબદાર વ્યક્તિ (Responsible)", PURPLE, LIGHT_PURPLE, "●"),
    "TIME": ("સમયમર્યાદા (Timeline)", GREEN, LIGHT_GREEN, "◷"),
    "IND": ("મોનિટરિંગ સૂચકાંકો", TEAL, LIGHT_GREEN, "▥"),
    "WHY": ("શા માટે મહત્વપૂર્ણ?", ORANGE, LIGHT_ORANGE, "!"),
    "RISK": ("જોખમ / Risks", RED, LIGHT_RED, "!"),
    "DOC": ("દસ્તાવેજીકરણ", NAVY2, LIGHT_BLUE, "▤"),
    "TRIGGER": ("Trigger", ORANGE, LIGHT_ORANGE, "↗"),
}


def draw_info_card(
    slide, x, y, w, h,
    field_key,
    items,
    photo=None,
    number=None,
):
    label, accent, light, icon = FIELD_STYLE.get(
        field_key,
        ("Key Information", BLUE, LIGHT_BLUE, "•")
    )

    card = rounded_rect(
        slide, x, y, w, h,
        WHITE,
        line=RGBColor(0xD8, 0xE1, 0xE8),
    )
    add_shadow(card)

    # accent strip
    rounded_rect(
        slide, x, y, Inches(0.07), h,
        accent, radius=False
    )

    # icon
    add_icon_badge(
        slide,
        x + Inches(0.18),
        y + Inches(0.16),
        Inches(0.43),
        icon,
        accent,
    )

    add_text(
        slide,
        x + Inches(0.72),
        y + Inches(0.13),
        w - Inches(0.9),
        Inches(0.35),
        label,
        size=12,
        color=NAVY,
        bold=True,
    )

    body_x = x + Inches(0.22)
    body_y = y + Inches(0.64)
    body_w = w - Inches(0.44)
    body_h = h - Inches(0.78)

    if photo:
        photo_w = min(Inches(2.05), body_w * 0.34)
        try:
            photo.seek(0)
            slide.shapes.add_picture(
                photo,
                body_x,
                body_y,
                width=photo_w,
                height=body_h,
            )
            body_x += photo_w + Inches(0.15)
            body_w -= photo_w + Inches(0.15)
        except Exception:
            pass

    if number is not None:
        circle(
            slide,
            body_x,
            body_y,
            Inches(0.34),
            accent,
        )
        add_text(
            slide,
            body_x,
            body_y,
            Inches(0.34),
            Inches(0.34),
            str(number),
            size=10,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
            font="Arial",
        )
        body_x += Inches(0.48)
        body_w -= Inches(0.48)

    if field_key == "STEPS":
        add_bullets(
            slide,
            body_x,
            body_y,
            body_w,
            body_h,
            items,
            size=13,
            color=GRAY,
            bullet=True,
        )
    else:
        add_bullets(
            slide,
            body_x,
            body_y,
            body_w,
            body_h,
            items,
            size=13,
            color=GRAY,
            bullet=False,
        )


# ============================================================
# Layouts
# ============================================================

def build_cover(prs, left_logo, right_logo, title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)

    rounded_rect(
        slide, 0, 0, SLIDE_W, SLIDE_H,
        NAVY, radius=False
    )

    # Decorative graphic
    circle(
        slide, Inches(9.6), Inches(-0.5),
        Inches(4.5), NAVY2
    )
    circle(
        slide, Inches(10.5), Inches(4.8),
        Inches(3.8), TEAL
    )

    if left_logo:
        add_logo(
            slide, left_logo,
            Inches(0.55), Inches(0.42),
            Inches(1.25), Inches(1.25)
        )

    if right_logo:
        add_logo(
            slide, right_logo,
            Inches(11.45), Inches(0.42),
            Inches(1.25), Inches(1.25)
        )

    add_text(
        slide,
        Inches(1.1), Inches(2.05),
        Inches(11.1), Inches(0.55),
        "રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP)",
        size=18,
        color=RGBColor(0xB8, 0xDD, 0xF8),
        bold=True,
        align=PP_ALIGN.CENTER,
    )

    add_text(
        slide,
        Inches(1.0), Inches(2.75),
        Inches(11.3), Inches(1.4),
        title,
        size=32,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    add_text(
        slide,
        Inches(2.0), Inches(4.35),
        Inches(9.3), Inches(0.5),
        "Public Health Actions • Ahmedabad Municipal Corporation",
        size=17,
        color=RGBColor(0xD5, 0xE9, 0xF8),
        align=PP_ALIGN.CENTER,
    )

    add_text(
        slide,
        Inches(2.2), Inches(5.2),
        Inches(8.9), Inches(0.8),
        "Training & Implementation Guide",
        size=18,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )

    add_footer(slide, 1)


def build_intro_slide(prs, title, items, slide_no, left_logo, right_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        "",
        title,
        title,
        left_logo,
        right_logo,
        section="Introduction",
    )

    # large objective-style visual
    add_icon_badge(
        slide,
        Inches(0.65), Inches(1.65),
        Inches(1.0), "✓", TEAL
    )

    add_text(
        slide,
        Inches(1.9), Inches(1.55),
        Inches(10.5), Inches(0.65),
        title,
        size=25,
        color=NAVY,
        bold=True,
    )

    add_bullets(
        slide,
        Inches(1.9), Inches(2.25),
        Inches(10.2), Inches(3.8),
        items,
        size=17,
        color=DARK,
        bullet=True,
        gap=10,
    )

    add_footer(slide, slide_no)


def build_index_slide(prs, modules, slide_no, left_logo, right_logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide, "", "Module Index",
        "મોડ્યુલ ઇન્ડેક્સ / Module Index",
        left_logo, right_logo,
        section="Contents",
    )

    all_subs = []
    for mod_no, mod in sorted(modules.items(), key=lambda z: int(z[0])):
        all_subs.extend(
            [(mod_no, mod["title"], sub_no, sub["title"])
             for sub_no, sub in sorted(
                 mod["subs"].items(),
                 key=lambda z: tuple(map(int, z[0].split(".")))
             )]
        )

    # Index is generated from estimated final order.
    rows = all_subs[:24]

    cols = 2
    rows_per_col = math.ceil(len(rows) / cols)

    for col in range(cols):
        subset = rows[col * rows_per_col:(col + 1) * rows_per_col]
        x = Inches(0.65 + col * 6.15)
        y = Inches(1.55)

        for i, (mod_no, mod_title, sub_no, sub_title) in enumerate(subset):
            yy = y + Inches(0.38) * i

            rounded_rect(
                slide,
                x, yy,
                Inches(5.65), Inches(0.31),
                WHITE,
                line=RGBColor(0xD8, 0xE1, 0xE8),
            )

            circle(
                slide,
                x + Inches(0.06), yy + Inches(0.04),
                Inches(0.23), NAVY2
            )

            add_text(
                slide,
                x + Inches(0.06), yy + Inches(0.04),
                Inches(0.23), Inches(0.23),
                sub_no,
                size=6.5,
                color=WHITE,
                bold=True,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
                font="Arial",
            )

            display = sub_title[:54] + ("…" if len(sub_title) > 54 else "")
            add_text(
                slide,
                x + Inches(0.37), yy + Inches(0.02),
                Inches(4.55), Inches(0.26),
                display,
                size=8.5,
                color=DARK,
                bold=True,
                anchor=MSO_ANCHOR.MIDDLE,
            )

            # Page number placeholder
            add_text(
                slide,
                x + Inches(5.02), yy + Inches(0.02),
                Inches(0.48), Inches(0.26),
                "—",
                size=8,
                color=TEAL,
                bold=True,
                align=PP_ALIGN.RIGHT,
                anchor=MSO_ANCHOR.MIDDLE,
            )

    add_footer(slide, slide_no)


def build_module_overview(
    prs,
    module_no,
    module_title,
    subs,
    slide_no,
    left_logo,
    right_logo,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_header(
        slide,
        module_no,
        module_title,
        f"Module {module_no}: {module_title}",
        left_logo,
        right_logo,
    )

    add_text(
        slide,
        Inches(0.65), Inches(1.35),
        Inches(11.9), Inches(0.4),
        "Module Overview",
        size=19,
        color=NAVY,
        bold=True,
    )

    sub_items = list(subs.items())

    for i, (sub_no, sub) in enumerate(sub_items[:12]):
        col = i % 3
        row = i // 3

        x = Inches(0.65) + col * Inches(4.05)
        y = Inches(1.95) + row * Inches(1.22)

        rounded_rect(
            slide,
            x, y,
            Inches(3.72), Inches(0.96),
            WHITE,
            line=RGBColor(0xD8, 0xE1, 0xE8),
        )

        circle(
            slide,
            x + Inches(0.15), y + Inches(0.18),
            Inches(0.55),
            [TEAL, BLUE, PURPLE][col],
        )

        add_text(
            slide,
            x + Inches(0.15), y + Inches(0.18),
            Inches(0.55), Inches(0.55),
            sub_no,
            size=9,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
            font="Arial",
        )

        title = sub["title"]
        add_text(
            slide,
            x + Inches(0.85), y + Inches(0.14),
            Inches(2.68), Inches(0.66),
            title,
            size=11,
            color=DARK,
            bold=True,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    add_footer(slide, slide_no)


def build_workflow_slide(
    prs,
    module_no,
    module_title,
    sub,
    slide_no,
    left_logo,
    right_logo,
    photo=None,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_header(
        slide,
        module_no,
        module_title,
        f"{sub['number']}  {sub['title']}",
        left_logo,
        right_logo,
    )

    steps = sub.get("STEPS", []) or sub.get("OTHER", [])
    if not steps:
        steps = ["મુખ્ય પ્રક્રિયાની માહિતી ઉપલબ્ધ નથી."]

    # Keep workflow readable
    steps = steps[:8]

    x0 = Inches(0.72)
    y0 = Inches(1.72)
    card_w = Inches(2.72)
    card_h = Inches(1.12)

    for i, step in enumerate(steps):
        row = i // 4
        col = i % 4

        x = x0 + col * Inches(3.05)
        y = y0 + row * Inches(1.62)

        rounded_rect(
            slide, x, y,
            card_w, card_h,
            WHITE,
            line=RGBColor(0xD8, 0xE1, 0xE8),
        )

        circle(
            slide,
            x + Inches(0.14),
            y + Inches(0.16),
            Inches(0.47),
            TEAL if i % 2 == 0 else BLUE,
        )

        add_text(
            slide,
            x + Inches(0.14),
            y + Inches(0.16),
            Inches(0.47),
            Inches(0.47),
            str(i + 1),
            size=9,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
            font="Arial",
        )

        lines = wrap_text(step, 48)
        add_text(
            slide,
            x + Inches(0.75),
            y + Inches(0.12),
            Inches(1.80),
            Inches(0.84),
            "\n".join(lines[:5]),
            size=11.5,
            color=DARK,
            bold=True,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        if i < len(steps) - 1 and col < 3:
            line(
                slide,
                x + card_w,
                y + card_h / 2,
                x + card_w + Inches(0.30),
                y + card_h / 2,
                TEAL,
                2,
            )

    # Photo panel when available
    if photo:
        rounded_rect(
            slide,
            Inches(10.0), Inches(5.15),
            Inches(2.45), Inches(1.45),
            WHITE,
            line=RGBColor(0xD8, 0xE1, 0xE8),
        )
        try:
            photo.seek(0)
            slide.shapes.add_picture(
                photo,
                Inches(10.12), Inches(5.27),
                width=Inches(2.2),
                height=Inches(1.18),
            )
        except Exception:
            pass

    add_footer(slide, slide_no)


def build_content_slide(
    prs,
    module_no,
    module_title,
    sub,
    slide_no,
    left_logo,
    right_logo,
    asset_files=None,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_header(
        slide,
        module_no,
        module_title,
        f"{sub['number']}  {sub['title']}",
        left_logo,
        right_logo,
    )

    # Detect photo
    photo = None
    groups = keyword_group_for_submodule(sub)

    for group in groups:
        asset = find_asset(asset_files, group)
        if asset:
            try:
                photo = open(asset, "rb")
                break
            except Exception:
                pass

    available = []

    for key in ["OBJ", "STEPS", "WHO", "TIME", "IND", "WHY", "RISK", "DOC", "TRIGGER"]:
        items = [x for x in sub.get(key, []) if x]
        if items:
            available.append((key, items))

    if not available:
        available = [("OBJ", ["માહિતી ઉપલબ્ધ નથી."])]

    # Special workflow layout for many steps
    if len(sub.get("STEPS", [])) >= 4:
        build_workflow_slide(
            prs,
            module_no,
            module_title,
            sub,
            slide_no,
            left_logo,
            right_logo,
            photo,
        )
        # Remove the blank slide created above.
        xml_slides = prs.slides._sldIdLst
        xml_slides.remove(xml_slides[-2])
        return

    # 2-column professional card layout
    left = []
    right = []

    for i, item in enumerate(available):
        (left if i % 2 == 0 else right).append(item)

    def draw_column(items, x):
        y = Inches(1.55)
        width = Inches(6.08)

        if not items:
            return

        remaining = 5.30

        for key, values in items:
            # Dynamic card height
            chars = max(45, int(width / Inches(0.12)))
            lines = 0
            for value in values:
                lines += max(1, math.ceil(len(value) / chars))

            h = max(
                0.98,
                min(2.35, 0.55 + lines * 0.27 + len(values) * 0.10)
            )

            if remaining < h:
                h = max(0.90, remaining)

            draw_info_card(
                slide,
                x,
                y,
                width,
                Inches(h),
                key,
                values[:8],
                photo if (key == "OBJ" and photo) else None,
            )

            y += Inches(h) + GAP
            remaining -= h + 0.22

            if remaining < 0.65:
                break

    draw_column(left, Inches(0.52))
    draw_column(right, Inches(6.73))

    add_footer(slide, slide_no)


def build_summary_slide(
    prs,
    module_no,
    module_title,
    subs,
    slide_no,
    left_logo,
    right_logo,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_header(
        slide,
        module_no,
        module_title,
        "Key Takeaways / મુખ્ય મુદ્દાઓ",
        left_logo,
        right_logo,
    )

    takeaways = []

    for sub_no, sub in list(subs.items())[:8]:
        obj = all_text(sub.get("OBJ", []))
        if obj:
            takeaways.append(f"{sub_no}: {obj}")
        else:
            takeaways.append(f"{sub_no}: {sub['title']}")

    if not takeaways:
        takeaways = ["Moduleના મુખ્ય મુદ્દાઓ Word documentમાંથી ઉપલબ્ધ નથી."]

    for i, text in enumerate(takeaways):
        col = i % 2
        row = i // 2

        x = Inches(0.65) + col * Inches(6.15)
        y = Inches(1.62) + row * Inches(1.05)

        rounded_rect(
            slide,
            x, y,
            Inches(5.75), Inches(0.82),
            WHITE,
            line=RGBColor(0xD8, 0xE1, 0xE8),
        )

        circle(
            slide,
            x + Inches(0.15), y + Inches(0.18),
            Inches(0.44),
            TEAL,
        )

        add_text(
            slide,
            x + Inches(0.15), y + Inches(0.18),
            Inches(0.44), Inches(0.44),
            "✓",
            size=10,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
            font="Arial",
        )

        add_text(
            slide,
            x + Inches(0.75), y + Inches(0.10),
            Inches(4.75), Inches(0.62),
            "\n".join(wrap_text(text, 62)[:3]),
            size=10.5,
            color=DARK,
            bold=True,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    add_footer(slide, slide_no)


# ============================================================
# Build presentation
# ============================================================

def generate_presentation(
    modules,
    left_logo,
    right_logo,
    title,
    asset_files,
    add_module_overviews=True,
    add_summaries=True,
):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Cover
    build_cover(prs, left_logo, right_logo, title)

    # Intro
    build_intro_slide(
        prs,
        "પ્રસ્તાવના",
        [
            "રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ અંતર્ગત જાહેર આરોગ્ય કામગીરી માટે માર્ગદર્શન.",
            "ક્ષયરોગના દર્દીની ઓળખ, તપાસ, નોંધણી, સારવાર અને અનુસરણમાં ગુણવત્તાયુક્ત જાહેર આરોગ્ય કામગીરી.",
        ],
        len(prs.slides) + 1,
        left_logo,
        right_logo,
    )

    build_intro_slide(
        prs,
        "હેતુ",
        [
            "TB સેવાઓને સમયસર, ગુણવત્તાયુક્ત અને દર્દી-કેન્દ્રિત બનાવવી.",
            "NTEPની જાહેર આરોગ્ય કામગીરીમાં જવાબદારી, મોનિટરિંગ અને documentation મજબૂત કરવું.",
        ],
        len(prs.slides) + 1,
        left_logo,
        right_logo,
    )

    build_intro_slide(
        prs,
        "માર્ગદર્શિકાના મુખ્ય ઉદ્દેશ્યો",
        [
            "Presumptive TBથી treatment અને follow-up સુધીની પ્રક્રિયાને standardize કરવી.",
            "જવાબદાર વ્યક્તિઓ અને સમયમર્યાદા સ્પષ્ટ કરવી.",
            "Nikshay documentation અને monitoring indicatorsને મજબૂત બનાવવું.",
        ],
        len(prs.slides) + 1,
        left_logo,
        right_logo,
    )

    # Index
    build_index_slide(
        prs,
        modules,
        len(prs.slides) + 1,
        left_logo,
        right_logo,
    )

    # Modules
    for mod_no, mod in sorted(modules.items(), key=lambda z: int(z[0])):
        if add_module_overviews:
            build_module_overview(
                prs,
                mod_no,
                mod["title"],
                mod["subs"],
                len(prs.slides) + 1,
                left_logo,
                right_logo,
            )

        for sub_no, sub in sorted(
            mod["subs"].items(),
            key=lambda z: tuple(map(int, z[0].split("."))),
        ):
            build_content_slide(
                prs,
                mod_no,
                mod["title"],
                sub,
                len(prs.slides) + 1,
                left_logo,
                right_logo,
                asset_files,
            )

        if add_summaries:
            build_summary_slide(
                prs,
                mod_no,
                mod["title"],
                mod["subs"],
                len(prs.slides) + 1,
                left_logo,
                right_logo,
            )

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output, len(prs.slides)


# ============================================================
# Streamlit UI
# ============================================================

st.title("🏥 AMC NTEP — Professional PPT Auto-Designer")
st.caption(
    "Gujarati + English Word document → professional, infographic-style PowerPoint"
)

with st.sidebar:
    st.header("⚙️ Presentation Settings")

    presentation_title = st.text_input(
        "Presentation title",
        value="જાહેર આરોગ્ય કામગીરી (Public Health Actions)",
    )

    add_overviews = st.checkbox(
        "Add module overview slides",
        value=True,
    )

    add_summaries = st.checkbox(
        "Add module summary slides",
        value=True,
    )

    st.divider()

    st.markdown("### Branding")
    left_logo_file = st.file_uploader(
        "AMC / Government logo",
        type=["png", "jpg", "jpeg"],
        key="left_logo",
    )

    right_logo_file = st.file_uploader(
        "NTEP / TB logo",
        type=["png", "jpg", "jpeg"],
        key="right_logo",
    )

    st.divider()

    st.markdown("### 📷 Visual Asset Library")
    asset_uploads = st.file_uploader(
        "Upload photos / maps / equipment images",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="assets",
        help="Use filenames such as CBNAAT.png, Truenat.png, microscope.jpg, migration_map.png",
    )

uploaded_docx = st.file_uploader(
    "📄 Upload your Gujarati + English Word document (.docx)",
    type=["docx"],
)

if uploaded_docx:
    try:
        modules = parse_word(uploaded_docx)

        total_subs = sum(
            len(m["subs"]) for m in modules.values()
        )

        st.success(
            f"Detected {len(modules)} modules and {total_subs} sub-modules."
        )

        with st.expander("🔎 Preview parsed structure"):
            for mod_no, mod in sorted(
                modules.items(),
                key=lambda z: int(z[0]),
            ):
                st.markdown(
                    f"### Module {mod_no}: {mod['title']}"
                )

                for sub_no, sub in sorted(
                    mod["subs"].items(),
                    key=lambda z: tuple(map(int, z[0].split("."))),
                ):
                    st.markdown(
                        f"**{sub_no} — {sub['title']}**"
                    )

                    for field in [
                        "OBJ", "STEPS", "WHO", "TIME",
                        "IND", "WHY", "RISK", "DOC", "TRIGGER"
                    ]:
                        if sub.get(field):
                            st.write(
                                f"{field}:",
                                sub[field][:4]
                            )

        if asset_uploads:
            st.info(
                f"{len(asset_uploads)} visual assets uploaded."
            )

        if total_subs == 0:
            st.error(
                "No sub-modules detected. Use headings such as 1.1, 1.2, 2.1 etc. in the Word file."
            )
        else:
            if st.button(
                "🚀 Generate Professional PPT",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner(
                    "Designing cover, index, module slides, workflows, cards and visual layouts..."
                ):
                    output, slide_count = generate_presentation(
                        modules=modules,
                        left_logo=left_logo_file,
                        right_logo=right_logo_file,
                        title=presentation_title,
                        asset_files=save_uploaded_assets(
                            asset_uploads or [],
                            asset_dir="assets"
                        ),
                        add_module_overviews=add_overviews,
                        add_summaries=add_summaries,
                    )

                st.success(
                    f"✅ Presentation generated successfully — {slide_count} slides."
                )

                st.download_button(
                    "📥 Download AMC NTEP Professional PPT",
                    data=output,
                    file_name="AMC_NTEP_Professional_Public_Health_Actions.pptx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    ),
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"Error while reading/generating the presentation: {e}")

else:
    st.info(
        "Upload your Word document to begin. "
        "For best visual results, also upload AMC/NTEP logos and relevant photos."
    )
