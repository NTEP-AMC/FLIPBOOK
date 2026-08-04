import streamlit as st
import streamlit.components.v1 as components
import mammoth
import weasyprint
import base64
import re
from bs4 import BeautifulSoup, NavigableString, Tag
import tempfile
import os
from datetime import date

st.set_page_config(page_title="AMC NTEP Manual Generator", layout="wide")
st.title("AMC NTEP - Official Booklet & Flipbook Generator")

# ============================================================
# HELPERS
# ============================================================

def get_image_base64(filepath):
    """Convert a local image to a base64 data URI (empty string if missing)."""
    if os.path.exists(filepath):
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            mime_type = "image/png" if filepath.lower().endswith(".png") else "image/jpeg"
            return f"data:{mime_type};base64,{encoded_string}"
    return ""


def first_existing(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return None


# ============================================================
# SMALL HAND-DRAWN ICON SET (inline SVG, currentColor based)
# Used for module dividers and in-content field badges.
# ============================================================

ICONS = {
    "stethoscope": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M5 3v6a4 4 0 0 0 8 0V3" stroke-linecap="round"/><circle cx="18" cy="16" r="3"/><path d="M13 8v3a5 5 0 0 0 5 5" stroke-linecap="round"/><path d="M5 3H3.5M13 3h1.5" stroke-linecap="round"/></svg>',
    "clipboard": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1H9z"/><path d="M8.5 13l2 2 4.5-4.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "pill": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3.5" y="9.5" width="17" height="8" rx="4" transform="rotate(-40 12 13.5)"/><path d="M9.5 15.8L14.2 11" stroke-linecap="round"/></svg>',
    "microscope": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M9 20h9" stroke-linecap="round"/><path d="M11 20a5 5 0 1 1 4-8" /><path d="M12 12l4-4 2 2-4 4" /><path d="M16 6l2-2" stroke-linecap="round"/></svg>',
    "radar": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.2"/><circle cx="12" cy="12" r="0.9" fill="currentColor" stroke="none"/><path d="M12 4v-2M20 12h2M12 20v2M4 12H2" stroke-linecap="round"/></svg>',
    "badge": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="9" r="6"/><path d="M9 20l3-3 3 3-1-6.5h-4z"/><path d="M9 9l2 2 4-4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "heart": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 20s-7-4.6-9.5-9A5 5 0 0 1 12 6a5 5 0 0 1 9.5 5c-2.5 4.4-9.5 9-9.5 9z"/><path d="M4 12h3l1.5-3L11 15l1.5-4H20" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "document": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4" /><path d="M9 12h7M9 16h7M9 8h3" stroke-linecap="round"/></svg>',
    "target": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/></svg>',
    "bell": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 10a6 6 0 0 1 12 0c0 4 1.5 5.5 1.5 5.5H4.5S6 14 6 10z"/><path d="M10 19a2 2 0 0 0 4 0" stroke-linecap="round"/></svg>',
    "checklist": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 6h2M4 12h2M4 18h2" stroke-linecap="round"/><path d="M9 6h11M9 12h11M9 18h11" stroke-linecap="round"/></svg>',
    "lightbulb": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M9 18h6M10 21h4"/><path d="M12 3a6 6 0 0 0-3.5 10.9c.6.5 1 1.2 1 2.1h5c0-.9.4-1.6 1-2.1A6 6 0 0 0 12 3z"/></svg>',
    "users": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="9" cy="8" r="3"/><path d="M3 20a6 6 0 0 1 12 0" stroke-linecap="round"/><circle cx="17.5" cy="9.5" r="2.3"/><path d="M15 20a5 5 0 0 1 6.5-4.8" stroke-linecap="round"/></svg>',
    "userbadge": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0" stroke-linecap="round"/><path d="M9.5 12.5l1.5 1.5 3-3" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "clock": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="8.5"/><path d="M12 7v5l3.5 2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "filetext": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4"/><path d="M9 12h7M9 16h7M9 8h3" stroke-linecap="round"/></svg>',
    "chart": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "alert": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 3l10 18H2z" stroke-linejoin="round"/><path d="M12 10v4" stroke-linecap="round"/><circle cx="12" cy="17.3" r="0.9" fill="currentColor" stroke="none"/></svg>',
}

# Rotating colour palette — one theme per module, cycles if there are more
# than 7 modules in the uploaded document.
PALETTE = [
    {"color": "#1F5FA8", "light": "#DCEAFB"},
    {"color": "#0F7A6B", "light": "#DCF3EE"},
    {"color": "#6A3FA0", "light": "#EDE3F7"},
    {"color": "#C9660B", "light": "#FBE7D4"},
    {"color": "#B23A48", "light": "#F8DEE1"},
    {"color": "#2E7D32", "light": "#DFF3E0"},
    {"color": "#3949AB", "light": "#E1E4FA"},
]

FIELD_KEYWORDS = [
    (("objective",), "target"),
    (("trigger", "when"), "bell"),
    (("what to do",), "checklist"),
    (("why",), "lightbulb"),
    (("whom",), "users"),
    (("responsible",), "userbadge"),
    (("timeline",), "clock"),
    (("documentation",), "filetext"),
    (("monitoring",), "chart"),
    (("supervisor", "checklist", "quality"), "checklist"),
    (("if not done", "risk"), "alert"),
]

MODULE_PATTERN = re.compile(r'^Module\s+(\d+)\s*:\s*(.+)$', re.IGNORECASE)


def pick_module_icon(title):
    t = title.lower()
    if "diagnos" in t:
        return "stethoscope"
    if "before treatment" in t or "initiation" in t:
        return "clipboard"
    if "during treatment" in t:
        return "pill"
    if "laboratory" in t or "clinical" in t:
        return "microscope"
    if "surveillance" in t:
        return "radar"
    if "completion" in t or "outcome" in t:
        return "badge"
    if "post-treatment" in t or "follow-up" in t or "follow up" in t:
        return "heart"
    return "document"


def pick_field_icon(label_text):
    t = label_text.lower()
    for keys, icon in FIELD_KEYWORDS:
        if any(k in t for k in keys):
            return icon
    return "document"


def add_module_theming(soup):
    """
    Detects 'Module N: <Title>' headings (in ANY tag - h1/h2/h3/p, since Word
    docs are rarely styled consistently), and for each one:
      - inserts a full-page colour-themed divider with a matching icon
      - normalises the heading itself into a real <h1>
      - wraps everything up to the next module into a themed <div> so that
        headings/tables inside inherit that module's colour via CSS vars
    Returns the number of modules detected.
    """
    candidates = []
    for tag in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p']):
        txt = tag.get_text().strip()
        if len(txt) > 90:
            continue
        m = MODULE_PATTERN.match(txt)
        if m:
            candidates.append((tag, m))

    if not candidates:
        return 0

    module_meta = []
    for i, (tag, m) in enumerate(candidates):
        num = m.group(1)
        title = m.group(2).strip()
        pal = PALETTE[i % len(PALETTE)]
        icon_key = pick_module_icon(title)
        module_meta.append({"num": num, "title": title, "color": pal["color"], "light": pal["light"]})

        divider_html = f'''
        <div class="module-divider" data-mod="{num}" style="--mod-color:{pal['color']}; --mod-color-light:{pal['light']};">
          <div class="module-divider-num">{num.zfill(2)}</div>
          <div class="module-divider-inner">
            <div class="module-divider-icon">{ICONS.get(icon_key, ICONS['document'])}</div>
            <div class="module-divider-eyebrow">MODULE {num}</div>
            <div class="module-divider-title">{title}</div>
            <div class="module-divider-rule"></div>
          </div>
        </div>
        '''
        tag.insert_before(BeautifulSoup(divider_html, 'html.parser'))

        new_h1 = soup.new_tag('h1')
        new_h1.string = f"Module {num}: {title}"
        tag.replace_with(new_h1)

    # Wrap each module's content (divider + heading + everything until the
    # next divider) into a themed container.
    top_level = [c for c in list(soup.contents) if isinstance(c, Tag)]
    divider_indices = [i for i, t in enumerate(top_level) if t.get('class') and 'module-divider' in t.get('class')]

    for j, start in enumerate(divider_indices):
        end = divider_indices[j + 1] if j + 1 < len(divider_indices) else len(top_level)
        group = top_level[start:end]
        meta = module_meta[j]
        wrapper = soup.new_tag('div')
        wrapper['class'] = 'module-block'
        wrapper['style'] = f"--mod-color:{meta['color']}; --mod-color-light:{meta['light']};"
        group[0].insert_before(wrapper)
        for g in group:
            wrapper.append(g.extract())

    return len(module_meta)


def add_field_badges(soup):
    """
    Finds list items shaped like:
      <li><strong>લેબલ (English Label):<br/>...details...</strong></li>
    (the common pattern this kind of SOP/manual content uses) and turns the
    leading label into a small icon badge, keeping the details as normal text.
    """
    count = 0
    for li in soup.find_all('li'):
        strong = li.find('strong', recursive=False)
        if not strong:
            continue
        contents = list(strong.contents)
        br_idx = None
        for i, c in enumerate(contents):
            if isinstance(c, Tag) and c.name == 'br':
                br_idx = i
                break
        if br_idx is None:
            continue

        label_nodes = contents[:br_idx]
        label_text = ''.join(n.get_text() if isinstance(n, Tag) else str(n) for n in label_nodes).strip()
        m = re.search(r'\(([^)]{2,40})\)\s*:?\s*$', label_text)
        if not m:
            continue

        icon_key = pick_field_icon(m.group(1))
        rest_nodes = contents[br_idx + 1:]

        new_strong = soup.new_tag('strong')
        for n in rest_nodes:
            new_strong.append(n.extract() if isinstance(n, Tag) else NavigableString(str(n)))

        badge_html = (
            f'<div class="field-badge">'
            f'<span class="field-badge-icon">{ICONS.get(icon_key, ICONS["document"])}</span>'
            f'<span class="field-badge-label">{label_text}</span></div>'
        )
        strong.clear()
        strong.decompose()
        li.append(BeautifulSoup(badge_html, 'html.parser'))
        if rest_nodes:
            body_div = soup.new_tag('div')
            body_div['class'] = 'field-body'
            body_div.append(new_strong)
            li.append(body_div)
        count += 1

    return count


# ============================================================
# ASSETS
# ============================================================

amc_logo_b64 = get_image_base64("Amdavad_Municipal_Corporation_logo.png")
ntep_logo_b64 = get_image_base64("1-s2.0-S0019570720303152-gr1.jpg")

bg_path = first_existing([
    "banner_sidi-saiyyad-jali_902.jpg",
    "riverfront.jpg",
    "image_e9f81d.jpg",
])
bg_image_b64 = get_image_base64(bg_path) if bg_path else ""

# ============================================================
# SIDEBAR — COVER PAGE CONTROLS
# ============================================================

with st.sidebar:
    st.header("Cover Page Details")
    cover_title = st.text_input("Main Title", "NTEP OPERATIONAL MANUAL")
    cover_subtitle = st.text_input(
        "Subtitle", "National Tuberculosis Elimination Programme"
    )
    cover_org_line1 = st.text_input("Organisation Line 1", "Ahmedabad Municipal Corporation")
    cover_org_line2 = st.text_input("Organisation Line 2 (Gujarati)", "અમદાવાદ મ્યુનિસિપલ કોર્પોરેશન")
    cover_year = st.text_input("Year / Edition", str(date.today().year))
    header_title_guj = st.text_input(
        "Running Header (Gujarati)",
        "રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC",
    )
    theme = st.selectbox(
        "Cover Colour Theme",
        ["Government Navy & Gold", "Health Teal & Coral", "Maroon & Gold (Classic Gazette)"],
        index=0,
    )
    st.markdown("---")
    add_theming = st.checkbox("Auto-theme each Module (colour + icon per section)", value=True)
    add_badges = st.checkbox("Add icon badges to field labels (Objective, Trigger, etc.)", value=True)

THEMES = {
    "Government Navy & Gold": {"primary": "#0B2545", "accent": "#C6A15B", "accent_light": "#E7D6AE", "text_on_primary": "#FFFFFF"},
    "Health Teal & Coral": {"primary": "#0F3D3E", "accent": "#E9724C", "accent_light": "#F2C6B4", "text_on_primary": "#FFFFFF"},
    "Maroon & Gold (Classic Gazette)": {"primary": "#5C1A1B", "accent": "#D4AF37", "accent_light": "#EAD9A0", "text_on_primary": "#FFFFFF"},
}
T = THEMES[theme]

# ============================================================
# UPLOAD & GENERATE
# ============================================================

uploaded_docx = st.file_uploader("Upload Content Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:

    with st.spinner("Extracting content..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(uploaded_docx.read())
            tmp_docx_path = tmp_docx.name
        with open(tmp_docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            raw_html = result.value
        os.remove(tmp_docx_path)

    with st.spinner("Designing module themes & content graphics..."):
        soup = BeautifulSoup(raw_html, "html.parser")

        modules_found = 0
        if add_theming:
            modules_found = add_module_theming(soup)
        if add_badges:
            add_field_badges(soup)

        if modules_found:
            st.info(f"Detected {modules_found} module(s) in the document — each was auto-themed with its own colour and icon.")

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Gujarati:wght@400;600;700&family=Playfair+Display:wght@600;700;800&family=Lora:wght@400;500;600&display=swap');

            :root {{
                --primary: {T['primary']};
                --accent: {T['accent']};
                --accent-light: {T['accent_light']};
                --on-primary: {T['text_on_primary']};
            }}

            * {{ box-sizing: border-box; }}

            body {{
                font-family: 'Noto Sans Gujarati', 'Lora', serif;
                line-height: 1.7;
                color: #1a1a1a;
                text-align: justify;
                margin: 0;
            }}

            /* ---------------------------------------------------- */
            /* RUNNING HEADER                                       */
            /* ---------------------------------------------------- */
            #page-header {{ position: running(pageHeader); width: 100%; }}

            .header-table {{
                width: 100%;
                border-collapse: collapse;
                border-bottom: 2px solid var(--primary);
                margin-bottom: 6px;
            }}
            .header-table td {{ padding-bottom: 8px; vertical-align: middle; border: none; }}
            .h-left {{ text-align: left; width: 15%; }}
            .h-center {{
                text-align: center; width: 70%;
                font-size: 14px; font-weight: 700; color: var(--primary);
                letter-spacing: 0.3px;
            }}
            .h-right {{ text-align: right; width: 15%; }}
            .hdr-logo {{ height: 46px; object-fit: contain; }}

            /* ---------------------------------------------------- */
            /* PAGE SETUP                                            */
            /* ---------------------------------------------------- */
            @page {{
                size: A4;
                margin: 3.2cm 2cm 2.2cm 2cm;
                background-color: #ffffff;
                @top-center {{ content: element(pageHeader); width: 100%; }}
                @bottom-center {{
                    content: counter(page);
                    font-family: 'Lora', serif;
                    color: var(--primary);
                    font-size: 11px;
                }}
            }}

            @page cover {{
                margin: 0;
                @top-center {{ content: none; }}
                @bottom-center {{ content: none; }}
            }}

            /* ---------------------------------------------------- */
            /* COVER PAGE                                            */
            /* ---------------------------------------------------- */
            .cover-page {{
                page: cover;
                page-break-after: always;
                position: relative;
                width: 21cm;
                height: 29.7cm;
                background: #fbf9f4;
                overflow: hidden;
            }}
            .cover-photo-strip {{
                position: absolute;
                top: 0; left: 0; right: 0;
                height: 9cm;
                background-image: url('{bg_image_b64}');
                background-size: cover;
                background-position: center;
            }}
            .cover-photo-strip::after {{
                content: "";
                position: absolute;
                inset: 0;
                background: linear-gradient(180deg, rgba(11,37,69,0.35) 0%, var(--primary) 92%);
            }}
            .cover-tricolor {{
                position: absolute;
                top: 0; left: 0; right: 0;
                height: 6px;
                background: linear-gradient(90deg, #FF9933 33%, #FFFFFF 33% 66%, #138808 66%);
                z-index: 5;
            }}
            .cover-emblem-row {{
                position: relative;
                z-index: 4;
                padding: 1.4cm 1.6cm 0 1.6cm;
                display: table;
                width: 100%;
            }}
            .cover-emblem-cell {{
                display: table-cell;
                vertical-align: middle;
                width: 33.3%;
                text-align: center;
            }}
            .cover-emblem-cell.left {{ text-align: left; }}
            .cover-emblem-cell.right {{ text-align: right; }}
            .c-logo {{
                height: 92px;
                background: #ffffff;
                border-radius: 50%;
                padding: 6px;
                border: 3px solid var(--accent);
            }}
            .cover-eyebrow {{
                color: #ffffff;
                font-family: 'Lora', serif;
                font-size: 13px;
                letter-spacing: 3px;
                text-transform: uppercase;
                opacity: 0.9;
            }}
            .cover-title-block {{
                position: relative;
                z-index: 4;
                margin-top: 6.6cm;
                padding: 0 2cm;
                text-align: center;
            }}
            .cover-title-rule {{
                width: 90px;
                height: 3px;
                background: var(--accent);
                margin: 0 auto 18px auto;
            }}
            .cover-title {{
                font-family: 'Playfair Display', serif;
                font-size: 46px;
                font-weight: 800;
                color: var(--primary);
                letter-spacing: 0.5px;
                margin: 0;
                line-height: 1.25;
            }}
            .cover-subtitle {{
                font-family: 'Lora', serif;
                font-size: 19px;
                font-weight: 500;
                color: #333333;
                margin-top: 14px;
            }}
            .cover-divider {{
                width: 60%;
                margin: 34px auto;
                border: none;
                border-top: 1px solid var(--accent);
            }}
            .cover-org-block {{ text-align: center; font-family: 'Lora', serif; }}
            .cover-org-line1 {{ font-size: 20px; font-weight: 600; color: var(--primary); }}
            .cover-org-line2 {{
                font-size: 20px; font-family: 'Noto Sans Gujarati', sans-serif;
                color: var(--primary); margin-top: 4px;
            }}
            .cover-year {{
                margin-top: 10px; font-size: 14px; letter-spacing: 2px;
                color: var(--accent); text-transform: uppercase; font-weight: 600;
            }}
            .cover-footer-bar {{
                position: absolute;
                bottom: 0; left: 0; right: 0;
                z-index: 4;
                background: var(--primary);
                color: var(--on-primary);
                padding: 16px 2cm;
                text-align: center;
                font-family: 'Lora', serif;
                font-size: 12px;
                letter-spacing: 1px;
                border-top: 4px solid var(--accent);
            }}

            /* ---------------------------------------------------- */
            /* MODULE DIVIDER PAGES (auto-generated, one per module) */
            /* ---------------------------------------------------- */
            @page module-divider-page {{
                margin: 0;
                @top-center {{ content: none; }}
                @bottom-center {{ content: none; }}
            }}
            .module-divider {{
                page: module-divider-page;
                page-break-before: always;
                page-break-after: always;
                width: 21cm;
                height: 29.7cm;
                position: relative;
                background: var(--mod-color-light);
                display: flex;
                align-items: center;
                justify-content: center;
                overflow: hidden;
            }}
            .module-divider::before {{
                content: ""; position: absolute; top: 0; left: 0; right: 0;
                height: 12px; background: var(--mod-color);
            }}
            .module-divider::after {{
                content: ""; position: absolute; bottom: 0; left: 0; right: 0;
                height: 12px; background: var(--mod-color);
            }}
            .module-divider-num {{
                position: absolute; top: 50%; left: 50%;
                transform: translate(-50%, -50%);
                font-family: 'Playfair Display', serif;
                font-size: 480px; font-weight: 800;
                color: var(--mod-color); opacity: 0.10; line-height: 1;
            }}
            .module-divider-inner {{ position: relative; text-align: center; z-index: 1; }}
            .module-divider-icon {{
                width: 90px; height: 90px; margin: 0 auto 20px auto;
                color: var(--mod-color); background: #fff; border-radius: 50%;
                border: 3px solid var(--mod-color);
                display: flex; align-items: center; justify-content: center;
                padding: 18px;
            }}
            .module-divider-icon svg {{ width: 100%; height: 100%; }}
            .module-divider-eyebrow {{
                font-family: 'Lora', serif; letter-spacing: 4px; font-size: 15px;
                color: var(--mod-color); font-weight: 700; margin-bottom: 10px;
            }}
            .module-divider-title {{
                font-family: 'Playfair Display', serif; font-size: 32px; font-weight: 800;
                color: #1a1a1a; max-width: 14cm; margin: 0 auto;
            }}
            .module-divider-rule {{
                width: 70px; height: 3px; background: var(--mod-color);
                margin: 18px auto 0 auto;
            }}

            /* ---------------------------------------------------- */
            /* MODULE-THEMED CONTENT                                 */
            /* ---------------------------------------------------- */
            .module-block h1 {{
                page-break-before: always;
                color: var(--mod-color);
                font-family: 'Playfair Display', serif;
                font-size: 24px;
                border-bottom: 3px solid var(--mod-color);
                padding-bottom: 8px;
                margin-top: 0;
            }}
            .module-block h2 {{
                color: var(--mod-color); font-weight: 700; font-size: 18px;
                margin-top: 24px; border-left: 4px solid var(--mod-color); padding-left: 10px;
            }}
            .module-block h3 {{
                color: var(--mod-color); font-weight: 700; font-size: 15px;
                margin-top: 18px; border-left: 3px solid var(--mod-color); padding-left: 8px;
            }}
            .module-block th {{ background-color: var(--mod-color); }}

            /* Fallback (non-themed) content headings, when auto-theming is off
               or for text before the first module */
            .content > h1 {{
                page-break-before: always;
                color: var(--primary);
                font-family: 'Playfair Display', serif;
                font-size: 26px;
                border-bottom: 3px solid var(--accent);
                padding-bottom: 8px;
                margin-top: 0;
            }}
            .content > h2 {{
                color: var(--primary); font-weight: 700; font-size: 19px;
                margin-top: 28px; border-left: 4px solid var(--accent); padding-left: 10px;
            }}
            .content > h3 {{
                color: var(--primary); font-weight: 700; font-size: 16px; margin-top: 22px;
            }}

            .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
            .content li {{ margin-bottom: 8px; }}
            .content table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
            .content th, .content td {{ border: 1px solid #999; padding: 9px 10px; text-align: left; font-size: 13.5px; }}
            .content th {{ background-color: var(--primary); color: #fff; font-weight: 700; text-align: center; }}
            .content tr:nth-child(even) td {{ background-color: #f7f5ef; }}

            /* ---------------------------------------------------- */
            /* FIELD BADGES (Objective / Trigger / Timeline / ...)   */
            /* ---------------------------------------------------- */
            .field-badge {{
                display: inline-flex;
                align-items: center;
                gap: 6px;
                background: var(--mod-color-light, var(--accent-light));
                color: var(--mod-color, var(--primary));
                border: 1px solid var(--mod-color, var(--accent));
                border-radius: 14px;
                padding: 3px 10px 3px 6px;
                font-size: 12.5px;
                font-weight: 700;
                margin-bottom: 5px;
            }}
            .field-badge-icon {{ width: 15px; height: 15px; display: inline-block; }}
            .field-badge-icon svg {{ width: 100%; height: 100%; }}
            .field-body {{ font-weight: 500; margin-top: 2px; margin-bottom: 4px; text-align: justify; }}
        </style>
        </head>
        <body>

            <!-- Running Header (content pages only) -->
            <div id="page-header">
                <table class="header-table">
                    <tr>
                        <td class="h-left"><img src="{amc_logo_b64}" class="hdr-logo" alt="AMC Logo"></td>
                        <td class="h-center">{header_title_guj}</td>
                        <td class="h-right"><img src="{ntep_logo_b64}" class="hdr-logo" alt="NTEP Logo"></td>
                    </tr>
                </table>
            </div>

            <!-- ============== COVER PAGE ============== -->
            <div class="cover-page">
                <div class="cover-tricolor"></div>
                <div class="cover-photo-strip"></div>

                <div class="cover-emblem-row">
                    <div class="cover-emblem-cell left"><img src="{amc_logo_b64}" class="c-logo" alt="AMC Logo"></div>
                    <div class="cover-emblem-cell"><span class="cover-eyebrow">Government of Gujarat</span></div>
                    <div class="cover-emblem-cell right"><img src="{ntep_logo_b64}" class="c-logo" alt="NTEP Logo"></div>
                </div>

                <div class="cover-title-block">
                    <div class="cover-title-rule"></div>
                    <div class="cover-title">{cover_title}</div>
                    <div class="cover-subtitle">{cover_subtitle}</div>

                    <hr class="cover-divider">

                    <div class="cover-org-block">
                        <div class="cover-org-line1">{cover_org_line1}</div>
                        <div class="cover-org-line2">{cover_org_line2}</div>
                        <div class="cover-year">Edition {cover_year}</div>
                    </div>
                </div>

                <div class="cover-footer-bar">
                    FOR OFFICIAL USE&nbsp;&nbsp;•&nbsp;&nbsp;NTEP – AHMEDABAD MUNICIPAL CORPORATION&nbsp;&nbsp;•&nbsp;&nbsp;{cover_year}
                </div>
            </div>

            <!-- ============== CONTENT ============== -->
            <div class="content">
                {str(soup)}
            </div>
        </body>
        </html>
        """

    with st.spinner("Generating high-quality PDF..."):
        pdf_bytes = weasyprint.HTML(string=full_html, base_url=".").write_pdf()

    st.success("Manual & Flipbook generated successfully!")

    st.download_button(
        label="📄 Download Official PDF Manual",
        data=pdf_bytes,
        file_name="AMC_NTEP_Operational_Manual.pdf",
        mime="application/pdf",
    )

    st.markdown("---")
    st.header("📖 3D Interactive Flipbook")

    b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_data_uri = f"data:application/pdf;base64,{b64_pdf}"

    flipbook_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <link href="https://cdn.jsdelivr.net/npm/dflip/css/dflip.min.css" rel="stylesheet">
        <link href="https://cdn.jsdelivr.net/npm/dflip/css/themify-icons.min.css" rel="stylesheet">
        <style>
            body {{ margin: 0; padding: 0; background-color: #f4f4f9; }}
            ._df_book {{ height: 100vh !important; }}
        </style>
    </head>
    <body>
        <div class="_df_book" webgl="true" backgroundcolor="#f4f4f9"
             source="{pdf_data_uri}" id="df_manual">
        </div>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/dflip/js/dflip.min.js"></script>
    </body>
    </html>
    """

    with st.spinner("Rendering 3D flipbook viewer..."):
        components.html(flipbook_html, height=750, scrolling=False)
