import streamlit as st
import streamlit.components.v1 as components
import mammoth
import weasyprint
import base64
from bs4 import BeautifulSoup
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
# SIDEBAR — LET THE USER TWEAK COVER TEXT WITHOUT TOUCHING CODE
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

THEMES = {
    "Government Navy & Gold": {
        "primary": "#0B2545",
        "accent": "#C6A15B",
        "accent_light": "#E7D6AE",
        "text_on_primary": "#FFFFFF",
    },
    "Health Teal & Coral": {
        "primary": "#0F3D3E",
        "accent": "#E9724C",
        "accent_light": "#F2C6B4",
        "text_on_primary": "#FFFFFF",
    },
    "Maroon & Gold (Classic Gazette)": {
        "primary": "#5C1A1B",
        "accent": "#D4AF37",
        "accent_light": "#EAD9A0",
        "text_on_primary": "#FFFFFF",
    },
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

    with st.spinner("Building the booklet layout..."):
        soup = BeautifulSoup(raw_html, "html.parser")

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
            /* COVER PAGE — CLEAN OFFICIAL GAZETTE STYLE             */
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

            /* Faint heritage motif, confined to a strip so it never
               muddies the whole page */
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
                background: linear-gradient(180deg,
                    rgba(11,37,69,0.35) 0%,
                    var(--primary) 92%);
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

            .cover-org-block {{
                text-align: center;
                font-family: 'Lora', serif;
            }}
            .cover-org-line1 {{
                font-size: 20px;
                font-weight: 600;
                color: var(--primary);
            }}
            .cover-org-line2 {{
                font-size: 20px;
                font-family: 'Noto Sans Gujarati', sans-serif;
                color: var(--primary);
                margin-top: 4px;
            }}
            .cover-year {{
                margin-top: 10px;
                font-size: 14px;
                letter-spacing: 2px;
                color: var(--accent);
                text-transform: uppercase;
                font-weight: 600;
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
            /* TITLE / DIVIDER PAGE (before main content, optional) */
            /* ---------------------------------------------------- */
            .section-divider {{
                page-break-before: always;
                page-break-after: always;
                height: 100%;
            }}

            /* ---------------------------------------------------- */
            /* CONTENT FORMATTING                                    */
            /* ---------------------------------------------------- */
            .content h1 {{
                page-break-before: always;
                color: var(--primary);
                font-family: 'Playfair Display', serif;
                font-size: 26px;
                border-bottom: 3px solid var(--accent);
                padding-bottom: 8px;
                margin-top: 0;
            }}
            .content h2 {{
                color: var(--primary);
                font-weight: 700;
                font-size: 19px;
                margin-top: 28px;
                border-left: 4px solid var(--accent);
                padding-left: 10px;
            }}
            .content h3 {{
                color: var(--primary);
                font-weight: 700;
                font-size: 16px;
                margin-top: 22px;
            }}
            .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
            .content li {{ margin-bottom: 8px; }}
            .content table {{
                width: 100%;
                border-collapse: collapse;
                margin: 15px 0;
            }}
            .content th, .content td {{
                border: 1px solid #999;
                padding: 9px 10px;
                text-align: left;
                font-size: 13.5px;
            }}
            .content th {{
                background-color: var(--primary);
                color: #fff;
                font-weight: 700;
                text-align: center;
            }}
            .content tr:nth-child(even) td {{ background-color: #f7f5ef; }}
        </style>
        </head>
        <body>

            <!-- Running Header (applies to content pages only) -->
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
