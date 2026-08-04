import streamlit as st
import streamlit.components.v1 as components
import mammoth
import weasyprint
import base64
from bs4 import BeautifulSoup
import tempfile
import os

st.set_page_config(page_title="AMC NTEP Manual Generator", layout="wide")
st.title("AMC NTEP - Official Booklet & Flipbook Generator")

# --- Helper Function: Convert local image to Base64 ---
def get_image_base64(filepath):
    if os.path.exists(filepath):
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            mime_type = "image/png" if filepath.lower().endswith(".png") else "image/jpeg"
            return f"data:{mime_type};base64,{encoded_string}"
    return ""

# 1. Load Logos
amc_logo_b64 = get_image_base64("Amdavad_Municipal_Corporation_logo.png")
ntep_logo_b64 = get_image_base64("1-s2.0-S0019570720303152-gr1.jpg") 

# 2. Load Heritage Background (Checks which one you uploaded)
bg_image_b64 = ""
possible_bgs = ["banner_sidi-saiyyad-jali_902.jpg", "riverfront.jpg", "image_e9f81d.jpg"]
for bg in possible_bgs:
    if os.path.exists(bg):
        bg_image_b64 = get_image_base64(bg)
        break

uploaded_docx = st.file_uploader("Upload Content Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:
    with st.spinner("Extracting Gujarati content..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(uploaded_docx.read())
            tmp_docx_path = tmp_docx.name

        with open(tmp_docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            raw_html = result.value
        os.remove(tmp_docx_path) 

    with st.spinner("Applying Classic Book Theme..."):
        soup = BeautifulSoup(raw_html, 'html.parser')

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Gujarati:wght@400;700&display=swap');
                
                body {{ 
                    font-family: 'Noto Sans Gujarati', sans-serif; 
                    line-height: 1.6; 
                    color: #000; 
                    text-align: justify;
                }}
                
                /* ----------------------------------------------------- */
                /* 1. RUNNING HEADER (Absolute Positioning = Perfect)    */
                /* ----------------------------------------------------- */
                header {{ 
                    position: running(pageHeader); 
                    width: 100%;
                    height: 70px;
                    border-bottom: 2px solid #000;
                    margin-bottom: 20px;
                }}
                
                /* This forces the logos to the exact edges unconditionally */
                .hdr-amc {{ position: absolute; left: 0; top: 0; height: 60px; max-width: 90px; object-fit: contain; }}
                .hdr-ntep {{ position: absolute; right: 0; top: 0; height: 60px; max-width: 90px; object-fit: contain; }}
                
                .hdr-text {{ 
                    text-align: center; 
                    width: 100%; 
                    padding-top: 20px; 
                    font-size: 15px; 
                    font-weight: bold; 
                    color: #000; 
                }}

                /* ----------------------------------------------------- */
                /* 2. PAGE SETTINGS                                      */
                /* ----------------------------------------------------- */
                @page {{
                    size: A4;
                    margin: 3.5cm 2.5cm 2.5cm 2.5cm;
                    background-color: #ffffff; 
                    
                    @top-center {{ content: element(pageHeader); }}
                    @bottom-center {{ content: counter(page); font-family: 'Arial', sans-serif; }}
                }}

                /* ----------------------------------------------------- */
                /* 3. CLASSIC CONSTITUTION-STYLE COVER PAGE              */
                /* ----------------------------------------------------- */
                @page cover {{
                    margin: 0cm; 
                    @top-center {{ content: none; }} 
                    @bottom-center {{ content: none; }} 
                }}

                .cover-page {{
                    page: cover; 
                    page-break-after: always;
                    position: relative;
                    width: 21cm;
                    height: 29.7cm;
                    background-color: #0A192F; /* Very dark, classic navy/slate */
                    overflow: hidden;
                    text-align: center;
                }}

                /* Blurred Heritage Background */
                .cover-bg {{
                    position: absolute;
                    top: 0; left: 0; right: 0; bottom: 0;
                    background-image: url('{bg_image_b64}');
                    background-size: cover;
                    background-position: center;
                    filter: blur(5px);
                    opacity: 0.15; /* Keeps it subtle so text is readable */
                    z-index: 1;
                }}

                /* Ornate Gold Border */
                .cover-border {{
                    position: absolute;
                    top: 1.5cm; left: 1.5cm; right: 1.5cm; bottom: 1.5cm;
                    border: 4px solid #D4AF37; /* Classic Gold */
                    outline: 1px solid #D4AF37;
                    outline-offset: -10px;
                    z-index: 2;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                    padding: 2cm;
                    box-sizing: border-box;
                }}

                /* Cover Logos */
                .cover-logos {{
                    position: relative;
                    width: 100%;
                    height: 120px;
                }}
                .c-logo-amc {{ position: absolute; left: 0; top: 0; height: 110px; }}
                .c-logo-ntep {{ position: absolute; right: 0; top: 0; height: 110px; }}

                /* Title Block */
                .cover-title-box {{
                    background-color: rgba(10, 25, 47, 0.85);
                    border: 2px solid #D4AF37;
                    padding: 40px 20px;
                    margin: 1cm 0;
                }}

                .cover-title {{
                    font-family: 'Georgia', serif;
                    font-size: 50px;
                    font-weight: bold;
                    color: #FFFFFF;
                    text-transform: uppercase;
                    letter-spacing: 2px;
                    margin: 0;
                    line-height: 1.3;
                }}

                .cover-subtitle {{
                    font-family: 'Georgia', serif;
                    font-size: 22px;
                    color: #D4AF37;
                    margin-top: 15px;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                }}

                .cover-footer {{
                    font-family: 'Georgia', serif;
                    font-size: 18px;
                    color: #FFFFFF;
                }}

                /* ----------------------------------------------------- */
                /* 4. CONTENT FORMATTING                                 */
                /* ----------------------------------------------------- */
                .content h1 {{ page-break-before: always; color: #000; border-bottom: 2px solid #000; padding-bottom: 5px; margin-top: 0; }}
                .content h2, .content h3 {{ color: #000; font-weight: bold; margin-top: 25px; }}
                .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
                .content li {{ margin-bottom: 8px; }}
                .content table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                .content th, .content td {{ border: 1px solid #000; padding: 10px; text-align: left; }}
                .content th {{ background-color: #f2f2f2; color: #000; font-weight: bold; text-align: center; }}
                
                /* Ensures highlighted text in Word stays highlighted in PDF */
                mark {{ background-color: #FFFF00; color: #000; }}
                
            </style>
        </head>
        <body>
            <!-- Running Header -->
            <header>
                <img src="{amc_logo_b64}" class="hdr-amc" alt="AMC Logo">
                <div class="hdr-text">રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC</div>
                <img src="{ntep_logo_b64}" class="hdr-ntep" alt="NTEP Logo">
            </header>
            
            <!-- Classic Cover Page -->
            <div class="cover-page">
                <div class="cover-bg"></div>
                <div class="cover-border">
                    
                    <div class="cover-logos">
                        <img src="{amc_logo_b64}" class="c-logo-amc" alt="AMC Logo">
                        <img src="{ntep_logo_b64}" class="c-logo-ntep" alt="NTEP Logo">
                    </div>
                    
                    <div class="cover-title-box">
                        <div class="cover-title">Public Health<br>Action</div>
                        <div class="cover-subtitle">Operational Manual</div>
                    </div>
                    
                    <div class="cover-footer">
                        National Tuberculosis Elimination Program<br>
                        Ahmedabad Municipal Corporation<br><br>
                        &copy; 2026
                    </div>
                    
                </div>
            </div>
            
            <!-- Content -->
            <div class="content">
                {str(soup)}
            </div>
        </body>
        </html>
        """

    with st.spinner("Generating High-Quality PDF..."):
        pdf_bytes = weasyprint.HTML(string=full_html).write_pdf()

    st.success("Manual & Flipbook Generated Successfully!")
    
    st.download_button(
        label="📄 Download Official PDF Manual",
        data=pdf_bytes,
        file_name="AMC_NTEP_Operational_Manual.pdf",
        mime="application/pdf"
    )

    st.markdown("---")
    st.header("📖 3D Interactive Flipbook")

    b64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
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
    
    with st.spinner("Rendering 3D Flipbook Viewer..."):
        components.html(flipbook_html, height=750, scrolling=False)
