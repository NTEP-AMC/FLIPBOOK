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

# 2. Load Heritage Background (Safely)
bg_image_b64 = ""
possible_bgs = ["banner_sidi-saiyyad-jali_902.jpg", "riverfront.jpg", "image_e9f81d.jpg"]
for bg in possible_bgs:
    if os.path.exists(bg):
        bg_image_b64 = get_image_base64(bg)
        break

# Safely construct the CSS for the background only if it exists
bg_css_rule = f"background-image: url('{bg_image_b64}');" if bg_image_b64 else "background-color: transparent;"

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

    with st.spinner("Applying Perfect Government Alignment..."):
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
                /* 1. RUNNING HEADER (Bulletproof Table Layout)          */
                /* ----------------------------------------------------- */
                #page-header {{ 
                    position: running(pageHeader); 
                    width: 100%;
                }}
                
                .header-table {{
                    width: 100%;
                    border-collapse: collapse;
                    border-bottom: 2px solid #000;
                    margin-bottom: 10px;
                }}
                
                .header-table td {{ padding-bottom: 10px; vertical-align: middle; border: none; }}
                
                .h-left {{ text-align: left; width: 15%; }}
                .h-center {{ text-align: center; width: 70%; font-size: 16px; font-weight: bold; color: #000; }}
                .h-right {{ text-align: right; width: 15%; }}
                
                .hdr-logo {{ height: 60px; object-fit: contain; }}

                /* ----------------------------------------------------- */
                /* 2. PAGE SETTINGS                                      */
                /* ----------------------------------------------------- */
                @page {{
                    size: A4;
                    margin: 3.5cm 2cm 2.5cm 2cm;
                    background-color: #ffffff; 
                    
                    @top-center {{ 
                        content: element(pageHeader); 
                        width: 100%; 
                    }}
                    @bottom-center {{ 
                        content: counter(page); 
                        font-family: 'Arial', sans-serif; 
                    }}
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
                    background-color: #0A192F; /* Very dark slate/navy */
                    box-sizing: border-box;
                    padding: 1.5cm; /* Outer margin */
                    text-align: center;
                }}

                .cover-bg {{
                    position: absolute;
                    top: 0; left: 0; right: 0; bottom: 0;
                    {bg_css_rule}
                    background-size: cover;
                    background-position: center;
                    filter: blur(5px);
                    opacity: 0.15; 
                    z-index: 1;
                }}

                /* Using Padding instead of Absolute Positioning keeps logos inside */
                .cover-border {{
                    position: relative;
                    width: 100%; 
                    height: 100%;
                    border: 4px solid #D4AF37; /* Gold */
                    outline: 1px solid #D4AF37;
                    outline-offset: -10px;
                    z-index: 2;
                    box-sizing: border-box;
                    padding: 2cm;
                }}

                .cover-logos-table {{
                    width: 100%;
                    border-collapse: collapse;
                }}
                
                .c-logo {{
                    height: 110px;
                    background-color: #ffffff; 
                    border-radius: 50%; 
                    padding: 5px;
                }}

                .cover-title-box {{
                    background-color: rgba(10, 25, 47, 0.85);
                    border: 2px solid #D4AF37;
                    padding: 40px 20px;
                    margin-top: 3cm;
                    margin-bottom: 5cm;
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
                    width: 100%;
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
                
            </style>
        </head>
        <body>
            <!-- Header -->
            <div id="page-header">
                <table class="header-table">
                    <tr>
                        <td class="h-left"><img src="{amc_logo_b64}" class="hdr-logo" alt="AMC Logo"></td>
                        <td class="h-center">રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC</td>
                        <td class="h-right"><img src="{ntep_logo_b64}" class="hdr-logo" alt="NTEP Logo"></td>
                    </tr>
                </table>
            </div>
            
            <!-- Classic Cover Page -->
            <div class="cover-page">
                <div class="cover-bg"></div>
                <div class="cover-border">
                    
                    <!-- Logos Table inside the Border ensures they never bleed out -->
                    <table class="cover-logos-table">
                        <tr>
                            <td style="text-align: left;"><img src="{amc_logo_b64}" class="c-logo" alt="AMC Logo"></td>
                            <td style="text-align: right;"><img src="{ntep_logo_b64}" class="c-logo" alt="NTEP Logo"></td>
                        </tr>
                    </table>
                    
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
        # Added base_url="." to prevent Weasyprint from crashing on empty URLs
        pdf_bytes = weasyprint.HTML(string=full_html, base_url=".").write_pdf()

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
