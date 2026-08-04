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

def get_image_base64(filepath):
    if os.path.exists(filepath):
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            mime_type = "image/png" if filepath.lower().endswith(".png") else "image/jpeg"
            return f"data:{mime_type};base64,{encoded_string}"
    return ""

amc_logo_b64 = get_image_base64("Amdavad_Municipal_Corporation_logo.png")
ntep_logo_b64 = get_image_base64("1-s2.0-S0019570720303152-gr1.jpg") 

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

    with st.spinner("Applying Official Government Theme..."):
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
                /* RUNNING HEADER (AMC Left, NTEP Right)                 */
                /* ----------------------------------------------------- */
                header {{ 
                    position: running(pageHeader); 
                    width: 100%;
                    margin-bottom: 20px;
                }}
                
                .header-table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    border-bottom: 2px solid #000; 
                }}
                .header-table td {{ vertical-align: middle; padding-bottom: 10px; }}
                
                .header-left {{ width: 20%; text-align: left; }}
                .header-left img {{ height: 65px; max-width: 100px; object-fit: contain; }}
                
                .header-center {{ width: 60%; text-align: center; font-weight: bold; font-size: 16px; color: #000; }}
                
                .header-right {{ width: 20%; text-align: right; }}
                .header-right img {{ height: 65px; max-width: 100px; object-fit: contain; }}

                /* ----------------------------------------------------- */
                /* PAGE SETTINGS                                         */
                /* ----------------------------------------------------- */
                @page {{
                    size: A4;
                    margin: 3.5cm 2.5cm 2.5cm 2.5cm;
                    background-color: #ffffff; 
                    
                    @top-center {{ content: element(pageHeader); }}
                    @bottom-center {{ content: counter(page); font-family: 'Arial', sans-serif; }}
                }}

                /* ----------------------------------------------------- */
                /* FORMAL GOVERNMENT COVER PAGE                          */
                /* ----------------------------------------------------- */
                @page cover {{
                    margin: 0cm; 
                    background-color: #ffffff; 
                    @top-center {{ content: none; }} 
                    @bottom-center {{ content: none; }} 
                }}

                .cover-container {{
                    page: cover; 
                    page-break-after: always;
                    height: 29.7cm; 
                    width: 21cm;
                    padding: 3cm; 
                    box-sizing: border-box;
                    font-family: 'Arial', sans-serif; 
                    /* Thick, authoritative double border */
                    border: 12px solid #004B87;
                    outline: 2px solid #004B87;
                    outline-offset: -20px;
                    text-align: center;
                    position: relative;
                }}

                .cover-logos {{ 
                    width: 100%; 
                    margin-top: 1cm;
                    margin-bottom: 3cm; 
                }}
                .logo-amc {{ float: left; height: 130px; }}
                .logo-ntep {{ float: right; height: 130px; }}
                .cover-logos::after {{ content: ""; clear: both; display: table; }}

                .cover-department {{
                    font-size: 22px;
                    color: #333;
                    text-transform: uppercase;
                    margin-bottom: 2cm;
                    font-weight: bold;
                    letter-spacing: 1px;
                }}

                .cover-title {{ 
                    font-size: 55px; 
                    font-weight: 900; 
                    color: #000; 
                    text-transform: uppercase; 
                    border-top: 3px solid #004B87;
                    border-bottom: 3px solid #004B87;
                    padding: 30px 0;
                    margin-bottom: 2cm;
                    line-height: 1.2;
                }}
                
                .cover-subtitle {{ 
                    font-size: 26px; 
                    color: #004B87; 
                    font-weight: bold;
                }}
                
                .cover-footer {{
                    position: absolute;
                    bottom: 3cm;
                    left: 0;
                    width: 100%;
                    text-align: center;
                    font-size: 18px;
                    font-weight: bold;
                    color: #000;
                }}

                /* ----------------------------------------------------- */
                /* CONTENT FORMATTING                                    */
                /* ----------------------------------------------------- */
                .content h1 {{ page-break-before: always; color: #000; border-bottom: 2px solid #000; padding-bottom: 5px; margin-top: 0; }}
                .content h2, .content h3 {{ color: #000; font-weight: bold; margin-top: 25px; }}
                .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
                .content li {{ margin-bottom: 8px; }}
                .content table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                .content th, .content td {{ border: 1px solid #000; padding: 10px; text-align: left; background-color: white; }}
                .content th {{ background-color: #f2f2f2; color: #000; font-weight: bold; text-align: center; }}
                
            </style>
        </head>
        <body>
            <!-- Running Header -->
            <header>
                <table class="header-table">
                    <tr>
                        <td class="header-left"><img src="{amc_logo_b64}" alt="AMC Logo"></td>
                        <td class="header-center">રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC</td>
                        <td class="header-right"><img src="{ntep_logo_b64}" alt="NTEP Logo"></td>
                    </tr>
                </table>
            </header>
            
            <!-- Formal Government Cover Page -->
            <div class="cover-container">
                <div class="cover-logos">
                    <img src="{amc_logo_b64}" class="logo-amc" alt="AMC Logo">
                    <img src="{ntep_logo_b64}" class="logo-ntep" alt="NTEP Logo">
                </div>
                
                <div class="cover-department">
                    Ahmedabad Municipal Corporation<br>
                    National Tuberculosis Elimination Program
                </div>
                
                <div class="cover-title">
                    Public Health<br>Action
                </div>
                
                <div class="cover-subtitle">
                    Operational Guidelines
                </div>
                
                <div class="cover-footer">
                    Ahmedabad, Gujarat &copy; 2026
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
