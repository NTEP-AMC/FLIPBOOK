import streamlit as st
import mammoth
import weasyprint
import base64
from bs4 import BeautifulSoup
import tempfile
import os

st.set_page_config(page_title="AMC NTEP Manual Generator", layout="wide")
st.title("AMC NTEP - Official Booklet Generator")
st.write("Upload your Gujarati Word document to generate a formatted PDF manual.")

# --- Helper Function: Convert local image to Base64 ---
def get_image_base64(filepath):
    if os.path.exists(filepath):
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            mime_type = "image/png" if filepath.lower().endswith(".png") else "image/jpeg"
            return f"data:{mime_type};base64,{encoded_string}"
    return ""

# Read the logos you uploaded to GitHub
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

    with st.spinner("Applying Government Theme and Cover Page..."):
        soup = BeautifulSoup(raw_html, 'html.parser')

        # --- HTML & CSS Construction ---
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Gujarati:wght@400;700&display=swap');
                
                /* Base Styles */
                body {{ 
                    font-family: 'Noto Sans Gujarati', sans-serif; 
                    line-height: 1.6; 
                    color: #111; 
                    text-align: justify;
                }}
                
                /* ----------------------------------------------------- */
                /* 1. RUNNING HEADER (Table layout fixes the squishing)  */
                /* ----------------------------------------------------- */
                header {{ 
                    position: running(pageHeader); 
                    width: 100%;
                    margin-bottom: 20px;
                }}
                
                /* Using a table forces perfect left/center/right alignment in WeasyPrint */
                .header-table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    border-bottom: 2px solid #004B87; 
                }}
                .header-table td {{ vertical-align: middle; padding-bottom: 10px; }}
                
                .header-left {{ width: 20%; text-align: left; }}
                .header-left img {{ height: 55px; max-width: 100px; object-fit: contain; }}
                
                .header-center {{ width: 60%; text-align: center; font-weight: bold; font-size: 15px; color: #004B87; }}
                
                .header-right {{ width: 20%; text-align: right; }}
                .header-right img {{ height: 55px; max-width: 100px; object-fit: contain; }}

                /* ----------------------------------------------------- */
                /* 2. PAGE SETTINGS (The Light Yellow Theme)             */
                /* ----------------------------------------------------- */
                @page {{
                    size: A4;
                    margin: 3.5cm 2cm 2cm 2cm;
                    background-color: #FFFAEC; /* Soft professional light yellow/cream */
                    
                    @top-center {{ content: element(pageHeader); }}
                    @bottom-center {{ content: counter(page); font-family: 'Arial', sans-serif; }}
                }}

                /* ----------------------------------------------------- */
                /* 3. NEW COVER PAGE DESIGN (Corporate/Gov Report Style) */
                /* ----------------------------------------------------- */
                @page cover {{
                    margin: 0cm; 
                    background-color: #ffffff; /* Clean white background */
                    @top-center {{ content: none; }} 
                    @bottom-center {{ content: none; }} 
                }}

                .cover-container {{
                    page: cover; 
                    page-break-after: always;
                    height: 29.7cm; /* Full A4 height */
                    border-left: 45px solid #004B87; /* Thick authoritative blue spine */
                    padding: 3cm 2cm 2cm 4cm; /* Pushed in to clear the spine */
                    box-sizing: border-box;
                    font-family: 'Arial', sans-serif; 
                }}

                .cover-logos {{ 
                    text-align: right; /* Logos top right */
                    margin-bottom: 5cm; 
                }}
                .cover-logos img {{ 
                    height: 90px; 
                    margin-left: 25px; 
                }}

                .cover-title {{ 
                    font-size: 52px; 
                    font-weight: 900; 
                    color: #004B87; 
                    text-transform: uppercase; 
                    letter-spacing: 1px;
                    line-height: 1.1;
                    border-bottom: 5px solid #FFC000; /* Subtle gold accent line */
                    padding-bottom: 25px;
                    margin-bottom: 25px;
                }}
                
                .cover-subtitle {{ 
                    font-size: 22px; 
                    color: #444; 
                    font-weight: bold;
                }}
                
                .cover-footer {{
                    position: absolute;
                    bottom: 3cm;
                    font-size: 14px;
                    color: #666;
                }}

                /* ----------------------------------------------------- */
                /* 4. CONTENT FORMATTING                                 */
                /* ----------------------------------------------------- */
                .content h1 {{ 
                    page-break-before: always; 
                    color: #004B87;
                    border-bottom: 2px solid #FFC000; 
                    padding-bottom: 5px;
                    margin-top: 0;
                }}
                .content h2, .content h3 {{ color: #222; font-weight: bold; margin-top: 25px; }}
                .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
                .content li {{ margin-bottom: 8px; }}
                .content table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                .content th, .content td {{ border: 1px solid #333; padding: 10px; text-align: left; background-color: white; }}
                .content th {{ background-color: #004B87; color: white; font-weight: bold; text-align: center; }}
                
            </style>
        </head>
        <body>
            <!-- The Running Header using a Table -->
            <header>
                <table class="header-table">
                    <tr>
                        <td class="header-left"><img src="{ntep_logo_b64}" alt="NTEP Logo"></td>
                        <td class="header-center">રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC</td>
                        <td class="header-right"><img src="{amc_logo_b64}" alt="AMC Logo"></td>
                    </tr>
                </table>
            </header>
            
            <!-- The New Front Cover -->
            <div class="cover-container">
                <div class="cover-logos">
                    <img src="{ntep_logo_b64}" alt="NTEP Logo">
                    <img src="{amc_logo_b64}" alt="AMC Logo">
                </div>
                
                <div class="cover-title">
                    Public Health<br>Action
                </div>
                <div class="cover-subtitle">
                    National Tuberculosis Elimination Program<br>
                    Ahmedabad Municipal Corporation
                </div>
                
                <div class="cover-footer">
                    <strong>Operational Manual</strong> &copy; 2026
                </div>
            </div>
            
            <!-- The Extracted Word Document Content -->
            <div class="content">
                {str(soup)}
            </div>
        </body>
        </html>
        """

    with st.spinner("Generating High-Quality PDF..."):
        pdf_bytes = weasyprint.HTML(string=full_html).write_pdf()

    st.success("Manual Generated Successfully!")
    
    st.download_button(
        label="📄 Download Official PDF Manual",
        data=pdf_bytes,
        file_name="AMC_NTEP_Operational_Manual.pdf",
        mime="application/pdf"
    )
