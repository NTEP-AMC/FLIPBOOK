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
                    color: #000; 
                    text-align: justify;
                }}
                
                /* ----------------------------------------------------- */
                /* 1. RUNNING HEADER (For content pages only)            */
                /* ----------------------------------------------------- */
                header {{ 
                    position: running(pageHeader); /* Turns this into a reusable element */
                    width: 100%;
                    border-bottom: 2px solid #004B87;
                    padding-bottom: 10px;
                    margin-bottom: 20px;
                }}
                .logo-left {{ float: left; height: 60px; max-width: 130px; object-fit: contain; }}
                .logo-right {{ float: right; height: 60px; max-width: 130px; object-fit: contain; }}
                .header-title {{ 
                    text-align: center; 
                    font-weight: bold; 
                    padding-top: 15px; 
                    font-size: 16px; 
                    color: #004B87;
                }}

                /* ----------------------------------------------------- */
                /* 2. PAGE SETTINGS (The Light Yellow Theme)             */
                /* ----------------------------------------------------- */
                @page {{
                    size: A4;
                    margin: 3.5cm 2cm 2cm 2cm;
                    background-color: #FFFAEC; /* Soft professional light yellow/cream */
                    
                    /* Inject the running header into the top margin */
                    @top-center {{ content: element(pageHeader); }}
                    
                    @bottom-center {{ 
                        content: counter(page); 
                        font-family: 'Arial', sans-serif;
                    }}
                }}

                /* ----------------------------------------------------- */
                /* 3. COVER PAGE DESIGN (Blue & Yellow)                  */
                /* ----------------------------------------------------- */
                /* Create a special 'named page' for the cover so it ignores margins and headers */
                @page cover {{
                    margin: 0cm; 
                    background: linear-gradient(135deg, #004B87 60%, #FFC000 60%); /* Crisp Blue/Yellow diagonal split */
                    @top-center {{ content: none; }} /* Hide header */
                    @bottom-center {{ content: none; }} /* Hide page number */
                }}

                .cover-container {{
                    page: cover; /* Apply the special page settings */
                    page-break-after: always;
                    height: 29.7cm; /* Full A4 height */
                    text-align: center;
                    color: white;
                    font-family: 'Arial', sans-serif; /* English titles look better in Arial */
                    padding-top: 6cm;
                    box-sizing: border-box;
                }}

                .cover-logos {{ margin-bottom: 40px; }}
                
                /* Add a white circle background to logos so they pop against the dark blue */
                .cover-logos img {{ 
                    height: 120px; 
                    margin: 0 20px; 
                    background-color: white; 
                    padding: 15px; 
                    border-radius: 50%; 
                    box-shadow: 0px 4px 10px rgba(0,0,0,0.3);
                }}

                .cover-title {{ 
                    font-size: 55px; 
                    font-weight: bold; 
                    text-transform: uppercase; 
                    letter-spacing: 2px;
                    margin-bottom: 10px;
                    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
                }}
                
                .cover-subtitle {{ 
                    font-size: 24px; 
                    margin-top: 80px; 
                    color: #333; /* Dark text for the yellow portion of the background */
                    font-weight: bold;
                }}

                /* ----------------------------------------------------- */
                /* 4. CONTENT FORMATTING                                 */
                /* ----------------------------------------------------- */
                .content h1 {{ 
                    page-break-before: always; 
                    color: #004B87;
                    border-bottom: 2px solid #FFC000; 
                    padding-bottom: 5px;
                }}
                .content h2, .content h3 {{ color: #333; font-weight: bold; }}
                .content ul, .content ol {{ margin-left: 20px; padding-left: 10px; }}
                .content li {{ margin-bottom: 8px; }}
                .content table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                .content th, .content td {{ border: 1px solid #000; padding: 8px; text-align: left; background-color: white; }}
                .content th {{ background-color: #004B87; color: white; font-weight: bold; }}
                
            </style>
        </head>
        <body>
            <!-- The Running Header (Hidden on cover, visible on content pages) -->
            <header>
                <img src="{ntep_logo_b64}" class="logo-left" alt="NTEP Logo">
                <img src="{amc_logo_b64}" class="logo-right" alt="AMC Logo">
                <div class="header-title">રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP) - AMC</div>
            </header>
            
            <!-- The Unique Front Cover -->
            <div class="cover-container">
                <div class="cover-logos">
                    <img src="{ntep_logo_b64}" alt="NTEP Logo">
                    <img src="{amc_logo_b64}" alt="AMC Logo">
                </div>
                <div class="cover-title">Public Health Action<br>NTEP</div>
                <div class="cover-subtitle">Presented by<br>Ahmedabad Municipal Corporation</div>
            </div>
            
            <!-- The Extracted Word Document Content -->
            <div class="content">
                {str(soup)}
            </div>
        </body>
        </html>
        """

    with st.spinner("Generating PDF (Rendering design and fonts)..."):
        pdf_bytes = weasyprint.HTML(string=full_html).write_pdf()

    st.success("Manual Generated Successfully!")
    
    st.download_button(
        label="📄 Download Official PDF Manual",
        data=pdf_bytes,
        file_name="AMC_NTEP_Operational_Manual.pdf",
        mime="application/pdf"
    )
