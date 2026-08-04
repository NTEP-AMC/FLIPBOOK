import streamlit as st
import mammoth
import weasyprint
import base64
from bs4 import BeautifulSoup
import tempfile
import os

# --- Configuration ---
st.set_page_config(page_title="AMC NTEP Flipbook Generator", layout="wide")
st.title("AMC NTEP - Automated Flipbook & PDF Generator")

# --- UI: File Uploads ---
st.sidebar.header("Upload Assets")
uploaded_docx = st.sidebar.file_uploader("Upload Content Word Document (.docx)", type=["docx"])
uploaded_logo = st.sidebar.file_uploader("Upload AMC NTEP Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
theme_color = st.sidebar.color_picker("Pick a Theme Color", "#004B87") # Default blue

if uploaded_docx is not None:
    # --- Step 1: Convert DOCX to HTML using Mammoth ---
    # Mammoth focuses on semantic conversion, keeping your text exactly as is but stripping messy Word styles.
    with st.spinner("Extracting content from Word Document..."):
        # We need a temporary file because mammoth expects a file-like object
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(uploaded_docx.read())
            tmp_docx_path = tmp_docx.name

        with open(tmp_docx_path, "rb") as docx_file:
            # You can define custom style maps here if your word doc has specific styles
            # style_map = "p[style-name='Heading 1'] => h1.module-title:fresh"
            result = mammoth.convert_to_html(docx_file)
            raw_html = result.value
        
        os.remove(tmp_docx_path) # Cleanup

    # --- Step 2: Inject Styling, Cover, and Logos (HTML Manipulation) ---
    with st.spinner("Applying theme and structuring booklet..."):
        soup = BeautifulSoup(raw_html, 'html.parser')
        
        # 1. Handle Logo
        logo_html = ""
        if uploaded_logo:
             logo_bytes = uploaded_logo.read()
             logo_b64 = base64.b64encode(logo_bytes).decode()
             logo_mime = uploaded_logo.type
             logo_html = f'<img src="data:{logo_mime};base64,{logo_b64}" class="cover-logo" alt="AMC NTEP Logo">'

        # 2. Build the full HTML Document
        # We use CSS for pagination (@page), page breaks, and styling
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                /* Base Booklet Styles */
                body {{ font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; }}
                
                /* Print & PDF Settings (WeasyPrint uses these) */
                @page {{
                    size: A4;
                    margin: 2cm;
                    @bottom-center {{ content: counter(page); }}
                }}
                
                /* Cover Page Styling */
                .cover-page {{
                    text-align: center;
                    page-break-after: always;
                    padding-top: 100px;
                }}
                .cover-logo {{ max-width: 250px; margin-bottom: 30px; }}
                .cover-title {{ font-size: 3em; color: {theme_color}; margin-bottom: 20px; }}
                
                /* Module & Content Styling */
                h1 {{ color: {theme_color}; page-break-before: always; border-bottom: 2px solid {theme_color}; padding-bottom: 10px; }}
                h2 {{ color: #444; margin-top: 30px; }}
                
                /* Ensure tables look good */
                table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: {theme_color}; color: white; }}
                
                /* Custom Index (Placeholder - building a dynamic index requires parsing the DOM) */
                .index-page {{ page-break-after: always; }}
            </style>
        </head>
        <body>
            <!-- Generated Cover Page -->
            <div class="cover-page">
                {logo_html}
                <h1 class="cover-title">AMC NTEP Operational Booklet</h1>
                <h2>Departmental Guidelines & Modules</h2>
            </div>
            
            <!-- (Optional) You would generate an index here by finding all h1/h2 tags -->
            
            <!-- The Extracted Content -->
            <div class="content">
                {str(soup)}
            </div>
        </body>
        </html>
        """

    # --- Step 3: Generate PDF using WeasyPrint ---
    with st.spinner("Generating PDF..."):
        pdf_bytes = weasyprint.HTML(string=full_html).write_pdf()

    # --- Step 4: Display & Download ---
    st.success("Booklet Generated Successfully!")
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📄 Download PDF Booklet",
            data=pdf_bytes,
            file_name="AMC_NTEP_Booklet.pdf",
            mime="application/pdf"
        )
    
    # --- Step 5: The Flipbook Integration (The Tricky Part) ---
    st.header("Interactive Flipbook Preview")
    
    # Note on Flipbooks in Streamlit:
    # Streamlit doesn't have a native 'flipbook' widget. You have to embed HTML/JS.
    # The most robust way is to use Streamlit's components.html to inject a library like turn.js.
    # However, Turn.js expects pages to be separate DIVs, not a continuous scroll.
    # We must format our HTML specifically for the flipbook structure.
    
    flipbook_html = f"""
    <html>
    <head>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/turn.js/3/turn.min.js"></script>
        <style>
            #flipbook {{ width: 800px; height: 600px; margin: 0 auto; }}
            #flipbook .page {{ background: white; border: 1px solid #ccc; padding: 20px; overflow: hidden; }}
            /* Add styles to handle mammoth's continuous HTML into pages (requires complex JS pagination) */
        </style>
    </head>
    <body>
        <div id="flipbook">
            <!-- In a production app, you need logic to split the 'full_html' into distinct page DIVs here -->
            <div class="page">Cover Page<br>{logo_html}</div>
            <div class="page">Page 1 Content...</div>
            <div class="page">Page 2 Content...</div>
        </div>
        <script>
            $("#flipbook").turn({{ width: 800, height: 600, autoCenter: true }});
        </script>
    </body>
    </html>
    """
    
    import streamlit.components.v1 as components
    # components.html(flipbook_html, height=650)
    st.info("To render the true 3D flipbook, the continuous HTML must be paginated into images or distinct DIVs. The PDF is ready for download above.")

else:
    st.info("Please upload a .docx file to begin.")
