import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

st.set_page_config(page_title="AMC NTEP PPT Generator", layout="wide")
st.title("AMC NTEP - Professional Bilingual PPT Generator")
st.markdown("Upload your Gujarati/English Word document. The app will extract the structured sections and generate a formatted PowerPoint.")

# --- Helper Function: Smart Parser for Bilingual Text ---
def parse_ntep_document(docx_file):
    doc = docx.Document(docx_file)
    slides_data = []
    current_slide = None
    
    # Keywords to identify sections in your specific document
    keywords = [
        "ઉદ્દેશ્ય", "અમલીકરણનો સમય", "શું કરવું?", "શા માટે?", 
        "લક્ષિત જૂથ", "જવાબદાર વ્યક્તિ", "સમયમર્યાદા", "દસ્તાવેજીકરણ", 
        "મોનિટરિંગ સૂચકાંકો", "સુપરવાઈઝર ગુણવત્તા", "જો કામગીરી ન થાય"
    ]
    
    current_key = "intro"
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        # Detect New Module or Sub-Module (e.g., "1.1 Presumptive TB...")
        if re.match(r'^(Module|\d+\.\d+)', text, re.IGNORECASE):
            if current_slide:
                slides_data.append(current_slide)
            current_slide = {"title": text, "content": {}}
            current_key = "intro"
            continue
            
        if current_slide is None:
            current_slide = {"title": "પ્રસ્તાવના (Introduction)", "content": {}}
            
        # Detect if paragraph is one of the target headers
        found_key = False
        for key in keywords:
            if text.startswith(key) or text.startswith(f"{key} ("):
                current_key = key
                current_slide["content"][current_key] = text
                found_key = True
                break
                
        # If it's not a header, append it to the current active key
        if not found_key:
            if current_key not in current_slide["content"]:
                current_slide["content"][current_key] = text
            else:
                current_slide["content"][current_key] += f"\n{text}"
                
    if current_slide:
        slides_data.append(current_slide)
        
    return slides_data

# --- Helper Function: Set Formatting (Gujarati Support) ---
def format_text(run, font_size, bold=False, color=None):
    # Nirmala UI or Shruti are standard Windows fonts that support Gujarati well
    run.font.name = 'Nirmala UI' 
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color

# --- Helper Function: Generate Structured PPTX ---
def generate_ppt(slides_data):
    prs = Presentation()
    
    # Brand Colors based on your references
    navy_blue = RGBColor(10, 47, 81)
    teal = RGBColor(32, 163, 158)
    dark_gray = RGBColor(60, 60, 60)
    
    # 1. Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "રાષ્ટ્રીય ક્ષયરોગ નિવારણ કાર્યક્રમ (NTEP)\nજાહેર આરોગ્ય કાર્યવાહી (Public Health Actions)"
    format_text(title.text_frame.paragraphs[0].runs[0], 36, True, navy_blue)
    
    subtitle.text = "ઓપરેશનલ માર્ગદર્શિકા (Operational Manual)\nAhmedabad Municipal Corporation"
    format_text(subtitle.text_frame.paragraphs[0].runs[0], 20, False, dark_gray)
    
    # 2. Content Slides based on parsed data
    content_slide_layout = prs.slide_layouts[1] # Title and Content
    
    for data in slides_data:
        slide = prs.slides.add_slide(content_slide_layout)
        title_shape = slide.shapes.title
        body_shape = slide.placeholders[1]
        
        # Set Title
        title_shape.text = data["title"]
        format_text(title_shape.text_frame.paragraphs[0].runs[0], 28, True, navy_blue)
        
        # Set Content Layout
        tf = body_shape.text_frame
        tf.clear() # Clear default formatting
        
        for key, text_content in data["content"].items():
            # Add Header (e.g., "શું કરવું?")
            p_header = tf.add_paragraph()
            run_header = p_header.add_run()
            run_header.text = f"{key.upper()}: "
            format_text(run_header, 16, True, teal)
            
            # Add Body text for that section
            cleaned_text = text_content.replace(key, "").replace(":", "", 1).strip()
            if cleaned_text:
                run_body = p_header.add_run()
                run_body.text = cleaned_text
                format_text(run_body, 14, False, dark_gray)
            
            # Add some spacing
            p_header.space_after = Pt(12)

    # Save to memory buffer
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- Main App Execution ---
uploaded_docx = st.file_uploader("Upload Bilingual Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:
    with st.spinner("Parsing Gujarati/English Document Structure..."):
        parsed_data = parse_ntep_document(uploaded_docx)
        
    with st.spinner("Generating Structured PowerPoint..."):
        ppt_file = generate_ppt(parsed_data)
        
    st.success("PowerPoint Generated Successfully!")
    
    st.download_button(
        label="📊 Download Professional PPTX",
        data=ppt_file,
        file_name="AMC_NTEP_Operational_Manual.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
