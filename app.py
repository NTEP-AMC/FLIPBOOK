import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

st.set_page_config(page_title="AMC NTEP PPT Generator", layout="wide")
st.title("AMC NTEP - Professional Infographic PPT Generator")
st.markdown("Upload a `.docx` file. The app will split sections by **Headings** and generate a visually structured PowerPoint presentation.")

# --- Helper Function: Extract content from Word ---
def extract_content_from_docx(docx_file):
    doc = docx.Document(docx_file)
    slides_data = []
    current_slide = {"title": "Introduction", "content": []}
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        # If it's a Heading 1 or 2, treat it as a new slide trigger
        if para.style.name.startswith('Heading'):
            if current_slide["content"] or current_slide["title"] != "Introduction":
                slides_data.append(current_slide)
            current_slide = {"title": text, "content": []}
        else:
            current_slide["content"].append(text)
            
    if current_slide["content"]:
        slides_data.append(current_slide)
        
    return slides_data

# --- Helper Function: Create Infographic PPT ---
def create_ppt(slides_data):
    prs = Presentation()
    
    # Define custom colors (AMC / NTEP Theme)
    theme_dark_blue = RGBColor(10, 25, 47)
    theme_teal = RGBColor(32, 163, 158)
    theme_light_gray = RGBColor(240, 240, 240)
    
    # 1. Create Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "NTEP Public Health Actions"
    subtitle.text = "Operational Manual & Infographics\nAhmedabad Municipal Corporation"
    
    # 2. Create Content Slides
    blank_slide_layout = prs.slide_layouts[6] # Blank layout for custom drawing
    
    for data in slides_data:
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Draw Header Banner (Dark Blue)
        header_shape = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(10), Inches(1.2) # 1 is MSO_SHAPE.RECTANGLE
        )
        header_shape.fill.solid()
        header_shape.fill.fore_color.rgb = theme_dark_blue
        header_shape.line.fill.background()
        
        # Add Title Text to Banner
        txBox_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        tf_title = txBox_title.text_frame
        p_title = tf_title.paragraphs[0]
        p_title.text = data["title"]
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(255, 255, 255)
        
        # Draw Content Background Box (Light Gray for infographic feel)
        content_bg = slide.shapes.add_shape(
            1, Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
        )
        content_bg.fill.solid()
        content_bg.fill.fore_color.rgb = theme_light_gray
        content_bg.line.color.rgb = theme_teal
        content_bg.line.width = Pt(2)
        
        # Add Content Text
        txBox_content = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
        tf_content = txBox_content.text_frame
        tf_content.word_wrap = True
        
        for idx, paragraph_text in enumerate(data["content"]):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = paragraph_text
            p.font.size = Pt(16)
            p.space_after = Pt(14)

    # Save to memory
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- Streamlit UI ---
uploaded_docx = st.file_uploader("Upload Content Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:
    with st.spinner("Analyzing document structure..."):
        slides_data = extract_content_from_docx(uploaded_docx)
        
    with st.spinner("Generating Infographic PPTX..."):
        ppt_file = create_ppt(slides_data)
        
    st.success("PowerPoint Generated Successfully!")
    
    st.download_button(
        label="📊 Download Professional PPTX",
        data=ppt_file,
        file_name="AMC_NTEP_Infographic.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
