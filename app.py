import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io
import re

st.set_page_config(page_title="AMC NTEP Smart Summarizer", layout="wide")
st.title("AMC NTEP - Module-Wise Infographic Generator")
st.markdown("Upload your 70-page document. This app will skim the content, group it by **Module**, and generate a high-impact, 1-page dashboard per module.")

# --- Theme Colors ---
NAVY_BLUE = RGBColor(10, 47, 81)
TEAL_GREEN = RGBColor(32, 163, 158)
LIGHT_BG = RGBColor(245, 247, 250)
WHITE = RGBColor(255, 255, 255)
DARK_GRAY = RGBColor(60, 60, 60)
ACCENT_ORANGE = RGBColor(230, 126, 34)

# --- Helper: Apply Font Formatting ---
def format_text(run, font_size, bold=False, color=DARK_GRAY):
    run.font.name = 'Nirmala UI' # Supports Gujarati
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color

# --- Smart Parser: Group by Module & Sub-Module ---
def extract_module_summaries(docx_file):
    doc = docx.Document(docx_file)
    modules = []
    current_module = None
    current_sub = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        # Detect Module (e.g., "Module 1: At Diagnosis")
        if re.match(r'^Module\s*\d+', text, re.IGNORECASE):
            if current_module:
                modules.append(current_module)
            current_module = {"title": text, "sub_modules": []}
            current_sub = None
            continue
            
        # Detect Sub-Module (e.g., "1.1 Presumptive TB Evaluation...")
        if current_module and re.match(r'^\d+\.\d+', text):
            current_sub = {"title": text, "summary": ""}
            current_module["sub_modules"].append(current_sub)
            continue
            
        # Grab the first meaningful sentence as a summary for the infographic card
        if current_sub and not current_sub["summary"]:
            # Ignore headers like "ઉદ્દેશ્ય (Objective):" and grab actual text
            if len(text) > 20 and not text.endswith(":"):
                # Truncate to keep the card clean (approx 120 chars)
                summary = text[:120] + "..." if len(text) > 120 else text
                current_sub["summary"] = summary
                
    if current_module:
        modules.append(current_module)
        
    return modules

# --- Draw Infographic Grid ---
def draw_module_slide(slide, module_data):
    # 1. Draw Module Header
    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.4), Inches(12.33), Inches(1))
    header.fill.solid()
    header.fill.fore_color.rgb = NAVY_BLUE
    header.line.color.rgb = TEAL_GREEN
    header.line.width = Pt(2)
    
    tf_header = header.text_frame
    tf_header.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_header.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = module_data["title"].upper()
    format_text(run, 24, bold=True, color=WHITE)

    # 2. Draw Sub-Module Grid (Max 6 per slide for good UI, 3 columns x 2 rows)
    col_width = Inches(3.8)
    row_height = Inches(2.2)
    start_x = Inches(0.7)
    start_y = Inches(1.8)
    
    for idx, sub in enumerate(module_data["sub_modules"][:6]): # Limiting to 6 for clean layout
        row = idx // 3
        col = idx % 3
        
        box_x = start_x + (col * (col_width + Inches(0.3)))
        box_y = start_y + (row * (row_height + Inches(0.4)))
        
        # Background Card
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, box_x, box_y, col_width, row_height)
        card.fill.solid()
        card.fill.fore_color.rgb = LIGHT_BG
        card.line.color.rgb = TEAL_GREEN
        card.line.width = Pt(1.5)
        
        # Card Header Banner
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, box_x, box_y, col_width, Inches(0.6))
        banner.fill.solid()
        banner.fill.fore_color.rgb = TEAL_GREEN
        banner.line.fill.background()
        
        tf_banner = banner.text_frame
        tf_banner.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_banner = tf_banner.paragraphs[0]
        run_banner = p_banner.add_run()
        # Extract just the number and short title
        short_title = sub["title"][:45] + "..." if len(sub["title"]) > 45 else sub["title"]
        run_banner.text = short_title
        format_text(run_banner, 12, bold=True, color=WHITE)
        
        # Card Summary Text
        txBox = slide.shapes.add_textbox(box_x + Inches(0.1), box_y + Inches(0.7), col_width - Inches(0.2), row_height - Inches(0.8))
        tf_content = txBox.text_frame
        tf_content.word_wrap = True
        p_content = tf_content.paragraphs[0]
        run_content = p_content.add_run()
        run_content.text = sub["summary"]
        format_text(run_content, 11, bold=False, color=DARK_GRAY)

# --- Main PPTX Generator ---
def generate_summary_ppt(modules):
    prs = Presentation()
    # Use 16:9 Widescreen for modern infographics
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # Create Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = title_slide.shapes.title
    title.text = "NTEP Public Health Actions\nModule Summaries"
    format_text(title.text_frame.paragraphs[0].runs[0], 36, True, NAVY_BLUE)
    
    # Create Module Slides
    for mod in modules:
        slide = prs.slides.add_slide(blank_layout)
        draw_module_slide(slide, mod)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- Streamlit UI ---
uploaded_docx = st.file_uploader("Upload 70-Page Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:
    with st.spinner("Skimming document & extracting Module highlights..."):
        modules_data = extract_module_summaries(uploaded_docx)
        
        if not modules_data:
            st.error("Could not find 'Module' headers. Make sure your Word document uses 'Module 1:', 'Module 2:' format.")
        else:
            with st.spinner(f"Drawing Infographic Dashboards for {len(modules_data)} Modules..."):
                ppt_file = generate_summary_ppt(modules_data)
                
            st.success("Condensed Infographic PowerPoint Generated!")
            
            st.download_button(
                label="📊 Download Module-Wise PPTX",
                data=ppt_file,
                file_name="AMC_NTEP_Module_Summary.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
