import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io
import re

st.set_page_config(page_title="AMC NTEP Infographic PPT", layout="wide")
st.title("AMC NTEP - Advanced Infographic Layout Generator")

# --- Theme Colors ---
NAVY_BLUE = RGBColor(10, 47, 81)
TEAL_GREEN = RGBColor(32, 163, 158)
LIGHT_BG = RGBColor(245, 247, 250)
WHITE = RGBColor(255, 255, 255)
DARK_GRAY = RGBColor(60, 60, 60)

# --- Helper: Apply Gujarati Font ---
def format_text(run, font_size, bold=False, color=DARK_GRAY):
    run.font.name = 'Nirmala UI'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color

# --- Helper: Draw Styled Infographic Box ---
def draw_info_box(slide, left, top, width, height, title, content, border_color=TEAL_GREEN):
    # Background Shape
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.color.rgb = border_color
    shape.line.width = Pt(2)
    
    # Title Banner inside the box
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.4))
    banner.fill.solid()
    banner.fill.fore_color.rgb = border_color
    banner.line.fill.background()
    
    tf_banner = banner.text_frame
    p = tf_banner.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title
    format_text(run, 14, bold=True, color=WHITE)
    
    # Content Text Box
    txBox = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.5), width - Inches(0.2), height - Inches(0.6))
    tf = txBox.text_frame
    tf.word_wrap = True
    p_content = tf.paragraphs[0]
    run_content = p_content.add_run()
    run_content.text = content
    format_text(run_content, 12, bold=False, color=DARK_GRAY)

# --- Helper: Draw Workflow Steps (Circles and Arrows) ---
def draw_workflow_step(slide, step_num, text, left, top):
    # Number Circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, Inches(0.5), Inches(0.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = NAVY_BLUE
    circle.line.fill.background()
    tf_circle = circle.text_frame
    tf_circle.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_circ = tf_circle.paragraphs[0]
    p_circ.alignment = PP_ALIGN.CENTER
    run_circ = p_circ.add_run()
    run_circ.text = str(step_num)
    format_text(run_circ, 14, bold=True, color=WHITE)
    
    # Text Box next to circle
    txBox = slide.shapes.add_textbox(left + Inches(0.6), top, Inches(4), Inches(0.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p_text = tf.paragraphs[0]
    run_text = p_text.add_run()
    run_text.text = text
    format_text(run_text, 11, bold=False, color=DARK_GRAY)
    
    # Down Arrow (if not the first step, draw above)
    if step_num > 1:
        arrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, left + Inches(0.15), top - Inches(0.35), Inches(0.2), Inches(0.25))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = TEAL_GREEN
        arrow.line.fill.background()

# --- Main PPT Generator ---
def generate_infographic_ppt(docx_file):
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # --- Dummy Data Extraction (Simulating your parser) ---
    # In a full setup, this would parse your Word doc dynamically.
    # Here, we use structure to build the exact visual you want.
    slide = prs.slides.add_slide(blank_layout)
    
    # 1. Main Header Banner
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1))
    header.fill.solid()
    header.fill.fore_color.rgb = NAVY_BLUE
    tf = header.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "PHA 1.1: PRESUMPTIVE TB EVALUATION & MICROBIOLOGICAL CONFIRMATION"
    format_text(run, 18, bold=True, color=WHITE)

    # 2. Left Column (Objectives & Details)
    draw_info_box(slide, left=Inches(0.5), top=Inches(1.5), width=Inches(4), height=Inches(2), 
                  title="OBJECTIVE (ઉદ્દેશ્ય)", 
                  content="શંકાસ્પદ TB દર્દીની યોગ્ય તપાસ કરી માઇક્રોબાયોલોજિકલ પુષ્ટિ પ્રાપ્ત કરવી.")
                  
    draw_info_box(slide, left=Inches(0.5), top=Inches(3.8), width=Inches(4), height=Inches(1.5), 
                  title="WHO (જવાબદાર વ્યક્તિ)", 
                  content="MO / STS / Lab Technician / CHO / Staff Nurse", border_color=NAVY_BLUE)

    draw_info_box(slide, left=Inches(0.5), top=Inches(5.6), width=Inches(4), height=Inches(1.5), 
                  title="TIMELINE (સમયમર્યાદા)", 
                  content="પ્રથમ મુલાકાત દરમિયાન (૦-૨૪ કલાકમાં તપાસ શરૂ કરવી).", border_color=RGBColor(200, 80, 80))

    # 3. Right Column (Workflow Flowchart)
    # Background for Workflow
    workflow_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(1.5), Inches(4.8), Inches(5.6))
    workflow_bg.fill.solid()
    workflow_bg.fill.fore_color.rgb = LIGHT_BG
    workflow_bg.line.color.rgb = TEAL_GREEN
    
    wf_title = slide.shapes.add_textbox(Inches(5), Inches(1.6), Inches(4.4), Inches(0.5))
    run_wft = wf_title.text_frame.paragraphs[0].add_run()
    run_wft.text = "STEP-BY-STEP WORKFLOW"
    format_text(run_wft, 14, bold=True, color=NAVY_BLUE)

    # Drawing the steps mathematically
    steps = [
        "શંકાસ્પદ TB દર્દીનું ઇતિહાસ, લક્ષણો (ખાંસી, તાવ) પૂછો.",
        "શારીરિક તપાસ કરો.",
        "સંપૂર્ણ ગુણવત્તાવાળા નમૂના માટે માર્ગદર્શન આપો.",
        "CBNAAT / Truenat / Cartridge આધારિત NAAT દ્વારા તપાસ કરો.",
        "પરિણામની પુષ્ટિ કરો (TB Detected / Not Detected)."
    ]
    
    start_top = 2.2
    for idx, text in enumerate(steps):
        draw_workflow_step(slide, step_num=idx+1, text=text, left=Inches(5), top=Inches(start_top))
        start_top += 0.85 # Space between steps

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- Streamlit UI ---
uploaded_docx = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if uploaded_docx is not None:
    with st.spinner("Drawing Flowcharts and Infographic Layouts..."):
        ppt_file = generate_infographic_ppt(uploaded_docx)
        
    st.success("Infographic PowerPoint Generated Successfully!")
    
    st.download_button(
        label="📊 Download Flowchart/Infographic PPTX",
        data=ppt_file,
        file_name="AMC_NTEP_Visual_Infographic.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
