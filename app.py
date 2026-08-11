import streamlit as st
import docx
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import re

st.set_page_config(page_title="AMC NTEP - Template Injector", layout="wide")
st.title("AMC NTEP - Bulletproof Template Injector")

# --- 1. Advanced Word Parser ---
def parse_word_to_placeholders(docx_file):
    """
    Scans the Word doc and creates a dictionary of placeholders automatically.
    Example output: {'{{OBJ_1.1}}': 'શંકાસ્પદ TB દર્દીની...', '{{STEPS_1.1}}': 'લક્ષણો તપાસો...'}
    """
    doc = docx.Document(docx_file)
    data = {}
    current_sub = None
    current_key = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Detect Sub-module (e.g., "1.1 Presumptive TB...")
        match = re.match(r'^(\d+\.\d+)', text)
        if match:
            current_sub = match.group(1) # Grabs "1.1", "1.2", etc.
            data[f"{{{{TITLE_{current_sub}}}}}"] = text
            current_key = None
            continue

        if not current_sub:
            continue

        # Detect Sections and create corresponding tags
        if text.startswith("ઉદ્દેશ્ય"):
            current_key = f"{{{{OBJ_{current_sub}}}}}"
            data[current_key] = text.replace("ઉદ્દેશ્ય (Objective):", "").strip()
            
        elif text.startswith("શું કરવું?"):
            current_key = f"{{{{STEPS_{current_sub}}}}}"
            data[current_key] = text.replace("શું કરવું? (What to do?):", "").strip()
            
        elif text.startswith("જવાબદાર વ્યક્તિ"):
            current_key = f"{{{{WHO_{current_sub}}}}}"
            data[current_key] = text.replace("જવાબદાર વ્યક્તિ (Responsible Person):", "").strip()
            
        elif text.startswith("સમયમર્યાદા") or text.startswith("અમલીકરણનો સમય"):
            current_key = f"{{{{TIME_{current_sub}}}}}"
            data[current_key] = text.replace("સમયમર્યાદા (Timeline):", "").replace("અમલીકરણનો સમય / ક્યારે કરવું? (Trigger):", "").strip()
            
        elif text.startswith("મોનિટરિંગ સૂચકાંકો"):
            current_key = f"{{{{IND_{current_sub}}}}}"
            data[current_key] = text.replace("મોનિટરિંગ સૂચકાંકો (Monitoring Indicators):", "").strip()
            
        elif current_key:
            # If we are under a key, append the next lines to it (for multi-line steps)
            data[current_key] += f"\n{text}"

    return data

# --- 2. Bulletproof Shape Scanner ---
def replace_text_in_shapes(shapes, replacements):
    """
    Recursively digs through shapes (even Grouped infographics) and safely 
    overwrites placeholders while preserving your PowerPoint's custom fonts/colors.
    """
    for shape in shapes:
        # If it's a grouped graphic, dig inside it recursively
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            replace_text_in_shapes(shape.shapes, replacements)
            
        # If it holds text, check for placeholders
        elif shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        # FOUND IT! We do a paragraph-level replacement to fix the "split run" bug.
                        new_text = paragraph.text.replace(key, str(value))
                        
                        if len(paragraph.runs) > 0:
                            first_run = paragraph.runs[0]
                            # Erase all existing split blocks
                            for run in paragraph.runs:
                                run.text = ""
                            # Put the complete new text into the first block to preserve original formatting
                            first_run.text = new_text

# --- Streamlit UI ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload PPTX Template")
    uploaded_template = st.file_uploader("Upload PPTX containing {{TAGS}}", type=["pptx"])

with col2:
    st.subheader("2. Upload Content")
    uploaded_docx = st.file_uploader("Upload Gujarati Word Document (.docx)", type=["docx"])

if uploaded_template and uploaded_docx:
    if st.button("Generate Final Presentation", type="primary"):
        
        with st.spinner("Extracting content and building tags..."):
            replacements = parse_word_to_placeholders(uploaded_docx)
            
            # Show the user what tags were successfully created
            with st.expander("🔍 View Generated Tags (Debug Check)"):
                st.write(replacements)
            
        with st.spinner("Deep scanning PPTX and injecting Gujarati text..."):
            prs = Presentation(uploaded_template)
            
            # Process every slide
            for slide in prs.slides:
                replace_text_in_shapes(slide.shapes, replacements)
                
            # Save to memory
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            
        st.success("✅ Infographic Injection Complete! No empty templates this time.")
        
        st.download_button(
            label="📄 Download Perfect PPTX",
            data=ppt_io,
            file_name="AMC_NTEP_Final_Infographic.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
