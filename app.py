import streamlit as st
import docx
from pptx import Presentation
import io

st.set_page_config(page_title="AMC NTEP PPT Generator", layout="wide")
st.title("AMC NTEP - Template-Based Infographic Generator")
st.markdown("""
**Instructions:**
1. Upload your designed **PowerPoint Template (.pptx)** containing placeholders like `{{OBJ_1.1}}`, `{{WHO_1.1}}`.
2. Upload your **Word Document (.docx)** containing the data.
3. The app will inject the text directly into your beautiful graphics!
""")

# --- Helper Function: Extract Data from Word ---
def extract_data_from_word(docx_file):
    """
    Reads the Word document and extracts data to map to placeholders.
    (This is a simplified mapper. You can adjust the keys based on your actual Word format).
    """
    doc = docx.Document(docx_file)
    extracted_data = {}
    
    current_key = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        # Example logic to map Word text to PPTX placeholders
        # In a real scenario, you might use regex to find sections like "ઉદ્દેશ્ય (Objective):"
        if "1.1 Presumptive TB" in text:
            current_key = "1.1"
        elif text.startswith("ઉદ્દેશ્ય"):
            extracted_data[f"{{{{OBJ_{current_key}}}}}"] = text.replace("ઉદ્દેશ્ય (Objective):", "").strip()
        elif text.startswith("શું કરવું?"):
            extracted_data[f"{{{{STEPS_{current_key}}}}}"] = text.replace("શું કરવું? (What to do?):", "").strip()
        elif text.startswith("જવાબદાર વ્યક્તિ"):
            extracted_data[f"{{{{WHO_{current_key}}}}}"] = text.replace("જવાબદાર વ્યક્તિ (Responsible Person):", "").strip()
        elif text.startswith("મોનિટરિંગ"):
            extracted_data[f"{{{{IND_{current_key}}}}}"] = text.replace("મોનિટરિંગ સૂચકાંકો (Monitoring Indicators):", "").strip()

    return extracted_data

# --- Helper Function: Inject Text into Template ---
def inject_text_to_ppt(template_file, replacements):
    """
    Scans every slide and every shape in the PPTX. 
    If it finds a placeholder (e.g., {{OBJ_1.1}}), it replaces it with the Word doc text
    while preserving your custom font, size, and colors.
    """
    prs = Presentation(template_file)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
                
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for key, value in replacements.items():
                        if key in run.text:
                            # Replace the placeholder with the actual Gujarati text
                            run.text = run.text.replace(key, value)

    # Save to memory buffer
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- Streamlit UI ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload PPTX Template")
    uploaded_template = st.file_uploader("Upload your blank designed template (.pptx)", type=["pptx"])

with col2:
    st.subheader("2. Upload Content")
    uploaded_docx = st.file_uploader("Upload Gujarati Word Document (.docx)", type=["docx"])

if uploaded_template and uploaded_docx:
    if st.button("Generate Final Presentation", type="primary"):
        with st.spinner("Extracting content from Word..."):
            # 1. Parse the Word document
            replacements = extract_data_from_word(uploaded_docx)
            
        with st.spinner("Injecting text into your graphic template..."):
            # 2. Inject into PPTX
            final_ppt = inject_text_to_ppt(uploaded_template, replacements)
            
        st.success("Perfect Graphic Presentation Generated Successfully!")
        
        st.download_button(
            label="📄 Download Exact Infographic PPTX",
            data=final_ppt,
            file_name="AMC_NTEP_Final_Infographic.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
