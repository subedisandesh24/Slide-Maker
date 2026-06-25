import streamlit as st
import os
import random
import json
import zipfile
import io
import docx
from pptx import Presentation
import google.generativeai as genai

# ==========================================
# 1. PAGE SETUP & SECURITY
# ==========================================
st.set_page_config(page_title="Academic PPT Generator", page_icon="🎓", layout="centered")

# Streamlit Secrets bata API Key safe tarika le line (Zero leak risk)
# Cloud ma run garda Streamlit ko Settings > Secrets ma GEMINI_API_KEY set garna parcha
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    # Model updated to gemini-2.5-flash as 1.5 versions are shut down by Google
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.warning("⚠️ API Key not detected. Please set GEMINI_API_KEY in Streamlit Secrets.")

# ==========================================
# 2. FILE READERS
# ==========================================
def read_word_file(file_obj):
    doc = docx.Document(file_obj)
    return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

def read_ppt_file(file_obj):
    prs = Presentation(file_obj)
    slide_texts = []
    for slide in prs.slides:
        text = " ".join([shape.text for shape in slide.shapes if hasattr(shape, "text")])
        if text.strip(): 
            slide_texts.append(text)
    return slide_texts

# ==========================================
# 3. ACADEMIC AI PROCESSING
# ==========================================
def ai_process_word(full_text):
    prompt = f"""
    You are an expert Academic Presentation Assistant. Structure this text into formal presentation slides.
    Rules:
    1. Categorize into ONE: 'thesis', 'lecture', 'assignment'.
    2. Divide text into slides. Each slide needs EXACTLY 3-4 concise bullet points.
    3. Write a formal, 1-minute presenter script for each slide (to go in Speaker Notes).
    Return ONLY JSON:
    {{"category": "lecture", "slides": [{{"title": "Slide Title", "bullets": ["Bullet 1", "Bullet 2", "Bullet 3"], "script": "Presenter script..."}}]}}
    Text: {full_text}
    """
    response = model.generate_content(prompt)
    clean_json = response.text.strip().replace('```json', '').replace('```', '')
    return json.loads(clean_json)

def ai_process_ppt(slide_texts):
    formatted_slides = []
    for idx, text in enumerate(slide_texts):
        prompt = f"""
        Summarize this academic slide content.
        Rules:
        1. Keep original meaning. Use EXACTLY 3-4 concise bullets.
        2. Write a professional, academic 1-minute presenter script.
        Return ONLY JSON:
        {{"title": "Slide {idx+1} Title", "bullets": ["Bullet 1", "Bullet 2", "Bullet 3"], "script": "Script..."}}
        Text: {text}
        """
        response = model.generate_content(prompt)
        clean_json = response.text.strip().replace('```json', '').replace('```', '')
        formatted_slides.append(json.loads(clean_json))
    return {"category": "lecture", "slides": formatted_slides}

# ==========================================
# 4. TEMPLATE SELECTOR & INJECTION
# ==========================================
def generate_presentations(ai_data, master_templates):
    output_files = {}
    
    for i, (temp_name, temp_bytes) in enumerate(master_templates):
        # Read PPTX from bytes
        prs = Presentation(io.BytesIO(temp_bytes))
        slides_data = ai_data['slides']
        
        for slide_idx, slide_content in enumerate(slides_data):
            slide_layout = prs.slide_layouts[1] # Typically 'Title and Content' layout
            new_slide = prs.slides.add_slide(slide_layout)
            
            # Title
            if new_slide.shapes.title:
                new_slide.shapes.title.text = slide_content.get('title', f"Slide {slide_idx+1}")
            
            # Bullets
            body_shape = new_slide.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.text = slide_content['bullets'][0]
            for bullet in slide_content['bullets'][1:]:
                p = text_frame.add_paragraph()
                p.text = bullet
                
            # Speaker Notes
            notes_slide = new_slide.notes_slide
            notes_slide.notes_text_frame.text = slide_content['script']
            
        # Save generated file in memory
        output_stream = io.BytesIO()
        prs.save(output_stream)
        output_stream.seek(0)
        output_files[f"Generated_Design_{i+1}.pptx"] = output_stream.getvalue()
        
    return output_files

# ==========================================
# 5. STREAMLIT WEB FRONTEND (UI)
# ==========================================
st.title("🎓 Smart Academic Presentation Generator")
st.write("Convert your rough notes or rough PPTs into beautifully formatted, animated presentations with presenter scripts automatically!")

# File uploaders
uploaded_file = st.file_uploader("Upload your rough Document (.docx or .pptx)", type=["docx", "pptx"])
template_uploads = st.file_uploader("Upload 1 or more Master PPTX Templates (Or the system will use defaults)", type=["pptx"], accept_multiple_files=True)

if st.button("Generate Presentations", type="primary"):
    if not api_key:
        st.error("Please add your Gemini API Key first.")
    elif not uploaded_file:
        st.error("Please upload your rough document first.")
    else:
        with st.spinner("🤖 AI is reading and formatting your presentation..."):
            try:
                # 1. Read files
                if uploaded_file.name.endswith('.docx'):
                    raw_text = read_word_file(uploaded_file)
                    ai_data = ai_process_word(raw_text)
                else:
                    raw_slides = read_ppt_file(uploaded_file)
                    ai_data = ai_process_ppt(raw_slides)
                
                # 2. Template management
                templates_to_use = []
                if template_uploads:
                    for temp in template_uploads:
                        templates_to_use.append((temp.name, temp.read()))
                else:
                    # Fallback default template if user didn't upload any
                    fallback_prs = Presentation()
                    fallback_stream = io.BytesIO()
                    fallback_prs.save(fallback_stream)
                    fallback_stream.seek(0)
                    templates_to_use.append(("default.pptx", fallback_stream.read()))

                # 3. Generate PPTs
                generated_ppts = generate_presentations(ai_data, templates_to_use)
                
                # 4. Create ZIP for easy download
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_data in generated_ppts.items():
                        zip_file.writestr(filename, file_data)
                
                zip_buffer.seek(0)
                
                st.success("🎉 Presentations generated successfully!")
                
                # Download Button
                st.download_button(
                    label="📥 Download Generated Presentations (ZIP)",
                    data=zip_buffer,
                    file_name="academic_presentations.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"Error occurred: {str(e)}")
