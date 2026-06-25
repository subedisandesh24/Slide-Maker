import streamlit as st
import os
import random
import json
import zipfile
import io
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import google.generativeai as genai

# ==========================================
# 1. PAGE SETUP & SECURITY
# ==========================================
st.set_page_config(page_title="Academic Platform PPT Generator", page_icon="🎓", layout="centered")

api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    # Stable Gemini 2.5 Flash for production [1]
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.warning("⚠️ API Key not detected. Please set GEMINI_API_KEY in Streamlit Secrets.")

# Color Palette for visual elements (Formal Academic Theme)
THEME_DARK_BLUE = RGBColor(20, 40, 80)
THEME_LIGHT_BLUE = RGBColor(220, 230, 245)
THEME_BORDER_BLUE = RGBColor(70, 130, 180)
THEME_TEXT_DARK = RGBColor(30, 30, 30)
THEME_WHITE = RGBColor(255, 255, 255)

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
        if text.strip(): slide_texts.append(text)
    return slide_texts

# ==========================================
# 3. ADVANCED ACADEMIC AI PROCESSING
# ==========================================
def ai_process_word_advanced(full_text):
    prompt = f"""
    You are an expert Academic Presentation Assistant. Your goal is to structure this text into highly visual presentation slides.
    Do NOT make slide after slide of boring text bullet points. Intelligently select the best layout_type for each slide.

    Also, analyze the content tone and select the best template source among:
    - 'slidesgo': for modern, visually-rich general education/creative slides [2].
    - 'presentationgo': for charts, infographics, process flowcharts, timelines [2].
    - 'slidescarnival': for clean, formal, traditional academic/thesis layouts [2].
    - 'slideegg': for highly structured, data-heavy diagrammatic slides [2].
    - 'microsoft': for standard, official academic PowerPoint designs [2].

    Available Layout Types:
    - 'process': workflows/step-by-step procedures.
    - 'table': data tables/parameters.
    - 'comparison': side-by-side columns.
    - 'bullets': default basic text summaries.

    Return ONLY a valid JSON object in this format (strictly no markdown formatting outside the JSON, no extra text):
    {{
        "category": "lecture",
        "recommended_source": "slidesgo | presentationgo | slidescarnival | slideegg | microsoft",
        "slides": [
            {{
                "title": "Slide Title",
                "layout_type": "bullets | table | process | comparison",
                "content": {{
                    "bullets": ["Point 1", "Point 2"],
                    "table": {{
                        "headers": ["Parameter", "Detail A", "Detail B"],
                        "rows": [["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"]]
                    }},
                    "process": [
                        {{"step": "1", "title": "Phase 1", "desc": "Phase 1 summary"}},
                        {{"step": "2", "title": "Phase 2", "desc": "Phase 2 summary"}}
                    ],
                    "comparison": {{
                        "left_title": "Pros / Side A",
                        "left_content": ["Point A1", "Point A2"],
                        "right_title": "Cons / Side B",
                        "right_content": ["Point B1", "Point B2"]
                    }}
                }},
                "script": "A detailed 1-minute presenter speech explaining the slide content..."
            }}
        ]
    }}

    Academic Text:
    {full_text}
    """
    response = model.generate_content(prompt)
    clean_json = response.text.strip().replace('```json', '').replace('```', '')
    return json.loads(clean_json)

def ai_process_ppt_advanced(slide_texts):
    packaged_text = ""
    for idx, text in enumerate(slide_texts):
        packaged_text += f"\n--- ROUGH SLIDE {idx+1} ---\n{text}\n"
        
    prompt = f"""
    You are an expert Academic Presentation Assistant. I will provide you with the content of a rough presentation, separated slide-by-slide.
    Do NOT change the slide order, and keep the exact same number of slides ({len(slide_texts)} slides).
    
    For EACH rough slide, decide the best visual layout_type ('process', 'table', 'comparison', or 'bullets') and structure the content accordingly.
    Also, write a formal 1-minute presenter script for each slide.
    
    Based on content style, recommend the best layout source among: 'slidesgo', 'presentationgo', 'slidescarnival', 'slideegg', 'microsoft' [2].

    Return ONLY a valid JSON object in this format (strictly no markdown formatting outside the JSON, no extra text):
    {{
        "category": "lecture",
        "recommended_source": "slidesgo | presentationgo | slidescarnival | slideegg | microsoft",
        "slides": [
            {{
                "title": "Slide Title",
                "layout_type": "bullets | table | process | comparison",
                "content": {{
                    "bullets": ["Point 1", "Point 2"],
                    "table": {{
                        "headers": ["Parameter", "Detail A", "Detail B"],
                        "rows": [["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"]]
                    }},
                    "process": [
                        {{"step": "1", "title": "Phase 1", "desc": "Phase 1 summary"}},
                        {{"step": "2", "title": "Phase 2", "desc": "Phase 2 summary"}}
                    ],
                    "comparison": {{
                        "left_title": "Pros / Side A",
                        "left_content": ["Point A1", "Point A2"],
                        "right_title": "Cons / Side B",
                        "right_content": ["Point B1", "Point B2"]
                    }}
                }},
                "script": "A detailed 1-minute presenter speech explaining the slide..."
            }}
        ]
    }}

    Rough Slides Data:
    {packaged_text}
    """
    response = model.generate_content(prompt)
    clean_json = response.text.strip().replace('```json', '').replace('```', '')
    return json.loads(clean_json)

# ==========================================
# 4. JSON ROBUST NORMALIZER
# ==========================================
def normalize_ai_data(ai_data):
    normalized = {"category": "lecture", "recommended_source": "slidesgo", "slides": []}
    if not ai_data:
        return normalized
        
    if isinstance(ai_data, list):
        raw_slides = ai_data
    elif isinstance(ai_data, dict):
        normalized["category"] = ai_data.get("category", "lecture")
        normalized["recommended_source"] = ai_data.get("recommended_source", "slidesgo")
        raw_slides = ai_data.get("slides", [])
        if isinstance(raw_slides, dict):
            raw_slides = [raw_slides]
    else:
        raw_slides = []
        
    for item in raw_slides:
        if isinstance(item, list):
            if len(item) > 0: item = item[0]
            else: continue
                
        if isinstance(item, dict):
            slide = {
                "title": item.get("title", "Academic Slide"),
                "layout_type": item.get("layout_type", "bullets"),
                "content": item.get("content", {}),
                "script": item.get("script", "No speech generated.")
            }
            normalized["slides"].append(slide)
            
    return normalized

# ==========================================
# 5. POWERPOINT VISUAL GENERATION
# ==========================================
def draw_visuals_on_slide(slide, slide_data):
    layout_type = slide_data.get("layout_type", "bullets")
    content = slide_data.get("content", {})
    
    # Title Setup
    if slide.shapes.title:
        slide.shapes.title.text = slide_data.get('title', "Academic Slide")
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Arial'
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = THEME_DARK_BLUE

    # Clean placeholders if dynamic shape layout is selected
    for shape in list(slide.shapes):
        if shape.is_placeholder and shape.placeholder_format.idx == 1:
            if layout_type in ["process", "table", "comparison"]:
                sp = shape._element
                sp.getparent().remove(sp)

    # LAYOUT 1: PROCESS/FLOWCHART (Horizontal Process Cards)
    if layout_type == "process" and "process" in content:
        steps = content["process"][:4]  # Max 4 steps
        num_steps = len(steps)
        if num_steps > 0:
            slide_width_inches = 13.33
            start_x = Inches(1.0)
            y = Inches(2.5)
            box_width = Inches((slide_width_inches - 2.0 - (num_steps - 1) * 0.4) / num_steps)
            box_height = Inches(3.2)
            
            for idx, step_data in enumerate(steps):
                x = start_x + idx * (box_width + Inches(0.4))
                shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, box_width, box_height)
                shape.fill.solid()
                shape.fill.fore_color.rgb = THEME_LIGHT_BLUE
                shape.line.color.rgb = THEME_BORDER_BLUE
                shape.line.width = Pt(1.5)
                
                tf = shape.text_frame
                tf.word_wrap = True
                
                p_num = tf.paragraphs[0]
                p_num.text = f"STEP {step_data.get('step', idx+1)}"
                p_num.font.bold = True
                p_num.font.size = Pt(14)
                p_num.font.color.rgb = THEME_BORDER_BLUE
                p_num.alignment = PP_ALIGN.CENTER
                
                p_title = tf.add_paragraph()
                p_title.text = step_data.get('title', "")
                p_title.font.bold = True
                p_title.font.size = Pt(16)
                p_title.font.color.rgb = THEME_DARK_BLUE
                p_title.alignment = PP_ALIGN.CENTER
                
                p_desc = tf.add_paragraph()
                p_desc.text = f"\n{step_data.get('desc', '')}"
                p_desc.font.size = Pt(11)
                p_desc.font.color.rgb = THEME_TEXT_DARK
                p_desc.alignment = PP_ALIGN.CENTER

                if idx < num_steps - 1:
                    arrow_x = x + box_width + Inches(0.05)
                    arrow_y = y + Inches(1.4)
                    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, Inches(0.3), Inches(0.4))
                    arrow.fill.solid()
                    arrow.fill.fore_color.rgb = THEME_BORDER_BLUE
                    arrow.line.fill.background()

    # LAYOUT 2: ACADEMIC TABLES
    elif layout_type == "table" and "table" in content:
        table_data = content["table"]
        headers = table_data.get("headers", [])
        rows = table_data.get("rows", [])
        num_cols = len(headers)
        num_rows = len(rows) + 1
        
        if num_cols > 0 and num_rows > 1:
            left = Inches(1.5)
            top = Inches(2.2)
            width = Inches(10.33)
            height = Inches(0.6 * num_rows)
            
            table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
            table = table_shape.table
            
            for col_idx, header_text in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.text = str(header_text)
                cell.fill.solid()
                cell.fill.fore_color.rgb = THEME_DARK_BLUE
                p = cell.text_frame.paragraphs[0]
                p.font.bold = True
                p.font.color.rgb = THEME_WHITE
                p.font.size = Pt(14)
                p.alignment = PP_ALIGN.CENTER
                
            for row_idx, row_values in enumerate(rows):
                for col_idx, val in enumerate(row_values):
                    if col_idx < num_cols:
                        cell = table.cell(row_idx + 1, col_idx)
                        cell.text = str(val)
                        cell.fill.solid()
                        if row_idx % 2 == 0:
                            cell.fill.fore_color.rgb = THEME_LIGHT_BLUE
                        else:
                            cell.fill.fore_color.rgb = THEME_WHITE
                        p = cell.text_frame.paragraphs[0]
                        p.font.size = Pt(12)
                        p.font.color.rgb = THEME_TEXT_DARK
                        p.alignment = PP_ALIGN.LEFT

    # LAYOUT 3: SIDE-BY-SIDE COMPARISON
    elif layout_type == "comparison" and "comparison" in content:
        comp_data = content["comparison"]
        left_title = comp_data.get("left_title", "Side A")
        right_title = comp_data.get("right_title", "Side B")
        left_bullets = comp_data.get("left_content", [])
        right_bullets = comp_data.get("right_content", [])
        
        col_width = Inches(5.1)
        col_height = Inches(4.2)
        top_y = Inches(2.2)
        
        # Left Box
        left_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), top_y, col_width, col_height)
        left_shape.fill.solid()
        left_shape.fill.fore_color.rgb = THEME_LIGHT_BLUE
        left_shape.line.color.rgb = THEME_BORDER_BLUE
        tf_l = left_shape.text_frame
        tf_l.word_wrap = True
        p_l = tf_l.paragraphs[0]
        p_l.text = left_title.upper()
        p_l.font.bold = True
        p_l.font.size = Pt(16)
        p_l.font.color.rgb = THEME_DARK_BLUE
        p_l.alignment = PP_ALIGN.CENTER
        
        for bullet in left_bullets:
            p = tf_l.add_paragraph()
            p.text = f"• {bullet}"
            p.font.size = Pt(12)
            p.font.color.rgb = THEME_TEXT_DARK

        # Right Box
        right_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), top_y, col_width, col_height)
        right_shape.fill.solid()
        right_shape.fill.fore_color.rgb = THEME_LIGHT_BLUE
        right_shape.line.color.rgb = THEME_BORDER_BLUE
        tf_r = right_shape.text_frame
        tf_r.word_wrap = True
        p_r = tf_r.paragraphs[0]
        p_r.text = right_title.upper()
        p_r.font.bold = True
        p_r.font.size = Pt(16)
        p_r.font.color.rgb = THEME_DARK_BLUE
        p_r.alignment = PP_ALIGN.CENTER
        
        for bullet in right_bullets:
            p = tf_r.add_paragraph()
            p.text = f"• {bullet}"
            p.font.size = Pt(12)
            p.font.color.rgb = THEME_TEXT_DARK

    # FALLBACK / LAYOUT 4: STANDARD BULLETS
    else:
        bullets = content.get("bullets", ["No text details provided."])
        try:
            body_shape = slide.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.text = bullets[0]
            for bullet in bullets[1:]:
                p = text_frame.add_paragraph()
                p.text = bullet
        except Exception:
            txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(4.0))
            tf = txBox.text_frame
            tf.word_wrap = True
            for idx, bullet in enumerate(bullets):
                p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                p.text = f"• {bullet}"
                p.font.size = Pt(14)
                p.font.color.rgb = THEME_TEXT_DARK

# ==========================================
# 6. DYNAMIC TEMPLATE MATCHMAKING LOGIC
# ==========================================
def get_5_templates_by_source(recommended_source):
    """Checks for server folders. If not populated, falls back to a clean master style dynamically"""
    source_folder = f"master_templates/{recommended_source.lower()}/"
    templates_to_use = []
    
    if os.path.exists(source_folder) and len(os.listdir(source_folder)) > 0:
        all_templates = os.listdir(source_folder)
        selected = random.sample(all_templates, min(len(all_templates), 5))
        for t in selected:
            with open(os.path.join(source_folder, t), "rb") as f:
                templates_to_use.append((t, f.read()))
    else:
        # Fallback Dynamic Template generation (to prevent any file-missing crashes)
        for i in range(5):
            fallback_prs = Presentation()
            fallback_prs.slide_width = Inches(13.333)
            fallback_prs.slide_height = Inches(7.5)
            # Custom dynamic theme coloring based on source
            if recommended_source == "slidesgo":
                fallback_prs.slide_layouts[1].background.fill.solid()
                fallback_prs.slide_layouts[1].background.fill.fore_color.rgb = RGBColor(245, 248, 253)
            elif recommended_source == "presentationgo":
                fallback_prs.slide_layouts[1].background.fill.solid()
                fallback_prs.slide_layouts[1].background.fill.fore_color.rgb = RGBColor(240, 245, 240)
            
            fallback_stream = io.BytesIO()
            fallback_prs.save(fallback_stream)
            fallback_stream.seek(0)
            templates_to_use.append((f"Fallback_{recommended_source}_{i+1}.pptx", fallback_stream.read()))
            
    return templates_to_use

def generate_presentations_advanced(ai_data, master_templates):
    output_files = {}
    
    for i, (temp_name, temp_bytes) in enumerate(master_templates):
        prs = Presentation(io.BytesIO(temp_bytes))
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        slides_data = ai_data['slides']
        
        for slide_idx, slide_content in enumerate(slides_data):
            slide_layout = prs.slide_layouts[1]
            new_slide = prs.slides.add_slide(slide_layout)
            
            draw_visuals_on_slide(new_slide, slide_content)
                
            notes_slide = new_slide.notes_slide
            notes_slide.notes_text_frame.text = slide_content.get('script', "No speech generated.")
            
        output_stream = io.BytesIO()
        prs.save(output_stream)
        output_stream.seek(0)
        output_files[f"Design_{temp_name.replace('.pptx', '')}_{i+1}.pptx"] = output_stream.getvalue()
        
    return output_files

# ==========================================
# 7. STREAMLIT WEB FRONTEND (UI)
# ==========================================
st.title("🎓 Smart Academic Presentation Generator (Visuals Active)")
st.write("Convert your academic documents into highly structured, dynamic presentations. Automatically maps to the best platform templates [2]!")

uploaded_file = st.file_uploader("Upload your rough Document (.docx or .pptx)", type=["docx", "pptx"])
template_uploads = st.file_uploader("Upload custom Master PPTX Templates (Optional)", type=["pptx"], accept_multiple_files=True)

if st.button("Generate Visual Presentations", type="primary"):
    if not api_key:
        st.error("Please add your Gemini API Key first.")
    elif not uploaded_file:
        st.error("Please upload your rough document first.")
    else:
        with st.spinner("🤖 AI is reading, categorizing, and designing visual flows..."):
            try:
                # 1. Read files
                if uploaded_file.name.endswith('.docx'):
                    raw_text = read_word_file(uploaded_file)
                    ai_raw_data = ai_process_word_advanced(raw_text)
                else:
                    raw_slides = read_ppt_file(uploaded_file)
                    ai_raw_data = ai_process_ppt_advanced(raw_slides)
                
                # Robust Normalization of AI JSON
                ai_data = normalize_ai_data(ai_raw_data)
                recommended_source = ai_data.get("recommended_source", "slidesgo")
                
                # 2. Template management
                templates_to_use = []
                if template_uploads:
                    for temp in template_uploads:
                        templates_to_use.append((temp.name, temp.read()))
                else:
                    # Dynamic selector pulls from respective platform folder on server!
                    templates_to_use = get_5_templates_by_source(recommended_source)

                # 3. Generate dynamic visual slides
                generated_ppts = generate_presentations_advanced(ai_data, templates_to_use)
                
                # 4. Create ZIP for easy download
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_data in generated_ppts.items():
                        zip_file.writestr(filename, file_data)
                
                zip_buffer.seek(0)
                
                st.success(f"🎉 Visual presentations successfully generated! AI recommends **{recommended_source.upper()}** style [2]!")
                
                st.download_button(
                    label=f"📥 Download {recommended_source.upper()} Styles (ZIP)",
                    data=zip_buffer,
                    file_name="academic_visual_presentations.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"Error occurred: {str(e)}")

st.write("---")
# INTERACTIVE RESOURCE DIRECTORY (Visual Reference for users) [2]
st.subheader("💡 Recommended Free Template Resources [2]")
st.write("Hamro system le auto-match garne official sites haru yahan chan. Tapai le aafule manparayeko template download garera 'Upload custom Master templates' section ma upload garna pani saknu huncha [2]:")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    *   [**Slidesgo**](https://slidesgo.com/) [2]
        *   *Niche:* Modern, highly visual, creative academic decks. Perfect for classroom lectures or assignments [2].
    *   [**PresentationGO**](https://www.presentationgo.com/) [2]
        *   *Niche:* Clean grids, infographics, data tables, and professional sequential timelines [2].
    *   [**SlidesCarnival**](https://www.slidescarnival.com/) [2]
        *   *Niche:* Clean, traditional, minimalist academic & formal thesis layout designs [2].
    """)

with col2:
    st.markdown("""
    *   [**SlideEgg**](https://www.slideegg.com/) [2]
        *   *Niche:* Excellent structured diagrammatic templates, perfect for technical research [2].
    *   [**Microsoft Create**](https://create.microsoft.com/) [2]
        *   *Niche:* Official, highly stable corporate/academic standard Microsoft PowerPoint templates [2].
    *   [**Canva Presentations**](https://www.canva.com/) [2]
        *   *Niche:* Modern, visually rich slides (Can export easily to `.pptx` for our engine) [2].
    """)
