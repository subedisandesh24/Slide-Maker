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
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
import google.generativeai as genai

# ==========================================
# 1. PAGE SETUP & SECURITY
# ==========================================
st.set_page_config(page_title="Academic Visual PPT Generator", page_icon="🎓", layout="centered")

api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.warning("⚠️ API Key not detected. Please set GEMINI_API_KEY in Streamlit Secrets.")

# ==========================================
# 2. DIVERSE DESIGN THEMES (replaces external template scraping)
# ==========================================
# Note: Sites like Slidesgo / Canva / PresentationGO / SlidesCarnival / Microsoft Create /
# SlideEgg do NOT offer a public API for downloading templates, and their templates are
# licensed/copyrighted assets — so an automated "fetch & shuffle from these sites" feature
# isn't something this script can do. Instead, this engine ships with several distinct,
# original color/style themes (inspired by the *visual niche* of each site) and randomly
# shuffles which theme is applied to each generated file, so every batch looks different.

THEMES = {
    "Modern_Slidesgo_Style": {
        "primary": RGBColor(20, 40, 80),
        "light": RGBColor(220, 230, 245),
        "border": RGBColor(70, 130, 180),
        "text": RGBColor(30, 30, 30),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(40, 167, 69),
        "bg": RGBColor(248, 249, 252),
    },
    "DataGrid_PresentationGO_Style": {
        "primary": RGBColor(15, 76, 92),
        "light": RGBColor(214, 240, 237),
        "border": RGBColor(43, 140, 130),
        "text": RGBColor(25, 35, 35),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(230, 126, 34),
        "bg": RGBColor(247, 252, 251),
    },
    "Vibrant_Canva_Style": {
        "primary": RGBColor(88, 24, 130),
        "light": RGBColor(238, 222, 248),
        "border": RGBColor(155, 89, 182),
        "text": RGBColor(35, 25, 40),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(241, 90, 134),
        "bg": RGBColor(250, 246, 252),
    },
    "Minimalist_SlidesCarnival_Style": {
        "primary": RGBColor(45, 45, 45),
        "light": RGBColor(235, 235, 235),
        "border": RGBColor(120, 120, 120),
        "text": RGBColor(30, 30, 30),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(0, 0, 0),
        "bg": RGBColor(250, 250, 250),
    },
    "Corporate_Microsoft_Style": {
        "primary": RGBColor(0, 90, 158),
        "light": RGBColor(213, 232, 246),
        "border": RGBColor(0, 120, 212),
        "text": RGBColor(32, 31, 30),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(16, 124, 16),
        "bg": RGBColor(247, 250, 253),
    },
    "Diagrammatic_SlideEgg_Style": {
        "primary": RGBColor(180, 60, 30),
        "light": RGBColor(252, 228, 214),
        "border": RGBColor(214, 100, 50),
        "text": RGBColor(40, 30, 25),
        "white": RGBColor(255, 255, 255),
        "accent": RGBColor(39, 78, 19),
        "bg": RGBColor(253, 248, 245),
    },
}

def get_shuffled_themes(n):
    """Randomly shuffles and picks n themes (with repeats allowed if n > number of themes)."""
    names = list(THEMES.keys())
    random.shuffle(names)
    if n <= len(names):
        return names[:n]
    # if more variants requested than themes available, cycle with reshuffles
    out = []
    while len(out) < n:
        random.shuffle(names)
        out.extend(names)
    return out[:n]

# ==========================================
# 3. FILE READERS
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
# 4. ADVANCED ACADEMIC AI PROCESSING
# ==========================================
def ai_process_word_advanced(full_text):
    prompt = f"""
    You are an expert Academic Presentation Assistant. Your goal is to structure this text into HIGHLY VISUAL, PREMIUM presentation slides with a 'WOW' factor.

    STRICT TEXT LIMITS:
    - Never write long paragraphs or sentences on slides.
    - Each slide bullet point must be strictly restricted to 3-5 words maximum.
    - Focus strictly on dynamic visual representation rather than plain text.
    - Put all extensive explanations in the 'script' (Speaker Notes) so the presenter can speak for 1 minute.

    Available Layout Types:
    1. 'process': Use for workflows, methodologies, or step-by-step procedures. (Horizontal process cards with badges)
    2. 'table': Use for comparative parameter matrix, structured comparison data.
    3. 'comparison': Use for contrasting two concepts, Pros vs Cons, Old vs New. (Side-by-side split cards)
    4. 'stat': Use if there's a strong numeric fact, percentage, metric, or major statistic (e.g., '95% Efficiency', '500+ Participants').
    5. 'bullets': Use only if none of the above visual layouts match.

    Return ONLY a valid JSON object in this format (strictly no markdown formatting outside the JSON, no extra text):
    {{
        "category": "lecture",
        "recommended_source": "slidesgo | presentationgo | slidescarnival | slideegg | microsoft",
        "slides": [
            {{
                "title": "Slide Title",
                "layout_type": "process | table | comparison | stat | bullets",
                "content": {{
                    "bullets": ["Keyword-focused bullet 1", "Keyword-focused bullet 2"],
                    "table": {{
                        "headers": ["Parameter", "Detail A", "Detail B"],
                        "rows": [["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"]]
                    }},
                    "process": [
                        {{"step": "1", "title": "Phase 1", "desc": "Phase 1 summary"}}
                    ],
                    "comparison": {{
                        "left_title": "Side A Title",
                        "left_content": ["Short point A1", "Short point A2"],
                        "right_title": "Side B Title",
                        "right_content": ["Short point B1", "Short point B2"]
                    }},
                    "stat": {{
                        "number": "95%",
                        "label": "Metric Description (max 5 words)"
                    }}
                }},
                "script": "A detailed, professional 1-minute presenter speech explaining the slide content..."
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
    You are an expert Academic Presentation Assistant. Process this rough slide deck. Maintain 1:1 slide order ({len(slide_texts)} slides).
    Strictly summarize the text into a premium visual design with a 'WOW' factor.

    STRICT TEXT LIMITS:
    - Never write long sentences.
    - Bullet points must be strictly 3-5 words maximum.
    - Put all extensive explanations in the 'script' (Speaker Notes) so the presenter can speak for 1 minute.

    For each slide, choose: 'process', 'table', 'comparison', 'stat', or 'bullets'.

    Return ONLY a valid JSON object in this format (strictly no markdown formatting outside the JSON, no extra text):
    {{
        "category": "lecture",
        "recommended_source": "slidesgo | presentationgo | slidescarnival | slideegg | microsoft",
        "slides": [
            {{
                "title": "Slide Title",
                "layout_type": "process | table | comparison | stat | bullets",
                "content": {{
                    "bullets": ["Keyword point 1", "Keyword point 2"],
                    "table": {{
                        "headers": ["Parameter", "Detail A", "Detail B"],
                        "rows": [["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"]]
                    }},
                    "process": [
                        {{"step": "1", "title": "Phase 1", "desc": "Phase 1 summary"}}
                    ],
                    "comparison": {{
                        "left_title": "Side A",
                        "left_content": ["Short point A1", "Short point A2"],
                        "right_title": "Side B",
                        "right_content": ["Short point B1", "Short point B2"]
                    }},
                    "stat": {{
                        "number": "85%",
                        "label": "Metric Description"
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
# 5. JSON ROBUST NORMALIZER
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
# 6. PROGRAMMATIC TRANSITION & TIMING TRIGGER (XML)
# ==========================================
def set_slide_transition_and_timing(slide, transition_type="fade", auto_advance_sec=60):
    sLD = slide._element
    transition = sLD.find(qn('p:transition'))
    if transition is None:
        transition = OxmlElement('p:transition')
        sLD.append(transition)

    for child in list(transition):
        transition.remove(child)

    if transition_type == "fade":
        effect = OxmlElement('p:fade')
        transition.append(effect)
    elif transition_type == "push":
        effect = OxmlElement('p:push')
        effect.set('dir', 'lt')
        transition.append(effect)
    elif transition_type == "wipe":
        effect = OxmlElement('p:wipe')
        effect.set('dir', 'lt')
        transition.append(effect)

    if auto_advance_sec > 0:
        transition.set('advClick', '0')
        transition.set('advTm', str(auto_advance_sec * 1000))

# ==========================================
# 7. PREMIUM WOW-FACTOR VISUAL DRAWING ENGINE (theme-aware)
# ==========================================
def apply_slide_branding_texture(slide, theme):
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = theme["primary"]
    top_bar.line.fill.background()

def draw_visuals_on_slide(slide, slide_data, theme):
    layout_type = slide_data.get("layout_type", "bullets")
    content = slide_data.get("content", {})

    # Theme background
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = theme["bg"]

    # Apply branding bar
    apply_slide_branding_texture(slide, theme)

    # Title Setup
    if slide.shapes.title:
        slide.shapes.title.text = slide_data.get('title', "Academic Slide")
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Arial'
        slide.shapes.title.text_frame.paragraphs[0].font.bold = True
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = theme["primary"]

    # Clean default placeholders
    for shape in list(slide.shapes):
        if shape.is_placeholder and shape.placeholder_format.idx == 1:
            if layout_type in ["process", "table", "comparison", "stat"]:
                sp = shape._element
                sp.getparent().remove(sp)

    # LAYOUT 1: PROCESS/FLOWCHART (Horizontal Process Cards)
    if layout_type == "process" and "process" in content:
        steps = content["process"][:4]
        num_steps = len(steps)
        if num_steps > 0:
            start_x = Inches(1.0)
            y = Inches(2.5)
            box_width = Inches((11.33 - (num_steps - 1) * 0.4) / num_steps)
            box_height = Inches(3.2)

            for idx, step_data in enumerate(steps):
                x = start_x + idx * (box_width + Inches(0.4))

                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, box_width, box_height)
                card.fill.solid()
                card.fill.fore_color.rgb = theme["white"]
                card.line.color.rgb = theme["border"]
                card.line.width = Pt(1.5)

                header_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, box_width, Inches(0.45))
                header_bar.fill.solid()
                header_bar.fill.fore_color.rgb = theme["primary"]
                header_bar.line.fill.background()

                badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + box_width/2 - Inches(0.3), y - Inches(0.3), Inches(0.6), Inches(0.6))
                badge.fill.solid()
                badge.fill.fore_color.rgb = theme["border"]
                badge.line.fill.background()
                badge.text_frame.text = str(idx + 1)
                p_b = badge.text_frame.paragraphs[0]
                p_b.font.bold = True
                p_b.font.size = Pt(14)
                p_b.font.color.rgb = theme["white"]
                p_b.alignment = PP_ALIGN.CENTER

                tf = card.text_frame
                tf.word_wrap = True
                p_space = tf.paragraphs[0]
                p_space.text = "\n"

                p_title = tf.add_paragraph()
                p_title.text = step_data.get('title', "").upper()
                p_title.font.bold = True
                p_title.font.size = Pt(14)
                p_title.font.color.rgb = theme["primary"]
                p_title.alignment = PP_ALIGN.CENTER

                p_desc = tf.add_paragraph()
                p_desc.text = f"\n{step_data.get('desc', '')}"
                p_desc.font.size = Pt(11)
                p_desc.font.color.rgb = theme["text"]
                p_desc.alignment = PP_ALIGN.CENTER

                if idx < num_steps - 1:
                    arrow_x = x + box_width + Inches(0.05)
                    arrow_y = y + Inches(1.4)
                    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, Inches(0.3), Inches(0.4))
                    arrow.fill.solid()
                    arrow.fill.fore_color.rgb = theme["border"]
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
                cell.text = str(header_text).upper()
                cell.fill.solid()
                cell.fill.fore_color.rgb = theme["primary"]
                p = cell.text_frame.paragraphs[0]
                p.font.bold = True
                p.font.color.rgb = theme["white"]
                p.font.size = Pt(13)
                p.alignment = PP_ALIGN.CENTER

            for row_idx, row_values in enumerate(rows):
                for col_idx, val in enumerate(row_values):
                    if col_idx < num_cols:
                        cell = table.cell(row_idx + 1, col_idx)
                        cell.text = str(val)
                        cell.fill.solid()
                        if row_idx % 2 == 0:
                            cell.fill.fore_color.rgb = theme["light"]
                        else:
                            cell.fill.fore_color.rgb = theme["white"]
                        p = cell.text_frame.paragraphs[0]
                        p.font.size = Pt(11)
                        p.font.color.rgb = theme["text"]
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

        left_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), top_y, col_width, col_height)
        left_shape.fill.solid()
        left_shape.fill.fore_color.rgb = theme["white"]
        left_shape.line.color.rgb = theme["border"]

        left_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), top_y, col_width, Inches(0.15))
        left_accent.fill.solid()
        left_accent.fill.fore_color.rgb = theme["primary"]
        left_accent.line.fill.background()

        tf_l = left_shape.text_frame
        tf_l.word_wrap = True
        p_l = tf_l.paragraphs[0]
        p_l.text = f"\n{left_title.upper()}"
        p_l.font.bold = True
        p_l.font.size = Pt(16)
        p_l.font.color.rgb = theme["primary"]
        p_l.alignment = PP_ALIGN.CENTER

        for bullet in left_bullets:
            p = tf_l.add_paragraph()
            p.text = f"• {bullet}"
            p.font.size = Pt(12)
            p.font.color.rgb = theme["text"]

        right_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), top_y, col_width, col_height)
        right_shape.fill.solid()
        right_shape.fill.fore_color.rgb = theme["white"]
        right_shape.line.color.rgb = theme["border"]

        right_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0), top_y, col_width, Inches(0.15))
        right_accent.fill.solid()
        right_accent.fill.fore_color.rgb = theme["border"]
        right_accent.line.fill.background()

        tf_r = right_shape.text_frame
        tf_r.word_wrap = True
        p_r = tf_r.paragraphs[0]
        p_r.text = f"\n{right_title.upper()}"
        p_r.font.bold = True
        p_r.font.size = Pt(16)
        p_r.font.color.rgb = theme["primary"]
        p_r.alignment = PP_ALIGN.CENTER

        for bullet in right_bullets:
            p = tf_r.add_paragraph()
            p.text = f"• {bullet}"
            p.font.size = Pt(12)
            p.font.color.rgb = theme["text"]

    # LAYOUT 4: BIG STAT / METRIC CALLOUT
    elif layout_type == "stat" and "stat" in content:
        stat_data = content["stat"]
        number = stat_data.get("number", "0%")
        label = stat_data.get("label", "Key Metric")

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.5), Inches(2.2), Inches(6.33), Inches(4.0))
        card.fill.solid()
        card.fill.fore_color.rgb = theme["white"]
        card.line.color.rgb = theme["border"]
        card.line.width = Pt(2.0)

        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.5), Inches(2.2), Inches(6.33), Inches(0.2))
        band.fill.solid()
        band.fill.fore_color.rgb = theme["primary"]
        band.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True

        p_num = tf.paragraphs[0]
        p_num.text = f"\n{number}"
        p_num.font.bold = True
        p_num.font.size = Pt(64)
        p_num.font.color.rgb = theme["accent"]
        p_num.alignment = PP_ALIGN.CENTER

        p_label = tf.add_paragraph()
        p_label.text = f"\n{label.upper()}"
        p_label.font.bold = True
        p_label.font.size = Pt(16)
        p_label.font.color.rgb = theme["primary"]
        p_label.alignment = PP_ALIGN.CENTER

    # FALLBACK / LAYOUT 5: STANDARD BULLETS WITH LEFT ACCENT
    else:
        bullets = content.get("bullets", ["No text details provided."])

        accent_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(2.2), Inches(0.15), Inches(4.0))
        accent_card.fill.solid()
        accent_card.fill.fore_color.rgb = theme["border"]
        accent_card.line.fill.background()

        txBox = slide.shapes.add_textbox(Inches(1.8), Inches(2.2), Inches(10.0), Inches(4.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        for idx, bullet in enumerate(bullets):
            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
            p.text = f"• {bullet}"
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(10)

# ==========================================
# 8. DIVERSE / SHUFFLED PRESENTATION GENERATION
# ==========================================
def build_base_presentation(custom_template_bytes=None):
    """Builds a blank 16:9 presentation, optionally starting from an uploaded custom template."""
    if custom_template_bytes:
        prs = Presentation(io.BytesIO(custom_template_bytes))
    else:
        prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def generate_presentations_advanced(ai_data, num_variants, custom_templates=None):
    """
    Generates `num_variants` distinct .pptx files. Each file gets a randomly
    shuffled design theme so that every batch of outputs looks different,
    even on repeated runs with the same source document.
    """
    output_files = {}
    slides_data = ai_data['slides']

    # Pick a shuffled order of themes for this batch
    theme_order = get_shuffled_themes(num_variants)

    for i, theme_name in enumerate(theme_order):
        theme = THEMES[theme_name]

        # If the user uploaded their own templates, cycle through them as the base;
        # otherwise start from a clean blank presentation.
        base_bytes = None
        if custom_templates:
            base_bytes = custom_templates[i % len(custom_templates)][1]

        prs = build_base_presentation(base_bytes)

        for slide_content in slides_data:
            slide_layout = prs.slide_layouts[1]
            new_slide = prs.slides.add_slide(slide_layout)

            draw_visuals_on_slide(new_slide, slide_content, theme)

            layout_type = slide_content.get("layout_type", "bullets")
            if layout_type == "process":
                set_slide_transition_and_timing(new_slide, transition_type="push", auto_advance_sec=60)
            elif layout_type == "table":
                set_slide_transition_and_timing(new_slide, transition_type="wipe", auto_advance_sec=60)
            else:
                set_slide_transition_and_timing(new_slide, transition_type="fade", auto_advance_sec=60)

            notes_slide = new_slide.notes_slide
            notes_slide.notes_text_frame.text = slide_content.get('script', "No speech generated.")

        output_stream = io.BytesIO()
        prs.save(output_stream)
        output_stream.seek(0)
        clean_theme_name = theme_name.replace("_Style", "")
        output_files[f"Design_{i+1}_{clean_theme_name}.pptx"] = output_stream.getvalue()

    return output_files

# ==========================================
# 9. STREAMLIT WEB FRONTEND (UI)
# ==========================================
st.title("🎓 Smart Academic Presentation Generator (Visuals Active)")
st.write("Convert your academic documents into highly structured, dynamic presentations. Focuses on **Tables, Process Flowcharts, and Comparisons** automatically!")
st.caption("Each generation produces a randomly shuffled set of distinct design themes — no two batches look identical.")

uploaded_file = st.file_uploader("Upload your rough Document (.docx or .pptx)", type=["docx", "pptx"])
template_uploads = st.file_uploader("Upload custom Master PPTX Templates (Optional)", type=["pptx"], accept_multiple_files=True)
num_variants = st.slider("How many design variants to generate?", min_value=1, max_value=6, value=5)

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

                # 2. Optional custom templates (used as a base, theme colors still applied on top)
                custom_templates = None
                if template_uploads:
                    custom_templates = [(t.name, t.read()) for t in template_uploads]

                # 3. Generate diverse, shuffled visual slides
                generated_ppts = generate_presentations_advanced(ai_data, num_variants, custom_templates)

                # 4. Create ZIP for easy download
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_data in generated_ppts.items():
                        zip_file.writestr(filename, file_data)

                zip_buffer.seek(0)

                st.success(f"🎉 {num_variants} diverse, shuffled presentations generated! AI suggests this content best matches a **{recommended_source.upper()}**-style layout.")

                st.download_button(
                    label="📥 Download Generated Presentations (ZIP)",
                    data=zip_buffer,
                    file_name="academic_visual_presentations.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"Error occurred: {str(e)}")

st.write("---")
st.subheader("💡 Recommended Free Template Resources")
st.write("These sites don't offer an API for auto-downloading templates, but you can grab a `.pptx` from any of them yourself and upload it above under 'Upload custom Master Templates' — the AI's visual layout engine will then build on top of it:")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    *   [**Slidesgo**](https://slidesgo.com/)
        *   *Niche:* Modern, highly visual, creative academic decks. Perfect for classroom lectures or assignments.
    *   [**PresentationGO**](https://www.presentationgo.com/)
        *   *Niche:* Clean grids, infographics, data tables, and professional sequential timelines.
    *   [**SlidesCarnival**](https://www.slidescarnival.com/)
        *   *Niche:* Clean, traditional, minimalist academic & formal thesis layout designs.
    """)

with col2:
    st.markdown("""
    *   [**SlideEgg**](https://www.slideegg.com/)
        *   *Niche:* Excellent structured diagrammatic templates, perfect for technical research.
    *   [**Microsoft Create**](https://create.microsoft.com/)
        *   *Niche:* Official, highly stable corporate/academic standard Microsoft PowerPoint templates.
    *   [**Canva Presentations**](https://www.canva.com/)
        *   *Niche:* Modern, visually rich slides (Can export easily to `.pptx` for our engine).
    """)
