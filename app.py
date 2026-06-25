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
from pptx.oxml import OxmlElement
from pptx.oxml.ns import qn
import google.generativeai as genai

# ==========================================
# 1. PAGE SETUP & SECURITY
# ==========================================
st.set_page_config(page_title="Academic Premium Visual PPT", page_icon="🎓", layout="centered")

api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.warning("⚠️ API Key not detected. Please set GEMINI_API_KEY in Streamlit Secrets.")

# WOW Factor Premium Color Palette (Strict Academic Standard)
THEME_DARK_BLUE = RGBColor(15, 34, 64)       # Deep Navy Blue (Corporate/Academic)
THEME_BORDER_BLUE = RGBColor(60, 110, 170)    # Sleek Accent Blue
THEME_LIGHT_GRAY = RGBColor(245, 246, 248)    # Clean off-white background texture [1]
THEME_CARD_BG = RGBColor(255, 255, 255)       # Solid White Cards
THEME_TEXT_DARK = RGBColor(40, 40, 40)        # Readable dark gray body text
THEME_WHITE = RGBColor(255, 255, 255)         # Pure White
THEME_ACCENT_GREEN = RGBColor(40, 167, 69)    # Safe green stats

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
# 3. EXTRA-TEXT-CONDENSED ACADEMIC AI ENGINE
# ==========================================
def ai_process_word_advanced(full_text):
    prompt = f"""
    You are an expert Academic Presentation Assistant. Your goal is to structure this text into HIGHLY VISUAL, PREMIUM presentation slides with a 'WOW' factor.
    
    STRICT TEXT LIMITS:
    - Never write long paragraphs or sentences on slides.
    - Each slide bullet point must be strictly restricted to 3-5 words maximum.
    - Focus strictly on dynamic visual representation rather than plain text.
    - Put all extensive explanations in the 'script' (Speaker Notes) so the presenter can speak for 1 minute [1].

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
    - Put all extensive explanations in the 'script' (Speaker Notes) so the presenter can speak for 1 minute [1].

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
