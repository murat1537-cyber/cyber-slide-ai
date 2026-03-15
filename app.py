import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import json
import io

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

def create_pptx(json_data):
    data = json.loads(json_data)
    prs = Presentation()
    
    bg_rgb = hex_to_rgb(data["presentation_metadata"]["global_background_color_hex"])
    accent_rgb = hex_to_rgb(data["presentation_metadata"]["global_accent_color_hex"])
    
    for slide_data in data["slides"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(bg_rgb[0], bg_rgb[1], bg_rgb[2])
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(1))
        title_p = title_box.text_frame.paragraphs[0]
        title_p.text = slide_data["slide_title"]
        title_p.font.bold = True
        title_p.font.size = Pt(32)
        title_p.font.color.rgb = RGBColor(accent_rgb[0], accent_rgb[1], accent_rgb[2])
        
        layout_type = slide_data.get("layout_type", "text_only")
        text_color = 255 if sum(bg_rgb) < 380 else 0
        
        if layout_type == "text_only":
            body_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
            body_frame = body_box.text_frame
            body_frame.word_wrap = True
            for bullet in slide_data["content_bullets"]:
                p = body_frame.add_paragraph()
                p.text = f"• {bullet}"
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(text_color, text_color, text_color)
                
        elif layout_type == "text_and_chart":
            text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.5), Inches(5))
            text_frame = text_box.text_frame
            text_frame.word_wrap = True
            for bullet in slide_data["content_bullets"]:
                p = text_frame.add_paragraph()
                p.text = f"• {bullet}"
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(text_color, text_color, text_color)
                
            visual_data = slide_data.get("visual_element", {})
            if visual_data.get("type") == "bar_chart":
                chart_data = CategoryChartData()
                chart_data.categories = visual_data.get("categories", ["A", "B", "C"])
                chart_data.add_series(visual_data.get("series_name", "Data"), tuple(visual_data.get("values", [1, 2, 3])))
                
                x, y, cx, cy = Inches(5.2), Inches(1.8), Inches(4.3), Inches(4)
                chart = slide.shapes.add_chart(
                    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
                ).chart
                chart.has_title = True
                chart.chart_title.text_frame.text = visual_data.get("title", "Chart")

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

st.set_page_config(page_title="Cyber-Slide AI", page_icon="🛡️", layout="wide")
st.title("🛡️ Cyber-Slide AI: Akıllı Sunum Mimarı")

API_KEY = st.secrets.get("GEMINI_API_KEY", "")
if not API_KEY:
    st.stop()

genai.configure(api_key=API_KEY)

with st.sidebar:
    st.header("⚙️ Eğitim Parametreleri")
    topic = st.text_input("Konu Nedir?", "Zero Trust Architecture in Cloud")
    language = st.selectbox("Sunum Dili", ["English", "Nederlands (Dutch)"])
    slide_count = st.slider("Slayt Sayısı", min_value=3, max_value=20, value=5)
    design_prompt = st.text_area("Tasarım", "Koyu gri arka plan, başlıklar neon yeşil olsun.")

if st.button("🚀 Sunumu Üret", type="primary"):
    with st.spinner("Slaytlar çiziliyor ve grafikler oluşturuluyor..."):
        model = genai.GenerativeModel('gemini-2.5-flash', generation_config={"response_mime_type": "application/json"})
        
        system_prompt = f"""
        You are an elite Cybersecurity Instructor. Generate a presentation.
        Topic: {topic}
        Language: {language}
        Slide Count: Exactly {slide_count} slides.
        Design Vibe: {design_prompt}
        
        Rules for Layouts:
        - "text_only": Standard slide.
        - "text_and_chart": You MUST invent realistic cybersecurity statistical data and output it for a Bar Chart.
        
        Make sure at least 1 or 2 slides use "text_and_chart" to show data.
        
        Output STRICTLY in this JSON structure:
        {{
          "presentation_metadata": {{
            "global_background_color_hex": "#222222",
            "global_accent_color_hex": "#00FF00"
          }},
          "slides": [
            {{
              "slide_number": 1,
              "layout_type": "text_and_chart",
              "slide_title": "Slide Title Here",
              "content_bullets": ["Short text 1", "Short text 2"],
              "visual_element": {{
                "type": "bar_chart",
                "title": "Ransomware Attacks by Year",
                "categories": ["2021", "2022", "2023", "2024"],
                "series_name": "Incidents",
                "values": [1200, 1500, 2100, 3100]
              }}
            }}
          ]
        }}
        """
        
        try:
            response = model.generate_content(system_prompt)
            ppt_file = create_pptx(response.text)
            st.success("🎉 Sunumunuz hazır!")
            st.download_button("📥 PowerPoint Dosyasını İndir", data=ppt_file, file_name="Siber_Sunum.pptx")
        except Exception as e:
            st.error(f"Hata: {e}")
