import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
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
    text_color = 255 if sum(bg_rgb) < 380 else 0
    
    for slide_data in data["slides"]:
        # Boş şablon
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Arka Plan
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(bg_rgb[0], bg_rgb[1], bg_rgb[2])
        
        # Başlık
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(1))
        title_p = title_box.text_frame.paragraphs[0]
        title_p.text = slide_data["slide_title"]
        title_p.font.bold = True
        title_p.font.size = Pt(32)
        title_p.font.color.rgb = RGBColor(accent_rgb[0], accent_rgb[1], accent_rgb[2])
        
        layout_type = slide_data.get("layout_type", "text_only")
        
        if layout_type == "text_only":
            body_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
            body_frame = body_box.text_frame
            body_frame.word_wrap = True
            for bullet in slide_data["content_bullets"]:
                p = body_frame.add_paragraph()
                p.text = f"• {bullet}"
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(text_color, text_color, text_color)
                
        elif layout_type == "text_with_image_placeholder":
            # Metin Kutusu (Sol Taraf)
            text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.5), Inches(5))
            text_frame = text_box.text_frame
            text_frame.word_wrap = True
            for bullet in slide_data["content_bullets"]:
                p = text_frame.add_paragraph()
                p.text = f"• {bullet}"
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(text_color, text_color, text_color)
                
            # GÖRSEL ALANI (Sağ Taraf - Yuvarlak Köşeli Şık Kutu)
            placeholder = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, 
                Inches(5.2), Inches(1.8), Inches(4.3), Inches(4.0)
            )
            
            # Kutunun rengini arka plana göre hafif zıt yapıyoruz
            box_color = 40 if sum(bg_rgb) < 380 else 220
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = RGBColor(box_color, box_color, box_color)
            placeholder.line.color.rgb = RGBColor(accent_rgb[0], accent_rgb[1], accent_rgb[2])
            
            # Kutu İçine Nano Banana 2 Promptunu Yazdırıyoruz
            p_text = placeholder.text_frame.paragraphs[0]
            nano_prompt = slide_data.get("image_prompt", "Cybersecurity concept art")
            p_text.text = f"🖼️ GÖRSEL ALANI\n\nBu kutuyu silip yerine Nano Banana 2 ile üreteceğiniz resmi koyun.\n\nKopyalamanız Gereken Prompt:\n'{nano_prompt}'"
            p_text.font.size = Pt(12)
            p_text.font.color.rgb = RGBColor(text_color, text_color, text_color)

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UYGULAMA ARAYÜZÜ ---
st.set_page_config(page_title="Cyber-Slide AI", page_icon="🛡️", layout="wide")
st.title("🛡️ Cyber-Slide AI: Akıllı Sunum Mimarı")
st.markdown("Kusursuz metinler ve Nano Banana 2 için özel görsel komutları üreten asistan.")

API_KEY = st.secrets.get("GEMINI_API_KEY", "")
if not API_KEY:
    st.error("API Key eksik!")
    st.stop()

genai.configure(api_key=API_KEY)

with st.sidebar:
    st.header("⚙️ Eğitim Parametreleri")
    topic = st.text_input("Konu Nedir?", "Zero Trust Architecture")
    language = st.selectbox("Sunum Dili", ["English", "Nederlands (Dutch)"])
    slide_count = st.slider("Slayt Sayısı", min_value=3, max_value=20, value=5)
    design_prompt = st.text_area("Tasarım", "Koyu arka plan, siberpunk neon yeşil başlıklar.")

if st.button("🚀 Sunumu Üret", type="primary"):
    with st.spinner("Slaytlar tasarlanıyor ve Nano Banana 2 komutları yazılıyor..."):
        model = genai.GenerativeModel('gemini-2.5-flash', generation_config={"response_mime_type": "application/json"})
        
        system_prompt = f"""
        You are an elite Cybersecurity Instructor. Generate a presentation.
        Topic: {topic}
        Language: {language}
        Slide Count: Exactly {slide_count} slides.
        Design Vibe: {design_prompt}
        
        Rules for Layouts:
        - "text_only": Standard slide, full width text.
        - "text_with_image_placeholder": Text on the left, and you MUST provide an "image_prompt".
        
        The "image_prompt" should be a highly detailed, professional English prompt designed for an AI image generator (like Nano Banana 2 / Midjourney). Include style keywords like "cyberpunk", "hacker aesthetic", "3D render", "high tech", "infographic style".
        
        Output STRICTLY in this JSON structure:
        {{
          "presentation_metadata": {{
            "global_background_color_hex": "#111111",
            "global_accent_color_hex": "#00FF00"
          }},
          "slides": [
            {{
              "slide_number": 1,
              "layout_type": "text_with_image_placeholder",
              "slide_title": "Slide Title Here",
              "content_bullets": ["Point 1", "Point 2"],
              "image_prompt": "A glowing digital padlock in a dark neon cyberspace, network nodes connecting, 3D isometric view, unreal engine 5, highly detailed, cybersecurity concept."
            }}
          ]
        }}
        """
        
        try:
            response = model.generate_content(system_prompt)
            ppt_file = create_pptx(response.text)
            st.success("🎉 Sunumunuz hazır!")
            st.download_button("📥 PowerPoint Dosyasını İndir", data=ppt_file, file_name="Siber_Sunum_Pro.pptx")
        except Exception as e:
            st.error(f"Hata: {e}")
