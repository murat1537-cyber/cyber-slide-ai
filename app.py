import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import json
import io
import requests
import base64
import urllib.parse

# --- YARDIMCI FONKSİYONLAR ---
def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

def fetch_image_from_url(url):
    """Verilen URL'den resmi indirir ve PPTX için hazırlar."""
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except Exception as e:
        print(f"Resim indirilemedi: {e}")
        return None

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
        
        # Metin Kutusu (Sol Taraf)
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.5), Inches(5))
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        for bullet in slide_data["content_bullets"]:
            p = text_frame.add_paragraph()
            p.text = f"• {bullet}"
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(text_color, text_color, text_color)
            
        # GÖRSEL MOTORU (Sağ Taraf)
        visual_data = slide_data.get("visual_element", {})
        v_type = visual_data.get("type")
        
        image_stream = None
        
        if v_type == "ai_image":
            # Fütüristik Görsel Üretimi (Pollinations API)
            prompt = visual_data.get("prompt", "cybersecurity concept")
            # Güvenlik ve tasarım için promptu güçlendiriyoruz
            enhanced_prompt = f"{prompt}, highly detailed, 8k resolution, cyberpunk aesthetic, professional"
            encoded_prompt = urllib.parse.quote(enhanced_prompt)
            image_url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=800&height=600&nologo=true"
            image_stream = fetch_image_from_url(image_url)
            
        elif v_type == "mermaid":
            # Teknik Şema Üretimi (Mermaid.ink)
            code = visual_data.get("code", "graph TD;\nA-->B;")
            # Mermaid kodunu base64'e çevir
            encoded_code = base64.b64encode(code.encode('utf-8')).decode('ascii')
            # Koyu arka plan için şema renklerini ayarlayan bir yapı
            mermaid_url = f"https://mermaid.ink/img/{encoded_code}?bgColor=!black"
            image_stream = fetch_image_from_url(mermaid_url)

        # Eğer resim başarıyla üretildiyse sağ tarafa ekle
        if image_stream:
            try:
                # Resmi sağa hizala: Left=5.2, Top=1.8, Genişlik=4.3 inç
                slide.shapes.add_picture(image_stream, Inches(5.2), Inches(1.8), width=Inches(4.3))
            except Exception as e:
                print(f"Slayta resim eklenirken hata: {e}")

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UYGULAMA ARAYÜZÜ ---
st.set_page_config(page_title="Cyber-Slide AI", page_icon="🛡️", layout="wide")
st.title("🛡️ Cyber-Slide AI: Akıllı Sunum Mimarı")
st.markdown("Fütüristik görseller ve teknik şemalarla donatılmış siber güvenlik sunumları hazırlayın.")

API_KEY = st.secrets.get("GEMINI_API_KEY", "")
if not API_KEY:
    st.error("API Key eksik!")
    st.stop()

genai.configure(api_key=API_KEY)

with st.sidebar:
    st.header("⚙️ Eğitim Parametreleri")
    topic = st.text_input("Konu Nedir?", "Man in the Middle Attack (MitM)")
    language = st.selectbox("Sunum Dili", ["English", "Nederlands (Dutch)"])
    slide_count = st.slider("Slayt Sayısı", min_value=3, max_value=20, value=5)
    design_prompt = st.text_area("Tasarım", "Koyu arka plan, neon mavi ve hacker estetiği.")

if st.button("🚀 Sunumu Üret", type="primary"):
    with st.spinner("Gemini düşünüyor, AI resimler çiziyor ve şemalar derleniyor... (Bu işlem 30-40 saniye sürebilir)"):
        model = genai.GenerativeModel('gemini-2.5-flash', generation_config={"response_mime_type": "application/json"})
        
        system_prompt = f"""
        You are an elite Cybersecurity Instructor. Generate a presentation.
        Topic: {topic}
        Language: {language}
        Slide Count: Exactly {slide_count} slides.
        Design Vibe: {design_prompt}
        
        CRITICAL VISUAL RULES:
        For EVERY slide, you MUST choose between two visual types:
        1. "ai_image": For abstract, futuristic, or conceptual slides. Provide a highly descriptive English prompt for an AI image generator (e.g., "A glowing digital padlock in a neon cyberspace, 3D render").
        2. "mermaid": For technical architectures, flows, or processes. Provide valid Mermaid.js flowchart code.
        
        Balance the presentation: Use both ai_image and mermaid slides.
        
        Output STRICTLY in this JSON structure:
        {{
          "presentation_metadata": {{
            "global_background_color_hex": "#111111",
            "global_accent_color_hex": "#00FFFF"
          }},
          "slides": [
            {{
              "slide_number": 1,
              "slide_title": "Slide Title",
              "content_bullets": ["Point 1", "Point 2"],
              "visual_element": {{
                "type": "ai_image",
                "prompt": "Hacker typing on a laptop with glowing blue code in the background, cinematic lighting"
              }}
            }},
            {{
              "slide_number": 2,
              "slide_title": "Attack Flow",
              "content_bullets": ["Step 1", "Step 2"],
              "visual_element": {{
                "type": "mermaid",
                "code": "graph LR\\n A[Attacker] --> B(Victim)\\n B --> C{{Server}}"
              }}
            }}
          ]
        }}
        """
        
        try:
            response = model.generate_content(system_prompt)
            ppt_file = create_pptx(response.text)
            st.success("🎉 Sunumunuz harika görsellerle hazırlandı!")
            st.download_button("📥 PowerPoint Dosyasını İndir", data=ppt_file, file_name="Gorsel_Siber_Sunum.pptx")
        except Exception as e:
            st.error(f"Hata: {e}")
