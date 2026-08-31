import base64
import datetime
import io
import json
import os
from docxtpl import DocxTemplate
import google.generativeai as genai
import streamlit as st

# ==========================================
# 0. ตั้งค่าหน้าเว็บ (ต้องไว้บนสุดเสมอ)
# ==========================================
st.set_page_config(
    page_title="ระบบช่วยร่างและตรวจสอบหนังสือราชการ", page_icon="📝", layout="wide"
)

# ==========================================
# ระบบปรับแต่ง Background & CSS หัวเว็บ
# ==========================================
def set_background(image_file="background.jpg"):
  bg_image_css = ""
  if os.path.exists(image_file):
    try:
      with open(image_file, "rb") as f:
        encoded_string = base64.b64encode(f.read()).decode()
        bg_image_css = f"""
                    background-image: 
                        linear-gradient(rgba(255, 255, 255, 0.85), rgba(255, 255, 255, 0.85)), 
                        url(data:image/jpeg;base64,{encoded_string});
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                """
    except Exception:
      bg_image_css = "background-color: #f8fafc;"
  else:
    bg_image_css = "background-color: #f8fafc;"

  st.markdown(
      f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600&display=swap');
        
        html, body, [class*="css"] {{
            font-family: 'Kanit', sans-serif !important;
        }}

        .stApp {{
            {bg_image_css}
        }}

        .block-container {{
            background-color: rgba(255, 255, 255, 0.95); 
            border-radius: 0px;
            padding: 2.5rem 3rem;
            margin-top: 0rem;
            margin-bottom: 0rem;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
            max-width: 100% !important;
        }}

        @keyframes fadeInUpSmooth {{
            0% {{ opacity: 0; transform: translateY(22px); }}
            100% {{ opacity: 1; transform: translateY(0); }}
        }}

        .custom-navbar {{
            background-color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 15px 40px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            border-radius: 0px;
            margin-top: -2.5rem;
            margin-left: -3rem;
            margin-right: -3rem;
            border-bottom: 1px solid #E2E8F0;
            flex-wrap: wrap;
            gap: 10px;
        }}
        .nav-brand {{
            font-size: 18px;
            font-weight: 700;
            color: #581c87;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .nav-brand img {{ height: 32px; width: auto; object-fit: contain; }}
        .nav-links {{
            display: flex; gap: 20px; align-items: center; flex-wrap: wrap;
        }}
        .nav-links a {{
            text-decoration: none; color: #64748B; font-weight: 500; font-size: 14px;
            transition: color 0.3s; display: flex; align-items: center; gap: 6px; white-space: nowrap;
        }}
        .nav-links a:hover {{ color: #7c3aed; }}
        
        .hero-section {{
            background: linear-gradient(135deg, #581c87, #7c3aed);
            color: white; padding: 40px 5%; text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            margin-left: -3rem; margin-right: -3rem; margin-bottom: 30px;
            animation: fadeInUpSmooth 1s cubic-bezier(0.16, 1, 0.3, 1) both;
        }}
        .hero-section h1 {{
            font-size: 32px !important; font-weight: 700 !important; margin-bottom: 10px !important;
            color: white !important; background: none !important; -webkit-text-fill-color: white !important;
            animation: none !important; text-align: center; padding: 0;
        }}
        .hero-section p {{ font-size: 18px; font-weight: 300; opacity: 0.9; margin: 0; color: #f3e8ff; text-align: center; }}

        .nav-toggle {{ display: none; }}
        .nav-toggle-label {{ display: none; font-size: 24px; cursor: pointer; color: #581c87; padding: 5px; }}

        @media (max-width: 1024px) {{
            .custom-navbar {{ justify-content: space-between; padding: 15px 20px; }}
            .nav-toggle-label {{ display: block; }}
            .nav-links {{ 
                display: none; width: 100%; flex-direction: column; align-items: flex-start;
                gap: 15px; padding-top: 15px; margin-top: 10px; border-top: 1px solid #E2E8F0;
            }}
            .nav-toggle:checked ~ .nav-links {{ display: flex; }}
            .hero-section h1 {{ font-size: 22px !important; }}
            .hero-section p {{ font-size: 15px; }}
        }}

        h2, h3 {{ color: #581c87 !important; font-weight: 600; }}

        .stButton>button, .stDownloadButton>button {{
            background: linear-gradient(135deg, #6d28d9, #7c3aed); color: white !important;
            font-weight: 500; font-size: 16px; border-radius: 12px; border: none;
            padding: 12px 24px; box-shadow: 0 4px 15px rgba(109, 40, 217, 0.3);
            transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1); width: 100%;
        }}
        .stButton>button:hover, .stDownloadButton>button:hover {{
            transform: translateY(-3px) scale(1.015); box-shadow: 0 8px 25px rgba(109, 40, 217, 0.45); color: white;
        }}

        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}
        </style>
        """,
      unsafe_allow_html=True,
  )

# เรียกใช้สไตล์ธีมหน้าเว็บ
set_background("background.jpg")

# Render HTML หัวเว็บ
st.markdown(
    """
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<div class="custom-navbar">
<div class="nav-brand">
<img src="https://lh3.googleusercontent.com/d/1Ib-E-X35YqI8vQl7wpar_UXdoQYdc_1N" alt="Logo KU">
ฝ่ายพัฒนานิสิต คณะศึกษาศาสตร์ มก
</div>

<input type="checkbox" id="nav-toggle" class="nav-toggle">
<label for="nav-toggle" class="nav-toggle-label">
<i class="fa-solid fa-bars"></i>
</label>

<div class="nav-links">
<a href="https://canva.link/cbn78xyohbndm6z" target="_blank"><i class="fa-solid fa-list-check"></i> ขั้นตอนการเสนอโครงการ</a>
<a href="https://drive.google.com/drive/u/1/folders/1HDGo2ImRk_Szo5gXn5JnCXsu6swffRux" target="_blank"><i class="fa-solid fa-download"></i> ดาวโหลดแบบฟอร์มต่าง ๆ</a>
<a href="https://canva.link/fmw17m6mt6o0pok" target="_blank"><i class="fa-solid fa-award"></i> Template เกียรติบัตร</a>
<a href="https://www.edu.ku.ac.th/" target="_blank"><i class="fa-solid fa-globe"></i> เว็บไซต์คณะศึกษาศาสตร์</a>
</div>
</div>

<div class="hero-section">
<h1>📝 ระบบช่วยร่างและตรวจสอบหนังสือราชการ</h1>
<p>เครื่องมือช่วยสร้างและตรวจสอบหนังสือราชการตามรูปแบบมาตรฐาน</p>
</div>
""",
    unsafe_allow_html=True,
)

# ==========================================
# ระบบดึงวันที่ปัจจุบัน (ภาษาไทย)
# ==========================================
def get_thai_date():
  thai_months = [
      "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
      "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
  ]
  now = datetime.datetime.now()
  thai_year = now.year + 543
  thai_month = thai_months[now.month - 1]
  return f"{now.day} {thai_month} {thai_year}"

# ==========================================
# ระบบแทนที่คำใน Template Word
# ==========================================
def generate_word_from_template(template_path, context):
  try:
    doc = DocxTemplate(template_path)
    doc.render(context)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
  except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการสร้างไฟล์ Word: {e}")
    return None

# ==========================================
# ฟังก์ชันช่วยแกะ JSON จาก AI
# ==========================================
def parse_ai_json(response_text):
  try:
    clean_text = response_text.strip()
    if clean_text.startswith("```json"):
      clean_text = clean_text[7:-3].strip()
    elif clean_text.startswith("```"):
      clean_text = clean_text[3:-3].strip()
    return json.loads(clean_text)
  except json.JSONDecodeError:
    st.error("เกิดข้อผิดพลาดในการอ่านข้อมูลจาก AI กรุณาลองใหม่อีกครั้ง")
    return None

# ==========================================
# ฟังก์ชันระบบ Dynamic Form (ตามวัตถุประสงค์)
# ==========================================
def render_dynamic_form():
    doc_type = st.selectbox(
        "เลือกวัตถุประสงค์ของการทำหนังสือ:",
        [
            "[1] ขอความอนุเคราะห์ต่างๆ",
            "[2] เชิญวิทยากร / ขอเชิญเป็นผู้ทรงคุณวุฒิตัดสินผลในโครงการ",
            "[3] เชิญเป็นประธานในพิธีเปิด",
            "[4] ขออนุมัตินำนิสิตทำกิจกรรมนอกสถานที่",
            "[5] อื่น ๆ (โปรดระบุรายละเอียด)"
        ]
    )
    st.markdown("---")
    
    data = {"doc_type": doc_type, "attachment": "ไม่มี"}

    if doc_type == "[1] ขอความอนุเคราะห์ต่างๆ":
        data['subject'] = st.text_input("1. เรื่องที่ต้องการขอความอนุเคราะห์", placeholder="เช่น ขอความอนุเคราะห์ใช้สถานที่, ขอความอนุเคราะห์ยืมอุปกรณ์, ขอลาเรียนโดยไม่ถือเป็นวันลา")
        data['receiver'] = st.text_input("2. หนังสือฉบับนี้เรียนใคร (ตำแหน่งผู้รับ)", placeholder="เช่น คณบดีคณะศึกษาศาสตร์, หัวหน้าภาควิชา...")
        data['project_name'] = st.text_input("3. ชื่อโครงการ")
        data['objective'] = st.text_area("4. วัตถุประสงค์ของโครงการแบบย่อ ๆ", placeholder="สรุปใจความสำคัญสั้นๆ 1-2 บรรทัด")
        col1, col2 = st.columns(2)
        with col1:
            data['date_str'] = st.text_input("5. วันที่ดำเนินโครงการ", placeholder="เช่น วันเสาร์ที่ 14 มีนาคม 2569, ระหว่างวันที่ 11 - 12 มีนาคม 2569")
        with col2:
            data['time_str'] = st.text_input("6. ช่วงเวลาที่จัดโครงการ", placeholder="เช่น 8.30 - 16.30 น.")
        data['location'] = st.text_input("7. สถานที่จัดโครงการ", placeholder="เช่น ห้องประชุมสารนิเทศยุพา วีระไวทยะ")
        data['request_detail'] = st.text_area("8. รายละเอียดที่นิสิตจะขอความอนุเคราะห์", placeholder="ระบุสิ่งที่ต้องการให้ชัดเจน เช่น ขอใช้ห้องประชุมพร้อมเครื่องเสียง, ขอยืมโต๊ะพับจำนวน 5 ตัว, ขอให้นิสิตลาเรียนในรายวิชาโดยไม่ถือเป็นวันลา")
        data['attachment'] = st.text_input("9. เอกสารแนบท้าย (ถ้ามี)", placeholder="ระบุชื่อเอกสารแนบ เช่น สำเนาโครงการ..., กำหนดการ... (หากไม่มีให้เว้นว่างไว้)")

    elif doc_type == "[2] เชิญวิทยากร / ขอเชิญเป็นผู้ทรงคุณวุฒิตัดสินผลในโครงการ":
        st.info("💡 ระบบตั้งค่าเริ่มต้น: เรื่อง ขอความอนุเคราะห์บุคลากรในสังกัดของท่านเป็นวิทยากร/ผู้ทรงคุณวุฒิตัดสินผลในโครงการ")
        data['subject'] = "ขอความอนุเคราะห์บุคลากรในสังกัดของท่านเป็นวิทยากร/ขอความอนุเคราะห์บุคลากรในสังกัดของท่านเป็นผู้ทรงคุณวุฒิตัดสินผลในโครงการ"
        data['receiver'] = st.text_input("1. หนังสือฉบับนี้เรียนใคร", placeholder="ตำแหน่งผู้บังคับบัญชาของวิทยากร เช่น คณบดีคณะศึกษาศาสตร์, ผู้อำนวยการโรงเรียน...")
        data['project_name'] = st.text_input("2. ชื่อโครงการ")
        data['objective'] = st.text_area("3. วัตถุประสงค์ของโครงการแบบย่อ ๆ", placeholder="สรุปใจความสำคัญสั้นๆ 1-2 บรรทัด")
        data['date_str'] = st.text_input("4. วันที่ดำเนินโครงการ", placeholder="เช่น วันเสาร์ที่ 14 มีนาคม 2569, ระหว่างวันที่ 11 - 12 มีนาคม 2569")
        data['expert_name'] = st.text_input("5. ชื่อ-นามสกุล ของวิทยากร / ผู้ทรงคุณวุฒิ", placeholder="ระบุพร้อมคำนำหน้า/ตำแหน่งทางวิชาการ เช่น ผศ.ดร.สมชาย ใจดี")
        data['topic'] = st.text_input("6. หัวข้อที่ต้องการเชิญมาบรรยาย", placeholder="เช่น การพัฒนาบุคลิกภาพ (หากเป็นผู้ทรงคุณวุฒิตัดสินผลให้เว้นว่างไว้)")
        data['expert_detail'] = st.text_area("7. รายละเอียด วัน-เวลา และสถานที่ ที่เชิญมาเป็นวิทยากร/ผู้ทรงคุณวุฒิตัดสินผล", placeholder="ระบุช่วงเวลาเฉพาะของวิทยากร/กรรมการ เช่น ในวันที่ 14 มี.ค. 2569 เวลา 09.00 - 12.00 น. ณ ห้องประชุมสารนิเทศ")
        data['attachment'] = st.text_input("8. เอกสารแนบท้าย (ถ้ามี)", placeholder="ระบุชื่อเอกสารแนบ เช่น กำหนดการโครงการ... (หากไม่มีให้เว้นว่างไว้)")

    elif doc_type == "[3] เชิญเป็นประธานในพิธีเปิด":
        st.info("💡 ระบบตั้งค่าเริ่มต้น: เรื่อง ขอเรียนเชิญเป็นประธานในพิธีเปิดโครงการ")
        data['subject'] = "ขอเรียนเชิญเป็นประธานในพิธีเปิดโครงการ"
        data['receiver'] = st.text_input("1. หนังสือฉบับนี้เรียนใคร", placeholder="ระบุตำแหน่ง หรือ ชื่อ-นามสกุล พร้อมตำแหน่ง เช่น คณบดีคณะศึกษาศาสตร์")
        data['project_name'] = st.text_input("2. ชื่อโครงการ")
        data['objective'] = st.text_area("3. วัตถุประสงค์ของโครงการแบบย่อ ๆ", placeholder="สรุปใจความสำคัญสั้นๆ 1-2 บรรทัด")
        data['date_str'] = st.text_input("4. วันที่ดำเนินโครงการ", placeholder="เช่น วันเสาร์ที่ 14 มีนาคม 2569, ระหว่างวันที่ 11 - 12 มีนาคม 2569")
        data['opening_detail'] = st.text_area("5. รายละเอียดวัน เวลา และสถานที่ ในพิธีเปิด", placeholder="ระบุช่วงเวลาเฉพาะของพิธีเปิด เช่น วันเสาร์ที่ 14 มีนาคม 2569 เวลา 09.00 - 10.00 น. ณ ห้องประชุมสารนิเทศ")
        data['attachment'] = st.text_input("6. เอกสารแนบท้าย (ถ้ามี)", placeholder="ระบุชื่อเอกสารแนบ เช่น กำหนดการพิธีเปิด... (หากไม่มีให้เว้นว่างไว้)")

    elif doc_type == "[4] ขออนุมัตินำนิสิตทำกิจกรรมนอกสถานที่":
        st.info("💡 ระบบตั้งค่าเริ่มต้น: เรื่อง ขออนุมัติทำกิจกรรมนอกสถานที่ | เรียน คณบดีคณะศึกษาศาสตร์")
        data['subject'] = "ขออนุมัติทำกิจกรรมนอกสถานที่"
        data['receiver'] = "คณบดีคณะศึกษาศาสตร์"
        data['project_name'] = st.text_input("1. ชื่อโครงการ")
        data['objective'] = st.text_area("2. วัตถุประสงค์ของโครงการแบบย่อ ๆ", placeholder="สรุปใจความสำคัญสั้นๆ 1-2 บรรทัด")
        data['date_str'] = st.text_input("5. วันที่ดำเนินโครงการ", placeholder="เช่น วันเสาร์ที่ 14 มีนาคม 2569, ระหว่างวันที่ 11 - 12 มีนาคม 2569")
        data['location'] = st.text_input("6. สถานที่จัดโครงการ", placeholder="เช่น โรงเรียนวัดดอนเจดีย์ อ.พนมทวน จ.กาญจนบุรี")
        data['student_count'] = st.text_input("7. จำนวนนิสิตเข้าร่วมกิจกรรมกี่คน?", placeholder="เช่น 50 คน")
        data['attachment'] = st.text_input("8. เอกสารแนบท้าย (ถ้ามี)", placeholder="ระบุชื่อเอกสารแนบ เช่น โครงการ, รายชื่อนิสิตผู้เข้าร่วม (หากไม่มีให้เว้นว่างไว้)")
        
    elif doc_type == "[5] อื่น ๆ (โปรดระบุรายละเอียด)":

        st.info("💡 ระบุรายละเอียดหนังสือที่ต้องการให้ AI ช่วยร่าง")

        data['subject'] = st.text_input("1. เรื่องที่จะดำเนินการ", placeholder="เช่น ขอความอนุเคราะห์สนับสนุนโครงการ, ขออนุมัติเลื่อนการจัดโครงการ ฯลฯ")

        data['receiver'] = st.text_input("2. หนังสือฉบับนี้เรียนใคร (ตำแหน่งผู้รับ)", placeholder="เช่น คณบดีคณะศึกษาศาสตร์")

        data['project_name'] = st.text_input("3. ชื่อโครงการ (ถ้ามี)", placeholder="ถ้าไม่มีให้เว้นว่างไว้")

        data['objective'] = st.text_area("4. สาเหตุ / วัตถุประสงค์ที่ต้องทำหนังสือฉบับนี้", placeholder="สรุปใจความสำคัญ หรือเหตุผลความจำเป็น")

        data['request_detail'] = st.text_area("5. รายละเอียดที่ต้องการให้ระบุในหนังสือ", placeholder="อยากให้ AI เขียนอธิบายอะไรบ้าง พิมพ์รายละเอียดมาได้เลยครับ")

        data['date_str'] = st.text_input("6. วันที่/ช่วงเวลา ที่เกี่ยวข้อง", placeholder="เช่น วันที่ดำเนินงาน หรือวันที่ต้องการขอเปลี่ยนแปลง")

        data['attachment'] = st.text_input("7. มีเอกสารแนบหรือไม่", placeholder="ถ้ามีให้พิมพ์ชื่อเอกสารแนบ ถ้าไม่มีพิมพ์ว่า ไม่มี")
    return data

def format_raw_info(data):
    raw_info = f"- ประเภทหนังสือ: {data.get('doc_type')}\n"
    key_map = {
        'subject': 'เรื่อง', 'receiver': 'เรียน (ผู้รับ)', 'project_name': 'ชื่อโครงการ',
        'objective': 'วัตถุประสงค์/สาเหตุ', 'date_str': 'วันที่เกี่ยวข้อง', 'time_str': 'เวลาที่จัด',
        'location': 'สถานที่', 'request_detail': 'รายละเอียดที่ต้องการให้ดำเนินการ', 'attachment': 'สิ่งที่ส่งมาด้วย',
        'expert_name': 'ชื่อวิทยากร/ผู้ทรงคุณวุฒิ', 'topic': 'หัวข้อที่บรรยาย', 
        'expert_detail': 'รายละเอียดกำหนดการบรรยาย/ตัดสินผล', 'opening_detail': 'รายละเอียดพิธีเปิด',
        'student_count': 'จำนวนนิสิตเข้าร่วม'
    }
    for key, value in data.items():
        if key != 'doc_type' and value:
            thai_key = key_map.get(key, key)
            raw_info += f"- {thai_key}: {value}\n"
    return raw_info

# ==========================================
# ระบบตั้งค่า API Key (ดึงอัตโนมัติจาก Secrets)
# ==========================================
if "GEMINI_API_KEY" in st.secrets:
  api_key = st.secrets["GEMINI_API_KEY"]
else:
  st.sidebar.header("⚙️ การตั้งค่า")
  api_key = st.sidebar.text_input(
      "ใส่ Google Gemini API Key ของคุณ", type="password"
  )

if not api_key:
  st.error(
      "⚠️ ไม่พบ API Key ในระบบ (กรุณาตั้งค่า Secrets ใน Streamlit Cloud หรือกรอก API Key ใน Sidebar)"
  )
  st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-3.5-flash-lite")

# ==========================================
# เมนูเลือกการทำงาน
# ==========================================
menu = st.sidebar.radio(
    "เลือกฟังก์ชันการทำงาน",
    [
        "ร่างหนังสือภายใน (บันทึกข้อความ)",
        "ร่างหนังสือภายนอก",
        "ตรวจทานหนังสือราชการ",
    ],
)

# ==========================================
# 1. ฟังก์ชัน: ร่างหนังสือภายใน
# ==========================================
if menu == "ร่างหนังสือภายใน (บันทึกข้อความ)":
  st.header("📄 ร่างหนังสือภายใน (บันทึกข้อความ)")

  org_type = st.radio(
      "ประเภทของหน่วยงานที่ออกหนังสือ:",
      ["สโมสรนิสิต / ชุมนุม", "สาขาวิชา / คณะ"],
      horizontal=True,
  )

  st.markdown("---")
  st.subheader("ส่วนหัวและข้อมูลติดต่อ")
  col1, col2 = st.columns(2)
  with col1:
    sender = st.text_input("ชื่อหน่วยงาน/ชมรม", placeholder="เช่น สโมสรนิสิตคณะศึกษาศาสตร์")
    coordinator_name = st.text_input("ชื่อผู้ประสานงาน", placeholder="เช่น นายใจดี มีสุข")
  with col2:
    date = st.text_input("วันที่", value=get_thai_date())
    coordinator_phone = st.text_input("เบอร์โทรศัพท์ผู้ประสานงาน", placeholder="เช่น ***-***-****")

  st.markdown("---")
  st.subheader("รายละเอียดข้อมูลในหนังสือ")
  form_data = render_dynamic_form()

  if st.button("✨ ให้ AI ร่างหนังสือภายใน", type="primary"):
    if not sender or not form_data.get('subject') or not form_data.get('receiver'):
      st.warning("⚠️ กรุณากรอกข้อมูลสำคัญ (ชื่อหน่วยงาน, เรื่อง, เรียน) ให้ครบถ้วนก่อนเริ่มประมวลผล")
    else:
      with st.spinner("พี่ AI กำลังเรียบเรียงและเกลาภาษาราชการทุกส่วน..."):
        target_template = "template_internal_club.docx" if org_type == "สโมสรนิสิต / ชุมนุม" else "template_internal_major.docx"
        raw_info = format_raw_info(form_data)

        prompt = f"""คุณคือหัวหน้างานสารบรรณระดับสูง หน้าที่ของคุณคือการนำข้อมูลดิบของนิสิต ไปเกลาและเรียบเรียงใหม่ทั้งหมดให้เป็น "ภาษาราชการทางการระดับสูงสุด"

ตัวอย่างรูปแบบภาษาที่ต้องการ (ใช้เป็นแนวทาง):
- การขออนุญาต/อนุมัติ: "ทางสโมสรนิสิตคณะศึกษาศาสตร์ ได้จัดโครงการ... โดยมีวัตถุประสงค์เพื่อ... ในการนี้ สโมสรนิสิตคณะศึกษาศาสตร์ จึงใคร่ขออนุญาตให้นิสิตที่มีรายชื่อดังต่อไปนี้..."
- การเรียนเชิญ: "ด้วยสาขาวิชาเอกการสอนวิทยาศาสตร์ ได้จัดโครงการ... ในการนี้จึงใคร่ขอเรียนเชิญท่านให้เกียรติเป็นประธานในพิธีเปิด..."
- คำลงท้าย: "จึงเรียนมาเพื่อโปรดพิจารณาอนุเคราะห์ด้วย จักขอบพระคุณยิ่ง" หรือ "จึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ จักขอบพระคุณยิ่ง" หรือปรับให้เข้ากับบริบทของหนังสือ

คำสั่งเกลาภาษาทุกส่วน:
1. "sender": เกลาชื่อส่วนราชการให้สมบูรณ์
2. "subject": เกลาชื่อเรื่องให้สั้น กระชับ ตรงประเด็น ห้ามยาวเกินไป
3. "receiver": เกลาชื่อผู้รับหนังสือให้เป็นตำแหน่งทางการ
4. "body": เรียบเรียงเนื้อหาจำนวน 2 ย่อหน้า ตามมาตรฐานหนังสือภายใน:
   - ย่อหน้าที่ 1: บังคับขึ้นต้นด้วย "ทาง..." หรือ "ด้วย..." หรือ "เนื่องด้วย..." อธิบายสาเหตุและความเป็นมา
   - ย่อหน้าที่ 2: บังคับขึ้นต้นด้วย "ในการนี้..." หรือ "ทั้งนี้..." ระบุความประสงค์ที่ต้องการให้ผู้รับพิจารณาหรือดำเนินการ
5. "conclusion": ย่อหน้าที่ 3 (สรุป) ให้แยกออกมา ใช้ประโยคสรุปทางการ เช่น "จึงเรียนมาเพื่อโปรดพิจารณา..." พร้อมลงท้าย "จักขอบพระคุณยิ่ง"

ข้อมูลดิบจากนิสิต:
- ชื่อส่วนราชการ: {sender}
{raw_info}

ส่งผลลัพธ์กลับมาเป็น JSON Format เท่านั้น:
{{
  "sender": "...",
  "subject": "...",
  "receiver": "...",
  "body": "...",
  "conclusion": "..."
}}"""

        try:
          response = model.generate_content(
              prompt, generation_config={"response_mime_type": "application/json"}
          )
          ai_data = parse_ai_json(response.text)

          if ai_data:
            st.success("พี่ร่างให้เรียบร้อยแล้วครับ น้องอย่าลืมเติมข้อมูลตรงช่องว่างที่เว้นไว้ (... หรือ [...]) นะครับ หากต้องการให้พี่ปรับแก้คำศัพท์ส่วนไหน หรือแก้ไขข้อมูลตรงไหน พิมพ์บอกมาได้เลยครับ!")

            st.markdown("### 📋 ตัวอย่างข้อความที่ AI ช่วยเกลาให้:")
            st.write(f"**ส่วนราชการ:** {ai_data.get('sender', '')}")
            st.write(f"**เรื่อง:** {ai_data.get('subject', '')}")
            st.write(f"**เรียน:** {ai_data.get('receiver', '')}")
            st.write(f"**ผู้ประสานงาน:** {coordinator_name} (โทร. {coordinator_phone})")
            st.info(f"**เนื้อหาหนังสือ (ย่อหน้าที่ 1 และ 2):**\n\n{ai_data.get('body', '')}")
            st.info(f"**ข้อความสรุป (ย่อหน้าที่ 3):**\n\n{ai_data.get('conclusion', '')}")

            context = {
                "sender": ai_data.get("sender", ""),
                "date": date,
                "subject": ai_data.get("subject", ""),
                "receiver": ai_data.get("receiver", ""),
                "body": ai_data.get("body", ""),
                "conclusion": ai_data.get("conclusion", ""),
                "coordinator_name": coordinator_name,
                "coordinator_phone": coordinator_phone,
            }

            if os.path.exists(target_template):
              docx_data = generate_word_from_template(target_template, context)
              if docx_data:
                st.download_button(
                    label="📥 ดาวน์โหลดหนังสือฉบับสมบูรณ์ (.docx)",
                    data=docx_data,
                    file_name=(f"บันทึกข้อความ_{ai_data.get('subject', 'document')}.docx"),
                    mime=("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                )
            else:
              st.error(f"⚠️ ไม่พบไฟล์แม่พิมพ์ {target_template} ในระบบ")
        except Exception as e:
          st.error(f"เกิดข้อผิดพลาดในการประมวลผล AI: {e}")

# ==========================================
# 2. ฟังก์ชัน: ร่างหนังสือภายนอก
# ==========================================
elif menu == "ร่างหนังสือภายนอก":
  st.header("🏢 ร่างหนังสือภายนอก")

  st.markdown("---")
  st.subheader("ส่วนหัวและข้อมูลติดต่อ")
  col1, col2 = st.columns(2)
  with col1:
    org = st.text_input("หน่วยงานผู้ออกหนังสือ", placeholder="เช่น คณะศึกษาศาสตร์ มหาวิทยาลัยเกษตรศาสตร์")
    coordinator_name = st.text_input("ชื่อผู้ประสานงาน", placeholder="เช่น นายใจดี มีสุข")
  with col2:
    date = st.text_input("วันที่", value=get_thai_date())
    coordinator_phone = st.text_input("เบอร์โทรศัพท์ผู้ประสานงาน", placeholder="เช่น ***-***-****")

  st.markdown("---")
  st.subheader("รายละเอียดข้อมูลในหนังสือ")
  form_data = render_dynamic_form()

  if st.button("✨ ให้ AI ร่างหนังสือภายนอก", type="primary"):
    if not org or not form_data.get('subject') or not form_data.get('receiver'):
      st.warning("⚠️ กรุณากรอกข้อมูลสำคัญ (หน่วยงานผู้ออกหนังสือ, เรื่อง, เรียน) ให้ครบถ้วนก่อนเริ่มประมวลผล")
    else:
      with st.spinner("พี่ AI กำลังเรียบเรียงและเกลาภาษาราชการทุกส่วน..."):
        raw_info = format_raw_info(form_data)

        prompt = f"""คุณคือหัวหน้างานสารบรรณระดับสูง หน้าที่ของคุณคือการนำข้อมูลดิบของนิสิตไปเกลาและเรียบเรียงใหม่ทั้งหมดให้เป็น "ภาษาราชการทางการระดับสูงสุด" สำหรับหนังสือภายนอก (ตราครุฑ)

ตัวอย่างรูปแบบภาษาที่ต้องการ (ใช้เป็นแนวทาง):
- การขอความอนุเคราะห์ (สปอนเซอร์): "ด้วยสโมสรนิสิตคณะศึกษาศาสตร์ มหาวิทยาลัยเกษตรศาสตร์ ได้รับอนุมัติให้จัดโครงการ... ในการนี้สโมสรนิสิตคณะศึกษาศาสตร์... จึงใคร่ขอความอนุเคราะห์ท่านสนับสนุน..."
- การเรียนเชิญเป็นวิทยากร: "เนื่องด้วยคณะศึกษาศาสตร์ มหาวิทยาลัยเกษตรศาสตร์ จะจัดโครงการ... ในการนี้ คณะศึกษาศาสตร์ มหาวิทยาลัยเกษตรศาสตร์ จึงขอความอนุเคราะห์ท่าน... มาเป็นวิทยากรให้ความรู้..."
- คำลงท้ายหนังสือภายนอก: "จึงเรียนมาเพื่อโปรดพิจารณา" (สั้นๆ ไม่มีคำว่าจักขอบพระคุณยิ่งต่อท้าย)

คำสั่งเกลาภาษาทุกส่วน:
1. "org": เกลาชื่อหน่วยงานผู้ออกหนังสือให้เป็นทางการ
2. "subject": เกลาชื่อเรื่องให้สั้น กระชับ ตรงประเด็น ห้ามยาวเกินไป
3. "receiver": เกลาชื่อ/ตำแหน่งผู้รับหนังสือให้สุภาพ
4. "attachment": เกลาข้อความสิ่งที่ส่งมาด้วย (ถ้ามี)
5. "body": เรียบเรียงเนื้อหาจำนวน 2 ย่อหน้า ตามมาตรฐานหนังสือภายนอก:
   - ย่อหน้าที่ 1: บังคับขึ้นต้นด้วย "ด้วย..." หรือ "เนื่องด้วย..."
   - ย่อหน้าที่ 2: บังคับขึ้นต้นด้วย "ในการนี้..."
6. "conclusion": ย่อหน้าที่ 3 (บทสรุป) ให้แยกออกมาต่างหาก โดยบังคับใช้ประโยคสรุป เช่น "จึงเรียนมาเพื่อโปรดพิจารณา" 

ข้อมูลดิบจากนิสิต:
- หน่วยงานผู้ออกหนังสือ: {org}
{raw_info}

ส่งผลลัพธ์กลับมาเป็น JSON Format เท่านั้น:
{{
  "org": "...",
  "subject": "...",
  "receiver": "...",
  "attachment": "...",
  "body": "...",
  "conclusion": "..."
}}"""

        try:
          response = model.generate_content(
              prompt, generation_config={"response_mime_type": "application/json"}
          )
          ai_data = parse_ai_json(response.text)

          if ai_data:
            st.success("พี่ร่างให้เรียบร้อยแล้วครับ น้องอย่าลืมเติมข้อมูลตรงช่องว่างที่เว้นไว้ (... หรือ [...]) นะครับ หากต้องการให้พี่ปรับแก้คำศัพท์ส่วนไหน หรือแก้ไขข้อมูลตรงไหน พิมพ์บอกมาได้เลยครับ!")

            st.markdown("### 📋 ตัวอย่างข้อความที่ AI ช่วยเกลาให้:")
            st.write(f"**หน่วยงานผู้ออกหนังสือ:** {ai_data.get('org', '')}")
            st.write(f"**เรื่อง:** {ai_data.get('subject', '')}")
            st.write(f"**เรียน:** {ai_data.get('receiver', '')}")
            st.write(f"**สิ่งที่ส่งมาด้วย:** {ai_data.get('attachment', '')}")
            st.write(f"**ผู้ประสานงาน:** {coordinator_name} (โทร. {coordinator_phone})")
            st.info(f"**เนื้อหาหนังสือ (ย่อหน้าที่ 1 และ 2):**\n\n{ai_data.get('body', '')}")
            st.info(f"**ข้อความสรุป (ย่อหน้าที่ 3):**\n\n{ai_data.get('conclusion', '')}")

            context = {
                "org": ai_data.get("org", ""),
                "date": date,
                "subject": ai_data.get("subject", ""),
                "receiver": ai_data.get("receiver", ""),
                "attachment": ai_data.get("attachment", ""),
                "body": ai_data.get("body", ""),
                "conclusion": ai_data.get("conclusion", ""),
                "coordinator_name": coordinator_name,
                "coordinator_phone": coordinator_phone,
            }

            target_template = "template_external.docx"
            if os.path.exists(target_template):
              docx_data = generate_word_from_template(target_template, context)
              if docx_data:
                st.download_button(
                    label="📥 ดาวน์โหลดหนังสือฉบับสมบูรณ์ (.docx)",
                    data=docx_data,
                    file_name=(f"หนังสือภายนอก_{ai_data.get('subject', 'document')}.docx"),
                    mime=("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                )
            else:
              st.error("⚠️ ไม่พบไฟล์แม่พิมพ์ template_external.docx ในระบบ")
        except Exception as e:
          st.error(f"เกิดข้อผิดพลาดในการประมวลผล AI: {e}")

# ==========================================
# 3. ฟังก์ชัน: ตรวจทานหนังสือราชการ
# ==========================================
elif menu == "ตรวจทานหนังสือราชการ":
  st.header("🔍 ให้ AI ช่วยตรวจทานหนังสือราชการ")

  draft_text_input = st.text_area(
      "วางเนื้อหาหนังสือราชการของคุณที่นี่เพื่อตรวจสอบคำผิดและภาษา", height=200
  )

  if st.button("🕵️‍♂️ เริ่มการตรวจทาน", type="primary"):
    if draft_text_input:
      with st.spinner("AI กำลังวิเคราะห์และตรวจทาน..."):
        prompt = f"""กรุณาตรวจร่างหนังสือราชการต่อไปนี้:
"{draft_text_input}"

โปรดชี้เป้าจุดที่ควรแก้ไข (คำผิด/ความเหมาะสม) และเกลาประโยคให้เป็นทางการมากขึ้น"""

        response = model.generate_content(prompt)
        st.write(response.text)
    else:
      st.error("กรุณาวางเนื้อหาหนังสือราชการก่อนกดตรวจทาน")
