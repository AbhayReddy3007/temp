# app.py
import os, re, tempfile, fitz, docx, requests
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR

# ---------------- CONFIG ----------------
GEMINI_API_KEY = "AIzaSyBtah4ZmuiVkSrJABE8wIjiEgunGXAbT3Q"  # 🔑 Replace or use st.secrets
TEXT_MODEL_NAME = "gemini-2.0-flash"

# ---------------- GEMINI HELPERS ----------------
def call_gemini(prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{TEXT_MODEL_NAME}:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        resp = requests.post(url, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        return f"⚠️ Gemini API error: {e}"

def generate_title(summary: str) -> str:
    prompt = f"""Read the following summary and create a short, clear, presentation-style title.
- Keep it under 10 words.
Summary:
{summary}
"""
    return call_gemini(prompt).strip()

def extract_slide_count(description: str, default=None):
    m = re.search(r"(\d+)\s*(slides?|sections?|pages?)", description, re.IGNORECASE)
    if m:
        total = int(m.group(1))
        return max(1, total - 1)
    return None if default is None else default - 1

def parse_points(points_text: str):
    points, current_title, current_content = [], None, []
    lines = [re.sub(r"[#*>`]", "", ln).rstrip() for ln in points_text.splitlines()]
    for line in lines:
        if not line or "Would you like" in line:
            continue
        m = re.match(r"^\s*(Slide|Section)\s*(\d+)\s*:\s*(.+)$", line, re.IGNORECASE)
        if m:
            if current_title:
                points.append({"title": current_title, "description": "\n".join(current_content)})
            current_title, current_content = m.group(3).strip(), []
            continue
        if line.strip().startswith("-"):
            text = line.lstrip("-").strip()
            if text:
                current_content.append(f"• {text}")
        elif line.strip().startswith(("•", "*")) or line.startswith("  "):
            text = line.lstrip("•*").strip()
            if text:
                current_content.append(f"- {text}")
        else:
            if line.strip():
                current_content.append(line.strip())
    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})
    return points

def generate_outline(description: str):
    num_slides = extract_slide_count(description, default=None)
    if num_slides:
        prompt = f"""Create a PowerPoint outline on: {description}.
Generate exactly {num_slides} content slides (excluding title slide).
Format:
Slide 1: <Title>
- Bullet
- Bullet
"""
    else:
        prompt = f"""Create a PowerPoint outline on: {description}.
Each slide should have a short title and 3–4 bullet points.
Format:
Slide 1: <Title>
- Bullet
- Bullet
"""
    outline_text = call_gemini(prompt)
    return parse_points(outline_text)

def edit_outline_with_feedback(outline, feedback: str):
    outline_text = "\n".join(
        [f"Slide {i+1}: {s['title']}\n{s['description']}" for i, s in enumerate(outline['slides'])]
    )
    prompt = f"""
You are an assistant improving a PowerPoint outline.

Current Outline:
Title: {outline['title']}
{outline_text}

Feedback:
{feedback}

Task:
- Apply feedback and refine outline.
- Keep the same format.
"""
    updated_points = parse_points(call_gemini(prompt))
    return {"title": outline['title'], "slides": updated_points}

def split_text(text: str, chunk_size: int = 8000, overlap: int = 300):
    chunks, start, n = [], 0, len(text)
    while start < n:
        end = min(start + chunk_size, n)
        chunks.append(text[start:end])
        if end == n:
            break
        start = max(0, end - overlap)
    return chunks

def summarize_long_text(full_text: str) -> str:
    if not full_text or not full_text.strip():
        return ""
    chunks = split_text(full_text, chunk_size=8000, overlap=400)
    if len(chunks) <= 1:
        prompt = f"""Summarize the following text thoroughly:
{full_text}"""
        return call_gemini(prompt).strip()
    analyses = []
    for idx, ch in enumerate(chunks, start=1):
        prompt = f"""Analyze CHUNK {idx}:
{ch}"""
        analyses.append(call_gemini(prompt))
    combined = "\n\n".join(analyses)
    return call_gemini(f"Combine these into a full exhaustive summary:\n{combined}")

# ---------------- FILE UTILS ----------------
def extract_text(path: str, filename: str) -> str:
    name = filename.lower()
    if name.endswith(".pdf"):
        text_parts = []
        doc = fitz.open(path)
        for page in doc:
            text_parts.append(page.get_text("text"))
        doc.close()
        return "\n".join(text_parts)
    if name.endswith(".docx"):
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    if name.endswith(".txt"):
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return ""

def sanitize_filename(name: str) -> str:
    return re.sub(r'[^A-Za-z0-9_.-]', '_', name)

def clean_title_text(title: str) -> str:
    return re.sub(r"\s+", " ", title.strip()) if title else "Presentation"

def hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# ---------------- PPT GENERATOR ----------------
def create_ppt(title, points, filename="output.pptx", title_size=30, text_size=22,
               font="Calibri", title_color="#5E2A84", text_color="#282828",
               background_color="#FFFFFF", template_path=None):
    """Create PPT using optional template"""
    if template_path:
        prs = Presentation(template_path)
        # Remove existing slides but keep theme
        for _ in range(len(prs.slides)):
            xml_slides = prs.slides._sldIdLst
            slide_id = xml_slides[0]
            xml_slides.remove(slide_id)
    else:
        prs = Presentation()

    title = clean_title_text(title)

    # Title Slide
    slide_layout = prs.slide_layouts[0] if template_path else prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = title
    except Exception:
        tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        tf = tb.text_frame
        p = tf.add_paragraph()
        p.text = title
        p.font.bold = True
        p.font.size = Pt(title_size)
        p.font.name = font
        p.font.color.rgb = hex_to_rgb(title_color)
        p.alignment = PP_ALIGN.CENTER

    # Content Slides
    for item in points:
        key_point = clean_title_text(item.get("title", ""))
        description = item.get("description", "")
        slide_layout = prs.slide_layouts[1] if template_path else prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)

        try:
            slide.shapes.title.text = key_point
        except Exception:
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8), Inches(1))
            tf = tb.text_frame
            p = tf.add_paragraph()
            p.text = key_point
            p.font.bold = True
            p.font.size = Pt(title_size)
            p.font.name = font
            p.font.color.rgb = hex_to_rgb(title_color)

        if description:
            tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(7.5), Inches(4))
            tf = tb.text_frame
            for line in description.splitlines():
                if line.strip():
                    p = tf.add_paragraph()
                    p.text = line.strip("•-* ").strip()
                    p.font.size = Pt(text_size)
                    p.font.name = font
                    p.font.color.rgb = hex_to_rgb(text_color)
                    p.level = 0

    prs.save(filename)
    return filename

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="PPT Generator", layout="wide")
st.title("🧠 AI PPT Generator")

defaults = {
    "messages": [], 
    "outline_chat": None, 
    "summary_text": None, 
    "summary_title": None, 
    "doc_chat_history": [],
    "title_size": 30,
    "text_size": 22,
    "font_choice": "Calibri",
    "title_color": "#5E2A84",
    "text_color": "#282828",
    "bg_color": "#FFFFFF"
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- Customization Panel ---
st.subheader("🎨 Customize PPT Style")
col1, col2 = st.columns(2)
with col1:
    st.session_state.title_size = st.number_input("📌 Title Font Size", 10, 100, st.session_state.title_size)
with col2:
    st.session_state.text_size = st.number_input("📝 Text Font Size", 8, 60, st.session_state.text_size)

st.session_state.font_choice = st.selectbox(
    "🔤 Font Family",
    ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"],
    index=0
)

col3, col4, col5 = st.columns(3)
with col3:
    st.session_state.title_color = st.color_picker("🎨 Title Color", st.session_state.title_color)
with col4:
    st.session_state.text_color = st.color_picker("📝 Text Color", st.session_state.text_color)
with col5:
    st.session_state.bg_color = st.color_picker("🌆 Background Color", st.session_state.bg_color)

# --- Template Option ---
st.subheader("📂 Template Option")
use_template = st.radio(
    "Would you like to generate the PPT in a template?",
    ("No", "Yes (use uploaded template)"),
    index=0,
)
template_file = None
if use_template == "Yes (use uploaded template)":
    template_file = st.file_uploader("📤 Upload PowerPoint Template (.pptx)", type=["pptx"])

# --- Upload File ---
uploaded_file = st.file_uploader("📄 Upload a document", type=["pdf", "docx", "txt"])
if uploaded_file:
    with st.spinner("Processing file..."):
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        text = extract_text(tmp_path, uploaded_file.name)
        os.remove(tmp_path)
        if text.strip():
            summary = summarize_long_text(text)
            title = generate_title(summary)
            st.session_state.summary_text = summary
            st.session_state.summary_title = title
            st.success("✅ Document processed successfully!")
        else:
            st.error("❌ Could not read text from file.")

# --- Chat Input ---
if prompt := st.chat_input("💬 Type a message..."):
    if st.session_state.summary_text:
        if any(w in prompt.lower() for w in ["ppt", "slides", "presentation"]):
            slides = generate_outline(st.session_state.summary_text + "\n\n" + prompt)
            st.session_state.outline_chat = {"title": st.session_state.summary_title, "slides": slides}
            st.session_state.doc_chat_history.append(("assistant", "✅ Generated PPT outline."))
        else:
            st.session_state.doc_chat_history.append(("user", prompt))
            reply = call_gemini(f"Answer using this document:\n{st.session_state.summary_text}\n\nQ:{prompt}")
            st.session_state.doc_chat_history.append(("assistant", reply))
    else:
        st.session_state.messages.append(("user", prompt))
        if "ppt" in prompt.lower():
            slides = generate_outline(prompt)
            title = generate_title(prompt)
            st.session_state.outline_chat = {"title": title, "slides": slides}
            st.session_state.messages.append(("assistant", f"✅ PPT outline generated! Title: **{title}**"))
        else:
            reply = call_gemini(prompt)
            st.session_state.messages.append(("assistant", reply))
    st.rerun()

# --- Outline Preview + Generate PPT ---
if st.session_state.outline_chat:
    outline = st.session_state.outline_chat
    st.subheader(f"📝 Preview Outline: {outline['title']}")
    for idx, slide in enumerate(outline["slides"], start=1):
        with st.expander(f"Slide {idx}: {slide['title']}", expanded=False):
            st.markdown(slide["description"].replace("\n", "\n\n"))

    new_title = st.text_input("📌 Edit Title", value=outline.get("title", "Untitled"))
    feedback_box = st.text_area("✏️ Feedback for outline (optional):")

    col6, col7 = st.columns(2)
    with col6:
        if st.button("🔄 Apply Feedback"):
            with st.spinner("Updating outline..."):
                updated_outline = edit_outline_with_feedback(outline, feedback_box)
                updated_outline["title"] = new_title.strip() if new_title else updated_outline["title"]
                st.session_state.outline_chat = updated_outline
                st.success("✅ Outline updated!")
                st.rerun()

    with col7:
    if st.button("✅ Generate PPT"):
        with st.spinner("Generating PPT..."):
            filename = f"{sanitize_filename(new_title)}.pptx"

            # ✅ FIX: Write uploaded template to a temp file if provided
            temp_template_path = None
            if template_file:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_tpl:
                    tmp_tpl.write(template_file.getvalue())
                    temp_template_path = tmp_tpl.name

            create_ppt(
                new_title,
                outline["slides"],
                filename,
                title_size=int(st.session_state.title_size),
                text_size=int(st.session_state.text_size),
                font=st.session_state.font_choice,
                title_color=st.session_state.title_color,
                text_color=st.session_state.text_color,
                background_color=st.session_state.bg_color,
                template_path=temp_template_path,
            )

            with open(filename, "rb") as f:
                st.download_button(
                    "⬇️ Download PPT",
                    data=f,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            # ✅ Cleanup temp template
            if temp_template_path and os.path.exists(temp_template_path):
                os.remove(temp_template_path)

