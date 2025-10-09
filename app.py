# app.py
import os, re, tempfile, fitz, docx, requests
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ---------------- CONFIG ----------------
GEMINI_API_KEY = "AIzaSyBtah4ZmuiVkSrJABE8wIjiEgunGXAbT3Q"
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
    prompt = f"Generate a short, clear PowerPoint title (under 10 words) for this summary:\n{summary}"
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
        if line.strip().startswith("-") or line.strip().startswith("•") or line.strip().startswith("*"):
            text = line.lstrip("-•*").strip()
            if text:
                current_content.append(f"• {text}")
        else:
            if line.strip():
                current_content.append(line.strip())
    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})
    return points

def generate_outline(description: str):
    num_slides = extract_slide_count(description, default=None)
    if num_slides:
        prompt = f"Create a PowerPoint outline on: {description}. Generate {num_slides} slides."
    else:
        prompt = f"Create a PowerPoint outline on: {description}. Each slide should have 3–4 bullet points."
    outline_text = call_gemini(prompt)
    return parse_points(outline_text)

def edit_outline_with_feedback(outline, feedback: str):
    outline_text = "\n".join(
        [f"Slide {i+1}: {s['title']}\n{s['description']}" for i, s in enumerate(outline['slides'])]
    )
    prompt = f"Refine the outline below based on feedback.\nOutline:\n{outline_text}\nFeedback:\n{feedback}"
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
    if not full_text.strip():
        return ""
    chunks = split_text(full_text, 8000, 400)
    if len(chunks) <= 1:
        return call_gemini(f"Summarize in detail:\n{full_text}")
    analyses = [call_gemini(f"Analyze CHUNK {i}:\n{ch}") for i, ch in enumerate(chunks, 1)]
    return call_gemini("Combine these analyses into a detailed summary:\n" + "\n\n".join(analyses))

# ---------------- FILE UTILS ----------------
def extract_text(path: str, filename: str) -> str:
    name = filename.lower()
    if name.endswith(".pdf"):
        doc = fitz.open(path)
        text = "\n".join(page.get_text("text") for page in doc)
        doc.close()
        return text
    elif name.endswith(".docx"):
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    elif name.endswith(".txt"):
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
               background_color="#FFFFFF", theme="Custom",
               bg_title_path=None, bg_slide_path=None):
    prs = Presentation()
    title = clean_title_text(title)

    def set_bg(slide, image_path):
        if not image_path:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(background_color)
            return
        slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_bg(slide, bg_title_path)

    tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(title_size)
    p.font.bold = True
    p.font.name = font
    p.font.color.rgb = hex_to_rgb(title_color)
    p.alignment = PP_ALIGN.CENTER

    for item in points:
        key_point = clean_title_text(item.get("title", ""))
        description = item.get("description", "")
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        set_bg(slide, bg_slide_path)

        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8.4), Inches(1.0))
        tf_title = tb_title.text_frame
        p_title = tf_title.add_paragraph()
        p_title.text = key_point
        p_title.font.bold = True
        p_title.font.size = Pt(title_size)
        p_title.font.name = font
        p_title.font.color.rgb = hex_to_rgb(title_color)
        p_title.alignment = PP_ALIGN.LEFT

        if description:
            slide_format = st.session_state.get("slide_format", "Full Text")

            # Adjust layout based on format
            if slide_format == "Text & Image":
                tb_body = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(5.0), Inches(4.2))
            else:
                tb_body = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(7.5), Inches(4.2))

            tf_body = tb_body.text_frame
            tf_body.word_wrap = True

            for line in description.splitlines():
                if line.strip():
                    p_body = tf_body.add_paragraph()
                    p_body.text = line.strip("•-* ").strip()
                    p_body.font.size = Pt(text_size)
                    p_body.font.name = font
                    p_body.font.color.rgb = hex_to_rgb(text_color)
                    p_body.level = 0

            # Optional placeholder for image
            if slide_format == "Text & Image":
                left = Inches(6.2)
                top = Inches(2.0)
                width = Inches(3.0)
                height = Inches(3.5)
                shape = slide.shapes.add_shape(1, left, top, width, height)  # 1 = Rectangle
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(235, 235, 235)
                line = shape.line
                line.color.rgb = RGBColor(180, 180, 180)
                shape.text = "Image Placeholder"

        tb_footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.4))
        tf_footer = tb_footer.text_frame
        p_footer = tf_footer.add_paragraph()
        p_footer.text = "Generated with AI"
        p_footer.font.size = Pt(10)
        p_footer.font.name = font
        p_footer.font.color.rgb = RGBColor(150, 150, 150)
        p_footer.alignment = PP_ALIGN.RIGHT

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
    "bg_color": "#FFFFFF",
    "theme": "Custom",
    "slide_format": "Full Text"
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

# --- Theme Dropdown ---
st.session_state.theme = st.selectbox(
    "🎭 Select Theme",
    ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"],
    index=["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"].index(st.session_state.theme)
)

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
        else:
            reply = call_gemini(prompt)
            st.session_state.messages.append(("assistant", reply))
    st.rerun()

# --- Outline Preview + Generate PPT ---
if st.session_state.outline_chat:
    outline = st.session_state.outline_chat
    st.subheader(f"📝 Preview Outline: {outline['title']}")

    # Per-slide display
    for idx, slide in enumerate(outline["slides"], start=1):
        with st.expander(f"Slide {idx}: {slide['title']}", expanded=False):
            st.markdown(slide["description"].replace("\n", "\n\n"))

    # Title and feedback
    new_title = st.text_input("📌 Edit Title", value=outline.get("title", "Untitled"))
    feedback_box = st.text_area("✏️ Feedback for outline (optional):")

    # Buttons: Format selector + Apply Feedback + Generate PPT
    col6, col7, col8 = st.columns([1, 1, 1])

    with col6:
        st.session_state.slide_format = st.selectbox(
            "📐 Slide Format",
            ["Full Text", "Text & Image"],
            index=["Full Text", "Text & Image"].index(st.session_state.slide_format),
            help="Choose how slides should be formatted."
        )

    with col7:
        if st.button("🔄 Apply Feedback"):
            with st.spinner("Updating outline..."):
                updated_outline = edit_outline_with_feedback(outline, feedback_box)
                updated_outline["title"] = new_title.strip() if new_title else updated_outline["title"]
                st.session_state.outline_chat = updated_outline
                st.success("✅ Outline updated!")
                st.rerun()

    with col8:
        if st.button("✅ Generate PPT"):
            with st.spinner("Generating PPT..."):
                filename = f"{sanitize_filename(new_title)}.pptx"
                if st.session_state.theme == "Dr.Reddys White Master":
                    bg_title = "/mnt/data/360_F_373501182_AW73b2wvfm9wBuar0JYwKBeF8NAUHDOH.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                elif st.session_state.theme == "Dr.Reddys Blue Master":
                    bg_title = "/mnt/data/studio-background-concept-abstract-empty-light-gradient-purple-studio-room-background-product_1258-52339.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                else:
                    bg_title = bg_slide = None

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
                    theme=st.session_state.theme,
                    bg_title_path=bg_title,
                    bg_slide_path=bg_slide,
                )

                with open(filename, "rb") as f:
                    st.download_button(
                        "⬇️ Download PPT",
                        data=f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
