# app.py
# Full-featured AI PPT Generator
# ~450+ lines version with per-slide format + per-slide image upload

import os
import re
import io
import tempfile
import fitz
import docx
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ---------------- CONFIG ----------------
# You should replace GEMINI_API_KEY with your actual key if different.
GEMINI_API_KEY = "AIzaSyBtah4ZmuiVkSrJABE8wIjiEgunGXAbT3Q"
TEXT_MODEL_NAME = "gemini-2.0-flash"

# ---------------- LLM / GEMINI HELPERS ----------------
def call_gemini(prompt: str, timeout: int = 120) -> str:
    """
    Call Gemini (Generative Language) API with a prompt.
    Returns plain text or an error string if something fails.
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{TEXT_MODEL_NAME}:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        resp = requests.post(url, json=payload, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        # Defensive extraction — structure may vary
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        return f"⚠️ Gemini API error: {e}"

def generate_title(summary: str) -> str:
    """
    Generate a short title (under 10 words) for a given summary using Gemini.
    """
    if not summary or not summary.strip():
        return "Presentation"
    prompt = f"Generate a short, clear PowerPoint title (under 10 words) for this summary:\n{summary}"
    return call_gemini(prompt)

# ---------------- PARSING / OUTLINE HELPERS ----------------
def extract_slide_count(description: str, default=None):
    """
    Heuristic: find patterns like '10 slides' or '8 sections' to extract the slide count.
    """
    if not description:
        return None
    m = re.search(r"(\d+)\s*(slides?|sections?|pages?)", description, re.IGNORECASE)
    if m:
        total = int(m.group(1))
        # keep at least 1
        return max(1, total - 1)
    return None if default is None else default - 1

def parse_points(points_text: str):
    """
    Parse Gemini output into a list of slides:
      [{"title": ..., "description": "• bullet\n• bullet"}, ...]
    This parser is intentionally flexible to handle multiple output styles.
    """
    points = []
    current_title = None
    current_content = []

    if not points_text:
        return points

    # Normalize CRLF and strip trailing spaces
    lines = [ln.rstrip() for ln in points_text.replace("\r\n", "\n").split("\n")]

    for line in lines:
        if not line.strip():
            # blank line — treat as separator
            continue

        # skip lines that are the assistant asking a question or options
        if "Would you like" in line or line.strip().lower().startswith("option"):
            continue

        # slide header e.g., "Slide 1: Title"
        m = re.match(r"^\s*(Slide|Section)\s*(\d+)\s*[:.-]\s*(.+)$", line, re.IGNORECASE)
        if m:
            # flush previous
            if current_title:
                points.append({"title": current_title, "description": "\n".join(current_content)})
            current_title = m.group(3).strip()
            current_content = []
            continue

        # bullet lines beginning with -, •, * or numbered
        if re.match(r"^\s*[\-\u2022\*\d\)]\s+", line):
            # remove markers
            cleaned = re.sub(r"^\s*[\-\u2022\*\d\)]\s*", "", line).strip()
            if cleaned:
                current_content.append(f"• {cleaned}")
            continue

        # If the line looks like "Title: description" treat as a slide title plus bullet
        m2 = re.match(r"^\s*([^:]{3,80})\s*:\s*(.+)$", line)
        if m2 and not current_title:
            current_title = m2.group(1).strip()
            # add the rest as bullet
            current_content.append(f"• {m2.group(2).strip()}")
            continue

        # Generic content — if no title yet, treat as title fallback else as paragraph bullet
        if not current_title:
            current_title = line.strip()
        else:
            current_content.append(line.strip())

    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})

    return points

def generate_outline(description: str):
    """
    Ask Gemini to create an outline for the description.
    """
    if not description or not description.strip():
        return []
    num_slides = extract_slide_count(description, default=None)
    if num_slides:
        prompt = f"Create a PowerPoint outline on: {description}. Generate {num_slides} slides."
    else:
        prompt = f"Create a PowerPoint outline on: {description}. Each slide should have 3–4 bullet points."

    outline_text = call_gemini(prompt)
    slides = parse_points(outline_text)
    return slides

def edit_outline_with_feedback(outline, feedback: str):
    """
    Send the entire outline and feedback to Gemini to get a revised outline.
    """
    if not outline or "slides" not in outline:
        return outline
    outline_text = "\n".join([f"Slide {i+1}: {s['title']}\n{s['description']}" for i, s in enumerate(outline['slides'])])
    prompt = f"Refine the outline below based on feedback.\nOutline:\n{outline_text}\n\nFeedback:\n{feedback}"
    new_points = parse_points(call_gemini(prompt))
    return {"title": outline.get("title", "Presentation"), "slides": new_points}

# ---------------- LONG TEXT SUMMARIZATION ----------------
def split_text(text: str, chunk_size: int = 8000, overlap: int = 300):
    chunks = []
    start = 0
    n = len(text)
    while start < n:
        end = min(start + chunk_size, n)
        chunks.append(text[start:end])
        if end == n:
            break
        start = max(0, end - overlap)
    return chunks

def summarize_long_text(full_text: str) -> str:
    """
    Summarize a long text by chunking and asking Gemini to analyze each chunk,
    then combining analyses.
    """
    if not full_text or not full_text.strip():
        return ""
    chunks = split_text(full_text, chunk_size=8000, overlap=400)
    if len(chunks) == 1:
        return call_gemini(f"Summarize in detail:\n{full_text}")
    analyses = []
    for i, ch in enumerate(chunks, 1):
        analyses.append(call_gemini(f"Analyze CHUNK {i}:\n{ch}"))
    combined = call_gemini("Combine these analyses into a detailed summary:\n\n" + "\n\n".join(analyses))
    return combined

# ---------------- FILE / TEXT EXTRACTION HELPERS ----------------
def extract_text(path: str, filename: str) -> str:
    """
    Extract text from PDF/DOCX/TXT.
    """
    name = filename.lower()
    try:
        if name.endswith(".pdf"):
            doc = fitz.open(path)
            full = []
            for page in doc:
                txt = page.get_text("text")
                if txt:
                    full.append(txt)
            doc.close()
            return "\n".join(full)
        elif name.endswith(".docx"):
            d = docx.Document(path)
            return "\n".join(p.text for p in d.paragraphs)
        elif name.endswith(".txt"):
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception as e:
        return ""
    return ""

def sanitize_filename(name: str) -> str:
    return re.sub(r'[^A-Za-z0-9_.-]', '_', name).strip("_")

def clean_title_text(title: str) -> str:
    return re.sub(r"\s+", " ", title.strip()) if title else "Presentation"

def hex_to_rgb(hex_color: str) -> RGBColor:
    """
    Convert #RRGGBB into pptx RGBColor
    """
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        hex_color = "000000"
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# ---------------- PPT CREATION ----------------
def _add_image_to_slide(slide, image_bytes, left, top, width=None, height=None):
    """
    Helper to add an image to a slide from bytes. If width/height provided, use them.
    """
    try:
        img_stream = io.BytesIO(image_bytes)
        if width and height:
            slide.shapes.add_picture(img_stream, left, top, width=width, height=height)
        elif width:
            slide.shapes.add_picture(img_stream, left, top, width=width)
        elif height:
            slide.shapes.add_picture(img_stream, left, top, height=height)
        else:
            slide.shapes.add_picture(img_stream, left, top)
    except Exception as e:
        # silently ignore image errors (we want PPT generation to continue)
        pass

def create_ppt(title, points, filename="output.pptx", title_size=30, text_size=22,
               font="Calibri", title_color="#5E2A84", text_color="#282828",
               background_color="#FFFFFF", theme="Custom",
               bg_title_path=None, bg_slide_path=None):
    """
    Create and save a PPTX using python-pptx.
    Respects per-slide formats stored in st.session_state.slide_formats.
    Also supports per-slide uploaded images stored in st.session_state.slide_images (bytes).
    """
    prs = Presentation()
    title = clean_title_text(title)

    def set_bg(slide, image_path):
        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            except Exception:
                # fallback to solid color fill if picture fails
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = hex_to_rgb(background_color)
        else:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(background_color)

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_bg(slide, bg_title_path)
    tb = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(8.5), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(title_size)
    p.font.bold = True
    p.font.name = font
    p.font.color.rgb = hex_to_rgb(title_color)
    p.alignment = PP_ALIGN.LEFT

    # Content slides
    for idx, item in enumerate(points, start=1):
        key_point = clean_title_text(item.get("title", f"Slide {idx}"))
        description = item.get("description", "")
        # per-slide format (default Full Text)
        slide_format = st.session_state.get("slide_formats", {}).get(idx, st.session_state.get("slide_format", "Full Text"))
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        set_bg(slide, bg_slide_path)

        # Title box
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.0), Inches(1.0))
        tf_title = tb_title.text_frame
        p_title = tf_title.add_paragraph()
        p_title.text = key_point
        p_title.font.size = Pt(title_size)
        p_title.font.bold = True
        p_title.font.name = font
        p_title.font.color.rgb = hex_to_rgb(title_color)
        p_title.alignment = PP_ALIGN.LEFT

        # Body box size depends on format
        if description:
            if slide_format == "Text & Image":
                tb_body = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(5.0), Inches(4.0))
            else:
                tb_body = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(8.0), Inches(4.0))

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

        # If Text & Image and user uploaded an image for this slide, insert it into right area
        if slide_format == "Text & Image":
            img_bytes = None
            slide_images = st.session_state.get("slide_images", {})
            if slide_images.get(idx):
                img_bytes = slide_images.get(idx)
            if img_bytes:
                # place image at right side (approx)
                left = Inches(6.0)
                top = Inches(1.6)
                width = Inches(3.0)
                # add image
                _add_image_to_slide(slide, img_bytes, left, top, width=width)
            else:
                # optional placeholder rectangle
                left = Inches(6.0)
                top = Inches(1.6)
                width = Inches(3.0)
                height = Inches(3.0)
                try:
                    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor(240, 240, 240)
                    shape.line.color.rgb = RGBColor(200, 200, 200)
                    shape.text = "Image Placeholder"
                except Exception:
                    pass

        # Footer
        tb_footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.4))
        tf_footer = tb_footer.text_frame
        p_footer = tf_footer.add_paragraph()
        p_footer.text = "Generated with AI"
        p_footer.font.size = Pt(10)
        p_footer.font.name = font
        p_footer.font.color.rgb = RGBColor(120, 120, 120)
        p_footer.alignment = PP_ALIGN.RIGHT

    # Save to file
    prs.save(filename)
    return filename

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="AI PPT Generator", layout="wide")
st.title("🧠 AI PPT Generator — Per-slide Formats & Images")

# Initialize session state defaults
_defaults = {
    "messages": [],
    "doc_chat_history": [],
    "outline_chat": None,         # dict {"title": ..., "slides": [...]}
    "summary_text": None,
    "summary_title": None,
    "title_size": 30,
    "text_size": 22,
    "font_choice": "Calibri",
    "title_color": "#5E2A84",
    "text_color": "#282828",
    "bg_color": "#FFFFFF",
    "theme": "Custom",
    "slide_format": "Full Text",  # default global
    "slide_formats": {},          # per-slide formats, stored as {slide_index: "Full Text"|"Text & Image"}
    "slide_images": {},           # per-slide uploaded images stored as {slide_index: bytes}
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ---------- Top customization panel ----------
st.subheader("🎨 Customize PPT Style")
colA, colB = st.columns(2)
with colA:
    st.session_state.title_size = st.number_input("📌 Title Font Size", min_value=12, max_value=80, value=st.session_state.title_size)
with colB:
    st.session_state.text_size = st.number_input("📝 Text Font Size", min_value=8, max_value=48, value=st.session_state.text_size)

st.session_state.font_choice = st.selectbox(
    "🔤 Font Family",
    ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"],
    index=0 if st.session_state.font_choice not in ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"] else ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"].index(st.session_state.font_choice)
)

colC, colD, colE = st.columns(3)
with colC:
    st.session_state.title_color = st.color_picker("🎨 Title Color", st.session_state.title_color)
with colD:
    st.session_state.text_color = st.color_picker("📝 Text Color", st.session_state.text_color)
with colE:
    st.session_state.bg_color = st.color_picker("🌆 Background Color", st.session_state.bg_color)

st.session_state.theme = st.selectbox(
    "🎭 Select Theme",
    ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"],
    index=0 if st.session_state.theme not in ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"] else ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"].index(st.session_state.theme)
)

st.markdown("---")

# ---------- Upload document (optional) ----------
st.markdown("### 📄 Upload a document (optional) — PDF, DOCX or TXT")
uploaded_file = st.file_uploader("Upload a document to generate slides from", type=["pdf", "docx", "txt"])
if uploaded_file:
    with st.spinner("Extracting text..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            text = extract_text(tmp_path, uploaded_file.name)
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

        if text and text.strip():
            st.session_state.summary_text = summarize_long_text(text)
            st.session_state.summary_title = generate_title(st.session_state.summary_text)
            st.success("✅ Document processed. You can now ask to generate slides (type 'ppt' or 'slides').")
        else:
            st.error("❌ Could not extract text from the uploaded file.")

# ---------- Chat input (generate outline or ask questions) ----------
st.markdown("### 💬 Chat / Prompt")
chat_prompt = st.chat_input("Type a message (ask for 'ppt' or 'slides' to create an outline)...")
if chat_prompt:
    if st.session_state.summary_text:
        # If document is loaded, provide answers or create slides based on doc
        if any(w in chat_prompt.lower() for w in ["ppt", "slides", "presentation"]):
            with st.spinner("Generating outline from document and prompt..."):
                slides = generate_outline(st.session_state.summary_text + "\n\n" + chat_prompt)
                st.session_state.outline_chat = {"title": st.session_state.summary_title or "Presentation", "slides": slides}
                # reset per-slide images/formats (optional)
                st.session_state.slide_formats = {}
                st.session_state.slide_images = {}
        else:
            st.session_state.doc_chat_history.append(("user", chat_prompt))
            reply = call_gemini(f"Answer using this document:\n{st.session_state.summary_text}\n\nQ:{chat_prompt}")
            st.session_state.doc_chat_history.append(("assistant", reply))
    else:
        # No document: normal chat or generate slides from prompt
        st.session_state.messages.append(("user", chat_prompt))
        if any(w in chat_prompt.lower() for w in ["ppt", "slides", "presentation"]):
            with st.spinner("Generating outline..."):
                slides = generate_outline(chat_prompt)
                title = generate_title(chat_prompt)
                st.session_state.outline_chat = {"title": title or "Presentation", "slides": slides}
                st.session_state.slide_formats = {}
                st.session_state.slide_images = {}
        else:
            reply = call_gemini(chat_prompt)
            st.session_state.messages.append(("assistant", reply))
    st.rerun()

# ---------- show small recent chat history ----------
if st.session_state.get("messages"):
    with st.expander("Recent Chat (local)", expanded=False):
        for role, txt in st.session_state["messages"][-8:]:
            if role == "user":
                st.markdown(f"**You:** {txt}")
            else:
                st.markdown(f"**Assistant:** {txt}")

if st.session_state.get("doc_chat_history"):
    with st.expander("Document Chat (context)", expanded=False):
        for role, txt in st.session_state["doc_chat_history"][-8:]:
            if role == "user":
                st.markdown(f"**You (doc):** {txt}")
            else:
                st.markdown(f"**Assistant (doc):** {txt}")

st.markdown("---")

# ---------- Outline preview + per-slide editing UI ----------
if st.session_state.outline_chat:
    outline = st.session_state.outline_chat
    st.subheader(f"📝 Preview Outline: {outline.get('title', 'Presentation')}")
    st.write("Click a slide to expand it. For each slide you can add per-slide feedback, choose a format, upload an image, and apply an edit.")

    # Iterate slides and show per-slide controls
    for idx, slide in enumerate(outline.get("slides", []), start=1):
        with st.expander(f"Slide {idx}: {slide.get('title', '')}", expanded=False):
            # Display description nicely
            desc = slide.get("description", "")
            if desc:
                st.markdown(desc.replace("\n", "\n\n"))
            else:
                st.write("_No content for this slide yet._")

            # Feedback box
            feedback_key = f"feedback_{idx}"
            feedback = st.text_area(f"✏️ Feedback for Slide {idx}", key=feedback_key, height=90)

            # Two-column layout: left for format & uploader, right for edit button
            col_left, col_right = st.columns([3, 1])

            with col_left:
                # Per-slide format select (Full Text / Text & Image)
                current = st.session_state.get("slide_formats", {}).get(idx, st.session_state.get("slide_format", "Full Text"))
                format_key = f"format_{idx}"
                selected_format = st.selectbox(
                    f"📐 Format for Slide {idx}",
                    ["Full Text", "Text & Image"],
                    index=0 if current not in ["Full Text", "Text & Image"] else ["Full Text", "Text & Image"].index(current),
                    key=format_key
                )
                # persist
                st.session_state["slide_formats"][idx] = selected_format

                # If Text & Image selected, show image uploader for that slide
                if selected_format == "Text & Image":
                    img_key = f"slide_image_{idx}"
                    uploaded_img = st.file_uploader(f"🖼 Upload image for Slide {idx} (optional)", type=["png", "jpg", "jpeg"], key=img_key)
                    if uploaded_img:
                        try:
                            img_bytes = uploaded_img.getvalue()
                            st.session_state["slide_images"][idx] = img_bytes
                            st.image(img_bytes, caption=f"Preview Slide {idx} image", use_column_width=True)
                        except Exception:
                            st.warning("Could not read uploaded image.")
                    else:
                        # if previously uploaded image exists in session, show preview and an option to remove
                        if st.session_state.get("slide_images", {}).get(idx):
                            st.image(st.session_state["slide_images"][idx], caption=f"Current Slide {idx} image (uploaded)", use_column_width=True)
                            if st.button(f"Remove image for Slide {idx}", key=f"remove_img_{idx}"):
                                st.session_state["slide_images"].pop(idx, None)
                                st.success("Image removed.")
                else:
                    # If format is full text, offer to remove any uploaded image (cleanup)
                    if st.session_state.get("slide_images", {}).get(idx):
                        if st.button(f"Remove image for Slide {idx}", key=f"remove_img_ft_{idx}"):
                            st.session_state["slide_images"].pop(idx, None)
                            st.success("Image removed.")

            with col_right:
                # Edit slide button -> send the slide content & feedback to Gemini to refine
                edit_btn_key = f"edit_btn_{idx}"
                if st.button(f"💡 Edit Slide {idx}", key=edit_btn_key):
                    with st.spinner(f"Applying feedback to Slide {idx}..."):
                        # Construct a helpful prompt
                        prompt_lines = [
                            "You are an assistant that updates a PowerPoint slide.",
                            f"Slide Title: {slide.get('title','')}",
                            "Slide Content:",
                            slide.get("description", ""),
                            "",
                            "User Feedback:",
                            feedback or "(no feedback provided)",
                            "",
                            "Return only the updated bullet points or short paragraph text. Do not include commentary or metadata."
                        ]
                        prompt = "\n".join(prompt_lines)
                        result = call_gemini(prompt)
                        # Try to parse returned text into bullets/slide structure
                        parsed = parse_points(result)
                        if parsed:
                            # Replace only this slide content with the first parsed slide (or merge)
                            st.session_state.outline_chat["slides"][idx - 1] = parsed[0]
                            st.success(f"✅ Slide {idx} updated successfully.")
                            st.rerun()
                        else:
                            # fallback: attempt to use returned text as bullets by splitting lines
                            bullets = []
                            for l in result.splitlines():
                                l = l.strip()
                                if not l:
                                    continue
                                # treat as bullet if has punctuation or length
                                cleaned = re.sub(r"^[\-\u2022\*\d\)\.]+\s*", "", l)
                                if cleaned:
                                    bullets.append(f"• {cleaned}")
                            if bullets:
                                st.session_state.outline_chat["slides"][idx - 1]["description"] = "\n".join(bullets)
                                st.success(f"✅ Slide {idx} updated (fallback bullets).")
                                st.rerun()
                            else:
                                # last resort: try splitting into sentences
                                sents = [s.strip() for s in re.split(r"[.!?]\s+", result) if len(s.strip()) > 3]
                                if sents:
                                    st.session_state.outline_chat["slides"][idx - 1]["description"] = "\n".join(f"• {s}" for s in sents[:6])
                                    st.success(f"✅ Slide {idx} updated (sentence fallback).")
                                    st.rerun()
                                else:
                                    st.warning("Could not parse Gemini response. Try rephrasing your feedback or providing more specifics.")

    # After slide loop: global title edit, outline feedback, default slide format, and generate button
    st.markdown("---")
    st.subheader("Outline Controls")
    new_title = st.text_input("📌 Edit Presentation Title", value=outline.get("title", "Presentation"))
    outline_feedback = st.text_area("✏️ Feedback for the whole outline (optional)", height=140, placeholder="E.g., 'Make slides shorter', 'Add one slide on ethics'")

    # default slide format selector (applies to slides without per-slide choice when generating)
    default_format = st.selectbox("Default Slide Format (applies where per-slide format not set)", ["Full Text", "Text & Image"], index=0 if st.session_state.get("slide_format") not in ["Full Text", "Text & Image"] else ["Full Text", "Text & Image"].index(st.session_state.get("slide_format", "Full Text")))
    st.session_state["slide_format"] = default_format

    col_apply, col_generate, col_clear = st.columns([1, 1, 1])
    with col_apply:
        if st.button("🔄 Apply Feedback (outline-level)"):
            with st.spinner("Updating outline with feedback..."):
                updated = edit_outline_with_feedback(outline, outline_feedback)
                if updated and updated.get("slides"):
                    updated["title"] = new_title.strip() if new_title else updated.get("title", "Presentation")
                    st.session_state.outline_chat = updated
                    st.success("✅ Outline updated with feedback.")
                    # reset per-slide images/formats? we keep them to preserve user's choices
                    st.rerun()
                else:
                    st.warning("No changes returned from the assistant. Try rephrasing feedback.")

    with col_generate:
        if st.button("✅ Generate PPT"):
            with st.spinner("Generating PPT..."):
                filename = f"{sanitize_filename(new_title or outline.get('title','Presentation'))}.pptx"
                # choose background images based on theme (if those files exist)
                if st.session_state.get("theme") == "Dr.Reddys White Master":
                    bg_title = "/mnt/data/360_F_373501182_AW73b2wvfm9wBuar0JYwKBeF8NAUHDOH.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                elif st.session_state.get("theme") == "Dr.Reddys Blue Master":
                    bg_title = "/mnt/data/studio-background-concept-abstract-empty-light-gradient-purple-studio-room-background-product_1258-52339.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                else:
                    bg_title = bg_slide = None

                # call create_ppt with all session settings
                create_ppt(
                    new_title or outline.get("title", "Presentation"),
                    outline.get("slides", []),
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

                # expose download
                try:
                    with open(filename, "rb") as f:
                        st.download_button("⬇️ Download PPT", data=f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                except Exception as e:
                    st.error(f"Failed to create download: {e}")

    with col_clear:
        if st.button("🧹 Clear Outline & Selections"):
            st.session_state.outline_chat = None
            st.session_state.slide_formats = {}
            st.session_state.slide_images = {}
            st.success("Cleared outline, formats and uploaded images.")
            st.rerun()

# End of app
