# app.py
# Full-featured AI PPT Generator (per-slide feedback + per-slide format + global controls)
# Restored and extended to ~400+ lines as requested.

import os
import re
import tempfile
import fitz
import docx
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ---------------- CONFIG ----------------
# Replace with your own Gemini/LLM key if needed
GEMINI_API_KEY = "AIzaSyBtah4ZmuiVkSrJABE8wIjiEgunGXAbT3Q"
TEXT_MODEL_NAME = "gemini-2.0-flash"

# ---------------- GEMINI / LLM HELPERS ----------------
def call_gemini(prompt: str) -> str:
    """
    Call Gemini (Generative Language) endpoint with the provided prompt.
    Returns generated text or an error string on failure.
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{TEXT_MODEL_NAME}:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        resp = requests.post(url, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        # Defensive access in case structure is different
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        return f"⚠️ Gemini API error: {e}"

def generate_title(summary: str) -> str:
    """Ask Gemini to produce a concise title for a summary."""
    prompt = f"Generate a short, clear PowerPoint title (under 10 words) for this summary:\n{summary}"
    return call_gemini(prompt).strip()

def extract_slide_count(description: str, default=None):
    """
    Look for phrases like '10 slides' or '8 sections' in the user's description.
    Returns number of slides (int) or None if not present.
    """
    m = re.search(r"(\d+)\s*(slides?|sections?|pages?)", description, re.IGNORECASE)
    if m:
        total = int(m.group(1))
        # keep at least 1 slide
        return max(1, total - 1)
    return None if default is None else default - 1

# ---------------- PARSER for Gemini output ----------------
def parse_points(points_text: str):
    """
    Parse Gemini/assistant output into a list of slides in this format:
    [
      {"title": "Slide 1 title", "description": "• point1\n• point2\n..."},
      ...
    ]
    Accepts several formats and tries to be robust.
    """
    points, current_title, current_content = [], None, []
    lines = [re.sub(r"[#*>`]", "", ln).rstrip() for ln in points_text.splitlines()]

    for line in lines:
        if not line:
            continue
        if "Would you like" in line:
            continue

        # matches "Slide 1: Title" or "Section 1: Title"
        m = re.match(r"^\s*(Slide|Section)\s*(\d+)\s*:\s*(.+)$", line, re.IGNORECASE)
        if m:
            if current_title:
                points.append({"title": current_title, "description": "\n".join(current_content)})
            current_title, current_content = m.group(3).strip(), []
            continue

        # bullet lines beginning with -, •, or *
        if re.match(r"^\s*[\-\u2022\*]\s+", line):
            text = re.sub(r"^\s*[\-\u2022\*]\s*", "", line).strip()
            if text:
                current_content.append(f"• {text}")
            continue

        # If none of the above, treat as plain content line (add to current content)
        if line.strip():
            # if there's no title yet, treat the first non-empty line as title fallback
            if not current_title:
                current_title = line.strip()
            else:
                current_content.append(line.strip())

    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})

    return points

# ---------------- Outline generation / editing ----------------
def generate_outline(description: str):
    """
    Generate a slide outline for given topic/description using Gemini.
    """
    num_slides = extract_slide_count(description, default=None)
    if num_slides:
        prompt = f"Create a PowerPoint outline on: {description}. Generate {num_slides} slides."
    else:
        prompt = f"Create a PowerPoint outline on: {description}. Each slide should have 3–4 bullet points."

    outline_text = call_gemini(prompt)
    return parse_points(outline_text)

def edit_outline_with_feedback(outline, feedback: str):
    """
    Send existing outline + feedback back to Gemini to refine the entire outline.
    """
    outline_text = "\n".join(
        [f"Slide {i+1}: {s['title']}\n{s['description']}" for i, s in enumerate(outline['slides'])]
    )
    prompt = f"Refine the outline below based on feedback.\nOutline:\n{outline_text}\nFeedback:\n{feedback}"
    updated_points = parse_points(call_gemini(prompt))
    # Keep title unchanged; return updated slides
    return {"title": outline['title'], "slides": updated_points}

# ---------------- Text chunking / summarization ----------------
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
    """
    Summarize a long document by chunking and combining analyses.
    """
    if not full_text.strip():
        return ""
    chunks = split_text(full_text, 8000, 400)
    if len(chunks) <= 1:
        return call_gemini(f"Summarize in detail:\n{full_text}")
    analyses = [call_gemini(f"Analyze CHUNK {i}:\n{ch}") for i, ch in enumerate(chunks, 1)]
    return call_gemini("Combine these analyses into a detailed summary:\n" + "\n\n".join(analyses))

# ---------------- FILE UTILITIES ----------------
def extract_text(path: str, filename: str) -> str:
    """
    Extract text from pdf/docx/txt.
    """
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

# ---------------- PPT CREATION ----------------
def create_ppt(title, points, filename="output.pptx", title_size=30, text_size=22,
               font="Calibri", title_color="#5E2A84", text_color="#282828",
               background_color="#FFFFFF", theme="Custom",
               bg_title_path=None, bg_slide_path=None):
    """
    Build a PowerPoint .pptx file using python-pptx.
    Honors per-slide formatting stored in st.session_state.slide_formats (if present).
    """
    prs = Presentation()
    title = clean_title_text(title)

    def set_bg(slide, image_path):
        if not image_path:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(background_color)
            return
        slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Title slide
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

    # Content slides
    for idx, item in enumerate(points, start=1):
        key_point = clean_title_text(item.get("title", ""))
        description = item.get("description", "")
        # obtain per-slide format choice; default to Full Text
        slide_format = st.session_state.get("slide_formats", {}).get(idx, st.session_state.get("slide_format", "Full Text"))

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        set_bg(slide, bg_slide_path)

        # Title textbox
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8.4), Inches(1.0))
        tf_title = tb_title.text_frame
        p_title = tf_title.add_paragraph()
        p_title.text = key_point
        p_title.font.bold = True
        p_title.font.size = Pt(title_size)
        p_title.font.name = font
        p_title.font.color.rgb = hex_to_rgb(title_color)
        p_title.alignment = PP_ALIGN.LEFT

        # Body textbox size depends on slide format
        if description:
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

            # Optional image placeholder when Text & Image chosen
            if slide_format == "Text & Image":
                left = Inches(6.2)
                top = Inches(2.0)
                width = Inches(3.0)
                height = Inches(3.5)
                shape = slide.shapes.add_shape(1, left, top, width, height)  # Rectangle
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(235, 235, 235)
                line = shape.line
                line.color.rgb = RGBColor(180, 180, 180)
                shape.text = "Image Placeholder"

        # Footer
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

# ------------- default session state keys -------------
defaults = {
    "messages": [],               # simple chat messages (user/assistant)
    "doc_chat_history": [],       # messages when a document is uploaded
    "outline_chat": None,         # current outline {'title': ..., 'slides': [...]}
    "summary_text": None,         # document summary text
    "summary_title": None,        # title generated from summary
    "title_size": 30,
    "text_size": 22,
    "font_choice": "Calibri",
    "title_color": "#5E2A84",
    "text_color": "#282828",
    "bg_color": "#FFFFFF",
    "theme": "Custom",
    "slide_format": "Full Text",   # default for whole deck (if per-slide format not provided)
    "slide_formats": {}            # per-slide formats stored as {slide_index: "Full Text" or "Text & Image"}
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ------------- Customization panel -------------
st.subheader("🎨 Customize PPT Style")
col1, col2 = st.columns(2)
with col1:
    st.session_state.title_size = st.number_input(
        "📌 Title Font Size", min_value=10, max_value=100, value=st.session_state.title_size, step=1
    )
with col2:
    st.session_state.text_size = st.number_input(
        "📝 Text Font Size", min_value=8, max_value=60, value=st.session_state.text_size, step=1
    )

st.session_state.font_choice = st.selectbox(
    "🔤 Font Family",
    ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"],
    index=max(0, ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"].index(st.session_state.font_choice) if st.session_state.font_choice in ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Helvetica", "Comic Sans MS"] else 0)
)

col3, col4, col5 = st.columns(3)
with col3:
    st.session_state.title_color = st.color_picker("🎨 Title Color", st.session_state.title_color)
with col4:
    st.session_state.text_color = st.color_picker("📝 Text Color", st.session_state.text_color)
with col5:
    st.session_state.bg_color = st.color_picker("🌆 Background Color", st.session_state.bg_color)

st.session_state.theme = st.selectbox(
    "🎭 Select Theme",
    ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"],
    index=0 if st.session_state.theme not in ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"] else ["Dr.Reddys White Master", "Dr.Reddys Blue Master", "Custom"].index(st.session_state.theme)
)

# ------------- Upload document (pdf/docx/txt) -------------
st.markdown("### 📄 Upload a document (optional)")
uploaded_file = st.file_uploader("Upload a document to create slides from", type=["pdf", "docx", "txt"])
if uploaded_file:
    with st.spinner("Processing file..."):
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        text = extract_text(tmp_path, uploaded_file.name)
        try:
            os.remove(tmp_path)
        except Exception:
            pass

        if text.strip():
            # Summarize the document and generate title using Gemini
            summary = summarize_long_text(text)
            title = generate_title(summary)
            st.session_state.summary_text = summary
            st.session_state.summary_title = title
            st.success("✅ Document processed successfully!")
        else:
            st.error("❌ Could not read text from file.")

# ------------- Chat input -------------
st.markdown("### 💬 Chat with the AI / Generate outline")
if prompt := st.chat_input("Type a message (ask for 'ppt' to create slides)..."):
    if st.session_state.summary_text:
        # If a document is loaded, provide answers using it or generate slides based on doc
        if any(w in prompt.lower() for w in ["ppt", "slides", "presentation"]):
            slides = generate_outline(st.session_state.summary_text + "\n\n" + prompt)
            st.session_state.outline_chat = {"title": st.session_state.summary_title, "slides": slides}
        else:
            st.session_state.doc_chat_history.append(("user", prompt))
            reply = call_gemini(f"Answer using this document:\n{st.session_state.summary_text}\n\nQ:{prompt}")
            st.session_state.doc_chat_history.append(("assistant", reply))
    else:
        # No document: treat as normal conversation or generate slides from prompt
        st.session_state.messages.append(("user", prompt))
        if "ppt" in prompt.lower():
            slides = generate_outline(prompt)
            title = generate_title(prompt)
            st.session_state.outline_chat = {"title": title, "slides": slides}
        else:
            reply = call_gemini(prompt)
            st.session_state.messages.append(("assistant", reply))
    st.rerun()

# Optional: display a short chat history so user sees the last messages
if st.session_state.get("messages"):
    st.markdown("#### Recent Chat")
    for role, text in st.session_state["messages"][-6:]:
        if role == "user":
            st.markdown(f"**You:** {text}")
        else:
            st.markdown(f"**Assistant:** {text}")

if st.session_state.get("doc_chat_history"):
    st.markdown("#### Document Chat")
    for role, text in st.session_state["doc_chat_history"][-6:]:
        if role == "user":
            st.markdown(f"**You (doc):** {text}")
        else:
            st.markdown(f"**Assistant (doc):** {text}")

# ------------- Outline preview + per-slide editing -------------
if st.session_state.outline_chat:
    outline = st.session_state.outline_chat
    st.subheader(f"📝 Preview Outline: {outline['title']}")

    # for each slide: show content, a feedback box, per-slide format dropdown, and an Edit button
    for idx, slide in enumerate(outline["slides"], start=1):
        with st.expander(f"Slide {idx}: {slide['title']}", expanded=False):
            st.markdown(slide["description"].replace("\n", "\n\n"))

            # Per-slide feedback text area
            feedback = st.text_area(f"✏️ Feedback for Slide {idx}", key=f"feedback_{idx}", height=80)

            # Per-slide format selection (Full Text or Text & Image)
            # Stored to st.session_state.slide_formats as {idx: "Full Text" / "Text & Image"}
            current_format = st.session_state.get("slide_formats", {}).get(idx, st.session_state.get("slide_format", "Full Text"))
            selected_format = st.selectbox(
                f"📐 Format for Slide {idx}",
                ["Full Text", "Text & Image"],
                index=0 if current_format not in ["Full Text", "Text & Image"] else ["Full Text", "Text & Image"].index(current_format),
                key=f"format_{idx}"
            )
            # save the selection
            st.session_state.slide_formats[idx] = selected_format

            # Per-slide edit button
            if st.button(f"💡 Edit Slide {idx}", key=f"edit_btn_{idx}"):
                with st.spinner(f"Updating Slide {idx}..."):
                    # instruct Gemini to return updated bullet points only
                    prompt = (
                        f"You are updating a PowerPoint slide based on feedback.\n\n"
                        f"Slide Title: {slide['title']}\n"
                        f"Slide Content:\n{slide['description']}\n\n"
                        f"Feedback:\n{feedback}\n\n"
                        f"Return ONLY the updated bullet points (each starting with • or -)."
                    )
                    updated_text = call_gemini(prompt)
                    updated_points = parse_points(updated_text)

                    if updated_points:
                        # Replace only this slide
                        st.session_state.outline_chat["slides"][idx - 1] = updated_points[0]
                        st.success(f"✅ Slide {idx} updated successfully!")
                        st.rerun()
                    else:
                        # fallback: try to detect bullets in the returned text without Slide headers
                        bullets = []
                        for line in updated_text.splitlines():
                            if re.match(r"^[\-\u2022\*\d\)]\s+", line.strip()):
                                bullets.append(re.sub(r"^[\-\u2022\*\d\)]\s*", "", line.strip()))
                        # if found bullets, put them into the slide description
                        if bullets:
                            st.session_state.outline_chat["slides"][idx - 1] = {
                                "title": slide['title'],
                                "description": "\n".join(f"• {b}" for b in bullets)
                            }
                            st.success(f"✅ Slide {idx} updated successfully (fallback)!")
                            st.rerun()
                        else:
                            # last resort: split into sentences if the reply contains sentences
                            if "." in updated_text:
                                sentences = [s.strip() for s in re.split(r"[.!?]", updated_text) if len(s.strip()) > 3]
                                if sentences:
                                    st.session_state.outline_chat["slides"][idx - 1] = {
                                        "title": slide['title'],
                                        "description": "\n".join(f"• {s}" for s in sentences[:6])  # limit to first few
                                    }
                                    st.success(f"✅ Slide {idx} updated (sentence fallback).")
                                    st.rerun()
                            # if nothing worked:
                            st.warning(f"⚠️ Could not parse updated content for Slide {idx}. Try rephrasing feedback.")

    # Global title edit and outline-level feedback
    new_title = st.text_input("📌 Edit Title", value=outline.get("title", "Untitled"))
    feedback_box = st.text_area("✏️ Feedback for outline (optional):", height=120)

    # Buttons: Apply feedback (global), Generate PPT
    col6, col7, col8 = st.columns([1, 1, 1])
    with col6:
        # Optionally allow setting a default slide format for new slides (applies when generating PPT)
        st.session_state.slide_format = st.selectbox(
            "Default Slide Format",
            ["Full Text", "Text & Image"],
            index=0 if st.session_state.slide_format not in ["Full Text", "Text & Image"] else ["Full Text", "Text & Image"].index(st.session_state.slide_format)
        )

    with col7:
        if st.button("🔄 Apply Feedback"):
            with st.spinner("Updating outline..."):
                updated_outline = edit_outline_with_feedback(outline, feedback_box)
                # ensure title uses any edited title
                updated_outline["title"] = new_title.strip() if new_title else updated_outline["title"]
                st.session_state.outline_chat = updated_outline
                st.success("✅ Outline updated!")
                st.rerun()

    with col8:
        if st.button("✅ Generate PPT"):
            with st.spinner("Generating PPT..."):
                filename = f"{sanitize_filename(new_title)}.pptx"
                # pick backgrounds per selected theme
                if st.session_state.theme == "Dr.Reddys White Master":
                    bg_title = "/mnt/data/360_F_373501182_AW73b2wvfm9wBuar0JYwKBeF8NAUHDOH.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                elif st.session_state.theme == "Dr.Reddys Blue Master":
                    bg_title = "/mnt/data/studio-background-concept-abstract-empty-light-gradient-purple-studio-room-background-product_1258-52339.jpg"
                    bg_slide = "/mnt/data/pastel-purple-color-solid-background-1920x1080.png"
                else:
                    bg_title = bg_slide = None

                # call create_ppt with current session settings and per-slide formats
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

                # present download button
                with open(filename, "rb") as f:
                    st.download_button(
                        "⬇️ Download PPT",
                        data=f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

# EOF
