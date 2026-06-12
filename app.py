import os
import re
import io
import base64
import streamlit as st

from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from streamlit_js_eval import streamlit_js_eval


TEMPLATE_FILE = "Template.docx"


# -----------------------------
# DOCX helpers
# -----------------------------
def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def safe_filename(s: str) -> str:
    s = re.sub(r'[\\/*?:"<>|]+', "_", s)
    return s.strip().replace(" ", "_")


def iter_all_paragraphs(container):
    for p in container.paragraphs:
        yield p
    for t in container.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from iter_all_paragraphs(cell)


def paragraph_full_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def set_paragraph_text(paragraph, text: str):
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def replace_text_in_paragraph(paragraph, replacements):
    full_text = paragraph_full_text(paragraph)
    new_text = full_text
    for old, new in replacements.items():
        if old:
            new_text = new_text.replace(old, new)
    if new_text != full_text:
        set_paragraph_text(paragraph, new_text)


def replace_text_in_container(container, replacements):
    for p in iter_all_paragraphs(container):
        replace_text_in_paragraph(p, replacements)


def replace_everywhere(doc, replacements):
    replace_text_in_container(doc, replacements)
    for section in doc.sections:
        replace_text_in_container(section.header, replacements)
        replace_text_in_container(section.footer, replacements)


def find_paragraph_by_contains(container, needles, case_insensitive=True):
    for p in iter_all_paragraphs(container):
        txt = paragraph_full_text(p)
        check = txt.lower() if case_insensitive else txt
        for needle in needles:
            nd = needle.lower() if case_insensitive else needle
            if nd in check:
                return p
    return None


def find_any_anchor(doc, needles):
    p = find_paragraph_by_contains(doc, needles)
    if p:
        return p, "body"
    for i, section in enumerate(doc.sections):
        p = find_paragraph_by_contains(section.header, needles)
        if p:
            return p, f"header_{i}"
        p = find_paragraph_by_contains(section.footer, needles)
        if p:
            return p, f"footer_{i}"
    return None, None


def insert_paragraph_after(paragraph, text="", style=None):
    new_para = paragraph._parent.add_paragraph()
    paragraph._p.addnext(new_para._p)
    if text:
        new_para.add_run(text)
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    return new_para


def insert_image_after(paragraph, image_bytes, width=Inches(4.5), caption=None):
    p_img = insert_paragraph_after(paragraph)
    run = p_img.add_run()
    run.add_picture(io.BytesIO(image_bytes), width=width)

    if caption:
        p_cap = insert_paragraph_after(p_img, caption)
        try:
            p_cap.runs[0].italic = True
        except Exception:
            pass
        return p_img, p_cap
    return p_img, None


def add_hyperlink(paragraph, text, url, color="0000FF", underline=True):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)

    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)

    new_run.append(rPr)
    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


# -----------------------------
# Business logic
# -----------------------------
def build_olt_label(equipment: str, custom_olt_label: str) -> str:
    eq = normalize_spaces(equipment).lower()
    if eq == "nokia lightspan mf-2":
        return "OLT MF-2"
    return normalize_spaces(custom_olt_label) or normalize_spaces(equipment)


def build_fuse_line(olt_label: str, equipment: str) -> str:
    return f"FUSE No: L3 {olt_label} – ({normalize_spaces(equipment)} Power tapping point)"


def insert_power_section_content(doc, data, rs_entries):
    """
    rs_entries: list of dicts like:
      {"name": "...", "load": "...", "img_bytes": b"..."}
    Inserts only the number of RS provided.
    """
    fuse_variants = [
        "FUSE No: L3 Nokia OLT MF-2 – ( Nokia Power tapping point)",
        "FUSE No: L3 Nokia OLT MF-2 – (Nokia Power tapping point)",
        "FUSE No: L3 Nokia OLT MF-2",
        "FUSE No:"
    ]
    anchor, _ = find_any_anchor(doc, fuse_variants)
    if not anchor:
        raise ValueError("Could not find FUSE anchor text in template.")

    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])
    set_paragraph_text(anchor, build_fuse_line(olt_label, data["equipment"]))

    last = anchor
    for idx, rs in enumerate(rs_entries, start=1):
        line = f"RS{idx} + {rs['name']} + {rs['load']}"
        p = insert_paragraph_after(last, line)
        last = p
        if rs.get("img_bytes"):
            img_p, _ = insert_image_after(p, rs["img_bytes"], width=Inches(4.5), caption=f"RS{idx}")
            last = img_p


def insert_supporting_documents(doc, tssr_pdf_name: str):
    if not tssr_pdf_name:
        return

    anchor, _ = find_any_anchor(doc, ["Supporting Documents", "Supporting Document", "Attachments"])
    if not anchor:
        # fallback: append at end
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph("")

    insert_paragraph_after(anchor, f"TSSR: {tssr_pdf_name}")


def insert_existing_rectifier_image(doc, rectifier_img_bytes):
    if not rectifier_img_bytes:
        return "No rectifier image provided."

    # More flexible anchors (edit this list to match your template wording)
    rect_anchors = [
        "Existing Rectifier",
        "Rectifier",
        "RECTIFIER",
        "Rectifier Photo",
        "Photo of Rectifier",
        "Existing DC Plant",
    ]

    anchor, where = find_any_anchor(doc, rect_anchors)
    if not anchor:
        # Fallback: insert near the end instead of erroring out
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph("")
        insert_paragraph_after(anchor, "Existing Rectifier (Auto-inserted):")
        insert_image_after(anchor, rectifier_img_bytes, width=Inches(5.0), caption="Existing Rectifier")
        return "Rectifier anchor not found; image inserted at end of document."

    insert_image_after(anchor, rectifier_img_bytes, width=Inches(5.0), caption="Existing Rectifier")
    return f"Rectifier image inserted under anchor found in {where}."


def generate_docx_bytes(data, rs_entries, rectifier_img_bytes, tssr_pdf_name):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

    doc = Document(TEMPLATE_FILE)

    replacements = {
        "CDO-604": data["site_name"],
        "MIN699": data["plaid"],
        "Nokia Lightspan MF-2": data["equipment"],
        "John Carlo Rabanes": data["prepared_by"],
        "OLT Rollout Engineer": data["position"],
        "< May 19- June 19, 2026 10:00AM-6:00PM>": data["target_datetime"],
        "May 19- June 19, 2026 10:00AM-6:00PM": data["target_datetime"],
    }
    replace_everywhere(doc, replacements)

    insert_power_section_content(doc, data, rs_entries)
    insert_supporting_documents(doc, tssr_pdf_name)
    rect_note = insert_existing_rectifier_image(doc, rectifier_img_bytes)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue(), rect_note


# -----------------------------
# Clipboard via streamlit-js-eval
# -----------------------------
def clipboard_image_to_bytes(eval_key: str):
    js = """
    async () => {
      try {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          for (const type of item.types) {
            if (type.startsWith('image/')) {
              const blob = await item.getType(type);
              const dataUrl = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.readAsDataURL(blob);
              });
              return dataUrl;
            }
          }
        }
        return null;
      } catch (e) {
        return "ERROR:" + e.toString();
      }
    }
    """
    data_url = streamlit_js_eval(js_expressions=js, key=eval_key, want_output=True)
    if not data_url:
        return None
    if isinstance(data_url, str) and data_url.startswith("ERROR:"):
        st.warning("Clipboard read blocked by browser permission/policy.")
        return None
    if isinstance(data_url, str) and data_url.startswith("data:image/") and "," in data_url:
        b64 = data_url.split(",", 1)[1]
        return base64.b64decode(b64)
    return None


def uploaded_file_to_bytes(uploaded):
    return uploaded.getvalue() if uploaded else None


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="MOP Automation", layout="wide")
st.title("MOP Automation (Streamlit Cloud)")

if not os.path.exists(TEMPLATE_FILE):
    st.error("Template.docx not found in repo root. Commit Template.docx next to app.py.")
    st.stop()

col1, col2 = st.columns(2)

with col1:
    site_name = st.text_input("Site Name")
    plaid = st.text_input("Plaid")
    equipment = st.text_input("Equipment", value="Nokia Lightspan MF-2")
    olt_label_custom = st.text_input("Custom OLT Label (if not Nokia MF-2)")
    prepared_by = st.text_input("Prepared By")
    position = st.text_input("Position", value="OLT Rollout Engineer")
    target_datetime = st.text_input("Target Date and Time", value="May 19- June 19, 2026 10:00AM-6:00PM")

with col2:
    rs_count = st.selectbox("Number of RS", options=[1, 2], index=1)

st.divider()
st.subheader("RS Details")

rs_entries = []

st.markdown("### RS1")
rs1_name = st.text_input("RS1 Rectifier Name", key="rs1_name")
rs1_load = st.text_input("RS1 Load Assignment", key="rs1_load")
rs1_upload = st.file_uploader("Upload RS1 image", type=["png", "jpg", "jpeg", "bmp"], key="rs1_upload")
if st.button("Paste RS1 from clipboard", key="paste_rs1_btn"):
    st.session_state["rs1_clip_bytes"] = clipboard_image_to_bytes("rs1_clip_eval")
rs1_img_bytes = uploaded_file_to_bytes(rs1_upload) or st.session_state.get("rs1_clip_bytes")
if rs1_img_bytes:
    st.image(rs1_img_bytes, caption="RS1 image", width=260)

rs_entries.append({"name": rs1_name.strip(), "load": rs1_load.strip(), "img_bytes": rs1_img_bytes})

if rs_count == 2:
    st.markdown("### RS2")
    rs2_name = st.text_input("RS2 Rectifier Name", key="rs2_name")
    rs2_load = st.text_input("RS2 Load Assignment", key="rs2_load")
    rs2_upload = st.file_uploader("Upload RS2 image", type=["png", "jpg", "jpeg", "bmp"], key="rs2_upload")
    if st.button("Paste RS2 from clipboard", key="paste_rs2_btn"):
        st.session_state["rs2_clip_bytes"] = clipboard_image_to_bytes("rs2_clip_eval")
    rs2_img_bytes = uploaded_file_to_bytes(rs2_upload) or st.session_state.get("rs2_clip_bytes")
    if rs2_img_bytes:
        st.image(rs2_img_bytes, caption="RS2 image", width=260)

    rs_entries.append({"name": rs2_name.strip(), "load": rs2_load.strip(), "img_bytes": rs2_img_bytes})

st.divider()
st.subheader("Other Attachments")

rect_upload = st.file_uploader("Existing Rectifier image (page 8)", type=["png", "jpg", "jpeg", "bmp"], key="rect_upload")
if st.button("Paste Rectifier from clipboard", key="paste_rect_btn"):
    st.session_state["rect_clip_bytes"] = clipboard_image_to_bytes("rect_clip_eval")
rect_img_bytes = uploaded_file_to_bytes(rect_upload) or st.session_state.get("rect_clip_bytes")
if rect_img_bytes:
    st.image(rect_img_bytes, caption="Existing rectifier", width=300)

tssr_pdf = st.file_uploader("TSSR PDF (filename will be written into Word)", type=["pdf"], key="tssr_pdf")

data = {
    "site_name": site_name.strip(),
    "plaid": plaid.strip(),
    "equipment": equipment.strip(),
    "olt_label_custom": olt_label_custom.strip(),
    "prepared_by": prepared_by.strip(),
    "position": position.strip(),
    "target_datetime": target_datetime.strip(),
}

# Validation
required_base = ["site_name", "plaid", "equipment", "prepared_by", "position", "target_datetime"]
required_rs = ["name", "load"]

if st.button("Generate MOP (.docx)", type="primary"):
    missing = [k for k in required_base if not data.get(k)]
    if missing:
        st.error("Missing fields: " + ", ".join(missing))
        st.stop()

    # RS validation based on count
    for i, rs in enumerate(rs_entries, start=1):
        if i > rs_count:
            continue
        if not rs.get("name") or not rs.get("load"):
            st.error(f"Missing RS{i} Rectifier Name or Load Assignment.")
            st.stop()

    try:
        docx_bytes, rect_note = generate_docx_bytes(
            data=data,
            rs_entries=rs_entries[:rs_count],
            rectifier_img_bytes=rect_img_bytes,
            tssr_pdf_name=(tssr_pdf.name if tssr_pdf else "")
        )

        out_name = f"MOP_{safe_filename(data['site_name'])}_{safe_filename(data['plaid'])}.docx"
        st.success("Generated. " + rect_note)
        st.download_button(
            "Download MOP",
            data=docx_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.exception(e)
