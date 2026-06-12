import os
import re
import io
import time
import base64
import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

TEMPLATE_FILE = "Template.docx"


# -----------------------------
# DOCX helpers (same logic)
# -----------------------------
def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


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
    """
    Insert image from bytes after paragraph.
    python-docx add_picture accepts a file-like object.
    """
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


def build_olt_label(equipment: str, custom_olt_label: str) -> str:
    eq = normalize_spaces(equipment).lower()
    if eq == "nokia lightspan mf-2":
        return "OLT MF-2"
    return normalize_spaces(custom_olt_label) or normalize_spaces(equipment)


def build_fuse_line(olt_label: str, equipment: str) -> str:
    return f"FUSE No: L3 {olt_label} – ({normalize_spaces(equipment)} Power tapping point)"


def insert_power_section_content(doc, data, rs1_img_bytes, rs2_img_bytes):
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
    new_fuse = build_fuse_line(olt_label, data["equipment"])

    set_paragraph_text(anchor, new_fuse)

    rs1_line = f"RS1 + {data['rs1_rectifier_name']} + {data['rs1_load_assignment']}"
    rs2_line = f"RS2 + {data['rs2_rectifier_name']} + {data['rs2_load_assignment']}"

    p1 = insert_paragraph_after(anchor, rs1_line)
    last = p1
    if rs1_img_bytes:
        img_p, _ = insert_image_after(p1, rs1_img_bytes, width=Inches(4.5), caption="RS1")
        last = img_p

    p2 = insert_paragraph_after(last, rs2_line)
    if rs2_img_bytes:
        insert_image_after(p2, rs2_img_bytes, width=Inches(4.5), caption="RS2")


def insert_supporting_documents(doc, pdf_filename):
    if not pdf_filename:
        return
    anchor, _ = find_any_anchor(doc, ["Supporting Documents"])
    if not anchor:
        raise ValueError("Could not find 'Supporting Documents' section.")
    p = insert_paragraph_after(anchor)
    p.add_run("TSSR: ")
    # In Streamlit Cloud we can't reliably link to local filesystem for the downloader.
    # So we just put the filename.
    p.add_run(pdf_filename)


def insert_existing_rectifier_image(doc, rectifier_img_bytes):
    if not rectifier_img_bytes:
        return
    anchor, _ = find_any_anchor(doc, ["Existing Rectifier"])
    if not anchor:
        raise ValueError("Could not find 'Existing Rectifier' section.")
    insert_image_after(anchor, rectifier_img_bytes, width=Inches(5.0), caption="Existing Rectifier")


def generate_docx_bytes(data, rs1_img_bytes, rs2_img_bytes, rectifier_img_bytes, tssr_pdf_name):
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

    insert_power_section_content(doc, data, rs1_img_bytes, rs2_img_bytes)
    insert_supporting_documents(doc, tssr_pdf_name)
    insert_existing_rectifier_image(doc, rectifier_img_bytes)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# -----------------------------
# Streamlit "paste image" via JS
# -----------------------------
def paste_image_widget(key: str, label: str):
    """
    Creates a paste button. When user pastes, it stores base64 PNG in st.session_state[key].
    Requires user interaction and browser permissions.
    """
    import streamlit.components.v1 as components

    html = f"""
    <div style="display:flex; gap:10px; align-items:center;">
      <button id="btn_{key}" type="button">{label}</button>
      <span id="status_{key}" style="font-family:Arial; font-size:12px; color:#444;"></span>
    </div>
    <script>
      const btn = document.getElementById("btn_{key}");
      const status = document.getElementById("status_{key}");

      async function readClipboardImage() {{
        try {{
          const items = await navigator.clipboard.read();
          for (const item of items) {{
            for (const type of item.types) {{
              if (type.startsWith("image/")) {{
                const blob = await item.getType(type);
                const reader = new FileReader();
                reader.onload = () => {{
                  const dataUrl = reader.result; // data:image/png;base64,...
                  const payload = {{
                    key: "{key}",
                    dataUrl: dataUrl
                  }};
                  window.parent.postMessage({{ isStreamlitMessage: true, type: "streamlit:setComponentValue", value: payload }}, "*");
                  status.textContent = "Image pasted.";
                }};
                reader.readAsDataURL(blob);
                return;
              }}
            }}
          }}
          status.textContent = "No image in clipboard.";
        }} catch (e) {{
          status.textContent = "Paste blocked by browser permissions.";
        }}
      }}

      btn.addEventListener("click", readClipboardImage);
    </script>
    """
    payload = components.html(html, height=50, key=f"paste_component_{key}")
    if payload and isinstance(payload, dict) and payload.get("key") == key:
        st.session_state[key] = payload.get("dataUrl")


def dataurl_to_bytes(data_url: str):
    if not data_url or "," not in data_url:
        return None
    header, b64 = data_url.split(",", 1)
    return base64.b64decode(b64)


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="MOP Automation", layout="wide")
st.title("MOP Automation (Streamlit)")

with st.sidebar:
    st.write("Template file required: `Template.docx` in the repo root.")

col1, col2 = st.columns(2)

with col1:
    site_name = st.text_input("Site Name", value="")
    plaid = st.text_input("Plaid", value="")
    equipment = st.text_input("Equipment", value="Nokia Lightspan MF-2")
    olt_label_custom = st.text_input("Custom OLT Label (if not Nokia MF-2)", value="")
    prepared_by = st.text_input("Prepared By", value="")
    position = st.text_input("Position", value="OLT Rollout Engineer")
    target_datetime = st.text_input("Target Date and Time", value="May 19- June 19, 2026 10:00AM-6:00PM")

with col2:
    rs1_rectifier_name = st.text_input("RS1 Rectifier Name", value="")
    rs1_load_assignment = st.text_input("RS1 Load Assignment", value="")
    rs2_rectifier_name = st.text_input("RS2 Rectifier Name", value="")
    rs2_load_assignment = st.text_input("RS2 Load Assignment", value="")

st.divider()

st.subheader("Images / Attachments")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("### RS1 Image")
    rs1_upload = st.file_uploader("Upload RS1 image", type=["png", "jpg", "jpeg", "bmp"], key="rs1_upload")
    paste_image_widget("rs1_paste_dataurl", "Paste RS1 from clipboard")
    rs1_pasted_bytes = dataurl_to_bytes(st.session_state.get("rs1_paste_dataurl", ""))
    if rs1_upload:
        st.image(rs1_upload, caption="RS1 uploaded", use_container_width=True)
    elif rs1_pasted_bytes:
        st.image(rs1_pasted_bytes, caption="RS1 pasted", use_container_width=True)

with c2:
    st.markdown("### RS2 Image")
    rs2_upload = st.file_uploader("Upload RS2 image", type=["png", "jpg", "jpeg", "bmp"], key="rs2_upload")
    paste_image_widget("rs2_paste_dataurl", "Paste RS2 from clipboard")
    rs2_pasted_bytes = dataurl_to_bytes(st.session_state.get("rs2_paste_dataurl", ""))
    if rs2_upload:
        st.image(rs2_upload, caption="RS2 uploaded", use_container_width=True)
    elif rs2_pasted_bytes:
        st.image(rs2_pasted_bytes, caption="RS2 pasted", use_container_width=True)

with c3:
    st.markdown("### Existing Rectifier (Page 8)")
    rectifier_upload = st.file_uploader("Upload rectifier image", type=["png", "jpg", "jpeg", "bmp"], key="rect_upload")
    paste_image_widget("rectifier_paste_dataurl", "Paste Rectifier from clipboard")
    rect_pasted_bytes = dataurl_to_bytes(st.session_state.get("rectifier_paste_dataurl", ""))
    if rectifier_upload:
        st.image(rectifier_upload, caption="Rectifier uploaded", use_container_width=True)
    elif rect_pasted_bytes:
        st.image(rect_pasted_bytes, caption="Rectifier pasted", use_container_width=True)

tssr_pdf = st.file_uploader("TSSR PDF (uploaded for reference; filename inserted into Word)", type=["pdf"], key="tssr_pdf")

st.divider()

data = {
    "site_name": site_name.strip(),
    "plaid": plaid.strip(),
    "equipment": equipment.strip(),
    "olt_label_custom": olt_label_custom.strip(),
    "prepared_by": prepared_by.strip(),
    "position": position.strip(),
    "target_datetime": target_datetime.strip(),
    "rs1_rectifier_name": rs1_rectifier_name.strip(),
    "rs1_load_assignment": rs1_load_assignment.strip(),
    "rs2_rectifier_name": rs2_rectifier_name.strip(),
    "rs2_load_assignment": rs2_load_assignment.strip(),
}

required = [
    "site_name", "plaid", "equipment", "prepared_by", "position",
    "target_datetime", "rs1_rectifier_name", "rs1_load_assignment",
    "rs2_rectifier_name", "rs2_load_assignment"
]

def uploaded_file_to_bytes(uploaded):
    if not uploaded:
        return None
    return uploaded.getvalue()

rs1_img_bytes = uploaded_file_to_bytes(rs1_upload) or rs1_pasted_bytes
rs2_img_bytes = uploaded_file_to_bytes(rs2_upload) or rs2_pasted_bytes
rectifier_img_bytes = uploaded_file_to_bytes(rectifier_upload) or rect_pasted_bytes
tssr_pdf_name = tssr_pdf.name if tssr_pdf else ""

if st.button("Generate MOP (.docx)", type="primary"):
    missing = [k for k in required if not data.get(k)]
    if missing:
        st.error("Missing fields: " + ", ".join(missing))
    else:
        try:
            docx_bytes = generate_docx_bytes(
                data,
                rs1_img_bytes=rs1_img_bytes,
                rs2_img_bytes=rs2_img_bytes,
                rectifier_img_bytes=rectifier_img_bytes,
                tssr_pdf_name=tssr_pdf_name
            )

            out_name = f"MOP_{safe_filename(data['site_name'])}_{safe_filename(data['plaid'])}.docx"
            st.success("Generated.")
            st.download_button(
                "Download MOP",
                data=docx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.exception(e)
