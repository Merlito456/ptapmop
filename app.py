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
    # preserve the first run's formatting if possible
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = text
        else:
            run.text = ""
    if not paragraph.runs:
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


def find_all_paragraphs_by_contains(container, needles, case_insensitive=True):
    """Find ALL paragraphs containing any needle."""
    results = []
    for p in iter_all_paragraphs(container):
        txt = paragraph_full_text(p)
        check = txt.lower() if case_insensitive else txt
        for needle in needles:
            nd = needle.lower() if case_insensitive else needle
            if nd in check:
                results.append(p)
                break
    return results


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


def insert_paragraph_after(paragraph, text=""):
    new_para = paragraph._parent.add_paragraph()
    paragraph._p.addnext(new_para._p)
    if text:
        new_para.add_run(text)
    return new_para


def insert_image_after(paragraph, image_bytes, width=Inches(4.5)):
    p_img = insert_paragraph_after(paragraph)
    run = p_img.add_run()
    run.add_picture(io.BytesIO(image_bytes), width=width)
    return p_img


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


def build_fuse_line_rs1(olt_label: str, equipment: str, fuse_no: str = "L3") -> str:
    return f"FUSE No: {fuse_no} {olt_label} – ({normalize_spaces(equipment)} Power tapping point)"


def build_fuse_line_rs2(olt_label: str, equipment: str, fuse_no: str = "L6") -> str:
    return f"FUSE No: {fuse_no} {olt_label} – ({normalize_spaces(equipment)} Power tapping point)"


def handle_rs_section(doc, data, rs_entries):
    """
    Strategy:
    - Find existing RS label lines like 'RS1 + ...' and 'RS2 + ...' and REPLACE them.
    - Find existing FUSE lines and REPLACE them.
    - Find existing 'RS1' / 'RS2' standalone captions and REPLACE them.
    - Find existing images under RS sections and REPLACE (remove old + insert new).

    Based on the template screenshot:
    - Line 1: "FUSE No: L3 OLT MF-2 – (Nokia Lightspan MF-2 Power tapping point)"  → replace
    - Line 2: "RS1 + Eltek Flatpack + F8"                                           → replace
    - [TABLE with fuse panel]
    - Line: "RS1"                                                                   → replace caption
    - [Section for RS2 if exists]
    - "PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: (RECTIFIER 1)"                 → update number
    - "FUSE No: L6 Nokia OLT MF-2 – (Nokia Power tapping point)"                   → replace
    - [TABLE with rectifier breaker]
    """
    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])
    equipment = data["equipment"]

    # -------------------------------------------------------
    # RS1 FUSE line replacement
    # -------------------------------------------------------
    rs1_fuse_needles = [
        "FUSE No: L3 Nokia OLT MF-2",
        "FUSE No: L3 OLT MF-2",
        "FUSE No: L3",
    ]
    rs1_fuse_para = find_paragraph_by_contains(doc, rs1_fuse_needles)
    if rs1_fuse_para:
        new_fuse1 = build_fuse_line_rs1(olt_label, equipment, "L3")
        set_paragraph_text(rs1_fuse_para, new_fuse1)

    # -------------------------------------------------------
    # RS1 label line: "RS1 + <rectifier> + <load>"
    # -------------------------------------------------------
    rs1_label_needles = ["RS1 +", "rs1 +"]
    rs1_label_para = find_paragraph_by_contains(doc, rs1_label_needles)
    if rs1_label_para and rs_entries:
        rs1 = rs_entries[0]
        new_rs1_label = f"RS1 + {rs1['name']} + {rs1['load']}"
        set_paragraph_text(rs1_label_para, new_rs1_label)

    # -------------------------------------------------------
    # RS1 standalone caption (italic "RS1" below the image)
    # Find paragraphs that are EXACTLY "RS1" (or close)
    # -------------------------------------------------------
    rs1_caption_para = None
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p).strip()
        if txt == "RS1":
            rs1_caption_para = p
            break

    if rs1_caption_para and rs_entries:
        rs1 = rs_entries[0]
        new_caption = f"RS1 – {rs1['name']} (Load: {rs1['load']})"
        set_paragraph_text(rs1_caption_para, new_caption)

    # -------------------------------------------------------
    # RS2 FUSE line replacement
    # -------------------------------------------------------
    rs2_fuse_needles = [
        "FUSE No: L6 Nokia OLT MF-2",
        "FUSE No: L6 OLT MF-2",
        "FUSE No: L6",
    ]
    rs2_fuse_para = find_paragraph_by_contains(doc, rs2_fuse_needles)
    if rs2_fuse_para:
        if len(rs_entries) >= 2:
            new_fuse2 = build_fuse_line_rs2(olt_label, equipment, "L6")
            set_paragraph_text(rs2_fuse_para, new_fuse2)
        else:
            # Only 1 RS — clear the RS2 FUSE line
            set_paragraph_text(rs2_fuse_para, "")

    # -------------------------------------------------------
    # RS2 label line: "RS2 + ..." or standalone "RS2"
    # -------------------------------------------------------
    rs2_label_needles = ["RS2 +", "rs2 +"]
    rs2_label_para = find_paragraph_by_contains(doc, rs2_label_needles)
    if rs2_label_para:
        if len(rs_entries) >= 2:
            rs2 = rs_entries[1]
            new_rs2_label = f"RS2 + {rs2['name']} + {rs2['load']}"
            set_paragraph_text(rs2_label_para, new_rs2_label)
        else:
            set_paragraph_text(rs2_label_para, "")

    rs2_caption_para = None
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p).strip()
        if txt == "RS2":
            rs2_caption_para = p
            break
    if rs2_caption_para:
        if len(rs_entries) >= 2:
            rs2 = rs_entries[1]
            new_caption2 = f"RS2 – {rs2['name']} (Load: {rs2['load']})"
            set_paragraph_text(rs2_caption_para, new_caption2)
        else:
            set_paragraph_text(rs2_caption_para, "")

    # -------------------------------------------------------
    # "PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: (RECTIFIER 1)"
    # Update rectifier number if needed
    # -------------------------------------------------------
    rectifier_label_needles = ["PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE"]
    rect_label_para = find_paragraph_by_contains(doc, rectifier_label_needles)
    if rect_label_para:
        if len(rs_entries) >= 2:
            rs2 = rs_entries[1]
            new_rect_label = (
                f"PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: "
                f"(RECTIFIER 1 – {rs2['name']})"
            )
            set_paragraph_text(rect_label_para, new_rect_label)
        else:
            # 1 RS only — hide or clear this line
            set_paragraph_text(rect_label_para, "")

    # -------------------------------------------------------
    # RS Images: Insert AFTER the RS label lines
    # Images go under the RS1 label and RS2 FUSE line
    # -------------------------------------------------------
    if rs_entries and rs_entries[0].get("img_bytes"):
        target = rs1_label_para or rs1_fuse_para
        if target:
            insert_image_after(target, rs_entries[0]["img_bytes"], width=Inches(4.5))

    if len(rs_entries) >= 2 and rs_entries[1].get("img_bytes"):
        target = rs2_label_para or rs2_fuse_para
        if target:
            insert_image_after(target, rs_entries[1]["img_bytes"], width=Inches(4.5))


def insert_supporting_documents(doc, tssr_pdf_name: str):
    if not tssr_pdf_name:
        return

    anchor, _ = find_any_anchor(
        doc,
        ["Supporting Documents", "Supporting Document", "Attachments"]
    )
    if not anchor:
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph("")

    insert_paragraph_after(anchor, f"TSSR: {tssr_pdf_name}")


def insert_existing_rectifier_image(doc, rectifier_img_bytes):
    if not rectifier_img_bytes:
        return "No rectifier image provided."

    rect_anchors = [
        "Existing Rectifier",
        "Rectifier",
        "Rectifier Photo",
        "Photo of Rectifier",
        "Existing DC Plant",
        "DC Plant",
    ]

    anchor, where = find_any_anchor(doc, rect_anchors)
    if not anchor:
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph("")
        insert_paragraph_after(anchor, "Existing Rectifier:")
        insert_image_after(anchor, rectifier_img_bytes, width=Inches(5.0))
        return "Rectifier anchor not found; image inserted at end of document."

    insert_image_after(anchor, rectifier_img_bytes, width=Inches(5.0))
    return f"Rectifier image inserted under '{paragraph_full_text(anchor)}' in {where}."


def generate_docx_bytes(data, rs_entries, rectifier_img_bytes, tssr_pdf_name):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

    doc = Document(TEMPLATE_FILE)

    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])

    # Global text replacements
    replacements = {
        # Site / plaid / equipment
        "CDO-604": data["site_name"],
        "MIN699": data["plaid"],

        # Equipment name (keep this BEFORE OLT label replacements)
        "Nokia Lightspan MF-2": data["equipment"],

        # Prepared by / position
        "John Carlo Rabanes": data["prepared_by"],
        "OLT Rollout Engineer": data["position"],

        # Target date
        "< May 19- June 19, 2026 10:00AM-6:00PM>": data["target_datetime"],
        "May 19- June 19, 2026 10:00AM-6:00PM": data["target_datetime"],
    }
    replace_everywhere(doc, replacements)

    # RS section (replace labels, captions, FUSE lines, insert images)
    handle_rs_section(doc, data, rs_entries)

    # Supporting documents
    insert_supporting_documents(doc, tssr_pdf_name)

    # Page 8 existing rectifier image
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
        st.warning(f"Clipboard paste failed: {data_url}")
        return None
    if isinstance(data_url, str) and data_url.startswith("data:image/") and "," in data_url:
        b64 = data_url.split(",", 1)[1]
        return base64.b64decode(b64)
    return None


def uploaded_file_to_bytes(uploaded):
    return uploaded.getvalue() if uploaded else None


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="MOP Automation", layout="wide")
st.title("MOP Automation")

if not os.path.exists(TEMPLATE_FILE):
    st.error(
        "Template.docx not found. Make sure Template.docx is committed to the repo root."
    )
    st.stop()

# ---- General Info ----
st.subheader("General Information")
col1, col2 = st.columns(2)

with col1:
    site_name       = st.text_input("Site Name", placeholder="e.g. CDO-707-HS")
    plaid           = st.text_input("Plaid", placeholder="e.g. MIN1338")
    equipment       = st.text_input("Equipment", value="Nokia Lightspan MF-2")
    olt_label_custom = st.text_input(
        "Custom OLT Label (leave blank if Nokia Lightspan MF-2)",
        placeholder="e.g. OLT MF-4",
    )

with col2:
    prepared_by     = st.text_input("Prepared By", placeholder="e.g. John Carlo Rabanes")
    position        = st.text_input("Position", value="OLT Rollout Engineer")
    target_datetime = st.text_input(
        "Target Date and Time",
        value="May 19- June 19, 2026 10:00AM-6:00PM",
    )
    rs_count        = st.selectbox("Number of RS", options=[1, 2], index=1)

st.divider()

# ---- RS Details ----
st.subheader("RS Details")

rs_entries = []

# RS1
with st.expander("RS1 Details", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        rs1_name = st.text_input("RS1 Rectifier Name", placeholder="e.g. Eltek Flatpack 1")
        rs1_load = st.text_input("RS1 Load Assignment", placeholder="e.g. F8")
    with c2:
        rs1_upload = st.file_uploader(
            "Upload RS1 image",
            type=["png", "jpg", "jpeg", "bmp"],
            key="rs1_upload",
        )
        if st.button("Paste RS1 from clipboard", key="paste_rs1_btn"):
            st.session_state["rs1_clip_bytes"] = clipboard_image_to_bytes("rs1_clip_eval")

    rs1_img_bytes = uploaded_file_to_bytes(rs1_upload) or st.session_state.get("rs1_clip_bytes")
    if rs1_img_bytes:
        st.image(rs1_img_bytes, caption=f"RS1 – {rs1_name} (Load: {rs1_load})", width=380)

rs_entries.append({
    "name": rs1_name.strip(),
    "load": rs1_load.strip(),
    "img_bytes": rs1_img_bytes,
})

# RS2 (conditional)
if rs_count == 2:
    with st.expander("RS2 Details", expanded=True):
        c1, c2 = st.columns(2)
        with c2:
            rs2_upload = st.file_uploader(
                "Upload RS2 image",
                type=["png", "jpg", "jpeg", "bmp"],
                key="rs2_upload",
            )
            if st.button("Paste RS2 from clipboard", key="paste_rs2_btn"):
                st.session_state["rs2_clip_bytes"] = clipboard_image_to_bytes("rs2_clip_eval")
        with c1:
            rs2_name = st.text_input("RS2 Rectifier Name", placeholder="e.g. Eltek Flatpack 2")
            rs2_load = st.text_input("RS2 Load Assignment", placeholder="e.g. L3")

        rs2_img_bytes = uploaded_file_to_bytes(rs2_upload) or st.session_state.get("rs2_clip_bytes")
        if rs2_img_bytes:
            st.image(rs2_img_bytes, caption=f"RS2 – {rs2_name} (Load: {rs2_load})", width=380)

    rs_entries.append({
        "name": rs2_name.strip(),
        "load": rs2_load.strip(),
        "img_bytes": rs2_img_bytes,
    })

st.divider()

# ---- Existing Rectifier ----
st.subheader("Existing Rectifier Photo (Page 8)")
c1, c2 = st.columns(2)
with c1:
    rect_upload = st.file_uploader(
        "Upload existing rectifier image",
        type=["png", "jpg", "jpeg", "bmp"],
        key="rect_upload",
    )
    if st.button("Paste Rectifier from clipboard", key="paste_rect_btn"):
        st.session_state["rect_clip_bytes"] = clipboard_image_to_bytes("rect_clip_eval")

rect_img_bytes = uploaded_file_to_bytes(rect_upload) or st.session_state.get("rect_clip_bytes")
if rect_img_bytes:
    st.image(rect_img_bytes, caption="Existing Rectifier", width=380)

st.divider()

# ---- TSSR PDF ----
st.subheader("Supporting Documents")
tssr_pdf = st.file_uploader(
    "TSSR PDF (filename will be written into Word)",
    type=["pdf"],
    key="tssr_pdf",
)
if tssr_pdf:
    st.success(f"PDF selected: {tssr_pdf.name}")

st.divider()

# ---- Generate ----
data = {
    "site_name":        site_name.strip(),
    "plaid":            plaid.strip(),
    "equipment":        equipment.strip(),
    "olt_label_custom": olt_label_custom.strip(),
    "prepared_by":      prepared_by.strip(),
    "position":         position.strip(),
    "target_datetime":  target_datetime.strip(),
}

required_base = [
    "site_name", "plaid", "equipment",
    "prepared_by", "position", "target_datetime"
]

if st.button("Generate MOP (.docx)", type="primary"):
    missing = [k for k in required_base if not data.get(k)]
    if missing:
        st.error("Missing fields: " + ", ".join(missing))
        st.stop()

    for i, rs in enumerate(rs_entries, start=1):
        if not rs.get("name"):
            st.error(f"RS{i} Rectifier Name is required.")
            st.stop()
        if not rs.get("load"):
            st.error(f"RS{i} Load Assignment is required.")
            st.stop()

    try:
        docx_bytes, rect_note = generate_docx_bytes(
            data=data,
            rs_entries=rs_entries,
            rectifier_img_bytes=rect_img_bytes,
            tssr_pdf_name=(tssr_pdf.name if tssr_pdf else ""),
        )

        out_name = (
            f"MOP_{safe_filename(data['site_name'])}"
            f"_{safe_filename(data['plaid'])}.docx"
        )

        st.success("MOP generated successfully.")
        st.info(rect_note)
        st.download_button(
            label="Download MOP",
            data=docx_bytes,
            file_name=out_name,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
    except Exception as e:
        st.exception(e)
