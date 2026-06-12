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


# =============================================================
# DOCX LOW-LEVEL HELPERS
# =============================================================

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def safe_filename(s: str) -> str:
    s = re.sub(r'[\\/*?:"<>|]+', "_", s)
    return s.strip().replace(" ", "_")


def iter_all_paragraphs(container):
    """Yield every paragraph in body/header/footer/table cells."""
    for p in container.paragraphs:
        yield p
    for t in container.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from iter_all_paragraphs(cell)


def paragraph_full_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def set_paragraph_text(paragraph, text: str):
    """Replace paragraph content keeping first run's formatting."""
    for i, run in enumerate(paragraph.runs):
        run.text = text if i == 0 else ""
    if not paragraph.runs:
        paragraph.add_run(text)


def replace_paragraph_text(paragraph, old: str, new: str):
    """
    Replace `old` with `new` inside a paragraph.
    Handles the case where Word splits text across multiple runs.
    """
    full = paragraph_full_text(paragraph)
    if old not in full:
        return False
    new_full = full.replace(old, new)
    set_paragraph_text(paragraph, new_full)
    return True


def replace_text_in_container(container, replacements: dict):
    """Apply multiple replacements to all paragraphs in a container."""
    for p in iter_all_paragraphs(container):
        full = paragraph_full_text(p)
        new_full = full
        for old, new in replacements.items():
            if old and old in new_full:
                new_full = new_full.replace(old, new)
        if new_full != full:
            set_paragraph_text(p, new_full)


def replace_everywhere(doc, replacements: dict):
    """Replace in body + all section headers/footers."""
    replace_text_in_container(doc, replacements)
    for section in doc.sections:
        replace_text_in_container(section.header, replacements)
        replace_text_in_container(section.footer, replacements)


def find_paragraph_containing(container, needles: list,
                               case_insensitive=True):
    """Return first paragraph containing any of the needles."""
    for p in iter_all_paragraphs(container):
        txt = paragraph_full_text(p)
        chk = txt.lower() if case_insensitive else txt
        for needle in needles:
            nd = needle.lower() if case_insensitive else needle
            if nd in chk:
                return p
    return None


def find_in_doc(doc, needles: list):
    """Search body then headers/footers. Return (paragraph, location_str)."""
    p = find_paragraph_containing(doc, needles)
    if p:
        return p, "body"
    for i, section in enumerate(doc.sections):
        p = find_paragraph_containing(section.header, needles)
        if p:
            return p, f"header_{i}"
        p = find_paragraph_containing(section.footer, needles)
        if p:
            return p, f"footer_{i}"
    return None, None


def insert_paragraph_after(ref_paragraph, text=""):
    """Insert a new paragraph immediately after ref_paragraph."""
    new_para = ref_paragraph._parent.add_paragraph()
    ref_paragraph._p.addnext(new_para._p)
    if text:
        new_para.add_run(text)
    return new_para


def replace_placeholder_with_image(doc, placeholder: str,
                                    image_bytes: bytes,
                                    width=Inches(5.0)):
    """
    Find the paragraph that contains `placeholder`,
    clear it, and insert the image in-place.
    Returns True if placeholder was found and replaced.
    """
    for p in iter_all_paragraphs(doc):
        if placeholder in paragraph_full_text(p):
            # Clear all runs
            for run in p.runs:
                run.text = ""
            # Insert picture in first run
            if p.runs:
                p.runs[0].add_picture(io.BytesIO(image_bytes), width=width)
            else:
                p.add_run().add_picture(io.BytesIO(image_bytes), width=width)
            return True
    return False


def add_hyperlink(paragraph, text: str, url: str,
                  color="0000FF", underline=True):
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
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


# =============================================================
# BUSINESS / MOP LOGIC
# =============================================================

def build_olt_label(equipment: str, custom_olt_label: str) -> str:
    if normalize_spaces(equipment).lower() == "nokia lightspan mf-2":
        return "OLT MF-2"
    return normalize_spaces(custom_olt_label) or normalize_spaces(equipment)


def build_fuse_line(fuse_no: str, olt_label: str, equipment: str) -> str:
    """
    e.g. "FUSE No: L3 OLT MF-2 – (Nokia Lightspan MF-2 Power tapping point)"
    """
    return (
        f"FUSE No: {fuse_no} {olt_label} "
        f"– ({normalize_spaces(equipment)} Power tapping point)"
    )


def process_rs_section(doc, data: dict, rs_entries: list):
    """
    Handles all RS-related replacements using placeholders + existing text.

    Template structure expected:

    --- RS LOAD SCHEDULE SECTION (page 2 area) ---
    PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: (RECTIFIER 1)
    FUSE No: {{load}}+ Equipment –(Nokia Power tapping point)   ← RS1 FUSE line
    {{RS1 Load schedule image}}                                  ← RS1 fuse image placeholder

    PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: (RECTIFIER 2)
    FUSE No: L6 Nokia OLT MF-2 – ( Nokia Power tapping point)  ← RS2 FUSE line
    {{RS2 Load schedule image}}                                  ← RS2 fuse image placeholder

    --- EXISTING RECTIFIER SECTION (page 8 area) ---
    {{RS1 EXISTING IMAGE}}                                       ← RS1 existing rectifier photo
    RECTIFIER 1                                                  ← RS1 label
    {{RS2 EXISTING IMAGE}}                                       ← RS2 existing rectifier photo
    RECTIFIER 2                                                  ← RS2 label
    """

    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])
    equipment  = data["equipment"]

    rs1 = rs_entries[0] if len(rs_entries) > 0 else None
    rs2 = rs_entries[1] if len(rs_entries) > 1 else None

    # ----------------------------------------------------------
    # 1.  "PROPOSED RECTIFIER SYSTEM..." header lines
    #     Update RECTIFIER number labels.
    #     If only 1 RS, clear the RECTIFIER 2 header.
    # ----------------------------------------------------------
    rectifier_header_paras = []
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p)
        if "PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE" in txt.upper():
            rectifier_header_paras.append(p)

    if len(rectifier_header_paras) >= 1 and rs1:
        new_header1 = (
            f"PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: "
            f"(RECTIFIER 1 – {rs1['name']})"
        )
        set_paragraph_text(rectifier_header_paras[0], new_header1)

    if len(rectifier_header_paras) >= 2:
        if rs2:
            new_header2 = (
                f"PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: "
                f"(RECTIFIER 2 – {rs2['name']})"
            )
            set_paragraph_text(rectifier_header_paras[1], new_header2)
        else:
            # 1 RS only → clear the second header
            set_paragraph_text(rectifier_header_paras[1], "")

    # ----------------------------------------------------------
    # 2.  RS1 FUSE line
    #     Template has: "FUSE No: {{load}}+ Equipment –..."
    #     Replace with the real FUSE line.
    #     Fuse number comes from rs1["fuse_no"]
    # ----------------------------------------------------------
    rs1_fuse_needles = [
        "{{load}}",            # our new placeholder
        "FUSE No: L3",         # fallback if template still has original text
        "FUSE No: L3 Nokia OLT MF-2",
        "FUSE No: L3 OLT MF-2",
    ]
    rs1_fuse_para, _ = find_in_doc(doc, rs1_fuse_needles)
    if rs1_fuse_para and rs1:
        set_paragraph_text(
            rs1_fuse_para,
            build_fuse_line(rs1["fuse_no"], olt_label, equipment)
        )

    # ----------------------------------------------------------
    # 3.  RS2 FUSE line
    #     Template has: "FUSE No: L6 Nokia OLT MF-2 – ..."
    # ----------------------------------------------------------
    rs2_fuse_needles = [
        "FUSE No: L6 Nokia OLT MF-2",
        "FUSE No: L6 OLT MF-2",
        "FUSE No: L6",
    ]
    rs2_fuse_para, _ = find_in_doc(doc, rs2_fuse_needles)
    if rs2_fuse_para:
        if rs2:
            set_paragraph_text(
                rs2_fuse_para,
                build_fuse_line(rs2["fuse_no"], olt_label, equipment)
            )
        else:
            set_paragraph_text(rs2_fuse_para, "")

    # ----------------------------------------------------------
    # 4.  RS1 Load schedule image  →  {{RS1 Load schedule image}}
    # ----------------------------------------------------------
    if rs1 and rs1.get("load_img_bytes"):
        found = replace_placeholder_with_image(
            doc,
            "{{RS1 Load schedule image}}",
            rs1["load_img_bytes"],
            width=Inches(5.0),
        )
        if not found:
            st.warning("Placeholder '{{RS1 Load schedule image}}' not found in template.")
    else:
        # Clear the placeholder text so it doesn't appear in output
        replace_everywhere(doc, {"{{RS1 Load schedule image}}": ""})

    # ----------------------------------------------------------
    # 5.  RS2 Load schedule image  →  {{RS2 Load schedule image}}
    # ----------------------------------------------------------
    if rs2 and rs2.get("load_img_bytes"):
        found = replace_placeholder_with_image(
            doc,
            "{{RS2 Load schedule image}}",
            rs2["load_img_bytes"],
            width=Inches(5.0),
        )
        if not found:
            st.warning("Placeholder '{{RS2 Load schedule image}}' not found in template.")
    else:
        replace_everywhere(doc, {"{{RS2 Load schedule image}}": ""})

    # ----------------------------------------------------------
    # 6.  RS1 EXISTING IMAGE  →  {{RS1 EXISTING IMAGE}}
    #     Caption: "RECTIFIER 1" already in template (keep it)
    # ----------------------------------------------------------
    if rs1 and rs1.get("existing_img_bytes"):
        found = replace_placeholder_with_image(
            doc,
            "{{RS1 EXISTING IMAGE}}",
            rs1["existing_img_bytes"],
            width=Inches(5.0),
        )
        if not found:
            st.warning("Placeholder '{{RS1 EXISTING IMAGE}}' not found in template.")
    else:
        replace_everywhere(doc, {"{{RS1 EXISTING IMAGE}}": ""})

    # Update RECTIFIER 1 caption to include name
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p).strip()
        if txt == "RECTIFIER 1":
            if rs1:
                set_paragraph_text(p, f"RECTIFIER 1 – {rs1['name']}")
            break

    # ----------------------------------------------------------
    # 7.  RS2 EXISTING IMAGE  →  {{RS2 EXISTING IMAGE}}
    #     Caption: "RECTIFIER 2" already in template (keep it or clear)
    # ----------------------------------------------------------
    if rs2 and rs2.get("existing_img_bytes"):
        found = replace_placeholder_with_image(
            doc,
            "{{RS2 EXISTING IMAGE}}",
            rs2["existing_img_bytes"],
            width=Inches(5.0),
        )
        if not found:
            st.warning("Placeholder '{{RS2 EXISTING IMAGE}}' not found in template.")
    else:
        replace_everywhere(doc, {"{{RS2 EXISTING IMAGE}}": ""})

    # Update or clear RECTIFIER 2 caption
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p).strip()
        if txt == "RECTIFIER 2":
            if rs2:
                set_paragraph_text(p, f"RECTIFIER 2 – {rs2['name']}")
            else:
                set_paragraph_text(p, "")
            break


def insert_supporting_documents(doc, tssr_pdf_name: str):
    if not tssr_pdf_name:
        return

    anchor, _ = find_in_doc(
        doc,
        ["Supporting Documents", "Supporting Document", "Attachments"]
    )
    if not anchor:
        anchor = doc.paragraphs[-1]

    p = anchor._parent.add_paragraph()
    anchor._p.addnext(p._p)
    p.add_run("TSSR: ")
    p.add_run(tssr_pdf_name)


def generate_docx_bytes(data: dict, rs_entries: list, tssr_pdf_name: str):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"Template.docx not found. Make sure it is committed to the repo."
        )

    doc = Document(TEMPLATE_FILE)

    # ----------------------------------------------------------
    # Global text replacements
    # (site name, plaid, equipment, prepared by, position, date)
    # ----------------------------------------------------------
    replacements = {
        # Original template text  →  form value
        "CDO-604":                          data["site_name"],
        "MIN699":                           data["plaid"],
        "Nokia Lightspan MF-2":             data["equipment"],
        "John Carlo Rabanes":               data["prepared_by"],
        "OLT Rollout Engineer":             data["position"],
        "< May 19- June 19, 2026 10:00AM-6:00PM>": data["target_datetime"],
        "May 19- June 19, 2026 10:00AM-6:00PM":    data["target_datetime"],

        # In case user added {{...}} placeholders for these too
        "{{SITE_NAME}}":      data["site_name"],
        "{{PLAID}}":          data["plaid"],
        "{{EQUIPMENT}}":      data["equipment"],
        "{{PREPARED_BY}}":    data["prepared_by"],
        "{{POSITION}}":       data["position"],
        "{{TARGET_DATETIME}}": data["target_datetime"],
    }
    replace_everywhere(doc, replacements)

    # ----------------------------------------------------------
    # RS section (FUSE lines, images, captions)
    # ----------------------------------------------------------
    process_rs_section(doc, data, rs_entries)

    # ----------------------------------------------------------
    # Supporting documents
    # ----------------------------------------------------------
    insert_supporting_documents(doc, tssr_pdf_name)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# =============================================================
# CLIPBOARD HELPER
# =============================================================

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
    result = streamlit_js_eval(js_expressions=js, key=eval_key, want_output=True)
    if not result:
        return None
    if isinstance(result, str) and result.startswith("ERROR:"):
        st.warning(f"Clipboard paste failed: {result}")
        return None
    if isinstance(result, str) and result.startswith("data:image/") and "," in result:
        return base64.b64decode(result.split(",", 1)[1])
    return None


def uploaded_to_bytes(uploaded):
    return uploaded.getvalue() if uploaded else None


def image_input_widget(label_upload: str, label_paste: str,
                       upload_key: str, paste_btn_key: str,
                       session_key: str, eval_key: str,
                       preview_caption: str = ""):
    """
    Reusable widget: file upload + clipboard paste button + preview.
    Returns bytes or None.
    """
    uploaded = st.file_uploader(
        label_upload,
        type=["png", "jpg", "jpeg", "bmp"],
        key=upload_key,
    )
    if st.button(label_paste, key=paste_btn_key):
        st.session_state[session_key] = clipboard_image_to_bytes(eval_key)

    img_bytes = uploaded_to_bytes(uploaded) or st.session_state.get(session_key)
    if img_bytes:
        st.image(img_bytes, caption=preview_caption, width=340)
    return img_bytes


# =============================================================
# STREAMLIT UI
# =============================================================

st.set_page_config(page_title="MOP Automation", layout="wide")
st.title("MOP Automation")

if not os.path.exists(TEMPLATE_FILE):
    st.error(
        "**Template.docx not found.**  "
        "Make sure `Template.docx` is committed to the root of your repository."
    )
    st.stop()

# ── General Information ──────────────────────────────────────
st.subheader("General Information")
g1, g2 = st.columns(2)

with g1:
    site_name        = st.text_input("Site Name",     placeholder="e.g. CDO-707-HS")
    plaid            = st.text_input("Plaid",          placeholder="e.g. MIN1338")
    equipment        = st.text_input("Equipment",      value="Nokia Lightspan MF-2")
    olt_label_custom = st.text_input(
        "Custom OLT Label (leave blank if Nokia Lightspan MF-2)",
        placeholder="e.g. OLT MF-4",
    )

with g2:
    prepared_by     = st.text_input("Prepared By",  placeholder="e.g. John Carlo Rabanes")
    position        = st.text_input("Position",      value="OLT Rollout Engineer")
    target_datetime = st.text_input(
        "Target Date and Time",
        value="May 19- June 19, 2026 10:00AM-6:00PM",
    )
    rs_count = st.selectbox("Number of RS (Rectifier Systems)", options=[1, 2], index=1)

st.divider()

# ── RS Details ───────────────────────────────────────────────
st.subheader("RS Details")

rs_entries = []

for rs_idx in range(1, rs_count + 1):
    with st.expander(f"RS{rs_idx} Details", expanded=True):
        fa, fb = st.columns(2)

        with fa:
            st.markdown(f"#### RS{rs_idx} Information")
            rs_name = st.text_input(
                f"RS{rs_idx} Rectifier Name",
                placeholder="e.g. Eltek Flatpack 1",
                key=f"rs{rs_idx}_name",
            )
            rs_load = st.text_input(
                f"RS{rs_idx} Load Assignment",
                placeholder="e.g. F8",
                key=f"rs{rs_idx}_load",
            )
            rs_fuse_no = st.text_input(
                f"RS{rs_idx} FUSE No",
                placeholder="e.g. L3",
                value="L3" if rs_idx == 1 else "L6",
                key=f"rs{rs_idx}_fuse_no",
            )

        with fb:
            st.markdown(f"#### RS{rs_idx} Load Schedule Image")
            st.caption("This is the FUSE panel / load schedule photo (page 2 area)")
            load_img = image_input_widget(
                label_upload   = f"Upload RS{rs_idx} Load Schedule image",
                label_paste    = f"Paste RS{rs_idx} Load Schedule from clipboard",
                upload_key     = f"rs{rs_idx}_load_img_upload",
                paste_btn_key  = f"paste_rs{rs_idx}_load_btn",
                session_key    = f"rs{rs_idx}_load_img_bytes",
                eval_key       = f"rs{rs_idx}_load_clip_eval",
                preview_caption= f"RS{rs_idx} Load Schedule",
            )

        st.markdown(f"#### RS{rs_idx} Existing Rectifier Photo (page 8 area)")
        st.caption("This is the photo of the physical existing rectifier")
        existing_img = image_input_widget(
            label_upload   = f"Upload RS{rs_idx} Existing Rectifier image",
            label_paste    = f"Paste RS{rs_idx} Existing Rectifier from clipboard",
            upload_key     = f"rs{rs_idx}_existing_img_upload",
            paste_btn_key  = f"paste_rs{rs_idx}_existing_btn",
            session_key    = f"rs{rs_idx}_existing_img_bytes",
            eval_key       = f"rs{rs_idx}_existing_clip_eval",
            preview_caption= f"RS{rs_idx} Existing Rectifier",
        )

        rs_entries.append({
            "name":              rs_name.strip(),
            "load":              rs_load.strip(),
            "fuse_no":           rs_fuse_no.strip(),
            "load_img_bytes":    load_img,
            "existing_img_bytes": existing_img,
        })

st.divider()

# ── Supporting Documents ─────────────────────────────────────
st.subheader("Supporting Documents")
tssr_pdf = st.file_uploader(
    "TSSR PDF (filename will be written into the Word document)",
    type=["pdf"],
    key="tssr_pdf",
)
if tssr_pdf:
    st.success(f"PDF ready: {tssr_pdf.name}")

st.divider()

# ── Generate ─────────────────────────────────────────────────
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
    "prepared_by", "position", "target_datetime",
]

if st.button("Generate MOP (.docx)", type="primary"):
    # Validate base fields
    missing = [k for k in required_base if not data.get(k)]
    if missing:
        st.error("Please fill in: " + ", ".join(missing))
        st.stop()

    # Validate RS fields
    rs_valid = True
    for i, rs in enumerate(rs_entries, start=1):
        if not rs.get("name"):
            st.error(f"RS{i} Rectifier Name is required.")
            rs_valid = False
        if not rs.get("load"):
            st.error(f"RS{i} Load Assignment is required.")
            rs_valid = False
        if not rs.get("fuse_no"):
            st.error(f"RS{i} FUSE No is required.")
            rs_valid = False
    if not rs_valid:
        st.stop()

    try:
        docx_bytes = generate_docx_bytes(
            data=data,
            rs_entries=rs_entries,
            tssr_pdf_name=(tssr_pdf.name if tssr_pdf else ""),
        )

        out_name = (
            f"MOP_{safe_filename(data['site_name'])}"
            f"_{safe_filename(data['plaid'])}.docx"
        )

        st.success("MOP generated successfully!")
        st.download_button(
            label="⬇️ Download MOP",
            data=docx_bytes,
            file_name=out_name,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
    except Exception as e:
        st.exception(e)