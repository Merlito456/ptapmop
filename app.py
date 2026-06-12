import os
import re
import io
import base64
from lxml import etree
import streamlit as st

from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from streamlit_js_eval import streamlit_js_eval


TEMPLATE_FILE = "Template.docx"

# XML namespaces used in Word documents
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# =============================================================
# DOCX LOW-LEVEL HELPERS
# =============================================================

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def safe_filename(s: str) -> str:
    s = re.sub(r'[\\/*?:"<>|]+', "_", s)
    return s.strip().replace(" ", "_")


def iter_all_paragraphs(container):
    """Yield all paragraphs in body/table cells recursively."""
    for p in container.paragraphs:
        yield p
    for t in container.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from iter_all_paragraphs(cell)


def paragraph_full_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def set_paragraph_text(paragraph, text: str):
    """Replace paragraph text preserving first run formatting."""
    for i, run in enumerate(paragraph.runs):
        run.text = text if i == 0 else ""
    if not paragraph.runs:
        paragraph.add_run(text)


def replace_text_in_paragraph_preserve_format(paragraph, replacements: dict):
    """
    Replace text in paragraph that may be split across multiple runs.
    Preserves first run formatting (bold, italic, underline, color, font).
    """
    full = paragraph_full_text(paragraph)
    if not full.strip():
        return

    new_full = full
    changed = False
    for old, new in replacements.items():
        if old and old in new_full:
            new_full = new_full.replace(old, new)
            changed = True

    if not changed:
        return

    if paragraph.runs:
        first_run = paragraph.runs[0]
        bold      = first_run.bold
        italic    = first_run.italic
        underline = first_run.underline
        font_name = first_run.font.name
        font_size = first_run.font.size
        try:
            font_color = (
                first_run.font.color.rgb
                if first_run.font.color and first_run.font.color.type
                else None
            )
        except Exception:
            font_color = None

        for run in paragraph.runs:
            run.text = ""

        first_run.text      = new_full
        first_run.bold      = bold
        first_run.italic    = italic
        first_run.underline = underline
        if font_name:
            first_run.font.name = font_name
        if font_size:
            first_run.font.size = font_size
        if font_color:
            try:
                first_run.font.color.rgb = font_color
            except Exception:
                pass
    else:
        paragraph.add_run(new_full)


def replace_text_in_container(container, replacements: dict):
    for p in iter_all_paragraphs(container):
        replace_text_in_paragraph_preserve_format(p, replacements)


# =============================================================
# RAW XML TEXT REPLACEMENT (for headers/footers with textboxes)
# =============================================================

def replace_in_xml_text_nodes(xml_element, replacements: dict):
    """
    Walk ALL <w:t> nodes in the XML tree and do plain string replacement.
    This reaches text inside:
    - text boxes  (<w:txbxContent>)
    - drawings
    - shapes
    - anything python-docx .paragraphs doesn't expose

    Note: this does NOT handle text split across runs.
    For split-run text in headers, we also run the run-reconstruction approach.
    """
    ns = WORD_NS
    for t_node in xml_element.iter(f"{{{ns}}}t"):
        if t_node.text:
            new_text = t_node.text
            for old, new in replacements.items():
                if old and old in new_text:
                    new_text = new_text.replace(old, new)
            if new_text != t_node.text:
                t_node.text = new_text


def replace_everywhere(doc, replacements: dict):
    """
    Replace text in:
    1. Body paragraphs + table cells (run-aware, format-preserving)
    2. Header/footer via python-docx paragraph access
    3. Header/footer raw XML (catches textboxes, drawings, shapes)
    """
    # 1. Body
    replace_text_in_container(doc, replacements)

    # 2 & 3. All sections headers/footers
    for section in doc.sections:
        # python-docx paragraph access (tables, plain paragraphs)
        replace_text_in_container(section.header, replacements)
        replace_text_in_container(section.footer, replacements)

        # Raw XML access (textboxes, drawings, shapes)
        replace_in_xml_text_nodes(section.header._element, replacements)
        replace_in_xml_text_nodes(section.footer._element, replacements)

        # Even/first page headers if present
        try:
            replace_text_in_container(section.even_page_header, replacements)
            replace_in_xml_text_nodes(
                section.even_page_header._element, replacements)
        except Exception:
            pass
        try:
            replace_text_in_container(section.first_page_header, replacements)
            replace_in_xml_text_nodes(
                section.first_page_header._element, replacements)
        except Exception:
            pass


# =============================================================
# FIND PARAGRAPH HELPERS
# =============================================================

def find_paragraph_containing(container, needles: list,
                               case_insensitive=True):
    for p in iter_all_paragraphs(container):
        txt = paragraph_full_text(p)
        chk = txt.lower() if case_insensitive else txt
        for needle in needles:
            nd = needle.lower() if case_insensitive else needle
            if nd in chk:
                return p
    return None


def find_in_doc(doc, needles: list):
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
    new_para = ref_paragraph._parent.add_paragraph()
    ref_paragraph._p.addnext(new_para._p)
    if text:
        new_para.add_run(text)
    return new_para


# =============================================================
# IMAGE PLACEHOLDER REPLACEMENT
# =============================================================

def replace_placeholder_with_image(doc, placeholders: list,
                                    image_bytes: bytes,
                                    width=Inches(5.0)):
    """
    Find paragraph containing any placeholder string,
    clear it, insert image in-place.
    Returns the matched placeholder or None.
    """
    for p in iter_all_paragraphs(doc):
        full = paragraph_full_text(p)
        for ph in placeholders:
            if ph in full:
                for run in p.runs:
                    run.text = ""
                if p.runs:
                    p.runs[0].add_picture(io.BytesIO(image_bytes), width=width)
                else:
                    p.add_run().add_picture(io.BytesIO(image_bytes), width=width)
                return ph
    return None


def clear_placeholders(doc, placeholders: list):
    mapping = {ph: "" for ph in placeholders}
    replace_everywhere(doc, mapping)


# =============================================================
# HYPERLINK HELPER
# =============================================================

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
# DEBUG HELPER
# =============================================================

def debug_list_all_text(doc) -> list:
    """
    List all text from:
    - body paragraphs + table cells
    - header/footer paragraphs + table cells
    - header/footer raw XML <w:t> nodes (catches textboxes)
    """
    lines = []

    # Body
    for p in iter_all_paragraphs(doc):
        txt = paragraph_full_text(p).strip()
        if txt:
            lines.append(f"[BODY] {txt}")

    # Headers / footers via python-docx
    for i, section in enumerate(doc.sections):
        for p in iter_all_paragraphs(section.header):
            txt = paragraph_full_text(p).strip()
            if txt:
                lines.append(f"[HEADER s{i}] {txt}")
        for p in iter_all_paragraphs(section.footer):
            txt = paragraph_full_text(p).strip()
            if txt:
                lines.append(f"[FOOTER s{i}] {txt}")

    # Headers / footers via raw XML (textboxes, shapes)
    ns = WORD_NS
    for i, section in enumerate(doc.sections):
        for t_node in section.header._element.iter(f"{{{ns}}}t"):
            txt = (t_node.text or "").strip()
            if txt:
                lines.append(f"[HEADER-XML s{i}] {txt}")
        for t_node in section.footer._element.iter(f"{{{ns}}}t"):
            txt = (t_node.text or "").strip()
            if txt:
                lines.append(f"[FOOTER-XML s{i}] {txt}")

    return lines


# =============================================================
# PLACEHOLDER VARIANTS
# =============================================================

RS1_LOAD_PH = [
    "{{RS1 Load schedule image}}",
    "{{RS1 Load schedule image}",
    "{{RS1 load schedule image}}",
    "{{RS1 load schedule image}",
]

RS2_LOAD_PH = [
    "{{RS2 Load schedule image}}",
    "{{RS2 Load schedule image}",
    "{{RS2 load schedule image}}",
    "{{RS2 load schedule image}",
]

RS1_EXIST_PH = [
    "{{RS1 EXISTING IMAGE}}",
    "{{RS1 EXISTING IMAGE}",
    "{{RS1 existing image}}",
    "{{RS1 existing image}",
]

RS2_EXIST_PH = [
    "{{RS2 EXISTING IMAGE}}",
    "{{RS2 EXISTING IMAGE}",
    "{{RS2 existing image}}",
    "{{RS2 existing image}",
]


# =============================================================
# BUSINESS LOGIC
# =============================================================

def build_olt_label(equipment: str, custom_olt_label: str) -> str:
    if normalize_spaces(equipment).lower() == "nokia lightspan mf-2":
        return "OLT MF-2"
    return normalize_spaces(custom_olt_label) or normalize_spaces(equipment)


def build_fuse_line(fuse_no: str, olt_label: str, equipment: str) -> str:
    return (
        f"FUSE No: {fuse_no} {olt_label} "
        f"– ({normalize_spaces(equipment)} Power tapping point)"
    )


def process_rs_section(doc, data: dict, rs_entries: list, warnings: list):
    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])
    equipment  = data["equipment"]

    rs1 = rs_entries[0] if len(rs_entries) > 0 else None
    rs2 = rs_entries[1] if len(rs_entries) > 1 else None

    # A. PROPOSED RECTIFIER SYSTEM headers
    rectifier_header_paras = []
    for p in iter_all_paragraphs(doc):
        if "PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE" in paragraph_full_text(p).upper():
            rectifier_header_paras.append(p)

    if len(rectifier_header_paras) >= 1 and rs1:
        set_paragraph_text(
            rectifier_header_paras[0],
            f"PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: "
            f"(RECTIFIER 1 – {rs1['name']})"
        )
    if len(rectifier_header_paras) >= 2:
        if rs2:
            set_paragraph_text(
                rectifier_header_paras[1],
                f"PROPOSED RECTIFIER SYSTEM LOAD BREAKER FUSE: "
                f"(RECTIFIER 2 – {rs2['name']})"
            )
        else:
            set_paragraph_text(rectifier_header_paras[1], "")

    # B. RS1 FUSE line
    rs1_fuse_para, _ = find_in_doc(doc, [
        "{{load}}", "FUSE No: L3 Nokia OLT MF-2",
        "FUSE No: L3 OLT MF-2", "FUSE No: L3",
        "{{load}}+ Equipment",
    ])
    if rs1_fuse_para and rs1:
        set_paragraph_text(
            rs1_fuse_para,
            build_fuse_line(rs1["fuse_no"], olt_label, equipment)
        )

    # C. RS2 FUSE line
    rs2_fuse_para, _ = find_in_doc(doc, [
        "FUSE No: L6 Nokia OLT MF-2",
        "FUSE No: L6 OLT MF-2",
        "FUSE No: L6",
    ])
    if rs2_fuse_para:
        if rs2:
            set_paragraph_text(
                rs2_fuse_para,
                build_fuse_line(rs2["fuse_no"], olt_label, equipment)
            )
        else:
            set_paragraph_text(rs2_fuse_para, "")

    # D. RS1 Load Schedule image
    if rs1 and rs1.get("load_img_bytes"):
        matched = replace_placeholder_with_image(
            doc, RS1_LOAD_PH, rs1["load_img_bytes"], width=Inches(5.0))
        if matched:
            warnings.append(f"✅ RS1 Load Schedule image inserted (matched: '{matched}').")
        else:
            warnings.append(f"⚠️ RS1 Load Schedule placeholder not found. Tried: {RS1_LOAD_PH}")
    else:
        clear_placeholders(doc, RS1_LOAD_PH)

    # E. RS2 Load Schedule image
    if rs2 and rs2.get("load_img_bytes"):
        matched = replace_placeholder_with_image(
            doc, RS2_LOAD_PH, rs2["load_img_bytes"], width=Inches(5.0))
        if matched:
            warnings.append(f"✅ RS2 Load Schedule image inserted (matched: '{matched}').")
        else:
            warnings.append(f"⚠️ RS2 Load Schedule placeholder not found. Tried: {RS2_LOAD_PH}")
    else:
        clear_placeholders(doc, RS2_LOAD_PH)

    # F. RS1 EXISTING IMAGE
    if rs1 and rs1.get("existing_img_bytes"):
        matched = replace_placeholder_with_image(
            doc, RS1_EXIST_PH, rs1["existing_img_bytes"], width=Inches(5.0))
        if matched:
            warnings.append(f"✅ RS1 Existing image inserted (matched: '{matched}').")
        else:
            warnings.append(
                f"⚠️ RS1 EXISTING IMAGE placeholder not found. Tried: {RS1_EXIST_PH}")
    else:
        clear_placeholders(doc, RS1_EXIST_PH)

    # RECTIFIER 1 caption
    for p in iter_all_paragraphs(doc):
        if paragraph_full_text(p).strip() == "RECTIFIER 1":
            set_paragraph_text(p, f"RECTIFIER 1 – {rs1['name']}" if rs1 else "")
            break

    # G. RS2 EXISTING IMAGE
    if rs2 and rs2.get("existing_img_bytes"):
        matched = replace_placeholder_with_image(
            doc, RS2_EXIST_PH, rs2["existing_img_bytes"], width=Inches(5.0))
        if matched:
            warnings.append(f"✅ RS2 Existing image inserted (matched: '{matched}').")
        else:
            warnings.append(
                f"⚠️ RS2 EXISTING IMAGE placeholder not found. Tried: {RS2_EXIST_PH}")
    else:
        clear_placeholders(doc, RS2_EXIST_PH)

    # RECTIFIER 2 caption
    for p in iter_all_paragraphs(doc):
        if paragraph_full_text(p).strip() == "RECTIFIER 2":
            set_paragraph_text(p, f"RECTIFIER 2 – {rs2['name']}" if rs2 else "")
            break


def insert_supporting_documents(doc, tssr_pdf_name: str):
    if not tssr_pdf_name:
        return
    anchor, _ = find_in_doc(
        doc,
        ["Supporting Documents", "Supporting Document",
         "SUPPORTING DOCUMENTS", "Attachments"]
    )
    if not anchor:
        anchor = doc.paragraphs[-1]
    p = anchor._parent.add_paragraph()
    anchor._p.addnext(p._p)
    p.add_run(f"TSSR: {tssr_pdf_name}")


def generate_docx_bytes(data: dict, rs_entries: list,
                         tssr_pdf_name: str):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError("Template.docx not found in repo root.")

    doc = Document(TEMPLATE_FILE)
    warnings = []

    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])

    # ----------------------------------------------------------
    # Build header title replacement
    # From debug output the header text (in textbox) likely is:
    # "Project FTTH: CDO-604_MIN995 Lightspan OLT MF-2 System Power Tapping"
    # We replace piece by piece to preserve surrounding text.
    # ----------------------------------------------------------
    old_header_title = (
        f"Project FTTH: CDO-604_MIN995 Lightspan OLT MF-2 System Power Tapping"
    )
    new_header_title = (
        f"Project FTTH: {data['site_name']}_{data['plaid']} "
        f"Lightspan {data['equipment']} System Power Tapping"
    )

    replacements = {
        # Header title (full match attempt first)
        old_header_title: new_header_title,

        # Individual pieces for robustness
        # (in case text is split across <w:t> nodes)
        "CDO-604": data["site_name"],
        "MIN699":  data["plaid"],
        "MIN995":  data["plaid"],

        # Equipment
        "Nokia Lightspan MF-2": data["equipment"],
        "Lightspan MF-2":       data["equipment"],

        # OLT label
        "OLT MF-2": olt_label,

        # Prepared by
        "John Carlo Rabanes": data["prepared_by"],

        # Position — template uses "OLT Engineer" in body
        # and may use "OLT Rollout Engineer" elsewhere
        "OLT Rollout Engineer": data["position"],
        "OLT Engineer":         data["position"],

        # Target date
        "< May 19- June 19, 2026 10:00AM-6:00PM>": data["target_datetime"],
        "May 19- June 19, 2026 10:00AM-6:00PM":    data["target_datetime"],

        # Generic placeholders
        "{{SITE_NAME}}":       data["site_name"],
        "{{PLAID}}":           data["plaid"],
        "{{EQUIPMENT}}":       data["equipment"],
        "{{PREPARED_BY}}":     data["prepared_by"],
        "{{POSITION}}":        data["position"],
        "{{TARGET_DATETIME}}": data["target_datetime"],
    }

    # replace_everywhere now hits textboxes via raw XML too
    replace_everywhere(doc, replacements)

    process_rs_section(doc, data, rs_entries, warnings)
    insert_supporting_documents(doc, tssr_pdf_name)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue(), warnings


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
    if isinstance(result, str) and "," in result and result.startswith("data:image/"):
        return base64.b64decode(result.split(",", 1)[1])
    return None


def uploaded_to_bytes(uploaded):
    return uploaded.getvalue() if uploaded else None


def image_input_widget(label_upload, label_paste,
                        upload_key, paste_btn_key,
                        session_key, eval_key,
                        preview_caption=""):
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
    st.error("**Template.docx not found** in repo root.")
    st.stop()

# ── Debug ─────────────────────────────────────────────────────
with st.expander("🔍 Debug: Inspect all text in template", expanded=False):
    if st.button("List all text (body + header + footer + XML textboxes)"):
        _doc = Document(TEMPLATE_FILE)
        lines = debug_list_all_text(_doc)
        st.code("\n".join(lines), language="text")
        st.caption(
            "HEADER-XML lines are from textboxes/drawings — "
            "this is what was previously invisible. "
            "Copy the exact text to verify replacements."
        )

st.divider()

# ── General Information ───────────────────────────────────────
st.subheader("General Information")
g1, g2 = st.columns(2)

with g1:
    site_name        = st.text_input("Site Name", placeholder="e.g. CDO-707-HS")
    plaid            = st.text_input("Plaid", placeholder="e.g. MIN1338")
    equipment        = st.text_input("Equipment", value="Nokia Lightspan MF-2")
    olt_label_custom = st.text_input(
        "Custom OLT Label (leave blank if Nokia Lightspan MF-2)",
        placeholder="e.g. OLT MF-4",
    )

with g2:
    prepared_by     = st.text_input("Prepared By",
                                     placeholder="e.g. John Carlo Rabanes")
    position        = st.text_input("Position", value="OLT Rollout Engineer")
    target_datetime = st.text_input(
        "Target Date and Time",
        value="May 19- June 19, 2026 10:00AM-6:00PM",
    )
    rs_count = st.selectbox(
        "Number of RS (Rectifier Systems)", options=[1, 2], index=1)

st.divider()

# ── RS Details ────────────────────────────────────────────────
st.subheader("RS Details")
rs_entries = []

for rs_idx in range(1, rs_count + 1):
    with st.expander(f"RS{rs_idx} Details", expanded=True):
        fa, fb = st.columns(2)

        with fa:
            st.markdown(f"#### RS{rs_idx} Information")
            rs_name    = st.text_input(
                f"RS{rs_idx} Rectifier Name",
                placeholder="e.g. Eltek Flatpack 1",
                key=f"rs{rs_idx}_name",
            )
            rs_load    = st.text_input(
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
            st.caption("Fuse panel / load schedule photo")
            load_img = image_input_widget(
                label_upload    = f"Upload RS{rs_idx} Load Schedule image",
                label_paste     = f"Paste RS{rs_idx} Load Schedule from clipboard",
                upload_key      = f"rs{rs_idx}_load_img_upload",
                paste_btn_key   = f"paste_rs{rs_idx}_load_btn",
                session_key     = f"rs{rs_idx}_load_img_bytes",
                eval_key        = f"rs{rs_idx}_load_clip_eval",
                preview_caption = f"RS{rs_idx} Load Schedule",
            )

        st.markdown(f"#### RS{rs_idx} Existing Rectifier Photo")
        st.caption("Physical photo of the existing rectifier")
        existing_img = image_input_widget(
            label_upload    = f"Upload RS{rs_idx} Existing Rectifier image",
            label_paste     = f"Paste RS{rs_idx} Existing Rectifier from clipboard",
            upload_key      = f"rs{rs_idx}_existing_img_upload",
            paste_btn_key   = f"paste_rs{rs_idx}_existing_btn",
            session_key     = f"rs{rs_idx}_existing_img_bytes",
            eval_key        = f"rs{rs_idx}_existing_clip_eval",
            preview_caption = f"RS{rs_idx} Existing Rectifier",
        )

        rs_entries.append({
            "name":               rs_name.strip(),
            "load":               rs_load.strip(),
            "fuse_no":            rs_fuse_no.strip(),
            "load_img_bytes":     load_img,
            "existing_img_bytes": existing_img,
        })

st.divider()

# ── Supporting Documents ──────────────────────────────────────
st.subheader("Supporting Documents")
tssr_pdf = st.file_uploader(
    "TSSR PDF (filename will be written into the Word document)",
    type=["pdf"], key="tssr_pdf",
)
if tssr_pdf:
    st.success(f"PDF ready: {tssr_pdf.name}")

st.divider()

# ── Generate ──────────────────────────────────────────────────
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
    missing = [k for k in required_base if not data.get(k)]
    if missing:
        st.error("Please fill in: " + ", ".join(missing))
        st.stop()

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
        docx_bytes, gen_warnings = generate_docx_bytes(
            data=data,
            rs_entries=rs_entries,
            tssr_pdf_name=(tssr_pdf.name if tssr_pdf else ""),
        )

        out_name = (
            f"MOP_{safe_filename(data['site_name'])}"
            f"_{safe_filename(data['plaid'])}.docx"
        )

        st.success("MOP generated successfully!")
        for w in gen_warnings:
            if w.startswith("✅"):
                st.success(w)
            elif w.startswith("⚠️"):
                st.warning(w)
            else:
                st.info(w)

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