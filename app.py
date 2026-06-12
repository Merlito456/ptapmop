import os
import re
import time
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Template.docx")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "bmp"}
ALLOWED_PDF_EXTENSIONS = {"pdf"}

app = Flask(__name__)
app.secret_key = "mop-secret-key"


def ensure_dirs():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def safe_filename(s: str) -> str:
    s = re.sub(r'[\\/*?:"<>|]+', "_", s)
    return s.strip().replace(" ", "_")


def allowed_file(filename, allowed_exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_exts


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


def insert_image_after(paragraph, image_path, width=Inches(4.5), caption=None):
    p_img = insert_paragraph_after(paragraph)
    run = p_img.add_run()
    run.add_picture(image_path, width=width)

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


def insert_power_section_content(doc, data):
    fuse_variants = [
        "FUSE No: L3 Nokia OLT MF-2 – ( Nokia Power tapping point)",
        "FUSE No: L3 Nokia OLT MF-2 – (Nokia Power tapping point)",
        "FUSE No: L3 Nokia OLT MF-2",
        "FUSE No:"
    ]

    anchor, where = find_any_anchor(doc, fuse_variants)
    if not anchor:
        return False, "Could not find FUSE anchor text in template."

    olt_label = build_olt_label(data["equipment"], data["olt_label_custom"])
    new_fuse = build_fuse_line(olt_label, data["equipment"])

    old_text = paragraph_full_text(anchor)
    replaced = old_text
    exacts = [
        "FUSE No: L3 Nokia OLT MF-2 – ( Nokia Power tapping point)",
        "FUSE No: L3 Nokia OLT MF-2 – (Nokia Power tapping point)",
    ]
    exact_hit = False
    for ex in exacts:
        if ex in replaced:
            replaced = replaced.replace(ex, new_fuse)
            exact_hit = True

    if exact_hit:
        set_paragraph_text(anchor, replaced)
    else:
        set_paragraph_text(anchor, new_fuse)

    rs1_line = f"RS1 + {data['rs1_rectifier_name']} + {data['rs1_load_assignment']}"
    rs2_line = f"RS2 + {data['rs2_rectifier_name']} + {data['rs2_load_assignment']}"

    p1 = insert_paragraph_after(anchor, rs1_line)
    last = p1
    if data.get("rs1_image"):
        img_p, _ = insert_image_after(p1, data["rs1_image"], width=Inches(4.5), caption="RS1")
        last = img_p

    p2 = insert_paragraph_after(last, rs2_line)
    if data.get("rs2_image"):
        insert_image_after(p2, data["rs2_image"], width=Inches(4.5), caption="RS2")

    return True, f"Inserted RS1/RS2 content under FUSE anchor in {where}."


def insert_supporting_documents(doc, pdf_path):
    if not pdf_path:
        return False, "No TSSR PDF selected."

    anchor, where = find_any_anchor(doc, ["Supporting Documents"])
    if not anchor:
        return False, "Could not find 'Supporting Documents' section."

    p = insert_paragraph_after(anchor)
    p.add_run("TSSR: ")
    file_url = "file:///" + pdf_path.replace("\\", "/")
    add_hyperlink(p, os.path.basename(pdf_path), file_url)
    return True, f"Inserted TSSR hyperlink under Supporting Documents in {where}."


def insert_existing_rectifier_image(doc, image_path):
    if not image_path:
        return False, "No Page 8 rectifier image selected."

    anchor, where = find_any_anchor(doc, ["Existing Rectifier"])
    if not anchor:
        return False, "Could not find 'Existing Rectifier' section."

    insert_image_after(anchor, image_path, width=Inches(5.0), caption="Existing Rectifier")
    return True, f"Inserted rectifier image under Existing Rectifier in {where}."


def save_upload(file_storage, subname):
    if not file_storage or not file_storage.filename:
        return ""

    filename = secure_filename(file_storage.filename)
    timestamp = str(int(time.time() * 1000))
    final_name = f"{subname}_{timestamp}_{filename}"
    path = os.path.join(UPLOAD_DIR, final_name)
    file_storage.save(path)
    return path


def save_base64_image(data_url, subname):
    if not data_url:
        return ""

    import base64

    if "," not in data_url:
        return ""

    header, encoded = data_url.split(",", 1)
    timestamp = str(int(time.time() * 1000))
    filename = f"{subname}_{timestamp}.png"
    path = os.path.join(UPLOAD_DIR, filename)

    with open(path, "wb") as f:
        f.write(base64.b64decode(encoded))

    return path


def generate_mop(data):
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

    insert_power_section_content(doc, data)
    insert_supporting_documents(doc, data.get("tssr_pdf"))
    insert_existing_rectifier_image(doc, data.get("rectifier_image_page8"))

    output_file = os.path.join(
        OUTPUT_DIR,
        f"MOP_{safe_filename(data['site_name'])}_{safe_filename(data['plaid'])}.docx"
    )
    doc.save(output_file)
    return output_file


@app.route("/", methods=["GET", "POST"])
def index():
    ensure_dirs()

    if request.method == "POST":
        try:
            data = {
                "site_name": request.form.get("site_name", "").strip(),
                "plaid": request.form.get("plaid", "").strip(),
                "equipment": request.form.get("equipment", "").strip(),
                "olt_label_custom": request.form.get("olt_label_custom", "").strip(),
                "prepared_by": request.form.get("prepared_by", "").strip(),
                "position": request.form.get("position", "").strip(),
                "target_datetime": request.form.get("target_datetime", "").strip(),
                "rs1_rectifier_name": request.form.get("rs1_rectifier_name", "").strip(),
                "rs1_load_assignment": request.form.get("rs1_load_assignment", "").strip(),
                "rs2_rectifier_name": request.form.get("rs2_rectifier_name", "").strip(),
                "rs2_load_assignment": request.form.get("rs2_load_assignment", "").strip(),
            }

            required = [
                "site_name", "plaid", "equipment", "prepared_by", "position",
                "target_datetime", "rs1_rectifier_name", "rs1_load_assignment",
                "rs2_rectifier_name", "rs2_load_assignment"
            ]
            missing = [f for f in required if not data[f]]
            if missing:
                flash("Missing fields: " + ", ".join(missing), "error")
                return redirect(url_for("index"))

            rs1_image = save_upload(request.files.get("rs1_image"), "rs1_image")
            rs2_image = save_upload(request.files.get("rs2_image"), "rs2_image")
            rectifier_image = save_upload(request.files.get("rectifier_image_page8"), "rectifier_image")
            tssr_pdf = save_upload(request.files.get("tssr_pdf"), "tssr_pdf")

            # clipboard-pasted images from hidden fields
            if not rs1_image:
                rs1_image = save_base64_image(request.form.get("rs1_image_paste", ""), "rs1_pasted")
            if not rs2_image:
                rs2_image = save_base64_image(request.form.get("rs2_image_paste", ""), "rs2_pasted")
            if not rectifier_image:
                rectifier_image = save_base64_image(request.form.get("rectifier_image_paste", ""), "rectifier_pasted")

            data["rs1_image"] = rs1_image
            data["rs2_image"] = rs2_image
            data["rectifier_image_page8"] = rectifier_image
            data["tssr_pdf"] = tssr_pdf

            output_file = generate_mop(data)
            return send_file(output_file, as_attachment=True)

        except Exception as e:
            flash(str(e), "error")
            return redirect(url_for("index"))

    return render_template("index.html")


if __name__ == "__main__":
    ensure_dirs()
    app.run(debug=True)