# app.py
import streamlit as st
from docx import Document
import io
import os
import re
import mammoth
from docx.shared import RGBColor

st.set_page_config(page_title="COA Placeholder Replacer", layout="wide")
st.title("📄 COA Placeholder Replacer — MOD / FAR (No calculations)")

# --- Configure template paths (update if needed) ---
TEMPLATE_PATHS = {
    "MOD": "/mnt/data/COA 7500-8000.docx",
    "FAR": "/mnt/data/[ 2025 ] LIPL FAR COA.docx"
}

# --- Utility: normalize obvious brace typos in template text ---
def normalize_broken_placeholders_in_doc(doc):
    """
    Fix some common broken placeholder formats like:
      ((BATCH_1}}  ->  {{BATCH_1}}
      {{BATCH_1))  ->  {{BATCH_1}}
    This function edits runs in-place to correct those mistakes before replacement.
    """
    # patterns to normalize: replace '((' -> '{{' and '))' -> '}}' only when they appear near placeholder names
    # We'll be conservative: only transform runs that contain either '((' or '))' or similar
    for para in doc.paragraphs:
        for run in para.runs:
            if "((" in run.text or "))" in run.text or "}}" in run.text and "((" in run.text:
                run.text = run.text.replace("((", "{{").replace("))", "}}").replace("}{", "}{")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if "((" in run.text or "))" in run.text or "}}" in run.text and "((" in run.text:
                            run.text = run.text.replace("((", "{{").replace("))", "}}").replace("}{", "}{")
    # headers/footers
    try:
        for section in doc.sections:
            header = section.header
            for para in header.paragraphs:
                for run in para.runs:
                    if "((" in run.text or "))" in run.text:
                        run.text = run.text.replace("((", "{{").replace("))", "}}")
            footer = section.footer
            for para in footer.paragraphs:
                for run in para.runs:
                    if "((" in run.text or "))" in run.text:
                        run.text = run.text.replace("((", "{{").replace("))", "}}")
    except Exception:
        pass


# --- Placeholder regex (matches {{KEY}} where KEY is letters, digits, underscore or hyphen) ---
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z0-9_\-]+)\s*\}\}")

# --- Replace placeholders in a paragraph while preserving style of first overlapping run ---
def replace_placeholders_in_paragraph(paragraph, replacements):
    runs = paragraph.runs
    if not runs:
        return

    # build full text and run offsets
    full_text = ""
    offsets = []  # list of (run, start_index, end_index)
    for run in runs:
        start = len(full_text)
        full_text += run.text
        end = len(full_text)
        offsets.append((run, start, end))

    # find placeholders in this paragraph
    matches = list(PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return

    # We'll process matches left-to-right; as we modify runs we also update offsets
    # To keep it simple, we work on the original offsets and runs, but clear overlapping runs as we go.
    for match in matches:
        key = match.group(1)
        if key not in replacements:
            continue
        replacement_text = str(replacements[key])

        placeholder_start, placeholder_end = match.start(), match.end()

        # find overlapping runs
        overlapping = [(r, s, e) for (r, s, e) in offsets if not (e <= placeholder_start or s >= placeholder_end)]
        if not overlapping:
            continue

        style_run, s0, e0 = overlapping[0]
        first_run, first_s, first_e = overlapping[0]
        last_run, last_s, last_e = overlapping[-1]

        # prefix (text in first run before placeholder)
        prefix_len = max(0, placeholder_start - first_s)
        prefix = first_run.text[:prefix_len]

        # suffix (text in last run after placeholder)
        suffix_start_in_last = placeholder_end - last_s
        suffix = last_run.text[suffix_start_in_last:]

        # clear overlapping runs
        for r, _, _ in overlapping:
            r.text = ""

        # set replacement into the first overlapping run
        new_text = prefix + replacement_text + suffix
        first_run.text = new_text

        # try to copy basic font attributes from style_run to first_run
        try:
            font = style_run.font
            first_run.font.name = font.name
            first_run.font.size = font.size
            first_run.font.bold = font.bold
            first_run.font.italic = font.italic
            first_run.font.underline = font.underline
            if font.color is not None and getattr(font.color, "rgb", None) is not None:
                first_run.font.color.rgb = font.color.rgb
        except Exception:
            # ignore attribute copy failures
            pass

def advanced_replace_text_preserving_style(doc, replacements):
    """
    Replace placeholders across the document while preserving basic style.
    Will run on main paragraphs, tables, headers and footers.
    """
    # Normalize obvious broken placeholders first
    normalize_broken_placeholders_in_doc(doc)

    # body paragraphs
    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, replacements)

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_placeholders_in_paragraph(para, replacements)

    # headers/footers
    try:
        for section in doc.sections:
            header = section.header
            for para in header.paragraphs:
                replace_placeholders_in_paragraph(para, replacements)
            footer = section.footer
            for para in footer.paragraphs:
                replace_placeholders_in_paragraph(para, replacements)
    except Exception:
        pass

# --- Convert DOCX to HTML for preview ---
def docx_to_html(docx_path):
    with open(docx_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        return result.value

# --- UI: choose COA type ---
coa_type = st.selectbox("Choose COA type", ["MOD", "FAR"])

template_path = TEMPLATE_PATHS.get(coa_type)
st.info(f"Using template: {os.path.basename(template_path) if template_path else 'Not found'}")

# If template not found, allow upload
if not template_path or not os.path.exists(template_path):
    st.warning("Template not found in expected path. Please upload the template (.docx) for this COA type.")
    uploaded = st.file_uploader(f"Upload {coa_type} template (.docx)", type=["docx"], key=f"upload_{coa_type}")
    if uploaded is not None:
        # save uploaded to a temporary file and use it
        tmp_path = f"/tmp/{coa_type}_template.docx"
        with open(tmp_path, "wb") as f:
            f.write(uploaded.read())
        template_path = tmp_path
        st.success("Template uploaded and saved for this session.")

# --- Build form fields for four batches (per your mapping) ---
st.markdown("### Enter COA values (four batches)")
with st.form("coa_form"):
    date_val = st.text_input("DATE (e.g., DD-MM-YYYY)", value="")
    # Batch labels (BATCH_1..BATCH_4)
    batch_1 = st.text_input("BATCH_1 (Batch 1 label)", value="")
    batch_2 = st.text_input("BATCH_2 (Batch 2 label)", value="")
    batch_3 = st.text_input("BATCH_3 (Batch 3 label)", value="")
    batch_4 = st.text_input("BATCH_4 (Batch 4 label)", value="")

    # Moisture M1..M4
    st.write("#### Moisture (M1..M4)")
    m1 = st.text_input("M1 (Batch 1 moisture %)", value="")
    m2 = st.text_input("M2 (Batch 2 moisture %)", value="")
    m3 = st.text_input("M3 (Batch 3 moisture %)", value="")
    m4 = st.text_input("M4 (Batch 4 moisture %)", value="")

    # Viscosity 2h: B1V1..B4V1
    st.write("#### Viscosity after 2 hours (B1V1..B4V1)")
    b1v1 = st.text_input("B1V1 (Batch 1 - 2h)", value="")
    b2v1 = st.text_input("B2V1 (Batch 2 - 2h)", value="")
    b3v1 = st.text_input("B3V1 (Batch 3 - 2h)", value="")
    b4v1 = st.text_input("B4V1 (Batch 4 - 2h)", value="")

    # Viscosity 24h: B1V2..B4V2
    st.write("#### Viscosity after 24 hours (B1V2..B4V2)")
    b1v2 = st.text_input("B1V2 (Batch 1 - 24h)", value="")
    b2v2 = st.text_input("B2V2 (Batch 2 - 24h)", value="")
    b3v2 = st.text_input("B3V2 (Batch 3 - 24h)", value="")
    b4v2 = st.text_input("B4V2 (Batch 4 - 24h)", value="")

    # pH PH1..PH4
    st.write("#### pH (PH1..PH4)")
    ph1 = st.text_input("PH1 (Batch 1 pH)", value="")
    ph2 = st.text_input("PH2 (Batch 2 pH)", value="")
    ph3 = st.text_input("PH3 (Batch 3 pH)", value="")
    ph4 = st.text_input("PH4 (Batch 4 pH)", value="")

    # Mesh MESH1..MESH4
    st.write("#### 200 Mesh (MESH1..MESH4)")
    mesh1 = st.text_input("MESH1 (Batch 1 % through 200 mesh)", value="")
    mesh2 = st.text_input("MESH2 (Batch 2 % through 200 mesh)", value="")
    mesh3 = st.text_input("MESH3 (Batch 3 % through 200 mesh)", value="")
    mesh4 = st.text_input("MESH4 (Batch 4 % through 200 mesh)", value="")

    # Bulk Density BD1..BD4
    st.write("#### Bulk Density (BD1..BD4)")
    bd1 = st.text_input("BD1 (Batch 1 Bulk Density)", value="")
    bd2 = st.text_input("BD2 (Batch 2 Bulk Density)", value="")
    bd3 = st.text_input("BD3 (Batch 3 Bulk Density)", value="")
    bd4 = st.text_input("BD4 (Batch 4 Bulk Density)", value="")

    # Fann @ 3' F1..F4
    st.write("#### Fann Viscosity @ 3' (F1..F4)")
    f1 = st.text_input("F1 (Batch 1 Fann 3min)", value="")
    f2 = st.text_input("F2 (Batch 2 Fann 3min)", value="")
    f3 = st.text_input("F3 (Batch 3 Fann 3min)", value="")
    f4 = st.text_input("F4 (Batch 4 Fann 3min)", value="")

    # Fann @ 30' FV1..FV4
    st.write("#### Fann Viscosity @ 30' (FV1..FV4)")
    fv1 = st.text_input("FV1 (Batch 1 Fann 30min)", value="")
    fv2 = st.text_input("FV2 (Batch 2 Fann 30min)", value="")
    fv3 = st.text_input("FV3 (Batch 3 Fann 30min)", value="")
    fv4 = st.text_input("FV4 (Batch 4 Fann 30min)", value="")

    submitted = st.form_submit_button("Generate COA")

# --- On submit: perform placeholder replacement and provide preview & download ---
if submitted:
    if not template_path or not os.path.exists(template_path):
        st.error("Template file not found. Please upload the template or correct TEMPLATE_PATHS.")
    else:
        # build replacements dictionary matching template placeholders exactly
        replacements = {
            # date (matching placeholder seen in FAR: DD-MM-YYYY)
            "DD-MM-YYYY": date_val,

            # batch labels
            "BATCH_1": batch_1,
            "BATCH_2": batch_2,
            "BATCH_3": batch_3,
            "BATCH_4": batch_4,

            # moisture
            "M1": m1,
            "M2": m2,
            "M3": m3,
            "M4": m4,

            # viscosities (2h)
            "B1V1": b1v1,
            "B2V1": b2v1,
            "B3V1": b3v1,
            "B4V1": b4v1,

            # viscosities (24h)
            "B1V2": b1v2,
            "B2V2": b2v2,
            "B3V2": b3v2,
            "B4V2": b4v2,

            # pH
            "PH1": ph1,
            "PH2": ph2,
            "PH3": ph3,
            "PH4": ph4,

            # mesh
            "MESH1": mesh1,
            "MESH2": mesh2,
            "MESH3": mesh3,
            "MESH4": mesh4,

            # bulk density
            "BD1": bd1,
            "BD2": bd2,
            "BD3": bd3,
            "BD4": bd4,

            # fann @ 3'
            "F1": f1,
            "F2": f2,
            "F3": f3,
            "F4": f4,

            # fann @ 30'
            "FV1": fv1,
            "FV2": fv2,
            "FV3": fv3,
            "FV4": fv4
        }

        # load template
        try:
            doc = Document(template_path)
        except Exception as e:
            st.error(f"Failed to open template: {e}")
            doc = None

        if doc:
            # perform replacements
            advanced_replace_text_preserving_style(doc, replacements)

            # save generated file to a BytesIO buffer
            output_filename = f"COA_{coa_type}_generated.docx"
            output_path = f"/tmp/{output_filename}"
            try:
                doc.save(output_path)
            except Exception as e:
                st.error(f"Failed to save generated DOCX: {e}")
                output_path = None

            if output_path and os.path.exists(output_path):
                # try preview
                try:
                    html = docx_to_html(output_path)
                    st.subheader("📄 Preview")
                    st.components.v1.html(f"<div style='padding:10px'>{html}</div>", height=700, scrolling=True)
                except Exception as e:
                    st.warning(f"Preview failed ({e}). You can still download the generated file below.")

                # Provide download
                with open(output_path, "rb") as f:
                    doc_bytes = f.read()
                buffer = io.BytesIO(doc_bytes)

                st.download_button(
                    label="📥 Download generated COA (DOCX)",
                    data=buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                st.success("COA generated. Check the downloaded file and open in Word to confirm formatting.")
