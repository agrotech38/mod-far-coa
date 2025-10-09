# app.py
import streamlit as st
from docx import Document
import io
import os
import re
import mammoth
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
    KOLKATA = ZoneInfo("Asia/Kolkata")
except Exception:
    import pytz
    KOLKATA = pytz.timezone("Asia/Kolkata")

st.set_page_config(page_title="ModFar COA Generator", layout="wide")
st.title("📄 ModFar COA Generator")

# --- Defaults: local paths (these match the files you uploaded) ---
DEFAULT_TEMPLATES = {
    "MOD": "PH LIPL MOD COA.docx",
    "FAR": "PH LIPL FAR COA.docx"
}

# --- Regex for placeholders like {{KEY}} ---
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z0-9_\-/]+)\s*\}\}")

# --- Utility to normalize common broken placeholders e.g. '((BATCH_1}}' -> '{{BATCH_1}}' ---
def normalize_broken_placeholders_in_doc(doc):
    for para in doc.paragraphs:
        for run in para.runs:
            if "((" in run.text or "))" in run.text:
                run.text = run.text.replace("((", "{{").replace("))", "}}")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if "((" in run.text or "))" in run.text:
                            run.text = run.text.replace("((", "{{").replace("))", "}}")
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

# --- Replace placeholders in a paragraph while preserving style of first overlapping run ---
def replace_placeholders_in_paragraph(paragraph, replacements):
    runs = paragraph.runs
    if not runs:
        return

    # Build full text and offsets
    full_text = ""
    offsets = []  # (run, start, end)
    for run in runs:
        start = len(full_text)
        full_text += run.text
        end = len(full_text)
        offsets.append((run, start, end))

    matches = list(PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return

    for match in matches:
        key = match.group(1)
        if key not in replacements:
            continue
        replacement_text = str(replacements[key])
        p_start, p_end = match.start(), match.end()

        overlapping = [(r, s, e) for (r, s, e) in offsets if not (e <= p_start or s >= p_end)]
        if not overlapping:
            continue

        first_run, first_s, first_e = overlapping[0]
        last_run, last_s, last_e = overlapping[-1]

        # prefix & suffix around placeholder inside first/last runs
        prefix_len = max(0, p_start - first_s)
        prefix = first_run.text[:prefix_len]
        suffix_start_in_last = p_end - last_s
        suffix = last_run.text[suffix_start_in_last:]

        # clear overlapping runs
        for r, _, _ in overlapping:
            r.text = ""

        new_text = prefix + replacement_text + suffix
        first_run.text = new_text

        # copy style from the first overlapping run (safe copy)
        try:
            font = first_run.font
            if font.name:
                first_run.font.name = font.name
            if font.size:
                first_run.font.size = font.size
            first_run.font.bold = font.bold
            first_run.font.italic = font.italic
            first_run.font.underline = font.underline
            if font.color and getattr(font.color, "rgb", None) is not None:
                first_run.font.color.rgb = font.color.rgb
        except Exception:
            pass

def advanced_replace_text_preserving_style(doc, replacements):
    normalize_broken_placeholders_in_doc(doc)

    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, replacements)

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

# --- Convert to HTML for preview using mammoth ---
def docx_to_html_bytes(docx_bytes):
    result = mammoth.convert_to_html(io.BytesIO(docx_bytes))
    return result.value

# --- Simplified template retrieval: only local default templates are used --- 
def get_template_bytes(coa_type):
    path = DEFAULT_TEMPLATES.get(coa_type)
    if path and os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    return None

# --- Default date (DD/MM/YYYY) in Asia/Kolkata timezone, editable by user ---
now_kolkata = datetime.now(KOLKATA)
default_ddmmyyyy_slash = now_kolkata.strftime("%d/%m/%Y")  # DD/MM/YYYY

st.markdown("### Enter values (fill Batch 1 completely, then Batch 2, then Batch 3, then Batch 4)")
coa_type = st.selectbox("Choose COA type", ["MOD", "FAR"])

# Use tabs: one tab per batch to enforce batch-by-batch entry
tab1, tab2, tab3, tab4 = st.tabs(["Batch 1", "Batch 2", "Batch 3", "Batch 4"])

# holder for batch inputs
batches = {"1": {}, "2": {}, "3": {}, "4": {}}

# Date input placed in Batch 1 tab (editable)
with tab1:
    st.subheader("Batch 1")
    date_field = st.text_input("Date (DD/MM/YYYY)", value=default_ddmmyyyy_slash, key="date_field")
    batches["1"]["BATCH"] = st.text_input("BATCH_1 (Label)", key="batch1_label")
    batches["1"]["M"] = st.text_input("M1 (Moisture)", key="m1")
    if coa_type == "MOD":
        batches["1"]["B1V1"] = st.text_input("B1V1 (30min viscosity)", key="b1v1_mod")
        batches["1"]["B1V2"] = st.text_input("B1V2 (60min viscosity)", key="b1v2_mod")
        batches["1"]["PH"] = st.text_input("PH1 (pH)", key="ph1_mod")
    else:
        batches["1"]["B1V1"] = st.text_input("B1V1 (2h viscosity)", key="b1v1_far")
        batches["1"]["B1V2"] = st.text_input("B1V2 (24h viscosity)", key="b1v2_far")
        batches["1"]["PH"] = st.text_input("PH1 (pH)", key="ph1_far")
        batches["1"]["MESH"] = st.text_input("MESH1 (200 mesh %)", key="mesh1")
        batches["1"]["BD"] = st.text_input("BD1 (Bulk Density)", key="bd1")
        batches["1"]["F"] = st.text_input("F1 (Fann 3')", key="f1")
        batches["1"]["FV"] = st.text_input("FV1 (Fann 30')", key="fv1")

with tab2:
    st.subheader("Batch 2")
    batches["2"]["BATCH"] = st.text_input("BATCH_2 (Label)", key="batch2_label")
    batches["2"]["M"] = st.text_input("M2 (Moisture)", key="m2")
    if coa_type == "MOD":
        batches["2"]["B1V1"] = st.text_input("B2V1 (30min viscosity)", key="b2v1_mod")
        batches["2"]["B1V2"] = st.text_input("B2V2 (60min viscosity)", key="b2v2_mod")
        batches["2"]["PH"] = st.text_input("PH2 (pH)", key="ph2_mod")
    else:
        batches["2"]["B1V1"] = st.text_input("B2V1 (2h viscosity)", key="b2v1_far")
        batches["2"]["B1V2"] = st.text_input("B2V2 (24h viscosity)", key="b2v2_far")
        batches["2"]["PH"] = st.text_input("PH2 (pH)", key="ph2_far")
        batches["2"]["MESH"] = st.text_input("MESH2 (200 mesh %)", key="mesh2")
        batches["2"]["BD"] = st.text_input("BD2 (Bulk Density)", key="bd2")
        batches["2"]["F"] = st.text_input("F2 (Fann 3')", key="f2")
        batches["2"]["FV"] = st.text_input("FV2 (Fann 30')", key="fv2")

with tab3:
    st.subheader("Batch 3")
    batches["3"]["BATCH"] = st.text_input("BATCH_3 (Label)", key="batch3_label")
    batches["3"]["M"] = st.text_input("M3 (Moisture)", key="m3")
    if coa_type == "MOD":
        batches["3"]["B1V1"] = st.text_input("B3V1 (30min viscosity)", key="b3v1_mod")
        batches["3"]["B1V2"] = st.text_input("B3V2 (60min viscosity)", key="b3v2_mod")
        batches["3"]["PH"] = st.text_input("PH3 (pH)", key="ph3_mod")
    else:
        batches["3"]["B1V1"] = st.text_input("B3V1 (2h viscosity)", key="b3v1_far")
        batches["3"]["B1V2"] = st.text_input("B3V2 (24h viscosity)", key="b3v2_far")
        batches["3"]["PH"] = st.text_input("PH3 (pH)", key="ph3_far")
        batches["3"]["MESH"] = st.text_input("MESH3 (200 mesh %)", key="mesh3")
        batches["3"]["BD"] = st.text_input("BD3 (Bulk Density)", key="bd3")
        batches["3"]["F"] = st.text_input("F3 (Fann 3')", key="f3")
        batches["3"]["FV"] = st.text_input("FV3 (Fann 30')", key="fv3")

with tab4:
    st.subheader("Batch 4")
    batches["4"]["BATCH"] = st.text_input("BATCH_4 (Label)", key="batch4_label")
    batches["4"]["M"] = st.text_input("M4 (Moisture)", key="m4")
    if coa_type == "MOD":
        batches["4"]["B1V1"] = st.text_input("B4V1 (30min viscosity)", key="b4v1_mod")
        batches["4"]["B1V2"] = st.text_input("B4V2 (60min viscosity)", key="b4v2_mod")
        batches["4"]["PH"] = st.text_input("PH4 (pH)", key="ph4_mod")
    else:
        batches["4"]["B1V1"] = st.text_input("B4V1 (2h viscosity)", key="b4v1_far")
        batches["4"]["B1V2"] = st.text_input("B4V2 (24h viscosity)", key="b4v2_far")
        batches["4"]["PH"] = st.text_input("PH4 (pH)", key="ph4_far")
        batches["4"]["MESH"] = st.text_input("MESH4 (200 mesh %)", key="mesh4")
        batches["4"]["BD"] = st.text_input("BD4 (Bulk Density)", key="bd4")
        batches["4"]["F"] = st.text_input("F4 (Fann 3')", key="f4")
        batches["4"]["FV"] = st.text_input("FV4 (Fann 30')", key="fv4")

# Generate button below tabs
if st.button("Generate COA"):
    template_bytes = get_template_bytes(coa_type)
    if template_bytes is None:
        st.error("Template file not found in server path. Ensure the template file is present in the app directory.")
    else:
        try:
            doc = Document(io.BytesIO(template_bytes))
        except Exception as e:
            st.error(f"Failed to open template as docx: {e}")
            doc = None

        if doc:
            replacements = {}
            # populate date placeholder - main requested format
            replacements["DD/MM/YYYY"] = date_field
            # also populate DD-MM-YYYY for templates that might use that format (keeps compatibility)
            replacements["DD-MM-YYYY"] = date_field.replace("/", "-")

            # fill batch placeholders
            for i in ("1", "2", "3", "4"):
                replacements[f"BATCH_{i}"] = batches[i].get("BATCH", "")
                replacements[f"M{i}"] = batches[i].get("M", "")
                replacements[f"B{i}V1"] = batches[i].get("B1V1", "")
                replacements[f"B{i}V2"] = batches[i].get("B1V2", "")
                replacements[f"PH{i}"] = batches[i].get("PH", "")
                if coa_type == "FAR":
                    replacements[f"MESH{i}"] = batches[i].get("MESH", "")
                    replacements[f"BD{i}"] = batches[i].get("BD", "")
                    replacements[f"F{i}"] = batches[i].get("F", "")
                    replacements[f"FV{i}"] = batches[i].get("FV", "")

            # perform replacements
            advanced_replace_text_preserving_style(doc, replacements)

            # save to bytes
            out_buffer = io.BytesIO()
            doc.save(out_buffer)
            out_bytes = out_buffer.getvalue()

            # preview
            try:
                html = docx_to_html_bytes(out_bytes)
                st.subheader("📄 Preview (HTML)")
                st.components.v1.html(f"<div style='padding:12px'>{html}</div>", height=700, scrolling=True)
            except Exception as e:
                st.warning(f"Preview (mammoth) failed: {e}. You can still download the DOCX.")

            filename = f"COA_{coa_type}_{batches['1'].get('BATCH','batch1')}.docx"
            st.download_button(
                label="📥 Download generated DOCX",
                data=out_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Generated. Open the downloaded DOCX in MS Word to confirm visual formatting.")
