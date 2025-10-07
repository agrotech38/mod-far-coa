import streamlit as st
from docx import Document
import datetime
import os
import pandas as pd
from io import BytesIO
import zipfile

# --- Replace placeholders robustly even in tables ---
def replace_placeholders(doc, replacements):
    import re
    pattern = re.compile(r"{{(.*?)}}")

    def replace_in_paragraph(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)
        for key, val in replacements.items():
            full_text = full_text.replace(f"{{{{{key}}}}}", str(val))
        paragraph.clear()
        paragraph.add_run(full_text)

    for p in doc.paragraphs:
        if pattern.search(p.text):
            replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if pattern.search(p.text):
                        replace_in_paragraph(p)

# --- Generate DOCX ---
def generate_docx(data, template_path="SALARY SLIP FORMAT.docx"):
    doc = Document(template_path)
    replace_placeholders(doc, data)
    file_name = f"salaryslip_{data['Name'].replace(' ', '_')}_{data['Month'].replace(' ', '_')}.docx"
    doc.save(file_name)
    return file_name

# --- Streamlit App ---
st.set_page_config(page_title="Salary Slip Bulk Generator", layout="wide")
st.title("📄 Bulk Salary Slip Generator")

# --- Bulk Generator ---
uploaded_file = st.file_uploader("📤 Upload Excel File for Bulk Generation", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ Uploaded {len(df)} row(s)")

    generated_files = []

    for index, row in df.iterrows():
        name = row["Name"]
        total = row["Salary"] + row["Bonus"] + row["Other"]
        net_advance = row["Advance_Till_Date"] - row["Advance_Deduct"]
        payable = total - (row["ESI"] + row["Advance_Deduct"] + row["MISC"])
        payment_date = datetime.datetime.now().strftime("%d %B %Y")

        data = {
            "Name": name,
            "Designation": row["Designation"],
            "Department": row["Department"],
            "Total_Days": row["Total_Days"],
            "Working_Days": row["Working_Days"],
            "Weekly_Off": row["Weekly_Off"],
            "Festival_Off": row["Festival_Off"],
            "Paid_Days": row["Paid_Days"],
            "Base": row["Base"],
            "Month": row["Month"],
            "Salary": row["Salary"],
            "Bonus": row["Bonus"],
            "Other": row["Other"],
            "Total": total,
            "ESI": row["ESI"],
            "Advance_Till_Date": row["Advance_Till_Date"],
            "Advance_Deduct": row["Advance_Deduct"],
            "Net_Advance": net_advance,
            "MISC": row["MISC"],
            "Payable": payable,
            "Payment_Date": payment_date,
        }

        docx_path = generate_docx(data)
        generated_files.append(docx_path)

        # --- Individual Download Button ---
        with open(docx_path, "rb") as f:
            st.download_button(
                label=f"📄 Download {os.path.basename(docx_path)}",
                data=f,
                file_name=os.path.basename(docx_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"download_{index}"
            )

    # --- Zip Download Option ---
    zip_path = "generated_salary_slips.zip"
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for file in generated_files:
            zipf.write(file)

    with open(zip_path, "rb") as zipf:
        st.download_button(
            label="⬇️ Download All Salary Slips as ZIP",
            data=zipf,
            file_name="generated_salary_slips.zip",
            mime="application/zip",
            key="zip_download_final"
        )
