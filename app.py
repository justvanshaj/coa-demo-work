import streamlit as st
from docx import Document
import random
import os
import mammoth
import io
from docx.shared import RGBColor
from datetime import datetime
import calendar

# --- Replace placeholders preserving font and color ---
def advanced_replace_text_preserving_style(doc, replacements):
    def replace_in_paragraph(paragraph):
        runs = paragraph.runs
        full_text = ''.join(run.text for run in runs)
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in full_text:
                new_runs = []
                accumulated = ""
                for run in runs:
                    accumulated += run.text
                    new_runs.append(run)
                    if placeholder in accumulated:
                        style_run = next((r for r in new_runs if placeholder in r.text), new_runs[0])
                        font = style_run.font
                        accumulated = accumulated.replace(placeholder, value)
                        for r in new_runs:
                            r.text = ''
                        if new_runs:
                            new_run = new_runs[0]
                            new_run.text = accumulated
                            # preserve font attributes where possible
                            try:
                                new_run.font.name = font.name
                            except:
                                pass
                            try:
                                new_run.font.size = font.size
                            except:
                                pass
                            try:
                                new_run.font.bold = font.bold
                            except:
                                pass
                            try:
                                new_run.font.italic = font.italic
                            except:
                                pass
                            try:
                                new_run.font.underline = font.underline
                            except:
                                pass
                            try:
                                new_run.font.color.rgb = font.color.rgb
                            except:
                                pass
                        break

    for para in doc.paragraphs:
        replace_in_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para)

# --- Generate DOCX ---
def generate_docx(data, template_path="template.docx", output_path="generated_coa.docx"):
    doc = Document(template_path)
    advanced_replace_text_preserving_style(doc, data)
    doc.save(output_path)
    return output_path

# --- Convert DOCX to HTML preview ---
def docx_to_html(docx_path):
    with open(docx_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        return result.value

# --- Deterministic component function (from user inputs) ---
def calculate_components_from_inputs(moisture, fat, air, ash, protein, gum):
    """
    Receives user-entered component values and returns them rounded to 2 decimals.
    """
    fat = round(float(fat), 2)
    air = round(float(air), 2)
    ash = round(float(ash), 2)
    protein = round(float(protein), 2)
    gum = round(float(gum), 2)
    # Ensure the numbers are non-negative
    fat = max(fat, 0.0)
    air = max(air, 0.0)
    ash = max(ash, 0.0)
    protein = max(protein, 0.0)
    gum = max(gum, 0.0)
    return gum, protein, ash, air, fat

# --- Streamlit App Starts ---
st.set_page_config(page_title="COA Generator", layout="wide")
st.title("🧪 COA Document Generator (Code-Based Template)")

with st.form("coa_form"):
    code = st.selectbox(
        "Select Product Code Range",
        [f"{i}-{i+500}" for i in range(500, 10001, 500)]
    )
    st.info(f"📄 Using template: COA {code}.docx")

    date = st.text_input("Date (e.g., July 2025)")

    # Auto-calculate Best Before (made more robust to common input cases)
    best_before = ""
    if date:
        parsed = None
        for fmt_try in ("%B %Y", "%b %Y", "%B, %Y", "%b, %Y"):
            try:
                # try with title-cased month to be more forgiving
                parsed = datetime.strptime(date.strip().title(), fmt_try)
                break
            except:
                try:
                    parsed = datetime.strptime(date.strip(), fmt_try)
                    break
                except:
                    parsed = None
        if parsed:
            year = parsed.year + 2
            month = parsed.month - 1
            if month == 0:
                month = 12
                year -= 1
            best_before = f"{calendar.month_name[month].upper()} {year}"
            st.success(f"Best Before auto-filled: {best_before}")
        else:
            st.warning("Enter Date in format: July 2025")

    batch_no = st.text_input("Batch Number")
    moisture = st.number_input("Moisture (%)", min_value=0.0, max_value=100.0, step=0.01, value=10.0)

    # --- NEW: explicit user-entered components with required ranges ---
    st.markdown("### Enter components (must be within specified ranges)")
    fat = st.number_input("Fat (%) — range 0.45 to 0.55", min_value=0.45, max_value=0.55, step=0.01, value=0.50, format="%.2f")
    air = st.number_input("Air (%) — range 2.90 to 3.10", min_value=2.90, max_value=3.10, step=0.01, value=3.00, format="%.2f")
    ash = st.number_input("Ash Content (%) — range 0.45 to 0.55", min_value=0.45, max_value=0.55, step=0.01, value=0.50, format="%.2f")
    protein = st.number_input("Protein (%) — range 2.45 to 2.55", min_value=2.45, max_value=2.55, step=0.01, value=2.50, format="%.2f")
    gum = st.number_input("Gum Content (%) — range 80.10 to 89.95", min_value=80.10, max_value=89.95, step=0.01, value=83.00, format="%.2f")

    # Optional convenience: let user allow automatic gum adjustment to make sum = 100
    auto_adjust_gum = st.checkbox("Auto-adjust gum to make sum 100 (only if result stays within allowed gum range)", value=False)

    ph = st.text_input("pH Level (e.g., 6.7)")
    mesh_200 = st.text_input("200 Mesh (%)")
    viscosity_2h = st.text_input("Viscosity After 2 Hours (CPS)")
    viscosity_24h = st.text_input("Viscosity After 24 Hours (CPS)")
    submitted = st.form_submit_button("Generate COA")

if submitted:
    # Validation of ranges (number_input enforces min/max, but double-check)
    errs = []
    def in_range(val, lo, hi):
        return (val >= lo - 1e-9) and (val <= hi + 1e-9)

    if not in_range(fat, 0.45, 0.55):
        errs.append("Fat must be between 0.45 and 0.55.")
    if not in_range(air, 2.90, 3.10):
        errs.append("Air must be between 2.90 and 3.10.")
    if not in_range(ash, 0.45, 0.55):
        errs.append("Ash Content must be between 0.45 and 0.55.")
    if not in_range(protein, 2.45, 2.55):
        errs.append("Protein must be between 2.45 and 2.55.")
    if not in_range(gum, 80.10, 89.95):
        errs.append("Gum Content must be between 80.10 and 89.95.")
    if not in_range(moisture, 0.0, 100.0):
        errs.append("Moisture must be between 0 and 100.")

    if errs:
        for e in errs:
            st.error(e)
    else:
        # compute sum and optionally auto-adjust gum
        total = round(moisture + fat + air + ash + protein + gum, 2)
        tolerance = 0.01

        if abs(total - 100.0) <= tolerance:
            # OK — exactly 100 (within tolerance)
            gum_final = round(gum, 2)
            gum_used_auto = False
        else:
            # sum is not 100 — try auto-adjust if requested
            if auto_adjust_gum:
                needed = round(100.0 - (moisture + fat + air + ash + protein), 2)
                if in_range(needed, 80.10, 89.95):
                    gum_final = round(needed, 2)
                    gum_used_auto = True
                    st.info(f"Gum auto-adjusted to {gum_final}% to make total 100.00%")
                else:
                    st.error(f"Auto-adjust failed: required gum would be {needed}%, which is outside allowed range (80.10–89.95). Please adjust inputs.")
                    st.stop()
            else:
                st.error(f"Components sum to {total}%. They must sum to 100.00%. Either correct the inputs or enable 'Auto-adjust gum'.")
                st.stop()

        # all good — produce values
        gum_val, protein_val, ash_val, air_val, fat_val = calculate_components_from_inputs(moisture, fat, air, ash, protein, gum_final)

        # Optional: show summary table for final composition
        try:
            import pandas as pd
            summary_df = pd.DataFrame({
                "Component": ["Moisture", "Fat", "Air", "Ash", "Protein", "Gum"],
                "Value (%)": [round(moisture,2), fat_val, air_val, ash_val, protein_val, gum_val]
            })
            st.subheader("Composition summary")
            st.dataframe(summary_df, width=600)
        except Exception:
            pass

        data = {
            "DATE": date,
            "BATCH_NO": batch_no,
            "BEST_BEFORE": best_before,
            "MOISTURE": f"{round(moisture,2)}%",
            "PH": ph,
            "MESH_200": f"{mesh_200}%",
            "VISCOSITY_2H": viscosity_2h,
            "VISCOSITY_24H": viscosity_24h,
            "GUM_CONTENT": f"{gum_val}%",
            "PROTEIN": f"{protein_val}%",
            "ASH_CONTENT": f"{ash_val}%",
            "AIR": f"{air_val}%",
            "FAT": f"{fat_val}%"
        }

        template_path = f"COA {code}.docx"
        output_path = "generated_coa.docx"

        if not os.path.exists(template_path):
            st.error(f"Template file 'COA {code}.docx' not found.")
        else:
            generate_docx(data, template_path=template_path, output_path=output_path)

            try:
                html = docx_to_html(output_path)
                st.subheader("📄 Preview")
                st.components.v1.html(f"<div style='padding:15px'>{html}</div>", height=700, scrolling=True)
            except:
                st.warning("Preview failed. You can still download the file below.")

            # Rename file based on batch & code
            safe_batch = (batch_no or "batch").replace("/", "_").replace("\\", "_").replace(" ", "_")
            final_filename = f"COA-{safe_batch}-{code}.docx"

            with open(output_path, "rb") as f:
                doc_bytes = f.read()
            buffer = io.BytesIO(doc_bytes)

            st.download_button(
                label="📥 Download COA (DOCX)",
                data=buffer,
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
