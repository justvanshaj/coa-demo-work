import streamlit as st
from docx import Document
import os
import mammoth
import io
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

# --- Helper: distribute target among variables within min/max using weighted mids ---
def distribute_within_bounds(target, names, mins, maxs, weights):
    """
    Water-filling style distribution:
    - target: total sum to allocate
    - names: list of names (for clarity)
    - mins, maxs: dict(name->min), dict(name->max)
    - weights: dict(name->weight) — initial preference (e.g., midpoints)
    Returns dict(name->value) or raises ValueError if infeasible.
    """
    eps = 1e-9
    # feasibility
    min_sum = sum(mins[n] for n in names)
    max_sum = sum(maxs[n] for n in names)
    if target + eps < min_sum or target - eps > max_sum:
        raise ValueError(f"Target {target} not achievable with given bounds (min_sum={min_sum}, max_sum={max_sum}).")

    # start with weighted allocation proportional to weights
    w_sum = sum(weights[n] for n in names)
    if w_sum <= 0:
        # uniform fallback
        vals = {n: target / len(names) for n in names}
    else:
        vals = {n: target * (weights[n] / w_sum) for n in names}

    # clamp iteratively and redistribute remainder until stable
    locked = {n: False for n in names}
    for _ in range(50):  # finite iterations
        changed = False
        remainder = target - sum(vals[n] for n in names if not locked[n]) - sum(vals[n] for n in names if locked[n])
        # clamp any values outside bounds
        for n in names:
            if not locked[n]:
                if vals[n] < mins[n]:
                    vals[n] = mins[n]
                    locked[n] = True
                    changed = True
                elif vals[n] > maxs[n]:
                    vals[n] = maxs[n]
                    locked[n] = True
                    changed = True
        if not changed:
            # redistribute remainder among unlocked based on weights
            unlocked = [n for n in names if not locked[n]]
            if not unlocked:
                break
            rem = target - sum(vals[n] for n in names)
            if abs(rem) < 1e-8:
                break
            w_un_sum = sum(weights[n] for n in unlocked)
            if w_un_sum <= 0:
                # equal split
                for n in unlocked:
                    vals[n] += rem / len(unlocked)
            else:
                for n in unlocked:
                    vals[n] += rem * (weights[n] / w_un_sum)
        # if nothing changed and remainder small -> ok
        if not changed and abs(target - sum(vals.values())) < 1e-8:
            break

    # final clamp enforce and tiny rounding
    for n in names:
        vals[n] = max(min(vals[n], maxs[n]), mins[n])
        vals[n] = round(vals[n], 2)

    # Final adjustment to match target exactly within small tolerance by nudging unlocked variables
    total = sum(vals[n] for n in names)
    diff = round(target - total, 2)
    if abs(diff) >= 0.01:
        # try to adjust any variable that isn't at a bound
        adjustable = [n for n in names if mins[n] + 1e-9 < vals[n] < maxs[n] - 1e-9]
        for n in adjustable:
            allow_low = round(vals[n] - mins[n], 2)
            allow_high = round(maxs[n] - vals[n], 2)
            move = max(-allow_low, min(allow_high, diff))
            if abs(move) >= 0.01:
                vals[n] = round(vals[n] + move, 2)
                diff = round(target - sum(vals.values()), 2)
                if abs(diff) < 0.01:
                    break
    # final feasibility check
    final_total = round(sum(vals[n] for n in names), 2)
    if abs(final_total - target) > 0.01:
        raise ValueError(f"Could not distribute to meet target. final_total={final_total}, target={target}")
    return vals

# --- Moisture-based automatic component calculation ---
def calculate_components(moisture):
    """
    Auto-calc fat, air, ash, protein, gum using the specified ranges such that:
    moisture + fat + air + ash + protein + gum = 100.0 (within 0.01)
    Ranges:
      fat: 0.45 - 0.55
      air: 2.90 - 3.10
      ash: 0.45 - 0.55
      protein: 2.45 - 2.55
      gum: 80.10 - 89.95
    """
    # validate moisture
    if moisture < 0 or moisture > 100:
        raise ValueError("Moisture must be between 0 and 100")

    remaining = round(100.0 - float(moisture), 4)

    # set ranges
    ranges = {
        "fat": (0.45, 0.55),
        "air": (2.90, 3.10),
        "ash": (0.45, 0.55),
        "protein": (2.45, 2.55),
        "gum": (80.10, 89.95)
    }

    # midpoints
    mids = {k: round((v[0] + v[1]) / 2.0, 4) for k, v in ranges.items()}

    # start with mids for non-gum
    others = ["fat", "air", "ash", "protein"]
    others_mid_sum = sum(mids[o] for o in others)

    # ideal gum to meet remaining
    gum_needed = round(remaining - others_mid_sum, 4)

    gum_min, gum_max = ranges["gum"]
    # if gum_needed inside gum range -> done (use mids for others and gum_needed)
    if gum_min - 1e-9 <= gum_needed <= gum_max + 1e-9:
        fat = round(mids["fat"], 2)
        air = round(mids["air"], 2)
        ash = round(mids["ash"], 2)
        protein = round(mids["protein"], 2)
        gum = round(gum_needed, 2)
        # final rounding adjust small floating error
        total = round(moisture + fat + air + ash + protein + gum, 2)
        if abs(total - 100.0) > 0.01:
            # tiny fix: adjust gum by the residual
            gum = round(gum + (100.0 - total), 2)
        return gum, protein, ash, air, fat

    # otherwise clamp gum to closest bound and distribute remaining among others
    # try with gum at mid then clamp:
    # We'll try both extremes of gum (min, max) to find feasible distribution for others.
    gum_try_list = []
    # prefer gum that is closest to needed
    if gum_needed < gum_min:
        gum_try_list = [gum_min, gum_max]
    else:
        gum_try_list = [gum_max, gum_min]

    for gum_try in gum_try_list:
        available_for_others = round(remaining - gum_try, 4)
        # prepare mins and maxs for others
        mins = {o: ranges[o][0] for o in others}
        maxs = {o: ranges[o][1] for o in others}
        # weights from midpoints
        weights = {o: mids[o] for o in others}
        try:
            allocated = distribute_within_bounds(available_for_others, others, mins, maxs, weights)
            # success
            fat = allocated["fat"]
            air = allocated["air"]
            ash = allocated["ash"]
            protein = allocated["protein"]
            gum = round(gum_try, 2)
            # final total check and tiny fix
            total = round(moisture + fat + air + ash + protein + gum, 2)
            if abs(total - 100.0) > 0.01:
                # apply tiny residual to gum if possible
                residual = round(100.0 - total, 2)
                new_gum = round(gum + residual, 2)
                if ranges["gum"][0] <= new_gum <= ranges["gum"][1]:
                    gum = new_gum
                    total = round(moisture + fat + air + ash + protein + gum, 2)
            if abs(total - 100.0) <= 0.01:
                return gum, protein, ash, air, fat
            # else try next gum_try
        except ValueError:
            continue

    # if both attempts failed, it's infeasible
    raise ValueError(
        f"Unable to compute components that satisfy ranges for moisture={moisture}. "
        f"Remaining={remaining}. Please check moisture value."
    )


# --- Streamlit App Starts ---
st.set_page_config(page_title="COA Generator", layout="wide")
st.title("🧪 COA Document Generator (Auto Components)")

with st.form("coa_form"):
    code = st.selectbox(
        "Select Product Code Range",
        [f"{i}-{i+500}" for i in range(500, 10001, 500)]
    )
    st.info(f"📄 Using template: COA {code}.docx")

    date = st.text_input("Date (e.g., July 2025)")

    # Auto-calculate Best Before (robust parsing)
    best_before = ""
    if date:
        parsed = None
        for fmt_try in ("%B %Y", "%b %Y", "%B, %Y", "%b, %Y"):
            try:
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
    moisture = st.number_input("Moisture (%)", min_value=0.0, max_value=99.0, step=0.01, value=10.0)

    st.markdown(
        "Component values (Fat, Air, Ash, Protein, Gum) will be **automatically calculated** "
        "to meet these constraints and sum to 100% with Moisture."
    )

    ph = st.text_input("pH Level (e.g., 6.7)")
    mesh_200 = st.text_input("200 Mesh (%)")
    viscosity_2h = st.text_input("Viscosity After 2 Hours (CPS)")
    viscosity_24h = st.text_input("Viscosity After 24 Hours (CPS)")
    submitted = st.form_submit_button("Generate COA")

if submitted:
    template_path = f"COA {code}.docx"
    output_path = "generated_coa.docx"

    # compute components automatically
    try:
        gum, protein, ash, air, fat = calculate_components(moisture)
    except Exception as e:
        st.error(f"Component calculation failed: {e}")
        st.stop()

    # show summary
    try:
        import pandas as pd
        summary_df = pd.DataFrame({
            "Component": ["Moisture", "Fat", "Air", "Ash", "Protein", "Gum"],
            "Value (%)": [round(moisture, 2), round(fat,2), round(air,2), round(ash,2), round(protein,2), round(gum,2)]
        })
        st.subheader("Composition summary (auto-calculated)")
        st.dataframe(summary_df, width=600)
        total = round(moisture + fat + air + ash + protein + gum, 2)
        st.info(f"Total = {total}%")
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
        "GUM_CONTENT": f"{round(gum,2)}%",
        "PROTEIN": f"{round(protein,2)}%",
        "ASH_CONTENT": f"{round(ash,2)}%",
        "AIR": f"{round(air,2)}%",
        "FAT": f"{round(fat,2)}%"
    }

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
