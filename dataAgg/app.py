import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ============================================
# Shared helper functions
# ============================================

def extract_numeric(pattern, text):
    match = re.search(pattern, str(text))
    if match:
        return match.group(1)
    return None


# ============================================
# Logic for Upload & Ask (uNa)
# ============================================

def process_excel_una(df):
    cols = df.columns.tolist()
    models = cols[1:-1]

    overall_scores = {"Model": []}
    tts_seconds = {"Model": []}
    upload_seconds = {"Model": []}

    overall_pattern = r"Overall Score:\s*([0-9]+(?:\.[0-9]+)?)\b"
    tts_pattern = r"TTS\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)"
    upload_pattern = r"UT\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)\s*s?\b"

    for model in models:
        col_data = df[model].dropna().astype(str).tolist()

        overall_vals = []
        tts_vals = []
        upload_vals = []

        for cell in col_data:
            if "Overall Score" in cell:
                val = extract_numeric(overall_pattern, cell)
                if val is not None:
                    overall_vals.append(float(val))

            if "TTS" in cell:
                val = extract_numeric(tts_pattern, cell)
                if val is not None:
                    tts_vals.append(float(val))

            if "UT" in cell:
                val = extract_numeric(upload_pattern, cell)
                if val is not None:
                    upload_vals.append(float(val))

        overall_scores["Model"].append(model)
        tts_seconds["Model"].append(model)
        upload_seconds["Model"].append(model)

        for i, v in enumerate(overall_vals, start=1):
            overall_scores.setdefault(str(i), []).append(v)

        for i, v in enumerate(tts_vals, start=1):
            tts_seconds.setdefault(str(i), []).append(v)

        for i, v in enumerate(upload_vals, start=1):
            upload_seconds.setdefault(str(i), []).append(v)

    # Write output to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame(overall_scores).to_excel(writer, sheet_name="overall_scores", index=False)
        pd.DataFrame(tts_seconds).to_excel(writer, sheet_name="tts_seconds", index=False)
        pd.DataFrame(upload_seconds).to_excel(writer, sheet_name="upload_seconds", index=False)

    output.seek(0)
    return output


# ============================================
# Logic for Web Search (ws)
# ============================================

def process_excel_ws(df):
    cols = df.columns.tolist()
    models = cols[1:-1]

    overall_scores = {"Model": []}
    transparency_scores = {"Model": []}
    tts_seconds = {"Model": []}

    overall_pattern = r"Overall Score:\s*([0-9]+(?:\.[0-9]+)?)\b"
    transparency_pattern = r"Transparency Score:\s*([0-9]+(?:\.[0-9]+)?)\b"
    tts_pattern = r"TTS\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)"

    for model in models:
        col_data = df[model].dropna().astype(str).tolist()

        overall_vals = []
        transparency_vals = []
        tts_vals = []

        for cell in col_data:
            if "Overall Score" in cell:
                val = extract_numeric(overall_pattern, cell)
                if val is not None:
                    overall_vals.append(float(val))

            if "Transparency Score" in cell:
                val = extract_numeric(transparency_pattern, cell)
                if val is not None:
                    transparency_vals.append(float(val))

            if "TTS" in cell:
                val = extract_numeric(tts_pattern, cell)
                if val is not None:
                    tts_vals.append(float(val))

        overall_scores["Model"].append(model)
        transparency_scores["Model"].append(model)
        tts_seconds["Model"].append(model)

        for i, v in enumerate(overall_vals, start=1):
            overall_scores.setdefault(str(i), []).append(v)

        for i, v in enumerate(transparency_vals, start=1):
            transparency_scores.setdefault(str(i), []).append(v)

        for i, v in enumerate(tts_vals, start=1):
            tts_seconds.setdefault(str(i), []).append(v)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame(overall_scores).to_excel(writer, sheet_name="overall_scores", index=False)
        pd.DataFrame(transparency_scores).to_excel(writer, sheet_name="transparency_scores", index=False)
        pd.DataFrame(tts_seconds).to_excel(writer, sheet_name="tts_seconds", index=False)

    output.seek(0)
    return output


# ============================================
# Streamlit UI
# ============================================

st.title("Excel Score Extractor")

file_type = st.selectbox(
    "Choose extraction type:",
    ("Upload & Ask (uNa)", "Web Search (ws)")
)

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")

if st.button("Extract Scores"):
    if uploaded_file is None:
        st.error("Please upload a file first.")
    else:
        df = pd.read_excel(uploaded_file)

        if file_type == "Upload & Ask (uNa)":
            output = process_excel_una(df)
            filename = "extracted_results_uNa.xlsx"
        else:
            output = process_excel_ws(df)
            filename = "extracted_results_ws.xlsx"

        st.success("Extraction complete!")

        st.download_button(
            label="📥 Download Extracted Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
