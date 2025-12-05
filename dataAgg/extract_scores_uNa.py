import pandas as pd
import re

def extract_numeric(pattern, text):
    match = re.search(pattern, str(text))
    if match:
        return match.group(1)
    return None

def normalize_lengths(d):
    max_len = max(len(v) for v in d.values())
    for k, v in d.items():
        while len(v) < max_len:
            v.append(None)

def process_excel(input_path, output_path):
    df = pd.read_excel(input_path)
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

    print(len(overall_scores), len(tts_seconds), len(upload_seconds))
    normalize_lengths(overall_scores)
    normalize_lengths(tts_seconds)
    normalize_lengths(upload_seconds)

    df_overall = pd.DataFrame(overall_scores)
    df_tts = pd.DataFrame(tts_seconds)
    df_upload = pd.DataFrame(upload_seconds)

    with pd.ExcelWriter(output_path) as writer:
        df_overall.to_excel(writer, sheet_name="overall_scores", index=False)
        df_tts.to_excel(writer, sheet_name="tts_seconds", index=False)
        df_upload.to_excel(writer, sheet_name="upload_seconds", index=False)

    print("Extraction complete! Output saved to:", output_path)


if __name__ == "__main__":
    process_excel("Ex_format_u&a.xlsx", "extracted_results_u&a.xlsx")
