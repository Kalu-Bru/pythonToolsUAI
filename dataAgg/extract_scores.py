import pandas as pd
import re

def extract_numeric(pattern, text):
    match = re.search(pattern, str(text))
    if match:
        return match.group(1)
    return None

def process_excel(input_path, output_path):
    df = pd.read_excel(input_path)

    # modify if columns names are changed
    models = {
        "UAI GPT 5": ("UAI GPT 5", "Screenshot"),
        "UAI GPT 5.1": ("UAI GPT 5.1", "Screenshot.1"),
        "OpenAI GPT 5.1": ("OpenAI GPT 5.1", "Screenshot.2")
    }

    overall_scores = {"Model": []}
    transparency_scores = {"Model": []}
    tts_seconds = {"Model": []}

    # modify if prescript is changed
    overall_pattern = r"Overall Score:\s*([0-9]*\.?[0-9]+)"
    transparency_pattern = r"Transparency Score:\s*([0-9]*\.?[0-9]+)"
    tts_pattern = r"TTS\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)"

    for model, (text_col, screenshot_col) in models.items():
        col_data = df[text_col].dropna().astype(str).tolist()
        overall_vals = []
        transparency_vals = []

        for cell in col_data:
            if "Overall Score" in cell:
                val = extract_numeric(overall_pattern, cell)
                if val is not None:
                    overall_vals.append(float(val))

            if "Transparency Score" in cell:
                val = extract_numeric(transparency_pattern, cell)
                if val is not None:
                    transparency_vals.append(float(val))

        screenshot_data = df[screenshot_col].dropna().astype(str).tolist()
        tts_vals = []

        for cell in screenshot_data:
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

    df_overall = pd.DataFrame(overall_scores)
    df_transparency = pd.DataFrame(transparency_scores)
    df_tts = pd.DataFrame(tts_seconds)

    with pd.ExcelWriter(output_path) as writer:
        df_overall.to_excel(writer, sheet_name="overall_scores", index=False)
        df_transparency.to_excel(writer, sheet_name="transparency_scores", index=False)
        df_tts.to_excel(writer, sheet_name="tts_seconds", index=False)

    print("Extraction complete! Output saved to:", output_path)


# change input file name (Ex_format.xlsx) if needed
# change output file name (extracted_results.xlsx) if needed
if __name__ == "__main__":
    process_excel("Ex_format.xlsx", "extracted_results.xlsx")

