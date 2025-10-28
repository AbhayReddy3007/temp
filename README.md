"""
score_drugs_full_abstracts.py

- Reads an Excel with columns (case-insensitive): Name, published_date, abstract
- Concatenates ALL abstracts per Name (sorted by published_date)
- Sends the full concatenated text to Gemini for scoring (no summarization)
- Applies exact MASH Resolution qualitative rules (user provided)
- Uses numeric thresholds for other endpoints (configurable)
- Saves results to an Excel file

Requirements:
    pip install pandas requests openpyxl

Notes:
- This sends the full concatenated abstracts to the LLM. If your concatenated texts are extremely large,
  your Gemini model may reject due to token limits. You explicitly requested "do not summarize",
  so this script does not summarize or chunk — it will attempt the full text.
- Adjust GEMINI endpoint/authentication if you use a different interface.
"""

import os
import json
import time
import zipfile
from typing import Dict, Any
import pandas as pd
import requests

# -----------------------
# USER CONFIG
# -----------------------
INPUT_EXCEL = "drugs_papers.xlsx"          # <-- change to your file path
SHEET_NAME = None                          # None = auto-handle (single or multi-sheet)
OUTPUT_EXCEL = "drug_scores_output_full.xlsx"

# Gemini API config: prefer storing API key in environment variable
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
# Example model path — change to your model id / path if needed
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "models/gemini-1.0")

# Endpoints to score (exact strings used in prompt and output)
ENDPOINTS = [
    "Weight loss(%)",
    "A1c reduction(%)",
    "MASH Resolution(%)",
    "ALT Reduction(%)"
]

# Numeric thresholds for endpoints OTHER THAN MASH (edit as desired)
# Format: { "<endpoint>": {5: value_for_5, 4: value_for_4, ... } }
thresholds: Dict[str, Dict[int, float]] = {
    "Weight loss(%)":   {5: 15.0, 4: 10.0, 3: 7.0, 2: 4.0, 1: 0.0},
    "A1c reduction(%)": {5: 1.5,  4: 1.0,  3: 0.7, 2: 0.4, 1: 0.0},
    # MASH Resolution is handled qualitatively below (do not include thresholds for it)
    "ALT Reduction(%)": {5: 40.0, 4: 25.0, 3: 15.0, 2: 7.0, 1: 0.0},
}
MAX_SCORE = 20

# -----------------------
# Utilities: robust Excel reading
# -----------------------
def read_excel_flex(path: str, sheet_name=None, engine="openpyxl") -> pd.DataFrame:
    """
    Read an Excel workbook robustly:
      - If sheet_name is None, handles single- or multi-sheet files.
      - If multiple sheets exist, concatenates them (checks columns).
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    # If user provided explicit sheet name or index, pass it directly
    x = pd.read_excel(path, sheet_name=sheet_name, engine=engine)

    if isinstance(x, dict):
        n = len(x)
        if n == 0:
            raise RuntimeError("Excel file has no sheets.")
        if n == 1:
            df = list(x.values())[0]
            print("Read workbook with 1 sheet.")
        else:
            # decide how to concat: check whether columns are identical across sheets
            col_sets = [set([c.strip() if isinstance(c, str) else c for c in df.columns]) for df in x.values()]
            all_same = all(col_sets[0] == s for s in col_sets)
            if all_same:
                df = pd.concat(x.values(), ignore_index=True)
                print(f"Concatenated {n} sheets (identical columns).")
            else:
                df = pd.concat(x.values(), ignore_index=True, sort=False)
                print(f"Concatenated {n} sheets (columns differed; missing values filled with NaN).")
    else:
        df = x
        print("Read single-sheet Excel (or sheet selected explicitly).")

    # normalize column names (strip)
    new_cols = []
    for c in df.columns:
        if isinstance(c, str):
            new_cols.append(c.strip())
        else:
            new_cols.append(c)
    df.columns = new_cols
    return df

# -----------------------
# Normalize and validate columns (case-insensitive)
# -----------------------
def map_required_columns(df: pd.DataFrame):
    """
    Ensure the DataFrame has Name, published_date, abstract columns (case-insensitive).
    Returns a DataFrame with columns renamed to exactly: Name, published_date, abstract
    """
    col_map = {}
    lower_to_col = { (c.lower() if isinstance(c, str) else c): c for c in df.columns }

    def find_col(candidate_names):
        for cand in candidate_names:
            key = cand.lower()
            if key in lower_to_col:
                return lower_to_col[key]
        return None

    # possible variants for each required column
    name_col = find_col(["name", "drug", "drug_name", "Name"])
    date_col = find_col(["published_date", "publish_date", "date", "published", "publication_date"])
    abstract_col = find_col(["abstract", "summary", "abstract_text", "Abstract"])

    missing = []
    if name_col is None:
        missing.append("Name")
    if date_col is None:
        missing.append("published_date")
    if abstract_col is None:
        missing.append("abstract")
    if missing:
        raise RuntimeError(f"Missing required column(s) in input file: {missing}. "
                           "Ensure your file has Name, published_date, and abstract columns (case-insensitive).")

    # rename to standardized names
    df = df.rename(columns={name_col: "Name", date_col: "published_date", abstract_col: "abstract"})
    return df

# -----------------------
# Gemini call (generic REST example)
# -----------------------
def call_gemini(prompt: str, max_tokens: int = 1200, temperature: float = 0.0) -> str:
    """
    Call Gemini-like REST endpoint. Edit for your provider if needed.
    Returns the model's text output.
    """
    if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
        raise RuntimeError("Set GEMINI_API_KEY environment variable or update GEMINI_API_KEY in the script.")

    # This URL pattern is illustrative. Replace if your deployment uses different path.
    url = f"https://generativelanguage.googleapis.com/v1beta2/{GEMINI_MODEL}:generateText"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GEMINI_API_KEY}"
    }
    body = {
        "prompt": {"text": prompt},
        "temperature": temperature,
        "maxOutputTokens": max_tokens
    }

    resp = requests.post(url, headers=headers, json=body, timeout=180)
    if resp.status_code != 200:
        raise RuntimeError(f"Gemini API error {resp.status_code}: {resp.text}")
    data = resp.json()

    # Try a few expected structures
    if isinstance(data, dict):
        if "candidates" in data and isinstance(data["candidates"], list) and len(data["candidates"]) > 0:
            return data["candidates"][0].get("content", "")
        if "output" in data and isinstance(data["output"], list):
            pieces = []
            for item in data["output"]:
                for content in item.get("content", []):
                    if "text" in content:
                        pieces.append(content["text"])
            if pieces:
                return "\n".join(pieces)
    # fallback
    return json.dumps(data)

# -----------------------
# Build scoring prompt (strict: use only provided abstracts)
# -----------------------
def build_scoring_prompt(drug_name: str, combined_text: str, thresholds: Dict[str, Dict[int, float]]) -> str:
    # Build numeric threshold lines for non-MASH endpoints
    thr_lines = []
    for ep in ENDPOINTS:
        if ep == "MASH Resolution(%)":
            continue
        if ep not in thresholds:
            raise ValueError(f"Thresholds missing for endpoint: {ep}")
        mapping = thresholds[ep]
        thr_lines.append(f"{ep}: " + ", ".join([f"score {s} if inferred improvement >= {mapping[s]}" for s in sorted(mapping.keys(), reverse=True)]))
    thr_text = "\n".join(thr_lines)

    # Exact MASH rules as requested
    mash_rules = (
        "MASH Resolution qualitative scoring rules (apply these EXACTLY):\n"
        " - 5: Evidence shows >=50% of patients achieved resolution with NO worsening of fibrosis.\n"
        " - 4: Evidence shows >=30% of patients achieved resolution with NO worsening of fibrosis.\n"
        " - 3: There is a resolution signal but some data indicates worsening of fibrosis in some patients.\n"
        " - 2: Mixed or ambiguous data regarding resolution (conflicting findings, low-quality evidence, or inconsistent results).\n"
        " - 1: No resolution observed in the provided abstracts.\n"
    )

    prompt = f"""
You are an expert clinical researcher. YOU MUST ONLY use the text between -----BEGIN ABSTRACTS----- and -----END ABSTRACTS----- below.
Do NOT use the web, external knowledge, or any information outside the provided abstracts. Use each and every abstract below; do NOT omit any study.

Drug: "{drug_name}"

TASK:
1) From the provided concatenated abstracts, extract evidence (and numeric estimates when present) for the following endpoints:
   {', '.join(ENDPOINTS)}
2) Apply the MASH Resolution rules EXACTLY as specified:
{mash_rules}

3) For the other endpoints, apply these numeric thresholds STRICTLY:
{thr_text}

4) Output EXACTLY valid JSON (no extra commentary) with this schema:
{{
  "drug": "<drug name>",
  "scores": {{
    "<endpoint name>": {{ "score": <1-5>, "reason": "<one-sentence justification drawn from the abstracts (cite brief fragment or date)>" }},
    ...
  }},
  "total_score": <0-20>,
  "max_score": {MAX_SCORE},
  "confidence": <0.0-1.0>   // decimal indicating confidence based ONLY on provided abstracts
}}

NOTES:
- Each "reason" must be one short sentence referencing the abstracts (e.g., 'Study [2021-06-15] reported ~32% resolution; no fibrosis worsening reported').
- If numeric percent is inferred for Weight/A1c/ALT, state the percent in the reason and apply numeric thresholds strictly.
- If evidence is limited, still assign a score but set confidence lower and explain brevity in the reason.
- The JSON object must be the only content in the model response.

-----BEGIN ABSTRACTS-----
{combined_text}
-----END ABSTRACTS-----
"""
    return prompt.strip()

# -----------------------
# Robust JSON extraction helper
# -----------------------
def extract_json_from_text(text: str):
    """Try to extract a JSON object substring from a text blob and parse it."""
    try:
        return json.loads(text)
    except Exception:
        # find the first '{' and the last '}' and try parsing that slice
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            candidate = text[start:end+1]
            try:
                return json.loads(candidate)
            except Exception:
                # last-ditch: try to fix common LLM mistakes (replace single quotes, remove trailing commas)
                cand2 = candidate.replace("'", '"')
                import re
                cand2 = re.sub(r",\s*}", "}", cand2)
                cand2 = re.sub(r",\s*]", "]", cand2)
                try:
                    return json.loads(cand2)
                except Exception:
                    return None
        return None

# -----------------------
# Main pipeline
# -----------------------
def main():
    print("Reading input Excel:", INPUT_EXCEL)
    df = read_excel_flex(INPUT_EXCEL, sheet_name=SHEET_NAME)
    df = map_required_columns(df)

    # ensure published_date parsed
    df["published_date"] = pd.to_datetime(df["published_date"], errors="coerce")

    # drop rows missing Name or abstract
    before_n = len(df)
    df = df.dropna(subset=["Name", "abstract"])
    after_n = len(df)
    if after_n < before_n:
        print(f"Dropped {before_n - after_n} rows missing Name or abstract.")

    # sort by published_date then group by Name
    grouped = df.sort_values("published_date").groupby("Name", sort=True)

    combined_per_drug = {}
    for name, group in grouped:
        parts = []
        for _, row in group.iterrows():
            date_str = ""
            if pd.notna(row["published_date"]):
                date_str = str(row["published_date"].date())
            parts.append(f"[{date_str}] {row['abstract']}")
        combined_text = "\n\n---\n\n".join(parts)
        combined_per_drug[name] = combined_text

    print(f"Found {len(combined_per_drug)} unique drugs. Beginning model scoring (this may take a while).")

    results = []
    for drug, combined_text in combined_per_drug.items():
        print(f"\nScoring drug: {drug}  (chars: {len(combined_text)})")
        # Build prompt (send full concatenated abstracts)
        prompt = build_scoring_prompt(drug, combined_text, thresholds)

        # Call Gemini with small retry loop
        raw_out = None
        for attempt in range(3):
            try:
                raw_out = call_gemini(prompt, max_tokens=1600, temperature=0.0)
                break
            except Exception as e:
                print(f"  Gemini call failed (attempt {attempt+1}): {e}")
                time.sleep(1 + attempt)
        if raw_out is None:
            print(f"  ERROR: failed to get response from Gemini for {drug}. Saving error placeholder.")
            results.append({
                "drug": drug,
                "error": "Gemini call failed after retries",
                "raw_output": None
            })
            continue

        # Try to parse JSON
        parsed = extract_json_from_text(raw_out)
        if parsed is None:
            print("  Warning: Could not parse JSON directly from model output. Storing raw output for manual inspection.")
            parsed = {
                "drug": drug,
                "raw_output": raw_out,
                "parse_error": True
            }
        else:
            # ensure total present: if not, compute from endpoint scores if possible
            if parsed.get("total_score") is None:
                s = 0
                ok = True
                for ep in ENDPOINTS:
                    try:
                        sc = parsed.get("scores", {}).get(ep, {}).get("score")
                        if isinstance(sc, (int, float)):
                            s += int(sc)
                        else:
                            ok = False
                    except Exception:
                        ok = False
                if ok:
                    parsed["total_score"] = s
            parsed.setdefault("max_score", MAX_SCORE)

        parsed["_char_len"] = len(combined_text)
        parsed["_raw_output"] = raw_out
        results.append(parsed)

        # courteous pause (avoid hammering the API)
        time.sleep(0.35)

    # Build output DataFrame
    rows = []
    for r in results:
        if not isinstance(r, dict):
            continue
        base = {
            "drug": r.get("drug"),
            "total_score": r.get("total_score"),
            "max_score": r.get("max_score", MAX_SCORE),
            "confidence": r.get("confidence"),
            "_char_len": r.get("_char_len"),
            "raw_output_excerpt": (r.get("_raw_output")[:1000] if r.get("_raw_output") else "")
        }
        scores = r.get("scores", {}) if isinstance(r.get("scores"), dict) else {}
        for ep in ENDPOINTS:
            ep_obj = scores.get(ep, {}) or {}
            base[f"{ep} - score"] = ep_obj.get("score")
            base[f"{ep} - reason"] = ep_obj.get("reason")
        rows.append(base)

    out_df = pd.DataFrame(rows)
    out_df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\nDone. Results saved to: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()

