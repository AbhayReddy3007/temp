"""
score_drugs_gemini25_flash.py

- Reads an Excel with columns (case-insensitive): Name, published_date, abstract
- Concatenates ALL abstracts per Name (sorted by published_date)
- Sends the full concatenated text to Gemini 2.5 Flash (no summarization)
- Applies exact MASH Resolution qualitative rules (as requested)
- Uses numeric thresholds for other endpoints (configurable)
- Saves results to an Excel file

USAGE:
 - Replace GEMINI_API_KEY with your real key/token (hardcoded as requested).
 - Ensure GEMINI_MODEL is "gemini-2.5-flash".
 - Run: python score_drugs_gemini25_flash.py

SECURITY WARNING: Hardcoding secrets in source is insecure. Do not commit this file to a public repo.
"""

import os
import json
import time
import re
from typing import Dict
import pandas as pd
import requests

# -----------------------
# CONFIG (edit these)
# -----------------------
INPUT_EXCEL = "drugs_papers.xlsx"          # path to your Excel file
SHEET_NAME = None                          # None = auto-handle single or multi-sheet
OUTPUT_EXCEL = "drug_scores_output_gemini25flash.xlsx"

# -----------------------
# Gemini config - HARDCODED as requested
# Replace the placeholder string with your real access token or API key.
# If it's an OAuth2 access token (e.g., starts with "ya29."), the script will use Authorization: Bearer <token>.
# If it's a Google API key (starts with "AIza"), the script will append ?key=<API_KEY> to the URL.
# -----------------------
GEMINI_API_KEY = "PASTE_YOUR_REAL_KEY_OR_ACCESS_TOKEN_HERE"
GEMINI_MODEL = "gemini-2.5-flash"

# Endpoints to score (exact strings used in prompt & output)
ENDPOINTS = [
    "Weight loss(%)",
    "A1c reduction(%)",
    "MASH Resolution(%)",
    "ALT Reduction(%)"
]

# Numeric thresholds for endpoints OTHER THAN MASH — modify to your desired cutoffs.
thresholds: Dict[str, Dict[int, float]] = {
    "Weight loss(%)":   {5: 15.0, 4: 10.0, 3: 7.0, 2: 4.0, 1: 0.0},
    "A1c reduction(%)": {5: 1.5,  4: 1.0,  3: 0.7, 2: 0.4, 1: 0.0},
    "ALT Reduction(%)": {5: 40.0, 4: 25.0, 3: 15.0, 2: 7.0, 1: 0.0},
}
MAX_SCORE = 20

# -----------------------
# Helper: robust Excel reading (single or multi-sheet)
# -----------------------
def read_excel_flex(path: str, sheet_name=None, engine="openpyxl") -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    x = pd.read_excel(path, sheet_name=sheet_name, engine=engine)
    if isinstance(x, dict):
        n = len(x)
        if n == 0:
            raise RuntimeError("Excel file contains no sheets.")
        if n == 1:
            df = list(x.values())[0]
            print("Read workbook with 1 sheet.")
        else:
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
    # normalize column names
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df

# -----------------------
# Map and validate required columns case-insensitively
# -----------------------
def map_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    lower_map = { (c.lower() if isinstance(c, str) else c): c for c in df.columns }
    def find_one(candidates):
        for cand in candidates:
            key = cand.lower()
            if key in lower_map:
                return lower_map[key]
        return None
    name_col = find_one(["name", "drug", "drug_name"])
    date_col = find_one(["published_date", "publish_date", "date", "publication_date"])
    abstract_col = find_one(["abstract", "summary", "abstract_text"])
    missing = []
    if name_col is None: missing.append("Name")
    if date_col is None: missing.append("published_date")
    if abstract_col is None: missing.append("abstract")
    if missing:
        raise RuntimeError(f"Missing required column(s): {missing}. Ensure the input has Name, published_date, and abstract (case-insensitive).")
    df = df.rename(columns={name_col: "Name", date_col: "published_date", abstract_col: "abstract"})
    return df

# -----------------------
# Gemini call for gemini-2.5-flash (hardcoded-key friendly)
# -----------------------
def call_gemini(prompt: str, max_tokens: int = 1600, temperature: float = 0.0) -> str:
    if not GEMINI_API_KEY or GEMINI_API_KEY.startswith("PASTE_YOUR"):
        raise RuntimeError("GEMINI_API_KEY not set. Edit the script and paste your API key/token into GEMINI_API_KEY.")
    model_id = GEMINI_MODEL or ""
    if model_id.startswith("models/"):
        model_id = model_id.split("models/", 1)[1]
    base_url = f"https://generativelanguage.googleapis.com/v1beta2/models/{model_id}:generateText"
    headers = {"Content-Type": "application/json"}
    url = base_url
    # Heuristic: Google API keys often start with 'AIza'
    if GEMINI_API_KEY.startswith("AIza"):
        url = base_url + f"?key={GEMINI_API_KEY}"
    else:
        headers["Authorization"] = f"Bearer {GEMINI_API_KEY}"
    body = {
        "prompt": {"text": prompt},
        "temperature": temperature,
        "maxOutputTokens": max_tokens
    }
    resp = requests.post(url, headers=headers, json=body, timeout=240)
    if resp.status_code != 200:
        msg = (
            f"Gemini call failed: status={resp.status_code}\n"
            f"Request URL: {url}\n"
            f"Model id used: {model_id}\n"
            f"Response body:\n{resp.text}\n"
        )
        if resp.status_code == 404:
            msg += (
                "HTTP 404: resource not found. Check that GEMINI_MODEL is 'gemini-2.5-flash' "
                "and that your key/token is valid and has access to this model."
            )
        raise RuntimeError(msg)
    data = resp.json()
    # Extract common shapes
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
    return json.dumps(data)

# -----------------------
# Prompt builder (strict MASH rules & "use only provided abstracts")
# -----------------------
def build_scoring_prompt(drug_name: str, combined_text: str, thresholds: Dict[str, Dict[int, float]]) -> str:
    thr_lines = []
    for ep in ENDPOINTS:
        if ep == "MASH Resolution(%)":
            continue
        if ep not in thresholds:
            raise ValueError(f"Thresholds missing for endpoint: {ep}")
        mapping = thresholds[ep]
        thr_lines.append(f"{ep}: " + ", ".join([f"score {s} if inferred improvement >= {mapping[s]}" for s in sorted(mapping.keys(), reverse=True)]))
    thr_text = "\n".join(thr_lines)
    # MASH rules exactly as requested
    mash_rules = (
        "MASH Resolution qualitative scoring rules (apply these EXACTLY):\n"
        " - 5: >=50% of patients achieved resolution WITH NO worsening of fibrosis.\n"
        " - 4: >=30% of patients achieved resolution WITH NO worsening of fibrosis.\n"
        " - 3: Resolution signal but SOME data indicates worsening of fibrosis.\n"
        " - 2: Mixed or ambiguous data on resolution.\n"
        " - 1: No resolution observed.\n"
    )
    prompt = f"""
You are a clinical research expert. YOU MUST ONLY USE THE TEXT BETWEEN -----BEGIN ABSTRACTS----- AND -----END ABSTRACTS----- BELOW.
Do NOT use the web or any external knowledge. Use every abstract provided; do NOT omit or summarize any abstract.

Drug: "{drug_name}"

TASKS:
1) From the concatenated abstracts below, infer/extract numeric effect sizes or qualitative findings for these endpoints:
   {', '.join(ENDPOINTS)}

2) Apply the MASH Resolution rules EXACTLY as specified:
{mash_rules}

3) For the numeric endpoints, apply these thresholds STRICTLY:
{thr_text}

4) Output EXACTLY a single valid JSON object (no extra text) using this schema:
{{
  "drug": "<drug name>",
  "scores": {{
    "<endpoint name>": {{ "score": <1-5>, "reason": "<one-sentence justification drawn from the abstracts (cite brief fragment or date)>" }},
    ...
  }},
  "total_score": <0-20>,
  "max_score": {MAX_SCORE},
  "confidence": <0.0-1.0>
}}

NOTES:
- Each "reason" must be one short sentence referencing the abstracts (e.g., 'Study [2021-06-15] reported ~32% resolution; no fibrosis worsening noted').
- If you infer a numeric percent for Weight/A1c/ALT, state the percent in the reason and apply numeric thresholds strictly.
- If evidence is limited, still assign a score but set confidence lower and explain in the reason.
- The JSON object must be the only content in the response.

-----BEGIN ABSTRACTS-----
{combined_text}
-----END ABSTRACTS-----
"""
    # trim leading/trailing whitespace and return
    return prompt.strip()

# -----------------------
# JSON extraction helper (resilient)
# -----------------------
def extract_json_from_text(text: str):
    try:
        return json.loads(text)
    except Exception:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            candidate = text[start:end+1]
            try:
                return json.loads(candidate)
            except Exception:
                cand2 = candidate.replace("'", '"')
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
    print("Reading input:", INPUT_EXCEL)
    df = read_excel_flex(INPUT_EXCEL, sheet_name=SHEET_NAME)
    df = map_required_columns(df)
    df["published_date"] = pd.to_datetime(df["published_date"], errors="coerce")
    before_n = len(df)
    df = df.dropna(subset=["Name", "abstract"])
    after_n = len(df)
    if after_n < before_n:
        print(f"Dropped {before_n - after_n} rows missing Name or abstract.")
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
    print(f"Found {len(combined_per_drug)} unique drugs. Beginning scoring...")

    results = []
    for drug, combined_text in combined_per_drug.items():
        print(f"\nScoring: {drug} (chars={len(combined_text)})")
        prompt = build_scoring_prompt(drug, combined_text, thresholds)
        raw_out = None
        for attempt in range(3):
            try:
                raw_out = call_gemini(prompt, max_tokens=2000, temperature=0.0)
                break
            except Exception as e:
                print(f"  Gemini call failed (attempt {attempt+1}): {e}")
                time.sleep(1 + attempt)
        if raw_out is None:
            print(f"  ERROR: failed to get response for {drug}. Storing error placeholder.")
            results.append({"drug": drug, "error": "Gemini call failed after retries", "raw_output": None})
            continue
        parsed = extract_json_from_text(raw_out)
        if parsed is None:
            print("  Warning: Could not parse JSON from model output. Storing raw output for inspection.")
            parsed = {"drug": drug, "raw_output": raw_out, "parse_error": True}
        else:
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
        time.sleep(0.35)

    # Flatten results to DataFrame
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
