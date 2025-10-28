"""
gemini_drug_scoring_full_abstracts.py

- Reads an Excel with columns: Name, published_date, abstract
- Combines ALL abstracts per drug (no summarization / no omission)
- Sends full concatenated abstracts to Gemini with strict instructions:
    * Use only the provided abstracts (no web)
    * Follow the MASH Resolution qualitative rules exactly
    * Apply numeric thresholds for other endpoints (user-provided)
    * Return EXACTLY valid JSON (schema described)
- Saves results to an output Excel.

Requirements:
pip install pandas requests openpyxl

Usage:
- Edit INPUT_EXCEL and thresholds below
- Set GEMINI_API_KEY and GEMINI_MODEL in environment, or edit the variables
- Run: python gemini_drug_scoring_full_abstracts.py
"""

import os
import json
import time
from typing import Dict
import pandas as pd
import requests

# -----------------------
# USER CONFIG
# -----------------------
INPUT_EXCEL = "drugs_papers.xlsx"    # your input file
SHEET_NAME = None                    # if needed
OUTPUT_EXCEL = "drug_scores_output_full_abstracts.xlsx"

# Gemini API config - set env vars or replace here (not recommended)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "models/gemini-1.0")  # adjust per your account

# endpoints (must match exactly in prompt & output)
ENDPOINTS = [
    "Weight loss(%)",
    "A1c reduction(%)",
    "MASH Resolution(%)",
    "ALT Reduction(%)"
]

# Numeric thresholds for endpoints OTHER THAN MASH (edit to your desired cutoffs).
# Keys must match ENDPOINTS exactly, but MASH is handled qualitatively and ignored here.
thresholds: Dict[str, Dict[int, float]] = {
    "Weight loss(%)":     {5: 15.0, 4: 10.0, 3: 7.0, 2: 4.0, 1: 0.0},
    "A1c reduction(%)":   {5: 1.5,  4: 1.0,  3: 0.7, 2: 0.4, 1: 0.0},
    # MASH Resolution is handled by qualitative rules — do not include thresholds here for it
    "ALT Reduction(%)":   {5: 40.0, 4: 25.0, 3: 15.0, 2: 7.0, 1: 0.0}
}
MAX_SCORE = 20

# -----------------------
# Gemini call (simple REST example)
# Adjust if your Gemini usage differs (client library, different URL, auth)
# -----------------------
def call_gemini(prompt: str, max_tokens: int = 1200, temperature: float = 0.0) -> str:
    if GEMINI_API_KEY is None or GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
        raise RuntimeError("Set GEMINI_API_KEY environment variable or update GEMINI_API_KEY in script.")
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
    # common shapes handled:
    if "candidates" in data and isinstance(data["candidates"], list):
        return data["candidates"][0].get("content", "")
    if "output" in data and isinstance(data["output"], list):
        pieces = []
        for item in data["output"]:
            for c in item.get("content", []):
                if "text" in c:
                    pieces.append(c["text"])
        return "\n".join(pieces)
    return json.dumps(data)

# -----------------------
# Build prompt (explicit MASH rules + "use every abstract" directive)
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

    mash_rules = (
        "MASH Resolution qualitative scoring rules (apply these EXACTLY):\n"
        " - 5: Evidence shows >=50% of patients achieved NASH resolution WITH NO worsening of fibrosis.\n"
        " - 4: Evidence shows >=30% of patients achieved NASH resolution WITH NO worsening of fibrosis.\n"
        " - 3: There is a resolution signal but some data indicates worsening of fibrosis in some patients.\n"
        " - 2: Mixed or ambiguous data regarding resolution (conflicting findings, low-quality evidence, or inconsistent results).\n"
        " - 1: No resolution observed in the provided abstracts."
    )

    prompt = f"""
You are a clinical research expert. You WILL ONLY use the text between -----BEGIN ABSTRACTS----- and -----END ABSTRACTS----- below to draw conclusions.
Do NOT use the web, any external sources, or your world knowledge beyond what is present in the abstracts. Use each and every abstract provided — do NOT omit any abstract or make assumptions based on outside knowledge.

Drug: "{drug_name}"

TASKS:
1) From the concatenated abstracts below, extract evidence and (where present) approximate numeric effect sizes or descriptive findings for each of these endpoints:
   {', '.join(ENDPOINTS)}.
   For each endpoint, the model should base conclusions only on the text provided and must cite brief supporting text from the abstracts in the "reason" (one short sentence).
2) APPLY the MASH Resolution qualitative rules exactly (below).
{mash_rules}

3) For the other endpoints, use the numeric scoring thresholds provided here:
{thr_text}

4) Return EXACTLY valid JSON (no extra text) in the schema:
{{
  "drug": "<drug name>",
  "scores": {{
    "<endpoint name>": {{ "score": <1-5>, "reason": "<one-sentence justification drawn from the abstracts>" }},
    ...
  }},
  "total_score": <0-20>,
  "max_score": 20,
  "confidence": <0.0-1.0>    // decimal expressing model's confidence based only on the material provided
}}

NOTES & RULES:
- Use only the provided abstracts. Do NOT search the web or rely on external facts.
- Use each abstract's content to support statements; include a short textual cue in each reason like: "Study X: [YYYY-MM-DD] reported ..." or quote short fragment (<=20 words) from the abstracts if present.
- For numeric endpoints (Weight loss, A1c, ALT), if you infer a numeric percent, state the inferred percent in the reason and apply the numeric thresholds strictly (choose highest score whose threshold is satisfied).
- For MASH Resolution, obey the qualitative rules above (choose the one that best matches the evidence).
- If evidence is limited or ambiguous for an endpoint, still assign a score but set confidence lower and explain why in the reason.
- The JSON output must be the only text in the response.

-----BEGIN ABSTRACTS-----
{combined_text}
-----END ABSTRACTS-----
"""
    return prompt.strip()

# -----------------------
# Main pipeline (no summarization; uses full concatenated abstracts)
# -----------------------
def main():
    # 1. Read Excel
    df = pd.read_excel(INPUT_EXCEL, sheet_name=SHEET_NAME, engine="openpyxl")
    df = df.rename(columns={c: c.strip() for c in df.columns})
    required = {"Name", "published_date", "abstract"}
    if not required.issubset(set(df.columns)):
        raise RuntimeError(f"Input must contain columns: {required}")
    df["published_date"] = pd.to_datetime(df["published_date"], errors="coerce")
    df = df.dropna(subset=["Name", "abstract"])

    # 2. Combine all abstracts per drug (sorted by date)
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

    results = []
    for drug, combined_text in combined_per_drug.items():
        print(f"Processing: {drug}  (chars={len(combined_text)})")
        # build prompt (we intentionally send the full combined_text)
        prompt = build_scoring_prompt(drug, combined_text, thresholds)
        # call Gemini (retry small number of times)
        raw_out = None
        for attempt in range(3):
            try:
                raw_out = call_gemini(prompt, max_tokens=1200, temperature=0.0)
                break
            except Exception as e:
                print(f"  Gemini call failed (attempt {attempt+1}): {e}")
                time.sleep(1 + attempt * 1.0)
        if raw_out is None:
            print(f"  Failed for {drug}; saving error.")
            results.append({
                "drug": drug,
                "error": "Gemini call failed after retries",
                "raw_output": None
            })
            continue

        # Try parsing JSON (model was asked to output JSON only)
        parsed = None
        try:
            parsed = json.loads(raw_out)
        except Exception:
            # try to extract JSON substring
            try:
                start = raw_out.find("{")
                end = raw_out.rfind("}") + 1
                parsed = json.loads(raw_out[start:end])
            except Exception:
                parsed = {"drug": drug, "raw_output": raw_out, "parse_error": True}

        # If total_score missing, compute from endpoint scores if possible
        if isinstance(parsed, dict) and parsed.get("total_score") is None:
            total = 0
            ok = True
            for ep in ENDPOINTS:
                sc = None
                try:
                    sc = parsed.get("scores", {}).get(ep, {}).get("score")
                except Exception:
                    sc = None
                if isinstance(sc, (int, float)):
                    total += int(sc)
                else:
                    ok = False
            if ok:
                parsed["total_score"] = total
            else:
                parsed["total_score"] = parsed.get("total_score")  # keep as-is (maybe None)

        parsed["_char_len"] = len(combined_text)
        results.append(parsed)
        time.sleep(0.35)

    # 4. Flatten results to dataframe and save
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
            "raw_output": r.get("raw_output", "") or (r if isinstance(r, str) else "")
        }
        scores = r.get("scores", {}) if isinstance(r.get("scores"), dict) else {}
        for ep in ENDPOINTS:
            ep_obj = scores.get(ep, {}) or {}
            base[f"{ep} - score"] = ep_obj.get("score")
            base[f"{ep} - reason"] = ep_obj.get("reason")
        rows.append(base)

    out_df = pd.DataFrame(rows)
    out_df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\nDone. Results saved to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
