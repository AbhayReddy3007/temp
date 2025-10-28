"""
gemini_drug_scoring.py

Requirements:
- pandas
- requests
- openpyxl (for pandas to read/write Excel)

Install: pip install pandas requests openpyxl
"""

import os
import json
import math
import time
import requests
import pandas as pd
from typing import Dict, Any, List

# -----------------------
# USER CONFIG
# -----------------------
INPUT_EXCEL = "drugs_papers.xlsx"   # path to your Excel
SHEET_NAME = None                   # or sheet name if needed
OUTPUT_EXCEL = "drug_scores_output.xlsx"

# Gemini API config (set as environment variables for safety)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "models/gemini-1.0")  # change to your model id

# Local limits used to decide whether to summarize first
CHARS_SUMMARY_THRESHOLD = 18000   # if concatenated abstracts exceed this, summarize first
CHUNK_CHARS = 8000                # when summarizing, chunk size to feed the model

# endpoints to score (exact names used later in prompt)
ENDPOINTS = [
    "Weight loss(%)",
    "A1c reduction(%)",
    "MASH Resolution(%)",
    "ALT Reduction(%)"
]

# Example thresholds: YOU SHOULD EDIT THESE to your exact "when to give 5/4/3/2/1"
# numeric meaning: minimum % (or absolute) improvement required for that score
# Keys must match ENDPOINTS strings exactly.
thresholds: Dict[str, Dict[int, float]] = {
    "Weight loss(%)":     {5: 15.0, 4: 10.0, 3: 7.0, 2: 4.0, 1: 0.0},
    "A1c reduction(%)":   {5: 1.5,  4: 1.0,  3: 0.7, 2: 0.4, 1: 0.0},
    "MASH Resolution(%)": {5: 60.0, 4: 45.0, 3: 30.0,2: 15.0,1: 0.0},
    "ALT Reduction(%)":   {5: 40.0, 4: 25.0, 3: 15.0,2: 7.0, 1: 0.0}
}
MAX_SCORE = 20


# -----------------------
# Helper: Gemini call
# -----------------------
def call_gemini(prompt: str, max_tokens: int = 1200, temperature: float = 0.0) -> str:
    """
    Call Gemini (Generative Language) REST API (example).
    Adjust endpoint/model name if your provider is different.
    Returns the model's text output (string).
    NOTE: This example uses the v1beta2 generateText endpoint style.
    """
    if GEMINI_API_KEY is None or GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
        raise RuntimeError("Please set GEMINI_API_KEY environment variable or update GEMINI_API_KEY in the script.")

    # Example Google Generative API endpoint (may require different auth depending on your setup)
    # If you use a cloud client or a different endpoint, replace this block with the proper call.
    url = f"https://generativelanguage.googleapis.com/v1beta2/{GEMINI_MODEL}:generateText"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GEMINI_API_KEY}"
    }
    body = {
        "prompt": {
            "text": prompt
        },
        "temperature": temperature,
        "maxOutputTokens": max_tokens
    }

    resp = requests.post(url, headers=headers, json=body, timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(f"Gemini API error {resp.status_code}: {resp.text}")
    data = resp.json()
    # structure depends on API; try a couple of known patterns
    if "candidates" in data and isinstance(data["candidates"], list):
        return data["candidates"][0].get("content", "")
    if "output" in data and isinstance(data["output"], list):
        # Some responses place text under output[0].content[0].text
        pieces = []
        for item in data["output"]:
            if "content" in item and isinstance(item["content"], list):
                for c in item["content"]:
                    if "text" in c:
                        pieces.append(c["text"])
        return "\n".join(pieces)
    # fallback: stringify entire json
    return json.dumps(data, indent=2)


# -----------------------
# Helper: summarizer (simple two-step)
# -----------------------
def summarize_long_text(long_text: str) -> str:
    """
    If text is too long, chunk and ask the model to summarize each chunk,
    then combine chunk summaries and ask for a final concise summary.
    Returns the final summary string.
    """
    # naive chunking by characters
    chunks = [long_text[i:i+CHUNK_CHARS] for i in range(0, len(long_text), CHUNK_CHARS)]
    chunk_summaries = []
    for i, ch in enumerate(chunks):
        prompt = (
            f"You are an expert scientific summarizer. Produce a concise (3-5 sentence) summary of this chunk "
            f"keeping key study results, effect sizes, and any numerical endpoints. Output only the summary.\n\n"
            f"Chunk {i+1}/{len(chunks)}:\n\n{ch}"
        )
        # small token budget for chunk summary
        text = call_gemini(prompt, max_tokens=400, temperature=0.0)
        chunk_summaries.append(text.strip())
        time.sleep(0.35)  # mild rate-limit spacing

    combined = "\n\n".join(chunk_summaries)
    final_prompt = (
        "You are an expert scientific summarizer. Combine the following chunk summaries into one coherent "
        "concise summary (max 8 sentences). Keep numeric outcomes and effect sizes, include study count if possible.\n\n"
        f"{combined}"
    )
    final_summary = call_gemini(final_prompt, max_tokens=600, temperature=0.0)
    return final_summary.strip()


# -----------------------
# Build scoring prompt
# -----------------------
def build_scoring_prompt(drug_name: str, combined_text: str, thresholds: Dict[str, Dict[int, float]]) -> str:
    """
    Build a deterministic prompt instructing the LLM to:
      - apply thresholds
      - return machine-readable JSON
      - include per-endpoint short justification
      - include total score and confidence
    """
    thr_lines = []
    for ep in ENDPOINTS:
        if ep not in thresholds:
            raise ValueError(f"Thresholds missing for endpoint: {ep}")
        mapping = thresholds[ep]
        thr_lines.append(f"{ep}: " + ", ".join([f"score {s} if >= {mapping[s]}" for s in sorted(mapping.keys(), reverse=True)]))

    thr_text = "\n".join(thr_lines)

    prompt = f"""
You are a clinical research expert. Below are concatenated abstracts and metadata for the drug "{drug_name}" (many papers).
Your tasks:
1) From the text below, infer expected evidence-based approximate numeric effect sizes and/or qualitative evidence for each of these endpoints:
   {', '.join(ENDPOINTS)}.
   You may state numbers (e.g., "weight loss ~12% (pooled from four RCTs)") if supported by the abstracts, or give best-effort estimates with explanation if exact numbers are not present.

2) Using the following scoring thresholds, assign a score 1-5 for each endpoint.
{thr_text}

3) Output EXACTLY valid JSON (no extra text) with the following schema:
{{
  "drug": "<drug name>",
  "scores": {{
    "<endpoint name>": {{ "score": <1-5>, "reason": "<one-sentence justification, include numbers if available>" }},
    ...
  }},
  "total_score": <0-20>,
  "max_score": 20,
  "confidence": <0.0-1.0>   // decimal expressing how confident you are about the scoring
}}

Notes:
- Apply the thresholds strictly: choose the highest score whose threshold is satisfied by the evidence/inferred numeric value.
- If evidence is insufficient to estimate a numeric improvement, you may still assign a score but set confidence lower and explain in the "reason".
- Keep each reason brief (1 sentence).
- DO NOT include any text outside the JSON object.

Now analyze the following concatenated abstracts (start of text below). Use only the information in the text to score; if you need to make reasonable inferences, note them explicitly in the reason.
-----BEGIN ABSTRACTS-----
{combined_text}
-----END ABSTRACTS-----
"""
    return prompt.strip()


# -----------------------
# Main pipeline
# -----------------------
def main():
    # 1. Read Excel and select columns
    df = pd.read_excel(INPUT_EXCEL, sheet_name=SHEET_NAME, engine="openpyxl")
    # keep only relevant columns
    df = df.rename(columns={c: c.strip() for c in df.columns})
    assert "Name" in df.columns and "abstract" in df.columns and "published_date" in df.columns, \
        "Excel must contain columns: Name, published_date, abstract"

    # coerce published_date to datetime for sorting
    df["published_date"] = pd.to_datetime(df["published_date"], errors="coerce")
    # drop rows with missing Name or abstract
    df = df.dropna(subset=["Name", "abstract"])

    # 2. Group by drug name and combine abstracts (sorted by published_date ascending)
    grouped = df.sort_values("published_date").groupby("Name", sort=True)
    combined_per_drug = {}
    for name, group in grouped:
        abstracts = []
        for _, row in group.iterrows():
            date_str = ""
            if pd.notna(row["published_date"]):
                date_str = str(row["published_date"].date())
            abstract_text = f"[{date_str}] {row['abstract']}"
            abstracts.append(abstract_text)
        combined_text = "\n\n---\n\n".join(abstracts)
        combined_per_drug[name] = combined_text

    # 3. For each drug: summarize if too long -> build scoring prompt -> call Gemini -> parse JSON
    results = []
    for drug, combined_text in combined_per_drug.items():
        print(f"\nProcessing drug: {drug}  (chars: {len(combined_text)})")

        # If very long, summarize first to keep prompt manageable
        if len(combined_text) > CHARS_SUMMARY_THRESHOLD:
            print("  -> Long text detected; summarizing before scoring...")
            try:
                summary = summarize_long_text(combined_text)
                text_for_model = summary + "\n\n[Original concatenated abstracts omitted for brevity]"
            except Exception as e:
                print("  Summarization failed:", e)
                text_for_model = combined_text[:CHUNK_CHARS]  # fallback to partial text
        else:
            text_for_model = combined_text

        prompt = build_scoring_prompt(drug, text_for_model, thresholds)

        # call gemini; allow simple retries
        attempt = 0
        raw_out = None
        while attempt < 3:
            try:
                raw_out = call_gemini(prompt, max_tokens=800, temperature=0.0)
                break
            except Exception as e:
                attempt += 1
                print(f"  Gemini call failed (attempt {attempt}): {e}")
                time.sleep(1.0 + attempt * 1.0)
        if raw_out is None:
            print(f"  Failed to get model output for {drug}; skipping.")
            continue

        # model is instructed to output only JSON; try to parse it
        parsed = None
        try:
            parsed = json.loads(raw_out)
        except Exception:
            # sometimes model returns code fences or extra text; try to find the JSON substring
            try:
                start = raw_out.find("{")
                end = raw_out.rfind("}") + 1
                json_text = raw_out[start:end]
                parsed = json.loads(json_text)
            except Exception as e:
                print("  Failed to parse JSON from model output. Raw output:\n", raw_out)
                # record raw text for manual inspection
                parsed = {
                    "drug": drug,
                    "scores": {},
                    "total_score": None,
                    "max_score": MAX_SCORE,
                    "confidence": None,
                    "raw_output": raw_out
                }

        # post-process to ensure numeric total exists; if not compute sum
        if isinstance(parsed, dict):
            if parsed.get("total_score") is None:
                # compute sum of numeric endpoint scores
                s = 0
                missing = False
                for ep in ENDPOINTS:
                    sc_obj = parsed.get("scores", {}).get(ep)
                    if isinstance(sc_obj, dict):
                        sc = sc_obj.get("score")
                        if isinstance(sc, (int, float)):
                            s += int(sc)
                        else:
                            missing = True
                    else:
                        missing = True
                if not missing:
                    parsed["total_score"] = s
                else:
                    parsed["total_score"] = parsed.get("total_score")  # keep None if can't compute

            parsed["max_score"] = parsed.get("max_score", MAX_SCORE)
            # attach the combined length for traceability
            parsed["_char_len"] = len(combined_text)

        results.append(parsed)

        # gentle pause to respect rate limits
        time.sleep(0.4)

    # 4. Save results to Excel
    # normalize results to table rows: one row per drug, plus nested score details as string
    rows = []
    for r in results:
        if not isinstance(r, dict):
            continue
        row = {
            "drug": r.get("drug"),
            "total_score": r.get("total_score"),
            "max_score": r.get("max_score"),
            "confidence": r.get("confidence"),
            "raw_output": r.get("raw_output", ""),
            "_char_len": r.get("_char_len", None),
        }
        scores = r.get("scores", {})
        for ep in ENDPOINTS:
            ep_obj = scores.get(ep, {})
            row[f"{ep} - score"] = ep_obj.get("score") if isinstance(ep_obj, dict) else None
            row[f"{ep} - reason"] = ep_obj.get("reason") if isinstance(ep_obj, dict) else None
        rows.append(row)
    out_df = pd.DataFrame(rows)
    out_df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\nDone. Results saved to {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
