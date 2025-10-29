# extract_outcomes_with_gemini.py
"""
Requirements:
  pip install google-genai pandas
Notes:
  - Recommended: set GEMINI_API_KEY as an environment variable.
  - If you really want to hardcode the API key, set API_KEY below (not recommended publicly).
References:
  Gemini docs / examples: see Google GenAI / Vertex docs for gemini-2.5-flash usage.
  (Example usage patterns used while building this script). 
"""

import os
import json
import re
import time
import pandas as pd

# Option A (recommended): from env var
API_KEY = os.environ.get("GEMINI_API_KEY")

# Option B (less secure) - uncomment to hardcode (only for local testing)
# API_KEY = "your_real_gemini_api_key_here"

if not API_KEY:
    raise RuntimeError("No GEMINI_API_KEY found. Set GEMINI_API_KEY env var or hardcode API_KEY in script.")

# Import the official client
try:
    from google import genai
except Exception as e:
    raise RuntimeError("Please install google-genai package: pip install google-genai") from e

def build_prompt(abstract_text):
    """
    Prompt instructs the model to output only JSON (no extra commentary).
    The model should return numeric percent values (float or int) for percentages,
    'yes'/'no'/'unknown' for resolution fields, and include a short supporting excerpt.
    """
    return f"""
You are a careful scientific extractor. Given the following research abstract, extract EXACTLY the following JSON object and nothing else:

{{
  "weight_loss_pct": <number or null>,         # % weight loss reported (as a number), or null if not reported
  "a1c_reduction_pct": <number or null>,       # % A1c reduction reported, or null
  "mash_resolution": <"yes"|"no"|"unknown">,    # whether MASH (NASH with metabolic dysfunction) resolved/remitted
  "alt_resolution": <"yes"|"no"|"unknown">,     # whether ALT (alanine aminotransferase) normalized/resolved
  "confidence": "<brief confidence estimate>", # short (e.g., 'high', 'low', or a short note)
  "supporting_text": "<short excerpt from the abstract used to justify>" 
}}

Rules:
- Return valid JSON only. Do NOT include explanations or extra text.
- Percent values must be numeric (e.g., 12.5). If abstract says 'no change' or doesn't report, use null.
- For resolution fields, use "yes" only if the abstract explicitly states resolution/normalization or provides clear numeric evidence; use "no" if it explicitly says not resolved; otherwise "unknown".
- Keep supporting_text short (<= 50 words), copy verbatim fragment that supports your extraction.

Abstract:
\"\"\"{abstract_text}\"\"\"
"""

def call_gemini(prompt, model="gemini-2.5-flash"):
    """
    Make a synchronous generate_content call to Gemini.
    This example uses the official genai client. Behavior may vary slightly by SDK version.
    """
    client = genai.Client(api_key=API_KEY)

    # Generate content. Adjust model ID if your account uses a different one.
    response = client.models.generate_content(
        model=model,
        contents=[prompt],
        # You can adjust temperature, max tokens etc. via config if SDK/version supports it.
    )

    # Extract text from the candidate parts (SDK returns parts; choose first candidate)
    # This extraction mirrors typical responses: candidates[0].content.parts
    text_parts = []
    try:
        candidate = response.candidates[0]
        for part in candidate.content.parts:
            if getattr(part, "text", None):
                text_parts.append(part.text)
    except Exception:
        # fallback to simpler attribute names if SDK differs
        try:
            text_parts.append(response.text)
        except Exception:
            raise RuntimeError("Unexpected response format from Gemini client; inspect `response` object.")
    return "\n".join(text_parts).strip()

# Fallback local regex extraction if model doesn't return JSON
def fallback_extract(abstract):
    res = {
        "weight_loss_pct": None,
        "a1c_reduction_pct": None,
        "mash_resolution": "unknown",
        "alt_resolution": "unknown",
        "confidence": "fallback_regex",
        "supporting_text": ""
    }

    # find percent-like numbers
    pct_matches = re.findall(r'([0-9]+(?:\.[0-9]+)?)\s*%|\b([0-9]+(?:\.[0-9]+)?)\s*percent', abstract, flags=re.I)
    # flatten matches
    flat = []
    for a,b in pct_matches:
        flat.append(a or b)
    # naive heuristics: first percent referencing weight or BMI -> weight_loss; A1c often "A1c" or "HbA1c"
    # find weight loss phrases
    weight_match = re.search(r'(weight (?:loss|reduction)[^\d\n\r]{0,40}([0-9]+(?:\.[0-9]+)?)[\s%]|lost\s([0-9]+(?:\.[0-9]+)?)\s%?)', abstract, flags=re.I)
    if weight_match:
        num = weight_match.group(2) or weight_match.group(3)
        try:
            res["weight_loss_pct"] = float(num)
            res["supporting_text"] += f"weight phrase: {weight_match.group(0)} "
        except:
            pass

    # A1c
    a1c_match = re.search(r'(A1c|HbA1c)[^\d\n\r]{0,40}([0-9]+(?:\.[0-9]+)?)\s*%?', abstract, flags=re.I)
    if a1c_match:
        try:
            res["a1c_reduction_pct"] = float(a1c_match.group(2))
            res["supporting_text"] += f"A1c phrase: {a1c_match.group(0)} "
        except:
            pass

    # MASH/NASH resolution heuristics
    if re.search(r'\b(resolv(ed|tion)|remission|remitted|resolved)\b', abstract, flags=re.I) and re.search(r'(MASH|NASH)', abstract, flags=re.I):
        res["mash_resolution"] = "yes"
        res["supporting_text"] += "MASH/NASH resolution phrase found. "
    elif re.search(r'(MASH|NASH)', abstract, flags=re.I):
        res["mash_resolution"] = "unknown"

    # ALT normalization
    if re.search(r'(ALT|alanine aminotransferase)[^\n]{0,80}\b(normaliz|normaliz|resolved|within normal)\b', abstract, flags=re.I):
        res["alt_resolution"] = "yes"
        res["supporting_text"] += "ALT normalization phrase found. "
    elif re.search(r'(ALT|alanine aminotransferase)', abstract, flags=re.I):
        res["alt_resolution"] = "unknown"

    return res

def parse_model_json(text):
    """
    Try to load JSON from the model's output. The model is instructed to produce JSON only.
    """
    # Trim anything before first { and after last } to be robust
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object detected")
    json_text = text[start:end+1]
    # Some models include trailing commas / comments - try to fix trivial issues
    json_text = re.sub(r",\s*}", "}", json_text)
    json_text = re.sub(r",\s*]", "]", json_text)
    return json.loads(json_text)

def process_csv(input_csv="abstracts.csv", output_csv="abstracts_with_outcomes.csv", model="gemini-2.5-flash"):
    df = pd.read_csv(input_csv)
    if 'abstract' not in df.columns:
        raise RuntimeError("CSV must contain a column named 'abstract'")

    results = []
    for idx, row in df.iterrows():
        abstract = str(row['abstract'])
        prompt = build_prompt(abstract)

        # call model, with simple retry
        try:
            raw = call_gemini(prompt, model=model)
        except Exception as e:
            print(f"[{idx}] Gemini call failed: {e}. Falling back to local regex.")
            parsed = fallback_extract(abstract)
            parsed['model_raw'] = ""
            results.append(parsed)
            time.sleep(0.2)
            continue

        # Try parsing JSON
        parsed = None
        try:
            parsed_json = parse_model_json(raw)
            # Normalize fields
            parsed = {
                "weight_loss_pct": parsed_json.get("weight_loss_pct"),
                "a1c_reduction_pct": parsed_json.get("a1c_reduction_pct"),
                "mash_resolution": parsed_json.get("mash_resolution"),
                "alt_resolution": parsed_json.get("alt_resolution"),
                "confidence": parsed_json.get("confidence"),
                "supporting_text": parsed_json.get("supporting_text"),
                "model_raw": raw
            }
        except Exception as e:
            # fallback to regex heuristics
            print(f"[{idx}] Failed to parse JSON from model output: {e}. Attempting regex fallback.")
            parsed = fallback_extract(abstract)
            parsed['model_raw'] = raw

        results.append(parsed)
        # polite pacing
        time.sleep(0.1)

    out_df = pd.concat([df.reset_index(drop=True), pd.DataFrame(results)], axis=1)
    out_df.to_csv(output_csv, index=False)
    print(f"Done. Saved to {output_csv}")

if __name__ == "__main__":
    # If you want to change input/output filenames or model id, edit below
    process_csv(input_csv="abstracts.csv", output_csv="abstracts_with_outcomes.csv", model="gemini-2.5-flash")
