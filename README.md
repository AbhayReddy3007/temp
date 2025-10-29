# extract_outcomes_gemini.py
"""
Usage:
    python extract_outcomes_gemini.py /path/to/abstracts.csv

Requirements:
    pip install google-genai pandas tqdm
    # or: pip install git+https://github.com/googleapis/python-genai.git
Notes:
    - The library will pick your API key from environment variable GEMINI_API_KEY.
      If you prefer to hardcode the key, set API_KEY variable below (not recommended for shared machines).
"""

import sys
import json
import time
import pandas as pd
from tqdm import tqdm

# --------- CONFIG ----------
# Option A (recommended): set GEMINI_API_KEY as environment variable
# export GEMINI_API_KEY="YOUR_API_KEY"
#
# Option B (if you insisted on hardcoding): set API_KEY below (not recommended)
API_KEY = None  # <-- replace with "ya29..." or your API key string ONLY IF you understand the risk

MODEL_NAME = "gemini-2.5-flash"
OUTPUT_CSV = "abstracts_with_outcomes.csv"
SLEEP_BETWEEN_REQUESTS = 0.35  # seconds; gentle pacing
MAX_ROWS = None  # set to integer for quick tests, or None to process all

# --------- Gemini client init ----------
try:
    from google import genai
except Exception as e:
    raise SystemExit(
        "Missing google-genai library. Install with:\n    pip install google-genai\n"
    )

if API_KEY:
    client = genai.Client(api_key=API_KEY)
else:
    client = genai.Client()  # will read GEMINI_API_KEY env var if set

# --------- prompt template ----------
PROMPT_TEMPLATE = """
You are a precise extractor. Given the abstract of a clinical / research paper, extract the following four outcomes if they are reported in the abstract. Output EXACTLY one valid JSON object (single-line) with these four keys:

1) weight_loss_pct -> numeric percent (e.g., 12.5) if the abstract reports a percentage of weight loss; otherwise null.
2) a1c_reduction_pct -> numeric percent (e.g., 1.2) if the abstract reports an A1c reduction percentage; otherwise null.
   - If A1c reduction is reported as absolute change in % (e.g., "A1c decreased from 8.5% to 7.3%"), compute the percent point change (8.5 -> 7.3 -> 1.2) and return 1.2.
   - If A1c is reported as relative percent reduction only (e.g., "10% relative reduction"), return 10.
3) mash_resolution -> one of "yes", "no", or "unclear". "yes" if the abstract explicitly states MASH (or NASH/MASH) resolved or showed resolution; "no" if explicitly states no resolution; "unclear" if not reported or ambiguous.
4) alt_resolution -> one of "yes", "no", or "unclear". "yes" if the abstract explicitly states ALT normalized or resolved; "no" if explicitly states not resolved; "unclear" otherwise.

Example output:
{"weight_loss_pct": 12.5, "a1c_reduction_pct": 1.2, "mash_resolution": "yes", "alt_resolution": "unclear"}

Now extract from this abstract (do not add any extra text, commentary, or explanation — only the JSON object):

-----
{abstract_here}
-----
"""

# --------- helper: parse model response ----------
def parse_json_from_response(resp_text):
    """
    Attempt to parse JSON from the model output text.
    If not parseable, return None.
    """
    try:
        # model should output a single-line JSON; but tolerate if there's surrounding whitespace
        text = resp_text.strip()
        # sometimes model wraps JSON in backticks or code fences; remove them
        if text.startswith("```"):
            # find last ```
            parts = text.split("```")
            for p in parts[::-1]:
                p = p.strip()
                if p.startswith("{") and p.endswith("}"):
                    text = p
                    break
        # find first { and last } to extract JSON substring
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            json_text = text[start:end+1]
            data = json.loads(json_text)
            return data
        else:
            # try direct json loads
            return json.loads(text)
    except Exception:
        return None

# --------- main ----------
def main(csv_path):
    df = pd.read_csv(csv_path)
    if "abstract" not in df.columns:
        raise SystemExit("CSV must contain a column named 'abstract' (lowercase).")
    if MAX_ROWS:
        df = df.head(MAX_ROWS)

    # prepare result columns
    df["weight_loss_pct"] = None
    df["a1c_reduction_pct"] = None
    df["mash_resolution"] = None
    df["alt_resolution"] = None
    df["gemini_raw"] = None

    # iterate rows
    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing"):
        abstract = str(row["abstract"])
        prompt = PROMPT_TEMPLATE.replace("{abstract_here}", abstract)

        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt
            )
            # many client responses expose .text convenience; otherwise inspect candidates
            resp_text = ""
            if hasattr(response, "text") and response.text:
                resp_text = response.text
            else:
                # try to safely pull first candidate text
                try:
                    parts = response.candidates[0].content.parts
                    # concatenate text parts
                    resp_text = "".join([p.text for p in parts if p.text])
                except Exception:
                    resp_text = str(response)

            df.at[idx, "gemini_raw"] = resp_text

            parsed = parse_json_from_response(resp_text)
            if parsed is None:
                # fallback: store raw and set unclear/null
                df.at[idx, "weight_loss_pct"] = None
                df.at[idx, "a1c_reduction_pct"] = None
                df.at[idx, "mash_resolution"] = "unclear"
                df.at[idx, "alt_resolution"] = "unclear"
            else:
                # normalize parsed values
                def norm_num(v):
                    if v is None: return None
                    try:
                        return float(v)
                    except Exception:
                        return None
                df.at[idx, "weight_loss_pct"] = norm_num(parsed.get("weight_loss_pct"))
                df.at[idx, "a1c_reduction_pct"] = norm_num(parsed.get("a1c_reduction_pct"))
                df.at[idx, "mash_resolution"] = parsed.get("mash_resolution") if parsed.get("mash_resolution") in ("yes","no","unclear") else "unclear"
                df.at[idx, "alt_resolution"] = parsed.get("alt_resolution") if parsed.get("alt_resolution") in ("yes","no","unclear") else "unclear"

        except Exception as e:
            # on error: save the error text and mark unclear
            df.at[idx, "gemini_raw"] = f"ERROR: {e}"
            df.at[idx, "mash_resolution"] = "unclear"
            df.at[idx, "alt_resolution"] = "unclear"

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    # save results
    df.to_csv(OUTPUT_CSV, index=False)
    # print summary
    total = len(df)
    wl_count = df["weight_loss_pct"].notnull().sum()
    a1c_count = df["a1c_reduction_pct"].notnull().sum()
    mash_yes = (df["mash_resolution"] == "yes").sum()
    alt_yes = (df["alt_resolution"] == "yes").sum()

    print(f"\nProcessed {total} abstracts.")
    print(f"Weight-loss % extracted for {wl_count} abstracts.")
    print(f"A1c reduction % extracted for {a1c_count} abstracts.")
    print(f"MASH resolution reported as 'yes' in {mash_yes} abstracts.")
    print(f"ALT resolution reported as 'yes' in {alt_yes} abstracts.")
    print(f"Results saved to: {OUTPUT_CSV}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_outcomes_gemini.py /path/to/abstracts.csv")
        sys.exit(1)
    csv_path = sys.argv[1]
    main(csv_path)
