"""
Transcription Script
====================
Transcribes the image-only rubric PDF using Gemini 3.5 Flash once,
and saves it to rubric_text.txt for use by text-only LLM APIs (Groq, Cerebras, Nvidia).
"""

import os, sys, io
if sys.platform.startswith('win'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from google import genai
from google.genai import types
from dotenv import load_dotenv

load_dotenv()

RUBRIC_PDF = r"Resume Review Doc.pdf"
OUTPUT_TXT = r"rubric_text.txt"

def main():
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        print("❌ GEMINI_API_KEY not found in .env. Needed once to transcribe the scanned PDF.")
        sys.exit(1)

    print("📄 Connecting to Gemini to transcribe scanned rubric...")
    client = genai.Client(api_key=api_key)

    if not os.path.exists(RUBRIC_PDF):
        print(f"❌ Rubric PDF not found: {RUBRIC_PDF}")
        sys.exit(1)

    print("📄 Uploading scanned rubric PDF (this might take a few seconds)...")
    rubric_file = client.files.upload(file=RUBRIC_PDF)
    print(f"   ✅ Rubric uploaded: {rubric_file.name}")

    print("🤖 transcribing rubric text with Gemini 3.5 Flash...")
    prompt = """This is a scanned PDF of resume evaluation rubrics/rules.
Please extract and transcribe the entire text content of this PDF page-by-page.
Be extremely detailed, preserving all text, tables, guidelines, categories, points, and instructions exactly.
Do not summarize. Output the raw text of the rubric."""

    response = client.models.generate_content(
        model="gemini-3.5-flash",
        contents=[rubric_file, prompt],
        config=types.GenerateContentConfig(
            max_output_tokens=8000
        )
    )

    rubric_text = response.text.strip()
    if not rubric_text:
        print("❌ Gemini returned empty transcription.")
        sys.exit(1)

    with open(OUTPUT_TXT, "w", encoding="utf-8") as f:
        f.write(rubric_text)

    print(f"🎉 SUCCESS! Transcribed rubric saved to: {OUTPUT_TXT} ({len(rubric_text):,} characters)")

if __name__ == "__main__":
    main()
