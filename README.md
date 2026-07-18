# BPHC Resume Reviewer

Automated resume review system using LLMs (Gemini, Groq, NVIDIA NIM). This tool downloads resumes from links provided in an Excel file, reads their text, evaluates them against a rubric, and saves the verdict (Green/Yellow/Red), numeric score, and tailored AI remarks into an output Excel file.

## 📁 Files in this Repository

| File | Purpose |
|------|---------|
| `review_resumes.py` | Original script using Google Gemini (`gemini-3.1-flash-lite`). Requires Google API Key. |
| `review_fast.py` | Faster, multi-LLM script using Llama 3.1 70B via NVIDIA NIM (primary) and Groq. |
| `review_with_cerebras.py` | Cerebras-based inference variant. |
| `requirements.txt` | Python dependencies. |
| `.env` | File for storing API keys securely. |

## ⚙️ Setup Instructions

### 1. Install Dependencies
Make sure you have Python installed. Run the following command to install the required libraries:
```bash
pip install -r requirements.txt
```
*(If you need to use a specific Python version, use `python -m pip install -r requirements.txt`)*

### 2. Configure API Keys
This tool can use multiple LLMs. Based on the script you use, you need the respective API keys.
Create a file named `.env` in this directory and add your keys:
```env
# For review_resumes.py (Gemini)
GEMINI_API_KEY=your_gemini_api_key_here

# For review_fast.py (NVIDIA NIM / Groq)
NVIDIA_API_KEY=your_nvidia_api_key_here
GROQ_API_KEY=your_groq_api_key_here
```

**Where to get the keys:**
- Gemini: https://aistudio.google.com/app/apikey
- NVIDIA NIM: https://build.nvidia.com/
- Groq: https://console.groq.com/keys

### 3. Prepare Your Input Files
1. **Excel Sheet**: Ensure you have an Excel sheet (like `Enrolled_Students_List_...xlsx` or `bhoovan_reviews.xlsx`) with student names and their resume URLs.
2. **Rubric PDF**: (Only for `review_resumes.py`) A PDF defining the evaluation criteria named `Resume Review Doc.pdf` in the root folder.

## ▶️ Running the Automation

You can run either of the following scripts based on your preference.

### Option A: Using Gemini (`review_resumes.py`)
Run the standard Gemini-based script:
```bash
python review_resumes.py
```
**What happens:**
1. It may prompt you for your Gemini API key if not found in `.env`.
2. It asks for any additional personal rules.
3. It downloads each resume, uploads it along with the rubric to Gemini, and generates the evaluation.
4. Results are saved after every student to `bhoovan_reviews_output.xlsx`.

### Option B: Fast Mode (`review_fast.py`)
Run the optimized, multi-LLM version which defaults to processing a limit of 50 resumes:
```bash
python review_fast.py --limit 50 --output Reviewed_50_Candidates.xlsx
```
**What happens:**
1. It uses NVIDIA NIM (Llama 3.1/3.3) for extremely fast inference.
2. It processes candidates one by one. If a link is invalid, it is skipped.
3. The results are saved after each candidate into the specified output Excel file.
4. It avoids re-evaluating students if they already have a score in the output file.

## 📊 Output Format
The resulting Excel file will have color-coded rows with new columns for the AI's review:
- **Green (Score 90-100 / 80-100)**: Excellent, meets most rubric criteria.
- **Yellow (Score 55-89 / 55-79)**: Average, needs minor/moderate cosmetic or structural improvements.
- **Red (Score < 55)**: Poor, fails key criteria, lacks metrics, or contains major red flags.

> **💡 Tip:** The output file is continually saved during execution, so you can safely stop the script and resume later without losing progress!
