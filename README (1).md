# 📄 ATS Resume Checker

A Streamlit app that estimates how well a resume will perform in an Applicant Tracking System (ATS) and tells you exactly what to improve. The AI part is powered by Claude.

## What it does

1. **Upload** a resume (PDF, DOCX or TXT, up to 5 MB).
2. *(Optional)* **Paste a job description** for a targeted keyword match.
3. Get:
   - an **estimated ATS score (0-100)** with a breakdown by category,
   - **prioritised improvements** (high / medium / low) with suggested rewrites,
   - **keywords found and missing**,
   - **automatic checks** (email, phone, LinkedIn, section headings, length),
   - a view of **the exact text an ATS would extract** from your file,
   - a **downloadable Markdown report**.

## How the score works

Claude rates seven categories from 0-100. The overall score is then calculated **in Python** as a weighted average, so it always matches the breakdown.

| Category | Weight |
|---|---|
| Formatting & ATS parseability | 20% |
| Keywords & relevance | 25% |
| Experience & achievements | 25% |
| Skills section | 10% |
| Education & certifications | 5% |
| Contact info & structure | 5% |
| Clarity & language | 10% |

**Score bands:** 85+ Excellent · 70-84 Good · 55-69 Needs improvement · below 55 Poor.

> ⚠️ This is an AI **estimate** for guidance. Real ATS products differ between employers and do not publish a universal score. Only add skills and achievements that are true for you.

The app also runs plain-Python checks (no AI) for things like a missing email/phone, scanned/image-only PDFs, tables, text boxes, images, and text hidden in DOCX headers/footers, because these are common reasons an ATS misreads a resume.

## Project files

```
app.py            # the whole app (UI + text extraction + Claude call + scoring)
requirements.txt  # Python dependencies
README.md         # this file
```

## Run it locally

You need Python 3.9+ (3.11 or 3.12 recommended) and an Anthropic API key from <https://console.anthropic.com>.

```bash
# 1. (recommended) create a fresh virtual environment
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# 2. install dependencies
pip install -r requirements.txt

# 3. run
streamlit run app.py
```

Then open the URL Streamlit prints (usually <http://localhost:8501>).

### Providing your API key

Use **one** of these:

| Method | How |
|---|---|
| Paste in the app | Type it into the **Anthropic API key** box in the sidebar (kept for that browser session only). |
| Secrets file (local) | Create `.streamlit/secrets.toml` containing `ANTHROPIC_API_KEY = "sk-ant-..."` |
| Environment variable | `export ANTHROPIC_API_KEY="sk-ant-..."` (Windows PowerShell: `$env:ANTHROPIC_API_KEY="sk-ant-..."`) |

**Never commit `.streamlit/secrets.toml` or your key to GitHub.**

### Choosing the model

Pick a model in the sidebar:

| Option | Model ID | Good for |
|---|---|---|
| Claude Sonnet 5.5 (default) | `claude-sonnet-5-5` | best balance of quality, speed, cost |
| Claude Haiku 5.5 | `claude-haiku-5-5` | fastest and cheapest |
| Claude Opus 5.5 | `claude-opus-5-5` | most thorough |

To force a specific model (for example if Anthropic releases a newer one), set `CLAUDE_MODEL = "model-id"` as a secret or environment variable. The sidebar picker is then hidden. Current model IDs: <https://docs.claude.com/en/docs/about-claude/models/overview>

## Deploy on Streamlit Community Cloud

1. Put `app.py`, `requirements.txt` and `README.md` in the **root** of a GitHub repository.
2. Go to <https://share.streamlit.io>, sign in with GitHub, and click **Create app**.
3. Choose your repository, branch (`main`) and main file path (`app.py`).
4. Open **Advanced settings** and (optional) paste your key into **Secrets**:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
5. Click **Deploy**. Later changes pushed to GitHub redeploy automatically.

### Who pays for the API calls?

- **Key stored in Secrets** → every visitor uses *your* credits. Reduce the risk by setting a spend limit in the Anthropic Console and, if possible, restricting who can view the app. The app also has a soft limit of 10 analyses per browser session (`RUNS_PER_SESSION_WITH_SERVER_KEY` in `app.py`); a determined user can bypass it by refreshing, so it is not real protection.
- **No key in Secrets** → each visitor pastes their own key in the sidebar and pays for their own usage.

## Privacy

The resume text (and job description) is sent to Anthropic's API to be analysed. The app does not save uploaded files or write them to disk. If you run this for other people, tell them that, and remember resumes contain personal data.

## Troubleshooting

| Problem | Fix |
|---|---|
| "Very little text could be extracted" | The PDF is a scan/image. Export a text-based PDF or DOCX from Word or Google Docs. |
| "API key was rejected" | Check for typos or extra spaces; make sure the key is active in the Console. |
| Error mentioning credit balance | Add credits to your Anthropic account. |
| 404 / "model not found" | The model ID is wrong or retired. Set `CLAUDE_MODEL` to a current ID. |
| "Rate limit reached" | Wait a minute and retry, or use Claude Haiku 5.5. |
| Old `.doc` file | Save it as `.docx` or PDF first. |
| Results look odd for a two-column resume | Check the **What the ATS sees** tab. If the text is jumbled, an ATS will struggle too, so use a single-column layout. |

## Limitations

- The AI sees extracted **text only**, not the visual design, so layout judgements are inferred from the extracted text and the automatic checks.
- Only the first 10 PDF pages and about 25,000 characters are analysed.
- Scores can vary slightly between runs because the model's wording and judgement are not perfectly deterministic.
