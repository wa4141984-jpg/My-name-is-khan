# # 💰 AI Expense Analyzer

## Project overview
AI Expense Analyzer is a Streamlit web app. You upload a CSV or Excel file of your expenses, and the app shows a dashboard, charts, a monthly budget check, and friendly tips written by Google Gemini.

## Problem the application solves
Many people do not know where their money goes. Looking at a long spreadsheet is hard. This app turns the spreadsheet into totals, charts and simple advice. You do **not** need to share a bank password.

## Features
- Upload `.csv` or `.xlsx` files (5 MB limit) and preview them
- Match your own column names to Date, Category, Description and Amount
- Cleans bad data: missing values, bad dates, non-numeric amounts, duplicates
- Refund handling: negative amounts can be subtracted or ignored
- Metric cards: total, average, count, highest expense
- Interactive Plotly charts: donut, bar, and line (trend over time)
- Sortable transaction table
- Monthly budget planner with a warning for incomplete months
- AI insights from Gemini (optional; the app works without a key)
- Downloads: cleaned CSV, summary text, AI report (Markdown)
- Currency display: PKR, USD, EUR, GBP (display only, no conversion)

## Technology stack
Python, Streamlit, Pandas, Plotly, Google Gen AI SDK (`google-genai`), openpyxl, Git/GitHub, Streamlit Community Cloud.

## Project structure
```
ai-expense-analyzer/
├── app.py             # the whole application
├── requirements.txt   # Python libraries
├── README.md          # this file
├── .gitignore         # files Git must ignore (see below)
└── .streamlit/
    └── secrets.toml   # YOUR API key (local only, never commit)
```

## Screenshots
Run the app, take screenshots, and save them in a folder called `screenshots/`. Then add lines like this:
```markdown
![Dashboard](screenshots/dashboard.png)
![Charts](screenshots/charts.png)
```
Use sample data only. Never share screenshots of real financial data.

## Prerequisites
- Python 3.10 or newer (https://www.python.org/downloads/) — tick **Add Python to PATH** on Windows
- Git (https://git-scm.com/downloads)
- A free GitHub account
- A Gemini API key (optional, only for AI insights)

## How to obtain a Gemini API key
1. Go to https://aistudio.google.com/ and sign in with a Google account.
2. Click **Get API key**, then **Create API key**.
3. Copy the key and keep it private. Never paste it into `app.py`, GitHub, or screenshots.

## Installation instructions
Open a terminal in the project folder, then follow the next section.

## Virtual environment setup
A virtual environment keeps this project's libraries separate from other projects.
```bash
python -m venv .venv
```
Activate it:
- Windows (PowerShell): `.venv\Scripts\Activate.ps1`
- Windows (Command Prompt): `.venv\Scripts\activate.bat`
- macOS / Linux: `source .venv/bin/activate`

Then install the libraries:
```bash
python -m pip install -r requirements.txt
```

## How to configure the API key
Choose **one** option.

**Option A: Streamlit secrets file (recommended).** Create `.streamlit/secrets.toml`:
```toml
GEMINI_API_KEY = "your_api_key_here"
```
Replace the text with your real key. This file is ignored by Git.

**Option B: environment variable.**
- Windows PowerShell: `$env:GEMINI_API_KEY="your_api_key_here"`
- macOS / Linux: `export GEMINI_API_KEY="your_api_key_here"`

Optional: set `GEMINI_MODEL` the same way if Google retires the default model. The default is `gemini-3.6-flash`.

## How to run the application locally
```bash
streamlit run app.py
```
Your browser opens at http://localhost:8501.

## Example input CSV format
```csv
Date,Category,Description,Amount
2026-10-01,Food,Groceries,2000
2026-10-03,Transport,Bus fare,500
2026-10-05,Education,Books,1500
2026-10-07,Entertainment,Games,1000
2026-10-10,Food,Restaurant,1200
```
Only **Amount** is required. Date, Category and Description are optional, but charts and budgets need dates. Negative amounts are treated as refunds.

## How to upload the project to GitHub
```bash
git init
git add app.py requirements.txt README.md .gitignore
git commit -m "First version of AI Expense Analyzer"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/ai-expense-analyzer.git
git push -u origin main
```
Create the empty repository on github.com first (do not add a README there). Check `git status` before committing: `.env`, `secrets.toml` and spreadsheets must **not** appear.

## How to deploy on Streamlit Community Cloud
1. Push the project to GitHub.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app** and choose your repository and the `main` branch.
4. Set **Main file path** to `app.py`.
5. Open **Advanced settings** and paste your secrets (next section).
6. Click **Deploy**, wait, then test the public URL with a sample CSV.

## How to configure Streamlit secrets
In **Advanced settings → Secrets** (or later: app menu → Settings → Secrets), paste:
```toml
GEMINI_API_KEY = "your_api_key_here"
```
Local environment variables only exist on your computer. Streamlit Cloud does not see them, so you must add secrets in the Cloud dashboard.

## Common errors and solutions
| Problem | Solution |
|---|---|
| `streamlit: command not found` | Activate the virtual environment, or run `python -m streamlit run app.py` |
| `ModuleNotFoundError` | Run `python -m pip install -r requirements.txt` |
| "No Gemini API key found" | Add `GEMINI_API_KEY` to secrets or environment, then restart the app |
| "rate limit or quota reached" | Wait a minute and try again |
| "model was not found" | Set `GEMINI_MODEL` to a current Gemini Flash model ID |
| "API key is correct and active" | Create a new key in Google AI Studio |
| "Please choose which column contains the Amount" | Pick the right column in the sidebar |
| Cloud build fails | Check `requirements.txt` is in the repo root and the file names are correct |

## Testing instructions
1. Run `python -m py_compile app.py` (checks syntax).
2. Upload the example CSV and check: total = 6,200, average = 1,240, 5 transactions, highest = 2,000 (Food).
3. Try a file with a bad amount (`abc`), a bad date, an empty file, and a `.txt` file. You should see friendly messages, not crashes.
4. Remove the API key and check that the dashboard still works.
5. Set a budget of 10,000 and select 2026-10. You should see 62% used and a warning that the month may be incomplete.
6. Download all three reports and open them.

## Privacy considerations and limitations
- Uploaded data stays in memory during your session; the app does not save it.
- Gemini receives only totals, counts and category names, not descriptions. You can preview this in the app. Category names are text you chose, so avoid putting private details in them.
- Reports never contain your API key.
- AI tips are estimates, not financial advice, and cannot guarantee savings.
- Currency is a label only. There is no currency conversion.
- Only the first sheet of an Excel file is read.
- Duplicate removal can remove real repeated purchases; you can turn it off.
- If you deploy publicly, anyone with the link can use your API quota. Consider making the app private.

## Future improvements
- Currency conversion, multiple Excel sheets, budgets per category
- Recurring-payment detection, saved settings, user login
- Automated test files and a test workflow on GitHubMy-name-is-khan
