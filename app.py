"""
AI Expense Analyzer
-------------------
Upload a CSV/Excel expense file, see a dashboard, plan a budget,
and (optionally) get AI tips from Google Gemini.

All money calculations are done with Python and Pandas.
Gemini only writes advice from numbers that Python already calculated.
"""

import hashlib
import json
import os
from datetime import date
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

# The Gemini SDK is optional at start-up: if it is missing, the dashboard still works.
try:
    from google import genai
    from google.genai import errors as genai_errors
    from google.genai import types as genai_types

    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
MAX_UPLOAD_MB = 5
# Stable Gemini Flash model ID. You can override it with a GEMINI_MODEL secret/env variable.
DEFAULT_MODEL = "gemini-3.6-flash"
CURRENCIES = ["PKR", "USD", "EUR", "GBP", "None"]

# Words we look for when guessing which column is which.
COLUMN_HINTS = {
    "Date": ["date", "day", "time", "when"],
    "Category": ["category", "type", "group"],
    "Description": ["description", "details", "note", "item", "memo", "name"],
    "Amount": ["amount", "cost", "price", "total", "spent", "value", "debit"],
}

SAMPLE_CSV = (
    "Date,Category,Description,Amount\n"
    "2026-10-01,Food,Groceries,2000\n"
    "2026-10-03,Transport,Bus fare,500\n"
    "2026-10-05,Education,Books,1500\n"
    "2026-10-07,Entertainment,Games,1000\n"
    "2026-10-10,Food,Restaurant,1200\n"
)


# ----------------------------------------------------------------------------
# 1. File loading
# ----------------------------------------------------------------------------
def load_file(file_name, file_bytes):
    """Read a CSV or Excel file into a DataFrame. Raises ValueError with a friendly message."""
    if not file_bytes:
        raise ValueError("The file is empty.")
    if len(file_bytes) > MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(f"The file is bigger than {MAX_UPLOAD_MB} MB. Please upload a smaller file.")

    name = file_name.lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(BytesIO(file_bytes), sep=None, engine="python", encoding_errors="replace")
        elif name.endswith(".xlsx"):
            df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")  # first sheet only
        else:
            raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")
    except ValueError:
        raise
    except Exception:
        raise ValueError("The file could not be read. Make sure it is a valid CSV or Excel (.xlsx) file.")

    df = df.dropna(how="all")
    if df.empty or len(df.columns) == 0:
        raise ValueError("The file has no data rows.")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def guess_column(columns, field):
    """Guess which column in the file matches Date, Category, Description or Amount."""
    lowered = {c: c.lower() for c in columns}
    for column, text in lowered.items():  # exact name first
        if text == field.lower():
            return column
    for hint in COLUMN_HINTS[field]:  # then partial match
        for column, text in lowered.items():
            if hint in text:
                return column
    return None


# ----------------------------------------------------------------------------
# 2. Cleaning and validation
# ----------------------------------------------------------------------------
def _to_amount(series):
    """Turn a column into numbers. Text like 'Rs 2,000' becomes 2000. Bad values become NaN."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    text = series.where(series.notna(), "").astype(str).str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(text, errors="coerce")


def clean_expenses(raw, mapping, remove_duplicates=True, refund_mode="net"):
    """
    Clean the uploaded data.
    mapping: {"Date": col or None, "Category": col or None, "Description": col or None, "Amount": col}
    refund_mode: "net"     -> negative amounts are refunds and reduce spending
                 "exclude" -> negative amounts are removed
    Returns (clean_df, report_dict). Raises ValueError if the data cannot be used.
    """
    amount_col = mapping.get("Amount")
    if not amount_col or amount_col not in raw.columns:
        raise ValueError("Please choose which column contains the Amount.")

    report = {"rows_in": len(raw), "invalid_amount": 0, "invalid_date": 0,
              "missing_category": 0, "duplicates_removed": 0,
              "negative_rows": 0, "negative_excluded": 0}

    df = pd.DataFrame(index=raw.index)
    df["Amount"] = _to_amount(raw[amount_col])
    df.loc[df["Amount"].abs() == float("inf"), "Amount"] = float("nan")

    # Date (optional)
    if mapping.get("Date") in raw.columns:
        df["Date"] = pd.to_datetime(raw[mapping["Date"]], errors="coerce", format="mixed")
        report["invalid_date"] = int(df["Date"].isna().sum())
    else:
        df["Date"] = pd.NaT

    # Category (optional)
    if mapping.get("Category") in raw.columns:
        cat = raw[mapping["Category"]]
        cat = cat.where(cat.notna(), "").astype(str).str.strip()
    else:
        cat = pd.Series("", index=raw.index)
    report["missing_category"] = int((cat == "").sum())
    df["Category"] = cat.replace("", "Uncategorized")

    # Description (optional)
    if mapping.get("Description") in raw.columns:
        desc = raw[mapping["Description"]]
        df["Description"] = desc.where(desc.notna(), "").astype(str).str.strip()
    else:
        df["Description"] = ""

    # Remove rows with unusable amounts
    bad = df["Amount"].isna()
    report["invalid_amount"] = int(bad.sum())
    df = df[~bad]

    # Refunds / negative amounts
    report["negative_rows"] = int((df["Amount"] < 0).sum())
    if refund_mode == "exclude":
        report["negative_excluded"] = report["negative_rows"]
        df = df[df["Amount"] >= 0]

    # Duplicates
    if remove_duplicates:
        before = len(df)
        df = df.drop_duplicates(subset=["Date", "Category", "Description", "Amount"])
        report["duplicates_removed"] = before - len(df)

    df = df[["Date", "Category", "Description", "Amount"]].reset_index(drop=True)
    if df.empty:
        raise ValueError("No valid expense rows were found. Check that the Amount column contains numbers.")
    return df, report


# ----------------------------------------------------------------------------
# 3. Calculations (all done in Python / Pandas)
# ----------------------------------------------------------------------------
def spending_by_category(df):
    grouped = df.groupby("Category")["Amount"].agg(Total="sum", Transactions="count").reset_index()
    total = grouped["Total"].sum()
    grouped["Share %"] = (grouped["Total"] / total * 100).round(1) if total != 0 else 0.0
    return grouped.sort_values("Total", ascending=False).reset_index(drop=True)


def spending_over_time(df):
    """Daily totals for short periods, monthly totals for long periods. Empty DataFrame if no dates."""
    dated = df.dropna(subset=["Date"])
    if dated.empty:
        return pd.DataFrame(columns=["Period", "Amount"])
    span_days = (dated["Date"].max() - dated["Date"].min()).days
    if span_days > 62:
        period = dated["Date"].dt.to_period("M").dt.to_timestamp()
    else:
        period = dated["Date"].dt.normalize()
    out = dated.groupby(period)["Amount"].sum().reset_index()
    out.columns = ["Period", "Amount"]
    return out


def monthly_totals(df):
    dated = df.dropna(subset=["Date"])
    if dated.empty:
        return pd.DataFrame(columns=["Month", "Amount"])
    out = dated.groupby(dated["Date"].dt.to_period("M").astype(str))["Amount"].sum().reset_index()
    out.columns = ["Month", "Amount"]
    return out


def compute_summary(df):
    positives = df[df["Amount"] > 0]
    highest = None
    if not positives.empty:
        row = positives.loc[positives["Amount"].idxmax()]
        highest = {"amount": float(row["Amount"]), "category": row["Category"],
                   "description": row["Description"],
                   "date": None if pd.isna(row["Date"]) else row["Date"].strftime("%Y-%m-%d")}
    dated = df.dropna(subset=["Date"])
    return {
        "total": float(df["Amount"].sum()),
        "average": float(df["Amount"].mean()),
        "count": int(len(df)),
        "highest": highest,
        "refunds_total": float(df.loc[df["Amount"] < 0, "Amount"].sum()),
        "by_category": spending_by_category(df),
        "monthly": monthly_totals(df),
        "date_min": None if dated.empty else dated["Date"].min(),
        "date_max": None if dated.empty else dated["Date"].max(),
    }


def calculate_budget(budget, spent):
    """Compare spending with a monthly budget."""
    if budget is None or budget <= 0:
        raise ValueError("Budget must be greater than zero.")
    return {
        "budget": float(budget),
        "spent": float(spent),
        "remaining": float(budget - spent),
        "utilization": float(spent / budget * 100),
        "status": "Over budget" if spent > budget else "Within budget",
    }


def month_warning(df, month_text, today=None):
    """Return a warning text if the chosen month may be incomplete, otherwise None."""
    dated = df.dropna(subset=["Date"])
    if dated.empty:
        return None
    today = today or date.today()
    period = pd.Period(month_text, freq="M")
    start, end = period.start_time, period.end_time.normalize()
    reasons = []
    if dated["Date"].min() > start:
        reasons.append(f"your data starts on {dated['Date'].min():%Y-%m-%d}")
    if dated["Date"].max() < end:
        reasons.append(f"your data ends on {dated['Date'].max():%Y-%m-%d}")
    if end.date() > today:
        reasons.append("this month is not finished yet")
    if not reasons:
        return None
    return ("This month may be incomplete (" + "; ".join(reasons) + "). "
            "Comparing partial spending with a full-month budget can make you look better than you really are.")


# ----------------------------------------------------------------------------
# 4. Charts
# ----------------------------------------------------------------------------
def make_pie_chart(by_category, currency):
    data = by_category[by_category["Total"] > 0]
    if data.empty:
        return None
    fig = px.pie(data, names="Category", values="Total", hole=0.45, title="Expenses by category")
    fig.update_traces(hovertemplate="%{label}<br>" + currency + " %{value:,.2f}<extra></extra>")
    return fig


def make_bar_chart(by_category, currency):
    if by_category.empty:
        return None
    fig = px.bar(by_category, x="Category", y="Total", title="Spending by category",
                 labels={"Total": f"Total ({currency})"}, text_auto=".2s")
    return fig


def make_line_chart(df, currency):
    trend = spending_over_time(df)
    if trend.empty:
        return None
    fig = px.line(trend, x="Period", y="Amount", markers=True, title="Spending over time",
                  labels={"Amount": f"Amount ({currency})", "Period": "Date"})
    return fig


# ----------------------------------------------------------------------------
# 5. Gemini AI
# ----------------------------------------------------------------------------
class AIError(Exception):
    """An error with a message that is safe to show to the user (never contains the API key)."""


def get_setting(name):
    """Look in Streamlit secrets first, then environment variables. Returns None if not found."""
    value = None
    try:
        value = st.secrets.get(name)
    except Exception:  # no secrets file exists
        value = None
    if not value:
        value = os.environ.get(name)
    value = str(value).strip() if value else None
    if not value or value == "your_api_key_here":
        return None
    return value


def build_ai_payload(summary, currency, budget_info=None):
    """Aggregated numbers only. No descriptions, no dates of single transactions."""
    cats = summary["by_category"]
    top = cats.head(15)
    payload = {
        "currency": currency if currency != "None" else "unspecified",
        "total_spent": round(summary["total"], 2),
        "average_transaction": round(summary["average"], 2),
        "transaction_count": summary["count"],
        "largest_single_transaction": None if not summary["highest"] else {
            "amount": round(summary["highest"]["amount"], 2), "category": summary["highest"]["category"]},
        "refunds_total": round(summary["refunds_total"], 2),
        "categories": [{"name": r["Category"], "total": round(r["Total"], 2),
                        "share_percent": r["Share %"], "transactions": int(r["Transactions"])}
                       for _, r in top.iterrows()],
        "other_categories_count": max(0, len(cats) - len(top)),
        "monthly_totals": [{"month": r["Month"], "total": round(r["Amount"], 2)}
                           for _, r in summary["monthly"].tail(24).iterrows()],
    }
    if budget_info:
        payload["monthly_budget"] = {k: (round(v, 2) if isinstance(v, float) else v)
                                     for k, v in budget_info.items()}
    return payload


AI_KEYS = ["summary", "major_categories", "unusual_spending", "repeated_patterns",
           "saving_ideas", "budget_recommendation"]


def build_prompt(payload):
    return (
        "You are a careful, friendly personal-finance assistant. Use simple English.\n"
        "Below is a JSON summary of a person's expenses. Use ONLY the numbers in it. "
        "Never invent totals. Do not say any transaction is fraud. Do not assume a high expense "
        "is unnecessary; say it may be worth reviewing because you do not know the context. "
        "Do not promise savings; your tips are estimates.\n"
        "Reply with ONLY a JSON object with these keys:\n"
        '"summary" (string, 2-3 sentences), '
        '"major_categories" (list of strings), '
        '"unusual_spending" (list of strings), '
        '"repeated_patterns" (list of strings), '
        '"saving_ideas" (list of strings, practical), '
        '"budget_recommendation" (string).\n\n'
        "EXPENSE SUMMARY:\n" + json.dumps(payload, indent=2)
    )


def validate_ai_response(text):
    """Check that the AI answer is valid JSON in the shape we expect. Raises AIError if not."""
    if not text or not text.strip():
        raise AIError("The AI returned an empty answer. Please try again.")
    cleaned = text.strip()
    if cleaned.startswith("```"):  # remove ```json fences if present
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        raise AIError("The AI answer was not in the expected format. Please try again.")
    if not isinstance(data, dict):
        raise AIError("The AI answer was not in the expected format. Please try again.")

    result = {}
    for key in AI_KEYS:
        value = data.get(key)
        if key in ("summary", "budget_recommendation"):
            if not isinstance(value, str) or not value.strip():
                raise AIError("The AI answer was missing some parts. Please try again.")
            result[key] = value.strip()[:1500]
        else:
            if not isinstance(value, list) or not all(isinstance(v, str) for v in value):
                raise AIError("The AI answer was missing some parts. Please try again.")
            result[key] = [v.strip()[:500] for v in value if v.strip()][:8]
    return result


def request_ai_insights(payload, api_key, model_name):
    """Call Gemini and return a validated dictionary. Raises AIError with a safe message."""
    if not GENAI_AVAILABLE:
        raise AIError("The google-genai package is not installed. Run: python -m pip install -r requirements.txt")
    if not api_key:
        raise AIError("No Gemini API key was found. Add GEMINI_API_KEY to your secrets or environment.")
    try:
        client = genai.Client(api_key=api_key, http_options=genai_types.HttpOptions(timeout=60000))
        response = client.models.generate_content(
            model=model_name,
            contents=build_prompt(payload),
            config=genai_types.GenerateContentConfig(
                temperature=0.3, response_mime_type="application/json"),
        )
        text = response.text
    except genai_errors.APIError as err:
        code = getattr(err, "code", None)
        if code == 429:
            raise AIError("Gemini rate limit or quota reached. Wait a minute and try again.")
        if code == 404:
            raise AIError(f"The model '{model_name}' was not found. Set GEMINI_MODEL to a current Flash model ID.")
        if code in (400, 401, 403):
            raise AIError("Gemini rejected the request. Check that your API key is correct and active.")
        if code is not None and code >= 500:
            raise AIError("The Gemini service is temporarily unavailable. Try again later.")
        raise AIError("Gemini returned an error. Please try again.")
    except Exception:
        raise AIError("Could not reach Gemini. Check your internet connection and try again.")
    return validate_ai_response(text)


# ----------------------------------------------------------------------------
# 6. Reports
# ----------------------------------------------------------------------------
def money(value, currency):
    prefix = "" if currency == "None" else currency + " "
    return f"{prefix}{value:,.2f}"


def date_range_text(summary):
    if summary["date_min"] is None:
        return "No valid dates"
    return f"{summary['date_min']:%Y-%m-%d} to {summary['date_max']:%Y-%m-%d}"


def safe_for_spreadsheet(df):
    """Stop spreadsheet programs from running text such as '=SUM(...)' as a formula."""
    out = df.copy()
    for col in ("Category", "Description"):
        out[col] = out[col].apply(lambda v: "'" + v if isinstance(v, str) and v[:1] in "=+-@" else v)
    return out


def cleaned_csv_bytes(df):
    out = safe_for_spreadsheet(df)
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    return out.to_csv(index=False).encode("utf-8")


def build_summary_text(summary, currency, budget_info=None, month_text=None):
    lines = ["EXPENSE SUMMARY", "===============",
             f"Date range: {date_range_text(summary)}",
             f"Currency: {'not specified' if currency == 'None' else currency}",
             f"Total expenses: {money(summary['total'], currency)}",
             f"Average transaction: {money(summary['average'], currency)}",
             f"Number of transactions: {summary['count']}"]
    if summary["highest"]:
        h = summary["highest"]
        lines.append(f"Highest expense: {money(h['amount'], currency)} ({h['category']})")
    lines += ["", "Spending by category:"]
    for _, r in summary["by_category"].iterrows():
        lines.append(f"- {r['Category']}: {money(r['Total'], currency)} ({r['Share %']}%)")
    if budget_info:
        lines += ["", f"Budget check ({month_text or 'all data'}):",
                  f"- Budget: {money(budget_info['budget'], currency)}",
                  f"- Spent: {money(budget_info['spent'], currency)}",
                  f"- Remaining: {money(budget_info['remaining'], currency)}",
                  f"- Used: {budget_info['utilization']:.1f}%",
                  f"- Status: {budget_info['status']}"]
    return "\n".join(lines) + "\n"


def build_ai_report(ai, summary, currency):
    def bullets(items):
        return "\n".join(f"- {i}" for i in items) if items else "- Nothing specific noted."
    return (
        "# AI Expense Insights\n\n"
        f"**Date range:** {date_range_text(summary)}  \n"
        f"**Total expenses:** {money(summary['total'], currency)}  \n\n"
        "> These suggestions are AI-generated estimates. They are not financial advice "
        "and do not guarantee any savings.\n\n"
        f"## Summary\n{ai['summary']}\n\n"
        f"## Major spending categories\n{bullets(ai['major_categories'])}\n\n"
        f"## Spending worth reviewing\n{bullets(ai['unusual_spending'])}\n\n"
        f"## Repeated patterns\n{bullets(ai['repeated_patterns'])}\n\n"
        f"## Ideas to reduce expenses\n{bullets(ai['saving_ideas'])}\n\n"
        f"## Budget recommendation\n{ai['budget_recommendation']}\n"
    )


# ----------------------------------------------------------------------------
# 7. The Streamlit page
# ----------------------------------------------------------------------------
def show_welcome():
    st.info("Welcome! Upload an expense file in the sidebar to begin.")
    st.markdown(
        "**How to use this app**\n"
        "1. Upload a `.csv` or `.xlsx` file in the sidebar.\n"
        "2. Check that the columns are matched correctly.\n"
        "3. Look at your dashboard, charts and budget.\n"
        "4. Optionally click **Generate AI insights**.\n\n"
        "Your file is only used in memory while the app is open. It is not saved."
    )
    st.download_button("Download a sample CSV", SAMPLE_CSV, "sample_expenses.csv", "text/csv")


def sidebar_inputs():
    """Draw the sidebar and return the user's choices."""
    st.sidebar.header("1. Upload")
    uploaded = st.sidebar.file_uploader("Expense file (.csv or .xlsx)", type=["csv", "xlsx"],
                                        help=f"Maximum size: {MAX_UPLOAD_MB} MB")
    st.sidebar.header("2. Settings")
    currency = st.sidebar.selectbox("Currency (display only, no conversion)", CURRENCIES)
    refund_label = st.sidebar.radio(
        "Negative amounts (refunds)",
        ["Subtract them from spending", "Ignore them"],
        help="Subtract: a -500 refund lowers your total. Ignore: negative rows are removed.")
    remove_dupes = st.sidebar.checkbox(
        "Remove exact duplicate rows", value=True,
        help="Turn off if you really bought the same thing, same day, same price, twice.")
    budget = st.sidebar.number_input("Monthly budget (optional, 0 = none)", min_value=0.0,
                                     value=0.0, step=100.0)
    return {"uploaded": uploaded, "currency": currency,
            "refund_mode": "net" if refund_label.startswith("Subtract") else "exclude",
            "remove_dupes": remove_dupes, "budget": budget}


def column_mapping(raw):
    """Let the user match their column names to Date, Category, Description, Amount."""
    st.sidebar.header("3. Match columns")
    columns = list(raw.columns)
    mapping = {}
    for field in ["Amount", "Date", "Category", "Description"]:
        guess = guess_column(columns, field)
        options = columns if field == "Amount" else ["(none)"] + columns
        index = options.index(guess) if guess in options else 0
        choice = st.sidebar.selectbox(f"{field} column", options, index=index)
        mapping[field] = None if choice == "(none)" else choice
    return mapping


def show_dashboard(df, summary, report, currency):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total expenses", money(summary["total"], currency))
    c2.metric("Average transaction", money(summary["average"], currency))
    c3.metric("Transactions", f"{summary['count']:,}")
    if summary["highest"]:
        c4.metric("Highest expense", money(summary["highest"]["amount"], currency),
                  help=f"Category: {summary['highest']['category']}")
    else:
        c4.metric("Highest expense", "n/a")
    st.caption(f"Date range: {date_range_text(summary)}")

    with st.expander("Data cleaning report"):
        st.write(f"- Rows in file: {report['rows_in']}")
        st.write(f"- Rows removed (amount missing or not a number): {report['invalid_amount']}")
        st.write(f"- Rows with a missing or invalid date (kept, but left out of date charts): {report['invalid_date']}")
        st.write(f"- Rows with no category (set to 'Uncategorized'): {report['missing_category']}")
        st.write(f"- Duplicate rows removed: {report['duplicates_removed']}")
        st.write(f"- Negative amounts found: {report['negative_rows']} "
                 f"({'removed' if report['negative_excluded'] else 'counted as refunds'})")

    st.subheader("Spending by category")
    st.dataframe(summary["by_category"], width="stretch", hide_index=True)
    st.subheader("All transactions")
    st.caption("Click a column header to sort.")
    st.dataframe(df, width="stretch", hide_index=True)


def show_charts(df, summary, currency):
    pie = make_pie_chart(summary["by_category"], currency)
    bar = make_bar_chart(summary["by_category"], currency)
    line = make_line_chart(df, currency)
    left, right = st.columns(2)
    with left:
        if pie:
            st.plotly_chart(pie, width="stretch")
        else:
            st.info("No positive spending to show in the pie chart.")
    with right:
        if bar:
            st.plotly_chart(bar, width="stretch")
    if line:
        st.plotly_chart(line, width="stretch")
    else:
        st.info("No valid dates were found, so the time chart is not available.")


def show_budget(df, budget, currency):
    """Draw the budget section. Returns (budget_info, month_text) for reports."""
    st.subheader("Budget planning")
    if budget <= 0:
        st.info("Enter a monthly budget in the sidebar to see this section.")
        return None, None
    dated = df.dropna(subset=["Date"])
    if dated.empty:
        st.warning("No valid dates, so the budget is compared with the total of ALL uploaded rows.")
        spent, month_text = float(df["Amount"].sum()), None
    else:
        months = sorted(dated["Date"].dt.to_period("M").astype(str).unique())
        month_text = st.selectbox("Choose a month", months, index=len(months) - 1)
        spent = float(dated.loc[dated["Date"].dt.to_period("M").astype(str) == month_text, "Amount"].sum())
        warning = month_warning(df, month_text)
        if warning:
            st.warning(warning)
    info = calculate_budget(budget, spent)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Budget", money(info["budget"], currency))
    c2.metric("Spent", money(info["spent"], currency))
    c3.metric("Remaining", money(info["remaining"], currency))
    c4.metric("Used", f"{info['utilization']:.1f}%")
    st.progress(min(info["utilization"] / 100, 1.0))
    (st.error if info["status"] == "Over budget" else st.success)(info["status"])
    return info, month_text


def show_ai(summary, currency, budget_info):
    st.subheader("AI insights")
    st.caption("AI suggestions are estimates, not financial advice, and cannot guarantee savings.")
    api_key = get_setting("GEMINI_API_KEY")
    model_name = get_setting("GEMINI_MODEL") or DEFAULT_MODEL
    if not api_key:
        st.warning("No Gemini API key found, so AI insights are turned off. "
                   "The rest of the app works normally. See the README to add GEMINI_API_KEY.")
        return None

    payload = build_ai_payload(summary, currency, budget_info)
    st.info("Only totals, counts and category names are sent to Gemini. "
            "Descriptions and single-transaction dates are NOT sent.")
    with st.expander("See exactly what will be sent"):
        st.json(payload)

    cache_id = hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()
    if st.button("Generate AI insights"):
        with st.spinner("Gemini is reading your summary..."):
            try:
                st.session_state["ai"] = {"id": cache_id, "data": request_ai_insights(payload, api_key, model_name)}
                st.success("Insights are ready.")
            except AIError as err:
                st.session_state.pop("ai", None)
                st.error(str(err))

    saved = st.session_state.get("ai")
    if not saved or saved["id"] != cache_id:
        return None
    ai = saved["data"]
    st.markdown(f"**Summary:** {ai['summary']}")
    sections = [("Major spending categories", "major_categories"),
                ("Spending worth reviewing", "unusual_spending"),
                ("Repeated patterns", "repeated_patterns"),
                ("Ideas to reduce expenses", "saving_ideas")]
    for title, key in sections:
        st.markdown(f"**{title}**")
        for item in ai[key] or ["Nothing specific noted."]:
            st.markdown(f"- {item}")
    st.markdown(f"**Budget recommendation:** {ai['budget_recommendation']}")
    return ai


def show_downloads(df, summary, currency, budget_info, month_text, ai):
    st.subheader("Download reports")
    st.download_button("Cleaned data (CSV)", cleaned_csv_bytes(df), "cleaned_expenses.csv", "text/csv")
    st.download_button("Expense summary (text)",
                       build_summary_text(summary, currency, budget_info, month_text),
                       "expense_summary.txt", "text/plain")
    if ai:
        st.download_button("AI insights (Markdown)", build_ai_report(ai, summary, currency),
                           "ai_insights.md", "text/markdown")
    else:
        st.caption("Generate AI insights first to download the AI report.")


def main():
    st.set_page_config(page_title="AI Expense Analyzer", page_icon="💰", layout="wide")
    st.title("💰 AI Expense Analyzer")
    st.write("Upload your expenses, see where your money goes, plan a budget, and get friendly AI tips.")

    options = sidebar_inputs()
    if options["uploaded"] is None:
        show_welcome()
        return

    try:
        raw = load_file(options["uploaded"].name, options["uploaded"].getvalue())
    except ValueError as err:
        st.error(str(err))
        return

    st.subheader("File preview")
    st.dataframe(raw.head(10), width="stretch", hide_index=True)

    mapping = column_mapping(raw)
    try:
        df, report = clean_expenses(raw, mapping, options["remove_dupes"], options["refund_mode"])
    except ValueError as err:
        st.error(str(err))
        return
    st.success(f"Loaded {len(df)} valid transactions.")

    currency = options["currency"]
    summary = compute_summary(df)
    tab_dash, tab_charts, tab_budget, tab_ai, tab_dl = st.tabs(
        ["Dashboard", "Charts", "Budget", "AI insights", "Downloads"])
    with tab_dash:
        show_dashboard(df, summary, report, currency)
    with tab_charts:
        show_charts(df, summary, currency)
    with tab_budget:
        budget_info, month_text = show_budget(df, options["budget"], currency)
    with tab_ai:
        ai = show_ai(summary, currency, budget_info)
    with tab_dl:
        show_downloads(df, summary, currency, budget_info, month_text, ai)


if __name__ == "__main__":
    main()
