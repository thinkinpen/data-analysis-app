import io
import re
import sys
import tempfile
from collections import Counter
from pathlib import Path
from typing import BinaryIO, Optional

import pandas as pd

try:
    import matplotlib.pyplot as plt
except Exception:  # pragma: no cover - optional dependency fallback
    plt = None

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency fallback
    PdfReader = None


APP_TITLE = "Financial Statement Analysis App"
SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xls", ".pdf")
STOP_WORDS = {
    "the", "and", "for", "that", "with", "you", "this", "from", "are", "was",
    "have", "has", "had", "but", "not", "all", "can", "will", "your", "into",
    "their", "there", "they", "them", "would", "could", "should", "about", "after",
    "before", "when", "where", "which", "while", "what", "how", "why", "been",
    "were", "being", "than", "then", "also", "such", "may", "any", "each", "our",
    "out", "use", "used", "using", "more", "most", "other", "some", "these", "those",
    "his", "her", "she", "him", "its", "it", "of", "to", "in", "on", "at", "by",
    "an", "a", "or", "as", "is", "be", "if", "we", "us", "do", "does", "did"
}

STATEMENT_ORDER = [
    "Revenue",
    "Cost of Sales",
    "Gross Profit",
    "Operating Expenses",
    "Operating Profit",
    "Finance Cost",
    "Profit Before Tax",
    "Tax Expense",
    "Profit After Tax",
    "Cash and Cash Equivalents",
    "Trade Receivables",
    "Inventory",
    "Other Current Assets",
    "Current Assets",
    "Property, Plant and Equipment",
    "Non-Current Assets",
    "Total Assets",
    "Trade Payables",
    "Short-Term Debt",
    "Other Current Liabilities",
    "Current Liabilities",
    "Long-Term Debt",
    "Non-Current Liabilities",
    "Total Liabilities",
    "Share Capital",
    "Retained Earnings",
    "Equity",
]

LINE_ITEM_ALIASES = {
    "Revenue": [
        "revenue", "turnover", "sales", "gross earnings", "interest income", "operating income"
    ],
    "Cost of Sales": [
        "cost of sales", "cost of goods sold", "cost of revenue", "direct cost", "cost of turnover"
    ],
    "Gross Profit": ["gross profit"],
    "Operating Expenses": [
        "operating expenses", "administrative expenses", "selling and distribution", "selling expenses",
        "distribution expenses", "other operating expenses", "general and administrative expenses"
    ],
    "Operating Profit": [
        "operating profit", "operating income", "profit from operations", "operating earnings"
    ],
    "Finance Cost": ["finance cost", "finance costs", "interest expense", "borrowing costs"],
    "Profit Before Tax": [
        "profit before tax", "profit before taxation", "profit before income tax", "pbt"
    ],
    "Tax Expense": ["tax expense", "income tax", "taxation", "current tax", "deferred tax"],
    "Profit After Tax": [
        "profit after tax", "profit for the year", "net profit", "profit attributable", "pat"
    ],
    "Cash and Cash Equivalents": ["cash and cash equivalents", "cash and bank", "cash balances"],
    "Trade Receivables": ["trade receivables", "accounts receivable", "receivables", "trade debtors"],
    "Inventory": ["inventory", "inventories", "stock in trade", "stock"],
    "Other Current Assets": ["other current assets", "prepayments", "advances", "due from related parties"],
    "Current Assets": ["current assets", "total current assets"],
    "Property, Plant and Equipment": [
        "property plant and equipment", "property, plant and equipment", "ppe", "fixed assets"
    ],
    "Non-Current Assets": ["non-current assets", "total non-current assets", "fixed assets total"],
    "Total Assets": ["total assets", "assets total"],
    "Trade Payables": ["trade payables", "accounts payable", "trade creditors", "payables"],
    "Short-Term Debt": ["short term debt", "short-term borrowings", "bank overdraft", "current borrowings"],
    "Other Current Liabilities": ["other current liabilities", "accruals", "due to related parties"],
    "Current Liabilities": ["current liabilities", "total current liabilities"],
    "Long-Term Debt": ["long term debt", "long-term borrowings", "non-current borrowings", "lease liabilities"],
    "Non-Current Liabilities": ["non-current liabilities", "total non-current liabilities"],
    "Total Liabilities": ["total liabilities", "liabilities total"],
    "Share Capital": ["share capital", "issued capital", "ordinary share capital"],
    "Retained Earnings": ["retained earnings", "retained profits", "accumulated losses", "reserves"],
    "Equity": ["equity", "shareholders' funds", "total equity", "net assets"],
}

SECTION_MAP = {
    "Revenue": "Income Statement",
    "Cost of Sales": "Income Statement",
    "Gross Profit": "Income Statement",
    "Operating Expenses": "Income Statement",
    "Operating Profit": "Income Statement",
    "Finance Cost": "Income Statement",
    "Profit Before Tax": "Income Statement",
    "Tax Expense": "Income Statement",
    "Profit After Tax": "Income Statement",
    "Cash and Cash Equivalents": "Statement of Financial Position",
    "Trade Receivables": "Statement of Financial Position",
    "Inventory": "Statement of Financial Position",
    "Other Current Assets": "Statement of Financial Position",
    "Current Assets": "Statement of Financial Position",
    "Property, Plant and Equipment": "Statement of Financial Position",
    "Non-Current Assets": "Statement of Financial Position",
    "Total Assets": "Statement of Financial Position",
    "Trade Payables": "Statement of Financial Position",
    "Short-Term Debt": "Statement of Financial Position",
    "Other Current Liabilities": "Statement of Financial Position",
    "Current Liabilities": "Statement of Financial Position",
    "Long-Term Debt": "Statement of Financial Position",
    "Non-Current Liabilities": "Statement of Financial Position",
    "Total Liabilities": "Statement of Financial Position",
    "Share Capital": "Statement of Financial Position",
    "Retained Earnings": "Statement of Financial Position",
    "Equity": "Statement of Financial Position",
}


def detect_file_type(filename: str) -> str:
    lowered = filename.lower()
    if lowered.endswith(".csv"):
        return "csv"
    if lowered.endswith((".xlsx", ".xls")):
        return "excel"
    if lowered.endswith(".pdf"):
        return "pdf"
    raise ValueError("Unsupported file format. Please upload a CSV, Excel, or PDF file.")



def list_excel_sheets(file_obj: BinaryIO) -> list[str]:
    file_obj.seek(0)
    workbook = pd.ExcelFile(file_obj)
    return workbook.sheet_names



def read_dataframe(file_obj: BinaryIO, filename: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    file_type = detect_file_type(filename)
    file_obj.seek(0)

    if file_type == "csv":
        return pd.read_csv(file_obj)
    if file_type == "excel":
        return pd.read_excel(file_obj, sheet_name=sheet_name or 0)

    raise ValueError("Tabular analysis is only available for CSV and Excel files.")



def extract_pdf_text(file_obj: BinaryIO) -> tuple[str, int]:
    if PdfReader is None:
        raise RuntimeError("PDF support is not installed. Add 'pypdf' to requirements.txt.")

    file_obj.seek(0)
    reader = PdfReader(file_obj)
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages).strip(), len(reader.pages)


if st is not None:

    @st.cache_data
    def load_data(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
        return read_dataframe(uploaded_file, uploaded_file.name, sheet_name=sheet_name)

    @st.cache_data
    def load_pdf_text(uploaded_file) -> tuple[str, int]:
        return extract_pdf_text(uploaded_file)

else:

    def load_data(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
        return read_dataframe(
            uploaded_file,
            getattr(uploaded_file, "name", "uploaded_file.csv"),
            sheet_name=sheet_name,
        )

    def load_pdf_text(uploaded_file) -> tuple[str, int]:
        return extract_pdf_text(uploaded_file)



def normalize_label(text: str) -> str:
    text = re.sub(r"[^a-z0-9\s]", " ", str(text).lower())
    text = re.sub(r"\s+", " ", text).strip()
    return text



def parse_number(value) -> Optional[float]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    if text in {"-", "—", "na", "n/a", "nil"}:
        return 0.0

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    text = text.replace(",", "").replace("₦", "").replace("$", "")
    text = text.replace("£", "").replace("€", "")
    text = re.sub(r"\s+", "", text)

    if text.endswith("%"):
        return None

    try:
        value = float(text)
    except ValueError:
        return None

    return -value if negative else value



def match_standard_label(label: str) -> Optional[str]:
    normalized = normalize_label(label)
    for standard_label, aliases in LINE_ITEM_ALIASES.items():
        for alias in aliases:
            if alias in normalized:
                return standard_label
    return None



def apply_cleaning_options(
    df: pd.DataFrame,
    selected_columns: Optional[list[str]] = None,
    remove_duplicates: bool = False,
    fill_missing: str = "Do nothing",
) -> pd.DataFrame:
    work_df = df.copy()

    if selected_columns:
        missing = [col for col in selected_columns if col not in work_df.columns]
        if missing:
            raise KeyError(f"Selected columns not found: {missing}")
        work_df = work_df[selected_columns].copy()

    if remove_duplicates:
        work_df = work_df.drop_duplicates()

    if fill_missing == "Drop rows with missing values":
        work_df = work_df.dropna()
    elif fill_missing == "Fill numeric with 0":
        numeric_cols = work_df.select_dtypes(include="number").columns
        work_df.loc[:, numeric_cols] = work_df[numeric_cols].fillna(0)
    elif fill_missing == "Fill text with blank":
        text_cols = work_df.select_dtypes(exclude="number").columns
        work_df.loc[:, text_cols] = work_df[text_cols].fillna("")
    elif fill_missing != "Do nothing":
        raise ValueError(f"Unknown fill_missing option: {fill_missing}")

    return work_df



def create_missing_report(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "column": df.columns,
            "missing_count": df.isna().sum().values,
            "missing_percent": (df.isna().mean().values * 100).round(2),
        }
    )



def tokenize_text(text: str) -> list[str]:
    return re.findall(r"[A-Za-z]{3,}", text.lower())



def summarize_pdf_text(text: str, page_count: int) -> dict:
    words = tokenize_text(text)
    filtered_words = [word for word in words if word not in STOP_WORDS]
    top_words = Counter(filtered_words).most_common(15)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    return {
        "pages": page_count,
        "characters": len(text),
        "words": len(words),
        "lines": len(lines),
        "top_words": pd.DataFrame(top_words, columns=["word", "count"]),
        "preview": "\n".join(lines[:30]) if lines else "",
    }



def extract_financial_rows_from_pdf(text: str) -> pd.DataFrame:
    rows: list[dict] = []
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            continue

        numbers = re.findall(r"\(?\d[\d,\.]*\)?", line)
        if len(numbers) < 2:
            continue

        standard_label = match_standard_label(line)
        if not standard_label:
            continue

        parsed_numbers = [parse_number(item) for item in numbers]
        parsed_numbers = [item for item in parsed_numbers if item is not None]
        if len(parsed_numbers) < 2:
            continue

        current_value = parsed_numbers[-1]
        prior_value = parsed_numbers[-2]
        rows.append(
            {
                "source_label": line,
                "standard_label": standard_label,
                "current_year": current_value,
                "prior_year": prior_value,
                "source": "PDF",
            }
        )

    return pd.DataFrame(rows)



def standardize_excel_financials(df: pd.DataFrame) -> pd.DataFrame:
    work_df = df.copy()
    work_df.columns = [str(col).strip() for col in work_df.columns]

    label_col = None
    for candidate in work_df.columns:
        normalized = normalize_label(candidate)
        if normalized in {"line item", "lineitem", "account", "description", "particulars", "item"}:
            label_col = candidate
            break
    if label_col is None:
        label_col = work_df.columns[0]

    numeric_candidates = []
    for col in work_df.columns:
        if col == label_col:
            continue
        parsed = work_df[col].map(parse_number)
        if parsed.notna().sum() > 0:
            numeric_candidates.append(col)

    if len(numeric_candidates) < 2:
        raise ValueError(
            "Excel file should contain a line-item column and at least two columns of amounts for current and prior year."
        )

    current_col = numeric_candidates[-1]
    prior_col = numeric_candidates[-2]

    rows: list[dict] = []
    for _, row in work_df.iterrows():
        label = row.get(label_col)
        standard_label = match_standard_label(label)
        if not standard_label:
            continue

        current_value = parse_number(row.get(current_col))
        prior_value = parse_number(row.get(prior_col))
        if current_value is None and prior_value is None:
            continue

        rows.append(
            {
                "source_label": str(label),
                "standard_label": standard_label,
                "current_year": current_value or 0.0,
                "prior_year": prior_value or 0.0,
                "source": "Excel",
            }
        )

    return pd.DataFrame(rows)



def coalesce_standardized_rows(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty:
        return pd.DataFrame(columns=["section", "line_item", "current_year", "prior_year"])

    grouped = (
        raw_df.groupby("standard_label", as_index=False)[["current_year", "prior_year"]]
        .sum()
        .rename(columns={"standard_label": "line_item"})
    )
    grouped["section"] = grouped["line_item"].map(SECTION_MAP).fillna("Other")
    grouped["order"] = grouped["line_item"].map({item: idx for idx, item in enumerate(STATEMENT_ORDER)})
    grouped["order"] = grouped["order"].fillna(9999)
    grouped = grouped.sort_values(["section", "order", "line_item"]).drop(columns=["order"])
    return grouped[["section", "line_item", "current_year", "prior_year"]].reset_index(drop=True)



def financial_value(statement_df: pd.DataFrame, line_item: str) -> float:
    series = statement_df.loc[statement_df["line_item"] == line_item, "current_year"]
    return float(series.iloc[0]) if not series.empty else 0.0



def prior_financial_value(statement_df: pd.DataFrame, line_item: str) -> float:
    series = statement_df.loc[statement_df["line_item"] == line_item, "prior_year"]
    return float(series.iloc[0]) if not series.empty else 0.0



def safe_divide(numerator: float, denominator: float) -> Optional[float]:
    if denominator in (0, 0.0):
        return None
    return numerator / denominator



def percentage_change(current: float, prior: float) -> Optional[float]:
    if prior in (0, 0.0):
        return None
    return (current - prior) / abs(prior)



def build_variance_table(statement_df: pd.DataFrame) -> pd.DataFrame:
    work = statement_df.copy()
    work["absolute_change"] = work["current_year"] - work["prior_year"]
    work["percentage_change"] = work.apply(
        lambda row: percentage_change(row["current_year"], row["prior_year"]), axis=1
    )
    return work



def build_ratio_table(statement_df: pd.DataFrame) -> pd.DataFrame:
    revenue = financial_value(statement_df, "Revenue")
    prior_revenue = prior_financial_value(statement_df, "Revenue")
    gross_profit = financial_value(statement_df, "Gross Profit")
    operating_profit = financial_value(statement_df, "Operating Profit")
    pbt = financial_value(statement_df, "Profit Before Tax")
    pat = financial_value(statement_df, "Profit After Tax")
    current_assets = financial_value(statement_df, "Current Assets")
    inventory = financial_value(statement_df, "Inventory")
    current_liabilities = financial_value(statement_df, "Current Liabilities")
    total_liabilities = financial_value(statement_df, "Total Liabilities")
    equity = financial_value(statement_df, "Equity")
    total_assets = financial_value(statement_df, "Total Assets")
    receivables = financial_value(statement_df, "Trade Receivables")
    payables = financial_value(statement_df, "Trade Payables")
    cost_of_sales = abs(financial_value(statement_df, "Cost of Sales"))

    ratios = [
        {
            "ratio": "Revenue Growth",
            "value": percentage_change(revenue, prior_revenue),
            "format": "percent",
        },
        {
            "ratio": "Gross Margin",
            "value": safe_divide(gross_profit, revenue),
            "format": "percent",
        },
        {
            "ratio": "Operating Margin",
            "value": safe_divide(operating_profit, revenue),
            "format": "percent",
        },
        {
            "ratio": "Net Margin",
            "value": safe_divide(pat, revenue),
            "format": "percent",
        },
        {
            "ratio": "Current Ratio",
            "value": safe_divide(current_assets, current_liabilities),
            "format": "times",
        },
        {
            "ratio": "Quick Ratio",
            "value": safe_divide(current_assets - inventory, current_liabilities),
            "format": "times",
        },
        {
            "ratio": "Debt to Equity",
            "value": safe_divide(total_liabilities, equity),
            "format": "times",
        },
        {
            "ratio": "Return on Assets",
            "value": safe_divide(pat, total_assets),
            "format": "percent",
        },
        {
            "ratio": "Return on Equity",
            "value": safe_divide(pat, equity),
            "format": "percent",
        },
        {
            "ratio": "Receivables Days",
            "value": safe_divide(receivables * 365, revenue),
            "format": "days",
        },
        {
            "ratio": "Inventory Days",
            "value": safe_divide(inventory * 365, cost_of_sales),
            "format": "days",
        },
        {
            "ratio": "Payables Days",
            "value": safe_divide(payables * 365, cost_of_sales),
            "format": "days",
        },
    ]
    return pd.DataFrame(ratios)



def format_metric(value: Optional[float], metric_format: str) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    if metric_format == "percent":
        return f"{value:.1%}"
    if metric_format == "days":
        return f"{value:.1f} days"
    if metric_format == "times":
        return f"{value:.2f}x"
    return f"{value:,.2f}"



def build_professional_commentary(statement_df: pd.DataFrame, ratio_df: pd.DataFrame) -> pd.DataFrame:
    comments: list[dict] = []

    variance_df = build_variance_table(statement_df)
    material_lines = variance_df.dropna(subset=["percentage_change"]).copy()
    material_lines["abs_pct"] = material_lines["percentage_change"].abs()
    material_lines = material_lines.sort_values("abs_pct", ascending=False)

    revenue = financial_value(statement_df, "Revenue")
    prior_revenue = prior_financial_value(statement_df, "Revenue")
    pat = financial_value(statement_df, "Profit After Tax")
    prior_pat = prior_financial_value(statement_df, "Profit After Tax")
    receivables = financial_value(statement_df, "Trade Receivables")
    prior_receivables = prior_financial_value(statement_df, "Trade Receivables")
    current_ratio = ratio_df.loc[ratio_df["ratio"] == "Current Ratio", "value"]
    debt_to_equity = ratio_df.loc[ratio_df["ratio"] == "Debt to Equity", "value"]

    revenue_growth = percentage_change(revenue, prior_revenue)
    if revenue_growth is not None:
        direction = "increased" if revenue_growth >= 0 else "declined"
        comments.append(
            {
                "category": "Performance",
                "observation": f"Revenue {direction} by {abs(revenue_growth):.1%} year on year.",
            }
        )

    pat_growth = percentage_change(pat, prior_pat)
    if pat_growth is not None:
        direction = "increased" if pat_growth >= 0 else "declined"
        comments.append(
            {
                "category": "Profitability",
                "observation": f"Profit after tax {direction} by {abs(pat_growth):.1%} relative to the prior year.",
            }
        )

    receivable_growth = percentage_change(receivables, prior_receivables)
    if revenue_growth is not None and receivable_growth is not None and receivable_growth > revenue_growth + 0.10:
        comments.append(
            {
                "category": "Working Capital",
                "observation": "Trade receivables grew faster than revenue, which may indicate collection pressure or slower customer conversion.",
            }
        )

    if not current_ratio.empty and current_ratio.iloc[0] is not None:
        value = current_ratio.iloc[0]
        if value < 1:
            comments.append(
                {
                    "category": "Liquidity",
                    "observation": f"Current ratio is {value:.2f}x, suggesting current obligations exceed near-term asset cover.",
                }
            )
        else:
            comments.append(
                {
                    "category": "Liquidity",
                    "observation": f"Current ratio is {value:.2f}x, indicating current assets cover current liabilities.",
                }
            )

    if not debt_to_equity.empty and debt_to_equity.iloc[0] is not None:
        value = debt_to_equity.iloc[0]
        if value > 2:
            comments.append(
                {
                    "category": "Leverage",
                    "observation": f"Debt to equity stands at {value:.2f}x, indicating a relatively leveraged capital structure.",
                }
            )

    for _, row in material_lines.head(5).iterrows():
        line_item = row["line_item"]
        if line_item in {"Revenue", "Profit After Tax"}:
            continue
        change = row["percentage_change"]
        if change is None:
            continue
        direction = "increased" if change >= 0 else "decreased"
        comments.append(
            {
                "category": "Movement",
                "observation": f"{line_item} {direction} by {abs(change):.1%} year on year.",
            }
        )

    if not comments:
        comments.append(
            {
                "category": "General",
                "observation": "The extracted dataset is too limited for robust professional commentary. Consider a cleaner source file or a more text-readable PDF.",
            }
        )

    return pd.DataFrame(comments)



def build_analysis_package_from_table(df: pd.DataFrame) -> dict:
    standardized_raw = standardize_excel_financials(df)
    standardized = coalesce_standardized_rows(standardized_raw)
    if standardized.empty:
        raise ValueError(
            "No recognizable financial statement line items were found. The Excel file should contain rows like Revenue, Profit Before Tax, Total Assets, Equity, Trade Receivables, Current Liabilities, and similar items."
        )

    variance = build_variance_table(standardized)
    ratios = build_ratio_table(standardized)
    commentary = build_professional_commentary(standardized, ratios)

    return {
        "raw_extracted": standardized_raw,
        "standardized": standardized,
        "variance": variance,
        "ratios": ratios,
        "commentary": commentary,
    }



def build_analysis_package_from_pdf(text: str) -> dict:
    raw_extracted = extract_financial_rows_from_pdf(text)
    standardized = coalesce_standardized_rows(raw_extracted)
    if standardized.empty:
        raise ValueError(
            "No recognizable financial statement values were extracted from the PDF. Use a text-based audited financial statement, not a scanned image, or upload the figures in Excel."
        )

    variance = build_variance_table(standardized)
    ratios = build_ratio_table(standardized)
    commentary = build_professional_commentary(standardized, ratios)
    pdf_summary = summarize_pdf_text(text, page_count=max(text.count("\f") + 1, 1))

    return {
        "raw_extracted": raw_extracted,
        "standardized": standardized,
        "variance": variance,
        "ratios": ratios,
        "commentary": commentary,
        "pdf_summary": pdf_summary,
    }



def analysis_package_to_excel_bytes(package: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, key in [
            ("Raw_Extracted", "raw_extracted"),
            ("Standardized_FS", "standardized"),
            ("Variance_Analysis", "variance"),
            ("Ratios", "ratios"),
            ("Commentary", "commentary"),
        ]:
            package[key].to_excel(writer, index=False, sheet_name=sheet_name)

        if "pdf_summary" in package:
            pd.DataFrame(
                {
                    "metric": ["pages", "words", "lines", "characters"],
                    "value": [
                        package["pdf_summary"]["pages"],
                        package["pdf_summary"]["words"],
                        package["pdf_summary"]["lines"],
                        package["pdf_summary"]["characters"],
                    ],
                }
            ).to_excel(writer, index=False, sheet_name="PDF_Summary")

    return output.getvalue()



def render_histogram(df: pd.DataFrame, column: str):
    if plt is None:
        raise RuntimeError("matplotlib is not available in this environment.")
    fig, ax = plt.subplots()
    ax.hist(df[column].dropna())
    ax.set_title(f"Histogram of {column}")
    ax.set_xlabel(column)
    ax.set_ylabel("Frequency")
    return fig



def render_bar_chart(df: pd.DataFrame, column: str):
    if plt is None:
        raise RuntimeError("matplotlib is not available in this environment.")
    value_counts = df[column].astype(str).value_counts().head(20)
    fig, ax = plt.subplots()
    value_counts.plot(kind="bar", ax=ax)
    ax.set_title(f"Bar Chart of {column}")
    ax.set_xlabel(column)
    ax.set_ylabel("Count")
    return fig



def render_line_chart(df: pd.DataFrame, column: str):
    if plt is None:
        raise RuntimeError("matplotlib is not available in this environment.")
    fig, ax = plt.subplots()
    ax.plot(df[column].dropna().reset_index(drop=True))
    ax.set_title(f"Line Chart of {column}")
    ax.set_xlabel("Index")
    ax.set_ylabel(column)
    return fig



def render_scatter_plot(df: pd.DataFrame, x_col: str, y_col: str):
    if plt is None:
        raise RuntimeError("matplotlib is not available in this environment.")
    fig, ax = plt.subplots()
    valid = df[[x_col, y_col]].dropna()
    ax.scatter(valid[x_col], valid[y_col])
    ax.set_title(f"Scatter Plot: {x_col} vs {y_col}")
    ax.set_xlabel(x_col)
    ax.set_ylabel(y_col)
    return fig



def render_generic_table_analysis(df: pd.DataFrame) -> None:
    st.success("File uploaded successfully.")

    with st.sidebar:
        st.header("Controls")
        st.write(f"Rows: {df.shape[0]}")
        st.write(f"Columns: {df.shape[1]}")

        selected_columns = st.multiselect(
            "Select columns to view",
            options=list(df.columns),
            default=list(df.columns),
        )
        remove_duplicates = st.checkbox("Remove duplicate rows")
        fill_missing = st.selectbox(
            "Handle missing values",
            [
                "Do nothing",
                "Drop rows with missing values",
                "Fill numeric with 0",
                "Fill text with blank",
            ],
        )

    work_df = apply_cleaning_options(
        df,
        selected_columns=selected_columns,
        remove_duplicates=remove_duplicates,
        fill_missing=fill_missing,
    )

    tab1, tab2, tab3, tab4 = st.tabs(["Preview", "Summary", "Missing Values", "Charts"])

    with tab1:
        st.subheader("Data Preview")
        st.dataframe(work_df, use_container_width=True)

    with tab2:
        st.subheader("Summary Statistics")
        numeric_df = work_df.select_dtypes(include="number")
        if numeric_df.empty:
            st.warning("No numeric columns available for summary statistics.")
        else:
            st.dataframe(numeric_df.describe().T, use_container_width=True)

    with tab3:
        st.subheader("Missing Values Report")
        st.dataframe(create_missing_report(work_df), use_container_width=True)

    with tab4:
        st.subheader("Charts")
        numeric_columns = list(work_df.select_dtypes(include="number").columns)
        categorical_columns = list(work_df.select_dtypes(exclude="number").columns)

        chart_type = st.selectbox(
            "Choose chart type", ["Bar Chart", "Histogram", "Line Chart", "Scatter Plot"]
        )

        if chart_type == "Histogram":
            if not numeric_columns:
                st.warning("No numeric columns available.")
            else:
                hist_col = st.selectbox("Select numeric column", numeric_columns)
                st.pyplot(render_histogram(work_df, hist_col))

        elif chart_type == "Bar Chart":
            if not categorical_columns:
                st.warning("No categorical columns available.")
            else:
                cat_col = st.selectbox("Select category column", categorical_columns)
                st.pyplot(render_bar_chart(work_df, cat_col))

        elif chart_type == "Line Chart":
            if not numeric_columns:
                st.warning("No numeric columns available.")
            else:
                line_col = st.selectbox("Select numeric column", numeric_columns)
                st.pyplot(render_line_chart(work_df, line_col))

        elif chart_type == "Scatter Plot":
            if len(numeric_columns) < 2:
                st.warning("At least two numeric columns are needed for a scatter plot.")
            else:
                x_col = st.selectbox("X-axis", numeric_columns, index=0)
                default_y_index = 1 if len(numeric_columns) > 1 else 0
                y_col = st.selectbox("Y-axis", numeric_columns, index=default_y_index)
                st.pyplot(render_scatter_plot(work_df, x_col, y_col))



def render_financial_statement_analysis(package: dict) -> None:
    standardized = package["standardized"]
    variance = package["variance"]
    ratios = package["ratios"]
    commentary = package["commentary"]
    excel_bytes = analysis_package_to_excel_bytes(package)

    st.success("Financial statement analysis completed.")
    st.download_button(
        label="Download financial analysis workbook",
        data=excel_bytes,
        file_name="financial_statement_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Standardized Statements",
        "Variance Analysis",
        "Ratios",
        "Professional Commentary",
    ])

    with tab1:
        st.subheader("Standardized Financial Statements")
        for section in ["Income Statement", "Statement of Financial Position"]:
            st.markdown(f"**{section}**")
            section_df = standardized[standardized["section"] == section].copy()
            st.dataframe(section_df, use_container_width=True)

    with tab2:
        st.subheader("Year-on-Year Variance")
        display_df = variance.copy()
        display_df["percentage_change"] = display_df["percentage_change"].map(
            lambda x: None if pd.isna(x) else f"{x:.1%}"
        )
        st.dataframe(display_df, use_container_width=True)

    with tab3:
        st.subheader("Key Financial Ratios")
        display_ratios = ratios.copy()
        display_ratios["formatted_value"] = display_ratios.apply(
            lambda row: format_metric(row["value"], row["format"]), axis=1
        )
        st.dataframe(display_ratios[["ratio", "formatted_value"]], use_container_width=True)

    with tab4:
        st.subheader("Professional Commentary")
        for _, row in commentary.iterrows():
            st.markdown(f"**{row['category']}**: {row['observation']}")



def render_pdf_analysis(text: str, page_count: int) -> None:
    summary = summarize_pdf_text(text, page_count)
    st.success("PDF uploaded successfully.")

    tab1, tab2, tab3 = st.tabs(["Overview", "Text Preview", "Top Words"])

    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Pages", summary["pages"])
        col2.metric("Words", summary["words"])
        col3.metric("Lines", summary["lines"])
        col4.metric("Characters", summary["characters"])

        if summary["words"] == 0:
            st.warning("No readable text was found in this PDF. It may be scanned images rather than text.")

    with tab2:
        st.subheader("Extracted Text Preview")
        st.text_area("PDF text", summary["preview"] or "No readable text found.", height=400)

    with tab3:
        st.subheader("Most Common Words")
        if summary["top_words"].empty:
            st.info("No keywords found.")
        else:
            st.dataframe(summary["top_words"], use_container_width=True)



def run_streamlit_app() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title("📊 Financial Statement Analysis App")
    st.write(
        "Upload an audited financial statement PDF or an Excel file with line items and two years of amounts. The app standardizes key accounts, calculates ratios, produces variance analysis, and exports an Excel workpaper."
    )

    analysis_mode = st.sidebar.radio(
        "Analysis Mode",
        ["Financial Statement Analysis", "Generic File Analysis"],
        index=0,
    )

    uploaded_file = st.file_uploader("Upload your file", type=["csv", "xlsx", "xls", "pdf"])

    if uploaded_file is None:
        st.info("Upload a file to begin.")
        return

    file_type = detect_file_type(uploaded_file.name)

    try:
        if analysis_mode == "Financial Statement Analysis":
            if file_type == "pdf":
                text, page_count = load_pdf_text(uploaded_file)
                package = build_analysis_package_from_pdf(text)
                with st.expander("Preview extracted PDF text"):
                    render_pdf_analysis(text, page_count)
                render_financial_statement_analysis(package)
                return

            sheet_name = None
            if file_type == "excel":
                sheet_names = list_excel_sheets(uploaded_file)
                sheet_name = st.sidebar.selectbox("Select Excel sheet", sheet_names)

            df = load_data(uploaded_file, sheet_name=sheet_name)
            package = build_analysis_package_from_table(df)
            render_financial_statement_analysis(package)
            return

        if file_type == "pdf":
            text, page_count = load_pdf_text(uploaded_file)
            render_pdf_analysis(text, page_count)
            return

        sheet_name = None
        if file_type == "excel":
            sheet_names = list_excel_sheets(uploaded_file)
            sheet_name = st.sidebar.selectbox("Select Excel sheet", sheet_names)

        df = load_data(uploaded_file, sheet_name=sheet_name)
        render_generic_table_analysis(df)
    except Exception as e:
        st.error(f"Could not analyze file: {e}")



def run_cli_app(argv: list[str]) -> int:
    if len(argv) < 2:
        print(
            "Streamlit is not installed in this environment.\n"
            "Use one of these commands:\n"
            "  python app.py <input_file>\n"
            "  python app.py --test\n"
            "Or install the full app dependencies and run:\n"
            "  pip install streamlit matplotlib pandas openpyxl pypdf\n"
            "  streamlit run app.py"
        )
        return 0

    input_path = Path(argv[1])
    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        return 1

    file_type = detect_file_type(input_path.name)
    with input_path.open("rb") as f:
        if file_type == "pdf":
            text, _page_count = extract_pdf_text(f)
            package = build_analysis_package_from_pdf(text)
        else:
            df = read_dataframe(f, input_path.name)
            package = build_analysis_package_from_table(df)

    print(f"{APP_TITLE}\n")
    print("Standardized Financial Statements:")
    print(package["standardized"].to_string(index=False))
    print("\nRatios:")
    print(package["ratios"].assign(formatted=lambda x: x.apply(lambda row: format_metric(row['value'], row['format']), axis=1))[["ratio", "formatted"]].to_string(index=False))
    print("\nProfessional Commentary:")
    print(package["commentary"].to_string(index=False))
    return 0



def _run_self_tests() -> int:
    sample = pd.DataFrame(
        {
            "Particulars": [
                "Revenue",
                "Cost of Sales",
                "Gross Profit",
                "Operating Expenses",
                "Operating Profit",
                "Finance Cost",
                "Profit Before Tax",
                "Tax Expense",
                "Profit After Tax",
                "Cash and Cash Equivalents",
                "Trade Receivables",
                "Inventory",
                "Current Assets",
                "Total Assets",
                "Trade Payables",
                "Current Liabilities",
                "Total Liabilities",
                "Equity",
            ],
            "2024": [1000, -600, 400, -150, 250, -20, 230, -70, 160, 90, 140, 120, 420, 950, 110, 260, 500, 450],
            "2023": [900, -540, 360, -130, 230, -18, 212, -62, 150, 80, 100, 110, 380, 900, 95, 240, 470, 430],
        }
    )

    standardized_raw = standardize_excel_financials(sample)
    assert not standardized_raw.empty
    assert "Revenue" in standardized_raw["standard_label"].values

    package = build_analysis_package_from_table(sample)
    standardized = package["standardized"]
    ratios = package["ratios"]
    commentary = package["commentary"]
    variance = package["variance"]

    assert not standardized.empty
    assert not ratios.empty
    assert not commentary.empty
    assert not variance.empty

    gross_margin = ratios.loc[ratios["ratio"] == "Gross Margin", "value"].iloc[0]
    assert abs(gross_margin - 0.4) < 1e-9

    current_ratio = ratios.loc[ratios["ratio"] == "Current Ratio", "value"].iloc[0]
    assert abs(current_ratio - (420 / 260)) < 1e-9

    revenue_change = variance.loc[variance["line_item"] == "Revenue", "percentage_change"].iloc[0]
    assert abs(revenue_change - ((1000 - 900) / 900)) < 1e-9

    parsed_negative = parse_number("(1,250)")
    assert parsed_negative == -1250.0
    parsed_positive = parse_number("2,500")
    assert parsed_positive == 2500.0

    pdf_text = """
    Revenue 900 1000
    Cost of Sales (540) (600)
    Gross Profit 360 400
    Operating Expenses (130) (150)
    Operating Profit 230 250
    Finance Cost (18) (20)
    Profit Before Tax 212 230
    Tax Expense (62) (70)
    Profit After Tax 150 160
    Trade Receivables 100 140
    Current Assets 380 420
    Current Liabilities 240 260
    Total Liabilities 470 500
    Equity 430 450
    Total Assets 900 950
    """
    pdf_rows = extract_financial_rows_from_pdf(pdf_text)
    assert not pdf_rows.empty
    assert "Revenue" in pdf_rows["standard_label"].values

    pdf_package = build_analysis_package_from_pdf(pdf_text)
    assert not pdf_package["standardized"].empty

    excel_bytes = analysis_package_to_excel_bytes(package)
    assert isinstance(excel_bytes, bytes)
    assert len(excel_bytes) > 0

    tokens = tokenize_text("Tax law tax audit finance finance finance")
    assert "finance" in tokens
    summary = summarize_pdf_text("Tax law tax audit finance finance finance", 2)
    assert summary["pages"] == 2
    assert int(summary["top_words"].iloc[0]["count"]) >= 1

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        pd.DataFrame({"a": [1, 2]}).to_excel(writer, index=False, sheet_name="First")
        pd.DataFrame({"b": [3, 4]}).to_excel(writer, index=False, sheet_name="Second")
    excel_buffer.seek(0)
    assert list_excel_sheets(excel_buffer) == ["First", "Second"]

    try:
        detect_file_type("demo.txt")
    except ValueError:
        pass
    else:
        raise AssertionError("Unsupported extension should raise ValueError")

    generic_df = apply_cleaning_options(sample, selected_columns=["Particulars", "2024"], fill_missing="Do nothing")
    assert list(generic_df.columns) == ["Particulars", "2024"]

    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / "analysis.xlsx"
        out_path.write_bytes(excel_bytes)
        assert out_path.exists()

    print("All self-tests passed.")
    return 0


if __name__ == "__main__":
    if "--test" in sys.argv:
        raise SystemExit(_run_self_tests())

    if st is not None:
        run_streamlit_app()
    else:
        raise SystemExit(run_cli_app(sys.argv))
