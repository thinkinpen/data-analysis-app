import io
import sys
import tempfile
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


APP_TITLE = "Data Analysis App"
SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xls")


def read_dataframe(file_obj: BinaryIO, filename: str) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame."""
    lowered = filename.lower()
    if lowered.endswith(".csv"):
        return pd.read_csv(file_obj)
    if lowered.endswith((".xlsx", ".xls")):
        return pd.read_excel(file_obj)
    raise ValueError("Unsupported file format. Please upload a CSV or Excel file.")


if st is not None:

    @st.cache_data
    def load_data(uploaded_file) -> pd.DataFrame:
        return read_dataframe(uploaded_file, uploaded_file.name)

else:

    def load_data(uploaded_file) -> pd.DataFrame:
        return read_dataframe(uploaded_file, getattr(uploaded_file, "name", "uploaded_file.csv"))


def apply_cleaning_options(
    df: pd.DataFrame,
    selected_columns: Optional[list[str]] = None,
    remove_duplicates: bool = False,
    fill_missing: str = "Do nothing",
) -> pd.DataFrame:
    """Apply column filtering and simple cleaning rules."""
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


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()


def save_dataframe(df: pd.DataFrame, path: str) -> None:
    target = Path(path)
    suffix = target.suffix.lower()
    if suffix == ".csv":
        df.to_csv(target, index=False)
        return
    if suffix == ".xlsx":
        target.write_bytes(dataframe_to_excel_bytes(df))
        return
    raise ValueError("Output file must end with .csv or .xlsx")


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


def run_streamlit_app() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title("📊 Data Analysis App")
    st.write("Upload a CSV or Excel file to clean, analyze, and visualize your data.")

    uploaded_file = st.file_uploader("Upload your file", type=["csv", "xlsx", "xls"])

    if uploaded_file is None:
        st.info("Upload a file to begin.")
        return

    try:
        df = load_data(uploaded_file)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return

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

    try:
        work_df = apply_cleaning_options(
            df,
            selected_columns=selected_columns,
            remove_duplicates=remove_duplicates,
            fill_missing=fill_missing,
        )
    except Exception as e:
        st.error(f"Could not clean data: {e}")
        return

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Preview", "Summary", "Missing Values", "Charts", "Export"]
    )

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

        try:
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
        except Exception as e:
            st.error(f"Could not render chart: {e}")

    with tab5:
        st.subheader("Export Cleaned Data")
        st.write("Download the cleaned or filtered version of your dataset.")
        try:
            excel_bytes = dataframe_to_excel_bytes(work_df)
            st.download_button(
                label="Download as Excel",
                data=excel_bytes,
                file_name="cleaned_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.warning(f"Excel export unavailable: {e}")

        st.download_button(
            label="Download as CSV",
            data=work_df.to_csv(index=False).encode("utf-8"),
            file_name="cleaned_data.csv",
            mime="text/csv",
        )


def run_cli_app(argv: list[str]) -> int:
    if len(argv) < 2:
        print(
            "Streamlit is not installed in this environment.\n"
            "Use one of these commands:\n"
            "  python app.py <input_file> [output_file]\n"
            "  python app.py --test\n"
            "Or install the full app dependencies and run:\n"
            "  pip install streamlit matplotlib pandas openpyxl\n"
            "  streamlit run app.py"
        )
        return 0

    input_path = Path(argv[1])
    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        return 1

    with input_path.open("rb") as f:
        df = read_dataframe(f, input_path.name)

    cleaned = apply_cleaning_options(df)

    print(f"{APP_TITLE}\n")
    print(f"Rows: {cleaned.shape[0]}")
    print(f"Columns: {cleaned.shape[1]}\n")
    print("Preview:")
    print(cleaned.head().to_string(index=False))
    print("\nMissing Values Report:")
    print(create_missing_report(cleaned).to_string(index=False))

    numeric_df = cleaned.select_dtypes(include="number")
    if not numeric_df.empty:
        print("\nSummary Statistics:")
        print(numeric_df.describe().T.to_string())

    if len(argv) >= 3:
        output_path = argv[2]
        save_dataframe(cleaned, output_path)
        print(f"\nSaved cleaned data to: {output_path}")

    return 0


def _run_self_tests() -> int:
    sample = pd.DataFrame(
        {
            "name": ["Ada", None, "Tolu", "Ada"],
            "score": [10, None, 30, 10],
            "group": ["A", "B", None, "A"],
        }
    )

    filtered = apply_cleaning_options(sample, selected_columns=["name", "score"])
    assert list(filtered.columns) == ["name", "score"]

    deduped = apply_cleaning_options(sample, remove_duplicates=True)
    assert deduped.shape[0] == 3

    zero_filled = apply_cleaning_options(sample, fill_missing="Fill numeric with 0")
    assert zero_filled.loc[1, "score"] == 0

    blank_filled = apply_cleaning_options(sample, fill_missing="Fill text with blank")
    assert blank_filled.loc[1, "name"] == ""
    assert blank_filled.loc[2, "group"] == ""

    dropped = apply_cleaning_options(sample, fill_missing="Drop rows with missing values")
    assert dropped.shape[0] == 2

    report = create_missing_report(sample)
    assert set(report.columns) == {"column", "missing_count", "missing_percent"}
    assert int(report.loc[report["column"] == "score", "missing_count"].iloc[0]) == 1

    csv_bytes = b"a,b\n1,2\n3,4\n"
    csv_df = read_dataframe(io.BytesIO(csv_bytes), "demo.csv")
    assert csv_df.shape == (2, 2)

    try:
        read_dataframe(io.BytesIO(b"test"), "demo.txt")
    except ValueError:
        pass
    else:
        raise AssertionError("Unsupported extension should raise ValueError")

    with tempfile.TemporaryDirectory() as tmpdir:
        csv_path = Path(tmpdir) / "out.csv"
        xlsx_path = Path(tmpdir) / "out.xlsx"
        save_dataframe(csv_df, str(csv_path))
        save_dataframe(csv_df, str(xlsx_path))
        assert csv_path.exists()
        assert xlsx_path.exists()

    print("All self-tests passed.")
    return 0


if __name__ == "__main__":
    if "--test" in sys.argv:
        raise SystemExit(_run_self_tests())

    if st is not None:
        run_streamlit_app()
    else:
        raise SystemExit(run_cli_app(sys.argv))
