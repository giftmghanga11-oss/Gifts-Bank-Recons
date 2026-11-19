from io import BytesIO
import pandas as pd
import streamlit as st
from reconciliation import export_to_excel, run_reconciliation

st.set_page_config(page_title="Bank & ERP Reconciliation", layout="wide")
st.title("Bank vs. ERP Reconciliation Tool")
st.caption(
    "Opposite-side matching (Bank Debits ↔ ERP Credits, Bank Credits ↔ ERP Debits) | "
    "Sorted ascending by amount | Strict < KSh 1.00 tolerance | Variance = Bank − ERP"
)


def load_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


col1, col2 = st.columns(2)
with col1:
    bank_file = st.file_uploader("Upload Bank Statement (CSV or XLSX)", type=["csv", "xlsx", "xls"])
with col2:
    erp_file = st.file_uploader("Upload ERP Cashbook (CSV or XLSX)", type=["csv", "xlsx", "xls"])

if bank_file and erp_file:
    bank_df = load_uploaded_file(bank_file)
    erp_df = load_uploaded_file(erp_file)

    results = run_reconciliation(bank_df, erp_df)
    summary_df = results["Reconciliation Summary"]

    st.subheader("Reconciliation Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    excel_bytes = export_to_excel(results)
    st.download_button(
        label="Download Reconciliation Workbook (.xlsx)",
        data=excel_bytes,
        file_name="bank_erp_reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tabs = st.tabs([
        "Bank Debits vs ERP Credits",
        "Bank Credits vs ERP Debits",
        "Matched Transactions",
        "Bank-Only Transactions",
        "ERP-Only Transactions",
        "Investigation & Duplicates",
    ])

    tab_keys = [
        "Bank Debits vs ERP Credits",
        "Bank Credits vs ERP Debits",
        "Matched Transactions",
        "Bank-Only Transactions",
        "ERP-Only Transactions",
        "Investigation & Duplicates",
    ]

    for tab, key in zip(tabs, tab_keys):
        with tab:
            st.dataframe(results[key], use_container_width=True, hide_index=True)
