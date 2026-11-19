import streamlit as st
import pandas as pd
from reconciliation import (
    reconcile,
    normalize_transactions,
    repair_cashbook_dataframe,
    build_reconciliation_excel,
)

st.set_page_config(page_title="Bank Reconciliation", page_icon="🏦", layout="wide")
st.title("🏦 Automated Bank Reconciliation")
st.caption("Bank details on the left • Variance in the middle • Cashbook details on the right")

with st.sidebar:
    st.header("Reconciliation settings")
    tolerance = st.number_input(
        "Maximum permitted difference (KSh)",
        min_value=0.0,
        value=1.0,
        step=0.1,
        format="%.2f",
        help="Differences strictly less than this amount are matched automatically."
    )

st.subheader("1. Upload your files")
left, right = st.columns(2)
with left:
    bank_file = st.file_uploader("Bank statement", type=["xlsx", "xls", "csv"], key="bank")
with right:
    cash_file = st.file_uploader("ERP cashbook", type=["xlsx", "xls", "csv"], key="cash")


def read_upload(upload):
    if upload is None:
        return None
    if upload.name.lower().endswith(".csv"):
        return pd.read_csv(upload)
    return pd.read_excel(upload)


def mapping_controls(df, prefix):
    cols = list(df.columns)
    st.markdown(f"**{prefix} column mapping**")

    default_mode = 0 if ("Debit Amount" in cols and "Credit Amount" in cols) else 1
    amount_mode = st.radio(
        f"{prefix}: amount format",
        ["Separate Debit & Credit columns", "Single Amount column"],
        index=default_mode,
        horizontal=True,
        key=f"{prefix}_amt_mode"
    )

    if amount_mode == "Separate Debit & Credit columns":
        dr_idx = cols.index("Debit Amount") if "Debit Amount" in cols else 0
        cr_idx = cols.index("Credit Amount") if "Credit Amount" in cols else 0
        debit_col = st.selectbox(f"{prefix}: debit column", cols, index=dr_idx, key=f"{prefix}_debit")
        credit_col = st.selectbox(f"{prefix}: credit column", cols, index=cr_idx, key=f"{prefix}_credit")
        return {"debit": debit_col, "credit": credit_col, "amount": None, "direction": None}

    amt_default = (
        cols.index("Paid amount")
        if "Paid amount" in cols
        else (cols.index("Payment amount") if "Payment amount" in cols else 0)
    )
    amount_col = st.selectbox(f"{prefix}: amount column", cols, index=amt_default, key=f"{prefix}_amount")

    dir_options = ["(none)"] + cols
    dir_default = dir_options.index("Payment no.") if "Payment no." in dir_options else 0
    direction_col = st.selectbox(
        f"{prefix}: debit/credit indicator column (optional)",
        dir_options,
        index=dir_default,
        key=f"{prefix}_direction"
    )
    return {
        "debit": None,
        "credit": None,
        "amount": amount_col,
        "direction": None if direction_col == "(none)" else direction_col,
    }


def style_report(df: pd.DataFrame, statuses: list[str]):
    def row_style(row):
        st_val = statuses[row.name] if row.name < len(statuses) else ""
        if st_val == "Bank Only":
            return ["background-color: #d9ead3; color: #1b4332"] * len(row)
        if st_val == "ERP Only":
            return ["background-color: #f4cccc; color: #7f1d1d"] * len(row)
        return [""] * len(row)

    return df.style.apply(row_style, axis=1)


if bank_file and cash_file:
    try:
        bank_df = read_upload(bank_file)
        cash_df = repair_cashbook_dataframe(read_upload(cash_file))

        st.subheader("2. Check your data")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Bank statement preview**")
            st.dataframe(bank_df.head(8), use_container_width=True)
            bank_map = mapping_controls(bank_df, "Bank")
        with c2:
            st.markdown("**ERP cashbook preview**")
            st.dataframe(cash_df.head(8), use_container_width=True)
            cash_map = mapping_controls(cash_df, "Cashbook")

        if st.button("Reconcile transactions", type="primary", use_container_width=True):
            try:
                bank_tx = normalize_transactions(bank_df, bank_map, "BANK")
                cash_tx = normalize_transactions(cash_df, cash_map, "ERP")
                results = reconcile(bank_tx, cash_tx, tolerance=tolerance)
                st.session_state["reconciliation"] = results
                st.success("Reconciliation report generated.")
            except Exception as exc:
                st.error(f"Could not reconcile these files: {exc}")

        results = st.session_state.get("reconciliation")
        if results is not None:
            st.subheader("3. Reconciliation report")
            summary = results["summary"]
            m1, m2, m3 = st.columns(3)
            m1.metric("Matched pairs", summary["matched"])
            m2.metric("Bank only (Green)", summary["bank_only"])
            m3.metric("Cashbook only (Red)", summary["erp_only"])

            report_df = results["report_df"]
            st.dataframe(
                style_report(report_df, results["statuses"]),
                use_container_width=True,
                hide_index=True
            )

            excel_bytes = build_reconciliation_excel(results)
            st.download_button(
                "⬇️ Download reconciliation report (.xlsx)",
                excel_bytes,
                file_name="NCBA_Bank_rec.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    except Exception as exc:
        st.error(f"Unable to read one of the uploaded files: {exc}")
else:
    st.markdown("Upload both files to generate the side-by-side bank reconciliation report.")
