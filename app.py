import streamlit as st
import pandas as pd
from reconciliation import reconcile, normalize_transactions, build_reconciliation_excel

st.set_page_config(page_title="Bank Reconciliation", page_icon="🏦", layout="wide")
st.title("🏦 Automated Bank Reconciliation")
st.caption("Upload your bank statement and ERP cashbook, map the columns, and review side-by-side matched and outstanding transactions.")

with st.sidebar:
    st.header("Reconciliation settings")
    tolerance = st.number_input(
        "Maximum permitted difference (KSh)",
        min_value=0.0,
        value=1.0,
        step=0.1,
        format="%.2f",
        help="A difference must be STRICTLY LESS than this value. At KSh 1.00, a difference of exactly KSh 1.00 will not match."
    )
    st.markdown("**Active matching rules**")
    st.markdown(
        "- **Opposite sides:** Bank Debits ↔ ERP Credits, Bank Credits ↔ ERP Debits\n"
        "- **Sorting:** Each side sorted ascending by amount\n"
        "- **Criterion:** Amount only (ignores date & description differences)\n"
        "- **Variance:** `Bank Amount − ERP Amount` (original amounts preserved)"
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
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    return pd.read_excel(upload)


def mapping_controls(df, prefix):
    cols = list(df.columns)
    st.markdown(f"**{prefix} column mapping**")
    date_col = st.selectbox(f"{prefix}: date column", ["(none)"] + cols, key=f"{prefix}_date")
    ref_col = st.selectbox(f"{prefix}: reference column", ["(none)"] + cols, key=f"{prefix}_ref")
    desc_col = st.selectbox(f"{prefix}: description column", ["(none)"] + cols, key=f"{prefix}_desc")

    amount_mode = st.radio(
        f"{prefix}: amount format",
        ["Separate Debit & Credit columns", "Single Amount column"],
        horizontal=True,
        key=f"{prefix}_amt_mode"
    )

    if amount_mode == "Separate Debit & Credit columns":
        debit_col = st.selectbox(f"{prefix}: debit column", cols, key=f"{prefix}_debit")
        credit_col = st.selectbox(f"{prefix}: credit column", cols, key=f"{prefix}_credit")
        amount_col = None
        direction_col = None
    else:
        amount_col = st.selectbox(f"{prefix}: amount column", cols, key=f"{prefix}_amount")
        direction_col = st.selectbox(f"{prefix}: debit/credit indicator column (optional)", ["(none)"] + cols, key=f"{prefix}_direction")
        debit_col = None
        credit_col = None

    return {
        "date": None if date_col == "(none)" else date_col,
        "reference": None if ref_col == "(none)" else ref_col,
        "description": None if desc_col == "(none)" else desc_col,
        "debit": debit_col,
        "credit": credit_col,
        "amount": amount_col,
        "direction": None if direction_col in (None, "(none)") else direction_col,
    }


def style_reconciliation_df(df: pd.DataFrame):
    if df.empty or "Status" not in df.columns:
        return df

    def highlight_row(row):
        status = str(row.get("Status", ""))
        if status == "Bank Only":
            return ["background-color: #d9ead3; color: #1b4332"] * len(row)
        if status == "ERP Only":
            return ["background-color: #f4cccc; color: #7f1d1d"] * len(row)
        return [""] * len(row)

    return df.style.apply(highlight_row, axis=1)


if bank_file and cash_file:
    try:
        bank_df = read_upload(bank_file)
        cash_df = read_upload(cash_file)
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

        st.info("Upload Excel or CSV files from NCBA and Sage X3. Map either separate Debit/Credit columns or a single signed Amount column.")
        if st.button("Reconcile transactions", type="primary", use_container_width=True):
            try:
                bank_tx = normalize_transactions(bank_df, bank_map, "BANK")
                cash_tx = normalize_transactions(cash_df, cash_map, "ERP")
                results = reconcile(bank_tx, cash_tx, tolerance=tolerance)
                st.session_state["reconciliation"] = results
                st.success("Reconciliation completed. Review each result below.")
            except Exception as exc:
                st.error(f"Could not reconcile these files: {exc}")

        results = st.session_state.get("reconciliation")
        if results is not None:
            st.subheader("3. Reconciliation summary")
            summary = results["summary"]
            cols = st.columns(5)
            for col, label, value in zip(
                cols,
                ["Matched pairs", "Bank only (Green)", "ERP only (Red)", "Needs investigation", "Net variance (KSh)"],
                [
                    summary["matched"],
                    summary["bank_only"],
                    summary["erp_only"],
                    summary["potential_duplicates"],
                    f"{summary['net_variance']:,.2f}",
                ]
            ):
                col.metric(label, value)

            tabs = st.tabs([
                "Bank Debits vs ERP Credits",
                "Bank Credits vs ERP Debits",
                "Matched",
                "Bank only",
                "ERP only",
                "Needs investigation",
                "Summary totals",
            ])
            tab_keys = [
                "debits_vs_credits",
                "credits_vs_debits",
                "matched",
                "bank_only",
                "erp_only",
                "investigate",
                "summary_table",
            ]

            for tab, key in zip(tabs, tab_keys):
                with tab:
                    df = results[key]
                    st.dataframe(style_reconciliation_df(df), use_container_width=True, hide_index=True)
                    if not df.empty:
                        st.download_button(
                            f"Download {key.replace('_', ' ').title()} (CSV)",
                            df.to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"{key}_transactions.csv",
                            mime="text/csv",
                            key=f"download_{key}"
                        )

            excel_bytes = build_reconciliation_excel(results)
            st.download_button(
                "⬇️ Download complete reconciliation workbook",
                excel_bytes,
                file_name="bank_reconciliation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    except Exception as exc:
        st.error(f"Unable to read one of the uploaded files: {exc}")
else:
    st.markdown("Upload both files to begin. The app will let you map the date, reference, description, and debit/credit columns.")
    st.markdown("**Updated matching policy**")
    st.markdown(
        "- **Opposite accounting sides:** Bank Debits match against ERP Credits; Bank Credits match against ERP Debits.\n"
        "- **Amount-first sorting & alignment:** Each side is sorted from smallest to largest and aligned side-by-side with blank spaces on the opposite side for unmatched rows.\n"
        "- **Strict KSh 1.00 tolerance:** Differences strictly below KSh 1.00 match automatically while preserving actual variance (`Bank − ERP`).\n"
        "- **One-to-one control:** Each transaction matches at most once; duplicate leftovers and unmatched entries are flagged for investigation."
    )
