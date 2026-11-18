import streamlit as st
import pandas as pd
from io import BytesIO
from reconciliation import reconcile, normalize_transactions

st.set_page_config(page_title="Bank Reconciliation", page_icon="🏦", layout="wide")
st.title("🏦 Automated Bank Reconciliation")
st.caption("Upload your bank statement and ERP cashbook, map the columns, and review matched and unmatched transactions.")

with st.sidebar:
    st.header("Reconciliation settings")
    tolerance = st.number_input("Maximum permitted difference (KSh)", min_value=0.0, value=1.0, step=0.1, format="%.2f",
                                help="A difference must be STRICTLY LESS than this value. At KSh 1.00, a difference of exactly KSh 1.00 will not match.")
    date_window = st.number_input("Date matching window (days)", min_value=0, max_value=60, value=5,
                                  help="Transactions further apart than this will not be automatically matched. Set to 0 to require the same date.")
    use_date = st.checkbox("Use transaction dates when matching", value=True)
    use_description = st.checkbox("Use descriptions/references to rank possible matches", value=True)

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
    amount_col = st.selectbox(f"{prefix}: amount column", cols, key=f"{prefix}_amount")
    desc_col = st.selectbox(f"{prefix}: description/reference column", ["(none)"] + cols, key=f"{prefix}_desc")
    direction_col = st.selectbox(f"{prefix}: debit/credit column (optional)", ["(none)"] + cols, key=f"{prefix}_direction")
    return {
        "date": None if date_col == "(none)" else date_col,
        "amount": amount_col,
        "description": None if desc_col == "(none)" else desc_col,
        "direction": None if direction_col == "(none)" else direction_col,
    }

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

        st.info("For the first version, upload Excel or CSV files. PDF statement extraction and saved bank-specific layouts can be added after we confirm your actual statement formats.")
        if st.button("Reconcile transactions", type="primary", use_container_width=True):
            try:
                bank_tx = normalize_transactions(bank_df, bank_map, "BANK")
                cash_tx = normalize_transactions(cash_df, cash_map, "CASHBOOK")
                results = reconcile(
                    bank_tx, cash_tx, tolerance=tolerance,
                    date_window_days=int(date_window) if use_date else None,
                    use_description=use_description
                )
                st.session_state["reconciliation"] = results
                st.success("Reconciliation completed. Review each result below.")
            except Exception as exc:
                st.error(f"Could not reconcile these files: {exc}")

        results = st.session_state.get("reconciliation")
        if results is not None:
            st.subheader("3. Reconciliation summary")
            summary = results["summary"]
            cols = st.columns(5)
            for col, label, value in zip(cols, ["Matched pairs", "Bank only", "Cashbook only", "Potential duplicates", "Amount differences"],
                                         [summary["matched"], summary["bank_only"], summary["cashbook_only"], summary["potential_duplicates"], summary["amount_differences"]]):
                col.metric(label, value)

            tabs = st.tabs(["Matched", "Bank only", "Cashbook only", "Needs investigation", "All results"])
            for tab, key in zip(tabs, ["matched", "bank_only", "cashbook_only", "investigate", "all"]):
                with tab:
                    df = results[key]
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    if not df.empty:
                        st.download_button(
                            f"Download {key.replace('_', ' ').title()} (CSV)",
                            df.to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"{key}_transactions.csv",
                            mime="text/csv",
                            key=f"download_{key}"
                        )
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for key in ["summary", "matched", "bank_only", "cashbook_only", "investigate", "all"]:
                    obj = results[key]
                    if isinstance(obj, pd.DataFrame):
                        obj.to_excel(writer, sheet_name=key[:31], index=False)
                    else:
                        pd.DataFrame([obj]).to_excel(writer, sheet_name="summary", index=False)
            st.download_button("⬇️ Download complete reconciliation workbook", output.getvalue(),
                               file_name="bank_reconciliation_report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary")
    except Exception as exc:
        st.error(f"Unable to read one of the uploaded files: {exc}")
else:
    st.markdown("Upload both files to begin. The app will let you select the date, amount, description and debit/credit columns.")
    st.markdown("**Initial matching policy**")
    st.markdown("- One bank entry can match only one cashbook entry, and vice versa.\n- Amount difference must be strictly less than the configured tolerance (default KSh 1).\n- Duplicate leftovers are flagged for investigation, not silently ignored.\n- Possible matches are ranked using date proximity and, optionally, description/reference similarity.")
