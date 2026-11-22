import pandas as pd
import streamlit as st

from reconciliation import (
    NONE_CHOICE,
    build_reconciliation_excel,
    normalize_transactions,
    reconcile,
    repair_cashbook_dataframe,
)

st.set_page_config(page_title="Bank Reconciliation", page_icon="🏦", layout="wide")
st.title("🏦 Automated Bank Reconciliation")
st.caption("Bank details on the left • Variance in the middle • Cashbook details on the right")

with st.sidebar:
    st.header("Reconciliation settings")
    tolerance = st.number_input(
        "Maximum permitted difference (KSh)",
        min_value=0.01,
        value=1.0,
        step=0.1,
        format="%.2f",
        help="Entries are matched only when the difference is strictly less than this amount.",
    )

st.subheader("1. Upload your files")
left, right = st.columns(2)
with left:
    bank_file = st.file_uploader("Bank statement", type=["xlsx", "xls", "csv"], key="bank")
with right:
    cash_file = st.file_uploader("ERP cashbook", type=["xlsx", "xls", "csv"], key="cash")


def read_upload(upload):
    upload.seek(0)
    if upload.name.lower().endswith(".csv"):
        return pd.read_csv(upload)
    return pd.read_excel(upload)


def default_index(options, preferred):
    """Index of the first preferred column name that exists in options, else 0."""
    for name in preferred:
        if name in options:
            return options.index(name)
    return 0


def style_report(df: pd.DataFrame, statuses: list[str]):
    def row_style(row):
        status = statuses[row.name] if row.name < len(statuses) else ""
        if status == "Bank Only":
            return ["background-color: #d9ead3; color: #1b4332"] * len(row)
        if status == "ERP Only":
            return ["background-color: #f4cccc; color: #7f1d1d"] * len(row)
        return [""] * len(row)

    return df.style.apply(row_style, axis=1)


def pairing_controls(bank_df, cash_df):
    """Lets the user choose which bank column is matched against which ERP column."""
    bank_opts = [NONE_CHOICE] + list(bank_df.columns)
    erp_opts = [NONE_CHOICE] + list(cash_df.columns)

    st.caption(
        "Choose which bank column is matched against which ERP column. "
        "If your bank or ERP has only one amount column, select it on both lines "
        "(money in and money out)."
    )

    h1, h2 = st.columns(2)
    h1.markdown("**Bank statement column**")
    h2.markdown("**ERP cashbook column**")

    st.markdown("💰 **Money IN** (bank credits ↔ ERP receipts)")
    c1, c2 = st.columns(2)
    bank_in = c1.selectbox(
        "Bank: money-in column", bank_opts,
        index=default_index(bank_opts, ["Credit Amount", "Credit", "Deposit", "Amount"]),
        key="bank_in", label_visibility="collapsed",
    )
    erp_in = c2.selectbox(
        "ERP: money-in column", erp_opts,
        index=default_index(erp_opts, ["Payment amount", "Paid amount", "Debit Amount", "Amount"]),
        key="erp_in", label_visibility="collapsed",
    )

    st.markdown("💸 **Money OUT** (bank debits ↔ ERP payments)")
    c3, c4 = st.columns(2)
    bank_out = c3.selectbox(
        "Bank: money-out column", bank_opts,
        index=default_index(bank_opts, ["Debit Amount", "Debit", "Withdrawal"]),
        key="bank_out", label_visibility="collapsed",
    )
    erp_out = c4.selectbox(
        "ERP: money-out column", erp_opts,
        index=default_index(erp_opts, ["Payment amount", "Paid amount", "Credit Amount", "Amount"]),
        key="erp_out", label_visibility="collapsed",
    )

    direction_col, out_keywords = None, ""
    if erp_in != NONE_CHOICE and erp_in == erp_out:
        st.markdown("🧭 **The ERP uses one amount column for both directions: how do we tell money out?**")
        d1, d2 = st.columns(2)
        dir_opts = ["(use the sign: negative = money out)"] + list(cash_df.columns)
        dir_pick = d1.selectbox(
            "ERP column that tells receipts from payments", dir_opts,
            index=default_index(dir_opts, ["Payment no.", "Payment method", "Type"]),
            key="erp_direction",
        )
        direction_col = None if dir_pick == dir_opts[0] else dir_pick
        if direction_col:
            out_keywords = d2.text_input(
                "A row is money OUT when that column contains (comma-separated)",
                value="ZATRF, ZABTR",
                key="erp_out_keywords",
                help="Everything else in that column is treated as money in (receipts).",
            )

    def none_to_none(v):
        return None if v == NONE_CHOICE else v

    return {
        "bank_in": none_to_none(bank_in),
        "bank_out": none_to_none(bank_out),
        "erp_in": none_to_none(erp_in),
        "erp_out": none_to_none(erp_out),
        "erp_direction": direction_col,
        "erp_out_keywords": out_keywords,
    }


if bank_file and cash_file:
    try:
        bank_df = read_upload(bank_file)
        cash_df = repair_cashbook_dataframe(read_upload(cash_file))
    except Exception as exc:
        st.error(f"Unable to read one of the uploaded files: {exc}")
        st.stop()

    st.subheader("2. Check your data")
    p1, p2 = st.columns(2)
    with p1:
        st.markdown(f"**Bank statement preview** ({len(bank_df)} rows)")
        st.dataframe(bank_df.head(8), use_container_width=True)
    with p2:
        st.markdown(f"**ERP cashbook preview** ({len(cash_df)} rows)")
        st.dataframe(cash_df.head(8), use_container_width=True)

    st.subheader("3. Match the columns")
    mapping = pairing_controls(bank_df, cash_df)

    nothing_selected = all(mapping[k] is None for k in ("bank_in", "bank_out", "erp_in", "erp_out"))
    if nothing_selected:
        st.warning("Select at least one amount column to reconcile.")

    signature = (bank_file.name, bank_file.size, cash_file.name, cash_file.size)
    if st.button("Reconcile transactions", type="primary", use_container_width=True, disabled=nothing_selected):
        try:
            bank_tx = normalize_transactions(bank_df, mapping["bank_in"], mapping["bank_out"], "BANK")
            erp_tx = normalize_transactions(
                cash_df, mapping["erp_in"], mapping["erp_out"], "ERP",
                direction_col=mapping["erp_direction"], out_keywords=mapping["erp_out_keywords"],
            )
            st.session_state["reconciliation"] = {
                "signature": signature,
                "results": reconcile(bank_tx, erp_tx, tolerance=tolerance, mapping=mapping),
            }
        except Exception as exc:
            st.session_state.pop("reconciliation", None)
            st.error(f"Could not reconcile these files: {exc}")

    stored = st.session_state.get("reconciliation")
    if stored and stored["signature"] == signature:
        results = stored["results"]
        summary = results["summary"]

        st.subheader("4. Reconciliation report")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Matched pairs", summary["matched"])
        m2.metric("Bank only (green)", summary["bank_only"])
        m3.metric("Cashbook only (red)", summary["erp_only"])
        m4.metric("Net variance", f"{summary['net_variance']:,.2f}")

        if results["skipped"]:
            st.warning(
                f"{len(results['skipped'])} row(s) could not be classified and are NOT in the report: "
                + "; ".join(f"row {r} ({why})" for r, why in results["skipped"][:10])
            )

        st.dataframe(
            style_report(results["report_df"], results["statuses"]),
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Download reconciliation report (.xlsx)",
            build_reconciliation_excel(results),
            file_name="Bank_rec.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
else:
    st.markdown("Upload both files to generate the side-by-side bank reconciliation report.")
