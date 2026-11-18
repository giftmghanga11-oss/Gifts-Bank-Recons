# Automated Bank Reconciliation

A Streamlit application for reconciling bank statements against ERP cashbooks.

## Current first-version features
- Upload bank statement and ERP cashbook in Excel (`.xlsx`, `.xls`) or CSV.
- Map each file's date, amount, description/reference and optional debit/credit columns.
- One-to-one matching: each transaction can be matched only once.
- Default KSh 1 tolerance is strict: absolute difference must be **less than KSh 1**. A difference of exactly KSh 1 does not match.
- Date-window control and optional description/reference ranking.
- Separate matched, bank-only, cashbook-only and investigation tabs.
- Download CSV tabs or a complete Excel reconciliation workbook.

## Run locally
1. Install Python 3.10 or later.
2. Open a terminal in this folder.
3. Create and activate a virtual environment (optional but recommended).
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
5. Start the app:
   ```bash
   streamlit run app.py
   ```

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository and upload these files.
2. Sign in to Streamlit Community Cloud using GitHub.
3. Select **New app**, choose the repository and `app.py`.
4. Deploy.

## Important accounting control
This tool proposes matches; it does not post journals or alter either source file. Review all unmatched and potential-duplicate items before finalising a bank reconciliation.

## Next development stage
Test against anonymised real examples from the bank and ERP exports. Then add PDF statement extraction, saved column mappings for each bank/ERP export, debit-credit sign configuration, manual match override, and a formal reconciliation statement with opening/closing balances.
