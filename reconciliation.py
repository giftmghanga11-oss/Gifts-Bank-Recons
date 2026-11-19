from io import BytesIO
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def _clean_amount(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float, np.number)):
        val = float(value)
        return None if np.isnan(val) else val
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("KSh", "").replace("KES", "").strip()
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1].strip()
        try:
            val = float(cleaned)
            return None if np.isnan(val) else val
        except ValueError:
            return None
    return None


def _direction_multiplier(value):
    if pd.isna(value):
        return 1.0
    text = str(value).strip().lower()
    if any(x in text for x in ["debit", "withdrawal", "payment", "dr"]):
        return -1.0
    if any(x in text for x in ["credit", "deposit", "receipt", "cr"]):
        return 1.0
    return 1.0


def normalize_transactions(df: pd.DataFrame, mapping: dict, source: str) -> pd.DataFrame:
    """
    Normalizes Bank or ERP rows into a clean DataFrame with:
    source_row, date, reference, description, amount (signed), abs_amount, flow ("DEBIT" or "CREDIT"), source
    """
    rows = []
    date_col = mapping.get("date")
    ref_col = mapping.get("reference")
    desc_col = mapping.get("description")
    debit_col = mapping.get("debit")
    credit_col = mapping.get("credit")
    amount_col = mapping.get("amount")
    direction_col = mapping.get("direction")

    for idx, row in df.iterrows():
        src_row = int(idx) + 2
        dt = pd.to_datetime(row[date_col], errors="coerce") if (date_col and date_col in df.columns) else pd.NaT
        ref = "" if not (ref_col and ref_col in df.columns) or pd.isna(row[ref_col]) else str(row[ref_col]).strip()
        desc = "" if not (desc_col and desc_col in df.columns) or pd.isna(row[desc_col]) else str(row[desc_col]).strip()

        if debit_col or credit_col:
            dr_val = abs(_clean_amount(row[debit_col]) or 0.0) if (debit_col and debit_col in df.columns) else 0.0
            cr_val = abs(_clean_amount(row[credit_col]) or 0.0) if (credit_col and credit_col in df.columns) else 0.0

            if dr_val > 0.0 and cr_val == 0.0:
                rows.append({
                    "source_row": src_row,
                    "date": dt,
                    "reference": ref,
                    "description": desc,
                    "amount": -dr_val,
                    "abs_amount": dr_val,
                    "flow": "DEBIT",
                    "source": source,
                })
            elif cr_val > 0.0 and dr_val == 0.0:
                rows.append({
                    "source_row": src_row,
                    "date": dt,
                    "reference": ref,
                    "description": desc,
                    "amount": cr_val,
                    "abs_amount": cr_val,
                    "flow": "CREDIT",
                    "source": source,
                })
            elif dr_val > 0.0 and cr_val > 0.0:
                net = cr_val - dr_val
                if net != 0.0:
                    rows.append({
                        "source_row": src_row,
                        "date": dt,
                        "reference": ref,
                        "description": desc,
                        "amount": net,
                        "abs_amount": abs(net),
                        "flow": "CREDIT" if net > 0 else "DEBIT",
                        "source": source,
                    })
        elif amount_col and amount_col in df.columns:
            raw = _clean_amount(row[amount_col])
            if raw is None or raw == 0.0:
                continue
            if direction_col and direction_col in df.columns:
                mult = _direction_multiplier(row[direction_col])
                signed_amt = abs(raw) * mult
            else:
                signed_amt = raw
            rows.append({
                "source_row": src_row,
                "date": dt,
                "reference": ref,
                "description": desc,
                "amount": signed_amt,
                "abs_amount": abs(signed_amt),
                "flow": "CREDIT" if signed_amt > 0 else "DEBIT",
                "source": source,
            })
        else:
            raise ValueError(f"{source}: valid amount or debit/credit column mapping is required.")

    cols = ["source_row", "date", "reference", "description", "amount", "abs_amount", "flow", "source"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows, columns=cols)


def _align_opposite_side(
    bank_side_df: pd.DataFrame,
    erp_side_df: pd.DataFrame,
    pair_label: str,
    tolerance: float,
) -> tuple[pd.DataFrame, set, set]:
    """
    Sorts both sides ascending by abs_amount, performs strict one-to-one matching
    where abs(bank_amt - erp_amt) < tolerance (ignoring dates & descriptions),
    and aligns rows side-by-side with Variance = Bank Amount - ERP Amount.
    Returns (aligned_df, matched_bank_rows, matched_erp_rows).
    """
    b_sorted = bank_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)
    e_sorted = erp_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)

    n_b = len(b_sorted)
    n_e = len(e_sorted)
    b_to_e: dict[int, int] = {}
    e_used: set[int] = set()
    matched_b_src_rows: set[int] = set()
    matched_e_src_rows: set[int] = set()

    e_start = 0
    for bi in range(n_b):
        b_amt = float(b_sorted.at[bi, "abs_amount"])
        while e_start < n_e and (b_amt - float(e_sorted.at[e_start, "abs_amount"])) >= tolerance:
            e_start += 1

        best_ei = None
        best_diff = tolerance
        for ei in range(e_start, n_e):
            e_amt = float(e_sorted.at[ei, "abs_amount"])
            if (e_amt - b_amt) >= tolerance:
                break
            if ei in e_used:
                continue
            diff = abs(b_amt - e_amt)
            if diff < best_diff:
                best_diff = diff
                best_ei = ei
                if best_diff == 0.0:
                    break

        if best_ei is not None:
            b_to_e[bi] = best_ei
            e_used.add(best_ei)
            matched_b_src_rows.add(int(b_sorted.at[bi, "source_row"]))
            matched_e_src_rows.add(int(e_sorted.at[best_ei, "source_row"]))

    aligned_rows = []
    b_ptr = 0
    e_ptr = 0

    while b_ptr < n_b or e_ptr < n_e:
        while e_ptr < n_e and e_ptr in e_used:
            e_ptr += 1

        has_b = b_ptr < n_b
        has_e = e_ptr < n_e

        if has_b and b_ptr in b_to_e:
            ei = b_to_e[b_ptr]
            b_amt = float(b_sorted.at[b_ptr, "abs_amount"])
            e_amt = float(e_sorted.at[ei, "abs_amount"])
            variance = b_amt - e_amt
            aligned_rows.append({
                "Pair": pair_label,
                "Status": "Matched",
                "Bank Row": int(b_sorted.at[b_ptr, "source_row"]),
                "Bank Ref": b_sorted.at[b_ptr, "reference"],
                "Bank Date": b_sorted.at[b_ptr, "date"],
                "Bank Description": b_sorted.at[b_ptr, "description"],
                "Bank Amount": b_amt,
                "Variance": variance,
                "ERP Amount": e_amt,
                "ERP Description": e_sorted.at[ei, "description"],
                "ERP Date": e_sorted.at[ei, "date"],
                "ERP Ref": e_sorted.at[ei, "reference"],
                "ERP Row": int(e_sorted.at[ei, "source_row"]),
            })
            b_ptr += 1
        elif has_b and (not has_e or float(b_sorted.at[b_ptr, "abs_amount"]) <= float(e_sorted.at[e_ptr, "abs_amount"])):
            b_amt = float(b_sorted.at[b_ptr, "abs_amount"])
            aligned_rows.append({
                "Pair": pair_label,
                "Status": "Bank Only",
                "Bank Row": int(b_sorted.at[b_ptr, "source_row"]),
                "Bank Ref": b_sorted.at[b_ptr, "reference"],
                "Bank Date": b_sorted.at[b_ptr, "date"],
                "Bank Description": b_sorted.at[b_ptr, "description"],
                "Bank Amount": b_amt,
                "Variance": b_amt,
                "ERP Amount": np.nan,
                "ERP Description": "",
                "ERP Date": pd.NaT,
                "ERP Ref": "",
                "ERP Row": np.nan,
            })
            b_ptr += 1
        elif has_e:
            e_amt = float(e_sorted.at[e_ptr, "abs_amount"])
            aligned_rows.append({
                "Pair": pair_label,
                "Status": "ERP Only",
                "Bank Row": np.nan,
                "Bank Ref": "",
                "Bank Date": pd.NaT,
                "Bank Description": "",
                "Bank Amount": np.nan,
                "Variance": -e_amt,
                "ERP Amount": e_amt,
                "ERP Description": e_sorted.at[e_ptr, "description"],
                "ERP Date": e_sorted.at[e_ptr, "date"],
                "ERP Ref": e_sorted.at[e_ptr, "reference"],
                "ERP Row": int(e_sorted.at[e_ptr, "source_row"]),
            })
            e_ptr += 1

    cols = [
        "Pair", "Status", "Bank Row", "Bank Ref", "Bank Date", "Bank Description",
        "Bank Amount", "Variance", "ERP Amount", "ERP Description", "ERP Date", "ERP Ref", "ERP Row"
    ]
    df_out = pd.DataFrame(aligned_rows, columns=cols) if aligned_rows else pd.DataFrame(columns=cols)
    return df_out, matched_b_src_rows, matched_e_src_rows


def _build_investigation_df(side_df: pd.DataFrame, matched_rows: set, source_label: str) -> pd.DataFrame:
    if side_df.empty:
        return pd.DataFrame(columns=["Source", "Flow", "Matched", "Reason", "Amount", "Date", "Reference", "Description", "Source Row"])

    counts = side_df["abs_amount"].value_counts().to_dict()
    flagged = []
    for _, r in side_df.iterrows():
        srow = int(r["source_row"])
        amt = float(r["abs_amount"])
        is_matched = srow in matched_rows
        dup_cnt = int(counts.get(amt, 1))
        is_dup = dup_cnt > 1

        if not is_matched or is_dup:
            reasons = []
            if not is_matched:
                reasons.append("Unmatched")
            if is_dup:
                reasons.append(f"Duplicate amount ({dup_cnt}x)")
            flagged.append({
                "Source": source_label,
                "Flow": r["flow"],
                "Matched": "Yes" if is_matched else "No",
                "Reason": " & ".join(reasons),
                "Amount": amt,
                "Date": r["date"],
                "Reference": r["reference"],
                "Description": r["description"],
                "Source Row": srow,
            })

    cols = ["Source", "Flow", "Matched", "Reason", "Amount", "Date", "Reference", "Description", "Source Row"]
    return pd.DataFrame(flagged, columns=cols) if flagged else pd.DataFrame(columns=cols)


def reconcile(bank: pd.DataFrame, cashbook: pd.DataFrame, tolerance: float = 1.0) -> dict:
    if tolerance <= 0:
        raise ValueError("Tolerance must be greater than zero.")

    bank_dr = bank[bank["flow"] == "DEBIT"]
    bank_cr = bank[bank["flow"] == "CREDIT"]
    erp_dr = cashbook[cashbook["flow"] == "DEBIT"]
    erp_cr = cashbook[cashbook["flow"] == "CREDIT"]

    # Bank Debits <-> ERP Credits; Bank Credits <-> ERP Debits
    dr_vs_cr, m_b_dr, m_e_cr = _align_opposite_side(bank_dr, erp_cr, "Bank Debits vs ERP Credits", tolerance)
    cr_vs_dr, m_b_cr, m_e_dr = _align_opposite_side(bank_cr, erp_dr, "Bank Credits vs ERP Debits", tolerance)

    all_df = pd.concat([dr_vs_cr, cr_vs_dr], ignore_index=True)
    matched_df = all_df[all_df["Status"] == "Matched"].copy()
    bank_only_df = all_df[all_df["Status"] == "Bank Only"].copy()
    erp_only_df = all_df[all_df["Status"] == "ERP Only"].copy()

    matched_bank_all = m_b_dr | m_b_cr
    matched_erp_all = m_e_dr | m_e_cr

    inv_parts = [
        _build_investigation_df(bank_dr, matched_bank_all, "BANK"),
        _build_investigation_df(bank_cr, matched_bank_all, "BANK"),
        _build_investigation_df(erp_dr, matched_erp_all, "ERP"),
        _build_investigation_df(erp_cr, matched_erp_all, "ERP"),
    ]
    non_empty_inv = [p for p in inv_parts if not p.empty]
    investigate_df = pd.concat(non_empty_inv, ignore_index=True) if non_empty_inv else pd.DataFrame(
        columns=["Source", "Flow", "Matched", "Reason", "Amount", "Date", "Reference", "Description", "Source Row"]
    )

    total_bank_amt = float(bank["abs_amount"].sum()) if not bank.empty else 0.0
    total_erp_amt = float(cashbook["abs_amount"].sum()) if not cashbook.empty else 0.0
    matched_bank_amt = float(matched_df["Bank Amount"].sum()) if not matched_df.empty else 0.0
    matched_erp_amt = float(matched_df["ERP Amount"].sum()) if not matched_df.empty else 0.0
    outstanding_bank_amt = float(bank_only_df["Bank Amount"].sum()) if not bank_only_df.empty else 0.0
    outstanding_erp_amt = float(erp_only_df["ERP Amount"].sum()) if not erp_only_df.empty else 0.0
    net_variance = float(all_df["Variance"].sum()) if not all_df.empty else 0.0

    summary_table = pd.DataFrame([
        {"Metric": "Total Bank Amount", "Value": total_bank_amt},
        {"Metric": "Total ERP Amount", "Value": total_erp_amt},
        {"Metric": "Matched Bank Amount", "Value": matched_bank_amt},
        {"Metric": "Matched ERP Amount", "Value": matched_erp_amt},
        {"Metric": "Outstanding Bank Amount (Green)", "Value": outstanding_bank_amt},
        {"Metric": "Outstanding ERP Amount (Red)", "Value": outstanding_erp_amt},
        {"Metric": "Net Variance (Bank - ERP)", "Value": net_variance},
        {"Metric": "Matched Pairs Count", "Value": int(len(matched_df))},
        {"Metric": "Bank-Only Count", "Value": int(len(bank_only_df))},
        {"Metric": "ERP-Only Count", "Value": int(len(erp_only_df))},
        {"Metric": "Items Flagged for Investigation", "Value": int(len(investigate_df))},
    ])

    summary = {
        "matched": int(len(matched_df)),
        "bank_only": int(len(bank_only_df)),
        "erp_only": int(len(erp_only_df)),
        "potential_duplicates": int(len(investigate_df)),
        "net_variance": net_variance,
        "total_bank_amount": total_bank_amt,
        "total_erp_amount": total_erp_amt,
        "outstanding_bank_amount": outstanding_bank_amt,
        "outstanding_erp_amount": outstanding_erp_amt,
    }

    return {
        "summary": summary,
        "summary_table": summary_table,
        "debits_vs_credits": dr_vs_cr,
        "credits_vs_debits": cr_vs_dr,
        "matched": matched_df,
        "bank_only": bank_only_df,
        "erp_only": erp_only_df,
        "investigate": investigate_df,
        "all": all_df,
    }


def build_reconciliation_excel(results: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    sheets = [
        ("Bank Debits vs ERP Credits", results.get("debits_vs_credits"), True),
        ("Bank Credits vs ERP Debits", results.get("credits_vs_debits"), True),
        ("Matched", results.get("matched"), False),
        ("Bank Only", results.get("bank_only"), True),
        ("ERP Only", results.get("erp_only"), True),
        ("Needs Investigation", results.get("investigate"), False),
        ("Summary Totals", results.get("summary_table"), False),
    ]

    for title, df, color_rows in sheets:
        ws = wb.create_sheet(title=title[:31])
        if df is None or df.empty:
            ws.append(["No records"])
            continue

        cols = list(df.columns)
        ws.append(cols)
        for c_idx in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font

        status_idx = cols.index("Status") if "Status" in cols else None

        for _, r in df.iterrows():
            row_vals = []
            for c in cols:
                val = r[c]
                if pd.isna(val):
                    row_vals.append(None)
                elif isinstance(val, pd.Timestamp):
                    row_vals.append(val.to_pydatetime())
                else:
                    row_vals.append(val)
            ws.append(row_vals)
            cur_row = ws.max_row

            if color_rows and status_idx is not None:
                st = str(r.iloc[status_idx])
                fill = green_fill if st == "Bank Only" else (red_fill if st == "ERP Only" else None)
                if fill:
                    for c_idx in range(1, len(cols) + 1):
                        ws.cell(row=cur_row, column=c_idx).fill = fill

        num_cols = {"Bank Amount", "Variance", "ERP Amount", "Amount", "Value"}
        for c_idx, col_name in enumerate(cols, start=1):
            if col_name in num_cols:
                for r_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

            max_len = len(str(col_name))
            for r_idx in range(2, min(ws.max_row + 1, 250)):
                max_len = max(max_len, len(str(ws.cell(row=r_idx, column=c_idx).value or "")))
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 42)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
