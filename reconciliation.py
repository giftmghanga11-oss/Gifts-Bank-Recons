import pandas as pd
from difflib import SequenceMatcher

def _clean_amount(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.replace(",", "").replace("KSh", "").replace("KES", "").strip()
        if value.startswith("(") and value.endswith(")"):
            value = "-" + value[1:-1]
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None

def _direction_multiplier(value):
    if pd.isna(value):
        return 1
    text = str(value).strip().lower()
    if any(x in text for x in ["debit", "withdrawal", "payment", "dr"]):
        return -1
    if any(x in text for x in ["credit", "deposit", "receipt", "cr"]):
        return 1
    return 1

def normalize_transactions(df, mapping, source):
    if mapping["amount"] not in df.columns:
        raise ValueError(f"{source}: selected amount column does not exist.")
    out = pd.DataFrame(index=df.index)
    out["source_row"] = df.index + 2
    out["date"] = pd.to_datetime(df[mapping["date"]], errors="coerce") if mapping.get("date") else pd.NaT
    out["description"] = df[mapping["description"]].fillna("").astype(str) if mapping.get("description") else ""
    raw_amounts = df[mapping["amount"]].apply(_clean_amount)
    if mapping.get("direction"):
        signs = df[mapping["direction"]].apply(_direction_multiplier)
        raw_amounts = raw_amounts * signs
    out["amount"] = raw_amounts
    out["source"] = source
    out["original_index"] = df.index
    out = out.dropna(subset=["amount"]).reset_index(drop=True)
    return out

def _similarity(a, b):
    a = str(a or "").lower().strip()
    b = str(b or "").lower().strip()
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

def reconcile(bank, cashbook, tolerance=1.0, date_window_days=5, use_description=True):
    if tolerance <= 0:
        raise ValueError("Tolerance must be greater than zero.")
    bank = bank.copy().reset_index(drop=True)
    cashbook = cashbook.copy().reset_index(drop=True)
    bank["status"] = "Bank only"
    cashbook["status"] = "Cashbook only"
    bank["matched_cashbook_row"] = None
    cashbook["matched_bank_row"] = None
    bank["amount_difference"] = None
    cashbook["amount_difference"] = None

    candidates = []
    for bi, b in bank.iterrows():
        for ci, c in cashbook.iterrows():
            diff = abs(float(b["amount"]) - float(c["amount"]))
            # Strictly less than tolerance: at tolerance=1.00, 1.00 is NOT a match.
            if diff >= tolerance:
                continue
            if date_window_days is not None and pd.notna(b["date"]) and pd.notna(c["date"]):
                days = abs((b["date"] - c["date"]).days)
                if days > date_window_days:
                    continue
            else:
                days = 0
            desc_score = _similarity(b["description"], c["description"]) if use_description else 0.0
            candidates.append((diff, days, -desc_score, bi, ci))

    # Greedy one-to-one assignment: each transaction can be consumed once only.
    candidates.sort()
    used_bank, used_cash = set(), set()
    for diff, days, neg_score, bi, ci in candidates:
        if bi in used_bank or ci in used_cash:
            continue
        used_bank.add(bi)
        used_cash.add(ci)
        bank.at[bi, "status"] = "Matched"
        cashbook.at[ci, "status"] = "Matched"
        bank.at[bi, "matched_cashbook_row"] = int(cashbook.at[ci, "source_row"])
        cashbook.at[ci, "matched_bank_row"] = int(bank.at[bi, "source_row"])
        bank.at[bi, "amount_difference"] = round(float(bank.at[bi, "amount"]) - float(cashbook.at[ci, "amount"]), 2)
        cashbook.at[ci, "amount_difference"] = round(float(bank.at[bi, "amount"]) - float(cashbook.at[ci, "amount"]), 2)

    # Unmatched rows that had at least one eligible candidate are possible duplicates/conflicts.
    eligible_bank = {x[3] for x in candidates}
    eligible_cash = {x[4] for x in candidates}
    for bi in range(len(bank)):
        if bi not in used_bank and bi in eligible_bank:
            bank.at[bi, "status"] = "Potential duplicate / needs investigation"
    for ci in range(len(cashbook)):
        if ci not in used_cash and ci in eligible_cash:
            cashbook.at[ci, "status"] = "Potential duplicate / needs investigation"

    bank_out = bank.copy()
    cash_out = cashbook.copy()
    matched_bank = bank_out[bank_out["status"] == "Matched"].copy()
    matched_cash = cash_out[cash_out["status"] == "Matched"].copy()
    matched = matched_bank.merge(
        matched_cash[["source_row", "description", "amount", "date"]],
        left_on="matched_cashbook_row", right_on="source_row",
        suffixes=("_bank", "_cashbook"), how="left"
    )
    matched = matched.rename(columns={
        "source_row_bank": "bank_source_row",
        "source_row_cashbook": "cashbook_source_row",
        "description_bank": "bank_description",
        "description_cashbook": "cashbook_description",
        "amount_bank": "bank_amount",
        "amount_cashbook": "cashbook_amount",
        "date_bank": "bank_date",
        "date_cashbook": "cashbook_date",
    })
    bank_only = bank_out[bank_out["status"] == "Bank only"].copy()
    cash_only = cash_out[cash_out["status"] == "Cashbook only"].copy()
    investigate = pd.concat([
        bank_out[bank_out["status"].str.contains("investigation", case=False)],
        cash_out[cash_out["status"].str.contains("investigation", case=False)]
    ], ignore_index=True)
    all_results = pd.concat([bank_out, cash_out], ignore_index=True)
    summary = {
        "matched": int(len(matched)),
        "bank_only": int(len(bank_only)),
        "cashbook_only": int(len(cash_only)),
        "potential_duplicates": int(len(investigate)),
        "amount_differences": int((matched["amount_difference"].fillna(0) != 0).sum()) if not matched.empty else 0,
        "bank_transactions": int(len(bank)),
        "cashbook_transactions": int(len(cashbook)),
    }
    return {
        "summary": summary,
        "matched": matched,
        "bank_only": bank_only,
        "cashbook_only": cash_only,
        "investigate": investigate,
        "all": all_results,
    }
