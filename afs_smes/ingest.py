"""Read and clean trial balance and general ledger exports.

Every change made to the source data is written to a cleaning log so the
clean-up is fully traceable (file, source row, issue, action taken).
"""
import re
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path

import pandas as pd

GL_ALIASES = {
    "date": ["date", "posting date", "transaction date", "doc date", "document date", "trans date", "gl date"],
    "journal": ["journal", "journal no", "journal number", "jv no", "jv", "voucher", "voucher no", "entry no",
                "document no", "doc no", "transaction id", "journal id", "je no"],
    "account": ["account", "account code", "account no", "account number", "acc code", "acct code", "gl code",
                "gl account", "code", "acc no"],
    "account_name": ["account name", "account description", "account title", "acc name", "gl name", "name"],
    "description": ["description", "narration", "memo", "particulars", "details", "line description"],
    "reference": ["reference", "ref", "ref no", "source", "source document"],
    "debit": ["debit", "dr", "debit amount", "debits"],
    "credit": ["credit", "cr", "credit amount", "credits"],
    "amount": ["amount", "net amount", "net"],
}
TB_ALIASES = {
    "account": GL_ALIASES["account"],
    "account_name": GL_ALIASES["account_name"] + ["description"],
    "opening": ["opening", "opening balance", "beginning balance", "balance b/f", "opening bal"],
    "opening_dr": ["opening dr", "opening debit", "opening balance dr", "beginning debit", "opening balance debit"],
    "opening_cr": ["opening cr", "opening credit", "opening balance cr", "beginning credit",
                   "opening balance credit"],
    "debit": ["debit", "dr", "movement dr", "period debit", "debits", "debit movement", "movement debit"],
    "credit": ["credit", "cr", "movement cr", "period credit", "credits", "credit movement", "movement credit"],
    "closing": ["closing", "closing balance", "ending balance", "balance", "balance c/f", "closing bal"],
    "closing_dr": ["closing dr", "closing debit", "closing balance dr", "ending debit", "closing balance debit"],
    "closing_cr": ["closing cr", "closing credit", "closing balance cr", "ending credit",
                   "closing balance credit"],
}
TOTAL_WORDS = re.compile(r"^\s*(grand\s+)?(sub\s*-?\s*)?totals?\b", re.I)


@dataclass
class Log:
    rows: list = field(default_factory=list)

    def add(self, file, row, issue, action):
        self.rows.append({"file": file, "row": row, "issue": issue, "action": action})


def _norm(s):
    return re.sub(r"[^a-z0-9/ ]", "", str(s).lower().replace("_", " ").replace(".", " ")).strip().replace("  ", " ")


def _read_raw(path, sheet=None):
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, header=None, dtype=object, keep_default_na=False, na_values=[""])
    return pd.read_excel(path, sheet_name=sheet or 0, header=None, dtype=object)


def _locate_header(raw, aliases, required):
    """Find the header row (within the first 30 rows) and map columns to canonical names."""
    lookup = {a: canon for canon, names in aliases.items() for a in names}
    best = None
    for i in range(min(30, len(raw))):
        cols = {}
        for j, v in enumerate(raw.iloc[i].tolist()):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            canon = lookup.get(_norm(v))
            if canon and canon not in cols:
                cols[canon] = j
        if best is None or len(cols) > len(best[1]):
            best = (i, cols)
    if best is None or not all(any(r in best[1] for r in alt) for alt in required):
        found = sorted(best[1]) if best else []
        raise ValueError(f"Could not find the header row. Columns recognised: {found}. "
                         f"Required (one of each group): {required}")
    return best


def parse_number(v):
    """Return (float, was_text). Handles '1,234.50', '(1,234.50)', '1 234,50-', blanks."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0, False
    if isinstance(v, (int, float)):
        return float(v), False
    s = str(v).strip().replace(" ", "").replace(" ", "")
    if s in ("", "-", "--"):
        return 0.0, False
    neg = s.startswith("(") and s.endswith(")") or s.endswith("-") or s.startswith("-")
    s = re.sub(r"[^0-9.,]", "", s)
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.search(r",\d{3}(,|$)", s) else s.replace(",", ".")
    if s == "":
        raise ValueError(f"not a number: {v!r}")
    return (-float(s) if neg else float(s)), True


def parse_code(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def parse_date(v, dayfirst=True):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None, False
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date() if hasattr(v, "date") else v, False
    if isinstance(v, date):
        return v, False
    try:
        return pd.to_datetime(str(v).strip(), dayfirst=dayfirst).date(), True
    except Exception:
        return None, True


def _text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _is_blank(row):
    return all(v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "" for v in row)


def read_tb(path, log, sheet=None):
    """Return DataFrame[account, account_name, opening, debit, credit, closing] (debit +, credit -)."""
    fname = Path(path).name
    raw = _read_raw(path, sheet)
    hdr, cols = _locate_header(raw, TB_ALIASES, [("account",), ("closing", "closing_dr")])
    if hdr > 0:
        log.add(fname, f"1-{hdr}", "Title/blank rows above the header", f"Skipped {hdr} row(s)")
    out, n_text, n_blank = [], 0, 0
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        xl = i + 1
        if _is_blank(row):
            n_blank += 1
            continue
        get = lambda k: row[cols[k]] if k in cols else None
        code_raw = get("account")
        code, name = parse_code(code_raw), _text(get("account_name"))
        if TOTAL_WORDS.match(code) or TOTAL_WORDS.match(name) or (code == "" and name == ""):
            log.add(fname, xl, f"Total/summary row ('{code or name or 'no account'}')", "Excluded from data")
            continue
        if code == "":
            log.add(fname, xl, f"Row without an account code ('{name}')", "Excluded - investigate")
            continue
        if isinstance(code_raw, str) and code_raw != code:
            log.add(fname, xl, f"Account code {code_raw!r} has stray spaces", f"Trimmed to '{code}'")
        vals = {}
        try:
            for k in ("opening", "opening_dr", "opening_cr", "debit", "credit", "closing", "closing_dr", "closing_cr"):
                vals[k], was_text = parse_number(get(k))
                n_text += was_text
        except ValueError as e:
            log.add(fname, xl, f"Unreadable amount ({e})", "Row excluded - investigate")
            continue
        opening = vals["opening"] + vals["opening_dr"] - vals["opening_cr"]
        closing = vals["closing"] + vals["closing_dr"] - vals["closing_cr"]
        debit, credit = vals["debit"], vals["credit"]
        if "debit" not in cols and "credit" not in cols:
            mv = closing - opening
            debit, credit = (mv, 0.0) if mv >= 0 else (0.0, -mv)
        out.append({"account": code, "account_name": name, "opening": round(opening, 2),
                    "debit": round(debit, 2), "credit": round(credit, 2), "closing": round(closing, 2), "row": xl})
    if n_blank:
        log.add(fname, "-", f"{n_blank} blank row(s)", "Removed")
    if n_text:
        log.add(fname, "-", f"{n_text} amount(s) stored as text (thousand separators, brackets)",
                "Converted to numbers")
    if "debit" not in cols and "credit" not in cols:
        log.add(fname, "-", "No movement columns in the trial balance", "Movement derived as closing less opening")
    if "opening" not in cols and "opening_dr" not in cols:
        log.add(fname, "-", "No opening balance column in the trial balance", "Opening balances taken as zero")
    return pd.DataFrame(out, columns=["account", "account_name", "opening", "debit", "credit", "closing", "row"])


def read_gl(path, log, sheet=None, dayfirst=True):
    """Return DataFrame[date, journal, account, account_name, description, reference, debit, credit]."""
    fname = Path(path).name
    raw = _read_raw(path, sheet)
    hdr, cols = _locate_header(raw, GL_ALIASES, [("account",), ("debit", "amount")])
    if hdr > 0:
        log.add(fname, f"1-{hdr}", "Title/blank rows above the header", f"Skipped {hdr} row(s)")
    out, n_text, n_blank, n_dates, n_neg = [], 0, 0, 0, 0
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        xl = i + 1
        if _is_blank(row):
            n_blank += 1
            continue
        get = lambda k: row[cols[k]] if k in cols else None
        code = parse_code(get("account"))
        label = " ".join(_text(get(k)) for k in ("date", "journal", "account", "account_name", "description"))
        if code == "" and TOTAL_WORDS.search(label.strip()) or TOTAL_WORDS.match(code):
            log.add(fname, xl, f"Total/summary row ('{label.strip()[:40]}')", "Excluded from data")
            continue
        try:
            dr, t1 = parse_number(get("debit"))
            cr, t2 = parse_number(get("credit"))
            amt, t3 = parse_number(get("amount"))
        except ValueError as e:
            log.add(fname, xl, f"Unreadable amount ({e})", "Row excluded - investigate")
            continue
        n_text += t1 + t2 + t3
        if "debit" not in cols:
            dr, cr = (amt, 0.0) if amt >= 0 else (0.0, -amt)
        if dr < 0 or cr < 0:
            n_neg += 1
            log.add(fname, xl, f"Negative amount in debit/credit column (Dr {dr:,.2f} / Cr {cr:,.2f})",
                    "Moved to the opposite column as a positive amount")
            net = dr - cr
            dr, cr = (net, 0.0) if net >= 0 else (0.0, -net)
        d, was_text = parse_date(get("date"), dayfirst)
        n_dates += bool(was_text and d)
        out.append({"date": d, "journal": parse_code(get("journal")), "account": code,
                    "account_name": _text(get("account_name")), "description": _text(get("description")),
                    "reference": _text(get("reference")), "debit": round(dr, 2), "credit": round(cr, 2), "row": xl})
    if n_blank:
        log.add(fname, "-", f"{n_blank} blank row(s)", "Removed")
    if n_text:
        log.add(fname, "-", f"{n_text} amount(s) stored as text", "Converted to numbers")
    if n_dates:
        log.add(fname, "-", f"{n_dates} date(s) stored as text", "Converted to dates "
                + ("(day-first)" if dayfirst else "(month-first)"))
    df = pd.DataFrame(out, columns=["date", "journal", "account", "account_name", "description", "reference",
                                    "debit", "credit", "row"])
    return _drop_duplicate_lines(df, fname, log)


def _drop_duplicate_lines(df, fname, log):
    """Remove an exact duplicate line only where that restores the balance of its journal."""
    if df.empty:
        return df
    key = ["date", "journal", "account", "description", "reference", "debit", "credit"]
    dup = df.duplicated(subset=key, keep="first")
    drop = []
    for idx in df.index[dup]:
        j = df.loc[idx, "journal"]
        lines = df[(df["journal"] == j) & ~df.index.isin(drop)]
        imbalance = round(lines["debit"].sum() - lines["credit"].sum(), 2)
        this = round(df.loc[idx, "debit"] - df.loc[idx, "credit"], 2)
        if j != "" and imbalance != 0 and imbalance == this:
            drop.append(idx)
            log.add(fname, int(df.loc[idx, "row"]),
                    f"Exact duplicate line in journal {j} (account {df.loc[idx, 'account']}, {abs(this):,.2f}) "
                    "which put the journal out of balance", "Duplicate removed - journal now balances")
    return df.drop(index=drop).reset_index(drop=True)


def harmonise_names(tb_frames, gl_frames, log):
    """One account name per code: the current-year TB name wins; GL spelling variants are reported."""
    names = {}
    for tb in tb_frames:
        for code, name in zip(tb["account"], tb["account_name"]):
            if name and code not in names:
                names[code] = name
    for gl in gl_frames:
        if gl is None or gl.empty:
            continue
        variants = gl[gl["account_name"] != ""].groupby("account")["account_name"].unique()
        changed = []
        for code, vs in variants.items():
            official = names.setdefault(code, vs[0])
            if any(v != official for v in vs):
                changed.append(code)
        if changed:
            shown = ", ".join(changed[:8]) + (f" (+{len(changed) - 8} more)" if len(changed) > 8 else "")
            log.add("General ledger", "-", f"{len(changed)} account(s) appear under more than one name or spelling: {shown}",
                    "Standardised to the trial balance name")
        gl["account_name"] = gl["account"].map(names).fillna(gl["account_name"])
    return names
