"""Map ledger accounts to AFS components (taxonomy leaves) and expense natures.

Order of precedence:
  1. the user's mapping file (csv/xlsx with columns account, afs_code[, nature])
  2. keyword rules on the account name, constrained by the account class
  3. a fallback per account class, flagged "Fallback - review"

Account class comes from the leading digit(s) of the account code (configurable):
  A asset, L liability, E equity, I income, C cost of sales, X expense, T income tax
"""
import re
from pathlib import Path

import pandas as pd

from .taxonomy import LEAF, NATURES, is_pl

DEFAULT_CLASSES = {"1": "A", "2": "L", "3": "E", "4": "I", "5": "C", "6": "X", "7": "X", "8": "T", "9": "X"}

# (classes, regex on lower-case account name, leaf)  - first match wins
RULES = [
    ("A", r"accum.*(depr|amort).*(software|intangible|licen|patent)|accum.*amort", "INTANG.D"),
    ("A", r"accum.*depr.*(build|land|premis|leasehold)", "PPE.LB.D"),
    ("A", r"accum.*depr.*(plant|machin|warehouse|factory|tool)", "PPE.PM.D"),
    ("A", r"accum.*depr.*(vehicle|motor|truck|car)", "PPE.MV.D"),
    ("A", r"accum.*depr", "PPE.FE.D"),
    ("A", r"(allowance|provision).*(doubtful|impair|bad|credit loss)", "TOR.ALLOW"),
    ("A", r"investment propert", "INVPROP"),
    ("A", r"\b(land|building|premises|leasehold improvement)", "PPE.LB.C"),
    ("A", r"plant|machin|warehouse equip|factory equip|tools", "PPE.PM.C"),
    ("A", r"vehicle|motor|truck", "PPE.MV.C"),
    ("A", r"furniture|fitting|fixture|office equip|computer equip|equipment", "PPE.FE.C"),
    ("A", r"software|intangible|licen[cs]e|patent|trademark|goodwill", "INTANG.C"),
    ("A", r"associate|joint venture|jointly controlled", "INVASSOC"),
    ("A", r"deferred tax", "DTA"),
    ("A", r"raw material|consumable|packing material", "INV.RM"),
    ("A", r"work.in.progress|wip", "INV.WIP"),
    ("A", r"inventor|merchandise|finished goods|stock", "INV.FG"),
    ("A", r"interest receivable|accrued interest", "TOR.INT"),
    ("A", r"prepay|prepaid|advance to supplier|deposit paid", "TOR.PREPAY"),
    ("A", r"contract asset|unbilled|accrued income|accrued revenue", "CONTA"),
    ("A", r"(income|corporate|profit) tax|tax (recoverable|refund)|prepaid tax", "CTA"),
    ("A", r"trade receivable|accounts? receivable|trade debtor|debtors", "TOR.TRADE"),
    ("A", r"petty cash|cash on hand|cash in hand", "CASH.HAND"),
    ("A", r"(short.term|fixed|time|term|call).{0,6}deposit", "CASH.DEP"),
    ("A", r"bank|cash", "CASH.BANK"),
    ("A", r"long.term (investment|loan)|non.current.*(investment|financial)", "OFA_NC"),
    ("A", r"investment|loan (to|receivable)|financial asset", "OFA_C"),
    ("A", r"receivable|vat (input|recoverable)|staff advance", "TOR.OTHER"),
    ("L", r"overdraft", "BORR_C.OD"),
    ("L", r"(finance |)lease.*(non.current|long.term)", "BORR_NC.LEASE"),
    ("L", r"lease", "BORR_C.LEASE"),
    ("L", r"(loan|borrowing|debenture|note payable).*(non.current|long.term)|long.term (loan|borrowing)",
     "BORR_NC.LOAN"),
    ("L", r"loan|borrowing|note payable", "BORR_C.LOAN"),
    ("L", r"deferred tax", "DTL"),
    ("L", r"(income|corporate|profit) tax", "CTL"),
    ("L", r"interest payable|accrued interest", "TOP.INT"),
    ("L", r"dividend", "TOP.DIV"),
    ("L", r"customer (advance|deposit)|contract liabilit|deferred (income|revenue)|unearned", "CONTL"),
    ("L", r"(provision|warrant).*(non.current|long.term)", "PROV_NC"),
    ("L", r"provision|warrant", "PROV_C"),
    ("L", r"(seniority|severance|gratuity|pension|retirement|long.service|employee benefit)", "EBO_NC"),
    ("L", r"accru", "TOP.ACCR"),
    ("L", r"trade payable|accounts? payable|trade creditor|creditors", "TOP.TRADE"),
    ("E", r"share premium|additional paid", "SP"),
    ("E", r"share capital|capital stock|ordinary share|paid.up|registered capital|common stock", "SC"),
    ("E", r"revaluation", "RESV.REVAL"),
    ("E", r"dividend|drawing", "RE.DIV"),
    ("E", r"retained|accumulated (profit|loss|deficit)|profit and loss", "RE"),
    ("E", r"reserve", "RESV.OTHER"),
    ("I", r"interest|investment income|dividend income", "FININC"),
    ("I", r"(gain|loss|profit).*(disposal|sale of (asset|ppe|property|equipment|vehicle))", "OI.DISP"),
    ("I", r"fair value", "OI.FV"),
    ("I", r"associate|joint venture", "SHAREASSOC"),
    ("I", r"service|fee|commission|rendering|maintenance contract", "REV.SERV"),
    ("I", r"sale|revenue|turnover", "REV.GOODS"),
    ("CXT", r"deferred tax", "TAX.DEF"),
    ("CXT", r"(income|corporate|profit) tax", "TAX.CUR"),
    ("CX", r"interest (expense|on)|finance (cost|charge)|loan interest|borrowing cost", "FINCOST"),
    ("CX", r"(gain|loss).*(disposal|sale of (asset|ppe|property|equipment|vehicle))", "OI.DISP"),
    ("X", r"sell|sales|distribution|marketing|advertis|promotion|delivery|freight out|carriage out|"
          r"commission|warrant", "DIST"),
]
FALLBACK = {"A": "TOR.OTHER", "L": "TOP.OTHER", "E": "RESV.OTHER", "I": "OI.OTHER", "C": "COS", "X": "ADMIN",
            "T": "TAX.CUR"}

NATURE_RULES = [
    (r"cost of (goods|sales|inventor)|purchases|material", "INVX"),
    (r"amorti[sz]ation", "AMORT"),
    (r"depreciation", "DEP"),
    (r"salar|wage|payroll|staff|bonus|pension|social security|employee", "EMP"),
    (r"rent|operating lease", "RENT"),
    (r"utilit|electric|water|telephone|internet", "UTIL"),
    (r"advertis|marketing|promotion", "ADV"),
    (r"freight|delivery|carriage|transport", "TRANSP"),
    (r"professional|audit|legal|consult", "PROF"),
    (r"repair|maintenance", "REPAIR"),
    (r"impairment|bad debt|doubtful|credit loss", "IMPAIR"),
    (r"warrant", "WARR"),
    (r"exchange|forex|fx ", "FX"),
]
EXPENSE_LEAVES = {"COS", "DIST", "ADMIN", "OTHEXP"}


def account_class(code, classes=None):
    classes = classes or DEFAULT_CLASSES
    for n in range(len(code), 0, -1):
        if code[:n] in classes:
            return classes[code[:n]]
    return "X"


def suggest(code, name, classes=None):
    """Return (leaf, nature, basis) for one account."""
    cls, low = account_class(code, classes), name.lower()
    leaf, basis = None, ""
    for rule_classes, pattern, target in RULES:
        if cls in rule_classes:
            m = re.search(pattern, low)
            if m:
                leaf, basis = target, f"Keyword rule: '{m.group(0)}'"
                break
    if leaf is None:
        leaf, basis = FALLBACK[cls], "Fallback - review"
        if cls in "CX":
            basis = "Account class default"
    nature = ""
    if leaf in EXPENSE_LEAVES:
        nature = next((n for pattern, n in NATURE_RULES if re.search(pattern, low)), "OTHER")
    return leaf, nature, basis


def load_user_mapping(path):
    path = Path(path)
    df = pd.read_csv(path, dtype=str) if path.suffix.lower() == ".csv" else pd.read_excel(path, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    if not {"account", "afs_code"} <= set(df.columns):
        raise ValueError("Mapping file needs the columns: account, afs_code (and optionally nature)")
    df = df.fillna("")
    return {r["account"].strip(): (r["afs_code"].strip(), r.get("nature", "").strip()) for _, r in df.iterrows()}


def build_mapping(accounts, classes=None, user_mapping=None):
    """accounts: {code: name}. Returns DataFrame[account, account_name, cls, leaf, nature, basis]."""
    user_mapping = user_mapping or {}
    natures = {n for n, _ in NATURES}
    rows = []
    for code in sorted(accounts):
        name = accounts[code]
        leaf, nature, basis = suggest(code, name, classes)
        if code in user_mapping and user_mapping[code][0]:
            u_leaf, u_nat = user_mapping[code]
            if u_leaf in LEAF:
                leaf, basis = u_leaf, "User mapping file"
                nature = u_nat if u_nat in natures else (nature if leaf in EXPENSE_LEAVES else "")
            else:
                basis = f"User mapping code '{u_leaf}' not in taxonomy - auto-mapped, review"
        if not is_pl(leaf) or leaf not in EXPENSE_LEAVES:
            nature = ""
        rows.append({"account": code, "account_name": name, "cls": account_class(code, classes),
                     "leaf": leaf, "nature": nature, "basis": basis})
    return pd.DataFrame(rows)
