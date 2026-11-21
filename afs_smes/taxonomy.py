"""AFS taxonomy: the financial-statement components that ledger accounts are mapped to.

Two levels:
  * leaf  - note-level component an account is mapped to (e.g. TOR.TRADE "Trade receivables")
  * face  - line item presented on the face of a primary statement (e.g. TOR "Trade and other receivables")

Sign convention for all ledger data: debit positive, credit negative.
"""

# face code -> (caption, statement, section)
FACES = [
    # --- Statement of financial position ---
    ("PPE", "Property, plant and equipment", "SoFP", "NCA"),
    ("INVPROP", "Investment property", "SoFP", "NCA"),
    ("INTANG", "Intangible assets", "SoFP", "NCA"),
    ("INVASSOC", "Investments in associates and jointly controlled entities", "SoFP", "NCA"),
    ("OFA_NC", "Other financial assets", "SoFP", "NCA"),
    ("DTA", "Deferred tax assets", "SoFP", "NCA"),
    ("INV", "Inventories", "SoFP", "CA"),
    ("TOR", "Trade and other receivables", "SoFP", "CA"),
    ("CONTA", "Contract assets", "SoFP", "CA"),
    ("CTA", "Current tax assets", "SoFP", "CA"),
    ("OFA_C", "Other financial assets", "SoFP", "CA"),
    ("CASH", "Cash and cash equivalents", "SoFP", "CA"),
    ("SC", "Share capital", "SoFP", "EQ"),
    ("SP", "Share premium", "SoFP", "EQ"),
    ("RESV", "Other reserves", "SoFP", "EQ"),
    ("RE", "Retained earnings", "SoFP", "EQ"),
    ("BORR_NC", "Borrowings", "SoFP", "NCL"),
    ("EBO_NC", "Employee benefit obligations", "SoFP", "NCL"),
    ("PROV_NC", "Provisions", "SoFP", "NCL"),
    ("DTL", "Deferred tax liabilities", "SoFP", "NCL"),
    ("TOP", "Trade and other payables", "SoFP", "CL"),
    ("CONTL", "Contract liabilities", "SoFP", "CL"),
    ("BORR_C", "Borrowings", "SoFP", "CL"),
    ("CTL", "Current tax liabilities", "SoFP", "CL"),
    ("PROV_C", "Provisions", "SoFP", "CL"),
    # --- Statement of comprehensive income (expenses by function) ---
    ("REV", "Revenue", "SoCI", "PL"),
    ("COS", "Cost of sales", "SoCI", "PL"),
    ("OI", "Other income", "SoCI", "PL"),
    ("DIST", "Distribution costs", "SoCI", "PL"),
    ("ADMIN", "Administrative expenses", "SoCI", "PL"),
    ("OTHEXP", "Other expenses", "SoCI", "PL"),
    ("FININC", "Finance income", "SoCI", "PL"),
    ("FINCOST", "Finance costs", "SoCI", "PL"),
    ("SHAREASSOC", "Share of profit of associates and jointly controlled entities", "SoCI", "PL"),
    ("TAX", "Income tax expense", "SoCI", "PL"),
]
FACE = {f[0]: f for f in FACES}

SECTION_TITLES = {
    "NCA": "Non-current assets", "CA": "Current assets", "EQ": "Equity",
    "NCL": "Non-current liabilities", "CL": "Current liabilities",
}

# leaf code, caption, face, movement class (drives GL movement-type classification)
LEAVES = [
    ("PPE.LB.C", "Land and buildings - cost or valuation", "PPE", "PPE_C"),
    ("PPE.LB.D", "Land and buildings - accumulated depreciation", "PPE", "PPE_D"),
    ("PPE.PM.C", "Plant and machinery - cost", "PPE", "PPE_C"),
    ("PPE.PM.D", "Plant and machinery - accumulated depreciation", "PPE", "PPE_D"),
    ("PPE.MV.C", "Motor vehicles - cost", "PPE", "PPE_C"),
    ("PPE.MV.D", "Motor vehicles - accumulated depreciation", "PPE", "PPE_D"),
    ("PPE.FE.C", "Furniture, fittings and equipment - cost", "PPE", "PPE_C"),
    ("PPE.FE.D", "Furniture, fittings and equipment - accumulated depreciation", "PPE", "PPE_D"),
    ("INVPROP", "Investment property", "INVPROP", ""),
    ("INTANG.C", "Intangible assets - cost", "INTANG", "INT_C"),
    ("INTANG.D", "Intangible assets - accumulated amortisation", "INTANG", "INT_D"),
    ("INVASSOC", "Investments in associates and jointly controlled entities", "INVASSOC", ""),
    ("OFA_NC", "Other financial assets - non-current", "OFA_NC", ""),
    ("DTA", "Deferred tax assets", "DTA", "DT"),
    ("INV.RM", "Raw materials and consumables", "INV", ""),
    ("INV.WIP", "Work in progress", "INV", ""),
    ("INV.FG", "Finished goods and goods for resale", "INV", ""),
    ("TOR.TRADE", "Trade receivables", "TOR", ""),
    ("TOR.ALLOW", "Allowance for impairment of trade receivables", "TOR", ""),
    ("TOR.OTHER", "Other receivables", "TOR", ""),
    ("TOR.PREPAY", "Prepayments", "TOR", ""),
    ("TOR.INT", "Interest receivable", "TOR", ""),
    ("CONTA", "Contract assets", "CONTA", ""),
    ("CTA", "Current tax assets", "CTA", ""),
    ("OFA_C", "Other financial assets - current", "OFA_C", ""),
    ("CASH.HAND", "Cash on hand", "CASH", "CASH"),
    ("CASH.BANK", "Bank balances", "CASH", "CASH"),
    ("CASH.DEP", "Short-term deposits", "CASH", "CASH"),
    ("SC", "Share capital", "SC", "EQ"),
    ("SP", "Share premium", "SP", "EQ"),
    ("RESV.REVAL", "Revaluation reserve", "RESV", "RESV"),
    ("RESV.OTHER", "Other reserves", "RESV", "RESV"),
    ("RE", "Retained earnings", "RE", "RE"),
    ("RE.DIV", "Dividends declared during the year", "RE", ""),
    ("BORR_NC.LOAN", "Bank loans - non-current portion", "BORR_NC", "BORR"),
    ("BORR_NC.LEASE", "Finance lease obligations - non-current portion", "BORR_NC", "BORR"),
    ("EBO_NC", "Long-term employee benefit obligations", "EBO_NC", ""),
    ("PROV_NC", "Provisions - non-current", "PROV_NC", "PROV"),
    ("DTL", "Deferred tax liabilities", "DTL", "DT"),
    ("TOP.TRADE", "Trade payables", "TOP", ""),
    ("TOP.ACCR", "Accrued expenses", "TOP", ""),
    ("TOP.OTHER", "Other payables (including indirect taxes)", "TOP", ""),
    ("TOP.INT", "Interest payable", "TOP", ""),
    ("TOP.DIV", "Dividends payable", "TOP", ""),
    ("CONTL", "Contract liabilities", "CONTL", ""),
    ("BORR_C.LOAN", "Bank loans - current portion", "BORR_C", "BORR"),
    ("BORR_C.LEASE", "Finance lease obligations - current portion", "BORR_C", "BORR"),
    ("BORR_C.OD", "Bank overdrafts", "BORR_C", "CASH"),
    ("CTL", "Current tax liabilities", "CTL", ""),
    ("PROV_C", "Provisions - current", "PROV_C", "PROV"),
    ("REV.GOODS", "Revenue from the sale of goods", "REV", ""),
    ("REV.SERV", "Revenue from the rendering of services", "REV", ""),
    ("COS", "Cost of sales", "COS", ""),
    ("OI.OTHER", "Sundry income", "OI", ""),
    ("OI.DISP", "Gain/(loss) on disposal of property, plant and equipment", "OI", ""),
    ("OI.FV", "Fair value gain/(loss) on investment property", "OI", ""),
    ("DIST", "Distribution costs", "DIST", ""),
    ("ADMIN", "Administrative expenses", "ADMIN", ""),
    ("OTHEXP", "Other expenses", "OTHEXP", ""),
    ("FININC", "Finance income", "FININC", ""),
    ("FINCOST", "Finance costs", "FINCOST", ""),
    ("SHAREASSOC", "Share of profit of associates and jointly controlled entities", "SHAREASSOC", ""),
    ("TAX.CUR", "Current tax expense", "TAX", ""),
    ("TAX.DEF", "Deferred tax expense", "TAX", ""),
]
LEAF = {l[0]: l for l in LEAVES}


def leaves_of(*faces):
    return [l[0] for l in LEAVES if l[2] in faces]


def is_pl(leaf):
    return leaf in LEAF and FACE[LEAF[leaf][2]][2] == "SoCI"


PPE_CLASSES = [
    ("LB", "Land and buildings"), ("PM", "Plant and machinery"),
    ("MV", "Motor vehicles"), ("FE", "Furniture, fittings and equipment"),
]
PPE_COST = [f"PPE.{c}.C" for c, _ in PPE_CLASSES]
PPE_DEP = [f"PPE.{c}.D" for c, _ in PPE_CLASSES]
BORR = ["BORR_NC.LOAN", "BORR_NC.LEASE", "BORR_C.LOAN", "BORR_C.LEASE"]
CASH_LEAVES = ["CASH.HAND", "CASH.BANK", "CASH.DEP", "BORR_C.OD"]
EXPENSE_FUNCTIONS = ["COS", "DIST", "ADMIN", "OTHEXP"]

# nature of expense (analysis of expenses note)
NATURES = [
    ("INVX", "Cost of inventories recognised as an expense"),
    ("EMP", "Employee benefits expense"),
    ("DEP", "Depreciation of property, plant and equipment"),
    ("AMORT", "Amortisation of intangible assets"),
    ("RENT", "Operating lease expense"),
    ("UTIL", "Utilities"),
    ("ADV", "Advertising and promotion"),
    ("TRANSP", "Freight and delivery"),
    ("PROF", "Professional fees"),
    ("REPAIR", "Repairs and maintenance"),
    ("IMPAIR", "Impairment loss on trade receivables"),
    ("WARR", "Warranty costs"),
    ("FX", "Net foreign exchange loss"),
    ("OTHER", "Other expenses"),
]

# movement types assigned to GL lines on roll-forward accounts
MVT_TYPES = [
    ("ADD", "Addition paid or payable in cash"),
    ("ADD_NC", "Addition financed by borrowings/finance lease (non-cash)"),
    ("DISP", "Disposal / derecognition"),
    ("DEP", "Depreciation / amortisation charge"),
    ("REVAL", "Revaluation"),
    ("PROCEEDS", "Borrowings drawn down (cash)"),
    ("REPAY", "Borrowings repaid (cash)"),
    ("RECLASS", "Transfer between current and non-current"),
    ("NONCASH", "Non-cash change in borrowings"),
    ("OCI", "Recognised in other comprehensive income"),
    ("PL", "Recognised in profit or loss"),
    ("TRANSFER", "Transfer between reserves"),
    ("DIVIDEND", "Dividend charged directly to retained earnings"),
    ("ISSUE", "Issue of shares"),
    ("USED", "Provision utilised"),
    ("REVERSED", "Provision reversed unused"),
    ("OTHER", "Other - investigate"),
]
