"""Illustrative data for the IFRS 18 working-paper generator.

Figures are adapted from the fictitious group in KPMG's "Guide to annual financial
statements - Initial application of IFRS 18" (October 2026), EUR thousands.
They exist only to demonstrate the workbook and prove that every check ties.
Replace them with the entity's trial balance (see assets/tb_template.csv).

Trial balance sign convention: debit = positive, credit = negative.
Columns: account, description, type (PL / OCI / BS), cy, py, ias1_line,
ifrs18_line, nature, mpm_item, transition_ref
"""

# fmt: off
TB = [
    # ---------------- Statement of profit or loss ----------------
    # Operating category
    ("4000", "Sales of goods and services",                         "PL", -102394, -96421, "Revenue", "Revenue", "", "", ""),
    ("5000", "Cost of sales - materials, conversion and overheads", "PL", 36833, 37798,  "Cost of sales", "Cost of sales", "Raw materials and other costs", "", ""),
    ("5010", "Cost of sales - depreciation",                        "PL", 3193, 3063,    "Cost of sales", "Cost of sales", "Depreciation", "", ""),
    ("5020", "Cost of sales - amortisation",                        "PL", 785, 795,      "Cost of sales", "Cost of sales", "Amortisation", "", ""),
    ("5030", "Cost of sales - employee benefits",                   "PL", 14737, 12932,  "Cost of sales", "Cost of sales", "Employee benefits", "", ""),
    ("5040", "Impairment (reversal) - property, plant and equipment","PL", -393, 1123,   "Cost of sales", "Cost of sales", "Impairment losses (IAS 36)", "Impairment (reversal) - PPE", ""),
    ("5041", "Impairment (reversal) - intangible assets",           "PL", -100, 285,     "Cost of sales", "Cost of sales", "Impairment losses (IAS 36)", "Impairment (reversal) - intangibles", ""),
    ("5050", "Inventory write-downs - routine obsolescence",        "PL", 345, 125,      "Cost of sales", "Cost of sales", "Inventory write-downs (IAS 2)", "", ""),
    ("5051", "Inventory write-down (reversal) - impairment event",  "PL", -10, 42,       "Cost of sales", "Cost of sales", "Inventory write-downs (IAS 2)", "Inventory write-down (reversal)", ""),
    ("5060", "Net interest income on defined benefit plan",         "PL", -23, -22,      "Cost of sales", "Interest and other financing expenses", "", "", "B"),
    ("5070", "Investment property maintenance",                     "PL", 50, 45,        "Cost of sales", "Administrative expenses", "Raw materials and other costs", "", "A"),
    ("6000", "Selling and distribution - depreciation",             "PL", 1056, 1013,    "Selling and distribution expenses", "Selling and distribution expenses", "Depreciation", "", ""),
    ("6010", "Selling and distribution - employee benefits",        "PL", 4875, 4278,    "Selling and distribution expenses", "Selling and distribution expenses", "Employee benefits", "", ""),
    ("6020", "Selling and distribution - other",                    "PL", 12391, 10574,  "Selling and distribution expenses", "Selling and distribution expenses", "Raw materials and other costs", "", ""),
    ("6100", "Administrative - depreciation",                       "PL", 641, 615,      "Administrative expenses", "Administrative expenses", "Depreciation", "", ""),
    ("6110", "Administrative - employee benefits",                  "PL", 2957, 2595,    "Administrative expenses", "Administrative expenses", "Employee benefits", "", ""),
    ("6120", "Administrative - other",                              "PL", 14149, 11218,  "Administrative expenses", "Administrative expenses", "Raw materials and other costs", "", ""),
    ("6200", "Research and development - depreciation",             "PL", 200, 180,      "Research and development expenses", "Research and development expenses", "Depreciation", "", ""),
    ("6210", "Research and development - employee benefits",        "PL", 909, 517,      "Research and development expenses", "Research and development expenses", "Employee benefits", "", ""),
    ("6300", "Impairment loss on goodwill",                         "PL", 116, 0,        "Other expenses", "Impairment losses on goodwill and remeasurement of disposal group", "Impairment losses (IAS 36)", "Impairment loss - goodwill", ""),
    ("6310", "Remeasurement loss - disposal group held for sale",   "PL", 35, 0,         "Other expenses", "Impairment losses on goodwill and remeasurement of disposal group", "Other remeasurements", "Remeasurement - disposal group", ""),
    ("6400", "Impairment loss on trade receivables and contract assets","PL", 200, 190,  "Impairment loss on trade receivables and contract assets", "Impairment loss on trade and other receivables and contract assets", "Other remeasurements", "", ""),
    ("6500", "Foreign exchange differences - trading items",        "PL", 283, 245,      "Finance costs", "Foreign exchange differences", "", "", "C"),
    ("6600", "Change in fair value of biological assets",           "PL", -587, -28,     "Other income", "Other operating income", "", "", ""),
    ("6610", "Government grants released",                          "PL", -238, 0,       "Other income", "Other operating income", "", "", ""),
    ("6620", "Gain on sale of property, plant and equipment",       "PL", -48, -16,      "Other income", "Other operating income", "", "", ""),
    ("6630", "Cash flow hedges - reclassified from hedging reserve (sales)","PL", -12, -3,"Other income", "Other operating income", "", "", ""),
    ("6640", "Economic hedges of sales - FX forwards at FVTPL",     "PL", -25, -18,      "Finance costs", "Other operating income", "", "", "F"),
    ("6700", "Settlement of pre-existing relationship with acquiree","PL", 326, 0,       "Other expenses", "Other operating expenses, including flood-related expenses", "", "", ""),
    ("6710", "Flood-related expenses",                              "PL", 519, 0,        "Other expenses", "Other operating expenses, including flood-related expenses", "", "Flood-related expenses", ""),
    ("6720", "Change in fair value of contingent consideration",    "PL", 20, 0,         "Other expenses", "Other operating expenses, including flood-related expenses", "", "", ""),
    ("6730", "Cash flow hedges - costs of hedging reclassified",    "PL", 6, 7,          "Other expenses", "Other operating expenses, including flood-related expenses", "", "", ""),
    ("6740", "Cash flow hedges (sales) - ineffective portion",      "PL", 35, 8,         "Finance costs", "Other operating expenses, including flood-related expenses", "", "", "E"),
    ("6750", "Net investment hedge - ineffective portion",          "PL", 1, 0,          "Other expenses", "Other operating expenses, including flood-related expenses", "", "", ""),
    # Investing category
    ("7000", "Share of profit of equity-accounted investees, net of tax","PL", -1141, -587,"Share of profit of equity-accounted investees, net of tax", "Share of profit of equity-accounted investees, net of tax", "", "", "G"),
    ("7100", "Rental income - investment property",                 "PL", -330, -302,    "Revenue", "Income from investment properties", "", "", "A"),
    ("7110", "Fair value gain - investment property",               "PL", -150, -60,     "Other income", "Income from investment properties", "", "", "A"),
    ("7200", "Interest income - debt securities and cash (EIR)",    "PL", -178, -141,    "Finance income", "Interest and other investing income", "", "", "D"),
    ("7210", "Interest income - net investment in finance lease",   "PL", -31, -10,      "Finance income", "Interest and other investing income", "", "", "D"),
    ("7220", "Dividend income - equity securities at FVOCI",        "PL", -26, -32,      "Finance income", "Interest and other investing income", "", "", "D"),
    ("7230", "Gain on derecognition of debt securities at FVOCI",   "PL", -64, 0,        "Finance income", "Interest and other investing income", "", "", "D"),
    ("7240", "Sovereign debt held for trading - FV change",         "PL", -49, 37,       "Finance costs", "Interest and other investing income", "", "", "D"),
    ("7250", "Impairment loss on debt securities (ECL)",            "PL", 59, 13,        "Finance costs", "Interest and other investing income", "", "", "D"),
    ("7260", "Foreign exchange differences - investments",          "PL", 25, 35,        "Finance costs", "Interest and other investing income", "", "", "C"),
    ("7300", "Remeasurement of pre-existing interest in acquiree",  "PL", -250, 0,       "Other income", "Remeasurement to fair value of pre-existing interest in an acquiree", "", "", ""),
    ("7400", "Equity securities mandatorily at FVTPL - FV change",  "PL", -508, -264,    "Finance income", "Financial assets at FVTPL - net change in fair value", "", "", "D"),
    # Financing category
    ("8000", "Interest expense - financial liabilities at amortised cost","PL", 1523, 1299,"Finance costs", "Interest and other financing expenses", "", "", ""),
    ("8010", "Unwind of discount - site restoration provision",     "PL", 60, 50,        "Finance costs", "Interest and other financing expenses", "", "", ""),
    ("8020", "Cash flow hedges (interest) - reclassified from OCI", "PL", -17, -12,      "Finance costs", "Interest and other financing expenses", "", "", ""),
    ("8030", "Cash flow hedges (interest) - ineffective portion",   "PL", 16, 8,         "Finance costs", "Interest and other financing expenses", "", "", ""),
    ("8040", "Foreign exchange differences - borrowings",           "PL", -122, -30,     "Finance costs", "Interest and other financing expenses", "", "", ""),
    # Income taxes and discontinued operations
    ("9000", "Income tax expense",                                  "PL", 3178, 2460,    "Income tax expense", "Income tax expense", "", "", ""),
    ("9100", "Profit (loss) from discontinued operation, net of tax","PL", -379, 422,    "Profit (loss) from discontinued operation, net of tax", "Profit (loss) from discontinued operation, net of tax", "", "", ""),

    # ---------------- Other comprehensive income ----------------
    ("9500", "Revaluation of property, plant and equipment",        "OCI", -200, 0,      "", "Revaluation of property, plant and equipment", "", "", ""),
    ("9510", "Remeasurements of defined benefit liability (asset)", "OCI", -72, 15,      "", "Remeasurements of defined benefit liability (asset)", "", "", ""),
    ("9520", "Equity investments at FVOCI - net change in FV",      "OCI", -141, -59,    "", "Equity investments at FVOCI - net change in fair value", "", "", ""),
    ("9530", "Equity-accounted investees - share of OCI (not reclassified)","OCI", -15, 3,"", "Equity-accounted investees - share of OCI (not reclassified)", "", "", ""),
    ("9540", "Tax on items not reclassified",                       "OCI", 137, 14,      "", "Related tax - items not reclassified", "", "", ""),
    ("9600", "Foreign operations - translation differences",        "OCI", -679, -471,   "", "Foreign operations - foreign currency translation differences", "", "", ""),
    ("9610", "Net investment hedge - net loss",                     "OCI", 3, 8,         "", "Net investment hedge - net loss", "", "", ""),
    ("9620", "Equity-accounted investees - share of OCI (reclassifiable)","OCI", 172, 166,"", "Equity-accounted investees - share of OCI (reclassifiable)", "", "", ""),
    ("9630", "Reclassification of FX differences on loss of significant influence","OCI", 20, 0,"", "Reclassification of foreign currency differences on loss of significant influence", "", "", ""),
    ("9640", "Cash flow hedges - effective portion of FV changes",  "OCI", 62, -95,      "", "Cash flow hedges - effective portion of changes in fair value", "", "", ""),
    ("9650", "Cash flow hedges - reclassified to profit or loss",   "OCI", 31, 12,       "", "Cash flow hedges - reclassified to profit or loss", "", "", ""),
    ("9660", "Cost of hedging reserve - changes in FV",             "OCI", -34, -10,     "", "Cost of hedging reserve - changes in fair value", "", "", ""),
    ("9670", "Cost of hedging reserve - reclassified to P&L",       "OCI", -8, -2,       "", "Cost of hedging reserve - reclassified to profit or loss", "", "", ""),
    ("9680", "Debt investments at FVOCI - net change in FV",        "OCI", -54, -60,     "", "Debt investments at FVOCI - net change in fair value", "", "", ""),
    ("9690", "Debt investments at FVOCI - reclassified to P&L",     "OCI", 64, 0,        "", "Debt investments at FVOCI - reclassified to profit or loss", "", "", ""),
    ("9699", "Tax on items that are or may be reclassified",        "OCI", -19, 48,      "", "Related tax - items that are or may be reclassified", "", "", ""),

    # ---------------- Statement of financial position (closing balances) ----------------
    ("1000", "Property, plant and equipment",                       "BS", 28490, 33230,  "", "Property, plant and equipment", "", "", ""),
    ("1100", "Intangible assets - non-current",                     "BS", 2279, 1134,    "", "Intangible assets (non-current)", "", "", ""),
    ("1110", "Goodwill",                                            "BS", 3832, 3407,    "", "Goodwill", "", "", ""),
    ("1200", "Biological assets",                                   "BS", 4730, 4056,    "", "Biological assets", "", "", ""),
    ("1300", "Investment property",                                 "BS", 1520, 400,     "", "Investment property", "", "", ""),
    ("1400", "Equity-accounted investees",                          "BS", 2489, 1948,    "", "Equity-accounted investees", "", "", ""),
    ("1500", "Financial assets - non-current",                      "BS", 3616, 3512,    "", "Financial assets (non-current)", "", "", ""),
    ("1600", "Deferred tax assets",                                 "BS", 2251, 2108,    "", "Deferred tax assets", "", "", ""),
    ("1700", "Employee benefit assets",                             "BS", 671, 731,      "", "Employee benefits (asset)", "", "", ""),
    ("2000", "Intangible assets - current (emission allowances)",   "BS", 115, 120,      "", "Intangible assets (current)", "", "", ""),
    ("2100", "Inventories",                                         "BS", 12148, 12119,  "", "Inventories", "", "", ""),
    ("2200", "Contract assets",                                     "BS", 1271, 782,     "", "Contract assets", "", "", ""),
    ("2300", "Financial assets - current",                          "BS", 662, 1032,     "", "Financial assets (current)", "", "", ""),
    ("2400", "Current tax assets",                                  "BS", 34, 60,        "", "Current tax assets", "", "", ""),
    ("2500", "Trade and other receivables",                         "BS", 32518, 22325,  "", "Trade and other receivables", "", "", ""),
    ("2600", "Prepayments",                                         "BS", 330, 1200,     "", "Prepayments", "", "", ""),
    ("2700", "Cash and cash equivalents",                           "BS", 1504, 1849,    "", "Cash and cash equivalents", "", "", ""),
    ("2800", "Assets held for sale",                                "BS", 14400, 0,      "", "Assets held for sale", "", "", ""),
    ("3000", "Share capital",                                       "BS", -14979, -14550,"", "Share capital", "", "", ""),
    ("3010", "Share premium",                                       "BS", -4777, -3500,  "", "Share premium", "", "", ""),
    ("3020", "Reserves - balance before current-period OCI",        "BS", -451, -30,     "", "Reserves", "", "", ""),
    ("3030", "Retained earnings - balance before current-period profit","BS", -13449, -8163,"", "Retained earnings", "", "", ""),
    ("3040", "Non-controlling interests - balance before current-period TCI","BS", -3281, -2718,"", "Non-controlling interests", "", "", ""),
    ("3100", "Loans and borrowings - non-current",                  "BS", -23311, -21268,"", "Loans and borrowings (non-current)", "", "", ""),
    ("3110", "Employee benefits - non-current",                     "BS", -912, -453,    "", "Employee benefits (non-current liability)", "", "", ""),
    ("3120", "Trade and other payables - non-current",              "BS", -290, -5,      "", "Trade and other payables (non-current)", "", "", ""),
    ("3130", "Deferred income - non-current",                       "BS", -1424, -1462,  "", "Deferred income (non-current)", "", "", ""),
    ("3140", "Provisions - non-current",                            "BS", -810, 0,       "", "Provisions (non-current)", "", "", ""),
    ("3150", "Deferred tax liabilities",                            "BS", -519, -406,    "", "Deferred tax liabilities", "", "", ""),
    ("3200", "Bank overdraft",                                      "BS", -334, -282,    "", "Bank overdraft", "", "", ""),
    ("3210", "Current tax liabilities",                             "BS", -4781, -1693,  "", "Current tax liabilities", "", "", ""),
    ("3220", "Loans and borrowings - current",                      "BS", -5794, -5656,  "", "Loans and borrowings (current)", "", "", ""),
    ("3230", "Employee benefits - current",                         "BS", -20, -388,     "", "Employee benefits (current liability)", "", "", ""),
    ("3240", "Trade and other payables - current",                  "BS", -24013, -21328,"", "Trade and other payables (current)", "", "", ""),
    ("3250", "Contract liabilities",                                "BS", -160, -166,    "", "Contract liabilities", "", "", ""),
    ("3260", "Provisions - current",                                "BS", -860, -1540,   "", "Provisions (current)", "", "", ""),
    ("3270", "Liabilities directly associated with assets held for sale","BS", -4410, 0,  "", "Liabilities directly associated with assets held for sale", "", "", ""),
]
# fmt: on

TRANSITION_EXPLANATIONS = [
    ("A", "Investment property", "Rental income (previously in revenue) and fair value changes (previously in other income) "
          "are classified in the investing category: the property generates a return individually and largely independently "
          "of other resources and investing in property is not a specified main business activity. Maintenance expense moved "
          "from cost of sales to administrative expenses.", "IFRS 18.53(c), 54, B47"),
    ("B", "Interest on liabilities that do not involve only the raising of finance",
          "Net interest on the defined benefit plan, previously in cost of sales, is classified in the financing category because "
          "it is identified for the purpose of applying IAS 19.", "IFRS 18.59(b), 60-61, B52-B54"),
    ("C", "Foreign exchange differences", "FX differences previously presented in net finance costs are classified in the same "
          "category as the income and expenses from the items that gave rise to them (operating, investing or financing).",
          "IFRS 18.B65-B67"),
    ("D", "Income and expenses on financial assets", "Interest, dividends, FV changes, derecognition gains and impairment on "
          "investments in financial assets, previously in finance income/costs, are classified in the investing category.",
          "IFRS 18.53(c), B46-B47, B60"),
    ("E", "Cash flow hedges - ineffectiveness", "The ineffective portion of designated hedges of forecast sales is classified "
          "in the same category as the effective portion (operating).", "IFRS 18.B71"),
    ("F", "Economic hedges", "Gains and losses on derivatives used to manage identified risks but not designated are classified "
          "in the category affected by the risk managed (operating - sales).", "IFRS 18.B70, B72"),
    ("G", "Share of profit of equity-accounted investees", "Always classified in the investing category; previously presented "
          "below net finance costs.", "IFRS 18.53(a), 55(a), B43"),
]

SETUP = {
    "entity": "Illustrative Group (demo data adapted from KPMG IFRS 18 guide, Oct 2026)",
    "cy_label": "2026",
    "py_label": "2025",
    "cy_end": "31 December 2026",
    "py_end": "31 December 2025",
    "currency": "EUR",
    "units": "thousands",
    "date_initial_application": "1 January 2026 (early adoption)",
    "early_adoption": "Yes",
    "smba_investing": "No",
    "smba_financing": "No",
    "expense_method": "Mixed",
    "tax_rate": 0.32,
    "nci_profit_cy": 497, "nci_profit_py": 351,
    "nci_oci_cy": 26, "nci_oci_py": 22,
    "pref_div_cy": 438, "pref_div_py": 438,
    "diluted_adj_cy": 61, "diluted_adj_py": 0,
    "wans_basic_cy": 3083, "wans_basic_py": 3060,
    "wans_diluted_cy": 3226, "wans_diluted_py": 3086,
    "owners_share_mpm": 0.991,
    "materiality": 450, "performance_materiality": 340, "trivial": 22,
}

# Statement of cash flows inputs (IAS 7 as amended by IFRS 18). Operating profit is linked, not input.
CASHFLOW = {
    "op_adjustments": [
        ("Depreciation", 5178, 5010),
        ("Amortisation", 736, 748),
        ("Deferred income released", -38, 0),
        ("(Reversal of) impairment losses on property, plant and equipment", -393, 1123),
        ("(Reversal of) impairment losses on intangible assets", -100, 285),
        ("Impairment losses on goodwill and remeasurement of disposal group", 151, 0),
        ("Change in fair value of biological assets", -587, -28),
        ("Change in fair value of derivatives at FVTPL (operating)", -25, -18),
        ("Change in fair value of contingent consideration", 20, 0),
        ("Ineffective portion of hedges", 36, 8),
        ("Foreign exchange differences (operating)", 283, 245),
        ("Gain on sale of property, plant and equipment", -48, -16),
        ("Equity-settled share-based payment", 755, 248),
    ],
    "wc_changes": [
        ("Intangible assets (emission allowances)", 5, -20),
        ("Inventories", -1843, -176),
        ("Contract assets", -489, -782),
        ("Trade and other receivables", -16946, -4920),
        ("Contract liabilities", -6, 166),
        ("Prepayments", 870, -305),
        ("Trade and other payables", 8727, -6841),
        ("Provisions and employee benefits", 26, 274),
    ],
    "tax_paid": (-400, -1910),
    "disc_ops_operating": (531, 17),
    "investing": [
        ("Interest received", 37, 29),
        ("Dividends received", 47, 32),
        ("Proceeds from sale of property, plant and equipment", 1179, 397),
        ("Proceeds from sale of investments", 1589, 534),
        ("Disposal of discontinued operation, net of cash disposed of", 10890, 0),
        ("Acquisition of subsidiary, net of cash acquired", -1799, 0),
        ("Acquisition of property, plant and equipment", -15857, -2228),
        ("Acquisition of investment property", -300, -40),
        ("Purchase of non-current biological assets", -305, -835),
        ("Acquisition of financial assets", -602, -342),
        ("Development expenditure", -1235, -503),
        ("Receipt of asset-related government grant", 130, 1462),
    ],
    "financing": [
        ("Proceeds from issue of share capital", 1550, 0),
        ("Proceeds from issue of convertible notes", 5000, 0),
        ("Proceeds from issue of redeemable preference shares", 2000, 0),
        ("Proceeds from loans and borrowings", 591, 4079),
        ("Proceeds from sale of treasury shares", 30, 0),
        ("Proceeds from exercise of share options", 50, 0),
        ("Proceeds from settlement of derivatives", 5, 11),
        ("Transaction costs related to loans and borrowings", -311, 0),
        ("Acquisition of non-controlling interests", -200, 0),
        ("Repurchase of treasury shares", 0, -280),
        ("Repayment of borrowings", -5055, -2445),
        ("Payment of lease liabilities", -554, -590),
        ("Dividends paid", -1243, -571),
        ("Interest paid", -1609, -1289),
    ],
    "opening_cash": (1567, 2226),
    "fx_on_cash": (-13, -26),
}

# Expenses capitalised in the period (needed for the IFRS 18.83(a) "total amount" disclosure)
CAPITALISED = {
    # function line: {nature: (cy, py)}
    "Cost of sales": {"Employee benefits": (0, 0)},
    "Research and development expenses": {"Depreciation": (249, 251), "Employee benefits": (1163, 1301)},
}
