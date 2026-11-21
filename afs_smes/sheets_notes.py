"""Notes to the financial statements (structure of Section 8), built with live formulas."""
from openpyxl.styles import Alignment

from .taxonomy import BORR, EXPENSE_FUNCTIONS, FACE, LEAF, NATURES, PPE_CLASSES, PPE_COST, PPE_DEP, leaves_of
from .xl import (B_BOTTOM, B_TOP, B_TOTAL, CY, F_BOLD, F_GREY, F_INPUT, F_ITALIC, F_NORMAL, F_SUB, F_TITLE, FILL_INPUT,
                 NUM, PCT, PY, RIGHT, face, glm, lead, nat, put, widths)

# order of the notes; a note with a face-line condition is included only when that line has balances
NOTE_PLAN = [
    ("general", None), ("basis", None), ("policies", None), ("judgements", None), ("revenue", None),
    ("otherinc", ("OI",)), ("expenses", None), ("finance", None), ("pbt", None), ("tax", None),
    ("ppe", ("PPE",)), ("invprop", ("INVPROP",)), ("intang", ("INTANG",)), ("invassoc", ("INVASSOC",)),
    ("ofa", ("OFA_NC", "OFA_C")), ("inv", ("INV",)), ("tor", ("TOR",)), ("cash", None), ("equity", None),
    ("borr", ("BORR_NC", "BORR_C")), ("top", ("TOP",)), ("prov", ("PROV_NC", "PROV_C")), ("ebo", ("EBO_NC",)),
    ("dtax", ("DTA", "DTL")), ("fininst", None), ("dividends", None), ("commitments", None),
    ("contingencies", None), ("related", None), ("events", None), ("approval", None),
]
DT = ["DTL", "DTA"]

POLICIES = [
    ("Revenue from contracts with customers",
     "Revenue is recognised to depict the transfer of promised goods or services to customers at the amount of consideration to which the "
     "Company expects to be entitled, excluding amounts collected on behalf of third parties such as value added tax. The Company identifies "
     "the contract and the promises in it, determines the transaction price, allocates it to the promises and recognises revenue when (or as) "
     "each promise is satisfied. Revenue from the sale of goods is recognised at the point in time when control of the goods passes to the "
     "customer, normally on delivery. Revenue from services is recognised over time as the services are rendered. Consideration received "
     "before the goods or services are transferred is presented as a contract liability; a right to consideration that is conditional on "
     "something other than the passage of time is presented as a contract asset."),
    ("Property, plant and equipment",
     "Items of property, plant and equipment are measured at cost less accumulated depreciation and any accumulated impairment losses, except "
     "land, which is carried at a revalued amount, being its fair value at the date of revaluation. Revaluation increases are recognised in "
     "other comprehensive income and accumulated in the revaluation reserve. Depreciation is charged so as to allocate the cost of assets less "
     "their residual values over their estimated useful lives, using the straight-line method. Land is not depreciated. The annual rates used "
     "are: [buildings 2%; plant and machinery 10%; motor vehicles 20%; furniture, fittings and equipment 10%]. Residual values, useful lives and "
     "depreciation methods are reviewed if there is an indication of a significant change since the last reporting date."),
    ("Intangible assets",
     "Purchased intangible assets (computer software) are stated at cost less accumulated amortisation and any accumulated impairment losses. "
     "They are amortised over their estimated useful life of [five] years using the straight-line method. If the useful life cannot be "
     "established reliably it is determined based on management's best estimate but does not exceed ten years."),
    ("Impairment of non-financial assets",
     "At each reporting date, property, plant and equipment, intangible assets and inventories are reviewed to determine whether there is any "
     "indication that they have suffered an impairment loss. If so, the recoverable amount of the affected asset (or group of related assets) "
     "is estimated and compared with its carrying amount; any shortfall is recognised immediately in profit or loss. Inventories are assessed "
     "by comparing the carrying amount of each item with its selling price less costs to complete and sell."),
    ("Inventories",
     "Inventories are stated at the lower of cost and estimated selling price less costs to complete and sell. Cost is calculated using the "
     "[weighted average cost / first-in, first-out] method and includes all costs of purchase and other costs incurred in bringing the "
     "inventories to their present location and condition."),
    ("Financial instruments",
     "Trade and other receivables, cash and bank balances, trade and other payables and bank loans are basic financial instruments. They are "
     "recognised initially at the transaction price (including transaction costs) and subsequently measured at amortised cost using the "
     "effective interest method; receivables and payables due within one year on normal credit terms are measured at the undiscounted amount "
     "of cash expected to be received or paid. At the end of each reporting period the carrying amounts of financial assets are reviewed for "
     "objective evidence of impairment; an impairment loss is recognised in profit or loss. Financial assets are derecognised when the "
     "contractual rights to the cash flows expire or are settled; financial liabilities when the obligation is discharged, cancelled or expires."),
    ("Cash and cash equivalents",
     "Cash and cash equivalents comprise cash on hand, demand deposits and short-term, highly liquid investments with a maturity of three "
     "months or less from the date of acquisition that are subject to an insignificant risk of changes in value, net of bank overdrafts "
     "repayable on demand."),
    ("Leases",
     "Leases are classified as finance leases when the terms of the lease transfer substantially all the risks and rewards incidental to "
     "ownership of the leased asset to the Company. All other leases are classified as operating leases. Rentals payable under operating "
     "leases are charged to profit or loss on a straight-line basis over the term of the lease."),
    ("Provisions",
     "A provision is recognised when the Company has a present obligation as a result of a past event, it is probable that an outflow of "
     "economic benefits will be required to settle it, and the amount can be estimated reliably. Provisions are measured at the best estimate "
     "of the amount required to settle the obligation at the reporting date. The provision for warranties covers the expected cost of "
     "repairing or replacing goods sold under assurance-type warranty terms."),
    ("Employee benefits",
     "Short-term employee benefits, including salaries, bonuses, paid leave and social security contributions, are recognised as an expense in "
     "the period in which the employees render the related service. [Describe any post-employment or other long-term benefit plans.]"),
    ("Borrowing costs", "All borrowing costs are recognised in profit or loss in the period in which they are incurred."),
    ("Income tax",
     "Income tax expense represents the sum of the tax currently payable and deferred tax. Current tax is based on taxable profit for the "
     "year. Deferred tax is recognised on differences between the carrying amounts of assets and liabilities in the financial statements and "
     "their corresponding tax bases (temporary differences), using tax rates that have been enacted or substantively enacted by the reporting "
     "date. Deferred tax assets are recognised only to the extent that it is probable that taxable profits will be available against which "
     "they can be used. Tax is recognised in profit or loss, except where it relates to an item recognised in other comprehensive income or "
     "directly in equity, in which case the tax is recognised there."),
    ("Foreign currency translation",
     "Transactions in currencies other than the functional currency are recorded at the spot exchange rate at the date of the transaction. "
     "Monetary items denominated in foreign currencies are retranslated at the closing rate at each reporting date. Exchange differences are "
     "recognised in profit or loss in the period in which they arise."),
    ("Dividends", "Dividends to the Company's shareholders are recognised as a liability in the period in which they are declared."),
]


def plan_notes(book):
    """Number the notes that apply to this entity."""
    n = 0
    for key, cond in NOTE_PLAN:
        if cond is None or any(f in book.active_faces for f in cond):
            n += 1
            book.notes[key] = n


class Notes:
    def __init__(self, book, ws, cfg):
        self.book, self.ws, self.cfg, self.r = book, ws, cfg, 5
        self.E = book.ent

    # ---- layout primitives ----
    def title(self, key, text):
        self.r += 1
        put(self.ws, f"A{self.r}", self.book.notes[key], F_SUB, align=Alignment(horizontal="left"))
        put(self.ws, f"B{self.r}", text, F_SUB)
        self.book.reg("Notes", "N_" + key, self.r)
        self.r += 1

    def sub(self, text):
        put(self.ws, f"B{self.r}", text, F_BOLD)
        self.r += 1

    def para(self, text, inp=False, placeholder=None):
        """Paragraph merged across the sheet. Formula text (starting with =) is supported."""
        ws, r = self.ws, self.r
        ws.merge_cells(f"B{r}:H{r}")
        is_placeholder = placeholder if placeholder is not None else ("[" in text and not text.startswith("="))
        put(ws, f"B{r}", text, F_INPUT if (inp or is_placeholder) else F_NORMAL,
            align=Alignment(wrap_text=True, vertical="top"), fill=FILL_INPUT if is_placeholder else None)
        ws.row_dimensions[r].height = 14.5 * max(1, -(-len(text) // 160)) + 3
        self.r += 2

    def heads(self, labels, start="C", years=False):
        col = ord(start)
        if years:
            labels = [self.book.year(CY), self.book.year(PY)]
        put(self.ws, f"B{self.r}", None, border=B_BOTTOM)
        for i, lab in enumerate(labels):
            c = put(self.ws, f"{chr(col + i)}{self.r}", lab, F_BOLD, "0" if years else None,
                    Alignment(horizontal="right", wrap_text=True, vertical="bottom"), B_BOTTOM)
        if not years and any(len(str(l)) > 14 for l in labels):
            self.ws.row_dimensions[self.r].height = 42
        self.r += 1

    def line(self, caption, formulas, bold=False, border=None, fmt=NUM, start="C", inp=False, key=None):
        font = F_BOLD if bold else F_NORMAL
        put(self.ws, f"B{self.r}", caption, font)
        for i, f in enumerate(formulas):
            if f is None:
                continue
            is_f = isinstance(f, str)
            put(self.ws, f"{chr(ord(start) + i)}{self.r}", ("=" + f) if is_f else f,
                F_INPUT if (inp or not is_f) else font, fmt, border=border)
        if key:
            self.book.reg("Notes", key, self.r)
        self.r += 1
        return self.r - 1

    def total(self, caption, first, last, ncols=2, start="C", key=None, double=True):
        cols = [chr(ord(start) + i) for i in range(ncols)]
        return self.line(caption, [f"SUM({c}{first}:{c}{last})" for c in cols], bold=True,
                         border=B_TOTAL if double else B_TOP, start=start, key=key)

    def leaf_table(self, faces, sign, total_caption, key):
        """Two-year breakdown of face line(s) by AFS component; returns the total row."""
        self.heads(None, years=True)
        leaves = [l for l in leaves_of(*faces) if l in self.book.active]
        first = self.r
        for l in leaves:
            self.line(LEAF[l][1], [sign + lead("close", CY, [l]), sign + lead("close", PY, [l])])
        t = self.total(total_caption, first, self.r - 1, key=key)
        tot_face = lambda fy: "+".join(f"{face('close', fy, f)}" for f in faces)
        for fy, col in ((CY, "C"), (PY, "D")):
            self.book.check(f"Note {self.book.notes[key]} {fy}: '{total_caption}' agrees to the statement of financial position",
                            f"Notes!{col}{t}-({sign}({tot_face(fy)}))")
        self.r += 1
        return t

    def ref(self, key, fy):
        return self.book.cell("Notes", key, fy)

    # ---- the notes ----
    def build(self):
        b, ws, E = self.book, self.ws, self.E
        put(ws, "A1", f"={E['entity_name']}", F_TITLE)
        put(ws, "A2", "Notes to the financial statements", F_SUB)
        put(ws, "A3", f'="For the year ended "&TEXT({E["reporting_date"]},"d mmmm yyyy")')
        put(ws, "A4", f'="Expressed in "&{E["currency"]}&". Blue text and yellow cells are entity-specific narrative to be completed or confirmed."', F_GREY)
        widths(ws, [5, 60, 16, 16, 16, 16, 16, 16])
        ws.sheet_view.showGridLines = False
        for key, _ in NOTE_PLAN:
            if key in b.notes:
                getattr(self, "n_" + key)()
        ws.freeze_panes = "A5"

    def n_general(self):
        E = self.E
        self.title("general", "General information")
        self.para(f'={E["entity_name"]}&" (the \'Company\') is a "&{E["legal_form"]}&" incorporated and domiciled in "&{E["country"]}'
                  f'&". The address of its registered office and principal place of business is "&{E["registered_office"]}'
                  f'&". The principal activities of the Company are "&{E["principal_activities"]}&"."', placeholder=False)
        self.para("[State the name of the parent and the ultimate controlling party, if any.]")

    def n_basis(self):
        E = self.E
        self.title("basis", "Basis of preparation and statement of compliance")
        self.para("These financial statements have been prepared in accordance with the IFRS for SMEs Accounting Standard (third edition) issued by "
                  "the International Accounting Standards Board. They comply with all the requirements of that Standard.")
        if self.cfg.get("early_application"):
            self.para("The third edition of the Standard is effective for annual periods beginning on or after 1 January 2027. The Company has "
                      "applied it before that date, as permitted. [Describe the transition: the requirements applied retrospectively, any "
                      "transition reliefs used and the effect on the comparative information.]")
        self.para(f'="The financial statements are presented in "&{E["currency"]}&", which is the Company\'s functional currency, and are '
                  f'prepared on the historical cost basis except for land, which is measured at a revalued amount. They comprise the statement of '
                  f'financial position, the statement of comprehensive income, the statement of changes in equity, the statement of cash flows and '
                  f'these notes, with comparative information for the preceding year."', placeholder=False)
        self.para("Going concern: management has assessed the Company's ability to continue as a going concern, taking into account all available "
                  "information about the future covering at least twelve months from the reporting date, and has not identified any material "
                  "uncertainty that may cast significant doubt upon it. The financial statements are therefore prepared on the going concern basis.")

    def n_policies(self):
        self.title("policies", "Material accounting policy information")
        self.para("The accounting policies set out below have been applied consistently to all periods presented.")
        n = self.book.notes["policies"]
        for i, (head, text) in enumerate(POLICIES, start=1):
            self.sub(f"{n}.{i}  {head}")
            self.para(text, inp="[" in text, placeholder=False)

    def n_judgements(self):
        self.title("judgements", "Significant judgements and key sources of estimation uncertainty")
        self.sub("Judgements in applying accounting policies")
        self.para("[Describe the judgements, apart from those involving estimations, that have the most significant effect on the amounts "
                  "recognised - for example the timing of transfer of control for revenue, or the classification of leases.]")
        self.sub("Key sources of estimation uncertainty")
        self.para("The key assumptions concerning the future that have a significant risk of causing a material adjustment to the carrying "
                  "amounts of assets and liabilities within the next financial year are: (a) the useful lives and residual values of property, "
                  "plant and equipment; (b) the fair value of land carried at a revalued amount; (c) the impairment allowance for trade "
                  "receivables; (d) the provision for warranties; and (e) the write-down of inventories to selling price less costs to complete "
                  "and sell. The carrying amounts concerned are shown in the related notes.", inp=True, placeholder=False)

    def n_revenue(self):
        b = self.book
        self.title("revenue", "Revenue from contracts with customers")
        self.sub("Disaggregation of revenue")
        self.heads(None, years=True)
        first = self.r
        for l in [x for x in leaves_of("REV") if x in b.active]:
            timing = " (point in time)" if l == "REV.GOODS" else " (over time)"
            self.line(LEAF[l][1] + timing, ["-" + lead("close", CY, [l]), "-" + lead("close", PY, [l])])
        t = self.total("Total revenue", first, self.r - 1, key="REV_T")
        for fy, col in ((CY, "C"), (PY, "D")):
            b.check(f"Note {b.notes['revenue']} {fy}: revenue agrees to the statement of comprehensive income",
                    f"Notes!{col}{t}-{b.cell('SoCI', 'REV', fy)}")
        self.r += 1
        self.sub("Contract balances")
        self.heads(None, years=True)
        self.line("Trade receivables, net of impairment allowance",
                  [lead("close", CY, ["TOR.TRADE", "TOR.ALLOW"]), lead("close", PY, ["TOR.TRADE", "TOR.ALLOW"])])
        if "CONTA" in b.active:
            self.line("Contract assets", [lead("close", CY, ["CONTA"]), lead("close", PY, ["CONTA"])])
        self.line("Contract liabilities (advances from customers)", ["-" + lead("close", CY, ["CONTL"]), "-" + lead("close", PY, ["CONTL"])])
        self.line("Revenue recognised from opening contract liabilities",
                  ["-" + lead("open", CY, ["CONTL"]), "-" + lead("open", PY, ["CONTL"])], inp=False)
        self.r += 1
        self.para("[Describe the nature of the goods and services promised, when promises are typically satisfied, significant payment terms, "
                  "obligations for returns, refunds and warranties, and confirm that opening contract liabilities were recognised as revenue in the year.]")

    def n_otherinc(self):
        b = self.book
        self.title("otherinc", "Other income")
        self.heads(None, years=True)
        first = self.r
        for l in [x for x in leaves_of("OI") if x in b.active]:
            self.line(LEAF[l][1], ["-" + lead("close", CY, [l]), "-" + lead("close", PY, [l])])
        self.total("Total other income", first, self.r - 1)

    def n_expenses(self):
        b = self.book
        self.title("expenses", "Analysis of expenses by nature")
        self.para("Expenses are presented by function in the statement of comprehensive income. Their nature is as follows:")
        funcs = EXPENSE_FUNCTIONS
        self.heads([FACE[f][1] for f in funcs] + ["Total", "Total"])
        put(self.ws, f"B{self.r}", None)
        for i, f in enumerate([b.year(CY)] * 5 + [b.year(PY)]):
            put(self.ws, f"{chr(ord('C') + i)}{self.r}", f, F_BOLD, "0", RIGHT)
        self.r += 1
        first = self.r
        for code, caption in NATURES:
            if code == "OTHER":
                continue
            r = self.r
            self.line(caption, [nat(CY, [code], [f]) for f in funcs] + [f"SUM(C{r}:F{r})", nat(PY, [code], funcs)])
        r = self.r   # "other" = everything not given a specific nature, so the note always agrees to the statement
        self.line("Other expenses",
                  [f"{face('close', CY, f)}-SUM({c}{first}:{c}{r - 1})" for f, c in zip(funcs, "CDEF")]
                  + [f"SUM(C{r}:F{r})", "+".join(face("close", PY, f) for f in funcs) + f"-SUM(H{first}:H{r - 1})"])
        t = self.total("Total cost of sales, distribution, administrative and other expenses", first, r, ncols=6, key="EXP_T")
        b.reg("Notes", "EXP_FIRST", first)
        b.check(f"Note {b.notes['expenses']} CY: expenses by nature agree to the expenses by function in the SoCI",
                f"Notes!G{t}+" + "+".join(b.cell("SoCI", f, CY) for f in funcs))
        b.check(f"Note {b.notes['expenses']} PY: expenses by nature agree to the expenses by function in the SoCI",
                f"Notes!H{t}+" + "+".join(b.cell("SoCI", f, PY) for f in funcs))

    def n_finance(self):
        self.title("finance", "Finance income and finance costs")
        self.heads(None, years=True)
        self.line("Interest income on financial assets at amortised cost",
                  ["-" + lead("close", CY, ["FININC"]), "-" + lead("close", PY, ["FININC"])])
        self.line("Interest expense on financial liabilities at amortised cost",
                  [lead("close", CY, ["FINCOST"]), lead("close", PY, ["FINCOST"])])
        self.r += 1
        self.para("No borrowing costs are capitalised: all borrowing costs are recognised as an expense when incurred.")

    def n_pbt(self):
        b = self.book
        self.title("pbt", "Profit before tax")
        self.para("Profit before tax is stated after charging/(crediting):")
        self.heads(None, years=True)
        items = [("Cost of inventories recognised as an expense", "INVX"), ("Employee benefits expense", "EMP"),
                 ("Depreciation of property, plant and equipment", "DEP"), ("Amortisation of intangible assets", "AMORT"),
                 ("Operating lease expense", "RENT"), ("Impairment loss on trade receivables", "IMPAIR"),
                 ("Net foreign exchange loss", "FX")]
        for cap, n in items:
            self.line(cap, [nat(CY, [n]), nat(PY, [n])])
        self.line("(Gain)/loss on disposal of property, plant and equipment",
                  [lead("close", CY, ["OI.DISP"]), lead("close", PY, ["OI.DISP"])])
        self.r += 1
        self.line("Average number of employees during the year", [self.E["employees_cy"], self.E["employees_py"]], fmt="#,##0")
        self.line("Research and development expenditure recognised as an expense", [0, 0])

    def n_tax(self):
        b, E = self.book, self.E
        self.title("tax", "Income tax")
        self.heads(None, years=True)
        f = self.r
        self.line("Current tax expense", [lead("close", CY, ["TAX.CUR"]), lead("close", PY, ["TAX.CUR"])])
        self.line("Deferred tax expense - origination and reversal of temporary differences",
                  [lead("close", CY, ["TAX.DEF"]), lead("close", PY, ["TAX.DEF"])])
        t = self.total("Income tax expense recognised in profit or loss", f, f + 1, key="TAX_T")
        self.line("Deferred tax recognised in other comprehensive income (revaluation of land)",
                  ["-" + glm(CY, DT, ["OCI"]), "-" + glm(PY, DT, ["OCI"])])
        self.r += 1
        self.sub("Reconciliation of income tax expense to accounting profit")
        self.heads(None, years=True)
        p = self.line("Profit before tax", [b.cell("SoCI", "PBT", CY), b.cell("SoCI", "PBT", PY)])
        self.line("Applicable tax rate", [E["tax_rate"], E["tax_rate"]], fmt=PCT)
        a = self.line("Tax at the applicable rate", [f"C{p}*{E['tax_rate']}", f"D{p}*{E['tax_rate']}"])
        ni = self.cfg.get("notes_inputs", {})
        self.line("Tax effect of expenses not deductible for tax purposes",
                  [float(ni.get("tax_nondeductible_cy", 0)), float(ni.get("tax_nondeductible_py", 0))])
        self.line("Tax effect of income not taxable", [0, 0])
        o = self.line("Other differences (balancing figure - investigate if significant)",
                      [f"C{t}-SUM(C{a}:C{a + 2})", f"D{t}-SUM(D{a}:D{a + 2})"])
        self.total("Income tax expense", a, o)
        self.line("Effective tax rate", [f"IF(C{p}=0,0,C{t}/C{p})", f"IF(D{p}=0,0,D{t}/D{p})"], fmt=PCT)
        self.r += 1
        self.sub("Current tax balances")
        self.heads(None, years=True)
        self.line("Current tax liabilities/(assets), net", ["-" + lead("close", CY, ["CTL", "CTA"]), "-" + lead("close", PY, ["CTL", "CTA"])])
        for fy, col in ((CY, "C"), (PY, "D")):
            b.check(f"Note {b.notes['tax']} {fy}: income tax expense agrees to the statement of comprehensive income",
                    f"Notes!{col}{t}+{b.cell('SoCI', 'TAX', fy)}")

    def _rollforward(self, classes, cost, dep, charge_caption, dep_caption, key, reval=True):
        """Cost / accumulated depreciation reconciliation for the current year (comparatives not required)."""
        b = self.book
        n = len(classes)
        tot = chr(ord("C") + n)
        cols = [chr(ord("C") + i) for i in range(n)]
        self.heads([c for c in classes] + ["Total"])

        def row(caption, fn, bold=False, border=None, k=None):
            r = self.r
            return self.line(caption, [fn(i) for i in range(n)] + [f"SUM(C{r}:{cols[-1]}{r})"], bold=bold, border=border, key=k)

        yr0 = f'="At 1 January "&TEXT(YEAR({self.E["reporting_date"]}),"0")'
        yr1 = f'="At 31 December "&TEXT(YEAR({self.E["reporting_date"]}),"0")'
        self.sub("Cost or valuation")
        o = row(None, lambda i: lead("open", CY, [cost[i]]))
        put(self.ws, f"B{o}", yr0)
        row("Additions", lambda i: glm(CY, [cost[i]], ["ADD", "ADD_NC"]))
        row("Disposals", lambda i: glm(CY, [cost[i]], ["DISP"]))
        if reval:
            row("Revaluation recognised in other comprehensive income", lambda i: glm(CY, [cost[i]], ["REVAL"]))
        last = self.r - 1
        oth = row("Other movements (unclassified GL movements - investigate)",
                  lambda i: f"{lead('close', CY, [cost[i]])}-SUM({cols[i]}{o}:{cols[i]}{last})")
        cc = row(None, lambda i: f"SUM({cols[i]}{o}:{cols[i]}{oth})", bold=True, border=B_TOP)
        put(self.ws, f"B{cc}", yr1, F_BOLD)
        self.r += 1
        self.sub(dep_caption)
        o2 = row(None, lambda i: "-" + lead("open", CY, [dep[i]]))
        put(self.ws, f"B{o2}", yr0)
        ch = row(charge_caption, lambda i: "-" + glm(CY, [dep[i]], ["DEP"]), k=key + "_CHARGE")
        row("Eliminated on disposals", lambda i: "-" + glm(CY, [dep[i]], ["DISP"]))
        last = self.r - 1
        oth2 = row("Other movements (unclassified GL movements - investigate)",
                   lambda i: f"-{lead('close', CY, [dep[i]])}-SUM({cols[i]}{o2}:{cols[i]}{last})")
        cd = row(None, lambda i: f"SUM({cols[i]}{o2}:{cols[i]}{oth2})", bold=True, border=B_TOP)
        put(self.ws, f"B{cd}", yr1, F_BOLD)
        self.r += 1
        self.sub("Carrying amount")
        c1 = row(None, lambda i: f"{cols[i]}{cc}-{cols[i]}{cd}", bold=True, border=B_TOTAL, k=key + "_NBV")
        put(self.ws, f"B{c1}", yr1, F_BOLD)
        c0 = row(None, lambda i: f"{cols[i]}{o}-{cols[i]}{o2}", bold=True)
        put(self.ws, f"B{c0}", yr0, F_BOLD)
        self.r += 1
        b.check(f"Note {b.notes[key]}: unclassified movements in the roll-forward are nil",
                f"ABS(Notes!{tot}{oth})+ABS(Notes!{tot}{oth2})")
        return tot, c1, c0, ch

    def n_ppe(self):
        b = self.book
        self.title("ppe", "Property, plant and equipment")
        tot, c1, c0, ch = self._rollforward([c for _, c in PPE_CLASSES], PPE_COST, PPE_DEP, "Depreciation charge for the year",
                                            "Accumulated depreciation and impairment", "ppe")
        b.check(f"Note {b.notes['ppe']}: carrying amount agrees to the SoFP (CY and PY)",
                f"ABS(Notes!{tot}{c1}-{b.cell('SoFP', 'PPE', CY)})+ABS(Notes!{tot}{c0}-{b.cell('SoFP', 'PPE', PY)})")
        b.check(f"Note {b.notes['ppe']}: depreciation charge agrees to the expense in profit or loss",
                f"Notes!{tot}{ch}-{nat(CY, ['DEP'])}")
        self.para("[Land was revalued on (date) by (name), an independent valuer, by reference to (method and significant assumptions). "
                  "Had land been carried under the cost model its carrying amount would have been (amount). State the carrying amount of "
                  "any assets pledged as security for liabilities and any contractual commitments to acquire property, plant and equipment.]")

    def n_intang(self):
        b = self.book
        self.title("intang", "Intangible assets")
        tot, c1, c0, ch = self._rollforward(["Computer software"], ["INTANG.C"], ["INTANG.D"], "Amortisation charge for the year",
                                            "Accumulated amortisation and impairment", "intang", reval=False)
        b.check(f"Note {b.notes['intang']}: carrying amount agrees to the SoFP (CY and PY)",
                f"ABS(Notes!{tot}{c1}-{b.cell('SoFP', 'INTANG', CY)})+ABS(Notes!{tot}{c0}-{b.cell('SoFP', 'INTANG', PY)})")
        self.para("Amortisation is included in administrative expenses in the statement of comprehensive income.", inp=True)

    def n_invprop(self):
        self.title("invprop", "Investment property")
        self.leaf_table(["INVPROP"], "", "Investment property", "invprop")
        self.para("[State the measurement basis (fair value through profit or loss, or cost model where fair value cannot be measured without "
                  "undue cost or effort), the methods and significant assumptions applied, whether an independent valuer was used, and "
                  "reconcile the opening and closing carrying amounts.]")

    def n_invassoc(self):
        self.title("invassoc", "Investments in associates and jointly controlled entities")
        self.leaf_table(["INVASSOC"], "", "Investments in associates and jointly controlled entities", "invassoc")
        self.para("[State the accounting policy (cost, equity or fair value model), the carrying amounts, dividends recognised as income and "
                  "the entity's share of profit or loss.]")

    def n_ofa(self):
        self.title("ofa", "Other financial assets")
        self.leaf_table(["OFA_NC", "OFA_C"], "", "Other financial assets", "ofa")
        self.para("[Describe the instruments, their measurement basis, terms and conditions.]")

    def n_inv(self):
        self.title("inv", "Inventories")
        self.leaf_table(["INV"], "", "Total inventories", "inv")
        self.heads(None, years=True)
        self.line("Inventories recognised as an expense during the year", [nat(CY, ["INVX"]), nat(PY, ["INVX"])])
        self.line("Write-downs of inventories recognised/(reversed) in profit or loss", [0, 0])
        self.line("Carrying amount of inventories pledged as security for liabilities", [0, 0])
        self.r += 1

    def n_tor(self):
        self.title("tor", "Trade and other receivables")
        self.leaf_table(["TOR"], "", "Total trade and other receivables", "tor")
        self.sub("Movement in the allowance for impairment of trade receivables")
        self.heads(None, years=True)
        f = self.r
        self.line("At 1 January", ["-" + lead("open", CY, ["TOR.ALLOW"]), "-" + lead("open", PY, ["TOR.ALLOW"])])
        self.line("Impairment loss recognised in profit or loss", [nat(CY, ["IMPAIR"]), nat(PY, ["IMPAIR"])])
        self.line("Amounts written off as uncollectible and other movements",
                  [f"-{lead('close', CY, ['TOR.ALLOW'])}-SUM(C{f}:C{f + 1})", f"-{lead('close', PY, ['TOR.ALLOW'])}-SUM(D{f}:D{f + 1})"])
        self.total("At 31 December", f, f + 2)
        self.r += 1

    def n_cash(self):
        b = self.book
        self.title("cash", "Cash and cash equivalents")
        self.heads(None, years=True)
        f = self.r
        for l in [x for x in leaves_of("CASH") if x in b.active]:
            self.line(LEAF[l][1], [lead("close", CY, [l]), lead("close", PY, [l])])
        t = self.total("Cash and cash equivalents in the statement of financial position", f, self.r - 1, double=False)
        od = self.line("Bank overdrafts repayable on demand (included in current borrowings)",
                       [lead("close", CY, ["BORR_C.OD"]), lead("close", PY, ["BORR_C.OD"])])
        c = self.line("Cash and cash equivalents in the statement of cash flows", [f"C{t}+C{od}", f"D{t}+D{od}"], bold=True,
                      border=B_TOTAL)
        for fy, col in ((CY, "C"), (PY, "D")):
            b.check(f"Note {b.notes['cash']} {fy}: cash and cash equivalents agree to the SoFP and to the SoCF",
                    f"ABS(Notes!{col}{t}-{b.cell('SoFP', 'CASH', fy)})+ABS(Notes!{col}{c}-{b.cell('SoCF', 'CLOSE', fy)})")
        self.r += 1
        self.para("[Disclose any significant cash and cash equivalent balances that are not available for use by the Company, and any "
                  "significant non-cash investing and financing transactions.]")

    def n_equity(self):
        E = self.E
        self.title("equity", "Share capital and reserves")
        self.heads(None, years=True)
        self.line("Number of ordinary shares issued and fully paid", [E["shares_issued_cy"], E["shares_issued_py"]], fmt="#,##0")
        self.line("Par value per share", [E["par_value"], E["par_value"]], fmt="#,##0.00")
        self.line("Share capital", ["-" + lead("close", CY, ["SC"]), "-" + lead("close", PY, ["SC"])])
        self.line("Share premium", ["-" + lead("close", CY, ["SP"]), "-" + lead("close", PY, ["SP"])])
        self.line("Number of shares issued during the year", [f"{E['shares_issued_cy']}-{E['shares_issued_py']}", None], fmt="#,##0")
        self.r += 1
        self.para("All shares are ordinary shares carrying one vote each and equal rights to dividends and to the repayment of capital. "
                  "[Describe rights, preferences and restrictions, shares reserved for issue and any shares held by the entity.]", inp=True, placeholder=False)
        self.sub("Reserves")
        self.para("Share premium records the amount received in excess of the par value of shares issued. The revaluation reserve records "
                  "revaluation increases on land, net of the related deferred tax; it is not available for distribution until realised.", inp=True, placeholder=False)

    def n_borr(self):
        b = self.book
        self.title("borr", "Borrowings")
        self.leaf_table(["BORR_NC", "BORR_C"], "-", "Total borrowings", "borr")
        self.para("[Describe the terms and conditions: interest rate, maturity, repayment schedule, restrictions and covenants; the carrying "
                  "amount of assets pledged as security; and any defaults or breaches of loan terms during the year.]")
        self.sub("Reconciliation of liabilities arising from financing activities")
        self.heads(None, years=True)
        f = self.r
        self.line("At 1 January", ["-" + lead("open", CY, BORR), "-" + lead("open", PY, BORR)])
        self.line("Cash flows: proceeds from borrowings", ["-" + glm(CY, BORR, ["PROCEEDS"]), "-" + glm(PY, BORR, ["PROCEEDS"])])
        self.line("Cash flows: repayment of borrowings", ["-" + glm(CY, BORR, ["REPAY"]), "-" + glm(PY, BORR, ["REPAY"])])
        self.line("Non-cash changes (new finance leases and other)",
                  [f"-{lead('close', CY, BORR)}-SUM(C{f}:C{f + 2})", f"-{lead('close', PY, BORR)}-SUM(D{f}:D{f + 2})"])
        self.total("At 31 December (excluding bank overdrafts)", f, f + 3)
        self.r += 1

    def n_top(self):
        self.title("top", "Trade and other payables")
        self.leaf_table(["TOP"], "-", "Total trade and other payables", "top")
        self.para("[Supplier finance arrangements: the Company has not entered into any supplier finance arrangement. If it has, disclose the "
                  "terms and conditions, the carrying amount of the liabilities concerned, the amount already paid to suppliers by the finance "
                  "provider and the range of payment due dates.]")

    def n_prov(self):
        b = self.book
        P = ["PROV_NC", "PROV_C"]
        self.title("prov", "Provisions")
        self.heads(["Warranties"])
        f = self.r
        self.line("At 1 January", ["-" + lead("open", CY, P)])
        self.line("Additional provisions recognised in the year", ["-" + glm(CY, P, ["ADD"])])
        self.line("Amounts used during the year", ["-" + glm(CY, P, ["USED"])])
        self.line("Unused amounts reversed", ["-" + glm(CY, P, ["REVERSED"])])
        o = self.line("Other movements (unclassified GL movements - investigate)", [f"-{lead('close', CY, P)}-SUM(C{f}:C{f + 3})"])
        t = self.total("At 31 December", f, o, ncols=1)
        self.line("Of which: current", ["-" + lead("close", CY, ["PROV_C"])])
        self.line("Of which: non-current", ["-" + lead("close", CY, ["PROV_NC"])])
        self.r += 1
        b.check(f"Note {b.notes['prov']}: unclassified movements in provisions are nil", f"Notes!C{o}")
        self.para("[Describe the nature of each obligation, the expected amount and timing of payments, and the uncertainties about them.]")

    def n_ebo(self):
        self.title("ebo", "Employee benefit obligations")
        self.leaf_table(["EBO_NC"], "-", "Employee benefit obligations", "ebo")
        self.para("[Describe the plans (defined contribution and/or defined benefit), amounts recognised as expense, the funding status and "
                  "the principal actuarial assumptions.]")

    def n_dtax(self):
        b = self.book
        self.title("dtax", "Deferred tax")
        self.para("Deferred tax liabilities/(assets) arise from accelerated tax depreciation and the revaluation of land. Movements were:", inp=True, placeholder=False)
        self.heads(None, years=True)
        f = self.r
        self.line("At 1 January", ["-" + lead("open", CY, DT), "-" + lead("open", PY, DT)])
        self.line("Charged/(credited) to profit or loss", [lead("close", CY, ["TAX.DEF"]), lead("close", PY, ["TAX.DEF"])])
        self.line("Charged/(credited) to other comprehensive income", ["-" + glm(CY, DT, ["OCI"]), "-" + glm(PY, DT, ["OCI"])])
        t = self.total("At 31 December - net deferred tax liability/(asset)", f, f + 2)
        for fy, col in ((CY, "C"), (PY, "D")):
            b.check(f"Note {b.notes['dtax']} {fy}: deferred tax roll-forward agrees to the SoFP",
                    f"Notes!{col}{t}+{lead('close', fy, DT)}")
        self.r += 1

    def n_fininst(self):
        self.title("fininst", "Financial instruments - carrying amounts by category")
        self.heads(None, years=True)
        fa = ["TOR.TRADE", "TOR.ALLOW", "TOR.OTHER", "TOR.INT", "CASH.HAND", "CASH.BANK", "CASH.DEP", "OFA_NC", "OFA_C"]
        fl = ["TOP.TRADE", "TOP.ACCR", "TOP.INT", "TOP.DIV"] + BORR + ["BORR_C.OD"]
        self.line("Financial assets measured at amortised cost", [lead("close", CY, fa), lead("close", PY, fa)])
        self.line("Financial liabilities measured at amortised cost", ["-" + lead("close", CY, fl), "-" + lead("close", PY, fl)])
        self.r += 1
        self.para("The Company holds no financial instruments measured at fair value through profit or loss. Land is the only item measured "
                  "at fair value; the valuation technique and inputs are described in the property, plant and equipment note. "
                  "[Add fair value measurement disclosures required by Section 12 for any other item measured at fair value.]", inp=True, placeholder=False)

    def n_dividends(self):
        E = self.E
        self.title("dividends", "Dividends")
        self.heads(None, years=True)
        d = self.line("Dividends declared during the year", [f"{lead('close', CY, ['RE.DIV'])}+{glm(CY, ['RE'], ['DIVIDEND'])}",
                                                             f"{lead('close', PY, ['RE.DIV'])}+{glm(PY, ['RE'], ['DIVIDEND'])}"])
        self.line("Dividend per share", [f"IF({E['shares_issued_cy']}=0,0,C{d}/{E['shares_issued_cy']})",
                                         f"IF({E['shares_issued_py']}=0,0,D{d}/{E['shares_issued_py']})"], fmt="#,##0.000")
        self.line("Dividends payable at year end", ["-" + lead("close", CY, ["TOP.DIV"]), "-" + lead("close", PY, ["TOP.DIV"])])
        self.r += 1

    def n_commitments(self):
        self.title("commitments", "Commitments")
        self.sub("Operating leases - the Company as lessee")
        self.para("The Company rents its office premises under a non-cancellable operating lease. Future minimum lease payments are:", inp=True, placeholder=False)
        self.heads(None, years=True)
        f = self.r
        for cap in ("Not later than one year", "Later than one year and not later than five years", "Later than five years"):
            c = self.line(cap, [0, 0])
            for col in "CD":
                self.ws[f"{col}{c}"].fill = FILL_INPUT
        self.total("Total future minimum lease payments", f, f + 2)
        self.line("Lease payments recognised as an expense in the year", [nat(CY, ["RENT"]), nat(PY, ["RENT"])])
        self.r += 1
        self.sub("Capital commitments")
        self.para("[State contractual commitments for the acquisition of property, plant and equipment and intangible assets, or that there are none.]")

    def n_contingencies(self):
        self.title("contingencies", "Contingent liabilities and contingent assets")
        self.para("[Describe the nature of each class of contingent liability, an estimate of its financial effect, the uncertainties about the "
                  "amount or timing of any outflow and the possibility of reimbursement - or state that there are none.]")

    def n_related(self):
        self.title("related", "Related party disclosures")
        self.para("[Name the parent and the ultimate controlling party. For each category of related party disclose the nature of the "
                  "relationship, the transactions during the year, outstanding balances, their terms, and any impairment recognised.]")
        self.heads(None, years=True)
        c = self.line("Key management personnel compensation (total)", [0, 0])
        for col in "CD":
            self.ws[f"{col}{c}"].fill = FILL_INPUT
        self.r += 1

    def n_events(self):
        self.title("events", "Events after the end of the reporting period")
        self.para("[Describe the nature and estimated financial effect of each material non-adjusting event after the reporting date - or state "
                  "that there were none between the reporting date and the date of authorisation.]")

    def n_approval(self):
        E = self.E
        self.title("approval", "Approval of the financial statements")
        self.para(f'="These financial statements were approved and authorised for issue by "&{E["authorised_by"]}&" on "'
                  f'&TEXT({E["authorised_date"]},"d mmmm yyyy")&"."', placeholder=False)
