"""Classify GL lines on roll-forward accounts into movement types.

The type is inferred from the other accounts in the same journal (e.g. a debit
to a PPE cost account against cash is an addition; against the revaluation
reserve it is a revaluation). The result is written to the GL sheet as an
editable input, so a reviewer can override any line.
"""
from .taxonomy import LEAF, is_pl


def classify(gl, leaf_of):
    """gl: DataFrame with journal, account, debit, credit. Returns list of movement types (one per line)."""
    mclass = gl["account"].map(lambda a: LEAF.get(leaf_of.get(a, ""), ("", "", "", ""))[3])
    pl = gl["account"].map(lambda a: is_pl(leaf_of.get(a, "")))
    out = [""] * len(gl)
    pos = {idx: n for n, idx in enumerate(gl.index)}
    for _, idxs in gl.groupby("journal", sort=False).groups.items():
        classes = [mclass[i] for i in idxs]
        present = set(classes)
        has_pl = any(pl[i] for i in idxs)
        for i in idxs:
            c = mclass[i]
            if not c or c == "CASH":
                continue
            debit = gl.at[i, "debit"] - gl.at[i, "credit"] > 0
            others = [mclass[k] for k in idxs if k != i]
            if c in ("PPE_C", "INT_C"):
                if "RESV" in present:
                    t = "REVAL"
                elif debit:
                    t = "ADD_NC" if "BORR" in present and "CASH" not in present else "ADD"
                else:
                    t = "DISP"
            elif c in ("PPE_D", "INT_D"):
                t = "DISP" if debit else "DEP"
            elif c == "BORR":
                if others and all(o == "BORR" for o in others):
                    t = "RECLASS"
                elif "CASH" in present:
                    t = "REPAY" if debit else "PROCEEDS"
                else:
                    t = "NONCASH"
            elif c == "DT":
                t = "OCI" if "RESV" in present else "PL"
            elif c == "RESV":
                t = "TRANSFER" if "RE" in present else "OCI"
            elif c == "RE":
                t = "TRANSFER" if "RESV" in present else ("DIVIDEND" if debit else "OTHER")
            elif c == "PROV":
                t = ("REVERSED" if has_pl else "USED") if debit else "ADD"
            elif c == "EQ":
                t = "ISSUE"
            else:
                t = ""
            out[pos[i]] = t
    return out
