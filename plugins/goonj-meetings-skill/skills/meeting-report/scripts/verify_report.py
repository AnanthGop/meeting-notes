#!/usr/bin/env python3
"""Verify a built meeting report in a single pass.

    python3 verify_report.py report.xlsx transcript.txt

Runs every check at once instead of one shell call per row:
  1. every Evidence quote appears verbatim in the transcript (case-insensitive;
     stutters, filler words and mis-transcribed terms must still be present), and
     no content row is missing its evidence
  2. no formula error values anywhere in the workbook
  3. the Summary counts match a recount of the sheets (when LibreOffice has
     recalculated the file; otherwise reported as "not recalculated")
  4. the content rules SKILL.md states are actually met: every decision is Decided
     or Deferred, a deferred decision names who it waits on, an open question names
     who answers it, dropdown columns hold only their allowed values, and every
     carry-forward row carries a recognised status
  5. lists Medium/Low confidence rows a human must confirm before circulating

Exit 0 = clean.  Exit 1 = something needs fixing.
"""
import json, os, re, sys, unicodedata

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed.  Fix:  pip install openpyxl")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_report import CARRY_STATUSES  # single source of truth for the vocabulary

MIN_FRAGMENT = 12          # shorter fragments match by chance; not worth checking
ELLIPSIS = re.compile(r"\.\.\.|…")
QUOTED = re.compile(r'"([^"]+)"')
CAVEAT = re.compile(r"\[[^\]]*\]\s*$")

EVIDENCE = "Evidence from transcript"
EVIDENCE_SHEETS = ["Action Items", "Decisions", "Open Questions", "Risks and Compliance",
                   "Carry Forward"]

# Allowed values per (sheet, English header). Mirrors the dropdowns build_report.py writes;
# a value outside the list is a silent miscount in the Summary, not a cosmetic slip.
ALLOWED = {
    ("Action Items", "Due basis"): ["Explicit", "Relational", "Conditional", "None"],
    ("Action Items", "Priority"): ["High", "Medium", "Low"],
    ("Action Items", "Status"): ["Open", "In Progress", "Closed", "Superseded", "Dropped"],
    ("Action Items", "Confidence"): ["High", "Medium", "Low"],
    ("Open Questions", "Status"): ["Open", "Answered", "Escalated", "Dropped"],
    ("Risks and Compliance", "Severity"): ["High", "Medium", "Low"],
    ("Carry Forward", "Status this meeting"): CARRY_STATUSES,
}


def s(v):
    return "" if v is None else str(v).strip()


def norm(text):
    """Collapse whitespace and unify quote characters so cosmetic differences don't fail."""
    text = unicodedata.normalize("NFC", text)
    text = (text.replace("’", "'").replace("‘", "'")
                .replace("“", '"').replace("”", '"')
                .replace("—", "-").replace("–", "-"))
    return re.sub(r"\s+", " ", text).strip()


def fragments(cell):
    """Evidence cell -> the substrings that must appear verbatim in the transcript.

    Handles: surrounding quotes, several quoted spans in one cell, '...' elisions,
    and a trailing [analyst caveat] which is ours and not from the transcript.
    """
    text = CAVEAT.sub("", cell).strip()
    spans = QUOTED.findall(text) or [text.strip('"')]
    out = []
    for span in spans:
        for frag in ELLIPSIS.split(span):
            frag = norm(frag).strip(' "\'-')
            if len(frag) >= MIN_FRAGMENT:
                out.append(frag)
    return out


def header_map(ws):
    """Column index by English header text (row 2 holds 'English\\nHindi')."""
    out = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=2, column=c).value
        if isinstance(v, str):
            out[v.split("\n")[0].strip()] = c
    return out


def rows_of(wb, sheet):
    """Content rows of a sheet as dicts keyed by English header. Blank rows are skipped."""
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    hm = header_map(ws)
    rows = []
    for r in range(3, ws.max_row + 1):
        row = {h: ws.cell(row=r, column=c).value for h, c in hm.items()}
        if any(s(v) for v in row.values()):
            row["_row"] = r
            rows.append(row)
    return rows


def recount(sheets):
    """Recompute every Summary stat the way its formula does. Keys are the English labels."""
    A, D, Q, R, C = (sheets[n] for n in ("Action Items", "Decisions", "Open Questions",
                                          "Risks and Compliance", "Carry Forward"))
    eq = lambda v, t: s(v).lower() == t.lower()
    sw = lambda v, t: s(v).lower().startswith(t.lower())
    return {
        "Action items — total": sum(1 for r in A if s(r.get("Sr"))),
        "Action items — open": sum(1 for r in A if eq(r.get("Status"), "Open")),
        "Action items — high priority": sum(1 for r in A if eq(r.get("Priority"), "High")
                                            and eq(r.get("Status"), "Open")),
        "Action items — no due date agreed": sum(1 for r in A if eq(r.get("Due basis"), "None")),
        "Decisions taken": sum(1 for r in D if sw(r.get("Decided / Deferred"), "Decided")),
        "Decisions deferred": sum(1 for r in D if sw(r.get("Decided / Deferred"), "Deferred")),
        "Open questions": sum(1 for r in Q if eq(r.get("Status"), "Open")),
        "High severity risks": sum(1 for r in R if eq(r.get("Severity"), "High")),
        "Items carried forward still open": sum(
            1 for r in C if sw(r.get("Status this meeting"), "Still open")
            or eq(r.get("Status this meeting"), "Not mentioned")),
    }


def summary_check(wb, expected):
    """Compare the Summary sheet's cached values with a recount.

    Returns "OK", "not recalculated" (no cached values - LibreOffice did not run), or a list
    of mismatches.
    """
    if "Summary" not in wb.sheetnames:
        return [{"stat": "(Summary sheet)", "problem": "sheet missing"}]
    ws = wb["Summary"]
    cached = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=2).value
        if isinstance(label, str) and " / " in label:
            en = label.split(" / ")[0].strip()
            if en in expected:
                cached[en] = ws.cell(row=r, column=3).value
    if not cached:
        return [{"stat": "(all)", "problem": "no stat rows found on the Summary sheet"}]
    if any(v is None for v in cached.values()):
        return "not recalculated"
    mismatches = []
    for en, v in cached.items():
        try:
            same = float(v) == float(expected[en])
        except (TypeError, ValueError):
            same = False
        if not same:
            mismatches.append({"stat": en, "summary_shows": v, "recount": expected[en]})
    return mismatches or "OK"


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    xlsx, txt = sys.argv[1], sys.argv[2]

    # Case is cosmetic in a transcript; everything else must match exactly.
    haystack = norm(open(txt, encoding="utf-8", errors="replace").read()).lower()
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    sheets = {n: rows_of(wb, n) for n in EVIDENCE_SHEETS}

    def ref(sheet, row):
        return {"sheet": sheet, "sr": row.get("Sr"), "row": row["_row"]}

    # 1 - evidence: present on every content row, and every quoted fragment in the transcript
    failures, no_evidence, caveat_only, checked = [], [], [], 0
    for sheet in EVIDENCE_SHEETS:
        for row in sheets[sheet]:
            cell = s(row.get(EVIDENCE))
            # Carry Forward only has new transcript words to quote when something was said
            # about the item; SKILL.md requires the quote for a Closed.
            required = (sheet != "Carry Forward"
                        or s(row.get("Status this meeting")).lower().startswith("closed"))
            if not cell:
                if required:
                    no_evidence.append({**ref(sheet, row), "reason": "Evidence cell is empty"})
                continue
            frags = fragments(cell)
            if not frags:
                if cell.startswith("[") and not CAVEAT.sub("", cell).strip():
                    caveat_only.append({**ref(sheet, row), "evidence": cell[:90]})
                else:
                    no_evidence.append({**ref(sheet, row), "reason":
                        f"no quoted fragment of {MIN_FRAGMENT}+ characters - quote more of the turn"})
                continue
            for frag in frags:
                checked += 1
                if frag.lower() not in haystack:
                    failures.append({**ref(sheet, row), "fragment": frag[:90]})

    # 2 - formula errors
    errors = [{"sheet": ws.title, "cell": c.coordinate, "value": c.value}
              for ws in wb.worksheets for r in ws.iter_rows() for c in r
              if isinstance(c.value, str) and c.value.startswith("#")]

    # 3 - Summary counts against a recount
    summary = summary_check(wb, recount(sheets))

    # 4 - content rules from SKILL.md
    rules, warnings = [], []
    for row in sheets["Decisions"]:
        t = s(row.get("Decided / Deferred"))
        if not (t.startswith("Decided") or t.startswith("Deferred")):
            rules.append({**ref("Decisions", row), "problem":
                f"'Decided / Deferred' is {t!r}; it must start with Decided or Deferred "
                "or the Summary COUNTIF misses it"})
        elif t.startswith("Deferred") and not s(row.get("If deferred — pending on")):
            rules.append({**ref("Decisions", row), "problem":
                "deferred decision names nobody it now waits on - that column is the point"})
    for row in sheets["Open Questions"]:
        if not s(row.get("Must be answered by")):
            rules.append({**ref("Open Questions", row), "problem":
                "no named person owes the answer - escalate it in the report instead of listing it"})
    for row in sheets["Action Items"]:
        basis, due = s(row.get("Due basis")), s(row.get("Due"))
        if basis == "Explicit" and not due:
            warnings.append({**ref("Action Items", row), "problem": "Due basis Explicit but Due is empty"})
        if basis == "None" and due:
            warnings.append({**ref("Action Items", row), "problem":
                f"Due basis None but Due says {due!r} - one of them is wrong"})
    for (sheet, col), allowed in ALLOWED.items():
        for row in sheets[sheet]:
            v = s(row.get(col))
            if v not in allowed:
                rules.append({**ref(sheet, row), "problem":
                    f"{col!r} is {v!r}; must be one of {', '.join(allowed)}"})

    # 5 - rows needing human confirmation
    review = [{"sr": row.get("Sr"), "confidence": s(row.get("Confidence")),
               "owner": row.get("Owner"), "action": s(row.get("Action"))[:70]}
              for row in sheets["Action Items"] if s(row.get("Confidence")) in ("Medium", "Low")]

    ok = not (failures or no_evidence or errors or rules or isinstance(summary, list))
    report = {
        "result": "PASS" if ok else "FAIL",
        "quote_fragments_checked": checked,
        "quote_failures": failures,
        "rows_with_no_evidence": no_evidence,
        "caveat_only_evidence": caveat_only,
        "formula_errors": errors,
        "summary_counts": summary,
        "content_rules": rules,
        "warnings": warnings,
        "needs_human_confirmation": review,
    }
    print(json.dumps(report, indent=2, ensure_ascii=False))

    if not ok:
        hints = []
        if failures:
            hints.append("A quote that does not match means the reading drifted; correct the "
                         "row rather than softening the quote.")
        if no_evidence:
            hints.append("A row without transcript evidence does not go in the report: quote "
                         "the words or drop the row.")
        if rules:
            hints.append("content_rules are SKILL.md rules the workbook breaks; fix the content JSON.")
        if isinstance(summary, list):
            hints.append("The Summary counts disagree with the sheets - rebuild before delivering.")
        print("\nFAIL - " + " ".join(hints), file=sys.stderr)
        sys.exit(1)
    notes = []
    if review:
        notes.append(f"{len(review)} row(s) are Medium/Low confidence and must be confirmed "
                     "by a person before this report is circulated")
    if caveat_only:
        notes.append(f"{len(caveat_only)} row(s) cite only an [unreadable] note - name them in the delivery")
    if summary == "not recalculated":
        notes.append("Summary counts were not recalculated (no LibreOffice); Excel will compute them on open")
    if notes:
        print("\nPASS - but " + "; ".join(notes) + ".", file=sys.stderr)


if __name__ == "__main__":
    main()
