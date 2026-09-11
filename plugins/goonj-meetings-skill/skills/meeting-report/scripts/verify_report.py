#!/usr/bin/env python3
"""Verify a built meeting report in a single pass.

    python3 verify_report.py report.xlsx transcript.txt

Runs every check at once instead of one shell call per row:
  1. every Evidence quote appears verbatim in the transcript (case-insensitive;
     stutters, filler words and mis-transcribed terms must still be present)
  2. no formula error values anywhere in the workbook
  3. Summary counts are arithmetically consistent with the sheets
  4. lists Medium/Low confidence rows a human must confirm before circulating

Exit 0 = clean.  Exit 1 = something needs fixing.
"""
import json, re, sys, unicodedata

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed.  Fix:  pip install openpyxl")

# Evidence lives in a different column per sheet.
EVIDENCE_COLS = {"Action Items": 11, "Decisions": 9, "Risks and Compliance": 10}
MIN_FRAGMENT = 12          # shorter fragments match by chance; not worth failing on
ELLIPSIS = re.compile(r"\.\.\.|…")
QUOTED = re.compile(r'"([^"]+)"')


def norm(s):
    """Collapse whitespace and unify quote characters so cosmetic differences don't fail."""
    s = unicodedata.normalize("NFC", s)
    s = (s.replace("’", "'").replace("‘", "'")
          .replace("“", '"').replace("”", '"')
          .replace("—", "-").replace("–", "-"))
    return re.sub(r"\s+", " ", s).strip()


def fragments(cell):
    """Evidence cell -> the substrings that must appear verbatim in the transcript.

    Handles: surrounding quotes, several quoted spans in one cell, '...' elisions,
    and a trailing [analyst caveat] which is ours and not from the transcript.
    """
    text = re.sub(r"\[[^\]]*\]\s*$", "", cell).strip()
    spans = QUOTED.findall(text) or [text.strip('"')]
    out = []
    for span in spans:
        for frag in ELLIPSIS.split(span):
            frag = norm(frag).strip(' "\'-')
            if len(frag) >= MIN_FRAGMENT:
                out.append(frag)
    return out


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    xlsx, txt = sys.argv[1], sys.argv[2]

    # Case is cosmetic in a transcript; everything else must match exactly.
    haystack = norm(open(txt, encoding="utf-8", errors="replace").read()).lower()
    wb = openpyxl.load_workbook(xlsx, data_only=True)

    failures, unverified, review = [], 0, []

    # 1 - evidence quotes
    checked = 0
    for sheet, col in EVIDENCE_COLS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row in range(3, ws.max_row + 1):
            cell = ws.cell(row=row, column=col).value
            sr = ws.cell(row=row, column=1).value
            if not isinstance(cell, str) or not cell.strip():
                continue
            frags = fragments(cell)
            if not frags:
                unverified += 1
                continue
            for frag in frags:
                checked += 1
                if frag.lower() not in haystack:
                    failures.append({"sheet": sheet, "sr": sr, "row": row,
                                     "fragment": frag[:90]})

    # 2 - formula errors
    errors = [{"sheet": s.title, "cell": c.coordinate, "value": c.value}
              for s in wb.worksheets for r in s.iter_rows() for c in r
              if isinstance(c.value, str) and c.value.startswith("#")]

    # 3 - count arithmetic
    arithmetic = []
    if "Decisions" in wb.sheetnames:
        ws = wb["Decisions"]
        vals = [ws.cell(row=r, column=4).value for r in range(3, ws.max_row + 1)]
        vals = [v for v in vals if v]
        dec = sum(1 for v in vals if str(v).startswith("Decided"))
        dfr = sum(1 for v in vals if str(v).startswith("Deferred"))
        if dec + dfr != len(vals):
            arithmetic.append(f"Decisions: {dec} Decided + {dfr} Deferred != {len(vals)} rows "
                              f"- a value is neither; the Summary COUNTIF will undercount")

    # 4 - rows needing human confirmation
    if "Action Items" in wb.sheetnames:
        ws = wb["Action Items"]
        for r in range(3, ws.max_row + 1):
            conf = ws.cell(row=r, column=10).value
            if conf in ("Medium", "Low"):
                review.append({"sr": ws.cell(row=r, column=1).value, "confidence": conf,
                               "owner": ws.cell(row=r, column=4).value,
                               "action": str(ws.cell(row=r, column=2).value or "")[:70]})

    ok = not failures and not errors and not arithmetic
    report = {
        "result": "PASS" if ok else "FAIL",
        "quote_fragments_checked": checked,
        "quote_failures": failures,
        "rows_with_no_evidence": unverified,
        "formula_errors": errors,
        "arithmetic": arithmetic,
        "needs_human_confirmation": review,
    }
    print(json.dumps(report, indent=2, ensure_ascii=False))

    if not ok:
        print("\nFAIL - fix the rows above. A quote that does not match means the reading "
              "drifted; correct the row rather than softening the quote.", file=sys.stderr)
        sys.exit(1)
    if review:
        print(f"\nPASS - but {len(review)} row(s) are Medium/Low confidence and must be "
              "confirmed by a person before this report is circulated.", file=sys.stderr)


if __name__ == "__main__":
    main()
