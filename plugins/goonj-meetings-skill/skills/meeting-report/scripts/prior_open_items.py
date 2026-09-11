#!/usr/bin/env python3
"""List every still-open item from the previous report in a workstream, for carry-forward.

    python3 prior_open_items.py <folder> <workstream> [--before YYYY-MM-DD]

Finds the newest report for the workstream in <folder>, <folder>/<workstream> or
<folder>/../<workstream>: any file whose name contains Meeting_Report_<Workstream>_YYYY-MM-DD,
whatever prefix or suffix surrounds it (Goonj_Meeting_Report_Leadership_2026-09-07 v2.xlsx
qualifies). Newest by the date in the filename, never by modification time; among files with
the same date, one that has the content sheets beats a Summary-only copy, then the plain name
beats a suffixed one. Prints JSON of the items still open in it:

  Action Items     Status is Open or In Progress
  Open Questions   Status is Open or Escalated
  Decisions        Decided / Deferred starts with Deferred
  Carry Forward    Status this meeting is Still open*, Not mentioned, or blank

Every item carries `first_raised`, the cross-meeting reference: "<date> · Action 3",
"<date> · Question 2", "<date> · Decision 1" for the report's own rows, and passed through
unchanged for rows that were already on its Carry Forward sheet, so the original reference
survives the whole chain. Copy it into the new Carry Forward row verbatim.

--before excludes reports dated on or after the meeting being written up, so re-running a
meeting never carries forward from its own report. Reads columns by header name, with aliases
for the headers earlier versions of the skill used, so older reports work too.
"""
import argparse, glob, json, os, re, sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed.  Fix:  pip install openpyxl")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_report import rows_of, s

CONTENT_SHEETS = ("Action Items", "Decisions", "Open Questions", "Carry Forward")

# Header names used by reports from earlier versions of the skill, per current header.
ALIASES = {
    "Sr": ("Sr", "Ref (original)"),
    "Open question": ("Open question", "Question"),
    "Open question (Hindi)": ("Open question (Hindi)", "Question (Hindi)"),
    "Must be answered by": ("Must be answered by", "Owes the answer"),
    "What it blocks / impact": ("What it blocks / impact", "What it blocks"),
    "By whom": ("By whom", "Decided by"),
    "If deferred — pending on": ("If deferred — pending on", "Waits on (if deferred)"),
    "Item carried forward": ("Item carried forward", "Item"),
    "Note": ("Note", "Update"),
}


def g(row, key):
    """Value under the current header, or under whichever older alias the row has."""
    for k in ALIASES.get(key, (key,)):
        if s(row.get(k)):
            return row[k]
    return None


def has_content_sheets(path):
    try:
        names = openpyxl.load_workbook(path, read_only=True).sheetnames
    except Exception:
        return False
    return any(n in names for n in CONTENT_SHEETS)


def find_prior(folder, workstream, before=None):
    """(date, path, alternates) of the newest matching report strictly before `before`."""
    folder = os.path.abspath(folder)
    names = {workstream, workstream.replace("/", "-").replace(" ", "_")}  # build_report's default
    dated = re.compile(r"Meeting_Report_(?:%s)_(\d{4}-\d{2}-\d{2})(.*)\.xlsx$"
                       % "|".join(re.escape(n) for n in names))
    dirs = [folder, os.path.join(folder, workstream),
            os.path.join(os.path.dirname(folder), workstream)]
    found = {}
    for d in dirs:
        for p in glob.glob(os.path.join(glob.escape(d), "*Meeting_Report_*.xlsx")):
            m = dated.search(os.path.basename(p))
            if not m or (before and m.group(1) >= before):
                continue
            found[p] = (m.group(1), has_content_sheets(p), m.group(2) == "", os.path.getmtime(p))
    if not found:
        return None, None, []
    ranked = sorted(found, key=lambda p: found[p], reverse=True)
    best = ranked[0]
    alternates = [p for p in ranked[1:] if found[p][0] == found[best][0]]
    return found[best][0], best, alternates


def item(sheet, kind, date, row, **fields):
    sr = g(row, "Sr")
    ref = f"{kind} {sr}" if str(sr).isdigit() else str(sr)   # older reports number rows A3, Q1
    out = {"sheet": sheet, "sr": sr, "first_raised": f"{date} · {ref}"}
    out.update({k: s(v) for k, v in fields.items() if s(v)})
    return out


def collect(wb, date):
    items = []
    for row in rows_of(wb, "Action Items"):
        st = s(row.get("Status"))
        if st in ("Open", "In Progress"):
            items.append(item("Action Items", "Action", date, row,
                              en=row.get("Action"), hi=row.get("Action (Hindi)"),
                              owner=row.get("Owner"), status=st, due=row.get("Due"),
                              due_basis=row.get("Due basis"), priority=row.get("Priority")))
    for row in rows_of(wb, "Open Questions"):
        st = s(row.get("Status"))
        if st in ("Open", "Escalated"):
            items.append(item("Open Questions", "Question", date, row,
                              en=g(row, "Open question"), hi=g(row, "Open question (Hindi)"),
                              owner=g(row, "Must be answered by"), status=st,
                              blocks=g(row, "What it blocks / impact")))
    for row in rows_of(wb, "Decisions"):
        t = s(row.get("Decided / Deferred"))
        if t.startswith("Deferred"):
            items.append(item("Decisions", "Decision", date, row,
                              en=row.get("Decision"), hi=row.get("Decision (Hindi)"),
                              owner=g(row, "If deferred — pending on") or g(row, "By whom"),
                              status=t))
    for row in rows_of(wb, "Carry Forward"):
        st = s(row.get("Status this meeting"))
        if st.startswith("Still open") or st == "Not mentioned" or not st:
            it = item("Carry Forward", "Carried", date, row,
                      en=g(row, "Item carried forward"), hi=row.get("Item (Hindi)"),
                      owner=row.get("Owner"), status=st or "(blank - was never reconciled)",
                      note=g(row, "Note"))
            # The original reference outranks this report's own row number. Older reports
            # split it across "From meeting" and "Ref (original)".
            original = s(row.get("First raised")) or (
                f"{s(row.get('From meeting'))} · {s(row.get('Ref (original)'))}"
                if s(row.get("From meeting")) else "")
            it["first_raised"] = original or it["first_raised"]
            items.append(it)
    return items


def main():
    ap = argparse.ArgumentParser(description="Open items from the previous report in a workstream.")
    ap.add_argument("folder", help="the transcript's folder")
    ap.add_argument("workstream")
    ap.add_argument("--before", metavar="YYYY-MM-DD",
                    help="ignore reports dated on or after this (the meeting being written up)")
    a = ap.parse_args()

    date, path, alternates = find_prior(a.folder, a.workstream, a.before)
    if not path:
        print(json.dumps({"source_report": None, "report_date": None, "count": 0, "items": [],
                          "note": f"No earlier Meeting_Report_{a.workstream}_*.xlsx found - "
                                  "this report is the first in the chain."}, indent=2))
        return
    wb = openpyxl.load_workbook(path, data_only=True)
    items = collect(wb, date)
    out = {"source_report": path, "report_date": date, "count": len(items), "items": items}
    if alternates:
        out["other_files_with_same_date"] = alternates   # tell the reader which copy was used
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
