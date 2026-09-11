#!/usr/bin/env python3
"""List every still-open item from the previous report in a workstream, for carry-forward.

    python3 prior_open_items.py <folder> <workstream> [--before YYYY-MM-DD]

Finds the newest Meeting_Report_<Workstream>_YYYY-MM-DD.xlsx in <folder>, <folder>/<workstream>
or <folder>/../<workstream> - newest by the date in the filename, never by modification time -
and prints JSON of the items still open in it:

  Action Items     Status is Open or In Progress
  Open Questions   Status is Open or Escalated
  Decisions        Decided / Deferred starts with Deferred
  Carry Forward    Status this meeting is Still open*, Not mentioned, or blank

Every item carries `first_raised`, the cross-meeting reference: "<date> · Action 3",
"<date> · Question 2", "<date> · Decision 1" for the report's own rows, and passed through
unchanged for rows that were already on its Carry Forward sheet, so the original reference
survives the whole chain. Copy it into the new Carry Forward row verbatim.

--before excludes reports dated on or after the meeting being written up, so re-running a
meeting never carries forward from its own report. Reads columns by header name, so reports
built by earlier versions of the skill work too.
"""
import argparse, glob, json, os, re, sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed.  Fix:  pip install openpyxl")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_report import rows_of, s

DATED = re.compile(r"_(\d{4}-\d{2}-\d{2})\.xlsx$")


def find_prior(folder, workstream, before=None):
    """(date, path) of the newest matching report strictly before `before`, else (None, None)."""
    folder = os.path.abspath(folder)
    names = {workstream, workstream.replace("/", "-").replace(" ", "_")}  # build_report's default
    dirs = [folder, os.path.join(folder, workstream),
            os.path.join(os.path.dirname(folder), workstream)]
    found = set()
    for d in dirs:
        for n in names:
            for p in glob.glob(os.path.join(glob.escape(d), f"Meeting_Report_{n}_????-??-??.xlsx")):
                m = DATED.search(p)
                if m and (not before or m.group(1) < before):
                    found.add((m.group(1), p))
    return max(found) if found else (None, None)


def item(sheet, kind, date, row, **fields):
    out = {"sheet": sheet, "sr": row.get("Sr"), "first_raised": f"{date} · {kind} {row.get('Sr')}"}
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
                              en=row.get("Open question"), hi=row.get("Open question (Hindi)"),
                              owner=row.get("Must be answered by"), status=st,
                              blocks=row.get("What it blocks / impact")))
    for row in rows_of(wb, "Decisions"):
        t = s(row.get("Decided / Deferred"))
        if t.startswith("Deferred"):
            items.append(item("Decisions", "Decision", date, row,
                              en=row.get("Decision"), hi=row.get("Decision (Hindi)"),
                              owner=row.get("If deferred — pending on") or row.get("By whom"),
                              status=t))
    for row in rows_of(wb, "Carry Forward"):
        st = s(row.get("Status this meeting"))
        if st.startswith("Still open") or st == "Not mentioned" or not st:
            it = item("Carry Forward", "Carried", date, row,
                      en=row.get("Item carried forward"), hi=row.get("Item (Hindi)"),
                      owner=row.get("Owner"), status=st or "(blank - was never reconciled)",
                      note=row.get("Note"))
            # The original reference outranks this report's own row number.
            it["first_raised"] = s(row.get("First raised")) or it["first_raised"]
            items.append(it)
    return items


def main():
    ap = argparse.ArgumentParser(description="Open items from the previous report in a workstream.")
    ap.add_argument("folder", help="the transcript's folder")
    ap.add_argument("workstream")
    ap.add_argument("--before", metavar="YYYY-MM-DD",
                    help="ignore reports dated on or after this (the meeting being written up)")
    a = ap.parse_args()

    date, path = find_prior(a.folder, a.workstream, a.before)
    if not path:
        print(json.dumps({"source_report": None, "report_date": None, "count": 0, "items": [],
                          "note": f"No earlier Meeting_Report_{a.workstream}_*.xlsx found - "
                                  "this report is the first in the chain."}, indent=2))
        return
    wb = openpyxl.load_workbook(path, data_only=True)
    items = collect(wb, date)
    print(json.dumps({"source_report": path, "report_date": date, "count": len(items),
                      "items": items}, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
