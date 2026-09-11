---
name: meeting-report
description: Turn a meeting transcript into a bilingual (English/Hindi) Excel meeting report — key points, action items with owner and due date, decisions taken vs deferred, open questions, risks, and open items carried forward from the previous meeting in the same workstream. Use for meeting minutes, summaries, action items, or "what did we decide" questions about a recorded call.
---

# Meeting report

Produces one Excel workbook per meeting from an auto-generated transcript. Every row is bilingual and every row cites the transcript words it came from, so a reader can check the reading without replaying the call.

## When to use

A meeting transcript (`.rtf`, `.txt`, `.docx`, `.vtt`) and a request for minutes, a summary, action items, or a meeting report. Also when someone asks "what did we decide" or "what's still open" about a recorded meeting.

## Non-negotiables

1. **Read the entire transcript before writing anything.** Commitments cluster in the last 15%, and the framing that makes them make sense is in the first 15%. Never work from a grep or a partial read.
2. **Every row cites evidence.** If you cannot quote the transcript words behind a row, the row does not go in. No inference presented as fact.
3. **Never invent a due date.** If no timing was stated, the Due basis is `None`. That is a finding, not a gap to fill.
4. **Flag what you are unsure of** in the Confidence column rather than guessing. These transcripts merge speakers; a wrong owner reaching a developer or a leader is the main failure mode.
5. **Evidence is verbatim — including the errors.** Copy the transcript's exact words. Keep the stutters (`you- you`), the filler (`you know`, `uh`), and the mis-transcribed terms: if the transcript says `DPTP`, the Evidence cell says `DPTP` even though the report body says DPDP. Normalising inside a quote defeats the point of having one, and `verify_report.py` will reject it. Use `...` to elide a middle section — each fragment either side is checked separately. Any caveat of your own goes in square brackets at the end, which is not checked.
6. **Write the content JSON in one file write, and build with one command.** Never assemble either incrementally across several edits — each round trip costs more wall-clock than the thinking did.

## Step 1 — Ask whether they want Hindi

**Ask before doing anything else**, so nobody waits through a transcript read to be asked a
question. Offer exactly two options:

- **English only** — the default choice, and noticeably faster
- **English + Hindi** — both, as paired columns

Use the session's question tool if there is one; otherwise ask in plain text and wait.

Three rules:

- **Ask once per run, not once per transcript.** If the request covers several meetings, one
  answer applies to all of them.
- **Do not ask if the request already answers it** — "make the report, English only", or
  "हिन्दी में भी चाहिए". Take them at their word and proceed.
- **If nobody can answer** (a scheduled or batch run), build English only and say so in the
  delivery, so the reader knows Hindi was skipped rather than lost.

Record the answer as `"language": "english"` or `"language": "bilingual"` at the top level of
the content JSON.

**If they choose English only, write no Hindi at all.** Omit `hi`, `rationale_hi`, `impact_hi`
and `exposure_hi` entirely, and make each `key_points` entry a single-element list. Do not
translate "just in case" and do not write empty Hindi strings — the whole saving is in not
generating it. The Hindi columns still exist in the workbook but are hidden, so the same file
can be filled in later without rebuilding.

## Step 2 — Load the glossary

Look for `meeting-report-glossary.md` in the transcript's folder, its parent, or its grandparent. It holds everything organisation-specific: canonical names of people and how transcripts mangle them, term and acronym variants, the workstream list, meeting archetypes, domain object codes and process order, approval thresholds, and recurring events that function as deadlines.

**Read it before reading the transcript.** Without it you will mis-attribute actions to misheard names.

If there is no glossary, say so once, work from the transcript alone, mark name-dependent rows `Medium` confidence, and offer to start a glossary from what you inferred. `reference/glossary-template.md` in this skill is the structure to follow — never treat its placeholder examples as real.

## Step 3 — Convert the transcript

**Preflight first:** `python3 -c "import openpyxl"`. **On Windows the interpreter is `py`
(or `python`), not `python3`** — use it there in this and every command below. If the
interpreter is not found or the import fails, **stop and tell the user what to install** — the
README section *What the computer needs* has copy-paste commands for Mac and Windows (Python 3,
then `pip install openpyxl`; LibreOffice on Windows/Linux for `.rtf`/`.docx`). If only openpyxl
is missing, offer to run the pip command for them. Never hand-roll a converter or a workbook to
get around a missing dependency.

```bash
python3 <skill-dir>/scripts/to_text.py "<file>.rtf" --outdir <tmp>
```

`<skill-dir>` is the folder holding this SKILL.md — `scripts/` and `reference/` live beside it,
**not** in the transcript folder, so always use the absolute path. The script handles `.rtf`,
`.docx`, `.txt` and `.vtt` (cue timestamps stripped), uses LibreOffice where it is installed and
falls back to macOS `textutil` where it is not, and prints the `.txt` path plus its line count.
`pandoc` does **not** read these RTFs — do not reach for it. Work in a scratch directory outside
the user's folders; write only the finished workbook into their folder.

Transcripts typically open with a header block:

```
Meeting Title: <title>
Date: <Mon D>
Meeting participants: <names>

Transcript:
<Speaker>: <text>
```

The title and participant list are frequently wrong or incomplete — the participant line often names only the recorder while five people speak. Rebuild both from the body.

## Step 4 — Establish meeting identity

- **Workstream** = the parent folder name. It drives carry-forward, so get it right. The glossary lists the valid workstreams.
- **Date** = the `Date:` header, resolved to a full year from the file's modification time.
- **Meeting type** — the glossary maps workstreams to archetypes. Each shifts the emphasis:
  - *Leadership / strategy* — Decisions and Open Questions carry the weight; actions are few and soft-dated.
  - *Operational / working session* — Action Items carry the weight, one row per screen or object, listed in the process order given in the glossary.
  - *Discovery / induction* — mostly Key Points and Open Questions. Expect few real actions and do not manufacture them.

## Step 5 — Read the transcript, knowing how it lies

These are machine transcripts of multilingual calls. Specifically:

- **Read to the last line.** `to_text.py` prints the line count; the file reader stops at
  2,000 lines, so a long call needs reading in offsets until the end. The commitments are at
  the end — stopping early is the one way to miss them all.
- **`Me:` is whoever recorded the call.** Not a name. Resolve it from the glossary or the surrounding context.
- **`Them:` and `Speaker A:` are not one person.** A single such block routinely contains three people talking in turn. Never attribute an action to `Them` — read the surrounding turns to work out who actually spoke, and mark Confidence `Medium` or `Low` when you cannot.
- **Names garble badly.** Normalise silently against the glossary's variant table. A name that is not in the glossary and not clearly spelled is a `Medium` confidence row, and worth flagging so someone can add it.
- **Scripts get swapped** — English transliterated into the local script and vice versa. Read for meaning, not spelling.
- **Stray foreign-language fragments** are transcription noise from crosstalk. Ignore them.
- **Repeated phrases** are stutter artefacts, not emphasis.

When the transcript is genuinely unreadable at a point that matters, say so in the Evidence cell in square brackets rather than smoothing it over.

## Step 6 — Extract into the content JSON

Write ONE JSON file (schema in Step 8) holding six sets. Each free-text field needs an English and a Hindi version — write natural Hindi, not machine translation, and keep proper nouns and system names in Latin script inside the Hindi text.

### Key points (8–12)
What someone who missed the call must know. Substance only — no "the team discussed X". Include disagreements and reversals: when a leader overrides a framing, that is a key point.

### Action items
One row per commitment. Columns: `Sr` · `Action` · `Action (Hindi)` · `Owner` · `Raised by` · `Due` · `Due basis` · `Priority` · `Status` · `Confidence` · `Evidence`.

- **Due basis** — `Explicit` (a date or deadline was stated) · `Relational` (tied to an event, e.g. "before the campaign", "by next call") · `Conditional` (starts when something else happens) · `None` (nothing agreed). Most actions in these meetings are Relational; `None` rows are the ones that quietly slip, so surface them in the Summary count.
- **Confidence** — `High` (owner named themselves, or was named directly and did not object) · `Medium` (inferred from context, or attributed by a third party without confirmation) · `Low` (speaker unclear in transcript).
- A commitment someone makes about themselves is High. A commitment someone assigns to an absent person is Medium at best.
- Status starts `Open` for everything.

### Decisions
One row per decision, with `Decided / Deferred` in its own column. A deferred decision must name **who it now waits on** — that column is the whole point of separating the two. Record the rationale where it was given; a decision without its reasoning gets relitigated next meeting.

### Open questions
Only questions with a **named person who owes an answer**. For each, state what it blocks in concrete terms, and quote the transcript words that raised it. A question with no owner is escalated in the report, not silently listed — `verify_report.py` rejects the row.

### Risks and compliance
Control gaps, data-privacy exposure, financial-control weaknesses, audit issues, reporting-integrity problems. Write **Exposure in business terms** — this is the column leadership reads, so no system vocabulary. The pattern that distinguishes a control failure from a bug is "…and nothing reports it". Severity: High / Medium / Low.

Where the glossary records a threshold or rule as unconfirmed, never state a figure heard in a transcript as settled — log it as an Open Question.

### Carry forward
See Step 7.

## Step 7 — Carry forward

**If the user says they do not want carry-forward** — "no carry forward", "earlier notes are not
relevant", "this report stands alone" — skip this whole step, leave `carry_forward` out of the
content JSON, and build with `--no-carry-forward`. That drops the sheet together with its Summary
count and its How to Use entries, so nothing in the workbook points at a sheet that is not there.
Do not delete the sheet from a built workbook by hand: the Summary formula would break. Otherwise:

```bash
python3 <skill-dir>/scripts/prior_open_items.py "<transcript folder>" "<Workstream>" --before <YYYY-MM-DD>
```

It finds the newest earlier `Meeting_Report_<Workstream>_*.xlsx` (same folder, or the workstream
folder) and prints every item still open in it: Action Items `Open`/`In Progress`, Open Questions
`Open`/`Escalated`, `Deferred` decisions, and its own Carry Forward rows still `Still open*` or
`Not mentioned`. `--before` is this meeting's date, so a re-run never reads its own report.
Do not write spreadsheet code to read the prior workbook.

- Carry across **every item the script lists**, not just the last meeting's — the newest report
  is then always the complete open-items picture, with no separate master file to keep in sync.
  A prior report's own open items and its carried items are both in the list.
- Reconcile each against what was said this time: `Closed` · `Still open` · `Still open - restated`
  · `Superseded` · `Not mentioned` — exactly these words; the column is a dropdown and the
  Summary counts on them. Quote the evidence for a `Closed`.
- Only mark `Closed` on clear evidence. `Not mentioned` is the honest call when the item simply
  did not come up — never quietly drop it. It still counts as open in the Summary.
- Also catch carry-forward items **admitted inside this transcript** — someone conceding they
  never sent a promised note or list is a carry-forward row even with no prior report on disk.
- **`first_raised` is the cross-meeting reference.** Copy it from the script's output verbatim —
  `2026-08-13 · Action 2`, `2026-08-06 · Question 1`. For an item admitted in this transcript
  with no prior report, write the meeting it was promised in if that is stated, else this
  meeting's date. **Never renumber** and never re-derive it from the new report's own rows.

If no prior report exists the script says so; still create the sheet and note that this is the
first in the chain.

## Step 8 — Build the workbook

**Do not write openpyxl code.** The plugin ships `scripts/build_report.py`, which owns every
column header, colour, row height, dropdown, conditional-format rule, Summary formula and the
whole How to Use sheet. You supply content only:

```bash
python3 <skill-dir>/scripts/build_report.py content.json \
  --out "<folder>/Meeting_Report_<Workstream>_<YYYY-MM-DD>.xlsx" \
  --language english|bilingual [--no-carry-forward]
```

`--language` may be omitted if the JSON carries a `language` key; the flag wins if both are
given. English-only hides the Hindi columns rather than removing them, because the Summary
formulas address columns by letter and dropping one would silently break every count.

It builds the workbook, recalculates it through LibreOffice (resolving the binary on Linux,
macOS and Windows), and prints the row counts and recalc status as JSON. If LibreOffice is
absent it says so and continues — the workbook is still correct, because Excel recalculates on
open; only automated verification of the counts is unavailable.

### content.json schema

Every field is a string unless noted. Omit a key and it renders empty; never invent a value to
fill one. When `language` is `english`, the `_hi` fields below are omitted entirely rather than
left blank, and `key_points` entries hold one element instead of two.

```json
{
  "language": "english|bilingual",
  "meeting": {
    "title_en": "", "title_hi": "", "workstream": "", "type": "",
    "date": "3 September 2026", "date_iso": "2026-09-03",
    "participants": "", "recorded_by": "", "source": ""
  },
  "key_points": [["english", "hindi"]],
  "actions": [{
    "sr": 1, "en": "", "hi": "", "owner": "", "raised_by": "", "due": "",
    "due_basis": "Explicit|Relational|Conditional|None",
    "priority": "High|Medium|Low", "status": "Open",
    "confidence": "High|Medium|Low", "evidence": ""
  }],
  "decisions": [{
    "sr": 1, "en": "", "hi": "", "type": "Decided|Deferred", "by": "",
    "pending_on": "", "rationale_en": "", "rationale_hi": "", "evidence": ""
  }],
  "questions": [{
    "sr": 1, "en": "", "hi": "", "answer_by": "",
    "impact_en": "", "impact_hi": "", "raised_in": "", "status": "Open", "evidence": ""
  }],
  "risks": [{
    "sr": 1, "en": "", "hi": "", "category": "", "exposure_en": "", "exposure_hi": "",
    "severity": "High|Medium|Low", "owner": "", "mitigation": "", "evidence": ""
  }],
  "carry_forward": [{
    "sr": 1, "en": "", "hi": "", "owner": "", "first_raised": "2026-08-13 · Action 2",
    "status": "Closed|Still open|Still open - restated|Superseded|Not mentioned",
    "note": "", "evidence": ""
  }]
}
```

`type` may carry a qualifier — `Decided (clarification)` — and the Summary formulas use wildcards
so those still count. Keep the qualifier after the base word. Every other dropdown column
(`due_basis`, `priority`, `status`, `confidence`, `severity`, carry-forward `status`) takes exactly
the listed words; anything else is a silent miscount, and `verify_report.py` rejects it.

`evidence` is required on every row of every sheet, with one exception: a carry-forward row whose
status is not `Closed` may leave it empty (nothing new was said). `raised_in` is where in this
meeting the question came up — an agenda item or topic, not a date. `source` is the transcript
file name.

A complete worked example — fictional organisation, fictional names — is in
`reference/example-content.json` with its transcript `reference/example-transcript.txt`; the pair
builds and verifies clean, so it is also the smoke test after any script change.

## Step 9 — Verify

One command, not one shell call per row:

```bash
python3 <skill-dir>/scripts/verify_report.py "<the .xlsx>" "<the converted transcript .txt>"
```

It checks every Evidence quote against the transcript and that no row is missing one, scans all
sheets for formula errors, recounts every Summary figure from the sheets (when LibreOffice has
recalculated the file), enforces the content rules above — every decision Decided or Deferred, a
deferred decision names who it waits on, an open question names who answers, dropdown columns hold
only their allowed words — and lists the Medium and Low confidence rows. Exit 0 is clean; exit 1
means fix something.

**A quote failure means your reading drifted — correct the row, never soften the quote.** The
commonest causes are dropping a stutter or filler word, and normalising a mis-transcribed name
or acronym inside the quotation marks. **A `rows_with_no_evidence` entry means the row has no
transcript words behind it — quote them or drop the row**; a fragment under 12 characters does
not count, so quote enough of the turn.

Re-run until it passes. `needs_human_confirmation`, `caveat_only_evidence` (rows whose only
evidence is an `[unreadable]` note) and `warnings` are not failures; they are what you flag to the
reader in Step 10.

## Step 10 — Deliver

The workbook is written straight into the user's folder, so it is already delivered. Tell them the folder and file name, and lead with what the meeting actually produced — how many actions and who holds the high-priority ones, what was decided against what was deferred, and anything High severity. Name the rows that need human confirmation before the report is circulated, including any whose only evidence is an `[unreadable]` note.

Do not restate the workbook's contents at length; they can open it.

If any name could not be resolved against the glossary, say so and offer to add it — the glossary is meant to improve every time it is used.
