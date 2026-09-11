# Meeting report glossary — TEMPLATE

Copy this file to your transcripts folder, rename it to `meeting-report-glossary.md`,
and replace every placeholder with your organisation's real values.

**Keep the completed glossary with your transcripts. Never commit it to this repository** —
it names real people and internal thresholds, and the repository may be public.

The skill looks for `meeting-report-glossary.md` in the transcript's folder, its parent,
or its grandparent. Anyone on the team can edit it; the next report picks up the change.

---

## People — canonical names and how transcripts mangle them

Add a row whenever a transcript mishears someone. This table is the single biggest
lever on report accuracy.

| Canonical name | Appears in transcripts as | Role in these meetings |
|---|---|---|
| Full Name (Short name) | Variant1, Variant2, Variant3 | What they own or decide |
| Full Name | Variant1, Variant2 | Usually the recorder, so labelled `Me:` |

## Organisation and term variants

Acronyms, product names and the organisation's own name. Transcription engines mangle
these constantly.

| Canonical | Appears as |
|---|---|
| YourOrg | Variant1, Variant2 |
| SYSTEM_NAME | Variant1, Variant2 |
| ACRONYM | Variant1, Variant2 |

## Workstreams

Each is a folder under the transcripts root. The folder name is the workstream.

`Workstream A` · `Workstream B` · `Workstream C`

## Meeting archetype by workstream

| Archetype | Workstreams | What carries the weight |
|---|---|---|
| Leadership / strategy | ... | Decisions and Open Questions. Few actions, softly dated. |
| Operational / working session | ... | Action Items, one row per screen or object. |
| Discovery / induction | ... | Key Points and Open Questions. Do not manufacture actions. |

## Domain object codes and process order

Only needed if you run operational working sessions that walk a system screen by screen.
List the codes in the order the process actually runs, so actions come out in that order.

`CODE1` Full name → `CODE2` Full name → `CODE3` Full name

Supporting objects: `CODE_X` Full name · `CODE_Y` Full name (proposed, does not exist yet).

## Approval thresholds and rules

State each rule, and mark clearly anything **not yet formally confirmed** — the skill
logs unconfirmed figures as Open Questions rather than reporting them as settled.

- Below <amount> — <what happens>
- <amount> to <amount> — <who approves>
- Above <amount> — <what happens>
- <Rule name> — **not yet formally confirmed.** Treat any figure heard in a transcript as
  unconfirmed.

## Recurring events

| Term | Meaning |
|---|---|
| EVENT_NAME | What it is, and why deadlines get pinned to it. |
| Next call / next update | How often these meetings recur, so "by next call" resolves to a real interval. |
