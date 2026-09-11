# Goonj Claude skills

Shared Claude skills for Goonj teams, distributed as a plugin.

Currently one skill: **meeting-report** — turns a meeting transcript into a bilingual
English/Hindi Excel report with key points, action items (owner, due date, priority),
decisions taken versus deferred, open questions, risks, and open items carried forward
from the previous meeting in the same workstream.

---

## Install (one time)

You need a paid Claude plan — Pro, Max, Team or Enterprise. Plugins are not available on
the free tier.

1. Open the **Claude desktop app**. If you are in Cowork, open the **Cowork** tab first.
2. In the left sidebar, open **Customize**, then the **Plugins** tab.
3. Choose **Add from a repository** and paste this repository's URL.
4. Install the **goonj-meetings-skill** plugin.

Command line, if you prefer:

```bash
/plugin marketplace add <owner>/<repo>
/plugin install goonj-meetings-skill@goonj-skills
```

## Set up the glossary (one time, per transcripts folder)

The skill reads an organisation glossary so it can correct mangled names and know your
workstreams. **The glossary is not in this repository** — it names real people, so it
lives with the transcripts instead.

1. Copy `plugins/goonj-meetings-skill/skills/meeting-report/reference/glossary-template.md`
2. Save it in your transcripts folder as **`meeting-report-glossary.md`**
3. Fill in the real names, term variants, workstreams and thresholds

Anyone on the team can edit it. Add a name the transcripts keep getting wrong and the next
report picks it up. If a Goonj glossary already exists in your shared transcripts folder,
you do not need to create one.

## Use it

Connect your transcripts folder, then ask in plain language:

> Make a meeting report from the 3 September Leadership transcript

The workbook is written next to the transcript as
`Meeting_Report_<Workstream>_<YYYY-MM-DD>.xlsx`.

---

## उपयोग कैसे करें (संक्षेप में)

1. Claude डेस्कटॉप ऐप में **Customize → Plugins → Add from a repository** से यह प्लगइन जोड़ें।
2. अपने ट्रांसक्रिप्ट फ़ोल्डर में `meeting-report-glossary.md` फ़ाइल रखें (टेम्पलेट ऊपर बताए पथ पर है)।
3. फिर बस इतना कहें: "3 सितंबर की Leadership मीटिंग की रिपोर्ट बना दीजिए"।

रिपोर्ट एक Excel फ़ाइल के रूप में उसी फ़ोल्डर में बन जाएगी जिसमें ट्रांसक्रिप्ट है। हर पंक्ति अंग्रेज़ी
और हिन्दी दोनों में होती है, और हर पंक्ति के साथ ट्रांसक्रिप्ट से प्रमाण दिया जाता है।

**ध्यान दें:** जिन पंक्तियों में Confidence कॉलम `Medium` या `Low` है, उन्हें आगे भेजने से पहले
एक बार जाँच लें। ट्रांसक्रिप्ट कभी-कभी नाम ग़लत सुनते हैं और वक्ताओं को आपस में मिला देते हैं।

---

## What the workbook contains

| Sheet | Contents |
|---|---|
| Summary | Meeting details, live counts, and 8–12 key points |
| Action Items | Every commitment, with owner, due date, priority and a confidence rating |
| Decisions | What was settled, and separately what was deferred and who it waits on |
| Open Questions | Unresolved questions, each with a named person who owes an answer |
| Risks and Compliance | Control gaps and exposure, written in business terms |
| Carry Forward | Open items from earlier meetings and where each now stands |
| How to Use | Bilingual legend for every column and status value |

Status, Priority and Confidence are dropdowns. Changing a Status updates the Summary
counts automatically, so the workbook works as a live tracker between meetings.

## Two things to know before circulating a report

1. **Check every row marked `Medium` or `Low` confidence.** Auto-transcripts mishear names
   and merge several speakers into one block. The skill flags what it is unsure of rather
   than guessing, but a person has to make the final call.
2. **`None` in the Due basis column means no timing was agreed on the call.** That is a
   finding, not a formatting gap. Those are the actions that slip.

## For maintainers — how the skill is structured

The skill ships two scripts, and the agent is told not to write spreadsheet code itself:

| File | Role |
|---|---|
| `scripts/build_report.py` | Takes a JSON of extracted content, writes the formatted workbook, recalculates it. Owns every colour, header, formula, dropdown and the whole How to Use sheet. |
| `scripts/verify_report.py` | One pass over the finished workbook: every Evidence quote checked against the transcript, formula errors scanned, counts reconciled, low-confidence rows listed. |

This keeps report formatting identical across every meeting and every person, and stops the
agent re-deriving the same layout on each run. Change a colour or add a column **in the script**,
not in `SKILL.md`.

Both scripts need `openpyxl` (`pip install openpyxl`). `build_report.py` finds LibreOffice on
Linux, macOS and Windows; if it is not installed the workbook is still correct — Excel
recalculates on open — but `verify_report.py` cannot check the counts. Set `SOFFICE=/path/to/soffice`
to point at a non-standard install.

## Maintainers

Bump `version` in `plugins/goonj-meetings-skill/.claude-plugin/plugin.json` on every change — that is what
pushes the update to everyone who has the plugin installed. Users refresh with
`/plugin marketplace update`.

Never commit a completed `meeting-report-glossary.md`, a transcript, or a generated report.
`.gitignore` blocks all three, but check before you push.

## Reusing this outside Goonj

Nothing organisation-specific lives in the skill — names, workstreams, object codes and
thresholds all sit in the glossary. Another organisation can use this by writing their own
glossary and changing nothing else.
