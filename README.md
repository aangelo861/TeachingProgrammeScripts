# Foundation Year Psychiatry Teaching Programme — Excel Automation

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Language](https://img.shields.io/badge/language-TypeScript-3178C6.svg)](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel)
[![Platform](https://img.shields.io/badge/platform-Excel%20Online-217346.svg)](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel)

*A beginner-friendly guide to using Excel Online scripts and AI tools to automate teaching feedback, attendance, and certificate workflows.*

This guide shows you how to use Microsoft Forms, an Excel workbook, and a handful of **[Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel)** to automate the admin behind a clinical teaching programme — turning raw feedback responses into session summaries, marking attendance automatically, flagging trainees who are missing sessions, and generating presenter certificates and letters.

The scripts in this repository are **worked examples**, built around one real teaching programme's workbook. They're meant to be read, understood, and **adapted** — including with the help of an AI assistant like Copilot, Claude, or ChatGPT — rather than copied in blindly. If your workbook looks different, that's expected; this guide shows you how to adjust the scripts to match it.

> [!NOTE]
> These are Office Scripts written in TypeScript, running against the `ExcelScript` API from Excel's **Automate** tab. They are **not** Node.js scripts — you can't run them with `ts-node` or `tsc` on your own computer; they only execute against a live workbook, inside Excel.

> [!TIP]
> Already comfortable with the ExcelScript API? Skip straight to [Technical reference](#technical-reference-the-scripts-in-detail) and [Configuration reference](#configuration-reference).

---

## Contents

- [Who this is for](#who-this-is-for)
- [What this project helps you automate](#what-this-project-helps-you-automate)
- [The basic idea](#the-basic-idea)
- [Before you start: key concepts](#before-you-start-key-concepts)
- [The workbook you need](#the-workbook-you-need)
- [Getting your Microsoft Forms responses into Excel](#getting-your-microsoft-forms-responses-into-excel)
- [Running your first Office Script](#running-your-first-office-script)
- [Using Copilot, Claude, ChatGPT, or another AI to adapt these scripts](#using-copilot-claude-chatgpt-or-another-ai-to-adapt-these-scripts)
- [Debugging scripts iteratively](#debugging-scripts-iteratively)
- [Your learning path through this repo](#your-learning-path-through-this-repo)
- [Worked example: one session, start to finish](#worked-example-one-session-start-to-finish)
- [What each script does (plain English)](#what-each-script-does-plain-english)
- [Adapting this repository to your own workbook](#adapting-this-repository-to-your-own-workbook)
- [Technical reference: the scripts in detail](#technical-reference-the-scripts-in-detail)
- [Configuration reference](#configuration-reference)
- [Common problems and fixes](#common-problems-and-fixes)
- [Maintenance](#maintenance)
- [Data protection and safe use with AI tools](#data-protection-and-safe-use-with-ai-tools)
- [Optional advanced setup: email, certificates and the full pipeline](#optional-advanced-setup-email-certificates-and-the-full-pipeline)
- [Glossary](#glossary)
- [License](#license)

---

## Who this is for

This guide is written for **programme administrators, medical educators, teaching leads, rota and admin teams, and clinicians** who use Excel day-to-day but have never written or run a script before.

You do **not** need a formal coding background. The example scripts happen to be written in TypeScript, but you're not expected to read or write TypeScript from scratch. You mainly need to be able to describe your workbook (sheet names, column headings) to an AI assistant, run a script, and recognise an error message well enough to ask for help fixing it. Microsoft's own documentation frames Office Scripts as a way to automate everyday Excel tasks using an Action Recorder and a Code Editor built directly into Excel — this project takes that same idea further, using AI assistance to help you adapt example scripts to your own workbook.

If you can open a workbook, find a ribbon tab, and copy-paste text, you can use this guide.

---

## What this project helps you automate

Running a teaching programme generates a lot of repetitive admin: reading through dozens of feedback forms by hand, working out who attended which session, chasing trainees who've missed several sessions in a row, and putting together certificates and thank-you letters for guest presenters. This repository contains three example Office Scripts that automate that admin inside Excel:

- **Summarising feedback** — turning a sheet of raw, one-row-per-response survey answers into one clear summary row per teaching session, with response counts, percentages, and comments.
- **Marking attendance** — matching feedback responses to a named trainee register, automatically marking who attended, and emailing the team a digest that flags trainees missing sessions.
- **Generating presenter letters and certificates** — building a personalised thank-you letter and certificate for a guest presenter from that session's feedback data.

These scripts are **examples, not a plug-and-play product**. They were built around one specific teaching programme's workbook and workflow. The rest of this guide shows you both how to run them as-is, and — just as importantly — how to adapt them, ideally with an AI assistant's help, to a workbook and workflow that looks like yours.

---

## The basic idea

At the simplest level, this whole project works like this:

```
Microsoft Forms responses  →  Excel workbook  →  Office Script  →  new sheet, email, or document
```

A teaching session happens. Attendees fill in a Microsoft Forms feedback survey afterwards. Those responses land as new rows in a sheet in your Excel workbook. You run an Office Script from Excel's **Automate** tab. The script reads that raw data, processes it, and writes its results back into the workbook — and for two of the three scripts, also sends an email that a further, optional pipeline turns into certificates and letters (see [Optional advanced setup](#optional-advanced-setup-email-certificates-and-the-full-pipeline)).

That's the whole idea. Everything else in this repository — the specific sheets, the column indexes, the Cloudflare Worker, the Google Apps Script — exists to support that one loop: **feedback in, useful output out.** The detailed, full-pipeline diagram lives later in this guide, in the advanced section, once the basic loop makes sense.

---

## Before you start: key concepts

A handful of terms come up constantly in this guide. If you already know Excel but have never touched a script, here's what you need:

| Term | Plain-English meaning |
| --- | --- |
| **Workbook** | The whole Excel file — the thing you open, e.g. `Teaching Programme.xlsx`. |
| **Worksheet / sheet** | One tab inside the workbook, e.g. `Sample Data` or `Attendance Record`. A workbook usually has several sheets. |
| **Microsoft Forms response sheet** | A sheet that Microsoft Forms automatically fills in — one row per person who submits the form, one column per question. |
| **Office Script** | A small program, written in TypeScript, that lives inside Excel and can read and change your workbook automatically. You run it from the **Automate** tab. |
| **TypeScript** | The programming language Office Scripts are written in — a more structured relative of JavaScript. You don't need to learn it to use this repo; you can describe what you want in plain English and have an AI assistant write or adapt the TypeScript for you. |
| **Column index** | A number a script uses to refer to a column, counting **from zero**. This trips up almost everyone at first — see below. |

**Zero-based column indexes, explained.** Excel shows column letters (A, B, C…), but scripts refer to columns by number, and that numbering starts at 0, not 1:

| Excel column | A | B | C | D | E | F |
| --- | --- | --- | --- | --- | --- | --- |
| Script index | 0 | 1 | 2 | 3 | 4 | 5 |

So if a script has a constant like `SAMPLE_EMAIL_COL = 4`, that means "the email address is in column **E**," not column D or F. When a script errors or produces the wrong result, a mismatched column index is one of the first things to check — see [Common problems and fixes](#common-problems-and-fixes).

**Where to learn more.** You don't need to read these before continuing, but they're useful background from Microsoft, and are referenced again in the [Glossary](#glossary):
- [Office Scripts overview](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) — what Office Scripts are and what they can do.
- [Tutorial: Create and format an Excel table](https://learn.microsoft.com/en-us/office/dev/scripts/tutorials/excel-tutorial) — a genuinely beginner-friendly walkthrough of the Action Recorder and Code Editor.

---

## The workbook you need

Before running any script, your workbook needs, at minimum, these sheets. **Names matter** — the scripts look sheets up by name, so if yours are named differently, you'll need to either rename your sheets or update the script's sheet-name constants (see [Adapting this repository](#adapting-this-repository-to-your-own-workbook)).

| Sheet | What it's for | Used by |
| --- | --- | --- |
| `Sample Data` | Raw survey export — one row per feedback response. This is the input every script ultimately reads from. | All three scripts |
| `Attendance Record` | A register grid: one row per trainee (Site, Specialty, CS Trainer, Inner/Outer track, Name, Email), one column per dated session, marked `Y` for attended. | `attendance.ts` |
| `Session Feedback` | The output of `sessionfeedback.ts` — one row per session, with response counts, percentages, and comments. | Written by `sessionfeedback.ts`; read by `attendance.ts` |
| `Letter creation` | A small input form (presenter name, date, title) used to generate one presenter's letter and certificate. | `lettergeneration.ts` |
| `Presenter CRM Dashboard` | A manually-maintained tracker of presenter relationships (last contact, outcome, follow-up). Not touched by any script — a purely human-maintained reference sheet. | Manual only, no script |

If your workbook is missing one of the first four sheets, create it — even as a blank tab with just a header row — before running the matching script. Most scripts expect the target sheet to already exist so they can write into it.

---

## Getting your Microsoft Forms responses into Excel

The scripts assume your feedback survey answers already live in an Excel sheet, one row per response, one column per question — which is exactly what Microsoft Forms produces automatically.

1. Open your survey in **Microsoft Forms** and go to the **Responses** tab.
2. Select **Open results in Excel** (older versions may say **Open in Excel**). This creates a workbook, stored in OneDrive or SharePoint, containing your response data.
3. In that exported sheet, the first few columns are always **respondent ID, start time, completion time, name, and email**, followed by one column per question — this is why scripts like `attendance.ts` reference a completion-time column and an email column by a fixed position near the start of the sheet.
4. Either keep working directly in that generated workbook, or copy the response sheet into your main teaching-programme workbook and name it `Sample Data`.
5. **Write down your column headings, in order.** You'll need this list when asking an AI assistant to adapt a script (see the next section), since the scripts refer to specific columns by position.
6. If you selected **Open results in Excel**, new responses sync automatically each time you reopen the workbook in Excel for the web. If you downloaded a static copy instead, you'll need to re-export periodically to pick up new responses.

---

## Running your first Office Script

This is the core loop you'll repeat for every script in this repo. The first time through, follow every step closely.

1. Open your workbook in **Excel for the web** (Excel Online) — Office Scripts run there fully. They also work in the Microsoft 365 desktop app on Windows, though some tenants only enable the **Automate** tab in the web version.
2. Along the top ribbon, click the **Automate** tab. If you don't see it, see [Common problems and fixes](#common-problems-and-fixes).
3. Click **New Script**. Excel opens a code editor with a small starter script already inside.
4. **Select all of the starter code and delete it.**
5. Paste in the full contents of the `.ts` file you want to run (e.g. `sessionfeedback.ts`).
6. Near the top of the script, find the block of constants (`SAMPLE_DATA_SHEET`, `OUTPUT_SHEET`, and similar). Check every sheet name against your actual tab names, and check every column index against your actual columns — remember, **zero-based** (see [Key concepts](#before-you-start-key-concepts)).
7. Click **Save**, and give the script a name you'll recognise later (e.g. "Session Feedback Summary").
8. Click **Run**.
9. The first time you run any script, Excel may ask you to **approve permissions** — accept these, since the script needs to read and write cells in your workbook.
10. Once it finishes, open the sheet the script writes to (e.g. `Session Feedback`) and check the output looks right.

If something goes wrong, don't worry — the [debugging workflow](#debugging-scripts-iteratively) below turns error messages into fixes, usually with an AI assistant doing most of the work.

---

## Using Copilot, Claude, ChatGPT, or another AI to adapt these scripts

The example scripts in this repo were built around one specific workbook. Yours will almost certainly look at least a little different — different sheet names, a different column order, different wording on your Microsoft Form. Rather than editing TypeScript by hand, the intended way to adapt these scripts is to **describe your workbook to an AI assistant** and ask it to adjust the script for you.

> [!TIP]
> Excel itself is also rolling out a **"Draft a script with AI"** option directly in the Automate tab (currently a preview feature, so it may not be available on every account yet). It's a useful starting point for a brand-new script, but for adapting the more involved scripts in this repo — with their multi-sheet cross-referencing and email payloads — a full conversation with an AI chat assistant, as described below, tends to give you more control.

**What to give the AI assistant.** For a good result, provide:
- The script you're adapting — the full `.ts` file, or the specific function you're changing.
- The names of your workbook's sheets.
- The column headings from your Microsoft Forms response sheet, in order.
- A short description of what you want the output to look like.
- Any error message Excel gives you, copied exactly, including the line number if one is shown.

Be explicit that the AI should **adapt the script to your workbook**, not assume your workbook matches the example in this repo. It's easy for an AI assistant to quietly assume your sheet is laid out exactly like the original — telling it your actual sheet names and column order up front avoids that.

**Recommended prompt template.** Something like this works well as a starting point — fill in your own details in the brackets, paste in the relevant script from this repo, and send it to your AI assistant of choice:

```
I have an Excel workbook with these sheets: [list your sheet names].
My Microsoft Forms responses are in a sheet called [sheet name].
These are the column headings, in order: [paste column headings].
I want an Office Script that: [describe what you want it to do].

Please adapt the following script to my workbook, and tell me exactly
which constants near the top I need to change:

[paste the script]
```

---

## Debugging scripts iteratively

You will very rarely get a script working perfectly on the first try — that's normal, and it's the reason this workflow exists. It's a short loop:

1. Ask your AI assistant for a first version of the script, or start from an existing one in this repo.
2. Paste it into Excel's **Automate → New Script** editor.
3. Click **Run**.
4. If it fails, copy the **exact error message and line number** Excel shows you.
5. Paste that error back to your AI assistant and ask it to explain what went wrong and fix it.
6. Paste the corrected script back into Excel and run it again.
7. Repeat until the output looks right.

Each pass through this loop tends to be quick. Most errors in practice come down to a sheet name that doesn't quite match, a column index that's off by one, or a date format that doesn't parse the way the script expects — all things an AI assistant can usually spot immediately once it sees the actual error text.

**Use the scripts' own built-in diagnostics before you go hunting.** `attendance.ts` and `lettergeneration.ts` both log detailed progress to the console as they run — open the **Output** pane at the bottom of the code editor after clicking Run to see it. Two things in there are especially useful:

- `lettergeneration.ts` logs a **column mapping verification** block that shows exactly which column letter and index each metric constant is reading, together with a short preview of that column's real header text. It's the fastest way to confirm a `COLS_*` constant is pointing at the question you think it is.
- `attendance.ts` logs a **cross-reference summary** at the end of every run — rows considered, exact-email matches, backup (parsed-name) matches, marks written, and separate counts for responses skipped because of no match, an ambiguous match, or an unparseable date. It's the fastest way to see *why* a particular trainee wasn't marked present, without reading the matching code at all.

---

## Your learning path through this repo

**Start small if you're new to Office Scripts.** Before touching the full attendance or letter-generation scripts, it's worth building some confidence with something tiny. Ask an AI assistant to help you write and run, in order:

1. A script that just reads one sheet and logs how many rows it has.
2. A script that writes a value into a specific cell on another sheet.
3. A script that counts how many rows meet a simple condition (e.g. "responses after a certain date").
4. A script that applies some basic formatting — a bold header, a colour fill — to a range.

Each of those is a five-minute script, and together they cover most of the basic moves the three real scripts in this repo make, just at a much smaller scale.

**Suggested order for this repo's scripts:**

1. Make sure your workbook has the sheets described in [The workbook you need](#the-workbook-you-need).
2. Run `sessionfeedback.ts` first, and check the `Session Feedback` output looks sensible.
3. Run `attendance.ts` next — it reads from `Session Feedback`, so it works best once that sheet is already populated.
4. Only once those two are working reliably, move on to `lettergeneration.ts` and the [optional email/certificate pipeline](#optional-advanced-setup-email-certificates-and-the-full-pipeline).

---

## Worked example: one session, start to finish

Here's an illustrative walk-through of the everyday loop, showing how the pieces connect:

1. A teaching session runs on a Thursday afternoon. Forty trainees attend and fill in the Microsoft Forms feedback survey afterwards.
2. Those 40 responses appear as 40 new rows in `Sample Data` — one row per trainee, each with a completion timestamp, ratings, and free-text comments.
3. You open the workbook and run `sessionfeedback.ts` from the **Automate** tab. It groups those 40 rows by session date, calculates the positive-response percentages and an average clarity rating, de-duplicates the free-text comments, and writes **one new row** into `Session Feedback` — preserving whatever hand-typed notes already existed in columns B and C for that date.
4. You then run `attendance.ts`. It works through the same 40 `Sample Data` rows, matches each respondent to a row in the `Attendance Record` register (by email first, then by parsing an NHS-style email address, then by name), and marks a green `"Y"` in that Thursday's column for everyone who responded. It also checks whether it's been at least three days since the last digest email — if so, it builds and sends one, including the new session's stats pulled from `Session Feedback`, plus a table of any trainees now on two or more consecutive missed sessions.
5. Separately, if a guest presenter spoke at that session, you fill in their name, the date, and the presentation title on the `Letter creation` sheet and run `lettergeneration.ts`. It filters `Sample Data` down to that one session date, calculates the headline percentages, and emails a structured payload that — via the [advanced pipeline](#optional-advanced-setup-email-certificates-and-the-full-pipeline) — turns into a PDF certificate and letter for the presenter.

That's the entire lifecycle of one teaching session: from form response, to summarised feedback, to marked attendance, to a team digest, to — where relevant — a presenter's certificate.

---

## What each script does (plain English)

| Script | In plain English |
| --- | --- |
| `sessionfeedback.ts` | Reads all the raw feedback responses and turns them into one tidy summary row per session — response count, percentages, comments. |
| `attendance.ts` | Works out who attended each session from the feedback responses, marks the register, and emails the team a digest that flags trainees missing sessions. |
| `lettergeneration.ts` | Builds a personalised thank-you letter and certificate for a guest presenter, based on that session's feedback. |

The full technical detail behind each of these is in [Technical reference](#technical-reference-the-scripts-in-detail).

---

## Adapting this repository to your own workbook

These scripts were built around one specific teaching-programme workbook. Treat them as **templates**, not a finished product — expect to change sheet names, column indexes, email addresses, thresholds, and letter wording to match your own setup, ideally with an AI assistant's help as described [above](#using-copilot-claude-chatgpt-or-another-ai-to-adapt-these-scripts).

**What to customise first** — the changes almost everyone needs to make:

- Sheet names (`SAMPLE_DATA_SHEET`, `OUTPUT_SHEET`, and similar constants).
- Column indexes that map onto your Microsoft Form's actual question order.
- Date columns and date format, if your form or region formats dates differently.
- Email addresses (`EMAIL_FROM`, `EMAIL_TO`, `EMAIL_REPLY_TO`).
- The attendance threshold and digest frequency, if two consecutive misses / a three-day gap doesn't suit your programme.
- Certificate and letter template links (`CERTIFICATE_DOC_URL`, `LETTER_DOC_URL`).
- Wording inside the generated letters — this lives in the Google Docs templates, not the script itself.

**What *not* to change unless you know what you're doing:**

- The `main(workbook: ExcelScript.Workbook)` function signature, including `async`/`await` where it's already used. Office Scripts looks specifically for a function named `main` taking this one parameter, and the two email-sending scripts need to stay `async` so they can wait for the email to actually finish sending before the script ends.
- The structure of the JSON payload sent to the Cloudflare Worker/Resend — the Google Apps Script on the other end expects a specific shape, and changing it without updating that script too will break the certificate/letter pipeline silently.
- The trainee-matching logic in `attendance.ts` (exact email → parsed NHS-style email → full name → surname → surname+firstname). This has several deliberate fallback steps to handle messy real-world data, and simplifying it may reduce match accuracy.
- The BST/DST-safe date-difference calculation behind `attendance.ts`'s digest gate — it deliberately builds calendar-day values from local date *components* via `Date.UTC()` rather than subtracting raw timestamps, specifically to avoid an off-by-one-day bug across UK clock changes. A naive simplification here will reintroduce that bug.

If you're not sure whether a change is safe, that's a good question to ask your AI assistant directly — paste in the function you're considering changing and ask what depends on it.

---

## Technical reference: the scripts in detail

*This section is the detailed, implementation-level reference — useful once you've got the basic loop working, or if you're already comfortable reading TypeScript.*

This repository is intentionally flat: everything lives in the root, with no subfolders, build config, or package-manager files to navigate.

| File | Role | Approx. size |
| --- | --- | --- |
| `sessionfeedback.ts` | Feedback summariser | ~800 lines |
| `attendance.ts` | Attendance marker + digest emailer | ~1,300 lines — the largest and most involved, since it combines matching, digest-building, and letter-drafting logic in one file |
| `lettergeneration.ts` | Letter/certificate generator | ~700 lines |
| `README.md` | This guide | — |
| `LICENSE` | The MIT License | — |

| File | Reads from | Writes to | Sends email? | Beginner description | When to run it |
| --- | --- | --- | --- | --- | --- |
| [`sessionfeedback.ts`](#sessionfeedbackts) | `Sample Data` | `Session Feedback` | No | Turns raw feedback rows into one summary row per session | After each session, once responses have come in |
| [`attendance.ts`](#attendancets) | `Attendance Record`, `Sample Data`, `Session Feedback` | `Attendance Record` (marks + digest state) | Yes — attendance/feedback digest to the team | Marks who attended and emails the team a digest flagging trainees missing sessions | After `sessionfeedback.ts`, whenever you want attendance updated |
| [`lettergeneration.ts`](#lettergenerationts) | `Letter creation`, `Sample Data` | `Letter creation` (form + status) | Yes — certificate/letter payload to the pipeline | Builds a thank-you letter and certificate for one guest presenter | Once per presenter, after filling in the `Letter creation` form |

### `sessionfeedback.ts`

**What it does.** Turns raw, one-row-per-response survey data into one summary row per session date.

For each date it: counts responses; de-duplicates free-text comments (dropping blanks, very short text, and cells that only contain a Likert label); computes the **positive %** (Agree/Strongly Agree/4/5) for topic relevance, speaker clarity, a combined figure, career interest, skills/knowledge, confidence, and satisfaction; averages a numeric clarity rating; drops sessions with fewer than `MIN_RESPONSES_TO_OUTPUT` responses; then writes and formats the output (header styling, %/date number formats, and red/amber/green conditional formatting).

**Preserves manual columns.** Columns **B** and **C** of `Session Feedback` are for hand-typed notes; the script reads the existing sheet and re-applies your B/C values to the matching date after rewriting. In practice, `attendance.ts` expects column B to hold the **facilitator/speaker name** and column C the **session title** — it pulls both into the digest email it builds, so it's worth typing them in consistently even though `sessionfeedback.ts` itself never reads them.

**Configure** (top of `main`): `SAMPLE_DATA_SHEET`, `OUTPUT_SHEET`, `MIN_RESPONSES_TO_OUTPUT`, and the zero-based column indexes (`START_TIME_COL`, `COLS_COMMENTS`, `COLS_TOPIC_RELEVANT`, `COLS_SPEAKER_CLEAR`, `COLS_CAREER_INTEREST`, `COLS_SKILLS_KNOWLEDGE`, `COLS_CONFIDENCE_MH`, `COLS_PROGRAMME_SATISFACTION`, `COLS_CLARITY_RATING`). Where a list has several columns, the **first non-blank** is used for metrics and **all** for comments (handles forms whose columns shifted between versions). Traffic-light thresholds live in `applyConditionalFormatting`.

### `attendance.ts`

**What it does.** Two jobs in one run:

1. **Marks attendance.** Cross-references each `Sample Data` response against the `Attendance Record` register. It matches a respondent to a trainee row by **exact email first**, then falls back to parsing `firstname.lastname[n]@nhs.net` and matching on full name → surname → surname+firstname. On a match it writes `"Y"` (green fill) into the column whose header date equals the response's completion date. Duplicate date columns and BST/DST date edge-cases are handled explicitly. This half of the script **always runs**, regardless of the digest gate below.

2. **Builds & emails a digest** (gated to send at most every ~3 calendar days; last-send date is stored in cell **Z100** and only updated on a successful send). The digest HTML includes: per-session attendance rate + survey metrics + comments (pulled from `Session Feedback`), a site breakdown, inner/outer-track breakdown, an **attendance-concerns table** for trainees on a run of **2+ consecutive missed sessions**, and **auto-drafted non-attendance letters** for concerned trainees who have an email on file.

**Configure** (top of `main`): the `EMAIL_*` constants ([see below](#configuration-reference)); the `Attendance Record` column map (`ATTENDANCE_FIRSTNAME_COL`, `…SURNAME_COL`, `…EMAIL_COL`, `DATE_COLUMNS_START`); the `Sample Data` columns (`SAMPLE_COMPLETION_TIME_COL`, `SAMPLE_EMAIL_COL`); and the `SF_*` indexes that map onto the `Session Feedback` sheet. The consecutive-miss threshold (`>= 2`) and digest gap (`< 3` days) are inline literals — see the [Maintenance](#maintenance) note about a misleading comment nearby.

### `lettergeneration.ts`

**What it does.** Generates a **teaching contribution letter** for a single presenter. You fill in a small form on the `Letter creation` sheet (the script rebuilds this form each run, preserving typed values):

- **C4** Presenter Name · **C5** Date of Presentation · **C6** Presentation Title

On run it filters `Sample Data` to that session date, computes the headline percentages (topic-relevant-and-clear, skills/knowledge, career interest), collects unique comments, and assembles a letter plus a machine-readable **`STRUCTURED_JSON`** block. It POSTs an HTML email (with that JSON embedded) to the Cloudflare Worker → Resend → the team inbox. A **Google Apps Script** watching that inbox parses the JSON, fills the **certificate** and **letter** Google Docs templates (keep the `{{Name}}` / `{{Date}}` placeholders), exports them as PDFs, and emails them out. The script writes a colour-coded status into **C8**: green **✅** on success (with a timestamp), orange **⚠️** for input problems (missing fields, unparseable date, no matching responses), red **❌** for a failed send.

**Configure** (top of `main`): the `EMAIL_*` constants; `CERTIFICATE_DOC_URL` and `LETTER_DOC_URL` (the two template docs, near the bottom of the file); and the `COLS_*` indexes for `Sample Data`. Letter wording, signatures and logos live in the **Google Docs templates**, not the script — edit a template and the next PDF updates automatically.

### Conventions

- One `main(workbook: ExcelScript.Workbook)` entry point per file — though the exact signature differs. `sessionfeedback.ts` is synchronous (`function main(...): void`), while `attendance.ts` and `lettergeneration.ts` are `async function main(...): Promise<void>`, since both need to `await` an email `fetch()` call before finishing.
- Sheet names and **zero-based** column indexes are named constants — always confirm yours.
- Output sheets are fully rewritten each run; anything that must survive (manual B/C columns, the Z100 digest date) is explicitly captured/restored.
- No secrets or personal data in source — recipients and endpoints are placeholder constants.
- A missing sheet is **not** a hard Excel error in any of the three scripts — each one logs a message (e.g. `"Source worksheet not found: Sample Data"`) to the console and returns quietly. Always check the Output pane, not just whether Excel showed a red banner — see [Common problems and fixes](#common-problems-and-fixes).

### Real column layout in the example workbook

Purely as a concrete illustration of the [zero-based indexing](#before-you-start-key-concepts) idea — **this is what one specific workbook looks like, not a spec you need to match:**

**`Attendance Record`:**

| Column | A | B | C | D | E | F | G | H | I onwards |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Index | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8+ |
| Holds | Site | Specialty | CS Trainer | CS Trainer email | Inner/Outer | First name | Surname | Email | One dated column per session |

**`Sample Data`** — only the columns the scripts actually reference: completion time in column **C** (index 2), respondent email in column **F** (index 5). The rest of the layout depends entirely on your Microsoft Form's own question order, which is exactly why the `COLS_*` constants exist rather than being hard-coded.

---

## Configuration reference

Both email-sending scripts share the same placeholder constants at the top of `main`. **Fill these in before use:**

| Constant | Replace with | Notes |
| --- | --- | --- |
| `EMAIL_WORKER_URL` | `https://YOUR-WORKER-SUBDOMAIN.workers.dev/` | Your Cloudflare Worker that proxies to Resend. |
| `EMAIL_FROM` | `onboarding@resend.dev` | Resend sandbox sender, or your verified domain. |
| `EMAIL_TO` | `your-inbox@example.com` | Inbox the Google Apps Script / team watches. |
| `EMAIL_REPLY_TO` | `your-reply-to@example.com` | Where replies should go. |
| `CERTIFICATE_DOC_URL` | `…/document/d/YOUR_CERTIFICATE_DOC_ID/edit` | *(lettergeneration.ts)* Certificate template. |
| `LETTER_DOC_URL` | `…/document/d/YOUR_LETTER_DOC_ID/edit` | *(lettergeneration.ts)* Letter template. |

The **Resend API key** is configured only as a **secret inside the Cloudflare Worker**, never in these files.

---

## Common problems and fixes

> [!IMPORTANT]
> Most of the failures below don't show up as a big red Excel error. A missing sheet or a bad constant usually just makes the script log a message and stop quietly. Always check the **Output** pane at the bottom of the code editor after clicking Run — that's where every `console.log` message, including the diagnostics mentioned in [Debugging scripts iteratively](#debugging-scripts-iteratively), actually appears.

| Problem | Likely cause | Fix |
| --- | --- | --- |
| I can't see the **Automate** tab | Your account/licence doesn't have Office Scripts enabled, or a desktop build without it turned on | Try Excel for the web first; ask your IT/Microsoft 365 admin to confirm Office Scripts is enabled for your tenant |
| Script runs, output looks empty, and nothing seems to happen | A `*_SHEET` constant doesn't match your actual tab name exactly, including capitalisation and spacing — the script logged this to the Output pane and stopped | Compare the constant to your tab name character-for-character, then update the constant |
| Numbers, percentages, or names come out wrong or blank | A column index constant points at the wrong column | Recount your columns using the [zero-based index table](#before-you-start-key-concepts), then update the constant |
| Dates don't match, or attendance isn't marked for someone who clearly responded | Your form or region uses a different date format, or a timezone/BST-DST edge case | Check how dates are actually stored in `Sample Data`; ask your AI assistant to adjust the date-parsing logic |
| Microsoft Forms changed its column order | Someone edited the form, which shifted which question sits in which column | Update the relevant `COLS_*` constants to match the new order |
| Script runs with no errors, but writes to the wrong place | A sheet name or column constant is technically valid but pointing at the wrong sheet/column | Double-check every constant at the top of the script against your workbook |
| Email-sending scripts fail immediately, every time | `EMAIL_WORKER_URL` still has its placeholder value, or the Worker isn't deployed | Set up the Worker as described in [Optional advanced setup](#optional-advanced-setup-email-certificates-and-the-full-pipeline); confirm the URL constant is correct |
| Email-sending scripts reach the Worker but still fail | Resend itself rejected the request | Check the Output pane for the HTTP status: **401** = API key invalid or revoked, **403** = sending domain/recipient not verified in Resend, **422** = payload failed validation, **429** = rate limit hit, **500** = Resend server error |

If none of these match what you're seeing, copy the exact error text and line number into your AI assistant along with the script — see [Debugging scripts iteratively](#debugging-scripts-iteratively).

<details>
<summary><strong>Real example from this project's history: what a column-index bug actually looks like</strong></summary>

At one point, the `SF_*` column constants that `attendance.ts` uses to read `Session Feedback` were off by two columns, because an earlier version didn't account for a facilitator-name column and a session-title column that sit early in that sheet. The symptoms were a good demonstration of why these constants matter:

- The "Responses" figure in the digest showed a session **title** instead of a number.
- The comments box showed a **name** instead of feedback text.
- "Topic Relevant" and "Speaker Clear" showed percentages like **1300%** and **400%** — because response and comment *counts* were being read out of the percentage columns and multiplied by 100.

If your own output ever shows text where you expect a number, or a percentage over 100%, that's a strong signal a `COLS_*` or `SF_*` constant has drifted from your actual sheet layout — recheck it against the [zero-based index table](#before-you-start-key-concepts).

</details>

---

## Maintenance

Microsoft Forms, and the workbook itself, will change over time — and when they do, these scripts can silently start reading the wrong column. Specifically:

- If the **Microsoft Form is edited** (a question added, removed, or reordered), the columns in `Sample Data` shift, and every `COLS_*` / `*_COL` constant that references a position may need updating.
- If a **sheet is renamed**, the corresponding `*_SHEET` constant needs updating to match.
- If your **attendance or digest policy changes** (for example, flagging 3 misses instead of 2, or sending digests weekly instead of every 3 days), the relevant inline values in `attendance.ts` need updating.

A reasonable habit: after any change to the Microsoft Form, re-run each script once against a small or test date range and check the output before trusting it on real data.

> [!NOTE]
> A code comment above the digest-gating logic in `attendance.ts` refers to a "7-day gate", but the actual check in the code is `daysDiff < 3` — a 3-day gate, matching this guide and the constant described above. If you ever want to change the digest frequency, search for `daysDiff < 3` rather than trusting that particular comment. Either way, attendance marking itself is never affected by this gate — it runs on every execution; only the digest email is held back.

---

## Data protection and safe use with AI tools

This workbook can contain **trainee names, NHS email addresses, attendance records, and free-text feedback** — real personal data. Keep that in mind whenever you're getting AI help with these scripts:

- **Remove or anonymise personal data** before pasting workbook screenshots, sample rows, or exported data into an AI chat tool, unless your organisation has specifically approved that tool for that kind of data.
- You can usually get all the help you need using just **column headings and sheet structure** — you rarely need to share real rows of data to get a script adapted or debugged.
- **Never paste API keys, secrets, or live credentials** into a script or into an AI chat. The Resend API key in particular should only ever live inside the Cloudflare Worker's secret storage, never in a script file or a chat message.
- If you're unsure whether a particular AI tool is approved for use with patient- or trainee-adjacent data at your organisation, check with your local information governance team before sharing anything beyond structure and headings.

---

## Optional advanced setup: email, certificates and the full pipeline

*Skip this section entirely if you only want the Excel side working — `sessionfeedback.ts`, and the attendance-marking half of `attendance.ts`, need none of this.*

<details>
<summary><strong>Click to expand: the email, certificate, and letter pipeline</strong></summary>

Two scripts — `attendance.ts` (for its digest email) and `lettergeneration.ts` (for certificates and letters) — send an email as part of what they do. Here's the full path that email takes:

```
                    ┌────────────────────────── EXCEL WORKBOOK ──────────────────────────┐
  Microsoft Form    │                                                                     │
  (feedback survey) │   Sample Data ──► sessionfeedback.ts ──► Session Feedback           │
        │           │   (raw responses)   (aggregate by date)   (1 row per session)       │
        └──────────►│                                                  │                  │
                    │   Attendance Record ◄── attendance.ts ───────────┘                  │
                    │   (register grid)       (mark "Y", build digest)                     │
                    │                                                                      │
                    │   Letter creation ──► lettergeneration.ts                            │
                    │   (presenter form)     (build per-session letter)                    │
                    └───────────────────────────────│─────────────────│──────────────────┘
                                                     │ HTTPS POST      │ HTTPS POST
                                                     ▼                 ▼
                                        ┌──────────────────────────────────────┐
                                        │  Cloudflare Worker (tunnel/proxy)     │
                                        │  holds the Resend API key as a secret │
                                        └──────────────────┬───────────────────┘
                                                           │ Resend API
                                                           ▼
                                        ┌──────────────────────────────────────┐
                                        │   Email inbox (Gmail)                 │
                                        └──────────────────┬───────────────────┘
                                                           │ on-receive trigger
                                                           ▼
                                        ┌──────────────────────────────────────┐
                                        │  Google Apps Script                   │
                                        │  • reads the JSON payload in the email│
                                        │  • fills Google Docs templates        │
                                        │  • exports certificate + letter PDFs  │
                                        │  • emails the presenter the PDFs      │
                                        │  • emails the team a readable summary │
                                        └──────────────────────────────────────┘
```

**Why a Cloudflare Worker?** Office Scripts can make outbound `fetch` calls but cannot safely hold an API key (the script is visible to every workbook user). So each script POSTs its email payload to a small Cloudflare Worker; the Worker stores the **Resend** API key as a secret and forwards the request to Resend. The Worker URL is the only endpoint the scripts know about.

> [!NOTE]
> The Cloudflare Worker, Resend account, Google Apps Script, and Google Docs templates described here are **not** part of this GitHub repository — they're separate pieces you create and host yourself. This repo contains only the three Office Scripts and this documentation; searching it for Worker or Apps Script source won't turn up anything.

Setting this up requires:

1. A **Cloudflare Worker** you control, configured with a Resend API key as a secret, that forwards posted email payloads to Resend's API.
2. A **Resend** account and sending domain (or their sandbox sender, for testing).
3. A **Google Apps Script** bound to the Gmail inbox that receives the pipeline's emails — set up to parse the embedded JSON payload, fill Google Docs templates, export PDFs, and send them on.
4. Two **Google Docs templates** (certificate and letter) with `{{Name}}` / `{{Date}}`-style placeholders that the Apps Script fills in.

This is genuinely optional. Many programmes will be well served by just the Excel-side feedback and attendance automation, with certificates and letters handled manually — you can always add this pipeline later.

</details>

---

## Glossary

| Term | Meaning |
| --- | --- |
| **Office Script** | A TypeScript program that runs inside Excel (via the Automate tab) and can read/write the workbook automatically. |
| **TypeScript** | A structured, typed relative of JavaScript; the language Office Scripts are written in. |
| **Workbook** | An entire Excel file. |
| **Worksheet / sheet** | One tab within a workbook. |
| **Microsoft Forms response sheet** | The Excel sheet Microsoft Forms automatically populates, one row per submitted response — see [Getting your Microsoft Forms responses into Excel](#getting-your-microsoft-forms-responses-into-excel). |
| **API** | Application Programming Interface — a defined way for one piece of software to ask another to do something, e.g. "send this email." |
| **Cloudflare Worker** | A small piece of code Cloudflare runs on your behalf; used here to safely hold a secret API key and forward requests to Resend. |
| **Resend** | An email-sending API/service used by the pipeline to actually deliver emails. |
| **Google Apps Script** | Google's scripting platform; used here to watch a Gmail inbox, read the JSON payload in incoming emails, and produce PDF certificates/letters from Google Docs templates. |
| **JSON payload** | A structured block of text (JavaScript Object Notation) that carries data — here, the details needed to fill in a certificate or letter. |
| **Constant** | A named value near the top of a script (e.g. `SAMPLE_DATA_SHEET`) that you're expected to check, and change if needed, to match your workbook. |
| **Column index** | The zero-based number a script uses to refer to a column (A = 0, B = 1, and so on). |
| **Output pane / console log** | The panel at the bottom of the code editor where `console.log` messages appear — both progress updates and the quiet "sheet not found"-style messages a script logs instead of throwing a visible error. Your first stop when debugging. |

---

## License

Released under the [MIT License](LICENSE). The data workbook and any real trainee
information are **not** part of this repository and are not covered by it.
