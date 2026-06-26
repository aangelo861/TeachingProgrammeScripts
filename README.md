# Foundation Year Psychiatry Teaching Programme — Excel Automation

A set of **[Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel)** (TypeScript that runs *inside* Excel) that automate the admin behind a clinical teaching programme: aggregating session feedback, marking attendance, flagging trainees with poor attendance, and emailing certificates + digests.

These scripts target the **`ExcelScript` API** and run from the **Automate** tab in Excel for the web / Microsoft 365. They are **not** Node.js — you cannot run them with `ts-node`/`tsc`; they execute against a live workbook.


---

## The big picture — how data flows

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

---

## How to run an Office Script

1. Open the workbook in **Excel for the web** or **Microsoft 365 desktop**.
2. **Automate** tab → **New Script** → paste the `.ts` file (replace the default `main`).
3. Fill in the [Configuration](#configuration) constants and confirm the sheet names / column indexes match your workbook.
4. Click **Run**. The logic lives in `main(workbook)`, which Office Scripts calls automatically.

---

## Scripts

| File | Reads | Writes | Sends email? |
| --- | --- | --- | --- |
| [`sessionfeedback.ts`](#sessionfeedbackts) | `Sample Data` | `Session Feedback` | No |
| [`attendance.ts`](#attendancets) | `Attendance Record`, `Sample Data`, `Session Feedback` | `Attendance Record` (marks + digest state) | Yes — attendance/feedback **digest** to the team |
| [`lettergeneration.ts`](#lettergenerationts) | `Letter creation`, `Sample Data` | `Letter creation` (form + status) | Yes — **certificate/letter** payload to the pipeline |

---

### `sessionfeedback.ts`

**What it does.** Turns raw, one-row-per-response survey data into one summary row per session date.

For each date it: counts responses; de-duplicates free-text comments (dropping blanks, very short text, and cells that only contain a Likert label); computes the **positive %** (Agree/Strongly Agree/4/5) for topic relevance, speaker clarity, a combined figure, career interest, skills/knowledge, confidence, and satisfaction; averages a numeric clarity rating; drops sessions with fewer than `MIN_RESPONSES_TO_OUTPUT` responses; then writes and formats the output (header styling, %/date number formats, and red/amber/green conditional formatting).

**Preserves manual columns.** Columns **B** and **C** of `Session Feedback` are for hand-typed notes; the script reads the existing sheet and re-applies your B/C values to the matching date after rewriting.

**Configure** (top of `main`): `SAMPLE_DATA_SHEET`, `OUTPUT_SHEET`, `MIN_RESPONSES_TO_OUTPUT`, and the zero-based column indexes (`START_TIME_COL`, `COLS_COMMENTS`, `COLS_TOPIC_RELEVANT`, `COLS_SPEAKER_CLEAR`, `COLS_CAREER_INTEREST`, `COLS_SKILLS_KNOWLEDGE`, `COLS_CONFIDENCE_MH`, `COLS_PROGRAMME_SATISFACTION`, `COLS_CLARITY_RATING`). Where a list has several columns, the **first non-blank** is used for metrics and **all** for comments (handles forms whose columns shifted between versions). Traffic-light thresholds live in `applyConditionalFormatting`.

---

### `attendance.ts`

**What it does.** Two jobs in one run:

1. **Marks attendance.** Cross-references each `Sample Data` response against the `Attendance Record` register. It matches a respondent to a trainee row by **exact email first**, then falls back to parsing `firstname.lastname[n]@nhs.net` and matching on full name → surname → surname+firstname. On a match it writes `"Y"` (green fill) into the column whose header date equals the response's completion date. Duplicate date columns and BST/DST date edge-cases are handled explicitly.

2. **Builds & emails a digest** (gated to send at most every ~3 calendar days; last-send date is stored in cell **Z100** and only updated on a successful send). The digest HTML includes: per-session attendance rate + survey metrics + comments (pulled from `Session Feedback`), a site breakdown, inner/outer-track breakdown, an **attendance-concerns table** for trainees on a run of **2+ consecutive missed sessions**, and **auto-drafted non-attendance letters** for concerned trainees who have an email on file.

**Configure** (top of `main`): the `EMAIL_*` constants ([see below](#configuration)); the `Attendance Record` column map (`ATTENDANCE_FIRSTNAME_COL`, `…SURNAME_COL`, `…EMAIL_COL`, `DATE_COLUMNS_START`); the `Sample Data` columns (`SAMPLE_COMPLETION_TIME_COL`, `SAMPLE_EMAIL_COL`); and the `SF_*` indexes that map onto the `Session Feedback` sheet. The consecutive-miss threshold (`>= 2`) and digest gap (`< 3` days) are inline literals.

---

### `lettergeneration.ts`

**What it does.** Generates a **teaching contribution letter** for a single presenter. You fill in a small form on the `Letter creation` sheet (the script rebuilds this form each run, preserving typed values):

- **C4** Presenter Name · **C5** Date of Presentation · **C6** Presentation Title

On run it filters `Sample Data` to that session date, computes the headline percentages (topic-relevant-and-clear, skills/knowledge, career interest), collects unique comments, and assembles a letter plus a machine-readable **`STRUCTURED_JSON`** block. It POSTs an HTML email (with that JSON embedded) to the Cloudflare Worker → Resend → the team inbox. A **Google Apps Script** watching that inbox parses the JSON, fills the **certificate** and **letter** Google Docs templates (keep the `{{Name}}` / `{{Date}}` placeholders), exports them as PDFs, and emails them out. The script writes a ✅/⚠️/❌ status into **C8**.

**Configure** (top of `main`): the `EMAIL_*` constants; `CERTIFICATE_DOC_URL` and `LETTER_DOC_URL` (the two template docs, near the bottom of the file); and the `COLS_*` indexes for `Sample Data`. Letter wording, signatures and logos live in the **Google Docs templates**, not the script — edit a template and the next PDF updates automatically.

---

## Configuration

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

## Workbook structure (reference)

The scripts expect these sheets (others in the workbook are scratch/archive):

| Sheet | Role |
| --- | --- |
| `Sample Data` | Raw survey export — one row per feedback response. Source for all three scripts. |
| `Attendance Record` | Register grid: trainee rows (Site, Specialty, CS Trainer, Inner/Outer, Name, Email) × dated session columns marked `Y`. |
| `Session Feedback` | Output of `sessionfeedback.ts` — one row per session with metrics. |
| `Letter creation` | Presenter input form for `lettergeneration.ts`. |
| `Presenter CRM Dashboard` | Manually-maintained presenter relationship tracker (last contact, outcome, follow-up). |

---

## Conventions

- One `main(workbook: ExcelScript.Workbook)` entry point per file.
- Sheet names and **zero-based** column indexes are named constants — always confirm yours.
- Output sheets are fully rewritten each run; anything that must survive (manual B/C columns, the Z100 digest date) is explicitly captured/restored.
- No secrets or personal data in source — recipients and endpoints are placeholder constants.

---

## License

Released under the [MIT License](LICENSE). The data workbook and any real trainee
information are **not** part of this repository and are not covered by it.
