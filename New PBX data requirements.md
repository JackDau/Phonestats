# Phone Data Requirements: YOLO Phone System (Garema)

**Prepared for:** IT service provider supporting the YOLO phone system
**Site:** Garema (one call queue, no "call me back" option)
**Purpose:** We review our phone activity each week in an internal dashboard (call volumes, wait times, answered and missed calls, busy half-hour periods, and calls handled per staff member). To do this for Garema, we need a call detail export from YOLO containing the information below.

---

## 1. Must have: one row per call, with these fields

| # | Field | What we need | Why we need it |
|---|-------|--------------|----------------|
| 1 | **Call ID** | A unique identifier for each call. If a call is transferred, every part of it carries the same ID | Avoids double counting transferred calls |
| 2 | **Call start date and time** | When the call first reached YOLO, to the second, in Canberra local time (daylight saving handled) | Daily and weekly volumes, half-hourly heatmaps, busiest times |
| 3 | **Queue entry time** | When the call joined the Garema queue (after any phone menu or greeting) | Separates time in the menu from time waiting for staff |
| 4 | **Answer time** | When a staff member picked up (blank if never answered) | Calculates how long the patient waited |
| 5 | **Call end time** | When the call finished | Call length |
| 6 | **Wait time (seconds)** | Seconds from queue entry to answer, or to hang-up if never answered | Average and maximum wait per day |
| 7 | **Talk time (seconds)** | Seconds from answer to end of call | Average call length |
| 8 | **Total duration (seconds)** | Seconds from start to end | Overall call length |
| 9 | **Direction** | Inbound, Outbound or Internal | Separate figures for calls in and calls out; exclude internal calls |
| 10 | **Caller number** | The patient's phone number, or "Withheld" if hidden | Repeat callers, follow-up of missed calls |
| 11 | **Number dialled** | Which Garema number was called | Confirms the call belongs to Garema |
| 12 | **Queue name** | The queue the call reached (blank if it never reached the queue) | Separates calls that hung up in the menu |
| 13 | **Call outcome** | One clear outcome for every inbound call (see section 2) | Answered vs missed counts and percentages |
| 14 | **Answered by** | Staff member name **and** extension of the person who answered (or made an outbound call) | Calls taken per staff member, average time to pick up, average call length |

**Outbound calls** need the same core fields: Call ID, start time, answer time (when the other party picked up), end time, talk time, staff name and extension, and number dialled.

## 2. Call outcome codes

Every inbound call should carry exactly one outcome, with a written description of what each code means. At minimum we need to tell apart:

- **Answered** by a staff member
- **Abandoned in queue** (patient hung up while waiting), with how long they waited
- **Abandoned in menu / greeting** (patient hung up before reaching the queue)
- **Timed out / overflowed** to another number or voicemail
- **Voicemail left**
- **After hours / closed message** played
- **Transferred** (see section 3)

## 3. Transfers

When a call is transferred, we need either:

- one row per part of the call, each sharing the same Call ID, with its own start time, end time and staff member, **or**
- a single row showing the first staff member who answered, the final staff member, and the number of transfers.

Please tell us which approach YOLO uses.

## 4. Nice to have

| Field | Why |
|-------|-----|
| Hold time (seconds) | Separates time on hold from active talk time |
| Position in queue when the caller joined | Shows how busy the queue was at the time |
| Number of staff logged in to the queue at the time | Helps plan staffing per half hour |
| Call recording reference | Quality and complaint follow-up |

Note: we use this data for staffing and patient access decisions. Our reception staff look after patients at the desk as well as the phones, so we do **not** need "rang but not answered" counts per staff member or individual efficiency scores.

## 5. Format and delivery

1. **File type:** CSV or Excel (.xlsx), one row per call (or per part of a call, sharing a Call ID), with a header row.
2. **Consistent columns:** column names and order stay the same from week to week. Please tell us before any change.
3. **Everything in one file:** inbound, outbound and internal calls for the chosen date range in a single export.
4. **Times:** dates and times in one consistent format, preferably `DD/MM/YYYY HH:MM:SS` in Canberra local time. Durations as whole seconds (plain numbers) or `HH:MM:SS`.
5. **Delivery:** ideally an automatic weekly export (Monday to Sunday) saved to our SharePoint folder or emailed, plus the ability for us to run an export for any date range ourselves. API access is a welcome alternative.
6. **History:** at least 12 months of call data kept and exportable.
7. **Data dictionary:** a document explaining every column and every outcome code.
8. **Sample file:** a sample export from a real week before we sign off, so we can test it in our dashboard.
9. **Privacy:** caller numbers are personal information under the Privacy Act 1988. Please confirm where call data is stored (Australian hosting preferred) and who can access it.

### Suggested column names

Our dashboard recognises a range of column names, so YOLO's own names may well work as they are. If you are setting up a custom report, these names are guaranteed to be recognised:

| Column name | Example |
|-------------|---------|
| Call ID | `Y102938` |
| Start Time | `03/11/2026 09:05:10` |
| Answer Time | `03/11/2026 09:06:15` |
| End Time | `03/11/2026 09:09:15` |
| Wait Time | `65` |
| Talk Time | `180` |
| Duration | `245` |
| Direction | `Inbound` |
| Caller Number | `0412 345 678` |
| Dialled Number | `02 6xxx xxxx` |
| Queue | `Garema` |
| Status | `Answered` |
| Agent | `Sam Lee` |
| Extension | `201` |

## 6. Questions for the provider

1. Which of the fields above are available in a standard YOLO report, and which need custom reporting?
2. Can reports be scheduled and delivered automatically? To which destinations?
3. Is there an API for call records? Is there an extra cost?
4. How is a transferred call represented?
5. How long is call data kept?
6. Can you supply a sample export and a data dictionary?

---

## Current problems with the reports supplied so far

The reports provided to date do not show:

- when each call started
- when each call ended
- how long the patient waited before being answered

These items (fields 2, 4, 5 and 6 above) are the minimum we need to measure how easily patients can reach Garema by phone.
