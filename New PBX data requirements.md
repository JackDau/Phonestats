# Phone Data Requirements for the New PBX

**Prepared for:** New IT service provider
**Purpose:** We analyse our phone activity each week in an internal dashboard (call volumes, wait times, answered and missed calls, busy periods, and per-queue results such as the Canberra Vasectomy line). To keep doing this on the new phone system, we need a call detail export containing the information below.

---

## 1. Must have: one row per call, with these fields

| # | Field | What we need | Why we need it |
|---|-------|--------------|----------------|
| 1 | **Call ID** | A unique identifier for each call, the same ID used across every report (call records, queue reports, transfers) | Lets us link a call's queue activity, answer details and transfers together without double counting |
| 2 | **Call start date and time** | When the call first reached our phone system, to the second, in Canberra local time (daylight saving handled), format `DD/MM/YYYY HH:MM:SS` | Daily and weekly volumes, half-hourly heatmaps, busiest times |
| 3 | **Queue entry time** | When the call joined a queue (after any phone menu) | Separates time spent in the menu from time spent waiting for staff |
| 4 | **Answer time** | When a staff member picked up (blank if never answered) | Calculates how long the patient waited |
| 5 | **Call end time** | When the call finished | Call length and completeness checks |
| 6 | **Wait time (seconds)** | Seconds from queue entry to answer, or to hang-up if never answered | Average and maximum wait time per day and per queue |
| 7 | **Talk time (seconds)** | Seconds from answer to end of call | Average call length |
| 8 | **Total duration (seconds)** | Seconds from start to end | Overall call length |
| 9 | **Direction** | Inbound, Outbound or Internal | Separate heatmaps for calls in and calls out; exclude internal calls |
| 10 | **Caller number** | The patient's phone number, or "Withheld" if hidden | Repeat caller analysis, callback lists for missed calls |
| 11 | **Number dialled** | Which of our numbers was called (main line, vasectomy line, health professionals line, fax, etc.) | Reporting per line |
| 12 | **Queue name** | Which queue the call went to (e.g. Appointments, General Enquiries, Health Professionals, Canberra Vasectomy) | Per-queue reporting |
| 13 | **Phone menu option chosen** | The menu key or option the caller selected, if a menu is used | Shows where callers are going and where they drop off |
| 14 | **Call outcome** | A clear outcome code for every inbound call (see section 2) | Answered vs missed counts and percentages |
| 15 | **Answered by** | Staff member name **and** extension of the person who answered | Per-staff calls taken, average time to pick up, average call length |
| 16 | **Site / location** | Which clinic or handset the call was answered at, if available | Separates sites and lets us exclude internal handsets |

**Outbound calls** need the same core fields: Call ID, start time, answer time (when the other party picked up), end time, talk time, staff name and extension, and number dialled.

## 2. Call outcome codes

Every inbound call should carry exactly one outcome, with a written description of what each code means. At minimum we need to tell apart:

- **Answered** by a staff member
- **Abandoned in queue** (patient hung up while waiting), with how long they waited
- **Abandoned in menu** (patient hung up before reaching a queue)
- **Timed out / overflowed** to another queue, number or voicemail
- **Voicemail left** (and which mailbox)
- **After hours / closed message** played
- **Callback requested** (if the system offers callbacks), plus when the callback was made and by whom
- **Transferred** (see section 3)

## 3. Transfers and multiple call legs

When a call is transferred or passes through more than one queue, we need either:

- one row per leg, each sharing the same Call ID, with the start time, end time, queue and staff member for that leg, **or**
- a single row with the first answering staff member, the final staff member, the number of transfers and total hold time.

Please tell us which approach the system uses.

## 4. Nice to have

| Field | Why |
|-------|-----|
| Hold time (seconds) | Separates time on hold from active talk time |
| Position in queue when the caller joined | Shows how busy the queue was at the time |
| Number of staff logged in to the queue at the time | Helps with staffing levels per half hour |
| Call recording reference | Quality and complaint follow-up |
| Call quality score (MOS) | Diagnosing audio complaints |

Note: we intend to use this data for staffing and patient access decisions. Our reception staff handle in-person patients as well as phones, so we do **not** need "ring no answer" counts per staff member or individual efficiency scores.

## 5. Format and delivery

1. **File type:** CSV or Excel (.xlsx), one row per call (or per leg, sharing a Call ID), with a header row.
2. **Consistent columns:** column names and order stay the same from week to week. Please notify us before any change.
3. **All queues in one file:** a single export covering every queue and direction for the chosen date range, rather than one file per queue.
4. **Times:** durations as whole seconds (plain numbers); dates and times in one consistent format, `DD/MM/YYYY HH:MM:SS`, Canberra local time.
5. **Delivery:** ideally an automatic weekly export (Monday to Sunday) saved to our SharePoint folder or emailed, plus the ability for us to run an export for any date range ourselves. API access is a welcome alternative.
6. **History:** at least 12 months of call data retained and exportable.
7. **Data dictionary:** a document explaining every column and every outcome code.
8. **Sample file:** a sample export from a real week before we sign off, so we can test it in our dashboard.
9. **Privacy:** caller numbers are personal information under the Privacy Act 1988. Please confirm where call data is stored (Australian hosting preferred) and who can access it.

## 6. Questions for the provider

1. What is the name and version of the PBX / phone platform?
2. Which of the fields above are available today in a standard report, and which need custom reporting?
3. Can reports be scheduled and delivered automatically? To which destinations?
4. Is there an API for call records? Is there an extra cost?
5. How is a call that passes through several queues or is transferred represented?
6. How long is call data kept?
7. Can you supply a sample export and a data dictionary?

---

## Current problems with the reports supplied so far

The reports provided to date do not show:

- when each call started
- when each call ended
- how long the patient waited before being answered

These three items (fields 2, 4, 5 and 6 above) are the minimum we need to measure patient access by phone.
