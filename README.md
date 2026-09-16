# Purchase Order Pipeline

An end-to-end automation pipeline that pulls purchase-order emails out of a Gmail
inbox, decrypts them through Microsoft's Office 365 Message Encryption (OME)
portal, extracts the PO data from the attached spreadsheet, and backfills
matching PO numbers into an Excel commission tracker — flagging duplicates and
data discrepancies along the way.

Built to replace a manual, error-prone daily process: someone previously had to
open each encrypted email by hand, wait for a one-time passcode, download the
attachment, and cross-reference it line-by-line against a 14,000+ row tracker.

## What it does

1. **Fetches** unread PO-request emails from Gmail over IMAP
2. **Decrypts** each one via Playwright, driving Microsoft's OME web portal
   through its email-based one-time-passcode flow
3. **Downloads and parses** the attached spreadsheet into structured records
4. **Merges** new records into two JSON files — `Fulfilled` (has a real PO
   number) and `Unfulfilled` (the source data had a note instead of a PO,
   e.g. a billing discrepancy)
5. **Backfills** the Excel tracker: writes in matching PO numbers, flags
   unresolved discrepancies with a highlighted row + comment, and flags
   duplicate PO numbers accidentally assigned to two different line items

## Why this was interesting to build

- **Gmail won't let a browser automate its own login** for an app-password
  account ("This browser or app may not be secure"), and IMAP app passwords
  aren't valid on the web login form either. Solution: drive Gmail over raw
  IMAP for reading mail, and use Playwright *only* for the one piece that
  actually requires a browser — Microsoft's OME decryption portal, which
  authenticates by emailing a one-time code rather than needing a Google
  session at all.
- **IMAP sequence numbers vs. UIDs** — an early version used sequence numbers
  (a message's position in the mailbox) under the assumption they were stable
  identifiers. They're not: mailbox activity between two runs can silently
  shift them, which caused an intermittent crash reading a *different*
  message than the one just searched for. Fixed by switching every
  search/fetch/store call to IMAP's real UID commands, which stay fixed for
  a message's lifetime, plus defensive handling so a malformed response is
  skipped instead of crashing the run.
- **Insertion-only side effects** — the tracker backfill only ever writes to
  a cell that was blank when the run started. It never rescans or modifies
  a row it didn't just fill, so re-running it is always safe and idempotent.

## Architecture

Everything is one class, `PurchaseOrderPipeline` (`po_pipeline.py`), split
into three independently-callable stages:

| Method | Responsibility |
|---|---|
| `fetch_emails()` | IMAP + OME decryption + spreadsheet parsing → returns new records |
| `merge_and_split(new_records)` | Merges into the two JSON files on disk, new data wins on conflicts |
| `backfill_tracker(fulfilled, unfulfilled, tracker_path)` | Writes PO numbers / flags into the Excel tracker |

`run(tracker_path)` chains all three for the full pipeline.

Two thin entry-point scripts sit on top:
- **`main.py`** — runs the whole pipeline (`pipeline.run(...)`)
- **`backfill_po.py`** — re-runs just the tracker backfill from whatever's
  currently in the JSON files, without touching email (useful for fast
  iteration on tracker logic)

## Setup

```bash
pip install -r requirements.txt
playwright install chromium
cp config.example.env .env   # then fill in your own values
python main.py
```

Required environment variables (see `config.example.env`):

| Variable | Description |
|---|---|
| `EMAIL_USER` | Gmail address the pipeline logs into |
| `GMAIL_APP_PASSWORD` | [App password](https://myaccount.google.com/apppasswords), not your regular password |
| `TRACKER_PATH` | Absolute path to the `.xlsm` tracker workbook |
| `SENDER_FILTER` | Only process emails from this sender |

## Stack

Python · Playwright · pandas · openpyxl · IMAP

---

*Note: `output/`, `downloads/`, and `debug/` are gitignored — they hold
generated data and downloaded attachments from real runs, not part of the
codebase itself.*
