# Onestop Line

Onestop Line is a simple web page that gives church members the phone number for Terry. Terry is a Telnyx phone voice assistant that answers read-only questions about the church's Onestop schedule (a shared Google Sheet). Callers can ask what is happening, when, and who is on a given role. Terry cannot change the schedule.

## Folders

- `web/` holds the Onestop Line page (`index.html`). It is a page body only; the hosting adds the doctype, head, and body tags.
- `refresh/` holds `build.py`, the script that builds Terry's instructions, and `head.txt`, the fixed opening part of those instructions.
- `skills/terry-onestop-refresh/` holds the Claude skill that runs the nightly refresh.

## How the nightly refresh works

A scheduled Claude task runs at midnight Eastern and follows the skill:

1. Download the Onestop sheet through the Google Drive connector.
2. Run `refresh/build.py` to build Terry's instructions from the current Monday to Sunday weekly tabs plus the calendar tab. The childcare column and childcare rows are removed.
3. Push only the instructions field to the Telnyx assistant. No other assistant settings are changed.

## Status

- Beta.
- Read-only: Terry answers questions but never edits the sheet.
- SMS (texting Terry) is planned.

## Privacy

Never commit sheet exports or generated instructions. They contain church members' names and schedules. The `.gitignore` blocks the common file types, but check `git status` before every commit.
