# EcoOils Project - Site Walkthrough Close-Out

HSSE logs a walkthrough finding with a BEFORE photo. The company's HSSE officer gets an email with the photo and a close-out link, adds the AFTER photo, and HSSE verifies. A daily digest chases open and overdue items. A pictorial before/after report can be printed or saved as PDF.

Files: index.html (app), Code.gs (Apps Script backend, line 1 is the only version), sw.js, wordmark.png, appsscript.json.
Follows EcoOils-Project-Common/RULES.md.

## Mail relay (v1.2.0)

Emails can be sent from a second Google account so the main account's daily quota is not used.
`MailRelay.gs` is deployed as its own Apps Script web app under the sending account.
In the main script set Script Properties `RELAY_URL` (the relay /exec URL) and `RELAY_SECRET`
(logged by the relay's `setup()`). Run `testRelay()` in the main script to check.
Optional: `RELAY_FALLBACK = 1` sends from the main account if the relay is down (default: off).
