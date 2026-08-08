# Staff Folder Manager

A small privacy-conscious manager dashboard for tracking staff-folder checklists and renewal dates (MOT, insurance, DBS review and other document review).

## Status and alert rules
- Green means acceptable.
- Red means information is missing/not confirmed, or a required date is already out of date.
- The daily alert includes red cells and expired dates only. Dates that have not passed never create an alert.
- `daily_staff_folder_alert.py` reads `latest-staff-folder.xlsx` and produces the privacy-minimised daily WhatsApp message body. It contains only staff initials/references and actions, never document scans or sensitive identifiers.

## v1 limitations
- Dashboard data lives only in the browser on the device where it is entered; export a JSON backup regularly.
- Device notifications are only checked while the app is open.
- An unattended daily WhatsApp delivery must be scheduled only after the dedicated Business account’s target group and outbound path have been safely verified.
- Do not store document scans, certificate numbers or other sensitive employment data in this app.

## Development
Open `index.html` via a static HTTP server (not `file://`) to test the PWA service worker.
