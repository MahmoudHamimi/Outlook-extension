# Inbox Assistant for Outlook Web

A browser extension (Chrome/Edge, Manifest V3) for **outlook.office.com**,
**outlook.office365.com**, and **outlook.live.com**. A small floating "IA"
button appears in the bottom-right corner of the inbox; click it to open the
panel with all nine features:

1. **Stale-email signals (automatic)** — continuously scans the visible message list and stamps a `⏰ Nd` badge directly onto any unread row that's sat unopened longer than the threshold (default 3 days, configurable in the "Stale" tab). No click required — it runs on its own via a `MutationObserver` plus a periodic re-check.
2. **Priority senders (automatic)** — add a sender's name or email in the "Priority" tab (or one-click "Add sender of open email"); any matching row in the message list automatically gets a `★ Priority` badge and a red side-accent, visible at a glance.
3. **Folders / manual categorization** — create folder names and assign the currently open email to one; browse assigned emails grouped by folder, with one click to reopen or remove any of them. This is a **client-side organizational layer stored by the extension** — it doesn't move messages via Outlook's own folder system (that needs Microsoft Graph API + OAuth, out of scope for a pure content-script extension) but works entirely offline/standalone.
4. **Follow-up tracker** — track the open email with an N-day reminder; overdue ones float to the top of the "Tracked" list.
5. **Meeting-request extractor** — scans the open email for dates/times and opens a prefilled "Add to Calendar" tab.
6. **Reply-tone checker** — flags curt/formal/passive-aggressive phrasing, on demand and by intercepting the Send button.
7. **Meeting cost calculator** — attendee count x salary-band hourly rate x duration.
8. **PDF export** — opens a clean, printable version of the open email (now including the sender's name) and triggers the print dialog so you can "Save as PDF."

Only 4 files: `manifest.json`, `content.js`, `icon.png`, this README.

Everything is **100% client-side and standalone**: all data (stale-threshold
setting, priority senders, folders, follow-ups) lives in `chrome.storage.local`
on the local browser profile. The extension makes no network requests of its
own and needs no permissions beyond `storage` — it only reads/reacts to the
Outlook web page already loaded in the tab.

## Load it (unpacked)

1. Go to `chrome://extensions` (or `edge://extensions`).
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and select this folder.
4. Open outlook.office.com (or outlook.live.com) — you should see the round "IA" button appear bottom-right.

No build step, no bundler — it's loaded exactly as-is.

## Why no external libraries

Manifest V3 extensions aren't allowed to fetch and execute remotely-hosted
code (Chrome Web Store policy + CSP), so this doesn't pull in a date-parsing
library or a PDF library from a CDN the way the taskpane version could.
Practically that means:

- **Meeting extraction** uses a compact regex-based date/time parser instead of a full natural-language date library — it'll catch "Monday, June 10 at 2pm," "6/10 at 14:00," "tomorrow at 3pm," etc., but won't be as smart as something like chrono-node about ambiguous phrasing.
- **PDF export** uses the browser's native print dialog (`window.print()` → "Save as PDF") instead of a bundled PDF library — one extra click, zero dependencies.

## About the DOM selectors (read this if something stops working)

Outlook on the web is a React app with an internal, undocumented, and
occasionally-changing DOM. This extension deliberately avoids relying on
Outlook's generated CSS class names (those churn often) and instead hooks
into things that are far more stable:

- **Which email is open** — read from the URL itself (`/mail/.../id/<itemId>/...`), not the DOM.
- **The open email's body** — `[role="document"]`, an ARIA role Outlook uses consistently for the message content region.
- **The open email's subject** — a heading element Outlook marks with `class="allowTextSelection"`.
- **The open email's sender** — an element whose accessible label/tooltip starts with `"From:"` or contains an `@` address.
- **Compose body** — `[contenteditable="true"][role="textbox"]`.
- **Send button** — `button[aria-label="Send"]`.
- **Message-list rows** (for stale/priority badges) — `[role="option"]`/`[role="row"]` items inside the `[role="listbox"]` message list. Unread state, sender, and received date are read from each row's accessible label, native tooltip `title` attributes, and bold-text rendering rather than any Outlook-specific class name, since those are the only reasonably stable signals available.

All of these live in one place — the `CONFIG` object at the top of
`content.js` — specifically so that if Microsoft changes Outlook's markup
and something stops matching, there's exactly one spot to update rather than
selectors scattered through the file. If a feature suddenly stops finding
the email body/subject/compose box, that config block is where to look
first (open DevTools on outlook.office.com, inspect the element, and swap
in whatever selector still matches).

## Known limitations

- **Follow-up tracking is manual-confirmation based** — there's no way for a content script to know a reply arrived; you click "Mark replied" yourself. True auto-detection needs Microsoft Graph API access (OAuth + admin consent), which is out of scope for a browser extension injecting into the page.
- **The Send-button interception is best-effort.** Outlook web doesn't expose a "before send" event to page scripts, so this hooks the Send button's click handler. It works well in the common cases (main compose window, inline reply) but a UI change could require updating `sendButtonSelectors` in `CONFIG`.
- **Meeting cost calculator inputs are manual** — reliably reading the live attendee list out of Outlook's meeting-compose UI would require selectors likely to break often, so attendee count and duration are typed in rather than auto-detected.
- **Stale-email detection depends on Outlook exposing a parseable received date somewhere in a row** (an accessible label or a native tooltip `title`). If a row has neither, it's silently skipped rather than guessed — you'll see a lower "flagged" count than the true number of stale unread emails in that case.
- **Unread detection uses OWA's bold-text convention as a fallback signal**, which is reliable in the default theme but could misfire under a heavily customized Outlook theme; the badge is additive/visual only and never changes anything about the email itself.
- **Folders are local labels, not real Outlook folders** — moving actual messages between Outlook folders requires Graph API/EWS permissions this extension intentionally doesn't request, to stay a pure client-side content script.
- Storage is local to the browser profile (`chrome.storage.local`) — tracked follow-ups, priority senders, and folders won't sync to another device unless you're signed into Chrome sync with extension data sync enabled.
