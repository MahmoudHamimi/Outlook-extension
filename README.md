# Inbox Assistant for Outlook Web

A browser extension (Chrome/Edge, Manifest V3) for **outlook.office.com**,
**outlook.office365.com**, and **outlook.live.com**. A small floating "IA"
button appears in the bottom-right corner of the inbox (with a red count
badge when something needs attention); click it to open the panel with all
seven tools:

1. **Stale-email signals (automatic)** — continuously scans the visible message list and stamps a `⏰ Nd` badge directly onto any unread row that's sat unopened longer than the threshold (default 3 days, configurable in the "Stale" tab). No click required. Opening the email clears its badge right away — see "Recent fixes" below for why this used to lag.
2. **Priority senders (automatic)** — add a sender's name or email in the "Priority" tab (or one-click "Add sender of open email"); any matching row in the message list automatically gets a `★ Priority` badge and a red side-accent, visible at a glance.
3. **Follow-up tracker** — track the open email with an N-day reminder; overdue ones float to the top of the "Tracked" list and count toward the badge on the floating IA button.
4. **Contacts insights** — pure-SVG donut charts showing who you receive email from most and who you send email to most, built entirely from whatever messages are currently loaded in the list (see "About the Contacts tab" below).
5. **Meeting cost calculator** — attendee count × salary-band hourly rate × duration, now with a currency picker, a recurrence dropdown that projects an annual cost (weekly/bi-weekly/monthly), a live stopwatch that ticks the running cost of a meeting in real time, and a saved history of past calculations.
6. **Multi-format export** — PDF (via the print dialog), plus one-click HTML, Markdown, and plain-text file downloads of the open email, sender included.
7. **Schedule Send quick-picker** — one-click presets (Tomorrow 8am, Monday 9am, Friday EOD) that appear both in the panel and as small pills injected next to Outlook's own Send button. This one is best-effort/experimental — see "Known limitations" below.

Only 4 files: `manifest.json`, `content.js`, `icon.png`, this README.

Everything is **100% client-side and standalone**: all data (stale-threshold
setting, priority senders, follow-ups, contacts snapshots, cost history)
lives in `chrome.storage.local` on the local browser profile. The extension
makes no network requests of its own and needs no permissions beyond
`storage` — it only reads/reacts to the Outlook web page already loaded in
the tab.

## Removed in 1.2.0

The **meeting-request extractor**, **reply-tone checker**, and **manual
folders** tabs from 1.1.0 have been removed entirely (code, storage keys,
and UI) to make room for Contacts insights, the upgraded cost calculator,
multi-format export, and Schedule Send.

## Load it (unpacked)

1. Go to `chrome://extensions` (or `edge://extensions`).
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and select this folder.
4. Open outlook.office.com (or outlook.live.com) — you should see the round "IA" button appear bottom-right.

No build step, no bundler — it's loaded exactly as-is.

## Recent fixes

- **Stale badge never clearing, confirmed root cause.** `isRowUnread()` had a
  bold-text fallback ("if any text in this row renders bold, treat it as
  unread") for skins where Outlook doesn't expose a clean unread signal. That
  fallback was matching the sender's **avatar initials** (e.g. "MH", "A"),
  which Outlook renders bold/600-weight permanently, whether the email is
  read or not — so once that fallback fired, a row could never be detected
  as "read" again and its stale badge could never clear, no matter how many
  times you rescanned. Fixed two ways: (1) unread detection now checks
  Outlook's own per-row "Mark as read" / "Mark as unread" toggle-button
  title first — confirmed via diagnostic to be an unambiguous, always-
  accurate signal of the row's current state — before falling back to
  anything else; (2) the bold-text fallback (now last-resort only) explicitly
  skips avatar/initials elements so it can't misfire the same way again.
- **MutationObserver only watching for added/removed nodes.** Outlook marks a
  row "read" by changing an *attribute* on the existing row, not by swapping
  the node out, so the observer now also watches `aria-label`/`class`/`title`
  attribute changes, and clicking any row also triggers a same-frame
  optimistic badge removal plus a couple of staggered rescans (300ms/900ms/2s)
  as a fallback in case Outlook is slow to update.

## About the Contacts tab

There's no Microsoft Graph API access here (see "Why no external
libraries" / folders note below), so "who I hear from most" is built purely
from whatever message rows Outlook has currently rendered in the list —
not your whole mailbox history. To get a fuller picture: open the folder
you care about (Inbox for "received from", Sent Items for "sent to"),
scroll down to let Outlook load more rows, then click "Scan this list".
Each scan replaces the previous snapshot for that folder (it doesn't
accumulate across scans, to avoid double-counting rows that re-render
during virtualized scrolling). The donut charts and legends are built with
plain SVG (a `stroke-dasharray` ring trick) — no charting library.

## Why no external libraries

Manifest V3 extensions aren't allowed to fetch and execute remotely-hosted
code (Chrome Web Store policy + CSP), so this doesn't pull in a charting
library, a date-parsing library, or a PDF library from a CDN the way a
taskpane version could. Practically that means:

- **The Contacts donut charts** are hand-built SVG rings, not a charting library.
- **PDF export** uses the browser's native print dialog (`window.print()` → "Save as PDF") instead of a bundled PDF library — one extra click, zero dependencies. HTML/Markdown/plain-text export use `Blob` + a temporary download link instead.
- **Schedule Send** drives Outlook's own native "Send later" dialog rather than shipping a scheduling engine of its own — see "Known limitations".

## About the DOM selectors (read this if something stops working)

Outlook on the web is a React app with an internal, undocumented, and
occasionally-changing DOM. This extension deliberately avoids relying on
Outlook's generated CSS class names (those churn often) and instead hooks
into things that are far more stable:

- **Which email is open** — read from the URL itself (`/mail/.../id/<itemId>/...`), not the DOM.
- **The open email's body** — `[role="document"]`, an ARIA role Outlook uses consistently for the message content region.
- **The open email's subject** — a heading element Outlook marks with `class="allowTextSelection"`.
- **The open email's sender** — an element whose accessible label/tooltip starts with `"From:"` or contains an `@` address.
- **Send button** — `button[aria-label="Send"]`.
- **Message-list rows** (for stale/priority badges and Contacts scanning) — `[role="option"]`/`[role="row"]` items inside the `[role="listbox"]` message list. Unread state, sender/recipient, and received date are read from each row's accessible label, native tooltip `title` attributes, and bold-text rendering rather than any Outlook-specific class name, since those are the only reasonably stable signals available.
- **Sent vs. received folder detection** (for Contacts) — the URL path (looks for `/sentitems/`) with a fallback to the selected item in the folder nav.
- **Schedule Send's "more options" caret and "Send later" menu item** — the LEAST stable selectors in this file; see "Known limitations".

All of these live in one place — the `CONFIG` object at the top of
`content.js` — specifically so that if Microsoft changes Outlook's markup
and something stops matching, there's exactly one spot to update rather than
selectors scattered through the file. If a feature suddenly stops finding
the email body/subject/compose box, that config block is where to look
first (open DevTools on outlook.office.com, inspect the element, and swap
in whatever selector still matches).

## Known limitations

- **Follow-up tracking is manual-confirmation based** — there's no way for a content script to know a reply arrived; you click "Mark replied" yourself. True auto-detection needs Microsoft Graph API access (OAuth + admin consent), which is out of scope for a browser extension injecting into the page.
- **Contacts insights are a snapshot of the loaded list, not your full mailbox** — see "About the Contacts tab" above. It's an honest, standalone alternative to a Graph API mail-analytics call, not a replacement for one.
- **Meeting cost calculator inputs are manual** — reliably reading a live attendee list out of Outlook's meeting-compose UI would require selectors likely to break often, so attendee count and duration are typed in rather than auto-detected.
- **Stale-email detection depends on Outlook exposing a parseable received date somewhere in a row** (an accessible label or a native tooltip `title`). If a row has neither, it's silently skipped rather than guessed — you'll see a lower "flagged" count than the true number of stale unread emails in that case.
- **Unread detection uses OWA's bold-text convention as a fallback signal**, which is reliable in the default theme but could misfire under a heavily customized Outlook theme; the badge is additive/visual only and never changes anything about the email itself.
- **Schedule Send is best-effort/experimental.** Outlook doesn't expose a stable, documented way for a content script to schedule a send. This feature tries, in order: (1) find the small caret/dropdown next to Send, (2) find a "Send later" menu item and click it, (3) find a calendar day button whose accessible label matches the target date and click it. Each step is wrapped so a miss falls back gracefully — at minimum, the computed date/time is always copied to your clipboard — and **the extension never clicks Outlook's own final Send/Schedule confirmation for you**; you always confirm the exact time in Outlook's own dialog. If Outlook changes this UI, the automation may stop at an earlier step than before — that's expected and handled, not a crash.
- Storage is local to the browser profile (`chrome.storage.local`) — tracked follow-ups, priority senders, contacts snapshots, and cost history won't sync to another device unless you're signed into Chrome sync with extension data sync enabled.
