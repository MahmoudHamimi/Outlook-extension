# InboxSentry for Outlook

A browser extension (Chrome/Edge, Manifest V3) for **outlook.office.com**,
**outlook.office365.com**, and **outlook.live.com**. A small floating "IS"
button appears in the bottom-right corner of the inbox (with a red count
badge when something needs attention); click it to open the panel with all
the tools below:

1. **Stale-email signals (automatic)** — continuously scans the visible message list and stamps a `⏰ Nd` badge directly onto any unread row that's sat unopened longer than the threshold (default 3 days, configurable in the "Stale" tab). Priority senders (see below) get their own, shorter threshold (default 1 day), so an important contact going quiet stands out faster than a newsletter doing the same — their badge is solid instead of outlined to tell them apart at a glance. Any individual flagged email can be **snoozed for a day** from the Stale tab if it's not actionable right now. No click required to flag; opening the email clears its badge right away — see "Recent fixes" below for why this used to lag.
2. **Priority senders (automatic)** — add a sender's name, email, or an `@domain.com` (to match everyone at a company) in the "Priority" tab, at one of two levels — **Normal** or **High**. Matching rows get a badge and a flat side-accent in the message list; High-priority badges pulse subtly so they stand out from Normal ones. The tab shows how many senders you're tracking and how many are visible in the current list, and a search box filters a long list.
3. **Follow-up tracker** — track the open email with an N-day reminder; overdue ones float to the top of the "Tracked" list and count toward the badge on the floating IA button.
4. **Contacts insights** — pure-SVG donut charts (top 5 + "Other") plus a full deduped sender list ("All senders (N)") showing who you receive email from most and who you send email to most. Each sender appears exactly once with an incrementing message count — never one row per email — and counts accumulate across scans instead of resetting (see "About the Contacts tab" below).
5. **Meeting cost calculator** — attendee count × salary-band hourly rate × duration, with a currency picker, a recurrence dropdown that projects an annual cost (weekly/bi-weekly/monthly), a live stopwatch that ticks the running cost of a meeting in real time, and a saved history of past calculations.
6. **Multi-format export** — PDF and HTML now use a minimalist, professional "letterhead" layout (gradient accent bar, small-caps brand line, serif subject heading, left-accented meta card, footer divider); plus one-click Markdown and plain-text file downloads of the open email, sender included.
7. **Schedule Send quick-picker** — one-click presets (Tomorrow 8am, Monday 9am, Friday EOD) *plus* a custom date/time picker for any minute of any day — both in the panel and as small pills (with an inline "Custom…" mini-picker) injected next to Outlook's own Send button. Times in the past are always blocked. This one is best-effort/experimental — see "Known limitations" below.

8. **Light / dark theme, density, and free positioning (new in 1.6.0)** — a header toggle (☾ / ☀) switches the whole panel between light and dark; a second toggle (▤) switches between normal and a **compact density** that tightens padding and font-size throughout the panel and nav rail; and the **"IS" launcher button itself is draggable** — press and drag it anywhere on screen to relocate the whole widget, the same way you'd drag any floating chat widget. A quick click still opens/closes the panel as always; only a real drag (a few pixels of movement) moves it, so nothing changes for anyone who just clicks it. All three default sensibly (theme from your OS preference, density to normal, position to bottom-right) and remember your choice after that.
9. **Keyword watch (new in 1.5.0, match modes added in 1.6.0)** — add words or phrases in the "Keywords" tab (e.g. `invoice`, `urgent`); any visible message whose subject or preview text matches one gets a `🔑` badge stamped on the row automatically, the same way Stale and Priority badges work. Each keyword picks a match mode: **Contains** (substring, the original 1.5.0 behavior — `urgent` also matches `urgently`), **Whole word** (`urgent` only matches on its own), or **Regex** (the keyword text is a case-insensitive regular expression; invalid patterns are rejected with an inline error rather than silently failing).
10. **Quick reply templates (new in 1.5.0)** — save reusable snippets in the "Compose" tab. Click into a compose field, then hit **Insert** to type the snippet in at your cursor (falls back to copying it to your clipboard if no compose field is focused); **Copy** always just copies.
11. **Attachment reminder (new in 1.5.0)** — best-effort safety net, on by default (toggle in the "Compose" tab): if your message text mentions "attach…" but no attachment can be found on the compose surface when you hit Send, a single confirmation prompt gives you a chance to double back before it actually sends. It never blocks a send you confirm.
12. **Email insights (new in 1.5.0)** — the "Insights" tab shows word count, estimated reading time, and link count for whatever's open in the reading pane, refreshing automatically as you switch emails.
13. **Contacts CSV export (new in 1.5.0)** — a "Download CSV" button on the Contacts tab exports the received/sent sender counts as a spreadsheet-ready file.

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
4. Open outlook.office.com (or outlook.live.com) — you should see the round "IS" button appear bottom-right.

No build step, no bundler — it's loaded exactly as-is.

## Recent changes

- **1.6.0 — density, drag-to-reposition, keyword match modes, gold accents.**
  A compact-density toggle (▤) tightens panel padding/font-size for small
  screens; the "IS" launcher button is now directly draggable — press and
  move it to relocate the whole widget anywhere on screen (persisted,
  clamped to stay on-screen) — a quick click still opens/closes the panel
  as before; keywords now pick a match mode — Contains (the old default),
  Whole word, or Regex — instead of always doing a plain substring match,
  with existing keywords upgraded to "Contains" automatically so nothing
  changes until you opt in. A small amount of gold was added purely as an
  accent (the logo mark's border, the header's top edge, and the
  High-priority badge/row signal) — not a palette change, just a
  highlight on the "this matters" cases.
- **1.5.0 — five new tools.** Keyword watch (auto-flag rows by subject/preview
  keyword, same mechanism as Priority senders), quick reply templates
  (save-and-insert-at-cursor snippets for compose), an attachment reminder
  (confirms before sending if "attach…" is mentioned but nothing's attached),
  an Insights tab (word count / reading time / link count for the open
  email), and CSV export for the Contacts tab. All five follow the existing
  patterns: local `chrome.storage.local` only, no new permissions, and every
  automation degrades to a manual fallback (clipboard copy, a plain confirm
  dialog) rather than failing silently.
- **Priority levels, domain matching, and search.** Priority senders now have a Normal or High level (High gets a bolder, subtly-pulsing badge); adding `@company.com` matches everyone at that domain instead of one address; and the Priority tab has a search box plus tracked/visible counts. Existing plain-text priority lists from 1.3.0 and earlier are upgraded automatically to Normal level the first time the extension runs — nothing is lost.
- **Priority-aware stale threshold and snoozing.** Priority senders can be flagged stale on a shorter, separately-configurable threshold (default 1 day) so an important contact going quiet stands out sooner; their badge and list entry are visually distinct from regular stale flags. Any single flagged email can now be snoozed for a day from the Stale tab.
- **Light / dark theme.** A header toggle switches the whole panel between light and dark; it starts from your OS preference and remembers your choice after that.
- **Visual redesign.** The panel now uses a flatter, more business-like palette (muted navy/slate instead of bright gradients), calmer priority highlighting (a solid side-accent instead of a glowing/pulsing wash on every row), and consistent spacing and typography throughout.
- **PDF/HTML export redesign.** Both now share one "letterhead" template — a
  thin accent bar, a small-caps "InboxSentry · Exported
  Email" brand line, a serif subject heading, a left-accented meta card for
  From/Exported, and a footer divider. No external fonts, images, or
  network requests — system fonts only, still fully offline.
- **Priority senders got real visual weight.** A matching row now gets a
  soft animated glow on its red side-accent, a subtle gradient background
  wash, and a small pulsing star in the corner; the `★ Priority` badge
  itself has a soft gradient fill and a gentle pulse — noticeable without
  being obnoxious.
- **Schedule Send: any time, not just three presets.** Both the panel tab
  and the inline pills next to Outlook's Send button now include a custom
  date + time picker (minute-level precision, not locked to 30-minute
  steps). Every scheduling path — presets, panel custom time, inline custom
  time — is now blocked centrally from targeting a time that's already in
  the past.
- **Contacts: deduped sender totals, not a per-email list.** Scanning now
  merges into persistent counts keyed by sender name, using each row's
  stable conversation id to guarantee a message is only ever counted once
  even across repeated scans — so the tab shows "Jane Doe — 14" rather than
  fourteen separate rows for Jane Doe. See "About the Contacts tab" below.

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
libraries" below), so "who I hear from most" is built purely from whatever
message rows Outlook has currently rendered in the list — not your whole
mailbox history. To get a fuller picture: open the folder you care about
(Inbox for "received from", Sent Items for "sent to"), scroll down to let
Outlook load more rows, then click "Scan this list" — repeat as often as
you like. Each sender appears exactly once, with a count that increments
by the number of *newly seen* messages each scan; every message row
carries a stable id (its DOM `id` or `data-convid`), and that id is
remembered so rescanning the same rows never double-counts them. Use
"Reset counts" to start over for both folders. The donut charts (top 5 +
"Other") and the "All senders" full list are built with plain SVG (a
`stroke-dasharray` ring trick) and HTML — no charting library.

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
- Storage is local to the browser profile (`chrome.storage.local`) — tracked follow-ups, priority senders, keywords, templates, contacts snapshots, and cost history won't sync to another device unless you're signed into Chrome sync with extension data sync enabled.
- **Attachment reminder can't see real attachments, only guess from the DOM.** It looks for elements OWA typically renders per-attachment near the compose box; if Outlook's markup for attachment chips changes, the reminder may fire even when a file is actually attached (or miss a case where one isn't) — it's a nudge, not a guarantee, and confirming always sends immediately.
- **Template "Insert" needs a compose field to still be focused/open.** It inserts at the last-focused compose body using the browser's own text-insertion API; if that field has since closed or lost track of your cursor, it falls back to copying the template to your clipboard instead.
