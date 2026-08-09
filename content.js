/* Inbox Assistant for Outlook Web — content.js
   Runs on outlook.office.com / outlook.office365.com / outlook.live.com.

   IMPORTANT: Outlook web's DOM is not publicly documented and changes over
   time. Every selector Outlook's markup depends on lives in the CONFIG
   object below so it's a single, obvious place to fix if a selector stops
   matching after a Microsoft update. Everything else degrades gracefully
   (shows "couldn't find that" instead of throwing) rather than breaking.
*/

(function () {
  "use strict";

  if (window.__inboxAssistantLoaded) return;
  window.__inboxAssistantLoaded = true;

  /* ============================== CONFIG ============================== */
  const CONFIG = {
    // Reading pane: the open email's body renders in an element with role="document".
    readingBodySelectors: ['[role="document"]'],
    // Reading pane: subject renders as a heading; OWA marks it with the allowTextSelection class.
    readingSubjectSelectors: [
      "h1.allowTextSelection",
      '[role="heading"].allowTextSelection',
      '[role="main"] [role="heading"]'
    ],
    // The Send button, wherever a compose surface renders it.
    sendButtonSelectors: ['button[aria-label="Send"]', 'button[name="Send"]'],
    // Best-effort, EXPERIMENTAL: the little dropdown/caret next to Send that opens
    // Outlook's own "Send later" menu. This is the least stable selector set in
    // this file — Outlook's send-options menu is an undocumented Fluent UI
    // control, so the Schedule Send feature always treats a miss here as a
    // graceful fallback (copy the time, let you pick it yourself) rather than
    // failing silently. See the Schedule section below for details.
    scheduleSendMoreOptionsSelectors: [
      'button[aria-label*="Send options" i]',
      'button[aria-label*="More send options" i]',
      'button[aria-label*="Schedule send" i]',
      'button[aria-label*="Send later" i]'
    ],
    // Reading pane: who the open email is from. OWA usually exposes this via an
    // aria-label/title starting with "From:" or containing an email address.
    readingSenderSelectors: [
      '[aria-label^="From:" i]',
      '[title^="From:" i]',
      'span[title*="@"]',
      'button[title*="@"]'
    ],
    // Message list (inbox view): each row is an "option"/"row" inside a listbox/grid.
    mailListContainerSelectors: ['div[role="listbox"]', 'div[aria-label*="message list" i]'],
    mailListItemSelectors: ['div[role="option"]', 'div[role="row"]'],
    followupStorageKey: "ia_followups_v1",
    priorityStorageKey: "ia_priority_senders_v1",
    contactsReceivedStorageKey: "ia_contacts_received_v1",
    contactsSentStorageKey: "ia_contacts_sent_v1",
    costHistoryStorageKey: "ia_cost_history_v1",
    // How many days an unread email can sit untouched before we flag it.
    staleThresholdDays: 3
  };

  /* ============================== UTIL ============================== */
  function qFirst(selectors, root = document) {
    for (const sel of selectors) {
      const el = root.querySelector(sel);
      if (el && el.offsetParent !== null) return el;
    }
    return null;
  }
  function qAllVisible(selectors, root = document) {
    const out = [];
    for (const sel of selectors) {
      root.querySelectorAll(sel).forEach((el) => {
        if (el.offsetParent !== null) out.push(el);
      });
    }
    return out;
  }
  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }
  function currentItemId() {
    // OWA URLs encode the open item as .../id/<longid>/...
    const m = window.location.href.match(/\/id\/([^/?#]+)/);
    if (m) return decodeURIComponent(m[1]);
    return null;
  }
  function storageGet(key) {
    return new Promise((resolve) => chrome.storage.local.get([key], (r) => resolve(r[key])));
  }
  function storageSet(key, value) {
    return new Promise((resolve) => chrome.storage.local.set({ [key]: value }, resolve));
  }
  function findSenderName() {
    for (const sel of CONFIG.readingSenderSelectors) {
      const el = document.querySelector(sel);
      if (!el || el.offsetParent === null) continue;
      const raw = (el.getAttribute("aria-label") || el.getAttribute("title") || el.textContent || "").trim();
      const cleaned = raw.replace(/^from:\s*/i, "").trim();
      if (cleaned) return cleaned;
    }
    return null;
  }
  function currentEmailInfo() {
    const id = currentItemId();
    if (!id) return null;
    const subjectEl = qFirst(CONFIG.readingSubjectSelectors);
    return {
      id,
      subject: subjectEl ? subjectEl.textContent.trim() : "(subject not found)",
      sender: findSenderName() || "(sender not found)",
      url: window.location.href
    };
  }
  function downloadBlob(filename, content, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  function safeFilename(subject) {
    return (subject || "email").replace(/[^a-z0-9\-_ ]/gi, "").trim().slice(0, 60) || "email";
  }
  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      return false;
    }
  }

  /* ---- Message-list row heuristics (used by stale, priority & contacts) ----
     OWA's list markup isn't documented and varies by skin, so these read only
     from accessible/robust signals: aria-label text, native tooltip "title"
     attributes, and bold-text rendering (OWA's convention for unread items).
     Anything that can't be determined confidently is simply skipped. */
  function isRowUnread(row) {
    // Most reliable signal: Outlook renders a per-row "Mark as read" / "Mark as
    // unread" toggle button whose title unambiguously reflects the row's
    // CURRENT state — "Mark as read" means clicking it would mark the row
    // read, i.e. it is currently unread; "Mark as unread" means the opposite.
    // Confirmed via diagnostic: this stayed accurate for both freshly-unread
    // and already-opened rows, unlike the fallbacks below.
    const buttons = row.querySelectorAll("button[title]");
    for (const b of buttons) {
      const t = (b.getAttribute("title") || "").trim().toLowerCase();
      if (t === "mark as read") return true;
      if (t === "mark as unread") return false;
    }
    const label = (row.getAttribute("aria-label") || "").trim();
    if (/^unread\b/i.test(label)) return true;
    if (row.querySelector('[class*="unread" i]')) return true;
    // Bold-text fallback, used only if neither signal above was found.
    // FIX: this used to also match the sender's avatar initials (e.g. "MH",
    // "A"), which Outlook renders bold/600-weight PERMANENTLY regardless of
    // read state — that false positive is why stale badges could never
    // clear once set. Avatar/initials elements are now explicitly skipped.
    const candidates = row.querySelectorAll("span, div");
    for (let i = 0; i < candidates.length && i < 12; i++) {
      const el = candidates[i];
      if (!el.textContent || !el.textContent.trim()) continue;
      if (el.closest('[class*="avatar" i]') || el.getAttribute("role") === "img") continue;
      const weight = window.getComputedStyle(el).fontWeight;
      if (weight === "bold" || parseInt(weight, 10) >= 600) return true;
    }
    return false;
  }
  function findRowDate(row) {
    const titled = row.querySelectorAll("[title]");
    for (const el of titled) {
      const t = el.getAttribute("title");
      if (!t) continue;
      const d = new Date(t);
      if (!isNaN(d.getTime()) && d.getTime() <= Date.now() + 60000) return d;
    }
    return null;
  }
  function findRowSender(row) {
    // Doubles as "find the row's contact" for the Contacts insights feature —
    // in Outlook's Sent Items view this same position shows the recipient
    // ("To: ...") instead of the sender, which is exactly what we want there.
    const label = row.getAttribute("aria-label") || "";
    const parts = label.split(/,|\u2022|\|/).map((s) => s.trim()).filter(Boolean);
    if (parts.length) {
      if (/^unread$/i.test(parts[0]) && parts[1]) return parts[1];
      if (!/^unread$/i.test(parts[0])) return parts[0];
    }
    const strong = row.querySelector("strong, b");
    if (strong && strong.textContent.trim()) return strong.textContent.trim();
    return null;
  }
  function priorityMatches(sender, prioritySenders) {
    if (!sender || !prioritySenders || prioritySenders.length === 0) return false;
    const s = sender.toLowerCase();
    return prioritySenders.some((p) => p && s.includes(p.toLowerCase()));
  }

  /* ============================== SHARED STATE ============================== */
  // Declared up top (rather than inline in each section) so every section can
  // safely reference these regardless of which order the sections initialize in.
  let staleThresholdDays = CONFIG.staleThresholdDays;
  let priorityCache = [];
  let lastStaleItems = []; // [{subject, sender, ageDays}] from the most recent scan
  let liveTimerInterval = null;
  let liveStartedAt = null;
  let liveElapsedBeforePause = 0;

  // Injected once into the real page (not the shadow root) so badges we add to
  // Outlook's own list rows render correctly; scoped with an "ia-" prefix to
  // avoid colliding with Outlook's own classes.
  const globalStyle = document.createElement("style");
  globalStyle.textContent = `
    .ia-row-badges { display:inline-flex; gap:4px; margin-left:8px; vertical-align:middle; }
    .ia-badge { display:inline-block; font-family:"Segoe UI",Arial,sans-serif; font-size:10px;
      font-weight:700; padding:1px 7px; border-radius:8px; line-height:16px; white-space:nowrap; }
    .ia-badge-stale { background:#FBE1DE; color:#B3261E; }
    .ia-badge-priority {
      background: linear-gradient(135deg,#FFF4CE,#FFDE85); color:#8A6D00;
      box-shadow: 0 0 0 1px rgba(138,109,0,.25);
      animation: ia-badge-pulse 2.2s ease-in-out infinite;
    }
    .ia-row-priority {
      position: relative !important;
      box-shadow: inset 3px 0 0 #B3261E !important;
      background: linear-gradient(90deg, rgba(179,38,30,.07), rgba(255,222,133,.14) 45%, transparent 92%) !important;
      animation: ia-priority-glow 2.8s ease-in-out infinite;
    }
    .ia-row-priority::before {
      content: "★";
      position: absolute; top: 2px; left: 3px; font-size: 9px; color: #B3261E;
      text-shadow: 0 0 5px rgba(179,38,30,.55);
      animation: ia-priority-star 1.9s ease-in-out infinite;
      pointer-events: none; z-index: 5;
    }
    @keyframes ia-priority-glow {
      0%, 100% { box-shadow: inset 3px 0 0 #B3261E; }
      50% { box-shadow: inset 3px 0 0 #B3261E, inset 0 0 16px rgba(179,38,30,.22); }
    }
    @keyframes ia-priority-star {
      0%, 100% { opacity: .5; transform: scale(1); }
      50% { opacity: 1; transform: scale(1.3); }
    }
    @keyframes ia-badge-pulse {
      0%, 100% { transform: scale(1); }
      50% { transform: scale(1.06); }
    }
  `;
  document.documentElement.appendChild(globalStyle);

  /* ============================== PANEL UI ============================== */
  const host = document.createElement("div");
  host.id = "inbox-assistant-host";
  host.style.cssText = "all:initial; position:fixed; z-index:2147483000; bottom:20px; right:20px;";
  document.documentElement.appendChild(host);
  const shadow = host.attachShadow({ mode: "open" });

  shadow.innerHTML = `
    <style>
      :host { all: initial; }
      * { box-sizing: border-box; font-family: "Segoe UI", Arial, sans-serif; }
      .toggle {
        width: 54px; height: 54px; border-radius: 50%;
        background: linear-gradient(135deg,#1B4F8C,#20B999); color:#fff;
        border:none; cursor:pointer; box-shadow:0 4px 16px rgba(27,79,140,.45); font-size:14px;
        font-weight:800; position:relative; display:flex; align-items:center; justify-content:center;
        transition: transform .15s ease;
      }
      .toggle:hover { transform:scale(1.06); }
      .toggle-badge {
        position:absolute; top:-4px; right:-4px; background:#B3261E; color:#fff; border-radius:10px;
        min-width:18px; height:18px; display:none; align-items:center; justify-content:center;
        font-size:10px; font-weight:700; padding:0 4px; box-shadow:0 0 0 2px #fff;
      }
      .panel {
        display:none; position:fixed; bottom:84px; right:20px; width:390px; max-height:78vh;
        background:#fff; border-radius:14px; box-shadow:0 10px 34px rgba(20,30,50,.28);
        overflow:hidden; flex-direction:column; border:1px solid #E1E5EA;
      }
      .panel.open { display:flex; }
      header { background: linear-gradient(135deg,#1B4F8C 0%, #2F6FB0 55%, #20B999 100%); color:#fff; padding:13px 15px; }
      header h1 { margin:0; font-size:15px; letter-spacing:.2px; }
      header p { margin:2px 0 0; font-size:10.5px; opacity:.9; }
      nav { display:flex; flex-wrap:wrap; background:#fff; border-bottom:1px solid #E1E5EA; }
      nav button {
        flex:1 1 25%; border:none; background:none; padding:7px 2px 6px; font-size:9.5px; cursor:pointer;
        color:#5B6472; border-bottom:3px solid transparent; white-space:nowrap;
        display:flex; flex-direction:column; align-items:center; gap:1px;
      }
      nav button .nav-icon { font-size:14px; line-height:1; }
      nav button.active { color:#1B4F8C; border-bottom-color:#1B4F8C; font-weight:700; background:#F5F8FC; }
      main { padding:12px 14px; overflow-y:auto; font-size:12.5px; color:#20242B; }
      section { display:none; }
      section.active { display:block; }
      h2 { font-size:12.5px; margin:0 0 6px; }
      .muted { color:#5B6472; font-size:11px; line-height:1.5; }
      label { display:block; font-size:11px; font-weight:600; margin:8px 0 3px; }
      input, select {
        width:100%; padding:5px 7px; border:1px solid #E1E5EA; border-radius:6px; font-size:12px;
      }
      button.primary {
        background: linear-gradient(135deg,#1B4F8C,#2F8FD1); color:#fff; border:none; border-radius:7px;
        padding:7px 12px; font-size:11.5px; font-weight:600; cursor:pointer; margin-top:7px;
        box-shadow:0 2px 6px rgba(27,79,140,.25); transition: transform .1s ease;
      }
      button.primary:hover { transform:translateY(-1px); }
      button.secondary {
        background:#fff; color:#1B4F8C; border:1px solid #C9D6E8; border-radius:7px; padding:5px 10px;
        font-size:10.5px; cursor:pointer; margin-right:6px; margin-top:5px; font-weight:600;
      }
      button.secondary:hover { background:#EDF1F7; }
      .item { border:1px solid #E1E5EA; border-radius:8px; padding:7px 8px; margin-bottom:7px; }
      .pill { display:inline-block; padding:2px 7px; border-radius:9px; font-size:10px; font-weight:600; }
      .pill.warn { background:#FCE9DA; color:#B26A00; }
      .pill.ok { background:#E4F3E5; color:#2E7D32; }
      .pill.due { background:#FBE1DE; color:#B3261E; }
      .empty { color:#5B6472; font-style:italic; font-size:11px; }
      .warning-list { margin:0; padding-left:16px; }
      .warning-list li { margin-bottom:5px; }
      .row-flex { display:flex; gap:6px; align-items:flex-end; }
      .row-flex input, .row-flex select { flex:1; }
      .tag-list { list-style:none; margin:8px 0 0; padding:0; }
      .tag-list li {
        display:flex; justify-content:space-between; align-items:center;
        border:1px solid #E1E5EA; border-radius:6px; padding:5px 8px; margin-bottom:5px; font-size:11.5px;
      }
      .tag-list button { background:none; border:none; color:#B3261E; cursor:pointer; font-size:11px; }
      .stat-card {
        background: linear-gradient(135deg,#EDF1F7,#F5F8FC); border:1px solid #E1E5EA; border-radius:10px;
        padding:12px 14px; margin-top:8px; text-align:center;
      }
      .stat-big {
        font-size:26px; font-weight:800; background: linear-gradient(135deg,#1B4F8C,#20B999);
        -webkit-background-clip:text; background-clip:text; color:#1B4F8C;
      }
      @supports (-webkit-background-clip:text) { .stat-big { color:transparent; } }
      .donut-row { display:flex; gap:14px; align-items:center; }
      .donut-legend { list-style:none; margin:0; padding:0; font-size:11px; flex:1; }
      .donut-legend li { display:flex; align-items:center; gap:6px; margin-bottom:4px; }
      .dot { width:9px; height:9px; border-radius:50%; display:inline-block; flex-shrink:0; }
      details.contact-full-list { margin-top:10px; border:1px solid #E1E5EA; border-radius:8px; }
      details.contact-full-list summary {
        cursor:pointer; padding:7px 9px; font-size:11px; font-weight:600; color:#1B4F8C; list-style:none;
      }
      details.contact-full-list summary::-webkit-details-marker { display:none; }
      ul.contact-full-list-items {
        list-style:none; margin:0; padding:0 9px 9px; max-height:180px; overflow-y:auto;
      }
      ul.contact-full-list-items li {
        display:flex; justify-content:space-between; font-size:11px; padding:4px 0; border-top:1px solid #F0F2F5;
      }
      ul.contact-full-list-items li:first-child { border-top:none; }
    </style>
    <button class="toggle" id="ia-toggle" title="Inbox Assistant">
      <span>IA</span>
      <span class="toggle-badge" id="ia-toggle-badge">0</span>
    </button>
    <div class="panel" id="ia-panel">
      <header><h1>Inbox Assistant</h1><p>Smart tools for your inbox</p></header>
      <nav id="ia-tabs">
        <button data-tab="followup" class="active"><span class="nav-icon">📌</span>Follow-ups</button>
        <button data-tab="stale"><span class="nav-icon">⏰</span>Stale</button>
        <button data-tab="priority"><span class="nav-icon">★</span>Priority</button>
        <button data-tab="contacts"><span class="nav-icon">📊</span>Contacts</button>
        <button data-tab="cost"><span class="nav-icon">💰</span>Cost</button>
        <button data-tab="export"><span class="nav-icon">📤</span>Export</button>
        <button data-tab="schedule"><span class="nav-icon">🗓️</span>Schedule</button>
      </nav>
      <main>
        <section id="tab-followup" class="active">
          <h2>Track this email</h2>
          <p class="muted">Flags the open email if you haven't marked it replied within N days.</p>
          <label>Remind me after (days)</label>
          <input type="number" id="ia-days" value="3" min="1" max="60" />
          <button class="primary" id="ia-track">Track this email</button>
          <p class="muted" id="ia-track-status"></p>
          <h2 style="margin-top:14px;">Tracked</h2>
          <div id="ia-followup-list"><p class="empty">Nothing tracked yet.</p></div>
        </section>

        <section id="tab-stale">
          <h2>Unopened &gt; <span id="ia-stale-threshold-label">3</span> days</h2>
          <p class="muted">Automatically scans the message list and flags any unread email that's sat unopened past the threshold — look for the <span class="pill due">⏰ Nd</span> badge directly on the row. Opening the email clears its badge right away.</p>
          <label>Flag unread emails older than (days)</label>
          <div class="row-flex">
            <input type="number" id="ia-stale-days" min="1" max="60" value="3" />
            <button class="primary" id="ia-stale-save" style="margin-top:0;">Save</button>
          </div>
          <h2 style="margin-top:14px;">Currently flagged (<span id="ia-stale-count">0</span>)</h2>
          <div id="ia-stale-list"><p class="empty">Nothing flagged in the visible list right now.</p></div>
          <button class="secondary" id="ia-stale-rescan">Rescan visible list</button>
        </section>

        <section id="tab-priority">
          <h2>Priority senders</h2>
          <p class="muted">Emails from these senders get a <span class="pill warn">★ Priority</span> badge and a red side-accent directly in the message list, automatically.</p>
          <label>Add sender (name or email)</label>
          <div class="row-flex">
            <input type="text" id="ia-priority-input" placeholder="e.g. jane@company.com" />
            <button class="primary" id="ia-priority-add" style="margin-top:0;">Add</button>
          </div>
          <button class="secondary" id="ia-priority-add-current">Add sender of open email</button>
          <ul class="tag-list" id="ia-priority-list"><li class="empty" style="border:none;">No priority senders yet.</li></ul>
        </section>

        <section id="tab-contacts">
          <h2>Who you hear from most</h2>
          <p class="muted">Each sender is counted once per unique message and never listed twice — scanning again only adds newly-loaded messages, so scroll to load more, open the folder you want (Inbox or Sent Items), then scan. Nothing leaves your browser.</p>
          <button class="primary" id="ia-contacts-scan">📊 Scan this list</button>
          <button class="secondary" id="ia-contacts-reset">Reset counts</button>
          <p class="muted" id="ia-contacts-status"></p>
          <h2 style="margin-top:14px;">Received from</h2>
          <div id="ia-contacts-received"><p class="empty">No inbox scan yet — open your Inbox and scan.</p></div>
          <h2 style="margin-top:14px;">Sent to</h2>
          <div id="ia-contacts-sent"><p class="empty">No Sent scan yet — open Sent Items and scan.</p></div>
        </section>

        <section id="tab-cost">
          <h2>Meeting cost calculator</h2>
          <div class="row-flex">
            <div style="flex:1;"><label>Attendee count</label><input type="number" id="ia-cost-attendees" min="1" value="4" /></div>
            <div style="flex:1;"><label>Currency</label>
              <select id="ia-cost-currency">
                <option value="$">USD ($)</option>
                <option value="€">EUR (€)</option>
                <option value="£">GBP (£)</option>
              </select>
            </div>
          </div>
          <label>Average salary band</label>
          <select id="ia-cost-salary">
            <option value="50000">~50,000/yr</option>
            <option value="75000" selected>~75,000/yr</option>
            <option value="100000">~100,000/yr</option>
            <option value="125000">~125,000/yr</option>
            <option value="150000">~150,000/yr</option>
            <option value="200000">~200,000/yr</option>
          </select>
          <div class="row-flex">
            <div style="flex:1;"><label>Duration (minutes)</label><input type="number" id="ia-cost-duration" min="5" step="5" value="30" /></div>
            <div style="flex:1;"><label>Recurs</label>
              <select id="ia-cost-recurrence">
                <option value="1">One-time</option>
                <option value="52">Weekly</option>
                <option value="26">Bi-weekly</option>
                <option value="12">Monthly</option>
              </select>
            </div>
          </div>
          <button class="primary" id="ia-cost-calc">Calculate</button>
          <div class="stat-card">
            <div class="stat-big" id="ia-cost-result">—</div>
            <p class="muted" id="ia-cost-note"></p>
            <p class="muted" id="ia-cost-fun"></p>
          </div>

          <h2 style="margin-top:14px;">⏱️ Live meeting timer</h2>
          <p class="muted">Start it when the meeting starts — watch the cost climb in real time.</p>
          <div class="stat-card">
            <div class="stat-big" id="ia-cost-live">$0.00</div>
            <p class="muted" id="ia-cost-live-elapsed">00:00</p>
            <button class="primary" id="ia-cost-live-toggle" style="margin-right:6px;">▶ Start</button>
            <button class="secondary" id="ia-cost-live-reset">Reset</button>
          </div>

          <h2 style="margin-top:14px;">History (<span id="ia-cost-history-count">0</span>)</h2>
          <div id="ia-cost-history"><p class="empty">No meetings calculated yet.</p></div>
          <button class="secondary" id="ia-cost-history-clear">Clear history</button>
        </section>

        <section id="tab-export">
          <h2>Export this email</h2>
          <p class="muted">PDF and HTML use a minimalist, professional letterhead layout (gradient accent bar, serif heading, meta card) — everything is generated locally in your browser, nothing is uploaded.</p>
          <button class="primary" id="ia-export-pdf">🖨️ PDF (print dialog)</button><br/>
          <button class="secondary" id="ia-export-html">🌐 HTML file</button>
          <button class="secondary" id="ia-export-md">📝 Markdown file</button>
          <button class="secondary" id="ia-export-txt">📄 Plain text file</button>
          <p class="muted" id="ia-export-status"></p>
        </section>

        <section id="tab-schedule">
          <h2>🗓️ Schedule Send Quick-Picker</h2>
          <p class="muted">These presets — plus a "Custom…" option — also appear as small pills next to Outlook's own Send button while composing. Picking one opens Outlook's native "Send later" dialog and gets as close as possible automatically — you always confirm the exact time yourself in Outlook's own dialog. Times in the past are blocked.</p>
          <div class="row-flex" style="flex-wrap:wrap;">
            <button class="secondary" data-preset="tomorrow8">Tomorrow 8am</button>
            <button class="secondary" data-preset="monday9">Monday 9am</button>
            <button class="secondary" data-preset="fridayEod">Friday EOD</button>
          </div>
          <h2 style="margin-top:14px;">Custom time</h2>
          <p class="muted">Pick any date and any time — not limited to the presets above.</p>
          <div class="row-flex">
            <div style="flex:1;"><label>Date</label><input type="date" id="ia-schedule-custom-date" /></div>
            <div style="flex:1;"><label>Time</label><input type="time" id="ia-schedule-custom-time" step="60" /></div>
          </div>
          <button class="primary" id="ia-schedule-custom-go">Schedule this time</button>
          <p class="muted" id="ia-schedule-status"></p>
        </section>
      </main>
    </div>
  `;

  const $ = (sel) => shadow.querySelector(sel);

  $("#ia-toggle").addEventListener("click", () => {
    $("#ia-panel").classList.toggle("open");
  });

  shadow.querySelectorAll("#ia-tabs button").forEach((btn) => {
    btn.addEventListener("click", () => {
      shadow.querySelectorAll("#ia-tabs button").forEach((b) => b.classList.remove("active"));
      shadow.querySelectorAll("main section").forEach((s) => s.classList.remove("active"));
      btn.classList.add("active");
      $("#tab-" + btn.dataset.tab).classList.add("active");
    });
  });

  async function updateToggleBadge() {
    const followUps = (await getFollowUps()).filter((f) => !f.replied);
    const overdue = followUps.filter(
      (f) => Math.floor((Date.now() - new Date(f.trackedAt).getTime()) / 86400000) >= f.days
    ).length;
    const total = overdue + lastStaleItems.length;
    const badge = $("#ia-toggle-badge");
    if (!badge) return;
    if (total > 0) {
      badge.style.display = "flex";
      badge.textContent = total > 99 ? "99+" : String(total);
    } else {
      badge.style.display = "none";
    }
  }

  /* =========================================================
     1) FOLLOW-UP TRACKER
     ========================================================= */
  async function getFollowUps() {
    return (await storageGet(CONFIG.followupStorageKey)) || [];
  }
  async function saveFollowUps(list) {
    await storageSet(CONFIG.followupStorageKey, list);
  }

  $("#ia-track").addEventListener("click", async () => {
    const id = currentItemId();
    const statusEl = $("#ia-track-status");
    if (!id) {
      statusEl.textContent = "Open a specific email in the reading pane first.";
      return;
    }
    const subjectEl = qFirst(CONFIG.readingSubjectSelectors);
    const subject = subjectEl ? subjectEl.textContent.trim() : "(subject not found)";
    const sender = findSenderName() || "(sender not found)";
    const days = parseInt($("#ia-days").value, 10) || 3;

    const list = await getFollowUps();
    if (list.some((f) => f.id === id)) {
      statusEl.textContent = "Already tracking this email.";
      return;
    }
    list.push({
      id,
      subject,
      sender,
      url: window.location.href,
      trackedAt: new Date().toISOString(),
      days,
      replied: false
    });
    await saveFollowUps(list);
    statusEl.textContent = `Tracking — you'll see it flagged after ${days} day(s).`;
    renderFollowUps();
  });

  async function renderFollowUps() {
    const list = (await getFollowUps()).filter((f) => !f.replied);
    const container = $("#ia-followup-list");
    if (list.length === 0) {
      container.innerHTML = '<p class="empty">Nothing tracked yet.</p>';
      updateToggleBadge();
      return;
    }
    const withElapsed = list.map((f) => ({
      ...f,
      elapsedDays: Math.floor((Date.now() - new Date(f.trackedAt).getTime()) / 86400000)
    }));
    // Overdue items float to the top.
    withElapsed.sort((a, b) => (b.elapsedDays - b.days) - (a.elapsedDays - a.days));

    container.innerHTML = withElapsed
      .map((f) => {
        const overdue = f.elapsedDays >= f.days;
        const pill = overdue
          ? '<span class="pill due">Follow up now</span>'
          : `<span class="pill ok">Due in ${Math.max(0, f.days - f.elapsedDays)}d</span>`;
        return `
        <div class="item" data-id="${escapeHtml(f.id)}">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <strong style="font-size:11.5px;">${escapeHtml(f.subject).slice(0, 60)}</strong>
          </div>
          <div class="muted">${f.sender ? escapeHtml(f.sender) + " &middot; " : ""}${pill} &middot; tracked ${new Date(f.trackedAt).toLocaleDateString()}</div>
          <button class="secondary" data-action="open">Open</button>
          <button class="secondary" data-action="replied">Mark replied</button>
        </div>`;
      })
      .join("");

    container.querySelectorAll(".item").forEach((el) => {
      const id = el.dataset.id;
      el.querySelector('[data-action="open"]').addEventListener("click", async () => {
        const list2 = await getFollowUps();
        const f = list2.find((x) => x.id === id);
        if (f) window.location.href = f.url;
      });
      el.querySelector('[data-action="replied"]').addEventListener("click", async () => {
        const list2 = await getFollowUps();
        const idx = list2.findIndex((x) => x.id === id);
        if (idx >= 0) {
          list2[idx].replied = true;
          await saveFollowUps(list2);
          renderFollowUps();
        }
      });
    });
    updateToggleBadge();
  }
  renderFollowUps();

  /* =========================================================
     1b) STALE-EMAIL SIGNALS — automatic, no click required.
     Scans the visible message list on a debounce (list re-renders
     constantly as OWA virtualizes rows) and stamps a badge onto any
     unread row whose received date is older than the threshold.

     FIX: the badge used to survive after you opened a stale email.
     That was because the old MutationObserver only watched for
     nodes being added/removed (childList), but Outlook usually
     marks a row "read" by changing an *attribute* (aria-label /
     class) on the existing row rather than swapping the node out —
     so no childList mutation ever fired and the badge never got
     recomputed. This version also watches attributes, and — since
     Outlook sometimes marks a row read a second or two after the
     click rather than instantly — a click on any row also triggers
     an optimistic same-frame badge removal plus a couple of
     follow-up rescans to reconcile with whatever Outlook ends up
     doing.
     ========================================================= */
  async function loadStaleThreshold() {
    const stored = await storageGet("ia_stale_days_v1");
    staleThresholdDays = stored || CONFIG.staleThresholdDays;
    $("#ia-stale-days").value = staleThresholdDays;
    $("#ia-stale-threshold-label").textContent = staleThresholdDays;
  }
  async function refreshPriorityCache() {
    priorityCache = (await storageGet(CONFIG.priorityStorageKey)) || [];
  }

  function applyRowSignals(row) {
    const old = row.querySelector(".ia-row-badges");
    if (old) old.remove();
    row.classList.remove("ia-row-priority");

    const unread = isRowUnread(row);
    const date = findRowDate(row);
    const sender = findRowSender(row);
    const badges = [];

    if (unread && date) {
      const ageDays = Math.floor((Date.now() - date.getTime()) / 86400000);
      if (ageDays >= staleThresholdDays) {
        badges.push(`<span class="ia-badge ia-badge-stale" title="Unopened for ${ageDays} day(s)">⏰ ${ageDays}d</span>`);
        lastStaleItems.push({
          subject: (row.getAttribute("aria-label") || "").split(",")[0] || "(email)",
          sender: sender || "(unknown sender)",
          ageDays
        });
      }
    }
    if (priorityMatches(sender, priorityCache)) {
      badges.push('<span class="ia-badge ia-badge-priority" title="Priority sender">★ Priority</span>');
      row.classList.add("ia-row-priority");
    }
    if (badges.length) {
      const wrap = document.createElement("span");
      wrap.className = "ia-row-badges";
      wrap.innerHTML = badges.join("");
      row.appendChild(wrap);
    }
  }

  function scanMailList() {
    lastStaleItems = [];
    const container = qFirst(CONFIG.mailListContainerSelectors) || document.body;
    const rows = qAllVisible(CONFIG.mailListItemSelectors, container);
    rows.forEach(applyRowSignals);
    renderStaleTab();
  }

  function renderStaleTab() {
    const countEl = $("#ia-stale-count");
    const listEl = $("#ia-stale-list");
    if (countEl && listEl) {
      countEl.textContent = lastStaleItems.length;
      if (lastStaleItems.length === 0) {
        listEl.innerHTML = '<p class="empty">Nothing flagged in the visible list right now.</p>';
      } else {
        listEl.innerHTML = lastStaleItems
          .sort((a, b) => b.ageDays - a.ageDays)
          .map(
            (it) => `
          <div class="item">
            <strong style="font-size:11.5px;">${escapeHtml(it.subject).slice(0, 60)}</strong>
            <div class="muted">${escapeHtml(it.sender)} &middot; <span class="pill due">${it.ageDays}d unopened</span></div>
          </div>`
          )
          .join("");
      }
    }
    updateToggleBadge();
  }

  let scanTimer = null;
  function scheduleScan() {
    clearTimeout(scanTimer);
    scanTimer = setTimeout(scanMailList, 500);
  }
  // Watch both node changes AND attribute changes (see FIX note above) so a
  // row that goes from unread -> read without being replaced still triggers
  // a rescan and clears its badge.
  new MutationObserver(scheduleScan).observe(document.body, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["aria-label", "class", "title"]
  });
  loadStaleThreshold().then(() => refreshPriorityCache().then(scanMailList));
  // OWA re-renders on its own, but a row's age can cross the threshold with no
  // DOM change at all (time just passes), so also re-check periodically.
  setInterval(scanMailList, 5 * 60 * 1000);

  // Belt-and-suspenders: optimistically clear a clicked row's stale badge
  // immediately (instant visual feedback), then reconcile with a couple of
  // staggered rescans in case Outlook takes a moment to actually mark it read.
  document.addEventListener(
    "click",
    (e) => {
      const row = e.target.closest(CONFIG.mailListItemSelectors.join(","));
      if (!row) return;
      const staleBadge = row.querySelector(".ia-badge-stale");
      if (staleBadge) staleBadge.remove();
      [300, 900, 2000].forEach((delay) => setTimeout(scanMailList, delay));
    },
    true
  );
  window.addEventListener("focus", scanMailList);

  $("#ia-stale-save").addEventListener("click", async () => {
    const days = Math.max(1, parseInt($("#ia-stale-days").value, 10) || 3);
    staleThresholdDays = days;
    await storageSet("ia_stale_days_v1", days);
    $("#ia-stale-threshold-label").textContent = days;
    scanMailList();
  });
  $("#ia-stale-rescan").addEventListener("click", scanMailList);

  /* =========================================================
     1c) PRIORITY SENDERS
     ========================================================= */
  async function getPrioritySenders() {
    return (await storageGet(CONFIG.priorityStorageKey)) || [];
  }
  async function savePrioritySenders(list) {
    await storageSet(CONFIG.priorityStorageKey, list);
    await refreshPriorityCache();
    scanMailList();
  }
  async function renderPriorityList() {
    const list = await getPrioritySenders();
    const el = $("#ia-priority-list");
    if (list.length === 0) {
      el.innerHTML = '<li class="empty" style="border:none;">No priority senders yet.</li>';
      return;
    }
    el.innerHTML = list
      .map((s, i) => `<li data-i="${i}"><span>${escapeHtml(s)}</span><button data-action="remove">Remove</button></li>`)
      .join("");
    el.querySelectorAll("li").forEach((li) => {
      const i = parseInt(li.dataset.i, 10);
      const removeBtn = li.querySelector('[data-action="remove"]');
      if (!removeBtn) return;
      removeBtn.addEventListener("click", async () => {
        const cur = await getPrioritySenders();
        cur.splice(i, 1);
        await savePrioritySenders(cur);
        renderPriorityList();
      });
    });
  }
  async function addPrioritySender(value) {
    const v = (value || "").trim();
    if (!v) return;
    const cur = await getPrioritySenders();
    if (cur.some((s) => s.toLowerCase() === v.toLowerCase())) return;
    cur.push(v);
    await savePrioritySenders(cur);
    renderPriorityList();
  }
  $("#ia-priority-add").addEventListener("click", () => {
    const input = $("#ia-priority-input");
    addPrioritySender(input.value);
    input.value = "";
  });
  $("#ia-priority-add-current").addEventListener("click", () => {
    const info = currentEmailInfo();
    if (!info || info.sender === "(sender not found)") return;
    addPrioritySender(info.sender);
  });
  renderPriorityList();

  /* =========================================================
     1d) CONTACTS INSIGHTS — who you receive from / send to most.
     Built entirely from whatever message rows are currently loaded
     in the visible list (no Microsoft Graph API involved, so it's a
     snapshot of what's rendered — scroll to load more rows before
     scanning for a fuller picture). Rendered as pure-SVG donut
     charts (a stroke-dasharray ring trick) with no charting library.
     ========================================================= */
  const DONUT_COLORS = ["#1B4F8C", "#20B999", "#E8A33D", "#B3261E", "#6C3FBF", "#2F8FD1", "#8A6D00", "#5B6472"];

  function detectFolderContext() {
    const path = window.location.pathname.toLowerCase();
    if (/\/sentitems\b/.test(path)) return "sent";
    if (/\/inbox\b/.test(path) || /\/mail\/0\/?$/.test(path) || path.endsWith("/mail")) return "received";
    const selected = document.querySelector('[aria-selected="true"], [aria-current="true"], [aria-current="page"]');
    const label = selected ? (selected.getAttribute("aria-label") || selected.textContent || "") : "";
    if (/sent/i.test(label)) return "sent";
    if (/inbox/i.test(label)) return "received";
    return "unknown";
  }

  function rowUniqueId(row) {
    // OWA row nodes carry a stable per-conversation id (either the element's
    // own id or a data-convid attribute) — used so a sender's count only
    // increments once per unique message, even if the same rows get rescanned.
    return row.id || row.getAttribute("data-convid") || null;
  }

  function scanContactsSnapshot(existing) {
    const context = detectFolderContext();
    const container = qFirst(CONFIG.mailListContainerSelectors) || document.body;
    const rows = qAllVisible(CONFIG.mailListItemSelectors, container);
    const counts = { ...((existing && existing.counts) || {}) };
    const seenIds = new Set((existing && existing.seenIds) || []);
    let newlyCounted = 0;
    rows.forEach((row) => {
      const name = findRowSender(row);
      if (!name) return;
      const uid = rowUniqueId(row);
      if (uid) {
        if (seenIds.has(uid)) return; // already counted in an earlier scan — no repeats
        seenIds.add(uid);
      }
      counts[name] = (counts[name] || 0) + 1;
      newlyCounted++;
    });
    // Cap how many ids we remember so storage doesn't grow without bound.
    const seenIdsArr = Array.from(seenIds);
    const cappedSeenIds = seenIdsArr.length > 2000 ? seenIdsArr.slice(seenIdsArr.length - 2000) : seenIdsArr;
    return {
      context,
      counts,
      seenIds: cappedSeenIds,
      scannedAt: new Date().toISOString(),
      rowsSeenThisScan: rows.length,
      newlyCountedThisScan: newlyCounted
    };
  }

  function topEntries(counts, n = 5) {
    const arr = Object.entries(counts)
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count);
    const top = arr.slice(0, n);
    const restTotal = arr.slice(n).reduce((s, x) => s + x.count, 0);
    if (restTotal > 0) top.push({ label: "Other", count: restTotal });
    return top.map((item, i) => ({
      ...item,
      color: item.label === "Other" ? "#C7CDD6" : DONUT_COLORS[i % DONUT_COLORS.length]
    }));
  }

  function buildDonutSVG(items, size = 132, thickness = 24) {
    const total = items.reduce((s, i) => s + i.count, 0) || 1;
    const r = (size - thickness) / 2;
    const cx = size / 2;
    const cy = size / 2;
    const circumference = 2 * Math.PI * r;
    let offset = 0;
    const circles = items
      .map((it) => {
        const frac = it.count / total;
        const dash = Math.max(frac * circumference, frac > 0 ? 0.5 : 0);
        const seg = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${it.color}" stroke-width="${thickness}" stroke-dasharray="${dash} ${circumference - dash}" stroke-dashoffset="${-offset}"></circle>`;
        offset += dash;
        return seg;
      })
      .join("");
    return `<svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}">
      <g transform="rotate(-90 ${cx} ${cy})">${circles}</g>
      <circle cx="${cx}" cy="${cy}" r="${r - thickness / 2 - 3}" fill="#fff"></circle>
      <text x="${cx}" y="${cy - 3}" text-anchor="middle" font-size="16" font-weight="800" fill="#20242B">${total}</text>
      <text x="${cx}" y="${cy + 13}" text-anchor="middle" font-size="8.5" fill="#5B6472">emails</text>
    </svg>`;
  }

  function renderContactChart(containerId, data, emptyLabel) {
    const el = $(containerId);
    if (!el) return;
    if (!data || !data.counts || Object.keys(data.counts).length === 0) {
      el.innerHTML = `<p class="empty">${emptyLabel}</p>`;
      return;
    }
    // One entry per unique sender — never one row per email — sorted by how
    // many messages have been counted from them, highest first.
    const allEntries = Object.entries(data.counts)
      .map(([label, count]) => ({ label, count }))
      .sort((a, b) => b.count - a.count);
    const grandTotal = allEntries.reduce((s, x) => s + x.count, 0);
    const chartItems = topEntries(data.counts, 5);
    const legend = chartItems
      .map((it) => {
        const pct = Math.round((it.count / grandTotal) * 100);
        return `<li><span class="dot" style="background:${it.color}"></span>${escapeHtml(it.label)} <span class="muted">${it.count} &middot; ${pct}%</span></li>`;
      })
      .join("");
    const fullList = allEntries
      .map((e) => `<li><span>${escapeHtml(e.label)}</span><span class="muted">${e.count}</span></li>`)
      .join("");
    el.innerHTML = `
      <div class="donut-row">
        ${buildDonutSVG(chartItems)}
        <ul class="donut-legend">${legend}</ul>
      </div>
      <details class="contact-full-list">
        <summary>All senders (${allEntries.length}) &middot; ${grandTotal} message(s) tracked</summary>
        <ul class="contact-full-list-items">${fullList}</ul>
      </details>
      <p class="muted" style="margin-top:6px;">Counts accumulate across scans — each message counted once &middot; last scanned ${new Date(data.scannedAt).toLocaleString()}</p>
    `;
  }

  async function refreshContactsTabs() {
    const received = await storageGet(CONFIG.contactsReceivedStorageKey);
    const sent = await storageGet(CONFIG.contactsSentStorageKey);
    renderContactChart("#ia-contacts-received", received, "No inbox scan yet — open your Inbox and scan.");
    renderContactChart("#ia-contacts-sent", sent, "No Sent scan yet — open Sent Items and scan.");
  }

  $("#ia-contacts-scan").addEventListener("click", async () => {
    const statusEl = $("#ia-contacts-status");
    const context = detectFolderContext();
    let existing = null;
    if (context === "sent") existing = await storageGet(CONFIG.contactsSentStorageKey);
    else if (context === "received") existing = await storageGet(CONFIG.contactsReceivedStorageKey);

    const snap = scanContactsSnapshot(existing);
    if (Object.keys(snap.counts).length === 0) {
      statusEl.textContent = "Couldn't read any names from the currently visible list — make sure a message list is open.";
      return;
    }
    if (snap.context === "sent") {
      await storageSet(CONFIG.contactsSentStorageKey, snap);
      statusEl.textContent = `Counted ${snap.newlyCountedThisScan} new sent message(s) this scan.`;
    } else if (snap.context === "received") {
      await storageSet(CONFIG.contactsReceivedStorageKey, snap);
      statusEl.textContent = `Counted ${snap.newlyCountedThisScan} new inbox message(s) this scan.`;
    } else {
      statusEl.textContent = `Scanned ${snap.rowsSeenThisScan} message(s) but couldn't tell if this is Inbox or Sent Items — open one of those folders to save the result.`;
      return;
    }
    refreshContactsTabs();
  });

  $("#ia-contacts-reset").addEventListener("click", async () => {
    await storageSet(CONFIG.contactsReceivedStorageKey, null);
    await storageSet(CONFIG.contactsSentStorageKey, null);
    $("#ia-contacts-status").textContent = "Counts reset.";
    refreshContactsTabs();
  });

  refreshContactsTabs();

  /* =========================================================
     2) MEETING COST CALCULATOR — a one-off estimate, a live
     running stopwatch that ticks the cost up in real time, and a
     saved history of past calculations.
     ========================================================= */
  function currencySymbol() {
    const el = $("#ia-cost-currency");
    return (el && el.value) || "$";
  }
  function computeHourlyRate(salary) {
    return (salary / 2080) * 1.3; // +30% for benefits/overhead
  }
  async function getCostHistory() {
    return (await storageGet(CONFIG.costHistoryStorageKey)) || [];
  }
  async function saveCostHistory(list) {
    await storageSet(CONFIG.costHistoryStorageKey, list.slice(0, 20));
  }
  function renderCostHistory() {
    getCostHistory().then((history) => {
      $("#ia-cost-history-count").textContent = history.length;
      const el = $("#ia-cost-history");
      if (history.length === 0) {
        el.innerHTML = '<p class="empty">No meetings calculated yet.</p>';
        return;
      }
      el.innerHTML = history
        .slice(0, 8)
        .map(
          (h) => `<div class="item">
            <strong style="font-size:11.5px;">${h.symbol}${h.cost.toFixed(0)}</strong>
            <span class="muted"> &middot; ${h.attendees} attendee(s) &middot; ${h.minutes}min &middot; ${escapeHtml(h.recurrenceLabel)}</span>
            <div class="muted">${new Date(h.ts).toLocaleString()}</div>
          </div>`
        )
        .join("");
    });
  }
  renderCostHistory();

  $("#ia-cost-calc").addEventListener("click", async () => {
    const attendees = Math.max(1, parseInt($("#ia-cost-attendees").value, 10) || 1);
    const salary = parseInt($("#ia-cost-salary").value, 10);
    const minutes = Math.max(1, parseInt($("#ia-cost-duration").value, 10) || 30);
    const recurTimesPerYear = parseInt($("#ia-cost-recurrence").value, 10) || 1;
    const symbol = currencySymbol();
    const hourlyRate = computeHourlyRate(salary);
    const hours = minutes / 60;
    const cost = attendees * hourlyRate * hours;
    const annual = cost * recurTimesPerYear;
    const recurrenceLabel = { "1": "one-time", "52": "weekly", "26": "bi-weekly", "12": "monthly" }[String(recurTimesPerYear)] || "one-time";

    $("#ia-cost-result").textContent = `${symbol}${cost.toFixed(0)}`;
    $("#ia-cost-note").textContent =
      `${attendees} attendee(s) × ~${symbol}${hourlyRate.toFixed(0)}/hr (incl. ~30% overhead) × ${hours.toFixed(2)} hr(s).` +
      (recurTimesPerYear > 1 ? ` If ${recurrenceLabel}, that's ~${symbol}${annual.toLocaleString(undefined, { maximumFractionDigits: 0 })}/year.` : "");
    $("#ia-cost-fun").textContent = `≈ ${Math.max(1, Math.round(cost / 5))} coffee run(s) worth of budget ☕ — a rough illustration, not a real conversion.`;

    const history = await getCostHistory();
    history.unshift({ ts: new Date().toISOString(), attendees, minutes, cost, symbol, recurrenceLabel });
    await saveCostHistory(history);
    renderCostHistory();
  });

  $("#ia-cost-history-clear").addEventListener("click", async () => {
    await saveCostHistory([]);
    renderCostHistory();
  });

  /* ---- Live meeting timer ---- */
  function formatElapsed(totalSeconds) {
    const m = Math.floor(totalSeconds / 60).toString().padStart(2, "0");
    const s = Math.floor(totalSeconds % 60).toString().padStart(2, "0");
    return `${m}:${s}`;
  }
  function tickLiveCost() {
    const attendees = Math.max(1, parseInt($("#ia-cost-attendees").value, 10) || 1);
    const salary = parseInt($("#ia-cost-salary").value, 10);
    const symbol = currencySymbol();
    const hourlyRate = computeHourlyRate(salary);
    const elapsedSeconds = liveElapsedBeforePause + (liveStartedAt ? (Date.now() - liveStartedAt) / 1000 : 0);
    const cost = attendees * hourlyRate * (elapsedSeconds / 3600);
    $("#ia-cost-live").textContent = `${symbol}${cost.toFixed(2)}`;
    $("#ia-cost-live-elapsed").textContent = formatElapsed(elapsedSeconds);
  }
  $("#ia-cost-live-toggle").addEventListener("click", async () => {
    const btn = $("#ia-cost-live-toggle");
    if (liveTimerInterval) {
      clearInterval(liveTimerInterval);
      liveTimerInterval = null;
      liveElapsedBeforePause += liveStartedAt ? (Date.now() - liveStartedAt) / 1000 : 0;
      liveStartedAt = null;
      btn.textContent = "▶ Resume";
      if (liveElapsedBeforePause >= 30) {
        const attendees = Math.max(1, parseInt($("#ia-cost-attendees").value, 10) || 1);
        const salary = parseInt($("#ia-cost-salary").value, 10);
        const symbol = currencySymbol();
        const hourlyRate = computeHourlyRate(salary);
        const cost = attendees * hourlyRate * (liveElapsedBeforePause / 3600);
        const history = await getCostHistory();
        history.unshift({
          ts: new Date().toISOString(),
          attendees,
          minutes: Math.round(liveElapsedBeforePause / 60),
          cost,
          symbol,
          recurrenceLabel: "live timer"
        });
        await saveCostHistory(history);
        renderCostHistory();
      }
    } else {
      liveStartedAt = Date.now();
      liveTimerInterval = setInterval(tickLiveCost, 500);
      btn.textContent = "⏸ Pause";
    }
  });
  $("#ia-cost-live-reset").addEventListener("click", () => {
    clearInterval(liveTimerInterval);
    liveTimerInterval = null;
    liveStartedAt = null;
    liveElapsedBeforePause = 0;
    $("#ia-cost-live-toggle").textContent = "▶ Start";
    $("#ia-cost-live").textContent = `${currencySymbol()}0.00`;
    $("#ia-cost-live-elapsed").textContent = "00:00";
  });

  /* =========================================================
     3) EXPORT — PDF via the native print dialog, plus HTML,
     Markdown, and plain-text file downloads generated locally
     with no external library and no network request.

     PDF/HTML now share one minimalist, professional "letterhead"
     template — a thin gradient top bar, a small caps brand line, a
     serif subject heading, and a left-accented meta card — instead
     of the old plain black-on-white print sheet. No external fonts
     or images are loaded (system fonts only), so it still works
     completely offline with zero network requests.
     ========================================================= */
  function currentReadingEmail() {
    const bodyEl = qFirst(CONFIG.readingBodySelectors);
    if (!bodyEl) return null;
    const subjectEl = qFirst(CONFIG.readingSubjectSelectors);
    const subject = subjectEl ? subjectEl.textContent.trim() : "Email";
    const sender = findSenderName() || "Unknown sender";
    return { subject, sender, bodyText: bodyEl.innerText || "" };
  }

  function buildStyledEmailHTML(info) {
    const exportedAt = new Date().toLocaleString();
    return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${escapeHtml(info.subject)}</title>
<style>
  @page { margin: 26mm 20mm; }
  * { box-sizing: border-box; }
  html, body { margin:0; padding:0; background:#fff; }
  body { font-family: "Segoe UI", Arial, sans-serif; color:#20242B; }
  .ia-topbar { height:6px; background: linear-gradient(90deg,#1B4F8C,#2F8FD1 50%,#20B999); }
  .ia-doc { max-width:720px; margin:0 auto; padding:34px 40px 54px; }
  .ia-brand {
    display:flex; align-items:center; gap:7px; font-size:10.5px; letter-spacing:.09em;
    text-transform:uppercase; color:#8892A0; font-weight:700; margin-bottom:20px;
  }
  .ia-brand .dot { width:6px; height:6px; border-radius:50%; background:#20B999; display:inline-block; }
  h1.ia-subject {
    font-family: Georgia, "Times New Roman", serif; font-size:24px; line-height:1.32;
    margin:0 0 18px; color:#151A22; font-weight:700;
  }
  .ia-meta-card {
    background:#F5F8FC; border:1px solid #E1E5EA; border-left:3px solid #1B4F8C;
    border-radius:6px; padding:12px 16px; margin-bottom:28px; font-size:12px; color:#4B5563;
  }
  .ia-meta-card div { margin:2px 0; }
  .ia-meta-card strong { color:#20242B; }
  .ia-body { font-size:13.5px; line-height:1.75; white-space:pre-wrap; color:#242A33; }
  .ia-footer {
    margin-top:44px; padding-top:14px; border-top:1px solid #E1E5EA; font-size:10px;
    color:#9AA3B0; display:flex; justify-content:space-between; letter-spacing:.02em;
  }
</style>
</head>
<body>
  <div class="ia-topbar"></div>
  <div class="ia-doc">
    <div class="ia-brand"><span class="dot"></span>Inbox Assistant &middot; Exported Email</div>
    <h1 class="ia-subject">${escapeHtml(info.subject)}</h1>
    <div class="ia-meta-card">
      <div><strong>From:</strong> ${escapeHtml(info.sender)}</div>
      <div><strong>Exported:</strong> ${escapeHtml(exportedAt)}</div>
    </div>
    <div class="ia-body">${escapeHtml(info.bodyText)}</div>
    <div class="ia-footer"><span>Exported from Outlook Web</span><span>Inbox Assistant</span></div>
  </div>
</body>
</html>`;
  }

  $("#ia-export-pdf").addEventListener("click", () => {
    const statusEl = $("#ia-export-status");
    const info = currentReadingEmail();
    if (!info) {
      statusEl.textContent = "Open an email in the reading pane first.";
      return;
    }
    const win = window.open("", "_blank");
    if (!win) {
      statusEl.textContent = "Pop-up blocked — allow pop-ups for Outlook to export.";
      return;
    }
    win.document.write(buildStyledEmailHTML(info));
    win.document.close();
    statusEl.textContent = 'Opened printable version — choose "Save as PDF" in the print dialog.';
    win.onload = () => win.print();
  });

  $("#ia-export-html").addEventListener("click", () => {
    const statusEl = $("#ia-export-status");
    const info = currentReadingEmail();
    if (!info) {
      statusEl.textContent = "Open an email in the reading pane first.";
      return;
    }
    downloadBlob(`${safeFilename(info.subject)}.html`, buildStyledEmailHTML(info), "text/html");
    statusEl.textContent = "Downloaded as HTML.";
  });

  $("#ia-export-md").addEventListener("click", () => {
    const statusEl = $("#ia-export-status");
    const info = currentReadingEmail();
    if (!info) {
      statusEl.textContent = "Open an email in the reading pane first.";
      return;
    }
    const md = `# ${info.subject}\n\n**From:** ${info.sender}  \n**Exported:** ${new Date().toLocaleString()}\n\n---\n\n${info.bodyText}\n`;
    downloadBlob(`${safeFilename(info.subject)}.md`, md, "text/markdown");
    statusEl.textContent = "Downloaded as Markdown.";
  });

  $("#ia-export-txt").addEventListener("click", () => {
    const statusEl = $("#ia-export-status");
    const info = currentReadingEmail();
    if (!info) {
      statusEl.textContent = "Open an email in the reading pane first.";
      return;
    }
    const txt = `${info.subject}\nFrom: ${info.sender}\nExported: ${new Date().toLocaleString()}\n\n${info.bodyText}\n`;
    downloadBlob(`${safeFilename(info.subject)}.txt`, txt, "text/plain");
    statusEl.textContent = "Downloaded as plain text.";
  });

  /* =========================================================
     4) SCHEDULE SEND QUICK-PICKER
     One-click presets (Tomorrow 8am, Monday 9am, Friday EOD),
     both injected next to Outlook's own Send button while
     composing AND available from the panel's Schedule tab.

     Fully client-side, no backend — but genuinely reliable, silent
     auto-scheduling isn't possible without one of two things this
     extension deliberately avoids: (a) depending on Outlook's
     undocumented, frequently-churning "Send later" dialog markup to
     blindly fill in a date, or (b) a background service worker that
     holds your unsent draft and fires it later, which would mean
     trusting a browser extension to hold and eventually transmit
     the contents of your email — more risk than a "quick preset"
     feature should carry. So instead this drives Outlook's OWN
     native Send-later flow as far as it can (opening the menu, the
     dialog, and — best effort — picking the right calendar day) and
     always leaves Outlook's own dialog open for you to confirm the
     final time and hit Outlook's own Send yourself. If any step
     can't find what it's looking for, it copies the computed
     date/time to your clipboard so you can paste/select it manually
     instead of guessing wrong silently.
     ========================================================= */
  const SCHEDULE_PRESETS = [
    { id: "tomorrow8", label: "Tomorrow 8am" },
    { id: "monday9", label: "Monday 9am" },
    { id: "fridayEod", label: "Friday EOD" }
  ];

  function computePresetDate(preset) {
    const now = new Date();
    if (preset === "tomorrow8") {
      const d = new Date(now);
      d.setDate(d.getDate() + 1);
      d.setHours(8, 0, 0, 0);
      return d;
    }
    if (preset === "monday9") {
      const d = new Date(now);
      const daysUntilMonday = ((1 - d.getDay()) + 7) % 7 || 7; // always the *next* Monday
      d.setDate(d.getDate() + daysUntilMonday);
      d.setHours(9, 0, 0, 0);
      return d;
    }
    if (preset === "fridayEod") {
      const d = new Date(now);
      const daysUntilFriday = ((5 - d.getDay()) + 7) % 7; // 0 if today IS Friday
      d.setDate(d.getDate() + daysUntilFriday);
      d.setHours(17, 0, 0, 0);
      if (d.getTime() <= now.getTime()) d.setDate(d.getDate() + 7); // this Friday's EOD already passed
      return d;
    }
    return now;
  }

  // Combines a native <input type="date"> value ("YYYY-MM-DD") and
  // <input type="time"> value ("HH:MM", minute-granularity — any time of
  // day, not restricted to fixed intervals) into a local Date, or null if
  // either input is empty/unparseable.
  function combineDateTimeInputs(dateStr, timeStr) {
    if (!dateStr || !timeStr) return null;
    const dateParts = dateStr.split("-").map(Number);
    const timeParts = timeStr.split(":").map(Number);
    if (dateParts.length < 3 || timeParts.length < 2) return null;
    const [y, m, d] = dateParts;
    const [hh, mm] = timeParts;
    if ([y, m, d, hh, mm].some((n) => Number.isNaN(n))) return null;
    return new Date(y, m - 1, d, hh, mm, 0, 0);
  }

  function formatPresetTarget(d) {
    return (
      d.toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" }) +
      " · " +
      d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" })
    );
  }

  function findComposeSendButton() {
    const boxes = qAllVisible(CONFIG.sendButtonSelectors);
    return boxes.length ? boxes[boxes.length - 1] : null;
  }

  async function attemptNativeScheduleSend(sendBtn, targetDate, statusEl) {
    // Block any attempt — preset, panel-tab custom time, or inline custom
    // time — to schedule a send in the past, regardless of entry point.
    if (!targetDate || isNaN(targetDate.getTime()) || targetDate.getTime() <= Date.now()) {
      statusEl.textContent = "That time is in the past — pick a time in the future.";
      return;
    }
    statusEl.textContent = "Opening Outlook's schedule menu…";
    const targetText = formatPresetTarget(targetDate);

    // Step 1: find a "more send options" caret near the Send button.
    let caret = null;
    for (const sel of CONFIG.scheduleSendMoreOptionsSelectors) {
      const candidates = qAllVisible([sel]);
      if (candidates.length) {
        caret =
          candidates.find((c) => Math.abs(c.getBoundingClientRect().left - sendBtn.getBoundingClientRect().right) < 80) ||
          candidates[0];
        break;
      }
    }
    if (!caret) {
      await copyToClipboard(targetText);
      statusEl.textContent = `Couldn't find Outlook's schedule-send arrow automatically. Target time (${targetText}) copied — look for the small arrow next to Send and paste it in.`;
      return;
    }
    caret.click();
    await new Promise((r) => setTimeout(r, 250));

    // Step 2: find a menu item mentioning "send later" / "schedule".
    const menuItems = Array.from(document.querySelectorAll('[role="menuitem"], [role="option"]'));
    const sendLater = menuItems.find((mi) => /send later|schedule send|custom time/i.test(mi.textContent || ""));
    if (!sendLater) {
      await copyToClipboard(targetText);
      statusEl.textContent = `Opened Outlook's send menu but couldn't find "Send later" automatically. Target time (${targetText}) copied to your clipboard.`;
      return;
    }
    sendLater.click();
    await new Promise((r) => setTimeout(r, 300));

    // Step 3: best effort — Fluent UI calendars label each day button with
    // the full accessible date, so look for one matching our target date.
    const fullLabel = targetDate.toLocaleDateString(undefined, {
      weekday: "long",
      month: "long",
      day: "numeric",
      year: "numeric"
    });
    const dayBtn = Array.from(document.querySelectorAll("button[aria-label]")).find((b) =>
      (b.getAttribute("aria-label") || "").toLowerCase().includes(fullLabel.toLowerCase())
    );
    await copyToClipboard(targetText);
    if (dayBtn) {
      dayBtn.click();
      statusEl.textContent = `Picked ${targetText} in Outlook's calendar — set the time and confirm "Send" in Outlook's own dialog.`;
    } else {
      statusEl.textContent = `Opened Outlook's schedule dialog — target time (${targetText}) copied to your clipboard since the date couldn't be auto-selected. Paste/select it, then confirm in Outlook's own dialog.`;
    }
  }

  function injectScheduleToolbar(sendBtn) {
    if (!sendBtn || sendBtn.dataset.iaScheduleInjected) return;
    const todayStr = new Date().toISOString().slice(0, 10);
    const container = document.createElement("span");
    container.setAttribute("data-ia-schedule-toolbar", "1");
    const wrap = document.createElement("span");
    wrap.style.cssText =
      "all:initial; display:inline-flex; gap:6px; align-items:center; margin:0 8px; font-family:'Segoe UI',Arial,sans-serif; vertical-align:middle;";
    wrap.innerHTML =
      SCHEDULE_PRESETS.map(
        (p) =>
          `<button data-preset="${p.id}" style="all:initial; cursor:pointer; font-family:inherit; font-size:11px; font-weight:600; color:#1B4F8C; background:#EDF1F7; border:1px solid #C9D6E8; border-radius:12px; padding:4px 10px;">🗓️ ${p.label}</button>`
      ).join("") +
      `<button data-preset="custom" style="all:initial; cursor:pointer; font-family:inherit; font-size:11px; font-weight:600; color:#1B4F8C; background:#fff; border:1px dashed #1B4F8C; border-radius:12px; padding:4px 10px;">⏱️ Custom…</button>`;

    // Any date + any minute-precision time — not limited to the three presets.
    const customPanel = document.createElement("span");
    customPanel.style.cssText =
      "all:initial; display:none; align-items:center; gap:6px; margin-top:6px; font-family:'Segoe UI',Arial,sans-serif;";
    customPanel.innerHTML = `
      <input type="date" min="${todayStr}" style="all:revert; font-size:11px; padding:3px 5px; border:1px solid #C9D6E8; border-radius:5px;" />
      <input type="time" step="60" style="all:revert; font-size:11px; padding:3px 5px; border:1px solid #C9D6E8; border-radius:5px;" />
      <button style="all:initial; cursor:pointer; font-family:inherit; font-size:11px; font-weight:600; color:#fff; background:#1B4F8C; border-radius:12px; padding:4px 10px;">Go</button>
    `;

    const status = document.createElement("span");
    status.style.cssText =
      "all:initial; display:block; font-family:'Segoe UI',Arial,sans-serif; font-size:10.5px; color:#5B6472; margin-top:4px; max-width:320px;";
    container.appendChild(wrap);
    container.appendChild(document.createElement("br"));
    container.appendChild(customPanel);
    container.appendChild(document.createElement("br"));
    container.appendChild(status);
    sendBtn.insertAdjacentElement("beforebegin", container);
    sendBtn.dataset.iaScheduleInjected = "1";

    wrap.querySelectorAll("button[data-preset]").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (btn.dataset.preset === "custom") {
          customPanel.style.display = customPanel.style.display === "none" ? "inline-flex" : "none";
          return;
        }
        const targetDate = computePresetDate(btn.dataset.preset);
        status.textContent = "Working…";
        await attemptNativeScheduleSend(sendBtn, targetDate, status);
      });
    });

    const dateInput = customPanel.querySelector('input[type="date"]');
    const timeInput = customPanel.querySelector('input[type="time"]');
    const goBtn = customPanel.querySelector("button");
    goBtn.addEventListener("click", async (e) => {
      e.preventDefault();
      e.stopPropagation();
      const combined = combineDateTimeInputs(dateInput.value, timeInput.value);
      if (!combined) {
        status.textContent = "Pick both a date and a time first.";
        return;
      }
      status.textContent = "Working…";
      await attemptNativeScheduleSend(sendBtn, combined, status);
    });
  }

  function scanForComposeSurfaces() {
    qAllVisible(CONFIG.sendButtonSelectors).forEach(injectScheduleToolbar);
  }
  new MutationObserver(scanForComposeSurfaces).observe(document.body, { childList: true, subtree: true });
  scanForComposeSurfaces();

  // Manual fallback from the panel tab, for whenever Outlook re-renders the
  // compose toolbar and our injected pills haven't reappeared yet.
  shadow.querySelectorAll("#tab-schedule [data-preset]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const statusEl = $("#ia-schedule-status");
      const sendBtn = findComposeSendButton();
      if (!sendBtn) {
        statusEl.textContent = "Open a compose or reply window first.";
        return;
      }
      const targetDate = computePresetDate(btn.dataset.preset);
      statusEl.textContent = "Working…";
      await attemptNativeScheduleSend(sendBtn, targetDate, statusEl);
    });
  });

  // Panel tab's own custom date/time picker (any minute, past times blocked
  // centrally inside attemptNativeScheduleSend).
  const scheduleCustomDateEl = $("#ia-schedule-custom-date");
  if (scheduleCustomDateEl) scheduleCustomDateEl.min = new Date().toISOString().slice(0, 10);
  $("#ia-schedule-custom-go").addEventListener("click", async () => {
    const statusEl = $("#ia-schedule-status");
    const sendBtn = findComposeSendButton();
    if (!sendBtn) {
      statusEl.textContent = "Open a compose or reply window first.";
      return;
    }
    const combined = combineDateTimeInputs($("#ia-schedule-custom-date").value, $("#ia-schedule-custom-time").value);
    if (!combined) {
      statusEl.textContent = "Pick both a date and a time first.";
      return;
    }
    statusEl.textContent = "Working…";
    await attemptNativeScheduleSend(sendBtn, combined, statusEl);
  });
})();
