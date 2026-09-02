/* InboxSentry for Outlook — content.js
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
    themeStorageKey: "ia_theme_v1",
    densityStorageKey: "ia_density_v1",
    panelPosStorageKey: "ia_panel_pos_v1",
    snoozeStorageKey: "ia_stale_snoozes_v1",
    staleDaysStorageKey: "ia_stale_days_v1",
    // How many days an unread email can sit untouched before we flag it as
    // stale. Priority senders get their own (shorter) threshold — see
    // priorityStaleThresholdDays / ia_stale_priority_days_v1 below — because
    // an important sender going quiet for 3 days is a bigger deal than a
    // newsletter doing the same.
    staleThresholdDays: 3,
    priorityStaleThresholdDays: 1,
    // New in 1.5.0 — keyword watch, quick templates, attachment reminder.
    keywordStorageKey: "ia_keywords_v1",
    templatesStorageKey: "ia_templates_v1",
    attachRemStorageKey: "ia_attach_reminder_enabled_v1",
    // Compose surface's editable body. Best-effort like everything else in
    // this file — used by Templates (insert at cursor) and the Attachment
    // reminder (scanning what was typed for the word "attach…").
    composeBodySelectors: [
      'div[aria-label="Message body"]',
      'div[role="textbox"][contenteditable="true"][aria-label*="body" i]',
      'div[contenteditable="true"][aria-label*="body" i]'
    ],
    // Any element OWA renders per-attachment on a compose surface (a chip,
    // a "remove attachment" button, etc.) — used only to decide whether the
    // Attachment reminder should speak up; never depended on for anything
    // that changes the email itself.
    composeAttachmentSelectors: [
      '[aria-label*="Remove attachment" i]',
      '[data-testid*="attachment" i]',
      'div[role="button"][aria-label*="attachment" i]'
    ]
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
  // Priority-sender entries are stored as { value, level } objects, where
  // level is "high" or "normal". migratePriorityEntry() upgrades the old
  // plain-string format (v1.3.0 and earlier) transparently the first time
  // it's read, so nobody loses their existing list on update.
  function migratePriorityEntry(entry) {
    if (typeof entry === "string") return { value: entry, level: "normal" };
    if (entry && typeof entry === "object" && entry.value) {
      return { value: entry.value, level: entry.level === "high" ? "high" : "normal" };
    }
    return null;
  }
  function migratePriorityList(list) {
    return (list || []).map(migratePriorityEntry).filter(Boolean);
  }
  // Returns the matching entry (so callers can read its level) or null.
  // A leading "@" in a stored value (e.g. "@bigclient.com") matches any
  // sender text containing that domain, which in practice flags every
  // message from anyone at that company.
  function priorityMatch(sender, prioritySenders) {
    if (!sender || !prioritySenders || prioritySenders.length === 0) return null;
    const s = sender.toLowerCase();
    let best = null;
    for (const p of prioritySenders) {
      if (!p || !p.value) continue;
      if (s.includes(p.value.toLowerCase())) {
        if (p.level === "high") return p; // high always wins outright
        if (!best) best = p;
      }
    }
    return best;
  }
  // Keywords used to be stored as a plain string[]. Upgrade each entry to
  // {text, mode} the first time the extension runs with the new version —
  // existing keywords default to "substring" (their old behavior), so
  // nothing changes for anyone until they explicitly pick a different mode.
  function migrateKeywordList(list) {
    return (list || [])
      .map((k) => {
        if (typeof k === "string") return { text: k, mode: "substring" };
        if (k && typeof k.text === "string") return { text: k.text, mode: k.mode || "substring" };
        return null;
      })
      .filter((k) => k && k.text);
  }
  // Returns the matching entry (so callers can show which keyword/mode hit)
  // or null. "substring" is a plain includes(); "whole" requires the match
  // to sit on a word boundary; "regex" treats the keyword text as a user
  // regex pattern (case-insensitive), skipped silently if it fails to
  // compile rather than throwing and breaking the scan for every row.
  function keywordMatch(haystackLower, keywords) {
    if (!haystackLower || !keywords || keywords.length === 0) return null;
    for (const k of keywords) {
      if (!k || !k.text) continue;
      if (k.mode === "regex") {
        try {
          const re = new RegExp(k.text, "i");
          if (re.test(haystackLower)) return k;
        } catch {
          // invalid pattern — skip rather than throw
        }
      } else if (k.mode === "whole") {
        const escaped = k.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp("\\b" + escaped + "\\b", "i");
        if (re.test(haystackLower)) return k;
      } else {
        if (haystackLower.includes(k.text)) return k;
      }
    }
    return null;
  }

  /* ============================== SHARED STATE ============================== */
  // Declared up top (rather than inline in each section) so every section can
  // safely reference these regardless of which order the sections initialize in.
  let staleThresholdDays = CONFIG.staleThresholdDays;
  let priorityStaleThresholdDays = CONFIG.priorityStaleThresholdDays;
  let priorityCache = []; // [{value, level}]
  let snoozeCache = {}; // { [uid]: isoExpiryString }
  let lastStaleItems = []; // [{uid, subject, sender, ageDays, isPriority, level}] from the most recent scan
  let liveTimerInterval = null;
  let liveStartedAt = null;
  let liveElapsedBeforePause = 0;
  let keywordCache = []; // [{text, mode}, ...] lowercase-text keyword entries for the Keywords tab
                          // mode is one of "substring" | "whole" | "regex"
  let lastComposeField = null; // most recently focused compose contenteditable, for Templates "Insert"

  // Injected once into the real page (not the shadow root) so badges we add to
  // Outlook's own list rows render correctly; scoped with an "ia-" prefix to
  // avoid colliding with Outlook's own classes.
  const globalStyle = document.createElement("style");
  globalStyle.textContent = `
    .ia-row-badges { display:inline-flex; gap:4px; margin-left:8px; vertical-align:middle; }
    .ia-badge { display:inline-block; font-family:"Segoe UI",Arial,sans-serif; font-size:11.5px;
      font-weight:600; padding:1px 7px; border-radius:4px; line-height:16px; white-space:nowrap;
      letter-spacing:.1px; }
    .ia-badge-stale { background:#FDECEA; color:#A32C1E; border:1px solid #F3C6BE; }
    .ia-badge-stale-priority { background:#A32C1E; color:#fff; border:1px solid #A32C1E; }
    .ia-badge-priority-normal { background:#EEF2FB; color:#24408E; border:1px solid #C8D3EE; }
    .ia-badge-priority-high {
      background:#24408E; color:#fff; border:1px solid var(--ia-gold-soft, #D9A441);
      box-shadow: 0 0 0 1px rgba(184,134,11,.35);
      animation: ia-badge-pulse 2.4s ease-in-out infinite;
    }
    .ia-badge-keyword { background:#2C8C7A; color:#fff; border:1px solid #2C8C7A; }
    /* Flat, business-style side accent instead of a glowing/pulsing wash —
       a thin solid rail is enough signal without feeling like a toy. */
    .ia-row-priority-normal { box-shadow: inset 3px 0 0 #6C86C4 !important; }
    .ia-row-priority-high { box-shadow: inset 3px 0 0 #24408E, inset 0 0 0 1px rgba(184,134,11,.18) !important; }
    @keyframes ia-badge-pulse {
      0%, 100% { opacity: 1; }
      50% { opacity: .72; }
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
      :host {
        all: initial;
        /* ---- Light theme (default) — navy/slate corporate palette ---- */
        --ia-bg: #FFFFFF;
        --ia-surface: #F6F7FA;
        --ia-surface-2: #EDF0F5;
        --ia-border: #DFE3EA;
        --ia-border-strong: #C7CEDA;
        --ia-text: #172033;
        --ia-heading: #0F1729;
        --ia-muted: #64708A;
        --ia-primary: #1E3A73;
        --ia-primary-dark: #142A57;
        --ia-primary-light: #2E52A0;
        --ia-primary-contrast: #FFFFFF;
        --ia-primary-tint: #EAF0FB;
        --ia-accent: #3D6FD4;
        --ia-success-bg: #E6F4EC; --ia-success-text: #1C6B3B;
        --ia-warn-bg: #FBF0DD; --ia-warn-text: #8A5A0C;
        --ia-danger-bg: #FBEAE8; --ia-danger-text: #9C2D20;
        --ia-shadow: 0 16px 40px rgba(15,23,41,.18), 0 2px 8px rgba(15,23,41,.08);
        --ia-shadow-sm: 0 1px 2px rgba(15,23,41,.08);
        --ia-focus-ring: 0 0 0 3px rgba(30,58,115,.16);
        /* Restrained gold accent — used only for High-priority signal and a
           couple of small decorative touches, never as a base UI color. */
        --ia-gold: #B8860B;
        --ia-gold-soft: #D9A441;
        --ia-gold-tint: #FBF3DF;
        /* Density: normal by default, tightened by [data-density="compact"]. */
        --ia-pad-md: 14px;
        --ia-pad-sm: 10px;
        --ia-font-base: 13px;
        --ia-font-sm: 11.5px;
        --ia-row-gap: 10px;
      }
      :host([data-density="compact"]) {
        --ia-pad-md: 8px;
        --ia-pad-sm: 6px;
        --ia-font-base: 11.5px;
        --ia-font-sm: 10.5px;
        --ia-row-gap: 6px;
      }
      :host([data-theme="dark"]) {
        --ia-bg: #161A24;
        --ia-surface: #1C212D;
        --ia-surface-2: #242A38;
        --ia-border: #323A4A;
        --ia-border-strong: #414B5E;
        --ia-text: #DDE2ED;
        --ia-heading: #F2F4F9;
        --ia-muted: #8D97AD;
        --ia-primary: #6E93E8;
        --ia-primary-dark: #557AD1;
        --ia-primary-light: #8AACF2;
        --ia-primary-contrast: #0D111A;
        --ia-primary-tint: #202B45;
        --ia-accent: #7DA0F0;
        --ia-success-bg: #16301F; --ia-success-text: #72CB92;
        --ia-warn-bg: #392910; --ia-warn-text: #E8AE5C;
        --ia-danger-bg: #391E1B; --ia-danger-text: #F0897A;
        --ia-shadow: 0 16px 40px rgba(0,0,0,.55), 0 2px 8px rgba(0,0,0,.3);
        --ia-shadow-sm: 0 1px 2px rgba(0,0,0,.35);
        --ia-focus-ring: 0 0 0 3px rgba(110,147,232,.22);
        --ia-gold: #D9A441;
        --ia-gold-soft: #E8C077;
        --ia-gold-tint: #3A2F14;
      }
      * { box-sizing: border-box; font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, "Inter", Arial, sans-serif; }
      ::selection { background: var(--ia-primary-tint); color: var(--ia-primary-dark); }

      /* ---- Floating launcher ---- */
      .toggle {
        width: 52px; height: 52px; border-radius: 14px;
        background: linear-gradient(155deg, var(--ia-primary-light), var(--ia-primary) 60%, var(--ia-primary-dark));
        color: var(--ia-primary-contrast);
        border:none; cursor:grab; box-shadow: var(--ia-shadow); font-size:14px;
        font-weight:700; position:relative; display:flex; align-items:center; justify-content:center;
        transition: transform .15s ease, box-shadow .15s ease; letter-spacing:.4px;
        touch-action:none; user-select:none;
      }
      .toggle:hover { transform: translateY(-1px) scale(1.03); box-shadow: var(--ia-shadow), 0 0 0 3px rgba(184,134,11,.18); }
      .toggle.dragging { cursor:grabbing; transform: scale(1.06); transition:none; }
      .toggle span:first-child { font-family: Georgia, "Times New Roman", serif; font-size:17px; font-weight:700; }
      .toggle-badge {
        position:absolute; top:-5px; right:-5px; background: var(--ia-danger-text); color:#fff; border-radius:9px;
        min-width:18px; height:18px; display:none; align-items:center; justify-content:center;
        font-size:11px; font-weight:700; padding:0 4px; box-shadow:0 0 0 2.5px var(--ia-bg);
      }

      /* ---- Panel shell ----
         Absolutely positioned relative to the host (not the viewport), and
         always opens upward from the launcher (bottom:62px puts the panel's
         bottom edge just above the button) regardless of where the host has
         been dragged to — so the launcher always ends up at the panel's
         bottom-right corner, the familiar floating-widget layout. Because
         the launcher and panel share the same fixed host, dragging the
         host moves both together with no separate logic needed. */
      .panel {
        display:none; position:absolute; bottom:62px; right:0; width:452px; max-height:82vh;
        background: var(--ia-bg); border-radius:12px; box-shadow: var(--ia-shadow);
        overflow:hidden; flex-direction:column; border:1px solid var(--ia-border); color: var(--ia-text);
      }
      .panel.open { display:flex; }
      header {
        background: linear-gradient(120deg, var(--ia-primary-dark), var(--ia-primary) 55%, var(--ia-primary-light));
        color: var(--ia-primary-contrast); position:relative;
        padding: var(--ia-pad-md) 16px; display:flex; align-items:center; justify-content:space-between; gap:8px;
        flex-shrink:0;
      }
      header::after {
        content:""; position:absolute; left:0; right:0; bottom:0; height:3px;
        background: linear-gradient(90deg, var(--ia-gold-soft) 0%, rgba(255,255,255,.45) 18%, rgba(255,255,255,0) 60%);
      }
      .brand-row { display:flex; align-items:center; gap:10px; min-width:0; }
      .brand-mark {
        width:30px; height:30px; border-radius:8px; flex-shrink:0; display:flex; align-items:center; justify-content:center;
        background: rgba(255,255,255,.14); border:1px solid var(--ia-gold-soft);
        font-family: Georgia, "Times New Roman", serif; font-weight:700; font-size:15px; letter-spacing:.2px;
      }
      .header-actions { display:flex; align-items:center; gap:6px; flex-shrink:0; }
      .density-toggle {
        all:initial; cursor:pointer; font-family:inherit; width:27px; height:27px; border-radius:7px;
        display:flex; align-items:center; justify-content:center; font-size:13px;
        background: rgba(255,255,255,.14); color: var(--ia-primary-contrast);
        transition: background .15s ease; border:1px solid rgba(255,255,255,.22);
      }
      .density-toggle:hover { background: rgba(255,255,255,.26); }
      header h1 { margin:0; font-size:16.5px; font-weight:700; letter-spacing:.15px; }
      header p { margin:1px 0 0; font-size:11px; opacity:.82; text-transform:uppercase; letter-spacing:.5px; }
      .theme-toggle {
        all:initial; cursor:pointer; font-family:inherit; width:27px; height:27px; border-radius:7px;
        display:flex; align-items:center; justify-content:center; font-size:15px;
        background: rgba(255,255,255,.14); color: var(--ia-primary-contrast); flex-shrink:0;
        transition: background .15s ease; border:1px solid rgba(255,255,255,.22);
      }
      .theme-toggle:hover { background: rgba(255,255,255,.26); }

      /* ---- Body: icon rail + content ---- */
      .panel-body { display:flex; flex:1; min-height:0; }
      nav {
        display:flex; flex-direction:column; align-items:stretch; flex-shrink:0; width:70px;
        background: var(--ia-surface); border-right:1px solid var(--ia-border);
        overflow-y:auto; padding:6px 0;
      }
      nav button {
        border:none; background:none; padding: var(--ia-pad-sm) 2px; font-size:10.5px; cursor:pointer;
        color: var(--ia-muted); white-space:nowrap; position:relative;
        display:flex; flex-direction:column; align-items:center; gap:3px; font-family:inherit; font-weight:600;
        transition: color .12s ease, background .12s ease; border-left:2.5px solid transparent;
      }
      :host([data-density="compact"]) nav { width:56px; }
      :host([data-density="compact"]) nav button .nav-icon { font-size:14px; width:22px; height:22px; }
      nav button .nav-icon {
        font-size:17px; line-height:1; width:28px; height:28px; border-radius:7px;
        display:flex; align-items:center; justify-content:center; transition: background .12s ease;
      }
      nav button:hover { color: var(--ia-primary); background: var(--ia-surface-2); }
      nav button.active {
        color: var(--ia-primary); font-weight:700; background: var(--ia-bg); border-left-color: var(--ia-primary);
      }
      nav button.active .nav-icon { background: var(--ia-primary-tint); }
      main {
        flex:1; min-width:0; padding: var(--ia-pad-md) 16px 16px; overflow-y:auto; font-size:14.5px;
        color: var(--ia-text); background: var(--ia-bg);
      }
      :host([data-density="compact"]) main { font-size:12.5px; }
      main::-webkit-scrollbar, nav::-webkit-scrollbar, ul.contact-full-list-items::-webkit-scrollbar { width:8px; }
      main::-webkit-scrollbar-thumb, nav::-webkit-scrollbar-thumb, ul.contact-full-list-items::-webkit-scrollbar-thumb {
        background: var(--ia-border-strong); border-radius:8px;
      }
      main::-webkit-scrollbar-track, nav::-webkit-scrollbar-track { background: transparent; }
      section { display:none; }
      section.active { display:block; animation: ia-fade-in .12s ease; }
      @keyframes ia-fade-in { from { opacity:.4; } to { opacity:1; } }

      h2 {
        font-size:12.5px; margin:0 0 7px; padding-bottom:6px; font-weight:700; color: var(--ia-heading);
        text-transform:uppercase; letter-spacing:.5px; border-bottom:1px solid var(--ia-border);
      }
      .muted { color: var(--ia-muted); font-size:12.5px; line-height:1.55; }
      label { display:block; font-size:12px; font-weight:700; margin:9px 0 3px; color: var(--ia-heading); text-transform:uppercase; letter-spacing:.3px; }
      input, select {
        width:100%; padding:6px 8px; border:1px solid var(--ia-border); border-radius:6px; font-size:14px;
        background: var(--ia-bg); color: var(--ia-text); font-family:inherit; transition: border-color .12s ease, box-shadow .12s ease;
      }
      input:focus, select:focus { outline:none; border-color: var(--ia-primary); box-shadow: var(--ia-focus-ring); }
      input[type="checkbox"] { width:15px; height:15px; accent-color: var(--ia-primary); }

      button.primary {
        background: linear-gradient(135deg, var(--ia-primary-light), var(--ia-primary)); color: var(--ia-primary-contrast);
        border:none; border-radius:7px;
        padding:8px 14px; font-size:13px; font-weight:700; cursor:pointer; margin-top:8px;
        box-shadow: var(--ia-shadow-sm); transition: filter .12s ease, transform .08s ease; font-family:inherit;
        letter-spacing:.1px;
      }
      button.primary:hover { filter:brightness(1.08); }
      button.primary:active { transform: translateY(1px); }
      button.secondary {
        background: var(--ia-bg); color: var(--ia-primary); border:1px solid var(--ia-border-strong); border-radius:7px; padding:6px 11px;
        font-size:12px; cursor:pointer; margin-right:6px; margin-top:6px; font-weight:700; font-family:inherit;
        transition: background .12s ease, border-color .12s ease;
      }
      button.secondary:hover { background: var(--ia-surface-2); border-color: var(--ia-primary); }
      button.secondary.danger { color: var(--ia-danger-text); }
      button.secondary.danger:hover { background: var(--ia-danger-bg); border-color: var(--ia-danger-text); }

      .item {
        border:1px solid var(--ia-border); border-radius:8px; padding:9px 10px; margin-bottom: var(--ia-row-gap);
        background: var(--ia-surface); box-shadow: var(--ia-shadow-sm);
      }
      .pill { display:inline-block; padding:2px 8px; border-radius:20px; font-size:11px; font-weight:700; letter-spacing:.2px; }
      .pill.warn { background: var(--ia-warn-bg); color: var(--ia-warn-text); }
      .pill.ok { background: var(--ia-success-bg); color: var(--ia-success-text); }
      .pill.due { background: var(--ia-danger-bg); color: var(--ia-danger-text); }
      .pill.high { background: var(--ia-primary); color: var(--ia-primary-contrast); }
      .pill.normal { background: var(--ia-primary-tint); color: var(--ia-primary); }
      .empty { color: var(--ia-muted); font-style:italic; font-size:12.5px; padding:10px 0; }
      .warning-list { margin:0; padding-left:16px; }
      .warning-list li { margin-bottom:5px; }
      .row-flex { display:flex; gap:7px; align-items:flex-end; }
      .row-flex input, .row-flex select { flex:1; }
      .tag-list { list-style:none; margin:8px 0 0; padding:0; }
      .tag-list li {
        display:flex; justify-content:space-between; align-items:center; gap:6px;
        border:1px solid var(--ia-border); border-radius:7px; padding: var(--ia-pad-sm) 9px; margin-bottom: var(--ia-row-gap); font-size: var(--ia-font-base);
        background: var(--ia-surface); transition: border-color .12s ease;
      }
      .tag-list li:hover { border-color: var(--ia-border-strong); }
      .tag-list li .tag-left { display:flex; align-items:center; gap:6px; overflow:hidden; }
      .tag-list li .tag-left span.name { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
      .tag-list button {
        background:none; border:none; color: var(--ia-danger-text); cursor:pointer; font-size:12px;
        font-weight:700; font-family:inherit; flex-shrink:0; padding:3px 6px; border-radius:5px;
        transition: background .12s ease;
      }
      .tag-list button:hover { background: var(--ia-danger-bg); }
      .stat-card {
        background: var(--ia-surface); border:1px solid var(--ia-border); border-radius:9px;
        padding:14px 16px; margin-top:8px; text-align:center; box-shadow: var(--ia-shadow-sm);
      }
      .stat-big { font-size:28.5px; font-weight:800; color: var(--ia-primary); letter-spacing:-.2px; }
      .donut-row { display:flex; gap:16px; align-items:center; }
      .donut-legend { list-style:none; margin:0; padding:0; font-size:12.5px; flex:1; }
      .donut-legend li { display:flex; align-items:center; gap:6px; margin-bottom:5px; color: var(--ia-text); }
      .dot { width:9px; height:9px; border-radius:2px; display:inline-block; flex-shrink:0; }
      details.contact-full-list { margin-top:10px; border:1px solid var(--ia-border); border-radius:8px; overflow:hidden; }
      details.contact-full-list summary {
        cursor:pointer; padding:8px 10px; font-size:12px; font-weight:700; color: var(--ia-primary); list-style:none;
        text-transform:uppercase; letter-spacing:.3px; background: var(--ia-surface);
      }
      details.contact-full-list summary::-webkit-details-marker { display:none; }
      ul.contact-full-list-items {
        list-style:none; margin:0; padding:2px 10px 10px; max-height:180px; overflow-y:auto;
      }
      ul.contact-full-list-items li {
        display:flex; justify-content:space-between; font-size:12.5px; padding:5px 0; border-top:1px solid var(--ia-border); color: var(--ia-text);
      }
      ul.contact-full-list-items li:first-child { border-top:none; }
      .field-hint { font-size:11.5px; color: var(--ia-muted); margin-top:4px; line-height:1.45; }
      .stat-inline { display:flex; gap:8px; margin:8px 0 2px; }
      .stat-inline .chip {
        flex:1; text-align:center; background: var(--ia-surface); border:1px solid var(--ia-border);
        border-radius:8px; padding:8px 4px;
      }
      .stat-inline .chip b { display:block; font-size:18.5px; color: var(--ia-primary); font-weight:800; }
      .stat-inline .chip span { font-size:10.5px; color: var(--ia-muted); text-transform:uppercase; letter-spacing:.3px; }
    </style>
    <button class="toggle" id="ia-toggle" title="InboxSentry — click to open, drag to move">
      <span>IS</span>
      <span class="toggle-badge" id="ia-toggle-badge">0</span>
    </button>
    <div class="panel" id="ia-panel">
      <header>
        <div class="brand-row">
          <div class="brand-mark">IS</div>
          <div><h1>InboxSentry</h1><p>Inbox Management Suite</p></div>
        </div>
        <div class="header-actions">
          <button class="density-toggle" id="ia-density-toggle" title="Toggle compact / normal density">▤</button>
          <button class="theme-toggle" id="ia-theme-toggle" title="Toggle light / dark theme">◐</button>
        </div>
      </header>
      <div class="panel-body">
      <nav id="ia-tabs">
        <button data-tab="followup" class="active"><span class="nav-icon">📌</span>Follow-ups</button>
        <button data-tab="stale"><span class="nav-icon">⏰</span>Stale</button>
        <button data-tab="priority"><span class="nav-icon">★</span>Priority</button>
        <button data-tab="contacts"><span class="nav-icon">📊</span>Contacts</button>
        <button data-tab="cost"><span class="nav-icon">💰</span>Cost</button>
        <button data-tab="export"><span class="nav-icon">📤</span>Export</button>
        <button data-tab="schedule"><span class="nav-icon">🗓️</span>Schedule</button>
        <button data-tab="keywords"><span class="nav-icon">🔑</span>Keywords</button>
        <button data-tab="compose"><span class="nav-icon">✍️</span>Compose</button>
        <button data-tab="insights"><span class="nav-icon">📖</span>Insights</button>
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
          <h2>Unopened &gt; threshold</h2>
          <p class="muted">Automatically scans the message list and flags any unread email that's sat unopened past the threshold — look for the <span class="pill due">⏰ Nd</span> badge directly on the row. Opening the email clears its badge right away.</p>
          <div class="stat-inline">
            <div class="chip"><b id="ia-stale-count">0</b><span>flagged now</span></div>
            <div class="chip"><b id="ia-stale-priority-count">0</b><span>are priority</span></div>
            <div class="chip"><b id="ia-stale-snoozed-count">0</b><span>snoozed</span></div>
          </div>
          <label>Flag unread emails older than (days)</label>
          <div class="row-flex">
            <input type="number" id="ia-stale-days" min="1" max="60" value="3" />
            <button class="primary" id="ia-stale-save" style="margin-top:0;">Save</button>
          </div>
          <label>Flag priority senders sooner, after (days)</label>
          <div class="row-flex">
            <input type="number" id="ia-stale-priority-days" min="0" max="60" value="1" />
            <button class="primary" id="ia-stale-priority-save" style="margin-top:0;">Save</button>
          </div>
          <p class="field-hint">Applies only to senders on your Priority list — everyone else still uses the threshold above.</p>
          <h2 style="margin-top:14px;">Currently flagged (<span id="ia-stale-count-2">0</span>)</h2>
          <div id="ia-stale-list"><p class="empty">Nothing flagged in the visible list right now.</p></div>
          <button class="secondary" id="ia-stale-rescan">Rescan visible list</button>
          <button class="secondary" id="ia-stale-clear-snoozes">Clear all snoozes</button>
        </section>

        <section id="tab-priority">
          <h2>Priority senders</h2>
          <p class="muted">Emails from these senders get a badge and a side-accent directly in the message list, automatically, and are flagged as stale sooner. Start a value with <b>@</b> (e.g. <code>@bigclient.com</code>) to match everyone at that domain.</p>
          <div class="stat-inline">
            <div class="chip"><b id="ia-priority-total">0</b><span>senders tracked</span></div>
            <div class="chip"><b id="ia-priority-visible">0</b><span>in current view</span></div>
          </div>
          <label>Add sender (name, email, or @domain.com)</label>
          <div class="row-flex">
            <input type="text" id="ia-priority-input" placeholder="e.g. jane@company.com" />
            <select id="ia-priority-level" style="flex:0 0 92px;">
              <option value="normal" selected>Normal</option>
              <option value="high">High</option>
            </select>
            <button class="primary" id="ia-priority-add" style="margin-top:0;">Add</button>
          </div>
          <button class="secondary" id="ia-priority-add-current">Add sender of open email</button>
          <label style="margin-top:12px;">Filter list</label>
          <input type="text" id="ia-priority-search" placeholder="Search priority senders…" />
          <ul class="tag-list" id="ia-priority-list"><li class="empty" style="border:none;">No priority senders yet.</li></ul>
        </section>

        <section id="tab-contacts">
          <h2>Who you hear from most</h2>
          <p class="muted">Each sender is counted once per unique message and never listed twice — scanning again only adds newly-loaded messages, so scroll to load more, open the folder you want (Inbox or Sent Items), then scan. Nothing leaves your browser.</p>
          <button class="primary" id="ia-contacts-scan">📊 Scan this list</button>
          <button class="secondary" id="ia-contacts-reset">Reset counts</button>
          <button class="secondary" id="ia-contacts-csv">⬇️ Download CSV</button>
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

        <section id="tab-keywords">
          <h2>Keyword watch</h2>
          <p class="muted">Add words or short phrases (e.g. <b>invoice</b>, <b>contract</b>, <b>urgent</b>). Any visible message whose subject or preview text contains one gets a <span class="ia-badge ia-badge-keyword" style="display:inline;">🔑 match</span> badge, automatically — no click required.</p>
          <div class="stat-inline">
            <div class="chip"><b id="ia-keyword-total">0</b><span>keywords tracked</span></div>
            <div class="chip"><b id="ia-keyword-visible">0</b><span>matches in view</span></div>
          </div>
          <label>Add a keyword or phrase</label>
          <div class="row-flex">
            <input type="text" id="ia-keyword-input" placeholder="e.g. invoice" />
            <select id="ia-keyword-mode" style="flex:0 0 118px;" title="Contains: matches anywhere, even inside a word (e.g. 'urgent' also matches 'urgently'). Whole word: only matches 'urgent' on its own. Regex: treat this as a regular expression.">
              <option value="substring" selected>Contains</option>
              <option value="whole">Whole word</option>
              <option value="regex">Regex</option>
            </select>
            <button class="primary" id="ia-keyword-add" style="margin-top:0;">Add</button>
          </div>
          <ul class="tag-list" id="ia-keyword-list"><li class="empty" style="border:none;">No keywords yet.</li></ul>
        </section>

        <section id="tab-compose">
          <h2>Quick reply templates</h2>
          <p class="muted">Save a snippet once, then click a compose field to place your cursor in it and hit <b>Insert</b> — it's typed in right where your cursor is. If no compose field is focused, <b>Copy</b> puts it on your clipboard instead so you can paste it yourself.</p>
          <label>Name</label>
          <input type="text" id="ia-template-name" placeholder="e.g. Meeting follow-up" />
          <label>Text</label>
          <input type="text" id="ia-template-text" placeholder="e.g. Thanks for your time today — here's a recap…" />
          <button class="primary" id="ia-template-add">Save template</button>
          <h2 style="margin-top:14px;">Saved (<span id="ia-template-count">0</span>)</h2>
          <div id="ia-template-list"><p class="empty">No templates saved yet.</p></div>

          <h2 style="margin-top:16px;">Attachment reminder</h2>
          <p class="muted">Before sending, if your message mentions "attach…" (or similar) but no attachment is found on the compose surface, InboxSentry asks you to confirm before it lets the send go through. It only ever asks — it never blocks a send you confirm.</p>
          <label style="display:flex; align-items:center; gap:7px; cursor:pointer;">
            <input type="checkbox" id="ia-attach-reminder-toggle" style="width:auto;" checked />
            <span style="font-weight:600;">Enabled</span>
          </label>
        </section>

        <section id="tab-insights">
          <h2>Open email at a glance</h2>
          <p class="muted">Quick stats for whatever's open in the reading pane right now — updates automatically as you switch emails.</p>
          <div class="stat-inline">
            <div class="chip"><b id="ia-insights-words">—</b><span>words</span></div>
            <div class="chip"><b id="ia-insights-read">—</b><span>min read</span></div>
            <div class="chip"><b id="ia-insights-links">—</b><span>link(s)</span></div>
          </div>
          <p class="muted" id="ia-insights-summary" style="margin-top:8px;">Open an email in the reading pane to see its stats.</p>
        </section>
      </main>
      </div>
    </div>
  `;

  const $ = (sel) => shadow.querySelector(sel);

  const toggleBtn = $("#ia-toggle");
  let suppressNextToggleClick = false;

  toggleBtn.addEventListener("click", () => {
    if (suppressNextToggleClick) {
      suppressNextToggleClick = false;
      return;
    }
    $("#ia-panel").classList.toggle("open");
  });

  /* ---- Theme (light / dark) ----
     Defaults to the OS/browser preference the first time the extension
     runs, then remembers whatever the person picks via the header toggle. */
  function applyTheme(theme) {
    host.setAttribute("data-theme", theme);
    const btn = $("#ia-theme-toggle");
    if (btn) btn.textContent = theme === "dark" ? "☀" : "☾";
  }
  async function loadTheme() {
    const stored = await storageGet(CONFIG.themeStorageKey);
    const theme = stored || (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
    applyTheme(theme);
  }
  $("#ia-theme-toggle").addEventListener("click", async (e) => {
    e.stopPropagation();
    const current = host.getAttribute("data-theme") === "dark" ? "dark" : "light";
    const next = current === "dark" ? "light" : "dark";
    applyTheme(next);
    await storageSet(CONFIG.themeStorageKey, next);
  });
  loadTheme();

  /* ---- Density (normal / compact) ----
     Shrinks panel padding/font-size for people who want to see more at
     once. Persisted like the theme; defaults to normal. */
  function applyDensity(density) {
    host.setAttribute("data-density", density);
    const btn = $("#ia-density-toggle");
    if (btn) btn.title = density === "compact" ? "Switch to normal density" : "Switch to compact density";
  }
  async function loadDensity() {
    const stored = await storageGet(CONFIG.densityStorageKey);
    applyDensity(stored === "compact" ? "compact" : "normal");
  }
  $("#ia-density-toggle").addEventListener("click", async (e) => {
    e.stopPropagation();
    const current = host.getAttribute("data-density") === "compact" ? "compact" : "normal";
    const next = current === "compact" ? "normal" : "compact";
    applyDensity(next);
    await storageSet(CONFIG.densityStorageKey, next);
  });
  loadDensity();

  /* ---- Drag-to-reposition ----
     The launcher button ("IS") is itself the drag handle — press and hold
     it, then move, to relocate the whole widget; a plain click (no real
     movement) still opens/closes the panel as before. This mirrors how
     floating chat/support widgets on the web are usually draggable, so it
     doesn't need its own separate control or explanation.

     The host is a fixed-position element normally anchored bottom-right
     via `bottom`/`right`. Dragging switches it to an explicit `left`/`top`
     (clamped to stay fully on-screen) and remembers that position across
     reloads. Because the launcher and the panel are both inside the same
     fixed host, moving the host moves both together — no separate logic
     needed to keep the open panel attached to the launcher. */
  function clampPos(x, y) {
    const w = host.offsetWidth || 460;
    const h = host.offsetHeight || 500;
    const maxX = Math.max(0, window.innerWidth - w);
    const maxY = Math.max(0, window.innerHeight - h);
    return { x: Math.min(Math.max(0, x), maxX), y: Math.min(Math.max(0, y), maxY) };
  }
  function applyPos(pos) {
    if (!pos) return;
    host.style.left = pos.x + "px";
    host.style.top = pos.y + "px";
    host.style.right = "auto";
    host.style.bottom = "auto";
  }
  async function loadPanelPos() {
    const stored = await storageGet(CONFIG.panelPosStorageKey);
    if (stored && typeof stored.x === "number" && typeof stored.y === "number") {
      applyPos(clampPos(stored.x, stored.y));
    }
  }
  (function setupToggleDrag() {
    const DRAG_THRESHOLD = 4; // px of movement before a press counts as a drag rather than a click
    let dragging = false;
    let moved = false;
    let pointerId = null;
    let startX = 0, startY = 0, originX = 0, originY = 0;

    function onPointerMove(e) {
      if (!dragging || e.pointerId !== pointerId) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      if (!moved && (Math.abs(dx) > DRAG_THRESHOLD || Math.abs(dy) > DRAG_THRESHOLD)) {
        moved = true;
        toggleBtn.classList.add("dragging");
        // Once a real drag starts, close the panel if open so the person
        // can see exactly where the launcher (and panel) will land.
        $("#ia-panel").classList.remove("open");
      }
      if (moved) {
        const pos = clampPos(originX + dx, originY + dy);
        applyPos(pos);
      }
    }
    async function onPointerUp(e) {
      if (!dragging || e.pointerId !== pointerId) return;
      dragging = false;
      toggleBtn.classList.remove("dragging");
      try { toggleBtn.releasePointerCapture(pointerId); } catch {}
      toggleBtn.removeEventListener("pointermove", onPointerMove);
      toggleBtn.removeEventListener("pointerup", onPointerUp);
      toggleBtn.removeEventListener("pointercancel", onPointerUp);
      if (moved) {
        const rect = host.getBoundingClientRect();
        const pos = clampPos(rect.left, rect.top);
        applyPos(pos);
        await storageSet(CONFIG.panelPosStorageKey, pos);
        // The pointerup that ends a drag also fires a click on the same
        // element right after — suppress just that one so dragging never
        // also toggles the panel open.
        suppressNextToggleClick = true;
      }
    }
    toggleBtn.addEventListener("pointerdown", (e) => {
      // Left click / primary touch only — ignore right-click, etc.
      if (e.button !== undefined && e.button !== 0) return;
      dragging = true;
      moved = false;
      pointerId = e.pointerId;
      const rect = host.getBoundingClientRect();
      originX = rect.left;
      originY = rect.top;
      startX = e.clientX;
      startY = e.clientY;
      try { toggleBtn.setPointerCapture(pointerId); } catch {}
      toggleBtn.addEventListener("pointermove", onPointerMove);
      toggleBtn.addEventListener("pointerup", onPointerUp);
      toggleBtn.addEventListener("pointercancel", onPointerUp);
    });
  })();
  loadPanelPos();

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
            <strong style="font-size:13px;">${escapeHtml(f.subject).slice(0, 60)}</strong>
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
    const stored = await storageGet(CONFIG.staleDaysStorageKey);
    staleThresholdDays = stored != null ? stored : CONFIG.staleThresholdDays;
    $("#ia-stale-days").value = staleThresholdDays;
    const storedPriority = await storageGet("ia_stale_priority_days_v1");
    priorityStaleThresholdDays = storedPriority != null ? storedPriority : CONFIG.priorityStaleThresholdDays;
    $("#ia-stale-priority-days").value = priorityStaleThresholdDays;
  }
  async function refreshPriorityCache() {
    priorityCache = migratePriorityList(await storageGet(CONFIG.priorityStorageKey));
  }
  async function refreshKeywordCache() {
    const migrated = migrateKeywordList(await storageGet(CONFIG.keywordStorageKey));
    keywordCache = migrated.map((k) => ({ text: k.text.toLowerCase(), mode: k.mode }));
  }
  async function refreshSnoozeCache() {
    const stored = (await storageGet(CONFIG.snoozeStorageKey)) || {};
    const now = Date.now();
    // Drop anything that already expired so the stored object doesn't grow forever.
    const live = {};
    Object.keys(stored).forEach((uid) => {
      if (new Date(stored[uid]).getTime() > now) live[uid] = stored[uid];
    });
    snoozeCache = live;
    if (Object.keys(live).length !== Object.keys(stored).length) {
      await storageSet(CONFIG.snoozeStorageKey, live);
    }
  }
  async function snoozeItem(uid, hours = 24) {
    if (!uid) return;
    const expiry = new Date(Date.now() + hours * 3600000).toISOString();
    snoozeCache[uid] = expiry;
    await storageSet(CONFIG.snoozeStorageKey, snoozeCache);
    scanMailList();
  }

  function applyRowSignals(row) {
    const old = row.querySelector(".ia-row-badges");
    if (old) old.remove();
    row.classList.remove("ia-row-priority-normal", "ia-row-priority-high");

    const unread = isRowUnread(row);
    const date = findRowDate(row);
    const sender = findRowSender(row);
    const match = priorityMatch(sender, priorityCache);
    const uid = rowUniqueId(row);
    const isSnoozed = uid && snoozeCache[uid] && new Date(snoozeCache[uid]).getTime() > Date.now();
    const badges = [];

    if (unread && date && !isSnoozed) {
      const ageDays = Math.floor((Date.now() - date.getTime()) / 86400000);
      const threshold = match ? Math.min(staleThresholdDays, priorityStaleThresholdDays) : staleThresholdDays;
      if (ageDays >= threshold) {
        const badgeClass = match ? "ia-badge-stale-priority" : "ia-badge-stale";
        const title = match ? `Priority sender, unopened for ${ageDays} day(s)` : `Unopened for ${ageDays} day(s)`;
        badges.push(`<span class="ia-badge ${badgeClass}" title="${title}">⏰ ${ageDays}d</span>`);
        lastStaleItems.push({
          uid,
          subject: (row.getAttribute("aria-label") || "").split(",")[0] || "(email)",
          sender: sender || "(unknown sender)",
          ageDays,
          isPriority: !!match,
          level: match ? match.level : null
        });
      }
    }
    if (match) {
      const badgeClass = match.level === "high" ? "ia-badge-priority-high" : "ia-badge-priority-normal";
      const label = match.level === "high" ? "🔺 Urgent" : "★ Priority";
      badges.push(`<span class="ia-badge ${badgeClass}" title="Priority sender (${match.level})">${label}</span>`);
      row.classList.add(match.level === "high" ? "ia-row-priority-high" : "ia-row-priority-normal");
    }
    if (keywordCache.length) {
      const haystack = (row.textContent || "").toLowerCase();
      const hit = keywordMatch(haystack, keywordCache);
      if (hit) {
        badges.push(`<span class="ia-badge ia-badge-keyword" title="Matched keyword (${escapeHtml(hit.mode)}): ${escapeHtml(hit.text)}">🔑 ${escapeHtml(hit.text)}</span>`);
      }
    }
    if (badges.length) {
      const wrap = document.createElement("span");
      wrap.className = "ia-row-badges";
      wrap.innerHTML = badges.join("");
      row.appendChild(wrap);
    }
  }

  function scanMailList() {
    // Pause the observer while we make our own DOM writes below, so
    // badge/class updates don't get picked up as "Outlook changed
    // something" and schedule another immediate rescan (see FIX note
    // on the observer above). It's reconnected right after.
    mailListObserver.disconnect();
    try {
      lastStaleItems = [];
      const container = qFirst(CONFIG.mailListContainerSelectors) || document.body;
      const rows = qAllVisible(CONFIG.mailListItemSelectors, container);
      rows.forEach(applyRowSignals);
      renderStaleTab();
    } finally {
      observeMailList();
    }
  }

  function renderStaleTab() {
    const countEl = $("#ia-stale-count");
    const countEl2 = $("#ia-stale-count-2");
    const priorityCountEl = $("#ia-stale-priority-count");
    const snoozedCountEl = $("#ia-stale-snoozed-count");
    const listEl = $("#ia-stale-list");
    if (countEl && listEl) {
      countEl.textContent = lastStaleItems.length;
      if (countEl2) countEl2.textContent = lastStaleItems.length;
      if (priorityCountEl) priorityCountEl.textContent = lastStaleItems.filter((i) => i.isPriority).length;
      if (snoozedCountEl) snoozedCountEl.textContent = Object.keys(snoozeCache).length;
      if (lastStaleItems.length === 0) {
        listEl.innerHTML = '<p class="empty">Nothing flagged in the visible list right now.</p>';
      } else {
        // Priority-flagged stale emails float to the top; ties broken by age.
        listEl.innerHTML = lastStaleItems
          .sort((a, b) => (b.isPriority - a.isPriority) || (b.ageDays - a.ageDays))
          .map((it, i) => {
            const pill = it.isPriority
              ? `<span class="pill due">🔥 ${it.ageDays}d &middot; priority</span>`
              : `<span class="pill due">${it.ageDays}d unopened</span>`;
            return `
          <div class="item" data-uid="${escapeHtml(it.uid || "")}" data-idx="${i}">
            <strong style="font-size:13px;">${escapeHtml(it.subject).slice(0, 60)}</strong>
            <div class="muted">${escapeHtml(it.sender)} &middot; ${pill}</div>
            ${it.uid ? '<button class="secondary" data-action="snooze">Snooze 1 day</button>' : ""}
          </div>`;
          })
          .join("");
        listEl.querySelectorAll('[data-action="snooze"]').forEach((btn) => {
          btn.addEventListener("click", () => {
            const uid = btn.closest(".item").dataset.uid;
            snoozeItem(uid, 24);
          });
        });
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
  //
  // FIX (flicker): applyRowSignals() itself writes to the DOM it's watching —
  // it removes/re-adds the ".ia-row-badges" node (a childList change) and
  // toggles "ia-row-priority-*" classes on the row (an attribute change in
  // attributeFilter). Those are exactly the mutations this observer listens
  // for, so every scan re-triggered another scan 500ms later, which redrew
  // the badge again, which re-triggered the observer again — an infinite
  // scan/redraw loop with no real change in between, seen as the badge
  // flickering (removed and recreated) on a ~500ms cadence forever. The fix
  // is to disconnect the observer for the duration of our own synchronous
  // DOM writes so only genuine Outlook-driven changes schedule a rescan.
  const mailListObserver = new MutationObserver(scheduleScan);
  function observeMailList() {
    mailListObserver.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["aria-label", "class", "title"]
    });
  }
  observeMailList();
  loadStaleThreshold().then(() =>
    refreshSnoozeCache().then(() => refreshPriorityCache().then(() => refreshKeywordCache().then(scanMailList)))
  );
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
    await storageSet(CONFIG.staleDaysStorageKey, days);
    scanMailList();
  });
  $("#ia-stale-priority-save").addEventListener("click", async () => {
    const days = Math.max(0, parseInt($("#ia-stale-priority-days").value, 10) || 0);
    priorityStaleThresholdDays = days;
    await storageSet("ia_stale_priority_days_v1", days);
    scanMailList();
  });
  $("#ia-stale-rescan").addEventListener("click", scanMailList);
  $("#ia-stale-clear-snoozes").addEventListener("click", async () => {
    snoozeCache = {};
    await storageSet(CONFIG.snoozeStorageKey, snoozeCache);
    scanMailList();
  });

  /* =========================================================
     1c) PRIORITY SENDERS
     ========================================================= */
  async function getPrioritySenders() {
    return migratePriorityList(await storageGet(CONFIG.priorityStorageKey));
  }
  async function savePrioritySenders(list) {
    await storageSet(CONFIG.priorityStorageKey, list);
    await refreshPriorityCache();
    scanMailList();
  }
  async function countPriorityInView() {
    const container = qFirst(CONFIG.mailListContainerSelectors) || document.body;
    const rows = qAllVisible(CONFIG.mailListItemSelectors, container);
    return rows.filter((r) => priorityMatch(findRowSender(r), priorityCache)).length;
  }
  async function renderPriorityList() {
    const list = await getPrioritySenders();
    const el = $("#ia-priority-list");
    const totalEl = $("#ia-priority-total");
    const visibleEl = $("#ia-priority-visible");
    if (totalEl) totalEl.textContent = list.length;
    if (visibleEl) visibleEl.textContent = await countPriorityInView();

    const filter = ($("#ia-priority-search") ? $("#ia-priority-search").value : "").trim().toLowerCase();
    // High priority first, then alphabetically within each level.
    const sorted = list
      .map((s, i) => ({ ...s, i }))
      .filter((s) => !filter || s.value.toLowerCase().includes(filter))
      .sort((a, b) => (b.level === "high") - (a.level === "high") || a.value.localeCompare(b.value));

    if (list.length === 0) {
      el.innerHTML = '<li class="empty" style="border:none;">No priority senders yet.</li>';
      return;
    }
    if (sorted.length === 0) {
      el.innerHTML = '<li class="empty" style="border:none;">No senders match that search.</li>';
      return;
    }
    el.innerHTML = sorted
      .map(
        (s) => `<li data-i="${s.i}">
          <span class="tag-left"><span class="pill ${s.level}">${s.level === "high" ? "High" : "Normal"}</span><span class="name">${escapeHtml(s.value)}</span></span>
          <button data-action="remove">Remove</button>
        </li>`
      )
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
  async function addPrioritySender(value, level) {
    const v = (value || "").trim();
    if (!v) return;
    const lvl = level === "high" ? "high" : "normal";
    const cur = await getPrioritySenders();
    const existingIdx = cur.findIndex((s) => s.value.toLowerCase() === v.toLowerCase());
    if (existingIdx >= 0) {
      cur[existingIdx].level = lvl; // re-adding an existing sender updates their level
    } else {
      cur.push({ value: v, level: lvl });
    }
    await savePrioritySenders(cur);
    renderPriorityList();
  }
  $("#ia-priority-add").addEventListener("click", () => {
    const input = $("#ia-priority-input");
    const level = $("#ia-priority-level").value;
    addPrioritySender(input.value, level);
    input.value = "";
  });
  $("#ia-priority-add-current").addEventListener("click", () => {
    const info = currentEmailInfo();
    if (!info || info.sender === "(sender not found)") return;
    addPrioritySender(info.sender, $("#ia-priority-level").value);
  });
  $("#ia-priority-search").addEventListener("input", () => renderPriorityList());
  renderPriorityList();

  /* =========================================================
     1c-ii) KEYWORD WATCH — same idea as priority senders, but
     matching on subject/preview text instead of who sent it.
     Badges are applied inside applyRowSignals() above, driven by
     keywordCache, so this section is just the storage + tab UI.
     ========================================================= */
  async function getKeywords() {
    return migrateKeywordList(await storageGet(CONFIG.keywordStorageKey));
  }
  async function saveKeywords(list) {
    await storageSet(CONFIG.keywordStorageKey, list);
    await refreshKeywordCache();
    scanMailList();
  }
  const KEYWORD_MODE_LABEL = { substring: "Contains", whole: "Whole word", regex: "Regex" };
  async function countKeywordMatchesInView() {
    if (!keywordCache.length) return 0;
    const container = qFirst(CONFIG.mailListContainerSelectors) || document.body;
    const rows = qAllVisible(CONFIG.mailListItemSelectors, container);
    return rows.filter((r) => {
      const haystack = (r.textContent || "").toLowerCase();
      return !!keywordMatch(haystack, keywordCache);
    }).length;
  }
  async function renderKeywordList() {
    const list = await getKeywords();
    const el = $("#ia-keyword-list");
    const totalEl = $("#ia-keyword-total");
    const visibleEl = $("#ia-keyword-visible");
    if (totalEl) totalEl.textContent = list.length;
    if (visibleEl) visibleEl.textContent = await countKeywordMatchesInView();
    if (list.length === 0) {
      el.innerHTML = '<li class="empty" style="border:none;">No keywords yet.</li>';
      return;
    }
    el.innerHTML = list
      .map(
        (k, i) => `<li data-i="${i}">
          <span class="tag-left"><span class="name">${escapeHtml(k.text)}</span><span class="pill normal" style="flex-shrink:0;">${escapeHtml(KEYWORD_MODE_LABEL[k.mode] || "Contains")}</span></span>
          <button data-action="remove">Remove</button>
        </li>`
      )
      .join("");
    el.querySelectorAll("li").forEach((li) => {
      const i = parseInt(li.dataset.i, 10);
      const removeBtn = li.querySelector('[data-action="remove"]');
      if (!removeBtn) return;
      removeBtn.addEventListener("click", async () => {
        const cur = await getKeywords();
        cur.splice(i, 1);
        await saveKeywords(cur);
        renderKeywordList();
      });
    });
  }
  $("#ia-keyword-add").addEventListener("click", async () => {
    const input = $("#ia-keyword-input");
    const modeSelect = $("#ia-keyword-mode");
    const v = (input.value || "").trim();
    if (!v) return;
    const mode = (modeSelect && modeSelect.value) || "substring";
    if (mode === "regex") {
      try {
        new RegExp(v, "i");
      } catch {
        input.style.borderColor = "var(--ia-danger-text)";
        input.title = "Not a valid regular expression";
        return;
      }
    }
    input.style.borderColor = "";
    input.title = "";
    const cur = await getKeywords();
    if (!cur.some((k) => k.text.toLowerCase() === v.toLowerCase() && k.mode === mode)) {
      cur.push({ text: v, mode });
      await saveKeywords(cur);
    }
    input.value = "";
    renderKeywordList();
  });
  $("#ia-keyword-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") $("#ia-keyword-add").click();
  });
  renderKeywordList();

  /* =========================================================
     1d) CONTACTS INSIGHTS — who you receive from / send to most.
     Built entirely from whatever message rows are currently loaded
     in the visible list (no Microsoft Graph API involved, so it's a
     snapshot of what's rendered — scroll to load more rows before
     scanning for a fuller picture). Rendered as pure-SVG donut
     charts (a stroke-dasharray ring trick) with no charting library.
     ========================================================= */
  const DONUT_COLORS = ["#24408E", "#2C8C7A", "#C77F1E", "#A32C1E", "#6C4FBF", "#4E7AC7", "#93590B", "#667085"];

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
    // Center hole/text colors follow the current theme so the donut doesn't
    // show a bright white disc in the middle of a dark panel.
    const isDark = host.getAttribute("data-theme") === "dark";
    const holeColor = isDark ? "#21242C" : "#fff";
    const textColor = isDark ? "#E7E9EE" : "#1A1D24";
    const subTextColor = isDark ? "#9AA1B0" : "#667085";
    return `<svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}">
      <g transform="rotate(-90 ${cx} ${cy})">${circles}</g>
      <circle cx="${cx}" cy="${cy}" r="${r - thickness / 2 - 3}" fill="${holeColor}"></circle>
      <text x="${cx}" y="${cy - 3}" text-anchor="middle" font-size="18.5" font-weight="800" fill="${textColor}">${total}</text>
      <text x="${cx}" y="${cy + 13}" text-anchor="middle" font-size="10" fill="${subTextColor}">emails</text>
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

  $("#ia-contacts-csv").addEventListener("click", async () => {
    const received = (await storageGet(CONFIG.contactsReceivedStorageKey)) || { counts: {} };
    const sent = (await storageGet(CONFIG.contactsSentStorageKey)) || { counts: {} };
    const rows = [["direction", "contact", "messages"]];
    Object.entries(received.counts || {})
      .sort((a, b) => b[1] - a[1])
      .forEach(([name, count]) => rows.push(["received", name, count]));
    Object.entries(sent.counts || {})
      .sort((a, b) => b[1] - a[1])
      .forEach(([name, count]) => rows.push(["sent", name, count]));
    if (rows.length === 1) {
      $("#ia-contacts-status").textContent = "Nothing to export yet — scan Inbox or Sent Items first.";
      return;
    }
    const csvEscape = (v) => `"${String(v).replace(/"/g, '""')}"`;
    const csv = rows.map((r) => r.map(csvEscape).join(",")).join("\r\n");
    downloadBlob(`inboxsentry-contacts-${new Date().toISOString().slice(0, 10)}.csv`, csv, "text/csv");
    $("#ia-contacts-status").textContent = "Downloaded contacts as CSV.";
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
            <strong style="font-size:13px;">${h.symbol}${h.cost.toFixed(0)}</strong>
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
  .ia-topbar { height:5px; background:#24408E; }
  .ia-doc { max-width:720px; margin:0 auto; padding:34px 40px 54px; }
  .ia-brand {
    display:flex; align-items:center; gap:7px; font-size:12px; letter-spacing:.09em;
    text-transform:uppercase; color:#8892A0; font-weight:700; margin-bottom:20px;
  }
  .ia-brand .dot { width:6px; height:6px; border-radius:50%; background:#24408E; display:inline-block; }
  h1.ia-subject {
    font-family: Georgia, "Times New Roman", serif; font-size:27.5px; line-height:1.32;
    margin:0 0 18px; color:#151A22; font-weight:700;
  }
  .ia-meta-card {
    background:#F5F8FC; border:1px solid #E1E5EA; border-left:3px solid #24408E;
    border-radius:6px; padding:12px 16px; margin-bottom:28px; font-size:14px; color:#4B5563;
  }
  .ia-meta-card div { margin:2px 0; }
  .ia-meta-card strong { color:#20242B; }
  .ia-body { font-size:15.5px; line-height:1.75; white-space:pre-wrap; color:#242A33; }
  .ia-footer {
    margin-top:44px; padding-top:14px; border-top:1px solid #E1E5EA; font-size:11.5px;
    color:#9AA3B0; display:flex; justify-content:space-between; letter-spacing:.02em;
  }
</style>
</head>
<body>
  <div class="ia-topbar"></div>
  <div class="ia-doc">
    <div class="ia-brand"><span class="dot"></span>InboxSentry &middot; Exported Email</div>
    <h1 class="ia-subject">${escapeHtml(info.subject)}</h1>
    <div class="ia-meta-card">
      <div><strong>From:</strong> ${escapeHtml(info.sender)}</div>
      <div><strong>Exported:</strong> ${escapeHtml(exportedAt)}</div>
    </div>
    <div class="ia-body">${escapeHtml(info.bodyText)}</div>
    <div class="ia-footer"><span>Exported from Outlook Web</span><span>InboxSentry</span></div>
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
          `<button data-preset="${p.id}" style="all:initial; cursor:pointer; font-family:inherit; font-size:12.5px; font-weight:600; color:#24408E; background:#EEF2FB; border:1px solid #C8D3EE; border-radius:6px; padding:4px 10px;">🗓️ ${p.label}</button>`
      ).join("") +
      `<button data-preset="custom" style="all:initial; cursor:pointer; font-family:inherit; font-size:12.5px; font-weight:600; color:#24408E; background:#fff; border:1px dashed #24408E; border-radius:6px; padding:4px 10px;">⏱️ Custom…</button>`;

    // Any date + any minute-precision time — not limited to the three presets.
    const customPanel = document.createElement("span");
    customPanel.style.cssText =
      "all:initial; display:none; align-items:center; gap:6px; margin-top:6px; font-family:'Segoe UI',Arial,sans-serif;";
    customPanel.innerHTML = `
      <input type="date" min="${todayStr}" style="all:revert; font-size:12.5px; padding:3px 5px; border:1px solid #C8D3EE; border-radius:5px;" />
      <input type="time" step="60" style="all:revert; font-size:12.5px; padding:3px 5px; border:1px solid #C8D3EE; border-radius:5px;" />
      <button style="all:initial; cursor:pointer; font-family:inherit; font-size:12.5px; font-weight:600; color:#fff; background:#24408E; border-radius:6px; padding:4px 10px;">Go</button>
    `;

    const status = document.createElement("span");
    status.style.cssText =
      "all:initial; display:block; font-family:'Segoe UI',Arial,sans-serif; font-size:12px; color:#667085; margin-top:4px; max-width:320px;";
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

  /* =========================================================
     5) QUICK REPLY TEMPLATES
     Saved snippets you can drop into whatever compose field last
     had focus. Uses document.execCommand("insertText"), the same
     approach browsers' own spellcheck-replace UI relies on, so it
     plays nicely with the compose box's own undo history — if it's
     ever unavailable (or nothing is focused), this falls back to
     copying the text to the clipboard instead, same "always leave
     the person with something usable" pattern as Schedule Send.
     ========================================================= */
  document.addEventListener(
    "focusin",
    (e) => {
      const field = e.target.closest ? e.target.closest(CONFIG.composeBodySelectors.join(",")) : null;
      if (field) lastComposeField = field;
    },
    true
  );

  async function getTemplates() {
    return (await storageGet(CONFIG.templatesStorageKey)) || [];
  }
  async function saveTemplates(list) {
    await storageSet(CONFIG.templatesStorageKey, list);
  }
  function insertIntoComposeField(field, text) {
    field.focus();
    try {
      const ok = document.execCommand("insertText", false, text);
      if (ok) return true;
    } catch (e) {
      /* fall through to clipboard */
    }
    return false;
  }
  async function renderTemplateList() {
    const list = await getTemplates();
    $("#ia-template-count").textContent = list.length;
    const el = $("#ia-template-list");
    if (list.length === 0) {
      el.innerHTML = '<p class="empty">No templates saved yet.</p>';
      return;
    }
    el.innerHTML = list
      .map(
        (t, i) => `<div class="item" data-i="${i}">
          <strong style="font-size:13px;">${escapeHtml(t.name)}</strong>
          <div class="muted" style="margin:3px 0 5px; white-space:pre-wrap;">${escapeHtml(t.text).slice(0, 140)}${t.text.length > 140 ? "…" : ""}</div>
          <button class="secondary" data-action="insert">Insert</button>
          <button class="secondary" data-action="copy">Copy</button>
          <button class="secondary danger" data-action="delete">Delete</button>
        </div>`
      )
      .join("");
    el.querySelectorAll(".item").forEach((item) => {
      const i = parseInt(item.dataset.i, 10);
      item.querySelector('[data-action="insert"]').addEventListener("click", async () => {
        const cur = await getTemplates();
        const t = cur[i];
        if (!t) return;
        const field = lastComposeField && document.contains(lastComposeField) ? lastComposeField : qFirst(CONFIG.composeBodySelectors);
        if (field && insertIntoComposeField(field, t.text)) return;
        const copied = await copyToClipboard(t.text);
        alert(copied ? "Couldn't find an open compose field — copied to your clipboard instead." : "Couldn't insert or copy — select the template text manually.");
      });
      item.querySelector('[data-action="copy"]').addEventListener("click", async () => {
        const cur = await getTemplates();
        const t = cur[i];
        if (!t) return;
        await copyToClipboard(t.text);
      });
      item.querySelector('[data-action="delete"]').addEventListener("click", async () => {
        const cur = await getTemplates();
        cur.splice(i, 1);
        await saveTemplates(cur);
        renderTemplateList();
      });
    });
  }
  $("#ia-template-add").addEventListener("click", async () => {
    const nameEl = $("#ia-template-name");
    const textEl = $("#ia-template-text");
    const name = (nameEl.value || "").trim();
    const text = (textEl.value || "").trim();
    if (!name || !text) return;
    const cur = await getTemplates();
    cur.push({ name, text });
    await saveTemplates(cur);
    nameEl.value = "";
    textEl.value = "";
    renderTemplateList();
  });
  renderTemplateList();

  /* =========================================================
     6) ATTACHMENT REMINDER
     Best-effort safety net, same spirit as Schedule Send: this
     NEVER blocks a send outright. If the open compose surface's
     text mentions "attach…" (or similar) but no attachment element
     can be found near it, a single confirm() gives the person a
     chance to double back — confirming sends immediately with no
     further interception.
     ========================================================= */
  const ATTACH_MENTION_RE = /\b(attach(ed|ment|ments)?|enclos(ed|ure))\b/i;
  let attachReminderEnabled = true;
  async function loadAttachReminderSetting() {
    const stored = await storageGet(CONFIG.attachRemStorageKey);
    attachReminderEnabled = stored !== false; // default on
    $("#ia-attach-reminder-toggle").checked = attachReminderEnabled;
  }
  $("#ia-attach-reminder-toggle").addEventListener("change", async (e) => {
    attachReminderEnabled = !!e.target.checked;
    await storageSet(CONFIG.attachRemStorageKey, attachReminderEnabled);
  });
  loadAttachReminderSetting();

  function composeSurfaceFor(sendBtn) {
    return (
      sendBtn.closest('[role="dialog"]') ||
      sendBtn.closest("form") ||
      sendBtn.closest('[aria-label*="compose" i]') ||
      document
    );
  }
  const attachBypass = new WeakSet();
  document.addEventListener(
    "click",
    (e) => {
      if (!attachReminderEnabled) return;
      const btn = e.target.closest ? e.target.closest(CONFIG.sendButtonSelectors.join(",")) : null;
      if (!btn) return;
      if (attachBypass.has(btn)) {
        attachBypass.delete(btn);
        return; // this click is our own re-dispatch after the person confirmed
      }
      const surface = composeSurfaceFor(btn);
      const bodyField = qFirst(CONFIG.composeBodySelectors, surface) || qFirst(CONFIG.composeBodySelectors);
      const text = bodyField ? bodyField.innerText || "" : "";
      if (!ATTACH_MENTION_RE.test(text)) return;
      const hasAttachment = qAllVisible(CONFIG.composeAttachmentSelectors, surface).length > 0;
      if (hasAttachment) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      const proceed = confirm(
        'Your message mentions "attach…" but InboxSentry couldn\'t find an attachment on this email. Send anyway?'
      );
      if (proceed) {
        attachBypass.add(btn);
        btn.click();
      }
    },
    true
  );

  /* =========================================================
     7) EMAIL INSIGHTS — word count, reading time, and link count
     for whatever's open in the reading pane right now.
     ========================================================= */
  function renderInsights() {
    const bodyEl = qFirst(CONFIG.readingBodySelectors);
    const wordsEl = $("#ia-insights-words");
    const readEl = $("#ia-insights-read");
    const linksEl = $("#ia-insights-links");
    const summaryEl = $("#ia-insights-summary");
    if (!wordsEl) return;
    if (!bodyEl) {
      wordsEl.textContent = "—";
      readEl.textContent = "—";
      linksEl.textContent = "—";
      summaryEl.textContent = "Open an email in the reading pane to see its stats.";
      return;
    }
    const text = bodyEl.innerText || "";
    const words = (text.match(/\S+/g) || []).length;
    const minutes = Math.max(1, Math.round(words / 200));
    const links = bodyEl.querySelectorAll("a[href]").length;
    wordsEl.textContent = words;
    readEl.textContent = minutes;
    linksEl.textContent = links;
    const info = currentEmailInfo();
    summaryEl.textContent = info
      ? `"${info.subject}" from ${info.sender}`
      : "Stats reflect the currently open email.";
  }
  shadow.querySelectorAll("#ia-tabs button").forEach((btn) => {
    if (btn.dataset.tab === "insights") btn.addEventListener("click", renderInsights);
  });
  setInterval(renderInsights, 4000);
  renderInsights();
})();
