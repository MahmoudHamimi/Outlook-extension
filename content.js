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
    // Compose windows (inline reply, popped-out new message) — editable body.
    composeBodySelectors: ['div[contenteditable="true"][role="textbox"]', 'div[contenteditable="true"]'],
    // The Send button, wherever a compose surface renders it.
    sendButtonSelectors: ['button[aria-label="Send"]', 'button[name="Send"]'],
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
    categoriesStorageKey: "ia_categories_v1",
    emailCategoryStorageKey: "ia_email_categories_v1",
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

  /* ---- Message-list row heuristics (used by the stale + priority signals) ----
     OWA's list markup isn't documented and varies by skin, so these read only
     from accessible/robust signals: aria-label text, native tooltip "title"
     attributes, and bold-text rendering (OWA's convention for unread items).
     Anything that can't be determined confidently is simply skipped. */
  function isRowUnread(row) {
    const label = (row.getAttribute("aria-label") || "").trim();
    if (/^unread\b/i.test(label)) return true;
    if (row.querySelector('[class*="unread" i]')) return true;
    const candidates = row.querySelectorAll("span, div");
    for (let i = 0; i < candidates.length && i < 12; i++) {
      const el = candidates[i];
      if (!el.textContent || !el.textContent.trim()) continue;
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

  // Injected once into the real page (not the shadow root) so badges we add to
  // Outlook's own list rows render correctly; scoped with an "ia-" prefix to
  // avoid colliding with Outlook's own classes.
  const globalStyle = document.createElement("style");
  globalStyle.textContent = `
    .ia-row-badges { display:inline-flex; gap:4px; margin-left:8px; vertical-align:middle; }
    .ia-badge { display:inline-block; font-family:"Segoe UI",Arial,sans-serif; font-size:10px;
      font-weight:700; padding:1px 7px; border-radius:8px; line-height:16px; white-space:nowrap; }
    .ia-badge-stale { background:#FBE1DE; color:#B3261E; }
    .ia-badge-priority { background:#FFF4CE; color:#8A6D00; }
    .ia-row-priority { box-shadow: inset 3px 0 0 #B3261E !important; }
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
        width: 52px; height: 52px; border-radius: 50%; background:#1B4F8C; color:#fff;
        border:none; cursor:pointer; box-shadow:0 2px 10px rgba(0,0,0,.3); font-size:20px;
      }
      .panel {
        display:none; position:fixed; bottom:82px; right:20px; width:380px; max-height:75vh;
        background:#fff; border-radius:10px; box-shadow:0 6px 24px rgba(0,0,0,.25);
        overflow:hidden; flex-direction:column; border:1px solid #E1E5EA;
      }
      .panel.open { display:flex; }
      header { background:#1B4F8C; color:#fff; padding:10px 14px; }
      header h1 { margin:0; font-size:14px; }
      nav { display:flex; flex-wrap:wrap; background:#fff; border-bottom:1px solid #E1E5EA; }
      nav button {
        flex:1 1 25%; border:none; background:none; padding:7px 3px; font-size:10px; cursor:pointer;
        color:#5B6472; border-bottom:3px solid transparent; white-space:nowrap;
      }
      nav button.active { color:#1B4F8C; border-bottom-color:#1B4F8C; font-weight:600; }
      main { padding:12px 14px; overflow-y:auto; font-size:12.5px; color:#20242B; }
      section { display:none; }
      section.active { display:block; }
      h2 { font-size:12.5px; margin:0 0 6px; }
      .muted { color:#5B6472; font-size:11px; line-height:1.5; }
      label { display:block; font-size:11px; font-weight:600; margin:8px 0 3px; }
      input, select {
        width:100%; padding:5px 7px; border:1px solid #E1E5EA; border-radius:5px; font-size:12px;
      }
      button.primary {
        background:#1B4F8C; color:#fff; border:none; border-radius:5px; padding:6px 11px;
        font-size:11.5px; cursor:pointer; margin-top:7px;
      }
      button.secondary {
        background:#fff; color:#1B4F8C; border:1px solid #1B4F8C; border-radius:5px; padding:4px 8px;
        font-size:10.5px; cursor:pointer; margin-right:6px; margin-top:4px;
      }
      .item { border:1px solid #E1E5EA; border-radius:6px; padding:7px; margin-bottom:7px; }
      .pill { display:inline-block; padding:2px 7px; border-radius:9px; font-size:10px; font-weight:600; }
      .pill.warn { background:#FCE9DA; color:#B26A00; }
      .pill.ok { background:#E4F3E5; color:#2E7D32; }
      .pill.due { background:#FBE1DE; color:#B3261E; }
      .empty { color:#5B6472; font-style:italic; font-size:11px; }
      #costResult { font-size:18px; font-weight:700; color:#1B4F8C; margin-top:4px; }
      .warning-list { margin:0; padding-left:16px; }
      .warning-list li { margin-bottom:5px; }
      .row-flex { display:flex; gap:6px; align-items:center; }
      .row-flex input, .row-flex select { flex:1; }
      .tag-list { list-style:none; margin:8px 0 0; padding:0; }
      .tag-list li {
        display:flex; justify-content:space-between; align-items:center;
        border:1px solid #E1E5EA; border-radius:6px; padding:5px 8px; margin-bottom:5px; font-size:11.5px;
      }
      .tag-list button { background:none; border:none; color:#B3261E; cursor:pointer; font-size:11px; }
      details.folder-group { border:1px solid #E1E5EA; border-radius:6px; margin-bottom:7px; }
      details.folder-group summary {
        cursor:pointer; padding:7px 9px; font-size:11.5px; font-weight:600; list-style:none;
        display:flex; justify-content:space-between; align-items:center;
      }
      details.folder-group summary::-webkit-details-marker { display:none; }
      details.folder-group .folder-items { padding:0 9px 9px; }
      .count-chip { background:#EDF1F7; color:#1B4F8C; border-radius:9px; padding:1px 8px; font-size:10px; font-weight:700; }
    </style>
    <button class="toggle" id="ia-toggle" title="Inbox Assistant">IA</button>
    <div class="panel" id="ia-panel">
      <header><h1>Inbox Assistant</h1></header>
      <nav id="ia-tabs">
        <button data-tab="followup" class="active">Follow-ups</button>
        <button data-tab="stale">Stale</button>
        <button data-tab="priority">Priority</button>
        <button data-tab="folders">Folders</button>
        <button data-tab="meeting">Meetings</button>
        <button data-tab="tone">Tone</button>
        <button data-tab="cost">Cost</button>
        <button data-tab="pdf">PDF</button>
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
          <p class="muted">Automatically scans the message list and flags any unread email that's sat unopened past the threshold — look for the <span class="pill due">⏰ Nd</span> badge directly on the row.</p>
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

        <section id="tab-folders">
          <h2>Categorize this email</h2>
          <p class="muted">Client-side folders stored by the extension — assign the open email to one, then browse by folder below.</p>
          <label>New folder name</label>
          <div class="row-flex">
            <input type="text" id="ia-folder-new" placeholder="e.g. Clients" />
            <button class="primary" id="ia-folder-create" style="margin-top:0;">Create</button>
          </div>
          <label style="margin-top:10px;">Assign open email to</label>
          <div class="row-flex">
            <select id="ia-folder-select"><option value="">No folders yet</option></select>
            <button class="primary" id="ia-folder-assign" style="margin-top:0;">Assign</button>
          </div>
          <p class="muted" id="ia-folder-status"></p>
          <h2 style="margin-top:14px;">Your folders</h2>
          <div id="ia-folder-groups"><p class="empty">No folders created yet.</p></div>
        </section>

        <section id="tab-meeting">
          <h2>Meeting-request extractor</h2>
          <p class="muted">Scans the open email for dates/times and opens a prefilled Outlook calendar event.</p>
          <button class="primary" id="ia-scan">Scan this email</button>
          <div id="ia-meeting-results"></div>
        </section>

        <section id="tab-tone">
          <h2>Reply-tone checker</h2>
          <p class="muted">Checks the open compose box for curt, overly formal, or passive-aggressive phrasing. Also intercepts Send automatically.</p>
          <button class="primary" id="ia-tone-check">Check tone now</button>
          <div id="ia-tone-results"></div>
        </section>

        <section id="tab-cost">
          <h2>Meeting cost calculator</h2>
          <label>Attendee count</label>
          <input type="number" id="ia-cost-attendees" min="1" value="4" />
          <label>Average salary band</label>
          <select id="ia-cost-salary">
            <option value="50000">~$50,000/yr</option>
            <option value="75000" selected>~$75,000/yr</option>
            <option value="100000">~$100,000/yr</option>
            <option value="125000">~$125,000/yr</option>
            <option value="150000">~$150,000/yr</option>
            <option value="200000">~$200,000/yr</option>
          </select>
          <label>Duration (minutes)</label>
          <input type="number" id="ia-cost-duration" min="5" step="5" value="30" />
          <button class="primary" id="ia-cost-calc">Calculate</button>
          <div id="ia-cost-result"></div>
          <p class="muted" id="ia-cost-note"></p>
        </section>

        <section id="tab-pdf">
          <h2>Download as PDF</h2>
          <p class="muted">Opens a print-formatted version of the open email — choose "Save as PDF" as the destination in the print dialog.</p>
          <button class="primary" id="ia-pdf">Download email as PDF</button>
          <p class="muted" id="ia-pdf-status"></p>
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
  }
  renderFollowUps();

  /* =========================================================
     1b) STALE-EMAIL SIGNALS — automatic, no click required.
     Scans the visible message list on a debounce (list re-renders
     constantly as OWA virtualizes rows) and stamps a badge onto any
     unread row whose received date is older than the threshold.
     ========================================================= */
  let staleThresholdDays = CONFIG.staleThresholdDays;
  let priorityCache = [];
  let lastStaleItems = []; // [{subject, sender, ageDays}] from the most recent scan, for the Stale tab

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
    if (!countEl || !listEl) return;
    countEl.textContent = lastStaleItems.length;
    if (lastStaleItems.length === 0) {
      listEl.innerHTML = '<p class="empty">Nothing flagged in the visible list right now.</p>';
      return;
    }
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

  let scanTimer = null;
  function scheduleScan() {
    clearTimeout(scanTimer);
    scanTimer = setTimeout(scanMailList, 500);
  }
  new MutationObserver(scheduleScan).observe(document.body, { childList: true, subtree: true });
  loadStaleThreshold().then(() => refreshPriorityCache().then(scanMailList));
  // OWA re-renders on its own, but a row's age can cross the threshold with no
  // DOM change at all (time just passes), so also re-check periodically.
  setInterval(scanMailList, 5 * 60 * 1000);

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
      li.querySelector('[data-action="remove"]').addEventListener("click", async () => {
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
     1d) FOLDERS / MANUAL CATEGORIZATION
     Client-side only: this organizes emails within the extension's
     own storage (it doesn't move messages via Outlook's own folders,
     since that would require server-side API permissions this
     extension doesn't request).
     ========================================================= */
  async function getFolders() {
    return (await storageGet(CONFIG.categoriesStorageKey)) || [];
  }
  async function saveFolders(list) {
    await storageSet(CONFIG.categoriesStorageKey, list);
  }
  async function getEmailFolders() {
    return (await storageGet(CONFIG.emailCategoryStorageKey)) || {};
  }
  async function saveEmailFolders(map) {
    await storageSet(CONFIG.emailCategoryStorageKey, map);
  }

  async function populateFolderSelect() {
    const folders = await getFolders();
    const sel = $("#ia-folder-select");
    sel.innerHTML = folders.length
      ? folders.map((f) => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`).join("")
      : '<option value="">No folders yet</option>';
  }

  async function renderFolderGroups() {
    const folders = await getFolders();
    const assigned = await getEmailFolders();
    const container = $("#ia-folder-groups");
    if (folders.length === 0) {
      container.innerHTML = '<p class="empty">No folders created yet.</p>';
      return;
    }
    const byFolder = {};
    folders.forEach((f) => (byFolder[f] = []));
    Object.entries(assigned).forEach(([id, entry]) => {
      if (byFolder[entry.folder]) byFolder[entry.folder].push({ id, ...entry });
    });

    container.innerHTML = folders
      .map((f) => {
        const items = byFolder[f] || [];
        const itemsHtml =
          items.length === 0
            ? '<p class="empty">No emails in this folder yet.</p>'
            : items
                .map(
                  (it) => `
              <div class="item" data-id="${escapeHtml(it.id)}">
                <strong style="font-size:11.5px;">${escapeHtml(it.subject).slice(0, 55)}</strong>
                <div class="muted">${escapeHtml(it.sender)}</div>
                <button class="secondary" data-action="open">Open</button>
                <button class="secondary" data-action="remove">Remove</button>
              </div>`
                )
                .join("");
        return `
        <details class="folder-group">
          <summary>${escapeHtml(f)} <span class="count-chip">${items.length}</span></summary>
          <div class="folder-items">
            ${itemsHtml}
            <button class="secondary" data-delete-folder="${escapeHtml(f)}" style="color:#B3261E; border-color:#B3261E;">Delete folder</button>
          </div>
        </details>`;
      })
      .join("");

    container.querySelectorAll(".item").forEach((el) => {
      const id = el.dataset.id;
      const openBtn = el.querySelector('[data-action="open"]');
      if (openBtn) {
        openBtn.addEventListener("click", async () => {
          const map = await getEmailFolders();
          if (map[id]) window.location.href = map[id].url;
        });
      }
      const removeBtn = el.querySelector('[data-action="remove"]');
      if (removeBtn) {
        removeBtn.addEventListener("click", async () => {
          const map = await getEmailFolders();
          delete map[id];
          await saveEmailFolders(map);
          renderFolderGroups();
        });
      }
    });
    container.querySelectorAll("[data-delete-folder]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const name = btn.dataset.deleteFolder;
        const folders2 = (await getFolders()).filter((f) => f !== name);
        await saveFolders(folders2);
        const map = await getEmailFolders();
        Object.keys(map).forEach((id) => {
          if (map[id].folder === name) delete map[id];
        });
        await saveEmailFolders(map);
        populateFolderSelect();
        renderFolderGroups();
      });
    });
  }

  $("#ia-folder-create").addEventListener("click", async () => {
    const input = $("#ia-folder-new");
    const name = input.value.trim();
    if (!name) return;
    const folders = await getFolders();
    if (!folders.some((f) => f.toLowerCase() === name.toLowerCase())) {
      folders.push(name);
      await saveFolders(folders);
    }
    input.value = "";
    populateFolderSelect();
    renderFolderGroups();
  });

  $("#ia-folder-assign").addEventListener("click", async () => {
    const statusEl = $("#ia-folder-status");
    const folder = $("#ia-folder-select").value;
    if (!folder) {
      statusEl.textContent = "Create a folder first.";
      return;
    }
    const info = currentEmailInfo();
    if (!info) {
      statusEl.textContent = "Open a specific email in the reading pane first.";
      return;
    }
    const map = await getEmailFolders();
    map[info.id] = { folder, subject: info.subject, sender: info.sender, url: info.url, assignedAt: new Date().toISOString() };
    await saveEmailFolders(map);
    statusEl.textContent = `Assigned to "${folder}".`;
    renderFolderGroups();
  });

  populateFolderSelect();
  renderFolderGroups();

  /* =========================================================
     2) MEETING-REQUEST EXTRACTOR
     Lightweight regex-based date/time extraction (no external
     libraries, since MV3 extensions can't load remote scripts).

     Strategy: find all "date anchors" (a day, however phrased) and all
     "time anchors" (a clock time, however phrased) separately, each with
     their character position in the text, then pair up anchors that sit
     close together. This catches far more real-world phrasing than a
     single combined regex — e.g. "let's sync Thursday morning", "how
     about 2-3pm on the 12th?", "next Tue at noon", "in 3 days at 10:30",
     "Fri 9/12 2:00 PM ET", or a bare "3pm works for me" with no date.
     ========================================================= */
  const MONTH_NAMES = ["january","february","march","april","may","june","july","august","september","october","november","december"];
  const MONTH_ABBR = { jan:"january", feb:"february", mar:"march", apr:"april", jun:"june", jul:"july", aug:"august", sep:"september", sept:"september", oct:"october", nov:"november", dec:"december" };
  const MONTHS_RE_PART = "january|february|march|april|may|june|july|august|september|october|november|december|jan\\.?|feb\\.?|mar\\.?|apr\\.?|jun\\.?|jul\\.?|aug\\.?|sept?\\.?|oct\\.?|nov\\.?|dec\\.?";
  const WEEKDAYS = ["sunday","monday","tuesday","wednesday","thursday","friday","saturday"];
  const WEEKDAYS_RE_PART = "sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tue|tues|wed|thu|thur|thurs|fri|sat";
  const MEETING_KEYWORDS = /\b(meet|meeting|call|sync|chat|catch up|discuss|standup|stand-up|invite|schedule|available|works for me|touch base|zoom|teams call)\b/i;

  function monthIndexFromName(raw) {
    const name = raw.toLowerCase().replace(/\.$/, "");
    const full = MONTH_ABBR[name] || name;
    return MONTH_NAMES.indexOf(full);
  }
  function weekdayIndexFromName(raw) {
    const name = raw.toLowerCase();
    const map = { sun:0, mon:1, tue:2, tues:2, wed:3, thu:4, thur:4, thurs:4, fri:5, sat:6 };
    if (map[name] !== undefined) return map[name];
    return WEEKDAYS.indexOf(name);
  }
  function to12to24(hourStr, meridiem) {
    let h = parseInt(hourStr, 10);
    if (meridiem) {
      const mer = meridiem.toLowerCase();
      if (mer === "pm" && h < 12) h += 12;
      if (mer === "am" && h === 12) h = 0;
    }
    return h;
  }

  /* ---- Date anchors ---- */
  // Patterns are tried most-specific-first; a later, less-specific match whose
  // characters overlap an already-accepted anchor is dropped (e.g. the numeric
  // pattern shouldn't also match the "08-12" tail of an ISO date, and a bare
  // weekday name shouldn't also fire inside "Monday, June 10").
  function findDateAnchors(text) {
    const anchors = [];
    const consumed = [];
    const today = new Date();

    function overlaps(index, length) {
      const s = index, e = index + length;
      return consumed.some(([cs, ce]) => s < ce && e > cs);
    }
    function push(index, length, label, y, mo, d) {
      if (overlaps(index, length)) return;
      consumed.push([index, index + length]);
      anchors.push({ index, length, label: label.trim(), y, mo, d });
    }

    let m;

    // "June 10", "Jun. 10th, 2026"
    const monthDayRe = new RegExp(`\\b(${MONTHS_RE_PART})\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,?\\s*(\\d{4}))?\\b`, "gi");
    while ((m = monthDayRe.exec(text)) !== null) {
      const mi = monthIndexFromName(m[1]);
      if (mi < 0) continue;
      push(m.index, m[0].length, m[0], m[3] ? parseInt(m[3], 10) : today.getFullYear(), mi, parseInt(m[2], 10));
    }

    // "10 June", "10th of June, 2026"
    const dayMonthRe = new RegExp(`\\b(\\d{1,2})(?:st|nd|rd|th)?\\s+(?:of\\s+)?(${MONTHS_RE_PART})(?:,?\\s*(\\d{4}))?\\b`, "gi");
    while ((m = dayMonthRe.exec(text)) !== null) {
      const mi = monthIndexFromName(m[2]);
      if (mi < 0) continue;
      push(m.index, m[0].length, m[0], m[3] ? parseInt(m[3], 10) : today.getFullYear(), mi, parseInt(m[1], 10));
    }

    // ISO "2026-06-10"
    const isoRe = /\b(\d{4})-(\d{1,2})-(\d{1,2})\b/g;
    while ((m = isoRe.exec(text)) !== null) {
      push(m.index, m[0].length, m[0], +m[1], +m[2] - 1, +m[3]);
    }

    // "the 12th", "on the 3rd" — day-of-month only, no explicit month named.
    const ordinalOnlyRe = /\bthe\s+(\d{1,2})(st|nd|rd|th)\b/gi;
    while ((m = ordinalOnlyRe.exec(text)) !== null) {
      const day = parseInt(m[1], 10);
      if (day < 1 || day > 31) continue;
      const t = new Date(today.getFullYear(), today.getMonth(), day);
      if (t < today) t.setMonth(t.getMonth() + 1); // roll to next month if this month's date already passed
      push(m.index, m[0].length, m[0], t.getFullYear(), t.getMonth(), t.getDate());
    }

    // Numeric "6/10", "06-10-2026" (assumes month/day, the common US convention)
    const numericRe = /\b(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?\b/g;
    while ((m = numericRe.exec(text)) !== null) {
      if (/^\d{1,2}[\/\-]7$/.test(m[0])) continue; // skip "24/7" etc.
      const mo = parseInt(m[1], 10) - 1;
      const d = parseInt(m[2], 10);
      if (mo < 0 || mo > 11 || d < 1 || d > 31) continue;
      let y = m[3] ? parseInt(m[3], 10) : today.getFullYear();
      if (y < 100) y += 2000;
      push(m.index, m[0].length, m[0], y, mo, d);
    }

    // "next Friday", "this Tuesday", "coming Mon", or a bare weekday like "Friday works great"
    const weekdayRe = new RegExp(`\\b(next|this|coming)?\\s*(${WEEKDAYS_RE_PART})\\b`, "gi");
    while ((m = weekdayRe.exec(text)) !== null) {
      const wd = weekdayIndexFromName(m[2]);
      if (wd < 0) continue;
      const qualifier = (m[1] || "").toLowerCase();
      const base = new Date(today);
      let diff = ((wd - base.getDay()) + 7) % 7; // 0 = that weekday is today
      if (qualifier === "next") diff = diff === 0 ? 7 : diff + 7; // "next Friday" = the Friday of next week
      base.setDate(base.getDate() + diff);
      push(m.index, m[0].length, m[0], base.getFullYear(), base.getMonth(), base.getDate());
    }

    // "today", "tomorrow"
    const todayRe = /\btoday\b/gi;
    while ((m = todayRe.exec(text)) !== null) {
      push(m.index, m[0].length, m[0], today.getFullYear(), today.getMonth(), today.getDate());
    }
    const tomorrowRe = /\btomorrow\b/gi;
    while ((m = tomorrowRe.exec(text)) !== null) {
      const t = new Date(today);
      t.setDate(t.getDate() + 1);
      push(m.index, m[0].length, m[0], t.getFullYear(), t.getMonth(), t.getDate());
    }

    // "in 3 days", "in 2 weeks"
    const inNRe = /\bin\s+(\d+)\s+(day|days|week|weeks)\b/gi;
    while ((m = inNRe.exec(text)) !== null) {
      const n = parseInt(m[1], 10);
      const unit = m[2].startsWith("week") ? 7 : 1;
      const t = new Date(today);
      t.setDate(t.getDate() + n * unit);
      push(m.index, m[0].length, m[0], t.getFullYear(), t.getMonth(), t.getDate());
    }

    // Next-week shorthand ("next week" with no specific day) — default to next week's Monday.
    const nextWeekRe = /\bnext week\b/gi;
    while ((m = nextWeekRe.exec(text)) !== null) {
      const base = new Date(today);
      const daysSinceMonday = (base.getDay() + 6) % 7; // Mon=0..Sun=6
      const mondayThisWeek = new Date(base);
      mondayThisWeek.setDate(base.getDate() - daysSinceMonday);
      const mondayNextWeek = new Date(mondayThisWeek);
      mondayNextWeek.setDate(mondayThisWeek.getDate() + 7);
      push(m.index, m[0].length, m[0], mondayNextWeek.getFullYear(), mondayNextWeek.getMonth(), mondayNextWeek.getDate());
    }

    return anchors;
  }

  /* ---- Time anchors (single times and ranges) ---- */
  function findTimeAnchors(text) {
    const anchors = []; // {index, length, label, h1, m1, h2?, m2?}
    let m;

    // Ranges first so the single-time pass below doesn't also match half of them:
    // "2-3pm", "2:00-3:30pm", "2pm to 3pm", "between 2 and 3pm"
    const rangeRe = /\b(?:between\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\s*(?:-|–|to|and)\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/gi;
    const consumed = [];
    while ((m = rangeRe.exec(text)) !== null) {
      const endMer = m[6];
      const startMer = m[3] || endMer; // "2-3pm" implies 2 shares 3's meridiem unless stated
      const h1 = to12to24(m[1], startMer);
      const h2 = to12to24(m[4], endMer);
      anchors.push({
        index: m.index, length: m[0].length, label: m[0].trim(),
        h1, m1: m[2] ? parseInt(m[2], 10) : 0,
        h2, m2: m[5] ? parseInt(m[5], 10) : 0
      });
      consumed.push([m.index, m.index + m[0].length]);
    }

    const isConsumed = (idx) => consumed.some(([s, e]) => idx >= s && idx < e);

    // Single "3pm", "3:30 PM"
    const timeRe = /\b(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/gi;
    while ((m = timeRe.exec(text)) !== null) {
      if (isConsumed(m.index)) continue;
      anchors.push({ index: m.index, length: m[0].length, label: m[0].trim(), h1: to12to24(m[1], m[3]), m1: m[2] ? parseInt(m[2], 10) : 0 });
    }

    // 24-hour "14:00", "09:30"
    const time24Re = /\b([01]?\d|2[0-3]):([0-5]\d)\b/g;
    while ((m = time24Re.exec(text)) !== null) {
      if (isConsumed(m.index)) continue;
      anchors.push({ index: m.index, length: m[0].length, label: m[0], h1: parseInt(m[1], 10), m1: parseInt(m[2], 10) });
    }

    // "noon", "midnight"
    const noonRe = /\b(noon|midnight)\b/gi;
    while ((m = noonRe.exec(text)) !== null) {
      const isNoon = m[1].toLowerCase() === "noon";
      anchors.push({ index: m.index, length: m[0].length, label: m[0], h1: isNoon ? 12 : 0, m1: 0 });
    }

    return anchors;
  }

  function extractCandidates(text) {
    const found = [];
    const seen = new Set();
    function addCandidate(label, start, end) {
      if (!start || isNaN(start.getTime())) return;
      const key = start.toISOString() + "|" + (end ? end.toISOString() : "");
      if (seen.has(key)) return;
      seen.add(key);
      found.push({ label: label.trim(), date: start, end: end || null });
    }

    const dateAnchors = findDateAnchors(text);
    const timeAnchors = findTimeAnchors(text);
    const usedTimeIdx = new Set();
    const NEAR = 45; // chars

    // Pair each date anchor with the closest nearby time anchor.
    dateAnchors.forEach((d) => {
      let best = null, bestDist = Infinity;
      timeAnchors.forEach((t, i) => {
        if (usedTimeIdx.has(i)) return;
        const dist = Math.min(
          Math.abs(t.index - (d.index + d.length)),
          Math.abs(d.index - (t.index + t.length))
        );
        if (dist <= NEAR && dist < bestDist) { best = { t, i }; bestDist = dist; }
      });
      const start = new Date(d.y, d.mo, d.d, 9, 0, 0, 0); // default 9:00 AM if no time nearby
      let end = null;
      let label = d.label;
      if (best) {
        usedTimeIdx.add(best.i);
        start.setHours(best.t.h1, best.t.m1, 0, 0);
        if (best.t.h2 !== undefined) {
          end = new Date(d.y, d.mo, d.d, best.t.h2, best.t.m2, 0, 0);
        }
        label = d.index < best.t.index ? `${d.label} ${best.t.label}` : `${best.t.label} ${d.label}`;
      }
      addCandidate(label, start, end);
    });

    // Time anchors not claimed by any date, but sitting near a meeting-ish keyword,
    // are assumed to mean "today" (or "tomorrow" if that time already passed today).
    timeAnchors.forEach((t, i) => {
      if (usedTimeIdx.has(i)) return;
      const windowText = text.slice(Math.max(0, t.index - 50), t.index + t.length + 20);
      if (!MEETING_KEYWORDS.test(windowText)) return;
      const now = new Date();
      const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), t.h1, t.m1, 0, 0);
      let end = null;
      let dayLabel = "today";
      if (start.getTime() < now.getTime()) {
        start.setDate(start.getDate() + 1);
        dayLabel = "tomorrow";
      }
      if (t.h2 !== undefined) {
        end = new Date(start.getFullYear(), start.getMonth(), start.getDate(), t.h2, t.m2, 0, 0);
      }
      addCandidate(`${t.label} (${dayLabel})`, start, end);
    });

    // Order by when the phrase appears in the email, most relevant first.
    return found.slice(0, 8);
  }

  $("#ia-scan").addEventListener("click", () => {
    const resultsEl = $("#ia-meeting-results");
    const bodyEl = qFirst(CONFIG.readingBodySelectors);
    if (!bodyEl) {
      resultsEl.innerHTML = '<p class="empty">Open an email in the reading pane, then scan.</p>';
      return;
    }
    const subjectEl = qFirst(CONFIG.readingSubjectSelectors);
    const subject = subjectEl ? subjectEl.textContent.trim() : "Meeting";
    const candidates = extractCandidates(bodyEl.innerText || "");

    if (candidates.length === 0) {
      resultsEl.innerHTML = '<p class="empty">No dates or times found in this email.</p>';
      return;
    }

    resultsEl.innerHTML = candidates
      .map((c, i) => {
        const rangeNote = c.end ? ` – ${c.end.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}` : "";
        return `
      <div class="item" data-idx="${i}">
        <div><strong>"${escapeHtml(c.label)}"</strong></div>
        <div class="muted">${c.date.toLocaleString()}${rangeNote}</div>
        <button class="secondary" data-action="add">Add to Calendar</button>
      </div>`;
      })
      .join("");

    resultsEl.querySelectorAll(".item").forEach((el) => {
      const idx = parseInt(el.dataset.idx, 10);
      el.querySelector('[data-action="add"]').addEventListener("click", () => {
        const start = candidates[idx].date;
        const end = candidates[idx].end || new Date(start.getTime() + 60 * 60 * 1000);
        const toISO = (d) =>
          `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}T${String(
            d.getHours()
          ).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:00`;
        const base = window.location.host.includes("live.com") ? "outlook.live.com" : "outlook.office.com";
        const url =
          `https://${base}/calendar/0/deeplink/compose?path=/calendar/action/compose&rru=addevent` +
          `&startdt=${encodeURIComponent(toISO(start))}&enddt=${encodeURIComponent(toISO(end))}` +
          `&subject=${encodeURIComponent(subject)}`;
        window.open(url, "_blank");
      });
    });
  });

  /* =========================================================
     3) REPLY-TONE CHECKER (manual + Send interception)
     ========================================================= */
  const PASSIVE_AGGRESSIVE_PHRASES = [
    "per my last email", "as previously stated", "as i mentioned before",
    "just to reiterate", "kindly note", "not sure if you saw my last email",
    "per our conversation", "as i said", "again,", "obviously"
  ];

  function analyzeTone(text) {
    const warnings = [];
    const trimmed = (text || "").trim();
    const lower = trimmed.toLowerCase();

    if (trimmed.length > 0 && trimmed.length < 25) {
      warnings.push("This message is very short — it may come across as curt.");
    }
    PASSIVE_AGGRESSIVE_PHRASES.forEach((phrase) => {
      if (lower.includes(phrase)) warnings.push(`Possibly passive-aggressive phrase: "${phrase}"`);
    });
    if ((trimmed.match(/\b[A-Z]{4,}\b/g) || []).length > 0) {
      warnings.push("Contains ALL-CAPS words, which can read as shouting.");
    }
    if ((trimmed.match(/!/g) || []).length >= 3) {
      warnings.push("Multiple exclamation marks may read as overly intense.");
    }
    if (["kindly", "esteemed", "herewith"].some((w) => lower.includes(w))) {
      warnings.push("Phrasing reads as overly formal/stiff for most email contexts.");
    }
    return warnings;
  }

  function findActiveComposeBody() {
    const boxes = qAllVisible(CONFIG.composeBodySelectors);
    if (boxes.length === 0) return null;
    // Prefer the currently focused box; otherwise the last (most recently opened) one.
    return boxes.find((b) => b.contains(document.activeElement)) || boxes[boxes.length - 1];
  }

  $("#ia-tone-check").addEventListener("click", () => {
    const resultsEl = $("#ia-tone-results");
    const box = findActiveComposeBody();
    if (!box) {
      resultsEl.innerHTML = '<p class="empty">Open a compose or reply window first.</p>';
      return;
    }
    const warnings = analyzeTone(box.innerText);
    resultsEl.innerHTML =
      warnings.length === 0
        ? '<p class="pill ok">Looks good — no tone issues detected.</p>'
        : '<ul class="warning-list">' +
          warnings.map((w) => `<li><span class="pill warn">Check</span> ${escapeHtml(w)}</li>`).join("") +
          "</ul>";
  });

  // Best-effort Send interception: OWA doesn't expose a "before send" event to
  // page scripts, so this hooks the Send button's click in the capture phase.
  let bypassOnce = false;
  document.addEventListener(
    "click",
    (e) => {
      if (bypassOnce) {
        bypassOnce = false;
        return;
      }
      const target = e.target.closest(CONFIG.sendButtonSelectors.join(","));
      if (!target) return;

      const box = findActiveComposeBody();
      if (!box) return;
      const warnings = analyzeTone(box.innerText);
      if (warnings.length === 0) return;

      e.preventDefault();
      e.stopImmediatePropagation();
      showToneWarningOverlay(warnings, target);
    },
    true
  );

  function showToneWarningOverlay(warnings, sendButton) {
    const overlay = document.createElement("div");
    overlay.style.cssText =
      "position:fixed; inset:0; background:rgba(0,0,0,.4); z-index:2147483001; display:flex; align-items:center; justify-content:center;";
    const box = document.createElement("div");
    box.style.cssText =
      "all:initial; background:#fff; font-family:'Segoe UI',Arial,sans-serif; border-radius:8px; padding:18px 20px; max-width:360px; box-shadow:0 8px 30px rgba(0,0,0,.3); color:#20242B;";
    box.innerHTML = `
      <h3 style="margin:0 0 8px; font-size:14px;">Tone check before sending</h3>
      <ul style="margin:0 0 12px; padding-left:18px; font-size:12.5px;">
        ${warnings.map((w) => `<li style="margin-bottom:5px;">${escapeHtml(w)}</li>`).join("")}
      </ul>
      <div style="text-align:right;">
        <button id="ia-cancel" style="background:#fff; border:1px solid #ccc; border-radius:5px; padding:6px 10px; margin-right:8px; cursor:pointer; font-size:12px;">Go back &amp; edit</button>
        <button id="ia-sendanyway" style="background:#1B4F8C; color:#fff; border:none; border-radius:5px; padding:6px 12px; cursor:pointer; font-size:12px;">Send anyway</button>
      </div>
    `;
    overlay.appendChild(box);
    document.body.appendChild(overlay);
    overlay.querySelector("#ia-cancel").addEventListener("click", () => overlay.remove());
    overlay.querySelector("#ia-sendanyway").addEventListener("click", () => {
      overlay.remove();
      bypassOnce = true;
      sendButton.click();
    });
  }

  /* =========================================================
     4) MEETING COST CALCULATOR
     ========================================================= */
  $("#ia-cost-calc").addEventListener("click", () => {
    const attendees = Math.max(1, parseInt($("#ia-cost-attendees").value, 10) || 1);
    const salary = parseInt($("#ia-cost-salary").value, 10);
    const minutes = Math.max(1, parseInt($("#ia-cost-duration").value, 10) || 30);
    const hourlyRate = (salary / 2080) * 1.3; // +30% for benefits/overhead
    const hours = minutes / 60;
    const cost = attendees * hourlyRate * hours;

    $("#ia-cost-result").textContent = `$${cost.toFixed(0)}`;
    $("#ia-cost-note").textContent =
      `${attendees} attendee(s) x ~$${hourlyRate.toFixed(0)}/hr (incl. ~30% overhead) x ${hours.toFixed(2)} hr(s). ` +
      `Weekly recurring would run ~$${(cost * 52).toLocaleString(undefined, { maximumFractionDigits: 0 })}/year.`;
  });

  /* =========================================================
     5) DOWNLOAD EMAIL AS PDF (via native print-to-PDF)
     ========================================================= */
  $("#ia-pdf").addEventListener("click", () => {
    const statusEl = $("#ia-pdf-status");
    const bodyEl = qFirst(CONFIG.readingBodySelectors);
    const subjectEl = qFirst(CONFIG.readingSubjectSelectors);
    if (!bodyEl) {
      statusEl.textContent = "Open an email in the reading pane first.";
      return;
    }
    const subject = subjectEl ? subjectEl.textContent.trim() : "Email";
    const sender = findSenderName() || "Unknown sender";
    const win = window.open("", "_blank");
    if (!win) {
      statusEl.textContent = "Pop-up blocked — allow pop-ups for Outlook to export.";
      return;
    }
    win.document.write(`
      <html><head><title>${escapeHtml(subject)}</title>
      <style>
        body{font-family:Arial,sans-serif; padding:32px; color:#111; max-width:700px; margin:auto;}
        h1{font-size:18px; border-bottom:1px solid #ccc; padding-bottom:10px;}
        .meta{color:#666; font-size:12px; margin-bottom:4px;}
        .meta.exported{margin-bottom:20px;}
        .body{font-size:13px; line-height:1.6; white-space:pre-wrap;}
      </style></head>
      <body>
        <h1>${escapeHtml(subject)}</h1>
        <div class="meta"><strong>From:</strong> ${escapeHtml(sender)}</div>
        <div class="meta exported">Exported ${new Date().toLocaleString()} from Outlook Web</div>
        <div class="body">${escapeHtml(bodyEl.innerText || "")}</div>
      </body></html>
    `);
    win.document.close();
    statusEl.textContent = "Opened printable version — choose \"Save as PDF\" in the print dialog.";
    win.onload = () => win.print();
  });
})();