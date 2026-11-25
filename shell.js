/* Shared shell: sidebar, page list, run log.
   To add a tool page: set its href below, then copy soon.html and replace the content. */
(function () {
  const ICONS = `<svg width="0" height="0" style="position:absolute" aria-hidden="true">
  <symbol id="home" viewBox="0 0 24 24"><path d="M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></symbol>
  <symbol id="layers" viewBox="0 0 24 24"><path d="M12 3l9 5-9 5-9-5z"/><path d="M3 13l9 5 9-5"/></symbol>
  <symbol id="grid" viewBox="0 0 24 24"><rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/></symbol>
  <symbol id="image" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="2.5"/><circle cx="9" cy="10" r="1.6"/><path d="M3 17l5-5 4 4 3-3 6 6"/></symbol>
  <symbol id="link" viewBox="0 0 24 24"><path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/></symbol>
  <symbol id="history" viewBox="0 0 24 24"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l3 2"/></symbol>
  <symbol id="sliders" viewBox="0 0 24 24"><path d="M4 6h8M18 6h2M4 12h2M12 12h8M4 18h10M20 18h0"/><circle cx="15" cy="6" r="2.2"/><circle cx="9" cy="12" r="2.2"/><circle cx="17" cy="18" r="2.2"/></symbol>
  <symbol id="sparkle" viewBox="0 0 24 24"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/><path d="M19 16v4M17 18h4"/></symbol>
  <symbol id="plus" viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></symbol>
  <symbol id="arrow" viewBox="0 0 24 24"><path d="M5 12h14M13 6l6 6-6 6"/></symbol>
  <symbol id="check" viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></symbol>
  <symbol id="clock" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></symbol>
  <symbol id="done" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l3 3 5-6"/></symbol>
</svg>`;
  window.PAGES = {
    dashboard:   { label: "Dashboard",                icon: "home",    href: "index.html",                   desc: "" },
    batches:     { label: "Listing Batches",          icon: "layers",  href: "soon.html?p=batches",          desc: "Track every listing batch from SAP and DAM data to Shopify import." },
    hub:         { label: "Vans Listing Hub",         icon: "grid",    href: "listinghub.html",              desc: "Merge SAP and DAM for Matrixify" },
    duplicates:  { label: "Matrixify Duplicate Checker", icon: "layers", href: "duplicates.html", desc: "Check Matrixify exports for duplicate records" },
    collections: { label: "Collection Standardizer",    icon: "sparkle", href: "collections.html",      desc: "Compare MY and SG collection rules and export SG changes" },
    images:      { label: "Vans Image Workbench",     icon: "image",   href: "vans_image.html",           desc: "Find and arrange product images by item code." },
    sku:         { label: "Vans Shopify SKU Mapping", icon: "link",    href: "soon.html?p=sku",              desc: "Map and sync SKUs with eRetail." },
    history:     { label: "Run History",              icon: "history", href: "soon.html?p=history",          desc: "Review activity recorded by the hub." },
    settings:    { label: "Settings",                 icon: "sliders", href: "soon.html?p=settings",         desc: "Mapping presets, channel codes and naming rules in one place." }
  };
  const GROUPS = [["Workspace", ["dashboard", "batches"]], ["Tools", ["hub", "collections", "images", "duplicates", "sku"]], ["System", ["history", "settings"]]];
  const cur = document.body.dataset.page, P = window.PAGES;
  document.body.insertAdjacentHTML("afterbegin", ICONS);
  const side = document.getElementById("side");
  if (side) side.innerHTML =
    `<a class="brand" href="${P.dashboard.href}" aria-label="Dotcom Operations Hub, go to dashboard" title="Dashboard"><div class="mark" aria-hidden="true"></div><div class="sidebar-label"><b>Dotcom<br>Operations Hub</b></div></a><button type="button" id="sidebar-pin" aria-controls="sidebar-nav" aria-pressed="false" aria-label="Keep sidebar expanded" title="Keep sidebar expanded">⇥</button><nav id="sidebar-nav" aria-label="Main">` +
    GROUPS.map(([g, ids], i) => `<div class="grp"${i ? "" : ' style="padding-top:0"'}>${g}</div>` + ids.map(id =>
      `<a href="${P[id].href}" aria-label="${P[id].label}" title="${P[id].label}"${id === cur ? ' aria-current="page"' : ""}><svg class="i"><use href="#${P[id].icon}"/></svg><span class="sidebar-label">${P[id].label}</span></a>`).join("")).join("") +
    `</nav><div class="me"><div class="av">EC</div><div class="sidebar-label"><b style="color:#fff;font-weight:600">E-commerce team</b><small>Internal workspace</small></div></div>`;
  if (side) {
    const key = "ops-sidebar-pinned";
    const pin = document.getElementById("sidebar-pin");

    function setPinned(pinned) {
      document.body.classList.toggle("sidebar-pinned", pinned);
      pin.setAttribute("aria-pressed", String(pinned));
      pin.title = pinned ? "Unpin sidebar" : "Keep sidebar expanded";
      pin.textContent = pinned ? "⇤" : "⇥";
    }

    function restorePin() {
      try { setPinned(localStorage.getItem(key) === "1"); } catch {}
    }

    document.body.classList.add("sidebar-rail");
    restorePin();
    pin.addEventListener("click", () => {
      const pinned = pin.getAttribute("aria-pressed") !== "true";
      setPinned(pinned);
      try { localStorage.setItem(key, pinned ? "1" : "0"); } catch {}
    });
    window.addEventListener("pageshow", restorePin);
  }

  document.querySelectorAll("[data-link]").forEach(a => { const p = P[a.dataset.link]; if (p) a.href = p.href; });

  const historyKey = "ops-activity-v1";
  const dateFormat = new Intl.DateTimeFormat("en-MY", {
    timeZone: "Asia/Kuala_Lumpur", dateStyle: "medium", timeStyle: "short"
  });
  let storageUnavailable = false;

  function duration(seconds) {
    if (seconds == null) return "—";
    if (seconds < 1) return "<1 sec";
    if (seconds < 60) return Math.round(seconds) + " sec";
    return Math.floor(seconds / 60) + " min " + Math.floor(seconds % 60) + " sec";
  }

  window.Runs = {
    get() {
      try {
        const records = JSON.parse(localStorage.getItem(historyKey) || "[]");
        return Array.isArray(records) ? records.filter(record => record && typeof record.activity === "string").map(record => ({
          ...record, t: Number.isFinite(record.time) ? dateFormat.format(record.time) : "—",
          batch: record.activity, items: record.items ?? "—", m: record.market || "", dur: duration(record.seconds)
        })) : [];
      } catch {
        storageUnavailable = true;
        return [];
      }
    },
    add(activity, { items, market = "", status = "Completed", started, file = "" } = {}) {
      const now = Date.now();
      const records = this.get();
      records.unshift({
        time: now, tool: P[cur]?.label || "Dotcom Operations Hub", activity, items, market, status, file,
        seconds: started ? Math.max(0, (now - started) / 1000) : null
      });
      try {
        localStorage.setItem(historyKey, JSON.stringify(records));
      } catch {
        storageUnavailable = true;
      }
    },
    render(tbody, { market = "all", limit = Infinity } = {}) {
      const records = this.get().filter(record => market === "all" ||
        String(record.market || "").split("/").includes(market)).slice(0, limit);
      tbody.replaceChildren();
      for (const record of records) {
        const row = tbody.insertRow();
        for (const value of [record.t, record.tool || "—", record.activity, record.file || "—",
          record.market || "—", record.items, record.status || "—", record.dur]) {
          row.insertCell().textContent = value;
        }
      }
      if (!records.length) {
        const cell = tbody.insertRow().insertCell();
        cell.colSpan = 8;
        cell.textContent = storageUnavailable ? "Activity storage is unavailable in this browser." :
          market === "all" ? "No activity yet. Open a tool to get started." : "No activity for this market yet.";
      }
    }
  };
  if (["hub", "collections", "images", "sku"].includes(cur)) Runs.add("Opened tool");
})();
