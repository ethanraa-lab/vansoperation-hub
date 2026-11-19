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
    collections: { label: "Vans Smart Collection",    icon: "sparkle", href: "soon.html?p=collections",      desc: "Create collections with conditions set automatically" },
    images:      { label: "Vans Image Workbench",     icon: "image",   href: "image-workbench.html",           desc: "Find and arrange product images by item code." },
    sku:         { label: "Vans Shopify SKU Mapping", icon: "link",    href: "soon.html?p=sku",              desc: "Map and sync SKUs with eRetail." },
    history:     { label: "Run History",              icon: "history", href: "soon.html?p=history",          desc: "Every run of every tool, in one list." },
    settings:    { label: "Settings",                 icon: "sliders", href: "soon.html?p=settings",         desc: "Mapping presets, channel codes and naming rules in one place." }
  };
  const GROUPS = [["Workspace", ["dashboard", "batches"]], ["Tools", ["hub", "collections", "images", "sku"]], ["System", ["history", "settings"]]];
  const cur = document.body.dataset.page, P = window.PAGES;
  document.body.insertAdjacentHTML("afterbegin", ICONS);
  const side = document.getElementById("side");
  if (side) side.innerHTML =
    `<a class="brand" href="${P.dashboard.href}" aria-label="Dotcom Operations Hub, go to dashboard"><div class="mark" aria-hidden="true"></div><div><b>Dotcom<br>Operations Hub</b></div></a><nav aria-label="Main">` +
    GROUPS.map(([g, ids], i) => `<div class="grp"${i ? "" : ' style="padding-top:0"'}>${g}</div>` + ids.map(id =>
      `<a href="${P[id].href}"${id === cur ? ' aria-current="page"' : ""}><svg class="i"><use href="#${P[id].icon}"/></svg>${P[id].label}</a>`).join("")).join("") +
    `</nav><div class="me"><div class="av">EC</div><div><b style="color:#fff;font-weight:600">E-commerce team</b><small>Internal workspace</small></div></div>`;
  document.querySelectorAll("[data-link]").forEach(a => { const p = P[a.dataset.link]; if (p) a.href = p.href; });

  /* run log shared by all pages (stored in this browser) */
  window.Runs = {
    demo: [
      { t: "Nov 12, 2:14 PM", batch: "Vans Malaysia · November Drop", items: 50, status: "Completed", dur: "5 min 41 sec", m: "MY" },
      { t: "Nov 11, 11:03 AM", batch: "Vans Singapore · Core Restock", items: 36, status: "Completed", dur: "4 min 28 sec", m: "SG" },
      { t: "Nov 10, 4:27 PM", batch: "Vans Malaysia · Classics Update", items: 48, status: "Completed", dur: "5 min 03 sec", m: "MY" }],
    get() { try { const s = JSON.parse(localStorage.getItem("ops-runs") || "null"); if (Array.isArray(s) && s.length) return s; } catch (e) {} return this.demo.slice(); },
    add(r) { const a = this.get(); a.unshift(r); try { localStorage.setItem("ops-runs", JSON.stringify(a.slice(0, 50))); } catch (e) {} }
  };
})();
