"use strict";

const SOURCE_ID = "Source: ID";
const SOURCE_TYPE = "Source: Type";
const RULE_COLUMNS = [
  "Inclusion: Type", "Inclusion: Match", "Condition: Field",
  "Condition: Relation", "Condition: Value", "Condition: Match"
];
const REQUIRED = ["ID", "Handle", "Title", SOURCE_ID, SOURCE_TYPE, ...RULE_COLUMNS];
const STATUSES = ["Same rules", "Different rules", "MY only", "SG only"];
const clean = value => String(value ?? "").trim();
const lower = value => clean(value).toLowerCase();
const numericId = value => /^[1-9]\d*$/.test(value);
const field = rule => lower(rule[2]);
const manual = rule => field(rule) === "manual selection";

function readCollections(workbook, market) {
  const sheet = workbook.Sheets.Collections;
  if (!sheet) throw new Error(market + " needs the current Matrixify Collections sheet.");
  const [headers = [], ...rows] = XLSX.utils.sheet_to_json(sheet, {
    header: 1, defval: "", raw: true, blankrows: false
  });
  const names = headers.map(clean);
  const missing = REQUIRED.filter(name => names.indexOf(name) === -1);
  if (missing.length) throw new Error(market + " is missing columns: " + missing.join(", ") + ".");
  if (REQUIRED.some(name => names.indexOf(name) !== names.lastIndexOf(name))) {
    throw new Error(market + " has duplicate required column headers.");
  }

  const byHandle = new Map();
  const byId = new Map();
  for (const cells of rows) {
    const row = Object.fromEntries(names.map((name, index) => [name, clean(cells[index])]));
    if (!row.ID && !row.Handle && !row.Title && !row[SOURCE_ID] &&
        !row[SOURCE_TYPE] && !RULE_COLUMNS.some(name => row[name])) continue;
    if (!row.ID || !row.Handle) {
      throw new Error(market + " has a populated row without an ID or Handle. Re-export Collections with all rows.");
    }

    const handle = lower(row.Handle);
    const existingHandle = byId.get(row.ID);
    if (existingHandle && existingHandle !== handle) {
      throw new Error(market + " repeats collection ID " + row.ID + " across handles.");
    }
    byId.set(row.ID, handle);
    if (!byHandle.has(handle)) {
      byHandle.set(handle, { id: row.ID, handle: row.Handle, title: row.Title, sources: new Map() });
    }
    const collection = byHandle.get(handle);
    if (collection.id !== row.ID) throw new Error(market + " repeats handle " + handle + " with different IDs.");
    if (!collection.title) collection.title = row.Title;
    if (!collection.metadata || lower(row["Top Row"]) === "true") {
      collection.metadata = Object.fromEntries(Object.entries(row).filter(([name]) =>
        name.startsWith("Metafield:") || ["Title", "Sort Order", "Template Suffix",
          "Products Count", "Published: Online Store"].includes(name)));
    }

    const rule = RULE_COLUMNS.map(name => row[name]);
    if (!row[SOURCE_ID] && !row[SOURCE_TYPE] && !rule.some(Boolean)) continue;
    const key = row[SOURCE_ID] || JSON.stringify([row[SOURCE_TYPE], row["Source: Title"], row["Source: Description"]]);
    if (!collection.sources.has(key)) {
      collection.sources.set(key, {
        id: row[SOURCE_ID], type: row[SOURCE_TYPE], title: row["Source: Title"] || "",
        description: row["Source: Description"] || "", rules: []
      });
    }
    const source = collection.sources.get(key);
    if (row[SOURCE_TYPE] && source.type !== row[SOURCE_TYPE]) {
      throw new Error(market + " has conflicting source types for " + handle + ".");
    }
    if (rule.some(Boolean)) source.rules.push(rule);
  }
  if (market === "MY" && !byHandle.size) throw new Error("The MY export contains no collections.");
  return byHandle;
}

function signature(collection) {
  const sources = [...collection.sources.values()].map(source => {
    const rules = source.rules.map(rule => JSON.stringify(rule.map((value, index) =>
      index === 4 ? value : lower(value))));
    return JSON.stringify([lower(source.type), [...new Set(rules)].sort()]);
  });
  return JSON.stringify(sources.sort());
}

function reviewReason(collection) {
  const sources = [...collection.sources.values()];
  if (sources.some(source => source.rules.some(manual))) return "Contains manual selections. Verify SG product membership.";
  if (sources.length !== 1) return "Requires exactly one source with exported conditions for automatic export.";
  const [source] = sources;
  if (!["products", "variants"].includes(lower(source.type))) return "Requires review of the source type.";
  if (!source.rules.length) return "Source has no exported rules.";
  for (const rule of source.rules) {
    if (field(rule) === "collection") return "References another collection. Verify the reference in SG.";
    if (!["include", "exclude"].includes(lower(rule[0])) ||
        !["all conditions", "any condition"].includes(lower(rule[1])) ||
        !rule[2] || !rule[3] || (!rule[4] && !["is empty", "is not empty"].includes(lower(rule[3]))) ||
        (rule[5] && !["and", "or"].includes(lower(rule[5])))) {
      return "Has incomplete or unrecognized condition fields.";
    }
  }
  return "";
}

function exportBlock(reference, destination) {
  const problem = reviewReason(reference);
  if (problem) return "MY: " + problem;
  if (!destination) return reference.title ? "" : "The MY title is missing.";
  const destinationProblem = reviewReason(destination);
  if (destinationProblem) return "SG: " + destinationProblem;
  const [mySource] = reference.sources.values();
  const [sgSource] = destination.sources.values();
  if (!numericId(destination.id) || !numericId(sgSource.id)) return "SG collection and source IDs are required.";
  if (lower(mySource.type) !== lower(sgSource.type)) return "The source types differ.";
  if (signature(reference) === signature(destination)) return "The collection rules already match.";
  return "";
}

function compareCollections(my, sg) {
  return [...new Set([...my.keys(), ...sg.keys()])].sort().map(handle => {
    const reference = my.get(handle);
    const destination = sg.get(handle);
    const status = !reference ? "SG only" : !destination ? "MY only" :
      signature(reference) === signature(destination) ? "Same rules" : "Different rules";
    return { handle, reference, destination, status, decision: "keep", target: "", applyMappedRules: false };
  });
}

function metadataDifferences(reference, destination) {
  const left = reference?.metadata ?? {};
  const right = destination?.metadata ?? {};
  return [...new Set([...Object.keys(left), ...Object.keys(right)])].sort()
    .filter(key => (left[key] || "") !== (right[key] || ""))
    .map(key => [key, left[key] || "", right[key] || ""]);
}

function workbookRows(entries, action) {
  const updating = action === "update";
  const identity = updating
    ? ["ID", "Handle", "Command", SOURCE_ID, "Source: Command", SOURCE_TYPE, "Source: Title", "Source: Description"]
    : ["Handle", "Title", "Command", SOURCE_TYPE, "Source: Title", "Source: Description", "Published: Online Store"];
  const output = [[...identity, ...RULE_COLUMNS]];
  for (const entry of entries) {
    const [mySource] = entry.reference.sources.values();
    const [sgSource] = entry.destination?.sources.values() ?? [];
    const prefix = updating
      ? [entry.destination.id, entry.destination.handle, "UPDATE", sgSource.id, "REPLACE",
          sgSource.type, sgSource.title, sgSource.description]
      : [entry.reference.handle, entry.reference.title, "NEW", mySource.type, mySource.title, mySource.description, false];
    const rows = new Set(mySource.rules.map(rule => JSON.stringify(rule)));
    for (const row of rows) output.push([...prefix, ...JSON.parse(row)]);
  }
  return output;
}

function initialize() {
  const get = id => document.getElementById(id);
  const message = get("message");
  const filter = get("filter");
  const decisions = { keep: "Keep", update: "Update SG rules", create: "Create in SG", map: "Map to existing SG" };
  let entries = [];
  let destinations = new Map();
  let revision = 0;
  for (const status of STATUSES) filter.add(new Option(status, status));

  function describe(text, error = false) {
    message.textContent = text;
    message.classList.toggle("error", error);
    message.hidden = !text;
  }

  function destinationFor(entry) {
    return entry.decision === "map" ? destinations.get(entry.target) : entry.destination;
  }

  function pendingExports() {
    const updates = entries.filter(entry => entry.decision === "update" ||
      (entry.decision === "map" && entry.applyMappedRules && destinationFor(entry)));
    const creates = entries.filter(entry => entry.decision === "create");
    const counts = new Map();
    for (const entry of updates) {
      const id = destinationFor(entry).id;
      counts.set(id, (counts.get(id) || 0) + 1);
    }
    const conflict = [...counts.values()].some(count => count > 1);
    return { updates, creates, conflict };
  }

  function syncDownloads() {
    const { updates, creates, conflict } = pendingExports();
    get("updates").disabled = !updates.length || conflict;
    get("creates").disabled = !creates.length;
    get("selection").textContent = conflict ? "Two updates target the same SG collection. Select only one." :
      updates.length + " updates · " + creates.length + " creations selected";
  }

  function element(tag, content) {
    const node = document.createElement(tag);
    node.textContent = content;
    return node;
  }

  function addRules(container, market, collection) {
    const side = document.createElement("section");
    side.append(element("strong", market + " · " + (collection?.title || "No collection selected")));
    if (collection) {
      if (![...collection.sources.values()].some(source => source.rules.length)) {
        side.append(element("p", "No rules in this export."));
      }
      for (const source of collection.sources.values()) {
        if (!source.rules.length) continue;
        side.append(element("p", source.type + (source.title ? " · " + source.title : "")));
        const list = document.createElement("ul");
        list.className = "rule-list";
        for (const rule of source.rules) list.append(element("li", rule.filter(Boolean).join(" · ")));
        side.append(list);
      }
    }
    container.append(side);
  }

  function render() {
    const search = lower(get("search").value);
    const tbody = get("rows");
    tbody.replaceChildren();
    for (const entry of entries) {
      const { handle, reference, destination, status } = entry;
      if (filter.value && status !== filter.value) continue;
      if (![handle, reference?.title, destination?.title, destinationFor(entry)?.handle,
        destinationFor(entry)?.title].some(value => lower(value).includes(search))) continue;
      const row = tbody.insertRow();
      const selection = row.insertCell();
      const identity = row.insertCell();
      identity.append(element("strong", reference?.title || destination?.title || handle), element("div", handle));
      const badge = element("span", status);
      badge.className = "status";
      badge.dataset.status = status;
      row.insertCell().append(badge);

      const comparison = row.insertCell();
      const details = document.createElement("details");
      details.append(element("summary", "Compare MY and SG rules"));
      comparison.append(details);
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      selection.append(checkbox);
      const note = element("small", "");
      comparison.append(note);

      function updateRow() {
        const chosen = destinationFor(entry);
        const blocked = reference ? exportBlock(reference, chosen) : "SG-only collection.";
        checkbox.disabled = Boolean(blocked);
        checkbox.checked = ["update", "create"].includes(entry.decision) || entry.applyMappedRules;
        const action = chosen ? "update" : "create";
        checkbox.setAttribute("aria-label", "Select " + handle + " for SG " + action);
        note.textContent = blocked || (chosen ? "Tick to update the SG collection." : "Tick to create this MY handle in SG.");
        if (entry.target && blocked && !exportBlock(reference)) {
          note.textContent += " Choose Create as a separate SG collection below to use the MY handle.";
        }
        note.className = "muted export-note";
        const pair = document.createElement("div");
        pair.className = "rule-pair";
        addRules(pair, "MY", reference);
        addRules(pair, "SG", chosen);
        details.replaceChildren(details.firstChild, pair);
        syncDownloads();
      }

      checkbox.onchange = () => {
        if (entry.target) entry.applyMappedRules = checkbox.checked;
        else entry.decision = checkbox.checked ? (destination ? "update" : "create") : "keep";
        syncDownloads();
      };

      if (reference && !destination && destinations.size) {
        const mapping = document.createElement("details");
        mapping.append(element("summary", "Create separately or map to an existing SG collection"));
        const target = document.createElement("select");
        target.setAttribute("aria-label", "SG mapping target for " + handle);
        target.add(new Option("Create as a separate SG collection", ""));
        for (const [key, collection] of [...destinations].sort(([a], [b]) => a.localeCompare(b))) {
          target.add(new Option(collection.title + " (" + key + ")", key));
        }
        target.value = entry.target;
        mapping.open = Boolean(entry.target);
        target.onchange = () => {
          entry.target = target.value;
          entry.decision = entry.target ? "map" : "keep";
          entry.applyMappedRules = false;
          updateRow();
          details.open = true;
          if (!entry.target && !checkbox.disabled) checkbox.focus();
        };
        mapping.append(target);
        comparison.append(mapping);
      }
      updateRow();
    }
    if (!tbody.rows.length) {
      const cell = tbody.insertRow().insertCell();
      cell.colSpan = 4;
      cell.textContent = "No collections match this filter.";
    }
    syncDownloads();
  }

  for (const id of ["my-file", "sg-file"]) {
    get(id).onchange = () => {
      revision++;
      destinations = new Map();
      entries = [];
      get("results").hidden = true;
      describe("");
    };
  }
  get("compare-form").onsubmit = async event => {
    event.preventDefault();
    const current = ++revision;
    const myFile = get("my-file").files[0];
    const sgFile = get("sg-file").files[0];
    get("compare").disabled = true;
    get("results").hidden = true;
    const started = Date.now();
    const file = [myFile.name, sgFile.name].join(" + ");
    describe("Reading exports…");
    try {
      if (typeof XLSX === "undefined") throw new Error("The Excel library could not load. Connect to the internet and reopen the page.");
      const [myData, sgData] = await Promise.all([myFile.arrayBuffer(), sgFile.arrayBuffer()]);
      if (current !== revision) return;
      const sgBytes = new Uint8Array(sgData);
      if (myData.byteLength === sgData.byteLength &&
          new Uint8Array(myData).every((byte, index) => byte === sgBytes[index])) {
        throw new Error("The uploads are identical. Choose one export from each store.");
      }
      const my = readCollections(XLSX.read(myData), "MY");
      const sg = readCollections(XLSX.read(sgData), "SG");
      entries = compareCollections(my, sg);
      window.Runs?.add("Compared MY and SG collections", { items: entries.length, market: "MY/SG", started, file });
      destinations = sg;
      get("search").value = "";
      filter.value = "";
      get("counts").textContent = STATUSES.map(status =>
        status + " " + entries.filter(entry => entry.status === status).length).join(" · ");
      get("counts").prepend(my.size + " MY · " + sg.size + " SG · ");
      get("results").hidden = false;
      describe("");
      render();
    } catch (error) {
      if (current === revision) {
        describe(error.message || "Could not read the exports.", true);
        window.Runs?.add("Collection comparison failed", { market: "MY/SG", status: "Failed", started, file });
      }
    } finally {
      get("compare").disabled = false;
    }
  };
  get("search").oninput = render;
  filter.onchange = render;
  const dialog = get("export-preview");
  get("close-preview").onclick = () => dialog.close();

  function preview(action) {
    const { updates, creates, conflict } = pendingExports();
    if (action === "update" && conflict) return;
    const chosen = (action === "update" ? updates : creates)
      .map(entry => ({ ...entry, destination: destinationFor(entry) }));
    if (!chosen.length) return;
    const rows = workbookRows(chosen, action);
    const kind = action === "update" ? "Updates" : "Creations";
    const filename = "Vans-SG-Collections-" + kind + ".xlsx";
    get("preview-title").textContent = "Preview SG " + kind.toLowerCase();
    get("preview-description").textContent = filename + " · " + chosen.length + " collections · " +
      (rows.length - 1) + " data rows";
    const table = get("preview-rows");
    const head = document.createElement("thead");
    const header = head.insertRow();
    for (const value of rows[0]) {
      const cell = element("th", value);
      cell.scope = "col";
      header.append(cell);
    }
    const body = document.createElement("tbody");
    for (const values of rows.slice(1)) {
      const row = body.insertRow();
      for (const value of values) row.insertCell().textContent = typeof value === "boolean" ?
        String(value).toUpperCase() : value;
    }
    table.replaceChildren(head, body);
    get("download-preview").onclick = () => {
      const book = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Collections");
      XLSX.writeFile(book, filename);
      window.Runs?.add("Downloaded SG collection " + kind.toLowerCase(), { items: chosen.length, market: "SG", file: filename });
      dialog.close();
    };
    dialog.showModal();
    window.Runs?.add("Previewed SG collection " + kind.toLowerCase(), { items: chosen.length, market: "SG" });
  }

  get("updates").onclick = () => preview("update");
  get("creates").onclick = () => preview("create");
  get("report").onclick = () => {
    const summary = [["MY Handle", "MY Title", "SG Handle", "SG Title", "SG ID", "Comparison", "Decision", "Export action"]];
    const rules = [["Review Handle", "Market", "Collection Handle", "Source ID", "Source Type", ...RULE_COLUMNS]];
    const metadata = [["Review Handle", "Field", "MY", "SG"]];
    for (const entry of entries) {
      const { reference, status, decision, handle } = entry;
      const destination = destinationFor(entry);
      summary.push([reference?.handle || "", reference?.title || "", destination?.handle || "",
        destination?.title || "", destination?.id || "", status, decisions[decision],
        entry.applyMappedRules || decision === "update" ? "UPDATE" : decision === "create" ? "NEW" : ""]);
      for (const [market, collection] of [["MY", reference], ["SG", destination]]) {
        for (const source of collection?.sources.values() ?? []) {
          for (const rule of source.rules) rules.push([handle, market, collection.handle, source.id, source.type, ...rule]);
        }
      }
      for (const difference of metadataDifferences(reference, destination)) metadata.push([handle, ...difference]);
    }
    const book = XLSX.utils.book_new();
    for (const [name, rows] of [["Review decisions", summary], ["Rule details", rules], ["Metadata differences", metadata]]) {
      XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
    }
    const filename = "Vans-MY-SG-Collection-Review.xlsx";
    XLSX.writeFile(book, filename);
    window.Runs?.add("Downloaded collection comparison report", { items: entries.length, market: "MY/SG", file: filename });
  };
}

if (typeof document !== "undefined") initialize();
