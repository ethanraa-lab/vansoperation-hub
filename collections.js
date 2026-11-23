"use strict";

const SOURCE_ID = "Source: ID";
const SOURCE_TYPE = "Source: Type";
const RULE_COLUMNS = [
  "Inclusion: Type", "Inclusion: Match", "Condition: Field",
  "Condition: Relation", "Condition: Value", "Condition: Match"
];
const REQUIRED = ["ID", "Handle", "Title", SOURCE_ID, SOURCE_TYPE, ...RULE_COLUMNS];
const STATUSES = ["Matched", "Different rules", "Missing in SG", "Review", "SG only"];
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

function compareCollections(my, sg) {
  const entries = [];
  for (const handle of new Set([...my.keys(), ...sg.keys()])) {
    const reference = my.get(handle);
    const destination = sg.get(handle);
    let status = destination ? "Different rules" : "Missing in SG";
    let note = "";
    let action = "";
    if (!reference) {
      status = "SG only";
    } else {
      const mySource = [...reference.sources.values()][0];
      const sgSource = destination && [...destination.sources.values()][0];
      const equivalent = destination && signature(reference) === signature(destination);
      note = reviewReason(reference) || (destination && reviewReason(destination)) || "";
      if (!note && equivalent) {
        status = "Matched";
      } else {
        if (!note && destination && lower(mySource.type) !== lower(sgSource.type)) {
          note = "MY and SG use different source types.";
        }
        if (!note && destination && (!numericId(destination.id) || !numericId(sgSource.id))) {
          note = "SG requires valid collection and source IDs for updates.";
        }
        if (!note && !destination && !reference.title) note = "MY title is missing.";
        if (note) {
          status = "Review";
          if (!destination) note = "Missing in SG. " + note;
        } else {
          action = destination ? "update" : "create";
        }
      }
      if (destination && reference.title !== destination.title) {
        note += (note ? " " : "") + "MY and SG titles differ.";
      }
    }
    entries.push({ handle, reference, destination, status, note, action });
  }
  return entries.sort((a, b) => a.handle.localeCompare(b.handle));
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
  const selected = new Set();
  let entries = [];
  let revision = 0;
  const filter = get("filter");
  filter.add(new Option("All statuses", ""));
  for (const status of STATUSES) filter.add(new Option(status, status));

  function describe(text, error = false) {
    message.textContent = text;
    message.classList.toggle("error", error);
    message.hidden = !text;
  }

  function visibleEntries() {
    const search = lower(get("search").value);
    return entries.filter(entry => (!filter.value || entry.status === filter.value) &&
      [entry.handle, entry.reference?.title, entry.destination?.title].some(value => lower(value).includes(search)));
  }

  function syncSelection() {
    const eligible = visibleEntries().filter(entry => entry.action);
    const checked = eligible.filter(entry => selected.has(entry.handle)).length;
    const toggle = get("select-visible");
    toggle.disabled = !eligible.length;
    toggle.checked = Boolean(eligible.length && checked === eligible.length);
    toggle.indeterminate = checked > 0 && checked < eligible.length;
    const updates = entries.filter(entry => entry.action === "update" && selected.has(entry.handle)).length;
    const creates = selected.size - updates;
    get("updates").disabled = !updates;
    get("creates").disabled = !creates;
    get("selection").textContent = updates + " updates · " + creates + " creations selected";
  }

  function addRules(details, market, collection) {
    const label = document.createElement("strong");
    label.textContent = market;
    const list = document.createElement("ul");
    list.className = "rule-list";
    const lines = [...(collection?.sources.values() ?? [])].flatMap(source =>
      source.rules.map(rule => source.type + " · " + rule.filter(Boolean).join(" · ")));
    for (const line of lines.length ? lines : [collection ? "No rules in export" : "Collection absent"]) {
      const item = document.createElement("li");
      item.textContent = line;
      list.append(item);
    }
    details.append(label, list);
  }

  function render() {
    const tbody = get("rows");
    tbody.replaceChildren();
    for (const entry of visibleEntries()) {
      const row = tbody.insertRow();
      const checkboxCell = row.insertCell();
      if (entry.action) {
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = selected.has(entry.handle);
        checkbox.setAttribute("aria-label", "Select " + entry.handle);
        checkbox.onchange = () => {
          if (checkbox.checked) selected.add(entry.handle);
          else selected.delete(entry.handle);
          syncSelection();
        };
        checkboxCell.append(checkbox);
      }
      const collection = row.insertCell();
      const title = document.createElement("strong");
      title.textContent = entry.reference?.title || entry.destination?.title || entry.handle;
      const handle = document.createElement("div");
      handle.className = "muted";
      handle.textContent = entry.handle;
      collection.append(title, handle);
      const status = document.createElement("span");
      status.className = "status";
      status.dataset.status = entry.status;
      status.textContent = entry.status;
      row.insertCell().append(status);
      const comparison = row.insertCell();
      comparison.append(entry.note);
      const details = document.createElement("details");
      const summary = document.createElement("summary");
      summary.textContent = "View MY and SG rules";
      details.append(summary);
      addRules(details, "MY", entry.reference);
      addRules(details, "SG", entry.destination);
      comparison.append(details);
    }
    if (!tbody.rows.length) {
      const cell = tbody.insertRow().insertCell();
      cell.colSpan = 4;
      cell.textContent = "No collections match this filter.";
    }
    syncSelection();
  }

  for (const id of ["my-file", "sg-file"]) {
    get(id).onchange = () => {
      revision++;
      selected.clear();
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
      selected.clear();
      get("search").value = "";
      filter.value = "";
      get("counts").textContent = STATUSES.map(status =>
        status + " " + entries.filter(entry => entry.status === status).length).join(" · ");
      get("results").hidden = false;
      describe("");
      render();
    } catch (error) {
      if (current === revision) describe(error.message || "Could not read the exports.", true);
    } finally {
      get("compare").disabled = false;
    }
  };
  get("search").oninput = render;
  filter.onchange = render;
  get("select-visible").onchange = event => {
    for (const entry of visibleEntries().filter(entry => entry.action)) {
      if (event.target.checked) selected.add(entry.handle);
      else selected.delete(entry.handle);
    }
    render();
  };

  function download(action) {
    const chosen = entries.filter(entry => entry.action === action && selected.has(entry.handle));
    if (!chosen.length) return;
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(workbookRows(chosen, action)), "Collections");
    XLSX.writeFile(book, "Vans-SG-Collections-" + (action === "update" ? "Updates" : "Creations") + ".xlsx");
  }
  get("updates").onclick = () => download("update");
  get("creates").onclick = () => download("create");
}

if (typeof document !== "undefined") initialize();
