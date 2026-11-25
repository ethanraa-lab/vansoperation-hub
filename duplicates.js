"use strict";

const clean = value => String(value ?? "").trim();
const normalized = value => clean(value).toLowerCase();
const collectionSheet = name => /^(?:(?:smart|custom) )?collections?$/i.test(name);
const handleKey = "handle";

function inspectRows(matrix, sheetName, extraColumn = "") {
  const headers = (matrix[0] || []).map(normalized);
  const column = name => headers.indexOf(name);
  const has = name => column(name) >= 0;
  const productSheet = /^products?$/i.test(sheetName) ||
    ["variant id", "variant sku", "variant barcode"].some(has);
  const records = matrix.slice(1).map((values, index) => ({ values, row: index + 2 }))
    .filter(({ values }) => values.some(value => clean(value) !== ""));
  const groups = new Map();

  function add(check, value, record, identity, foldCase = false) {
    const key = foldCase ? normalized(value) : clean(value);
    if (!key) return;
    const groupKey = JSON.stringify([check, key]);
    if (!groups.has(groupKey)) groups.set(groupKey, { check, value: clean(value), ids: new Set(), rows: [] });
    const group = groups.get(groupKey);
    group.ids.add(identity);
    group.rows.push(record.row);
  }

  const idsByHandle = new Map();
  for (const { values } of records) {
    const handle = normalized(values[column(handleKey)]);
    const id = clean(values[column("id")]);
    if (handle && id) {
      if (!idsByHandle.has(handle)) idsByHandle.set(handle, new Set());
      idsByHandle.get(handle).add(id);
    }
  }

  const variantIdsByOptions = new Map();
  function productIdentity(id, handle) {
    const matchingIds = idsByHandle.get(normalized(handle));
    return id || (matchingIds?.size === 1 ? [...matchingIds][0] : "") ||
      (handle ? "handle:" + normalized(handle) : "");
  }
  function optionsKey(values) {
    const options = [1, 2, 3].map(n => clean(values[column("option" + n + " value")]));
    return options.some(Boolean) ? JSON.stringify(options) : "";
  }
  if (productSheet) {
    for (const { values } of records) {
      const id = clean(values[column("variant id")]);
      const options = optionsKey(values);
      if (!id || !options) continue;
      const product = productIdentity(clean(values[column("id")]), clean(values[column(handleKey)]));
      if (!product) continue;
      const key = JSON.stringify([product, options]);
      if (!variantIdsByOptions.has(key)) variantIdsByOptions.set(key, new Set());
      variantIdsByOptions.get(key).add(id);
    }
  }

  for (const record of records) {
    const get = name => clean(record.values[column(name)]);
    const id = get("id");
    const handle = get(handleKey);
    const product = productIdentity(id, handle);
    const entity = productSheet ? "Product" : "Collection";
    if (productSheet || collectionSheet(sheetName)) {
      if (product && handle) add(entity + " handle on different records", handle, record, product, true);
      if (id && handle) add(entity + " ID with different handles", id, record, normalized(handle));
    }
    if (productSheet) {
      const variantId = get("variant id");
      const options = optionsKey(record.values);
      const matchingVariants = options && variantIdsByOptions.get(JSON.stringify([product, options]));
      const resolvedId = variantId || (matchingVariants?.size === 1 ? [...matchingVariants][0] : "");
      const variant = resolvedId ? "id:" + resolvedId : options ? "options:" + options : "";
      if (variantId && product) add("Variant ID on different products", variantId, record, product);
      for (const [field, label] of [["variant sku", "Variant SKU"], ["variant barcode", "Variant barcode"]]) {
        const value = get(field);
        if (variantId || product) {
          // Missing variant IDs/options can still reveal reuse across different products.
          const identity = JSON.stringify([product, variant || "unidentified variant"]);
          add(label, value, record, identity);
        }
      }
    }
    const signature = JSON.stringify(headers.map((name, index) =>
      name && name !== "row #" && name !== "top row" ? clean(record.values[index]) : ""));
    add("Repeated row", signature, record, record.row);
    if (extraColumn) add("Repeated values in " + extraColumn, get(normalized(extraColumn)), record, record.row);
  }

  const issues = [...groups.values()].filter(group => group.ids.size > 1).map(group => ({
    check: group.check, value: group.check === "Repeated row" ? "Matching row contents" : group.value,
    count: group.ids.size, rows: group.rows
  })).sort((a, b) => a.check.localeCompare(b.check) || a.rows[0] - b.rows[0]);
  return { issues, scanned: records.length };
}

function initialize() {
  const get = id => document.getElementById(id);
  const fileInput = get("workbook");
  const sheetSelect = get("worksheet");
  const extra = get("extra");
  const filter = get("filter");
  const scan = get("scan");
  const download = get("download");
  const status = get("status");
  let workbook;
  let filename;
  let issues = [];
  let revision = 0;

  function message(text, error = false) {
    status.hidden = false;
    status.className = error ? "error" : "muted";
    status.textContent = text;
  }

  function readSheet() {
    const sheet = workbook.Sheets[sheetSelect.value];
    if (!sheet?.["!ref"]) return [];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0, raw: false, defval: "", blankrows: true });
  }

  function prepareSheet() {
    get("results").hidden = true;
    const headers = (readSheet()[0] || []).map(clean).filter(Boolean);
    extra.replaceChildren(new Option("None", ""), ...[...new Set(headers)].map(name => new Option(name, name)));
    extra.disabled = !headers.length;
    scan.disabled = !headers.length;
    message(headers.length ? "Ready to check " + sheetSelect.value + "." : "This sheet has no column headers.");
  }

  fileInput.onchange = async () => {
    const version = ++revision;
    workbook = null;
    get("results").hidden = true;
    scan.disabled = extra.disabled = sheetSelect.disabled = true;
    sheetSelect.replaceChildren();
    extra.replaceChildren();
    const file = fileInput.files[0];
    if (!file) { status.hidden = true; return; }
    message("Opening " + file.name + "…");
    try {
      const csv = /\.csv$/i.test(file.name);
      const input = csv ? (await file.text()).replace(/^\uFEFF/, "") : await file.arrayBuffer();
      if (version !== revision) return;
      workbook = XLSX.read(input, { type: csv ? "string" : "array", raw: csv });
      filename = file.name;
      sheetSelect.replaceChildren(...workbook.SheetNames.map(name => new Option(name, name)));
      if (workbook.Sheets.Products) sheetSelect.value = "Products";
      sheetSelect.disabled = false;
      prepareSheet();
    } catch {
      if (version === revision) message("Could not open this workbook. Choose a Matrixify XLSX or CSV export.", true);
    }
  };

  function render() {
    const body = get("issues");
    body.replaceChildren();
    const visible = issues.filter(issue => !filter.value || issue.check === filter.value);
    for (const issue of visible) {
      const row = body.insertRow();
      const shownRows = issue.rows.slice(0, 20).join(", ");
      for (const value of [issue.check, issue.value, issue.count,
        issue.rows.length > 20 ? shownRows + " … (see report for all rows)" : shownRows]) {
        row.insertCell().textContent = value;
      }
    }
    if (!visible.length) {
      const cell = body.insertRow().insertCell();
      cell.colSpan = 4;
      cell.textContent = "No duplicate groups found for these checks.";
    }
    download.disabled = !issues.length;
  }

  sheetSelect.onchange = prepareSheet;
  extra.onchange = () => { get("results").hidden = true; };
  filter.onchange = render;
  scan.onclick = () => {
    const started = Date.now();
    const result = inspectRows(readSheet(), sheetSelect.value, extra.value);
    issues = result.issues;
    const checks = [...new Set(issues.map(issue => issue.check))];
    filter.replaceChildren(new Option("All checks", ""), ...checks.map(check => new Option(check, check)));
    message(sheetSelect.value + " · " + result.scanned + " rows checked · " + issues.length + " duplicate groups");
    get("results").hidden = false;
    render();
    window.Runs?.add("Checked Matrixify duplicates", { file: filename, items: result.scanned, started });
  };
  download.onclick = () => {
    const book = XLSX.utils.book_new();
    const rows = [["Worksheet", "Group", "Check", "Value", "Records", "Matching row count"],
      ...issues.map((issue, index) => [sheetSelect.value, index + 1, issue.check, issue.value, issue.count,
        issue.rows.length])];
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "Duplicate groups");
    const matches = [["Worksheet", "Group", "Check", "Excel row"]];
    issues.forEach((issue, index) => {
      for (const row of issue.rows) matches.push([sheetSelect.value, index + 1, issue.check, row]);
    });
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(matches), "Matching rows");
    const file = "Matrixify-Duplicate-Report.xlsx";
    XLSX.writeFile(book, file);
    window.Runs?.add("Downloaded Matrixify duplicate report", { file, items: issues.length });
  };
}

if (typeof document !== "undefined") initialize();
