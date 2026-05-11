/* ───────────────────────────────────────── */
/*  Kitting Flow — task pane logic           */
/* ───────────────────────────────────────── */

// --- Configuration --------------------------------------------------------
// Edit this list to change which columns appear. Names must match the headers
// in the "All Data" sheet exactly (trailing spaces matter).

const SHEET_NAME = "All Data";

// Columns shown on each row card (key info)
const CARD_FIELDS = {
  primary: "JDE Module",          // big bold line
  secondary: "JDE Description",   // small line under it
  group: "Business Unit",         // chip top-right
  hospital: "Hospital",           // shown in the middle line
  status: "Status",               // colored chip
  canShip: "Can Module Ship?",    // colored chip
};

// Columns shown in the detail drawer, grouped. Add or remove freely.
const DETAIL_GROUPS = [
  {
    title: "Identification",
    fields: [
      { name: "JDE Module", editable: false },
      { name: "JDE Description", editable: false },
      { name: "JDE Build number", editable: false },
      { name: "SAP Module", editable: false },
      { name: "SAP Serial Number", editable: false },
      { name: "SAP Description", editable: false },
      { name: "Business Unit", editable: false },
    ],
  },
  {
    title: "Status",
    fields: [
      { name: "Status", editable: true },
      { name: "Equipment status", editable: true },
      { name: "Anchor Status (Auto)", editable: false },
      { name: "BEK Status (Auto)", editable: false },
      { name: "Can Module Ship?", editable: true },
      { name: "FOL Status (Auto)", editable: false },
      { name: "Completion Percentage", editable: false },
    ],
  },
  {
    title: "Destination",
    fields: [
      { name: "Hospital", editable: true },
      { name: "Allocated Country", editable: true },
      { name: "HospitalContact", editable: true },
      { name: "RequestedShipDate", editable: true },
      { name: "Final Shipment Week", editable: true },
      { name: "Shipment Route", editable: true },
    ],
  },
  {
    title: "Ordering",
    fields: [
      { name: "Week To Raise Inbound", editable: true },
      { name: "9SE1 order number", editable: true },
      { name: "PO", editable: true },
      { name: "Week Inbound Raised", editable: true },
      { name: "Requested By", editable: true },
    ],
  },
  {
    title: "Comments",
    fields: [
      { name: "Ordering Comment ", editable: true, multiline: true },
      { name: "BEK Ordering Issue (comments)", editable: true, multiline: true },
      { name: "S&OP COMMENTS", editable: true, multiline: true },
      { name: "Kit build comments", editable: true, multiline: true },
    ],
  },
];

const PAGE_SIZE = 50;

// Fields to show when creating a brand-new row.
// `required: true` means the user must fill it in before saving.
// Any field name here must also exist in the All Data sheet headers.
const NEW_ROW_FIELDS = [
  { name: "JDE Module",          required: true },
  { name: "JDE Description",     required: true },
  { name: "Business Unit",       required: true },
  { name: "Hospital",            required: true },
  { name: "JDE Build number",    required: false },
  { name: "SAP Module",          required: false },
  { name: "SAP Serial Number",   required: false },
  { name: "SAP Description",     required: false },
  { name: "Allocated Country",   required: false },
  { name: "Status",              required: false },
  { name: "Equipment status",    required: false },
  { name: "Quantity in Module",  required: false },
  { name: "Final Shipment Week", required: false },
  { name: "RequestedShipDate",   required: false },
  { name: "Requested By",        required: false },
  { name: "9SE1 order number",   required: false },
  { name: "PO",                  required: false },
  { name: "Ordering Comment ",   required: false, multiline: true },
  { name: "S&OP COMMENTS",       required: false, multiline: true },
  { name: "Kit build comments",  required: false, multiline: true },
];

// Filters that drive the dropdowns above the search results
const FILTER_DROPDOWNS = [
  { id: "filter-bu",       column: "Business Unit",   limit: 30 },
  { id: "filter-status",   column: "Status",          limit: 30 },
  { id: "filter-can-ship", column: "Can Module Ship?", limit: 20 },
];

// --- State ----------------------------------------------------------------
let allHeaders = [];        // header names in order
let headerIndex = {};       // header name -> column index
let allRows = [];           // each row: { _row: <sheet row 1-based>, [header]: value }
let filteredRows = [];
let currentPage = 0;
let editingRow = null;      // row currently in the drawer
let editedValues = {};      // { columnName: newValue }
let isCreating = false;     // true when drawer is in create-new mode

// --- Office bootstrap -----------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    showToast("This add-in must run in Excel.", "error");
    return;
  }
  bindUi();
  loadData();
});

function bindUi() {
  document.getElementById("refresh-btn").addEventListener("click", loadData);
  document.getElementById("new-row-btn").addEventListener("click", openCreateDrawer);
  document.getElementById("search").addEventListener("input", debounce(applyFilters, 200));
  document.getElementById("clear-filters").addEventListener("click", clearFilters);
  FILTER_DROPDOWNS.forEach((f) => {
    document.getElementById(f.id).addEventListener("change", applyFilters);
  });
  document.getElementById("prev-page").addEventListener("click", () => { currentPage--; renderRows(); });
  document.getElementById("next-page").addEventListener("click", () => { currentPage++; renderRows(); });
  document.getElementById("close-drawer").addEventListener("click", closeDrawer);
  document.getElementById("cancel-edit").addEventListener("click", closeDrawer);
  document.getElementById("save-edit").addEventListener("click", onSaveClick);
}

// --- Data loading ---------------------------------------------------------
async function loadData() {
  showEmpty("Loading data from \"" + SHEET_NAME + "\" sheet…");
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const used = sheet.getUsedRange(true /*valuesOnly*/);
      used.load(["values", "rowCount", "columnCount", "rowIndex"]);
      await ctx.sync();

      if (!used.values || used.values.length < 2) {
        allHeaders = []; allRows = [];
        showEmpty("The \"" + SHEET_NAME + "\" sheet is empty.");
        return;
      }

      allHeaders = used.values[0].map((h) => (h === null ? "" : String(h)));
      headerIndex = {};
      allHeaders.forEach((h, i) => { headerIndex[h] = i; });

      // Row index 0 is the header. Sheet row = used.rowIndex + i + 1
      const baseRow = used.rowIndex + 1; // sheet row of first data row (1-based)
      allRows = [];
      for (let i = 1; i < used.values.length; i++) {
        const row = used.values[i];
        // Skip rows that are fully empty
        const hasValue = row.some((v) => v !== "" && v !== null && v !== undefined);
        if (!hasValue) continue;
        const obj = { _row: baseRow + i };
        for (let c = 0; c < allHeaders.length; c++) {
          obj[allHeaders[c]] = row[c];
        }
        allRows.push(obj);
      }
    });

    populateFilterDropdowns();
    applyFilters();
    showToast("Loaded " + allRows.length + " rows", "success");
  } catch (err) {
    console.error(err);
    showEmpty("Couldn't load data. Make sure the sheet is named exactly \"" + SHEET_NAME + "\".");
    showToast("Load failed: " + (err.message || err), "error");
  }
}

// --- Filters & search -----------------------------------------------------
function populateFilterDropdowns() {
  FILTER_DROPDOWNS.forEach((f) => {
    const el = document.getElementById(f.id);
    // preserve the first "All ..." option, clear the rest
    const first = el.querySelector("option");
    el.innerHTML = "";
    el.appendChild(first);
    const seen = new Set();
    for (const row of allRows) {
      const v = row[f.column];
      if (v === undefined || v === null || v === "") continue;
      const s = String(v);
      if (seen.has(s)) continue;
      seen.add(s);
      if (seen.size > f.limit) break;
    }
    [...seen].sort().forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v; opt.textContent = v;
      el.appendChild(opt);
    });
  });
}

function clearFilters() {
  document.getElementById("search").value = "";
  FILTER_DROPDOWNS.forEach((f) => { document.getElementById(f.id).value = ""; });
  applyFilters();
}

function applyFilters() {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const filterValues = {};
  FILTER_DROPDOWNS.forEach((f) => {
    const v = document.getElementById(f.id).value;
    if (v) filterValues[f.column] = v;
  });

  filteredRows = allRows.filter((row) => {
    for (const col of Object.keys(filterValues)) {
      if (String(row[col] ?? "") !== filterValues[col]) return false;
    }
    if (q) {
      // search across the most useful text fields
      const haystack = [
        row[CARD_FIELDS.primary], row[CARD_FIELDS.secondary],
        row["SAP Module"], row["SAP Serial Number"], row["SAP Description"],
        row["Hospital"], row["Allocated Country"], row["JDE Build number"],
        row["Set Reference"], row["9SE1 order number"], row["PO"],
        row["Ordering Comment "], row["S&OP COMMENTS"], row["Kit build comments"],
      ].map((v) => (v == null ? "" : String(v).toLowerCase())).join(" | ");
      if (!haystack.includes(q)) return false;
    }
    return true;
  });

  currentPage = 0;
  renderRows();
}

// --- Rendering ------------------------------------------------------------
function renderRows() {
  const container = document.getElementById("rows-container");
  const empty = document.getElementById("empty-state");
  const pag = document.getElementById("pagination");
  const count = document.getElementById("result-count");

  count.textContent = filteredRows.length + " result" + (filteredRows.length === 1 ? "" : "s");

  if (filteredRows.length === 0) {
    container.classList.add("hidden");
    pag.classList.add("hidden");
    empty.classList.remove("hidden");
    empty.innerHTML = "<p>No matching rows. Try a different search or clear filters.</p>";
    return;
  }
  empty.classList.add("hidden");
  container.classList.remove("hidden");
  pag.classList.remove("hidden");

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  if (currentPage >= totalPages) currentPage = totalPages - 1;
  if (currentPage < 0) currentPage = 0;
  const start = currentPage * PAGE_SIZE;
  const slice = filteredRows.slice(start, start + PAGE_SIZE);

  container.innerHTML = "";
  slice.forEach((row) => {
    container.appendChild(buildRowCard(row));
  });

  document.getElementById("page-info").textContent =
    "Page " + (currentPage + 1) + " of " + totalPages;
  document.getElementById("prev-page").disabled = currentPage === 0;
  document.getElementById("next-page").disabled = currentPage >= totalPages - 1;
}

function buildRowCard(row) {
  const card = document.createElement("div");
  card.className = "row-card";
  card.tabIndex = 0;

  const top = document.createElement("div");
  top.className = "row-card-top";
  const mod = document.createElement("div");
  mod.className = "row-module";
  mod.textContent = displayValue(row[CARD_FIELDS.primary]) || "(no module)";
  const bu = document.createElement("div");
  bu.className = "row-bu";
  bu.textContent = displayValue(row[CARD_FIELDS.group]);
  top.appendChild(mod);
  top.appendChild(bu);

  const mid = document.createElement("div");
  mid.className = "row-card-mid";
  const desc = displayValue(row[CARD_FIELDS.secondary]);
  const hosp = displayValue(row[CARD_FIELDS.hospital]);
  mid.innerHTML = (desc ? escapeHtml(desc) : "") +
                  (hosp ? " · <span class='row-hospital'>" + escapeHtml(hosp) + "</span>" : "");

  const bot = document.createElement("div");
  bot.className = "row-card-bot";
  const statusVal = displayValue(row[CARD_FIELDS.status]);
  const canShipVal = displayValue(row[CARD_FIELDS.canShip]);
  if (statusVal) bot.appendChild(buildChip(statusVal, classifyStatus(statusVal)));
  if (canShipVal) bot.appendChild(buildChip(canShipVal, classifyCanShip(canShipVal)));

  card.appendChild(top);
  card.appendChild(mid);
  card.appendChild(bot);
  card.addEventListener("click", () => openDrawer(row));
  card.addEventListener("keydown", (e) => { if (e.key === "Enter") openDrawer(row); });
  return card;
}

function buildChip(text, kind) {
  const c = document.createElement("span");
  c.className = "chip" + (kind ? " chip-" + kind : "");
  c.textContent = text;
  return c;
}

function classifyStatus(v) {
  const s = String(v).toLowerCase();
  if (s.includes("shipped") || s.includes("complete")) return "good";
  if (s.includes("issue") || s.includes("fail") || s.includes("not")) return "bad";
  if (s.includes("await") || s.includes("ordered") || s.includes("progress")) return "warn";
  return "";
}
function classifyCanShip(v) {
  const s = String(v).toLowerCase();
  if (s === "yes" || s.includes("ready")) return "good";
  if (s === "no" || s.includes("not")) return "bad";
  return "warn";
}

// --- Detail drawer --------------------------------------------------------
function openDrawer(row) {
  isCreating = false;
  editingRow = row;
  editedValues = {};
  document.getElementById("save-edit").textContent = "Save changes";
  document.getElementById("detail-title").textContent =
    displayValue(row[CARD_FIELDS.primary]) || "Module details";

  const fieldsEl = document.getElementById("detail-fields");
  fieldsEl.innerHTML = "";
  DETAIL_GROUPS.forEach((group) => {
    const g = document.createElement("div");
    g.className = "field-group";
    const t = document.createElement("div");
    t.className = "field-group-title";
    t.textContent = group.title;
    g.appendChild(t);
    group.fields.forEach((f) => {
      // Only show fields that exist in the sheet
      if (!(f.name in headerIndex)) return;
      g.appendChild(buildField(f, row[f.name]));
    });
    fieldsEl.appendChild(g);
  });

  document.getElementById("detail-drawer").classList.remove("hidden");
}

function buildField(def, value) {
  const wrap = document.createElement("div");
  wrap.className = "field" + (def.editable ? "" : " readonly");
  const lbl = document.createElement("label");
  lbl.className = "field-label";
  lbl.textContent = def.name.trim();
  wrap.appendChild(lbl);

  if (!def.editable) {
    const v = document.createElement("div");
    v.className = "field-value";
    v.textContent = displayValue(value);
    wrap.appendChild(v);
    return wrap;
  }

  let input;
  if (def.multiline) {
    input = document.createElement("textarea");
    input.className = "field-textarea";
    input.rows = 3;
  } else {
    input = document.createElement("input");
    input.type = "text";
    input.className = "field-input";
  }
  input.value = value == null ? "" : String(value);
  input.addEventListener("input", () => {
    editedValues[def.name] = input.value;
  });
  wrap.appendChild(input);
  return wrap;
}

function closeDrawer() {
  document.getElementById("detail-drawer").classList.add("hidden");
  document.getElementById("validation-msg").classList.add("hidden");
  editingRow = null;
  editedValues = {};
  isCreating = false;
}

function openCreateDrawer() {
  if (allHeaders.length === 0) {
    showToast("Data hasn't loaded yet — click Refresh first.", "error");
    return;
  }
  isCreating = true;
  editingRow = null;
  editedValues = {};
  document.getElementById("detail-title").textContent = "New row";
  document.getElementById("save-edit").textContent = "Create row";

  const fieldsEl = document.getElementById("detail-fields");
  fieldsEl.innerHTML = "";
  const group = document.createElement("div");
  group.className = "field-group";
  const title = document.createElement("div");
  title.className = "field-group-title";
  title.textContent = "New row — required fields marked with *";
  group.appendChild(title);

  NEW_ROW_FIELDS.forEach((f) => {
    if (!(f.name in headerIndex)) return; // skip fields not in the sheet
    group.appendChild(buildCreateField(f));
  });
  fieldsEl.appendChild(group);
  document.getElementById("detail-drawer").classList.remove("hidden");
}

function buildCreateField(def) {
  const wrap = document.createElement("div");
  wrap.className = "field";
  const lbl = document.createElement("label");
  lbl.className = "field-label";
  lbl.textContent = def.name.trim();
  if (def.required) {
    const star = document.createElement("span");
    star.className = "field-required-marker";
    star.textContent = "*";
    lbl.appendChild(star);
  }
  wrap.appendChild(lbl);

  let input;
  if (def.multiline) {
    input = document.createElement("textarea");
    input.className = "field-textarea";
    input.rows = 3;
  } else {
    input = document.createElement("input");
    input.type = "text";
    input.className = "field-input";
  }
  input.dataset.fieldName = def.name;
  if (def.required) input.dataset.required = "true";
  input.addEventListener("input", () => {
    editedValues[def.name] = input.value;
    input.classList.remove("field-error");
  });
  wrap.appendChild(input);
  return wrap;
}

function onSaveClick() {
  if (isCreating) {
    createNewRow();
  } else {
    saveEdits();
  }
}

async function saveEdits() {
  if (!editingRow) return;
  const changedKeys = Object.keys(editedValues);
  if (changedKeys.length === 0) {
    closeDrawer();
    return;
  }
  const btn = document.getElementById("save-edit");
  btn.disabled = true; btn.textContent = "Saving…";

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const sheetRow = editingRow._row; // 1-based
      changedKeys.forEach((col) => {
        const colIdx = headerIndex[col];
        if (colIdx === undefined) return;
        const addr = colLetter(colIdx) + sheetRow;
        const cell = sheet.getRange(addr);
        cell.values = [[ editedValues[col] ]];
      });
      await ctx.sync();
    });
    // update local cache
    changedKeys.forEach((col) => { editingRow[col] = editedValues[col]; });
    showToast("Saved " + changedKeys.length + " change" + (changedKeys.length === 1 ? "" : "s"), "success");
    renderRows();
    closeDrawer();
  } catch (err) {
    console.error(err);
    showToast("Save failed: " + (err.message || err), "error");
  } finally {
    btn.disabled = false; btn.textContent = "Save changes";
  }
}

async function createNewRow() {
  // Validate required fields
  const missing = [];
  const inputs = document.querySelectorAll("#detail-fields [data-field-name]");
  inputs.forEach((inp) => {
    inp.classList.remove("field-error");
    if (inp.dataset.required === "true") {
      const v = inp.value.trim();
      if (!v) {
        missing.push(inp.dataset.fieldName.trim());
        inp.classList.add("field-error");
      }
    }
  });

  const valMsg = document.getElementById("validation-msg");
  if (missing.length > 0) {
    valMsg.textContent = "Please fill in: " + missing.join(", ");
    valMsg.classList.remove("hidden");
    return;
  }
  valMsg.classList.add("hidden");

  const btn = document.getElementById("save-edit");
  btn.disabled = true; btn.textContent = "Creating…";

  try {
    let newRowNumber = 0;
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const used = sheet.getUsedRange(true);
      used.load(["rowCount", "rowIndex"]);
      await ctx.sync();

      // Append after the last used row.
      // Sheet row = used.rowIndex (0-based) + rowCount + 1 (to get 1-based next row)
      newRowNumber = used.rowIndex + used.rowCount + 1;

      // Build a single full-width row of values (empty string for unset cells)
      const rowValues = new Array(allHeaders.length).fill("");
      Object.keys(editedValues).forEach((col) => {
        const idx = headerIndex[col];
        if (idx !== undefined) rowValues[idx] = editedValues[col];
      });

      const lastColLetter = colLetter(allHeaders.length - 1);
      const addr = "A" + newRowNumber + ":" + lastColLetter + newRowNumber;
      sheet.getRange(addr).values = [rowValues];

      // Scroll the user to the new row so they can see it landed
      sheet.getRange("A" + newRowNumber).select();
      await ctx.sync();
    });

    // Update local cache so the new row shows up immediately in the list
    const newRowObj = { _row: newRowNumber };
    allHeaders.forEach((h) => { newRowObj[h] = editedValues[h] || ""; });
    allRows.push(newRowObj);

    showToast("New row created in sheet (row " + newRowNumber + ")", "success");
    populateFilterDropdowns();
    applyFilters();
    closeDrawer();
  } catch (err) {
    console.error(err);
    showToast("Create failed: " + (err.message || err), "error");
  } finally {
    btn.disabled = false; btn.textContent = "Create row";
  }
}

// --- Helpers --------------------------------------------------------------
function colLetter(idx) {
  // 0 -> A, 25 -> Z, 26 -> AA, etc.
  let s = ""; let n = idx;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}
function displayValue(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "number" && !Number.isInteger(v)) {
    return Math.abs(v) < 1 ? v.toFixed(3) : v.toFixed(2);
  }
  return String(v);
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));
}
function debounce(fn, ms) {
  let t; return function () {
    clearTimeout(t);
    const args = arguments, ctx = this;
    t = setTimeout(() => fn.apply(ctx, args), ms);
  };
}
function showEmpty(msg) {
  document.getElementById("empty-state").classList.remove("hidden");
  document.getElementById("empty-state").innerHTML = "<p>" + escapeHtml(msg) + "</p>";
  document.getElementById("rows-container").classList.add("hidden");
  document.getElementById("pagination").classList.add("hidden");
}
function showToast(msg, kind) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast " + (kind || "");
  setTimeout(() => { t.className = "toast hidden"; }, 2400);
}
