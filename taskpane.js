/* ───────────────────────────────────────── */
/*  Kitting Flow v2 — task pane logic        */
/* ───────────────────────────────────────── */

// =========================================================================
// CONFIG  — most things you'd tweak live here at the top
// =========================================================================

const SHEET_NAME = "All Data";

// Data range — DO NOT use getUsedRange (phantom formatting in this file).
const DATA_FIRST_COL = "A";
const DATA_LAST_COL = "AX";   // 50 real columns
const HEADER_ROW = 1;
const MAX_DATA_ROW = 20000;
const CHUNK_ROWS = 500;

// What appears on each Browse card
const CARD_FIELDS = {
  primary:   "JDE Module",
  secondary: "JDE Description",
  group:     "Business Unit",
  hospital:  "Hospital",
  status:    "Status",
  canShip:   "Can Module Ship?",
};

// Fields the search bar looks across
const SEARCH_FIELDS = [
  "JDE Module", "JDE Description", "JDE Build number", "Module_Build",
  "SAP Module", "SAP Serial Number", "SAP Description",
  "Hospital", "HospitalContact",
  "Loanset", "Set Reference", "Anchor Unique Reference (Auto)",
  "9SE1 order number", "PO", "FID ",
  "YU /PR", "Allocated Country",
  "Ordering Comment ", "S&OP COMMENTS", "Kit build comments",
  "BEK Ordering Issue (comments)",
];

// Top-of-page filter dropdowns
const FILTER_DROPDOWNS = [
  { id: "filter-bu",       column: "Business Unit", limit: 30 },
  { id: "filter-week-in",  column: "Week To Raise Inbound", limit: 60, sort: "weekish" },
  { id: "filter-forecast", column: "ForecastMonth", limit: 60, sort: "monthYear" },
  { id: "filter-bek",      column: "BEK Status (Auto)", limit: 15 },
  { id: "filter-fol",      column: "FOL Status (Auto)", limit: 15 },
];

// Values considered "hidden by default" — anything matching these in Status
// is suppressed when the "Hide shipped / #REF!" toggle is on.
const HIDDEN_STATUSES = ["EQUIPMENT SHIPPED", "EQUIPMENT DISSOLVED", "#REF!"];

// What counts as an "issue" for the Only-issues toggle
function rowIsIssue(row) {
  const bek = String(row["BEK Status (Auto)"] || "");
  const ship = String(row["Can Module Ship?"] || "");
  const issueLoc = String(row["ISSUE LOCATION (Auto)"] || "");
  const status = String(row["Status"] || "");
  if (bek === "Incomplete" || bek === "Do not ship") return true;
  if (ship === "FALSE") return true;
  if (issueLoc && issueLoc !== "NO CURRENT KNOWN ISSUE OR AWAITING REFILL") return true;
  if (status.includes("ISSUE")) return true;
  return false;
}

// Drawer field groups (unchanged from v1 mostly)
const DETAIL_GROUPS = [
  {
    title: "Identification",
    fields: [
      { name: "JDE Module", editable: false },
      { name: "JDE Description", editable: false },
      { name: "JDE Build number", editable: false },
      { name: "Module_Build", editable: false },
      { name: "SAP Module", editable: false },
      { name: "SAP Serial Number", editable: false },
      { name: "SAP Description", editable: false },
      { name: "Business Unit", editable: false },
      { name: "Loanset", editable: false },
    ],
  },
  {
    title: "Status",
    fields: [
      { name: "Status", editable: true },
      { name: "Equipment status", editable: true },
      { name: "BEK Status (Auto)", editable: false },
      { name: "Can Module Ship?", editable: true },
      { name: "FOL Status (Auto)", editable: false },
      { name: "Completion Percentage", editable: false },
      { name: "SAP Equipment Monitor Status (AUTO)", editable: false },
      { name: "SAP Object Type (Auto)", editable: false },
      { name: "SAP Storage Location (Auto)", editable: false },
      { name: "ISSUE LOCATION (Auto)", editable: false },
    ],
  },
  {
    title: "Destination",
    fields: [
      { name: "Hospital", editable: true },
      { name: "Allocated Country", editable: true },
      { name: "HospitalContact", editable: true },
      { name: "Final Shipment Week", editable: true },
      { name: "Set Reference", editable: true },
      { name: "FID ", editable: true },
    ],
  },
  {
    title: "Ordering",
    fields: [
      { name: "Week To Raise Inbound", editable: true },
      { name: "9SE1 order number", editable: true },
      { name: "PO", editable: true },
      { name: "YU /PR", editable: true },
      { name: "Week Inbound Raised", editable: true },
      { name: "Requested By", editable: true },
      { name: "Raise 9SE2", editable: true },
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

const NEW_ROW_FIELDS = [
  { name: "JDE Module", required: true },
  { name: "JDE Description", required: true },
  { name: "Business Unit", required: true },
  { name: "Hospital", required: true },
  { name: "JDE Build number", required: false },
  { name: "SAP Module", required: false },
  { name: "SAP Serial Number", required: false },
  { name: "Loanset", required: false },
  { name: "Allocated Country", required: false },
  { name: "Status", required: false },
  { name: "Quantity in Module", required: false },
  { name: "Final Shipment Week", required: false },
  { name: "Requested By", required: false },
  { name: "9SE1 order number", required: false },
  { name: "PO", required: false },
  { name: "Ordering Comment ", required: false, multiline: true },
  { name: "S&OP COMMENTS", required: false, multiline: true },
];

const PAGE_SIZE = 50;

// Week view: how many weeks in the future to bucket separately
const WEEK_LOOK_AHEAD = 4;

// =========================================================================
// STATE
// =========================================================================
let allHeaders = [];
let headerIndex = {};
let allRows = [];
let filteredRows = [];
let currentPage = 0;
let editingRow = null;
let editedValues = {};
let isCreating = false;
let currentView = "browse";

// =========================================================================
// BOOTSTRAP
// =========================================================================
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

  // Tabs
  document.querySelectorAll(".tab").forEach((t) => {
    t.addEventListener("click", () => switchView(t.dataset.view));
  });

  // Search & filters
  document.getElementById("search").addEventListener("input", debounce(applyFilters, 200));
  document.getElementById("clear-filters").addEventListener("click", clearFilters);
  document.getElementById("hide-shipped").addEventListener("change", applyFilters);
  document.getElementById("only-issues").addEventListener("change", applyFilters);
  FILTER_DROPDOWNS.forEach((f) => {
    document.getElementById(f.id).addEventListener("change", applyFilters);
  });

  document.getElementById("prev-page").addEventListener("click", () => { currentPage--; renderRows(); });
  document.getElementById("next-page").addEventListener("click", () => { currentPage++; renderRows(); });

  document.getElementById("close-drawer").addEventListener("click", closeDrawer);
  document.getElementById("cancel-edit").addEventListener("click", closeDrawer);
  document.getElementById("save-edit").addEventListener("click", onSaveClick);

  document.getElementById("week-refresh").addEventListener("click", renderWeekView);
}

function switchView(name) {
  currentView = name;
  document.querySelectorAll(".tab").forEach((t) => {
    t.classList.toggle("active", t.dataset.view === name);
  });
  document.querySelectorAll(".view").forEach((v) => {
    v.classList.toggle("active", v.id === "view-" + name);
  });
  if (name === "dashboard") renderDashboard();
  if (name === "week") renderWeekView();
}

// =========================================================================
// DATA LOADING
// =========================================================================
async function loadData() {
  showEmpty("Initialising…");
  allRows = [];
  allHeaders = [];
  headerIndex = {};

  try {
    showEmpty("Reading headers…");
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const addr = DATA_FIRST_COL + HEADER_ROW + ":" + DATA_LAST_COL + HEADER_ROW;
      const range = sheet.getRange(addr);
      range.load("values");
      await ctx.sync();
      allHeaders = range.values[0].map((h) => (h === null ? "" : String(h)));
      allHeaders.forEach((h, i) => { headerIndex[h] = i; });
    });

    const firstDataRow = HEADER_ROW + 1;
    let offset = 0;
    let emptyStreak = 0;
    const EMPTY_STREAK_LIMIT = 3;

    while (offset < MAX_DATA_ROW - firstDataRow) {
      const size = Math.min(CHUNK_ROWS, MAX_DATA_ROW - firstDataRow - offset);
      const a = firstDataRow + offset;
      const b = a + size - 1;
      let chunkNonEmpty = 0;

      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
        const addr = DATA_FIRST_COL + a + ":" + DATA_LAST_COL + b;
        const range = sheet.getRange(addr);
        range.load("values");
        await ctx.sync();

        const values = range.values;
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const hasValue = row.some((v) => v !== "" && v !== null && v !== undefined);
          if (!hasValue) continue;
          chunkNonEmpty++;
          const obj = { _row: a + i };
          for (let c = 0; c < allHeaders.length; c++) {
            obj[allHeaders[c]] = normalizeCellValue(allHeaders[c], row[c]);
          }
          allRows.push(obj);
        }
      });

      offset += size;
      if (chunkNonEmpty === 0) {
        emptyStreak++;
        if (emptyStreak >= EMPTY_STREAK_LIMIT) break;
      } else {
        emptyStreak = 0;
      }
      const approxPct = Math.min(99, Math.round((allRows.length / 16000) * 100));
      showEmpty("Loading… ~" + approxPct + "% (" + allRows.length.toLocaleString() + " rows)");
    }

    populateFilterDropdowns();
    applyFilters();
    if (currentView === "dashboard") renderDashboard();
    if (currentView === "week") renderWeekView();
    showToast("Loaded " + allRows.length.toLocaleString() + " rows", "success");
  } catch (err) {
    console.error(err);
    showEmpty("Couldn't load data. " + (err.message || err));
    showToast("Load failed: " + (err.message || err), "error");
  }
}

// =========================================================================
// FILTERS
// =========================================================================
function populateFilterDropdowns() {
  FILTER_DROPDOWNS.forEach((f) => {
    const el = document.getElementById(f.id);
    const first = el.querySelector("option");
    el.innerHTML = "";
    el.appendChild(first);
    const seen = new Set();
    for (const row of allRows) {
      const v = row[f.column];
      if (v === undefined || v === null || v === "") continue;
      const s = String(v);
      if (s === "#REF!") continue;
      if (!seen.has(s)) seen.add(s);
    }
    const sortFn = pickSorter(f.sort);
    [...seen].sort(sortFn).slice(0, f.limit).forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v; opt.textContent = v;
      el.appendChild(opt);
    });
  });
}

function pickSorter(kind) {
  if (kind === "weekish") {
    // Sort by parseable week number (1-53) first, then year values, then text
    return (a, b) => {
      const wa = parseWeek(a), wb = parseWeek(b);
      if (wa !== null && wb !== null) return wa - wb;
      if (wa !== null) return -1;
      if (wb !== null) return 1;
      return String(a).localeCompare(String(b));
    };
  }
  if (kind === "monthYear") {
    // "Jan-25" / "Feb-26" — sort chronologically
    return (a, b) => {
      const da = parseMonthYear(a), db = parseMonthYear(b);
      if (da && db) return da - db;
      if (da) return -1;
      if (db) return 1;
      return String(a).localeCompare(String(b));
    };
  }
  return (a, b) => String(a).localeCompare(String(b));
}

function parseMonthYear(s) {
  // "Jan-25" -> Date(2025, 0, 1)
  const m = String(s).trim().match(/^([A-Za-z]{3})-(\d{2,4})$/);
  if (!m) return null;
  const months = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
  const mo = months[m[1].toLowerCase()];
  if (mo === undefined) return null;
  let year = parseInt(m[2], 10);
  if (year < 100) year += 2000;
  return new Date(year, mo, 1).getTime();
}

function clearFilters() {
  document.getElementById("search").value = "";
  document.getElementById("hide-shipped").checked = true;
  document.getElementById("only-issues").checked = false;
  FILTER_DROPDOWNS.forEach((f) => { document.getElementById(f.id).value = ""; });
  applyFilters();
}

function applyFilters() {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const hideShipped = document.getElementById("hide-shipped").checked;
  const onlyIssues = document.getElementById("only-issues").checked;
  const filterValues = {};
  FILTER_DROPDOWNS.forEach((f) => {
    const v = document.getElementById(f.id).value;
    if (v) filterValues[f.column] = v;
  });

  filteredRows = allRows.filter((row) => {
    if (hideShipped) {
      const status = String(row["Status"] || "");
      if (HIDDEN_STATUSES.includes(status)) return false;
    }
    if (onlyIssues && !rowIsIssue(row)) return false;
    for (const col of Object.keys(filterValues)) {
      if (String(row[col] ?? "") !== filterValues[col]) return false;
    }
    if (q) {
      const haystack = SEARCH_FIELDS
        .map((f) => row[f] == null ? "" : String(row[f]).toLowerCase())
        .join(" | ");
      if (!haystack.includes(q)) return false;
    }
    return true;
  });

  currentPage = 0;
  renderRows();
}

// =========================================================================
// BROWSE RENDER
// =========================================================================
function renderRows() {
  const container = document.getElementById("rows-container");
  const empty = document.getElementById("empty-state");
  const pag = document.getElementById("pagination");
  const count = document.getElementById("result-count");

  count.textContent = filteredRows.length.toLocaleString() + " result" + (filteredRows.length === 1 ? "" : "s");

  if (filteredRows.length === 0) {
    container.classList.add("hidden");
    pag.classList.add("hidden");
    empty.classList.remove("hidden");
    empty.innerHTML = "<p>No matching rows. Try different filters or click Clear.</p>";
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
  slice.forEach((row) => container.appendChild(buildRowCard(row)));

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
  const loanset = displayValue(row["Loanset"]);
  let midHtml = "";
  if (desc) midHtml += escapeHtml(desc);
  if (hosp) midHtml += (midHtml ? " · " : "") + "<span class='row-hospital'>" + escapeHtml(hosp) + "</span>";
  if (loanset) midHtml += (midHtml ? " · " : "") + "Loanset " + escapeHtml(loanset);
  mid.innerHTML = midHtml;

  const bot = document.createElement("div");
  bot.className = "row-card-bot";
  const statusVal = displayValue(row[CARD_FIELDS.status]);
  const consign = displayValue(row["SAP Equipment Monitor Status (AUTO)"]);
  const bek = displayValue(row["BEK Status (Auto)"]);
  if (statusVal) bot.appendChild(buildChip(statusVal, classifyStatus(statusVal)));
  if (consign) bot.appendChild(buildChip(shortConsign(consign), classifyConsign(consign)));
  if (bek && bek !== "NOT IN BEK") bot.appendChild(buildChip("BEK: " + bek, classifyBek(bek)));

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
  if (s === "#ref!") return "bad";
  if (s.includes("shipped") || s.includes("built")) return "good";
  if (s.includes("issue") || s.includes("fail")) return "bad";
  if (s.includes("await") || s.includes("orthokit") || s.includes("topup")) return "warn";
  if (s.includes("dissolved")) return "info";
  return "";
}
function classifyConsign(v) {
  const s = String(v).toLowerCase();
  if (s.includes("free for use")) return "good";
  if (s.includes("blocked") || s.includes("unable")) return "bad";
  if (s.includes("at customer") || s.includes("reserved")) return "info";
  if (s.includes("dissolved") || s.includes("repair")) return "warn";
  return "";
}
function shortConsign(v) {
  // "loanset in consignment (E0010)" -> "In consignment"
  const m = String(v).match(/loanset\s+(.+?)\s+\(/i);
  if (m) return capitalize(m[1]);
  if (String(v).includes("UNABLE")) return "Unable to find";
  return v;
}
function classifyBek(v) {
  const s = String(v).toLowerCase();
  if (s === "complete" || s === "complete excess") return "good";
  if (s === "incomplete") return "warn";
  if (s === "do not ship") return "bad";
  return "";
}

// =========================================================================
// DETAIL DRAWER
// =========================================================================
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
  title.textContent = "New row — required fields marked *";
  group.appendChild(title);

  NEW_ROW_FIELDS.forEach((f) => {
    if (!(f.name in headerIndex)) return;
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
  if (isCreating) createNewRow();
  else saveEdits();
}

async function saveEdits() {
  if (!editingRow) return;
  const changedKeys = Object.keys(editedValues);
  if (changedKeys.length === 0) { closeDrawer(); return; }
  const btn = document.getElementById("save-edit");
  btn.disabled = true; btn.textContent = "Saving…";

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const sheetRow = editingRow._row;
      changedKeys.forEach((col) => {
        const colIdx = headerIndex[col];
        if (colIdx === undefined) return;
        const addr = colLetter(colIdx) + sheetRow;
        sheet.getRange(addr).values = [[ editedValues[col] ]];
      });
      await ctx.sync();
    });
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
  const missing = [];
  const inputs = document.querySelectorAll("#detail-fields [data-field-name]");
  inputs.forEach((inp) => {
    inp.classList.remove("field-error");
    if (inp.dataset.required === "true") {
      if (!inp.value.trim()) {
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
    let newRowNumber = HEADER_ROW + 1;
    for (const r of allRows) if (r._row >= newRowNumber) newRowNumber = r._row + 1;

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const rowValues = new Array(allHeaders.length).fill("");
      Object.keys(editedValues).forEach((col) => {
        const idx = headerIndex[col];
        if (idx !== undefined) rowValues[idx] = editedValues[col];
      });
      const lastColLetter = colLetter(allHeaders.length - 1);
      const addr = "A" + newRowNumber + ":" + lastColLetter + newRowNumber;
      sheet.getRange(addr).values = [rowValues];
      sheet.getRange("A" + newRowNumber).select();
      await ctx.sync();
    });

    const newRowObj = { _row: newRowNumber };
    allHeaders.forEach((h) => { newRowObj[h] = editedValues[h] || ""; });
    allRows.push(newRowObj);

    showToast("New row created (row " + newRowNumber + ")", "success");
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

// =========================================================================
// DASHBOARD
// =========================================================================
function renderDashboard() {
  const root = document.getElementById("dashboard-content");
  if (allRows.length === 0) {
    root.innerHTML = "<p class='empty-state'>Load data first.</p>";
    return;
  }

  // Only count "active" rows (not shipped/dissolved/#REF!)
  const active = allRows.filter((r) => !HIDDEN_STATUSES.includes(String(r["Status"] || "")));

  const bekCounts = countBy(active, "BEK Status (Auto)", (v) => v && v !== "NOT IN BEK");
  const folCounts = countBy(active, "FOL Status (Auto)", (v) => v && v !== "#REF!" && v !== "Please enter FOL ID");
  const hospitalCounts = countBy(active, "Hospital", (v) => v && !["Enter Ship To","ORTHOKIT","FORECAST ORDER","Not in Hospital Master",""].includes(v));

  // Summary headline tiles
  const issueCount = active.filter(rowIsIssue).length;
  const incompleteCount = active.filter((r) => String(r["BEK Status (Auto)"]) === "Incomplete").length;
  const doNotShipCount = active.filter((r) => String(r["BEK Status (Auto)"]) === "Do not ship").length;

  let html = "";

  html += dashSection("Overview", `
    <div class="dash-grid">
      ${tile("Active modules", active.length.toLocaleString(), "all non-shipped rows", "info")}
      ${tile("With issues", issueCount.toLocaleString(), "incomplete / blocked / KB issue", issueCount > 0 ? "bad" : "good", () => applyDashboardFilter({ onlyIssues: true }))}
      ${tile("BEK incomplete", incompleteCount.toLocaleString(), "needs completion", incompleteCount > 0 ? "warn" : "good", () => applyDashboardFilter({ bek: "Incomplete" }))}
      ${tile("Do not ship", doNotShipCount.toLocaleString(), "BEK flagged", doNotShipCount > 0 ? "bad" : "good", () => applyDashboardFilter({ bek: "Do not ship" }))}
    </div>
  `);

  const weekInCounts = countBy(active, "Week To Raise Inbound", (v) => v && v !== "#REF!");
  const forecastCounts = countBy(active, "ForecastMonth", (v) => v && v !== "#REF!");

  html += dashSection("Upcoming inbound weeks", listTile(sortCountsForWeek(weekInCounts), 10, () => {}));
  html += dashSection("Upcoming forecast months", listTile(sortCountsByMonthYear(forecastCounts), 10, () => {}));
  html += dashSection("BEK status", listTile(bekCounts, 6, (label) => applyDashboardFilter({ bek: label })));
  html += dashSection("FOL status", listTile(folCounts, 6, (label) => applyDashboardFilter({ fol: label })));
  html += dashSection("Top hospitals (active modules)", listTile(hospitalCounts, 10, (label) => applyDashboardFilter({ search: label })));

  root.innerHTML = html;

  // Wire up clickable list rows
  root.querySelectorAll("[data-dash-action]").forEach((el) => {
    el.addEventListener("click", () => {
      const action = el.dataset.dashAction;
      const value = el.dataset.dashValue;
      dispatchDashboardAction(action, value);
    });
  });
}

function dashSection(title, contentHtml) {
  return `
    <div class="dash-section">
      <div class="dash-section-title">${escapeHtml(title)}</div>
      ${contentHtml}
    </div>
  `;
}

function tile(label, value, sub, kind, onClick) {
  const cls = kind ? " " + kind : "";
  // We store the click action via a data attr — wired up after render.
  // Tile click actions are inlined as plain handlers here.
  return `
    <div class="dash-tile${cls}">
      <div class="dash-tile-label">${escapeHtml(label)}</div>
      <div class="dash-tile-value">${escapeHtml(value)}</div>
      <div class="dash-tile-sub">${escapeHtml(sub)}</div>
    </div>
  `;
}

function listTile(counts, max, onClickFactory) {
  // counts can be either a plain {label: count} object OR an array of [label, count] tuples
  // that has already been sorted as desired.
  let entries;
  if (Array.isArray(counts)) {
    entries = counts.slice(0, max);
  } else {
    entries = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, max);
  }
  if (entries.length === 0) return "<div class='dash-list-tile'><div class='dash-list-row'>No data</div></div>";
  const rows = entries.map(([label, n]) => `
    <div class="dash-list-row" data-dash-action="generic" data-dash-value="${escapeAttr(label)}">
      <span class="dash-list-name" title="${escapeAttr(label)}">${escapeHtml(label)}</span>
      <span class="dash-list-count">${n.toLocaleString()}</span>
    </div>
  `).join("");
  return `<div class="dash-list-tile">${rows}</div>`;
}

// Sort week-style values chronologically (1, 2, ..., 53, then year-strings, then text)
function sortCountsForWeek(counts) {
  return Object.entries(counts).sort((a, b) => {
    const wa = parseWeek(a[0]), wb = parseWeek(b[0]);
    if (wa !== null && wb !== null) return wa - wb;
    if (wa !== null) return -1;
    if (wb !== null) return 1;
    return String(a[0]).localeCompare(String(b[0]));
  });
}

// Sort month-year strings (Jan-25, Feb-25, ...) chronologically
function sortCountsByMonthYear(counts) {
  return Object.entries(counts).sort((a, b) => {
    const da = parseMonthYear(a[0]), db = parseMonthYear(b[0]);
    if (da && db) return da - db;
    if (da) return -1;
    if (db) return 1;
    return String(a[0]).localeCompare(String(b[0]));
  });
}

// Apply a dashboard tile click — switches to Browse and sets the relevant filter
function applyDashboardFilter(opts) {
  switchView("browse");
  document.getElementById("search").value = opts.search || "";
  document.getElementById("hide-shipped").checked = true;
  document.getElementById("only-issues").checked = !!opts.onlyIssues;
  document.getElementById("filter-bek").value = opts.bek || "";
  document.getElementById("filter-fol").value = opts.fol || "";
  document.getElementById("filter-week-in").value = opts.weekIn || "";
  document.getElementById("filter-forecast").value = opts.forecast || "";
  applyFilters();
}

function dispatchDashboardAction(action, value) {
  const candidates = ["filter-bek", "filter-fol", "filter-week-in", "filter-forecast", "filter-bu"];
  for (const id of candidates) {
    const sel = document.getElementById(id);
    for (const opt of sel.options) {
      if (opt.value === value) {
        switchView("browse");
        clearFilters();
        sel.value = value;
        applyFilters();
        return;
      }
    }
  }
  // Fallback: treat as a search term
  switchView("browse");
  clearFilters();
  document.getElementById("search").value = value;
  applyFilters();
}

function countBy(rows, col, filterFn) {
  const counts = {};
  for (const r of rows) {
    const v = r[col];
    if (v === undefined || v === null) continue;
    const s = String(v);
    if (s === "") continue;
    if (filterFn && !filterFn(s)) continue;
    counts[s] = (counts[s] || 0) + 1;
  }
  return counts;
}

// =========================================================================
// WEEK VIEW
// =========================================================================
function renderWeekView() {
  const root = document.getElementById("week-columns");
  if (allRows.length === 0) {
    root.innerHTML = "<p class='empty-state'>Load data first.</p>";
    return;
  }

  const today = new Date();
  const currentWeek = isoWeek(today);
  const currentYear = today.getFullYear();
  document.getElementById("week-current-label").textContent = "Current week: " + currentWeek;
  document.getElementById("week-config-info").textContent = "(ISO week " + currentWeek + " of " + currentYear + ")";

  // Filter to active rows that have a Final Shipment Week
  const active = allRows.filter((r) => !HIDDEN_STATUSES.includes(String(r["Status"] || "")));
  const buckets = {
    overdue: [],
    thisweek: [],
    nextweek: [],
    later: [],
    noweek: [],
  };

  for (const r of active) {
    const raw = r["Final Shipment Week"];
    const wk = parseWeek(raw);
    if (wk === null) { buckets.noweek.push(r); continue; }
    if (wk < currentWeek) buckets.overdue.push(r);
    else if (wk === currentWeek) buckets.thisweek.push(r);
    else if (wk === currentWeek + 1) buckets.nextweek.push(r);
    else if (wk <= currentWeek + WEEK_LOOK_AHEAD) buckets.later.push(r);
    else buckets.later.push(r);
  }

  // Sort each bucket — overdue: by week ascending (oldest first); others: by week
  buckets.overdue.sort((a, b) => parseWeek(a["Final Shipment Week"]) - parseWeek(b["Final Shipment Week"]));
  buckets.later.sort((a, b) => parseWeek(a["Final Shipment Week"]) - parseWeek(b["Final Shipment Week"]));

  let html = "";
  html += buildWeekColumn("Overdue", "overdue", buckets.overdue, "Final Shipment Week before week " + currentWeek);
  html += buildWeekColumn("This week (W" + currentWeek + ")", "thisweek", buckets.thisweek, "Final Shipment Week = " + currentWeek);
  html += buildWeekColumn("Next week (W" + (currentWeek + 1) + ")", "nextweek", buckets.nextweek, "Final Shipment Week = " + (currentWeek + 1));
  html += buildWeekColumn("Later (within " + WEEK_LOOK_AHEAD + " weeks)", "", buckets.later, "");
  html += buildWeekColumn("No shipment week set", "", buckets.noweek, "");

  root.innerHTML = html;

  // Wire clicks on week cards
  root.querySelectorAll(".week-card").forEach((c) => {
    c.addEventListener("click", () => {
      const sheetRow = parseInt(c.dataset.row, 10);
      const row = allRows.find((r) => r._row === sheetRow);
      if (row) openDrawer(row);
    });
  });
}

function buildWeekColumn(label, kindClass, rows, subLabel) {
  if (rows.length === 0) {
    return `
      <div class="week-col">
        <div class="week-col-header ${kindClass}">
          <strong>${escapeHtml(label)}</strong>
          <span class="week-col-count">0</span>
        </div>
        <div class="week-col-body"><div class="empty-state" style="padding:16px;">Nothing here.</div></div>
      </div>
    `;
  }
  const cards = rows.slice(0, 100).map((r) => {
    const mod = displayValue(r["JDE Module"]) || "(no module)";
    const hosp = displayValue(r["Hospital"]) || "—";
    const bu = displayValue(r["Business Unit"]) || "";
    return `
      <div class="week-card" data-row="${r._row}">
        <div class="week-card-top">${escapeHtml(mod)} <span style="font-weight:400;color:var(--text-muted);">· ${escapeHtml(bu)}</span></div>
        <div class="week-card-bot">${escapeHtml(hosp)}</div>
      </div>
    `;
  }).join("");
  const overflow = rows.length > 100 ? `<div style="text-align:center;padding:6px;font-size:11px;color:var(--text-muted);">+ ${rows.length - 100} more</div>` : "";
  return `
    <div class="week-col">
      <div class="week-col-header ${kindClass}">
        <strong>${escapeHtml(label)}</strong>
        <span class="week-col-count">${rows.length}</span>
      </div>
      <div class="week-col-body">${cards}${overflow}</div>
    </div>
  `;
}

// Parse "Final Shipment Week" values: "12", "15_2026", " 7 ", "2024", etc.
function parseWeek(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  const s = String(raw).trim();
  // Pattern "15_2026" — take the week part
  const m = s.match(/^(\d+)(?:_(\d{4}))?$/);
  if (m) {
    const wk = parseInt(m[1], 10);
    if (wk >= 1 && wk <= 53) return wk;
    // four-digit value means it's a year — ignore
    return null;
  }
  return null;
}

// ISO-8601 week number for a date
function isoWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

// =========================================================================
// HELPERS
// =========================================================================
// Excel stores dates as days-since-1900-01-01 (with a known leap-year bug).
// 45323 -> Jan 1 2024 etc. Returns a JS Date or null.
function excelSerialToDate(n) {
  if (typeof n !== "number" || !isFinite(n)) return null;
  if (n < 1 || n > 200000) return null; // sanity: rule out non-dates
  // Excel's epoch is 1899-12-30 (accounts for 1900 leap year bug)
  const ms = (n - 25569) * 86400 * 1000;
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return d;
}

const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// "45323" -> "Jan-24"   (used for ForecastMonth column)
function serialToMonthYear(n) {
  const d = excelSerialToDate(n);
  if (!d) return null;
  const yy = String(d.getUTCFullYear()).slice(-2);
  return MONTH_ABBR[d.getUTCMonth()] + "-" + yy;
}

// "45323" -> "01-Jan-24"  (used for general date-looking values)
function serialToDateStr(n) {
  const d = excelSerialToDate(n);
  if (!d) return null;
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const yy = String(d.getUTCFullYear()).slice(-2);
  return dd + "-" + MONTH_ABBR[d.getUTCMonth()] + "-" + yy;
}

// Columns whose values should be interpreted as month-year (e.g. "Jan-25")
const MONTH_YEAR_COLUMNS = new Set(["ForecastMonth"]);

// Columns whose values should be interpreted as date strings (e.g. "01-Jan-24")
const DATE_COLUMNS = new Set([
  "RequestedShipDate",
  "Approved to bring kit in incomplete (Date & initial)",
  "Date 9SE1 actually raised",
  "Shipment Date (Also look at auto column)",
  "Anchor Shipment Date (Auto)",
  "JDE Requested Date 06.06.24 (Auto)",
  "TECA Shipment Date ",
  "TECA BO FULFILMRNT DATE  ",
]);

// Normalize a raw cell value into what the UI should display.
// Numbers in date-y columns get formatted; everything else passes through.
function normalizeCellValue(colName, raw) {
  if (raw === null || raw === undefined || raw === "") return raw;
  if (typeof raw === "number") {
    if (MONTH_YEAR_COLUMNS.has(colName)) {
      const s = serialToMonthYear(raw);
      if (s) return s;
    }
    if (DATE_COLUMNS.has(colName)) {
      const s = serialToDateStr(raw);
      if (s) return s;
    }
  }
  return raw;
}

function colLetter(idx) {
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
function escapeAttr(s) { return escapeHtml(s); }
function capitalize(s) { return String(s).charAt(0).toUpperCase() + String(s).slice(1); }
function debounce(fn, ms) {
  let t; return function () {
    clearTimeout(t);
    const args = arguments, ctx = this;
    t = setTimeout(() => fn.apply(ctx, args), ms);
  };
}
function showEmpty(msg) {
  const e = document.getElementById("empty-state");
  e.classList.remove("hidden");
  e.innerHTML = "<p>" + escapeHtml(msg) + "</p>";
  document.getElementById("rows-container").classList.add("hidden");
  document.getElementById("pagination").classList.add("hidden");
}
function showToast(msg, kind) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast " + (kind || "");
  setTimeout(() => { t.className = "toast hidden"; }, 2600);
}
