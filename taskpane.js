/* ─────────────────────────────────────── */
/*  Kitting Flow v3                        */
/* ─────────────────────────────────────── */

// ── CONFIG ───────────────────────────────
const SHEET_NAME = "All Data";
const DATA_FIRST_COL = "A";
const DATA_LAST_COL  = "AX";
const HEADER_ROW     = 1;
const MAX_DATA_ROW   = 20000;
const CHUNK_ROWS     = 500;
const PAGE_SIZE      = 50;
const WEEK_LOOK_AHEAD = 4;

const CARD_FIELDS = {
  primary:   "JDE Module",
  secondary: "JDE Description",
  group:     "Business Unit",
  hospital:  "Hospital",
  status:    "Status",
};

const SEARCH_FIELDS = [
  "JDE Module","JDE Description","JDE Build number","Module_Build",
  "SAP Module","SAP Serial Number","SAP Description",
  "Hospital","HospitalContact",
  "Loanset","Set Reference","Anchor Unique Reference (Auto)",
  "9SE1 order number","PO","FID ",
  "YU /PR","Allocated Country",
  "Ordering Comment ","S&OP COMMENTS","Kit build comments",
  "BEK Ordering Issue (comments)",
];

// Hospitals that mean "not yet allocated to a real customer"
const UNALLOC_HOSPITALS = new Set([
  "Enter Ship To",
  "ORTHOKIT",
  "FORECAST ORDER",
  "Not in Hospital Master",
]);

const HIDDEN_STATUSES = ["EQUIPMENT SHIPPED","EQUIPMENT DISSOLVED","#REF!"];

function rowIsIssue(row) {
  const bek   = String(row["BEK Status (Auto)"] || "");
  const ship  = String(row["Can Module Ship?"] || "");
  const loc   = String(row["ISSUE LOCATION (Auto)"] || "");
  const stat  = String(row["Status"] || "");
  if (bek === "Incomplete" || bek === "Do not ship") return true;
  if (ship === "FALSE") return true;
  if (loc && loc !== "NO CURRENT KNOWN ISSUE OR AWAITING REFILL") return true;
  if (stat.includes("ISSUE")) return true;
  return false;
}

function rowIsAllocated(row) {
  const h = String(row["Hospital"] || "").trim();
  return h !== "" && !UNALLOC_HOSPITALS.has(h);
}

const FILTER_DROPDOWNS = [
  { id:"filter-bu",       column:"Business Unit",          limit:30 },
  { id:"filter-week-in",  column:"Week To Raise Inbound",  limit:60, sort:"weekish" },
  { id:"filter-forecast", column:"ForecastMonth",          limit:60, sort:"monthYear" },
  { id:"filter-fol",      column:"FOL Status (Auto)",      limit:15 },
];

const DETAIL_GROUPS = [
  { title:"Identification", fields:[
    {name:"JDE Module",editable:false},{name:"JDE Description",editable:false},
    {name:"JDE Build number",editable:false},{name:"Module_Build",editable:false},
    {name:"SAP Module",editable:false},{name:"SAP Serial Number",editable:false},
    {name:"SAP Description",editable:false},{name:"Business Unit",editable:false},
    {name:"Loanset",editable:false},
  ]},
  { title:"Status & allocation", fields:[
    {name:"Status",editable:true},{name:"Equipment status",editable:true},
    {name:"Can Module Ship?",editable:true},
    {name:"FOL Status (Auto)",editable:false},
    {name:"Completion Percentage",editable:false},
    {name:"SAP Equipment Monitor Status (AUTO)",editable:false},
    {name:"SAP Object Type (Auto)",editable:false},
    {name:"SAP Storage Location (Auto)",editable:false},
    {name:"ISSUE LOCATION (Auto)",editable:false},
  ]},
  { title:"Destination", fields:[
    {name:"Hospital",editable:true},{name:"Allocated Country",editable:true},
    {name:"HospitalContact",editable:true},{name:"Final Shipment Week",editable:true},
    {name:"Set Reference",editable:true},{name:"FID ",editable:true},
  ]},
  { title:"Ordering", fields:[
    {name:"Week To Raise Inbound",editable:true},{name:"9SE1 order number",editable:true},
    {name:"PO",editable:true},{name:"YU /PR",editable:true},
    {name:"Week Inbound Raised",editable:true},{name:"Requested By",editable:true},
    {name:"Raise 9SE2",editable:true},
  ]},
  { title:"Comments", fields:[
    {name:"Ordering Comment ",editable:true,multiline:true},
    {name:"BEK Ordering Issue (comments)",editable:true,multiline:true},
    {name:"S&OP COMMENTS",editable:true,multiline:true},
    {name:"Kit build comments",editable:true,multiline:true},
  ]},
];

const NEW_ROW_FIELDS = [
  {name:"JDE Module",required:true},{name:"JDE Description",required:true},
  {name:"Business Unit",required:true},{name:"Hospital",required:true},
  {name:"JDE Build number",required:false},{name:"SAP Module",required:false},
  {name:"SAP Serial Number",required:false},{name:"Loanset",required:false},
  {name:"Allocated Country",required:false},{name:"Status",required:false},
  {name:"Final Shipment Week",required:false},{name:"Requested By",required:false},
  {name:"9SE1 order number",required:false},{name:"PO",required:false},
  {name:"Ordering Comment ",required:false,multiline:true},
  {name:"S&OP COMMENTS",required:false,multiline:true},
];

// ── STATE ─────────────────────────────────
let allHeaders = [], headerIndex = {}, allRows = [], filteredRows = [];
let currentPage = 0, editingRow = null, editedValues = {}, isCreating = false;
let currentView = "browse";
let activeFilters = { alloc:"", reason:"", hideShipped:true, onlyIssues:false };

// ── BOOTSTRAP ─────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return showToast("Run in Excel.","error");
  bindUi();
  loadData();
});

function bindUi() {
  // Header
  document.getElementById("refresh-btn").addEventListener("click", loadData);
  document.getElementById("new-row-btn").addEventListener("click", openCreateDrawer);

  // Tabs
  document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => switchView(t.dataset.view)));

  // Filter toggle
  document.getElementById("filter-toggle").addEventListener("click", toggleFilterPanel);

  // Search
  document.getElementById("search").addEventListener("input", debounce(applyFilters, 200));

  // Allocation slicers
  document.querySelectorAll(".pill[data-slicer]").forEach(p => {
    p.addEventListener("click", () => onPillClick(p));
  });

  // Quick toggles (hide shipped / only issues)
  document.getElementById("toggle-hide-shipped").addEventListener("click", function() {
    activeFilters.hideShipped = !activeFilters.hideShipped;
    this.classList.toggle("active", activeFilters.hideShipped);
    applyFilters();
  });
  document.getElementById("toggle-only-issues").addEventListener("click", function() {
    activeFilters.onlyIssues = !activeFilters.onlyIssues;
    this.classList.toggle("active", activeFilters.onlyIssues);
    applyFilters();
  });

  // Dropdowns
  FILTER_DROPDOWNS.forEach(f => document.getElementById(f.id).addEventListener("change", applyFilters));

  // Clear
  document.getElementById("clear-filters").addEventListener("click", clearFilters);

  // Pagination
  document.getElementById("prev-page").addEventListener("click", () => { currentPage--; renderRows(); });
  document.getElementById("next-page").addEventListener("click", () => { currentPage++; renderRows(); });

  // Drawer
  document.getElementById("close-drawer").addEventListener("click", closeDrawer);
  document.getElementById("cancel-edit").addEventListener("click", closeDrawer);
  document.getElementById("save-edit").addEventListener("click", onSaveClick);

  // Week
  document.getElementById("week-refresh").addEventListener("click", renderWeekView);
}

// ── FILTER PANEL ──────────────────────────
let filterPanelOpen = false;
function toggleFilterPanel() {
  filterPanelOpen = !filterPanelOpen;
  const panel = document.getElementById("filter-panel");
  const btn   = document.getElementById("filter-toggle");
  panel.classList.toggle("open", filterPanelOpen);
  btn.classList.toggle("active", filterPanelOpen);
  btn.setAttribute("aria-expanded", filterPanelOpen);
}

function onPillClick(pill) {
  const slicer = pill.dataset.slicer;
  const val    = pill.dataset.val;
  // Deactivate siblings
  document.querySelectorAll(`.pill[data-slicer="${slicer}"]`).forEach(p => p.classList.remove("active"));
  pill.classList.add("active");
  activeFilters[slicer] = val;
  // Show / hide reason row
  if (slicer === "alloc") {
    document.getElementById("reason-row").style.display = val === "not-allocated" ? "" : "none";
    if (val !== "not-allocated") {
      activeFilters.reason = "";
      document.querySelectorAll('.pill[data-slicer="reason"]').forEach(p => p.classList.toggle("active", p.dataset.val === ""));
    }
  }
  updateFilterBadge();
  applyFilters();
}

function updateFilterBadge() {
  let count = 0;
  if (activeFilters.alloc) count++;
  if (activeFilters.reason) count++;
  if (!activeFilters.hideShipped) count++;   // non-default = active
  if (activeFilters.onlyIssues) count++;
  FILTER_DROPDOWNS.forEach(f => { if (document.getElementById(f.id).value) count++; });
  const badge = document.getElementById("filter-badge");
  badge.textContent = count;
  badge.classList.toggle("hidden", count === 0);
}

function clearFilters() {
  document.getElementById("search").value = "";
  activeFilters = { alloc:"", reason:"", hideShipped:true, onlyIssues:false };
  document.querySelectorAll(".pill[data-slicer]").forEach(p => {
    p.classList.toggle("active", p.dataset.val === "");
  });
  document.getElementById("toggle-hide-shipped").classList.add("active");
  document.getElementById("toggle-only-issues").classList.remove("active");
  document.getElementById("reason-row").style.display = "none";
  FILTER_DROPDOWNS.forEach(f => { document.getElementById(f.id).value = ""; });
  updateFilterBadge();
  applyFilters();
}

// ── FILTERS ───────────────────────────────
function applyFilters() {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const dropVals = {};
  FILTER_DROPDOWNS.forEach(f => {
    const v = document.getElementById(f.id).value;
    if (v) dropVals[f.column] = v;
  });

  filteredRows = allRows.filter(row => {
    if (activeFilters.hideShipped && HIDDEN_STATUSES.includes(String(row["Status"] || ""))) return false;
    if (activeFilters.onlyIssues && !rowIsIssue(row)) return false;
    if (activeFilters.alloc === "allocated"     && !rowIsAllocated(row)) return false;
    if (activeFilters.alloc === "not-allocated" &&  rowIsAllocated(row)) return false;
    if (activeFilters.reason) {
      if (String(row["Hospital"] || "").trim() !== activeFilters.reason) return false;
    }
    for (const col of Object.keys(dropVals)) {
      if (String(row[col] ?? "") !== dropVals[col]) return false;
    }
    if (q) {
      const haystack = SEARCH_FIELDS.map(f => row[f] == null ? "" : String(row[f]).toLowerCase()).join("|");
      if (!haystack.includes(q)) return false;
    }
    return true;
  });

  currentPage = 0;
  updateFilterBadge();
  renderRows();
}

function populateFilterDropdowns() {
  FILTER_DROPDOWNS.forEach(f => {
    const el = document.getElementById(f.id);
    const first = el.querySelector("option");
    el.innerHTML = ""; el.appendChild(first);
    const seen = new Set();
    for (const row of allRows) {
      const v = row[f.column];
      if (v == null || v === "" || String(v) === "#REF!") continue;
      seen.add(String(v));
    }
    const sorter = pickSorter(f.sort);
    [...seen].sort(sorter).slice(0, f.limit).forEach(v => {
      const o = document.createElement("option");
      o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}

function pickSorter(kind) {
  if (kind === "weekish") return (a,b) => {
    const wa = parseWeek(a), wb = parseWeek(b);
    if (wa!=null&&wb!=null) return wa-wb;
    if (wa!=null) return -1; if (wb!=null) return 1;
    return a.localeCompare(b);
  };
  if (kind === "monthYear") return (a,b) => {
    const da = parseMonthYear(a), db = parseMonthYear(b);
    if (da&&db) return da-db;
    if (da) return -1; if (db) return 1;
    return a.localeCompare(b);
  };
  return (a,b) => a.localeCompare(b);
}

// ── RENDER ROWS ───────────────────────────
function renderRows() {
  const container = document.getElementById("rows-container");
  const empty     = document.getElementById("empty-state");
  const pag       = document.getElementById("pagination");

  document.getElementById("result-count").textContent = filteredRows.length.toLocaleString() + " rows";

  if (!filteredRows.length) {
    container.classList.add("hidden"); pag.classList.add("hidden");
    empty.classList.remove("hidden");
    empty.innerHTML = "<p>No matching rows. Clear filters or try a different search.</p>";
    return;
  }
  empty.classList.add("hidden"); container.classList.remove("hidden"); pag.classList.remove("hidden");

  const total = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  currentPage = Math.max(0, Math.min(currentPage, total-1));
  const slice = filteredRows.slice(currentPage * PAGE_SIZE, (currentPage+1) * PAGE_SIZE);

  container.innerHTML = "";
  slice.forEach(row => container.appendChild(buildRowCard(row)));

  document.getElementById("page-info").textContent = (currentPage+1) + " / " + total;
  document.getElementById("prev-page").disabled = currentPage === 0;
  document.getElementById("next-page").disabled = currentPage >= total-1;
}

function buildRowCard(row) {
  const card = document.createElement("div");
  card.className = "row-card"; card.tabIndex = 0;

  // top row
  const top = document.createElement("div"); top.className = "row-card-top";
  const mod = document.createElement("div"); mod.className = "row-module";
  mod.textContent = displayValue(row[CARD_FIELDS.primary]) || "(no module)";
  const bu = document.createElement("div"); bu.className = "row-bu";
  bu.textContent = displayValue(row[CARD_FIELDS.group]);
  top.append(mod, bu);

  // mid row
  const mid = document.createElement("div"); mid.className = "row-card-mid";
  const hosp    = displayValue(row["Hospital"]);
  const loanset = displayValue(row["Loanset"]);
  const desc    = displayValue(row[CARD_FIELDS.secondary]);
  const allocated = rowIsAllocated(row);
  let midHtml = desc ? escapeHtml(desc) : "";
  if (hosp) midHtml += (midHtml?" · ":"") + `<span class='row-hospital'>${escapeHtml(hosp)}</span>`;
  if (loanset) midHtml += (midHtml?" · ":"") + "LS " + escapeHtml(loanset);
  // allocation badge inline
  const allocBadge = allocated
    ? `<span class='row-alloc-badge alloc-yes'>Allocated</span>`
    : `<span class='row-alloc-badge alloc-no'>${escapeHtml(getAllocReason(row))}</span>`;
  midHtml += " " + allocBadge;
  mid.innerHTML = midHtml;

  // bottom chips
  const bot = document.createElement("div"); bot.className = "row-card-bot";
  const stat = displayValue(row["Status"]);
  const fol  = displayValue(row["FOL Status (Auto)"]);
  if (stat) bot.appendChild(buildChip(stat, classifyStatus(stat)));
  if (fol && fol !== "Please enter FOL ID" && fol !== "#REF!") bot.appendChild(buildChip(fol, "info"));

  card.append(top, mid, bot);
  card.addEventListener("click", () => openDrawer(row));
  card.addEventListener("keydown", e => { if (e.key==="Enter") openDrawer(row); });
  return card;
}

function getAllocReason(row) {
  const h = String(row["Hospital"] || "").trim();
  if (h === "FORECAST ORDER") return "Forecast";
  if (h === "ORTHOKIT") return "OrthoKit";
  if (h === "Enter Ship To") return "Pending ship-to";
  if (h === "Not in Hospital Master") return "Not in master";
  if (h === "") return "No hospital";
  return "Not allocated";
}

function buildChip(text, kind) {
  const c = document.createElement("span");
  c.className = "chip" + (kind ? " chip-"+kind : "");
  c.textContent = text; return c;
}
function classifyStatus(v) {
  const s = v.toLowerCase();
  if (s==="equipment shipped"||s==="equipment built") return "good";
  if (s.includes("issue")||s.includes("fail")) return "bad";
  if (s.includes("await")||s.includes("orthokit")||s.includes("topup")) return "warn";
  if (s.includes("dissolved")) return "info";
  return "";
}

// ── LOAD DATA ─────────────────────────────
async function loadData() {
  showEmpty("Initialising…");
  allRows = []; allHeaders = []; headerIndex = {};

  try {
    showEmpty("Reading headers…");
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const r = sheet.getRange(DATA_FIRST_COL+HEADER_ROW+":"+DATA_LAST_COL+HEADER_ROW);
      r.load("values"); await ctx.sync();
      allHeaders = r.values[0].map(h => h==null?"":String(h));
      allHeaders.forEach((h,i) => { headerIndex[h]=i; });
    });

    const first = HEADER_ROW+1;
    let offset = 0, emptyStreak = 0;

    while (offset < MAX_DATA_ROW-first) {
      const size = Math.min(CHUNK_ROWS, MAX_DATA_ROW-first-offset);
      const a = first+offset, b = a+size-1;
      let nonEmpty = 0;

      await Excel.run(async ctx => {
        const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
        const r = sheet.getRange(DATA_FIRST_COL+a+":"+DATA_LAST_COL+b);
        r.load("values"); await ctx.sync();
        for (let i=0; i<r.values.length; i++) {
          const row = r.values[i];
          if (!row.some(v => v!==""&&v!==null&&v!==undefined)) continue;
          nonEmpty++;
          const obj = { _row: a+i };
          for (let c=0; c<allHeaders.length; c++) obj[allHeaders[c]] = normalizeCellValue(allHeaders[c], row[c]);
          allRows.push(obj);
        }
      });

      offset += size;
      if (nonEmpty===0) { if (++emptyStreak>=3) break; } else emptyStreak=0;
      const pct = Math.min(99, Math.round(allRows.length/16000*100));
      showEmpty("Loading… ~"+pct+"% ("+allRows.length.toLocaleString()+" rows)");
    }

    populateFilterDropdowns();
    applyFilters();
    if (currentView==="dashboard") renderDashboard();
    if (currentView==="week") renderWeekView();
    showToast("Loaded "+allRows.length.toLocaleString()+" rows","success");
  } catch(err) {
    console.error(err);
    showEmpty("Load failed: "+(err.message||err));
    showToast("Load failed: "+(err.message||err),"error");
  }
}

// ── VIEWS ─────────────────────────────────
function switchView(name) {
  currentView = name;
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.view===name));
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("active", v.id==="view-"+name));
  if (name==="dashboard") renderDashboard();
  if (name==="week") renderWeekView();
}

// ── DASHBOARD ─────────────────────────────
function renderDashboard() {
  const root = document.getElementById("dashboard-content");
  if (!allRows.length) { root.innerHTML="<p class='empty-state'>Load data first.</p>"; return; }

  const active = allRows.filter(r => !HIDDEN_STATUSES.includes(String(r["Status"]||"")));
  const allocated   = active.filter(rowIsAllocated);
  const unallocated = active.filter(r => !rowIsAllocated(r));
  const issueCount  = active.filter(rowIsIssue).length;
  const incompleteCount  = active.filter(r => String(r["BEK Status (Auto)"])==="Incomplete").length;
  const doNotShipCount   = active.filter(r => String(r["BEK Status (Auto)"])==="Do not ship").length;

  const bekCounts      = countBy(active, "BEK Status (Auto)", v => v&&v!=="NOT IN BEK");
  const folCounts      = countBy(active, "FOL Status (Auto)", v => v&&v!=="#REF!"&&v!=="Please enter FOL ID");
  const weekInCounts   = countBy(active, "Week To Raise Inbound", v => v&&v!=="#REF!");
  const forecastCounts = countBy(active, "ForecastMonth", v => v&&v!=="#REF!");
  const hospitalCounts = countBy(allocated, "Hospital",
    v => v&&!["Enter Ship To","ORTHOKIT","FORECAST ORDER","Not in Hospital Master",""].includes(v));

  // Not-allocated breakdown by reason
  const reasonCounts = countBy(unallocated, "Hospital", v => v!=="");

  let html = "";
  html += dashSection("Overview", `<div class="dash-grid">
    ${tile("Active modules", active.length.toLocaleString(), "non-shipped rows", "info")}
    ${tile("Allocated", allocated.length.toLocaleString(), "real hospital assigned", "good")}
    ${tile("Not allocated", unallocated.length.toLocaleString(), "no real hospital yet", unallocated.length>0?"warn":"")}
    ${tile("With issues", issueCount.toLocaleString(), "BEK / KB / ship flag", issueCount>0?"bad":"good")}
  </div>`);

  html += dashSection("Allocation breakdown", `<div class="dash-grid">
    ${tile("Forecast orders", (reasonCounts["FORECAST ORDER"]||0).toLocaleString(), "planned future builds", "info")}
    ${tile("OrthoKit", (reasonCounts["ORTHOKIT"]||0).toLocaleString(), "sitting in OrthoKit", "warn")}
    ${tile("Pending ship-to", (reasonCounts["Enter Ship To"]||0).toLocaleString(), "ship-to not yet set", "warn")}
    ${tile("Not in master", (reasonCounts["Not in Hospital Master"]||0).toLocaleString(), "hospital not in system", "bad")}
  </div>`);

  html += dashSection("BEK status", listTile(bekCounts, 6, true));
  html += dashSection("FOL status", listTile(folCounts, 6, true));
  html += dashSection("Upcoming inbound weeks", listTile(sortCountsForWeek(weekInCounts), 10, false));
  html += dashSection("Upcoming forecast months", listTile(sortCountsByMonthYear(forecastCounts), 10, false));
  html += dashSection("Top hospitals (allocated)", listTile(hospitalCounts, 10, true));

  root.innerHTML = html;

  root.querySelectorAll("[data-dash-val]").forEach(el => {
    el.addEventListener("click", () => dispatchDashboardAction(el.dataset.dashVal));
  });
}

function dashSection(title, inner) {
  return `<div class="dash-section"><div class="dash-section-title">${escapeHtml(title)}</div>${inner}</div>`;
}
function tile(label, value, sub, kind) {
  const cls = kind ? " "+kind : "";
  return `<div class="dash-tile${cls}">
    <div class="dash-tile-label">${escapeHtml(label)}</div>
    <div class="dash-tile-value">${escapeHtml(value)}</div>
    <div class="dash-tile-sub">${escapeHtml(sub)}</div>
  </div>`;
}
function listTile(counts, max, sortByCount) {
  let entries = Array.isArray(counts) ? counts : Object.entries(counts);
  if (sortByCount) entries = entries.sort((a,b)=>b[1]-a[1]);
  entries = entries.slice(0, max);
  if (!entries.length) return "<div class='dash-list-tile'><div class='dash-list-row'>No data</div></div>";
  return `<div class="dash-list-tile">${entries.map(([label,n])=>`
    <div class="dash-list-row" data-dash-val="${escapeAttr(label)}">
      <span class="dash-list-name" title="${escapeAttr(label)}">${escapeHtml(label)}</span>
      <span class="dash-list-count">${n.toLocaleString()}</span>
    </div>`).join("")}</div>`;
}
function sortCountsForWeek(counts) {
  return Object.entries(counts).sort((a,b)=>{
    const wa=parseWeek(a[0]),wb=parseWeek(b[0]);
    if(wa!=null&&wb!=null)return wa-wb; if(wa!=null)return -1; if(wb!=null)return 1;
    return a[0].localeCompare(b[0]);
  });
}
function sortCountsByMonthYear(counts) {
  return Object.entries(counts).sort((a,b)=>{
    const da=parseMonthYear(a[0]),db=parseMonthYear(b[0]);
    if(da&&db)return da-db; if(da)return -1; if(db)return 1;
    return a[0].localeCompare(b[0]);
  });
}
function dispatchDashboardAction(value) {
  const candidates = ["filter-fol","filter-week-in","filter-forecast","filter-bu"];
  for (const id of candidates) {
    const sel = document.getElementById(id);
    for (const opt of sel.options) {
      if (opt.value===value) {
        switchView("browse"); clearFilters();
        sel.value = value; applyFilters(); return;
      }
    }
  }
  switchView("browse"); clearFilters();
  document.getElementById("search").value = value;
  applyFilters();
}

// ── WEEK VIEW ─────────────────────────────
function renderWeekView() {
  const root = document.getElementById("week-columns");
  if (!allRows.length) { root.innerHTML="<p class='empty-state'>Load data first.</p>"; return; }

  const today = new Date();
  const cw = isoWeek(today);
  document.getElementById("week-current-label").textContent = "Current week: W"+cw;
  document.getElementById("week-config-info").textContent = "(ISO week "+cw+" of "+today.getFullYear()+")";

  const active = allRows.filter(r => !HIDDEN_STATUSES.includes(String(r["Status"]||"")));
  const buckets = { overdue:[], thisweek:[], nextweek:[], later:[], noweek:[] };

  for (const r of active) {
    const wk = parseWeek(r["Final Shipment Week"]);
    if (wk===null) { buckets.noweek.push(r); continue; }
    if (wk<cw) buckets.overdue.push(r);
    else if (wk===cw) buckets.thisweek.push(r);
    else if (wk===cw+1) buckets.nextweek.push(r);
    else buckets.later.push(r);
  }
  buckets.overdue.sort((a,b)=>parseWeek(a["Final Shipment Week"])-parseWeek(b["Final Shipment Week"]));
  buckets.later.sort((a,b)=>parseWeek(a["Final Shipment Week"])-parseWeek(b["Final Shipment Week"]));

  root.innerHTML = [
    weekCol("Overdue","overdue",buckets.overdue),
    weekCol("This week (W"+cw+")","thisweek",buckets.thisweek),
    weekCol("Next week (W"+(cw+1)+")","nextweek",buckets.nextweek),
    weekCol("Later","",buckets.later),
    weekCol("No week set","",buckets.noweek),
  ].join("");

  root.querySelectorAll(".week-card").forEach(c => {
    c.addEventListener("click", () => {
      const row = allRows.find(r => r._row===parseInt(c.dataset.row,10));
      if (row) openDrawer(row);
    });
  });
}
function weekCol(label, cls, rows) {
  const cards = rows.slice(0,100).map(r=>`
    <div class="week-card" data-row="${r._row}">
      <div class="week-card-top">${escapeHtml(displayValue(r["JDE Module"])||"(no module)")} <span style="font-weight:400;color:var(--muted)">· ${escapeHtml(displayValue(r["Business Unit"])||"")}</span></div>
      <div class="week-card-bot">${escapeHtml(displayValue(r["Hospital"])||"—")}</div>
    </div>`).join("");
  const more = rows.length>100 ? `<div style="text-align:center;padding:6px;font-size:11px;color:var(--muted)">+${rows.length-100} more</div>` : "";
  return `<div class="week-col">
    <div class="week-col-header ${cls}"><strong>${escapeHtml(label)}</strong><span class="week-col-count">${rows.length}</span></div>
    <div class="week-col-body">${rows.length ? cards+more : "<div class='empty-state' style='padding:12px'>Nothing here.</div>"}</div>
  </div>`;
}

// ── DRAWER ────────────────────────────────
function openDrawer(row) {
  isCreating=false; editingRow=row; editedValues={};
  document.getElementById("save-edit").textContent="Save changes";
  document.getElementById("detail-title").textContent=displayValue(row[CARD_FIELDS.primary])||"Module details";
  const el=document.getElementById("detail-fields"); el.innerHTML="";
  DETAIL_GROUPS.forEach(g=>{
    const div=document.createElement("div"); div.className="field-group";
    const t=document.createElement("div"); t.className="field-group-title"; t.textContent=g.title;
    div.appendChild(t);
    g.fields.forEach(f=>{ if(f.name in headerIndex) div.appendChild(buildField(f, row[f.name])); });
    el.appendChild(div);
  });
  document.getElementById("detail-drawer").classList.remove("hidden");
}
function buildField(def, value) {
  const wrap=document.createElement("div"); wrap.className="field"+(def.editable?"":" readonly");
  const lbl=document.createElement("label"); lbl.className="field-label"; lbl.textContent=def.name.trim();
  wrap.appendChild(lbl);
  if (!def.editable) {
    const v=document.createElement("div"); v.className="field-value"; v.textContent=displayValue(value);
    wrap.appendChild(v); return wrap;
  }
  const inp = def.multiline ? document.createElement("textarea") : document.createElement("input");
  inp.className=def.multiline?"field-textarea":"field-input";
  if (!def.multiline) inp.type="text";
  if (def.multiline) inp.rows=3;
  inp.value=value==null?"":String(value);
  inp.addEventListener("input",()=>{ editedValues[def.name]=inp.value; });
  wrap.appendChild(inp); return wrap;
}
function openCreateDrawer() {
  if (!allHeaders.length) { showToast("Load data first.","error"); return; }
  isCreating=true; editingRow=null; editedValues={};
  document.getElementById("save-edit").textContent="Create row";
  document.getElementById("detail-title").textContent="New row";
  const el=document.getElementById("detail-fields"); el.innerHTML="";
  const g=document.createElement("div"); g.className="field-group";
  const t=document.createElement("div"); t.className="field-group-title"; t.textContent="New row — required fields marked *";
  g.appendChild(t);
  NEW_ROW_FIELDS.forEach(f=>{
    if(!(f.name in headerIndex)) return;
    const wrap=document.createElement("div"); wrap.className="field";
    const lbl=document.createElement("label"); lbl.className="field-label"; lbl.textContent=f.name.trim();
    if(f.required){const s=document.createElement("span");s.className="field-required-marker";s.textContent="*";lbl.appendChild(s);}
    wrap.appendChild(lbl);
    const inp=f.multiline?document.createElement("textarea"):document.createElement("input");
    inp.className=f.multiline?"field-textarea":"field-input";
    if(!f.multiline)inp.type="text"; if(f.multiline)inp.rows=3;
    inp.dataset.fieldName=f.name; if(f.required)inp.dataset.required="true";
    inp.addEventListener("input",()=>{editedValues[f.name]=inp.value;inp.classList.remove("field-error");});
    wrap.appendChild(inp); g.appendChild(wrap);
  });
  el.appendChild(g);
  document.getElementById("detail-drawer").classList.remove("hidden");
}
function closeDrawer() {
  document.getElementById("detail-drawer").classList.add("hidden");
  document.getElementById("validation-msg").classList.add("hidden");
  editingRow=null; editedValues={}; isCreating=false;
}
function onSaveClick() { isCreating ? createNewRow() : saveEdits(); }
async function saveEdits() {
  if (!editingRow) return;
  const changed=Object.keys(editedValues);
  if (!changed.length) { closeDrawer(); return; }
  const btn=document.getElementById("save-edit");
  btn.disabled=true; btn.textContent="Saving…";
  try {
    await Excel.run(async ctx=>{
      const sheet=ctx.workbook.worksheets.getItem(SHEET_NAME);
      changed.forEach(col=>{
        const idx=headerIndex[col]; if(idx===undefined)return;
        sheet.getRange(colLetter(idx)+editingRow._row).values=[[editedValues[col]]];
      });
      await ctx.sync();
    });
    changed.forEach(col=>{editingRow[col]=editedValues[col];});
    showToast("Saved "+changed.length+" change"+(changed.length===1?"":"s"),"success");
    renderRows(); closeDrawer();
  } catch(err){showToast("Save failed: "+(err.message||err),"error");}
  finally{btn.disabled=false;btn.textContent="Save changes";}
}
async function createNewRow() {
  const missing=[];
  document.querySelectorAll("#detail-fields [data-field-name]").forEach(inp=>{
    inp.classList.remove("field-error");
    if(inp.dataset.required==="true"&&!inp.value.trim()){missing.push(inp.dataset.fieldName.trim());inp.classList.add("field-error");}
  });
  const vm=document.getElementById("validation-msg");
  if(missing.length){vm.textContent="Please fill in: "+missing.join(", ");vm.classList.remove("hidden");return;}
  vm.classList.add("hidden");
  const btn=document.getElementById("save-edit");
  btn.disabled=true; btn.textContent="Creating…";
  try {
    let nr=HEADER_ROW+1;
    for(const r of allRows) if(r._row>=nr) nr=r._row+1;
    await Excel.run(async ctx=>{
      const sheet=ctx.workbook.worksheets.getItem(SHEET_NAME);
      const vals=new Array(allHeaders.length).fill("");
      Object.keys(editedValues).forEach(col=>{const i=headerIndex[col];if(i!==undefined)vals[i]=editedValues[col];});
      sheet.getRange("A"+nr+":"+colLetter(allHeaders.length-1)+nr).values=[vals];
      sheet.getRange("A"+nr).select();
      await ctx.sync();
    });
    const obj={_row:nr};
    allHeaders.forEach(h=>{obj[h]=editedValues[h]||"";});
    allRows.push(obj);
    showToast("Row created (row "+nr+")","success");
    populateFilterDropdowns(); applyFilters(); closeDrawer();
  } catch(err){showToast("Create failed: "+(err.message||err),"error");}
  finally{btn.disabled=false;btn.textContent="Create row";}
}

// ── DATE / VALUE HELPERS ──────────────────
const MONTH_ABBR=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_YEAR_COLS=new Set(["ForecastMonth"]);
const DATE_COLS=new Set(["RequestedShipDate","Approved to bring kit in incomplete (Date & initial)",
  "Date 9SE1 actually raised","Shipment Date (Also look at auto column)",
  "Anchor Shipment Date (Auto)","JDE Requested Date 06.06.24 (Auto)",
  "TECA Shipment Date ","TECA BO FULFILMRNT DATE  "]);

function excelSerialToDate(n){
  if(typeof n!=="number"||!isFinite(n)||n<1||n>200000)return null;
  const d=new Date((n-25569)*86400*1000);
  return isNaN(d.getTime())?null:d;
}
function serialToMonthYear(n){
  const d=excelSerialToDate(n); if(!d)return null;
  return MONTH_ABBR[d.getUTCMonth()]+"-"+String(d.getUTCFullYear()).slice(-2);
}
function serialToDateStr(n){
  const d=excelSerialToDate(n); if(!d)return null;
  return String(d.getUTCDate()).padStart(2,"0")+"-"+MONTH_ABBR[d.getUTCMonth()]+"-"+String(d.getUTCFullYear()).slice(-2);
}
function normalizeCellValue(col,raw){
  if(raw===null||raw===undefined||raw==="")return raw;
  if(typeof raw==="number"){
    if(MONTH_YEAR_COLS.has(col)){const s=serialToMonthYear(raw);if(s)return s;}
    if(DATE_COLS.has(col)){const s=serialToDateStr(raw);if(s)return s;}
  }
  return raw;
}
function parseMonthYear(s){
  const m=String(s).trim().match(/^([A-Za-z]{3})-(\d{2,4})$/);
  if(!m)return null;
  const mo={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}[m[1].toLowerCase()];
  if(mo===undefined)return null;
  let y=parseInt(m[2],10); if(y<100)y+=2000;
  return new Date(y,mo,1).getTime();
}
function parseWeek(raw){
  if(raw===null||raw===undefined||raw==="")return null;
  const m=String(raw).trim().match(/^(\d+)(?:_\d{4})?$/);
  if(!m)return null;
  const wk=parseInt(m[1],10);
  return (wk>=1&&wk<=53)?wk:null;
}
function isoWeek(d){
  const date=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  const day=date.getUTCDay()||7;
  date.setUTCDate(date.getUTCDate()+4-day);
  const y=new Date(Date.UTC(date.getUTCFullYear(),0,1));
  return Math.ceil((((date-y)/86400000)+1)/7);
}
function countBy(rows,col,filterFn){
  const c={};
  for(const r of rows){
    const v=r[col]; if(v==null||v==="")continue;
    const s=String(v); if(filterFn&&!filterFn(s))continue;
    c[s]=(c[s]||0)+1;
  }
  return c;
}
function colLetter(idx){
  let s="",n=idx;
  while(n>=0){s=String.fromCharCode((n%26)+65)+s;n=Math.floor(n/26)-1;}
  return s;
}
function displayValue(v){
  if(v===undefined||v===null)return "";
  if(typeof v==="number"&&!Number.isInteger(v))return Math.abs(v)<1?v.toFixed(3):v.toFixed(2);
  return String(v);
}
function escapeHtml(s){
  return String(s).replace(/[&<>"']/g,c=>({'&':"&amp;",'<':"&lt;",'>':"&gt;",'"':"&quot;","'":'&#39;'}[c]));
}
function escapeAttr(s){return escapeHtml(s);}
function debounce(fn,ms){let t;return function(){clearTimeout(t);t=setTimeout(()=>fn.apply(this,arguments),ms);};}
function showEmpty(msg){
  document.getElementById("empty-state").classList.remove("hidden");
  document.getElementById("empty-state").innerHTML="<p>"+escapeHtml(msg)+"</p>";
  document.getElementById("rows-container").classList.add("hidden");
  document.getElementById("pagination").classList.add("hidden");
}
function showToast(msg,kind){
  const t=document.getElementById("toast");
  t.textContent=msg; t.className="toast "+(kind||"");
  setTimeout(()=>{t.className="toast hidden";},2600);
}
