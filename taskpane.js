/* ─────────────────────────────────────── */
/*  Kitting Flow v4 — built 2026-06-16 15:17  */
/* ─────────────────────────────────────── */

const SHEET_NAME = "All Data";
const DATA_FIRST_COL = "A", DATA_LAST_COL = "AX";
const HEADER_ROW = 1, MAX_DATA_ROW = 20000, CHUNK_ROWS = 500;
const PAGE_SIZE = 50;

const CARD_FIELDS = {
  primary:"JDE Module", secondary:"JDE Description",
  group:"Business Unit", hospital:"Hospital", status:"Status",
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

const UNALLOC_HOSPITALS = new Set([
  "Enter Ship To","ORTHOKIT","FORECAST ORDER","Not in Hospital Master",
]);
const HIDDEN_STATUSES = ["EQUIPMENT SHIPPED","EQUIPMENT DISSOLVED","#REF!"];

function rowIsIssue(row) {
  const bek=String(row["BEK Status (Auto)"]||""), ship=String(row["Can Module Ship?"]||"");
  const loc=String(row["ISSUE LOCATION (Auto)"]||""), stat=String(row["Status"]||"");
  if (bek==="Incomplete"||bek==="Do not ship") return true;
  if (ship==="FALSE") return true;
  if (loc&&loc!=="NO CURRENT KNOWN ISSUE OR AWAITING REFILL") return true;
  if (stat.includes("ISSUE")) return true;
  return false;
}
function rowIsAllocated(row) {
  const h = String(row["Hospital"]||"").trim();
  return h!==""&&!UNALLOC_HOSPITALS.has(h);
}

const FILTER_DROPDOWNS = [
  { id:"filter-bu",      column:"Business Unit",         limit:30 },
  { id:"filter-status",  column:"Status",                limit:15 },
  { id:"filter-week-in", column:"Week To Raise Inbound", limit:60, sort:"weekish" },
];

const DETAIL_GROUPS = [
  { title:"Identification", fields:[
    {name:"JDE Module"},{name:"JDE Description"},{name:"JDE Build number"},
    {name:"Module_Build"},{name:"SAP Module"},{name:"SAP Serial Number"},
    {name:"SAP Description"},{name:"Business Unit"},{name:"Loanset"},
  ]},
  { title:"Status & allocation", fields:[
    {name:"Status",editable:true},{name:"Equipment status",editable:true},
    {name:"Can Module Ship?",editable:true},{name:"FOL Status (Auto)"},
    {name:"Completion Percentage"},{name:"SAP Equipment Monitor Status (AUTO)"},
    {name:"SAP Object Type (Auto)"},{name:"SAP Storage Location (Auto)"},
    {name:"ISSUE LOCATION (Auto)"},
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
  {name:"JDE Build number"},{name:"SAP Module"},{name:"SAP Serial Number"},
  {name:"Loanset"},{name:"Allocated Country"},{name:"Status"},
  {name:"Final Shipment Week"},{name:"Requested By"},
  {name:"9SE1 order number"},{name:"PO"},
  {name:"Ordering Comment ",multiline:true},{name:"S&OP COMMENTS",multiline:true},
];

// ── STATE ──
let allHeaders=[], headerIndex={}, allRows=[], filteredRows=[];
let currentPage=0, editingRow=null, editedValues={}, isCreating=false;
let currentView="browse";
let activeFilters={ alloc:"", reason:"", hideShipped:true, onlyIssues:false, sapModule:"" };

// Track which dashboard sections / week columns are collapsed (persists during session)
const collapseState = { dash:{}, week:{} };

// ── BOOTSTRAP ──
Office.onReady((info)=>{
  if(info.host!==Office.HostType.Excel) return showToast("Run in Excel.","error");
  bindUi(); loadData();
});

function bindUi() {
  document.getElementById("refresh-btn").addEventListener("click", loadData);
  document.getElementById("new-row-btn").addEventListener("click", openCreateDrawer);
  document.getElementById("expand-btn").addEventListener("click", openFullscreen);
  document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => switchView(t.dataset.view)));
  document.getElementById("filter-toggle").addEventListener("click", toggleFilterPanel);
  document.getElementById("search").addEventListener("input", debounce(applyFilters, 200));
  document.querySelectorAll(".pill[data-slicer]").forEach(p => p.addEventListener("click", () => onPillClick(p)));
  document.getElementById("toggle-hide-shipped").addEventListener("click", function(){
    activeFilters.hideShipped = !activeFilters.hideShipped;
    this.classList.toggle("active", activeFilters.hideShipped);
    applyFilters();
  });
  document.getElementById("toggle-only-issues").addEventListener("click", function(){
    activeFilters.onlyIssues = !activeFilters.onlyIssues;
    this.classList.toggle("active", activeFilters.onlyIssues);
    applyFilters();
  });
  FILTER_DROPDOWNS.forEach(f => document.getElementById(f.id).addEventListener("change", applyFilters));
  document.getElementById("filter-sap").addEventListener("input", debounce(() => {
    activeFilters.sapModule = document.getElementById("filter-sap").value.trim();
    updateFilterBadge();
    applyFilters();
  }, 250));
  document.getElementById("clear-filters").addEventListener("click", clearFilters);
  document.getElementById("prev-page").addEventListener("click", () => { currentPage--; renderRows(); });
  document.getElementById("next-page").addEventListener("click", () => { currentPage++; renderRows(); });
  document.getElementById("close-drawer").addEventListener("click", closeDrawer);
  document.getElementById("cancel-edit").addEventListener("click", closeDrawer);
  document.getElementById("save-edit").addEventListener("click", onSaveClick);
}

// ── FILTER PANEL ──
let filterPanelOpen=false;
function toggleFilterPanel() {
  filterPanelOpen = !filterPanelOpen;
  document.getElementById("filter-panel").classList.toggle("open", filterPanelOpen);
  document.getElementById("filter-toggle").classList.toggle("active", filterPanelOpen);
  document.getElementById("filter-toggle").setAttribute("aria-expanded", filterPanelOpen);
}
function onPillClick(pill) {
  const slicer=pill.dataset.slicer, val=pill.dataset.val;
  document.querySelectorAll(`.pill[data-slicer="${slicer}"]`).forEach(p => p.classList.remove("active"));
  pill.classList.add("active");
  activeFilters[slicer] = val;
  if (slicer==="alloc") {
    document.getElementById("reason-row").style.display = val==="not-allocated" ? "" : "none";
    if (val!=="not-allocated") {
      activeFilters.reason = "";
      document.querySelectorAll('.pill[data-slicer="reason"]').forEach(p => p.classList.toggle("active", p.dataset.val===""));
    }
  }
  updateFilterBadge();
  applyFilters();
}
function updateFilterBadge() {
  let count=0;
  if (activeFilters.alloc) count++;
  if (activeFilters.reason) count++;
  if (!activeFilters.hideShipped) count++;
  if (activeFilters.onlyIssues) count++;
  if (activeFilters.sapModule) count++;
  FILTER_DROPDOWNS.forEach(f => { if (document.getElementById(f.id).value) count++; });  const badge = document.getElementById("filter-badge");
  badge.textContent = count;
  badge.classList.toggle("hidden", count===0);
}
function clearFilters() {
  document.getElementById("search").value = "";
  document.getElementById("filter-sap").value = "";
  activeFilters = { alloc:"", reason:"", hideShipped:true, onlyIssues:false, sapModule:"" };
  document.querySelectorAll(".pill[data-slicer]").forEach(p => p.classList.toggle("active", p.dataset.val===""));
  document.getElementById("toggle-hide-shipped").classList.add("active");
  document.getElementById("toggle-only-issues").classList.remove("active");
  document.getElementById("reason-row").style.display = "none";
  FILTER_DROPDOWNS.forEach(f => { document.getElementById(f.id).value = ""; });
  updateFilterBadge();
  applyFilters();
}

// ── FILTERS ──
function applyFilters() {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const dropVals = {};
  FILTER_DROPDOWNS.forEach(f => {
    const v = document.getElementById(f.id).value;
    if (v) dropVals[f.column] = v;
  });
  filteredRows = allRows.filter(row => {
    if (activeFilters.hideShipped && HIDDEN_STATUSES.includes(String(row["Status"]||""))) return false;
    if (activeFilters.onlyIssues && !rowIsIssue(row)) return false;
    if (activeFilters.alloc==="allocated"     && !rowIsAllocated(row)) return false;
    if (activeFilters.alloc==="not-allocated" &&  rowIsAllocated(row)) return false;
    if (activeFilters.reason && String(row["Hospital"]||"").trim()!==activeFilters.reason) return false;
    if (activeFilters.sapModule) {
      const rowSap = String(row["SAP Module"]||"").toLowerCase();
      if (!rowSap.includes(activeFilters.sapModule.toLowerCase())) return false;
    }
    for (const col of Object.keys(dropVals)) {
      if (String(row[col]??"") !== dropVals[col]) return false;
    }
    if (q) {
      const hay = SEARCH_FIELDS.map(f => row[f]==null?"":String(row[f]).toLowerCase()).join("|");
      if (!hay.includes(q)) return false;
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
      if (v==null||v===""||String(v)==="#REF!") continue;
      seen.add(String(v));
    }
    [...seen].sort(pickSorter(f.sort)).slice(0, f.limit).forEach(v => {
      const o = document.createElement("option");
      o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}
function pickSorter(kind) {
  if (kind==="weekish") return (a,b) => {
    const wa=parseWeek(a),wb=parseWeek(b);
    if(wa!=null&&wb!=null)return wa-wb;
    if(wa!=null)return -1; if(wb!=null)return 1;
    return a.localeCompare(b);
  };
  if (kind==="monthYear") return (a,b) => {
    const da=parseMonthYear(a),db=parseMonthYear(b);
    if(da&&db)return da-db; if(da)return -1; if(db)return 1;
    return a.localeCompare(b);
  };
  return (a,b) => a.localeCompare(b);
}

// ── RENDER ROWS ──
function renderRows() {
  const container=document.getElementById("rows-container");
  const empty=document.getElementById("empty-state");
  const pag=document.getElementById("pagination");
  document.getElementById("result-count").textContent = filteredRows.length.toLocaleString()+" rows";
  if (!filteredRows.length) {
    container.classList.add("hidden"); pag.classList.add("hidden");
    empty.classList.remove("hidden");
    empty.innerHTML = "<p>No matching rows.</p>";
    return;
  }
  empty.classList.add("hidden"); container.classList.remove("hidden"); pag.classList.remove("hidden");
  const total = Math.max(1, Math.ceil(filteredRows.length/PAGE_SIZE));
  currentPage = Math.max(0, Math.min(currentPage, total-1));
  const slice = filteredRows.slice(currentPage*PAGE_SIZE, (currentPage+1)*PAGE_SIZE);
  container.innerHTML = "";
  slice.forEach(row => container.appendChild(buildRowCard(row)));
  document.getElementById("page-info").textContent = (currentPage+1)+" / "+total;
  document.getElementById("prev-page").disabled = currentPage===0;
  document.getElementById("next-page").disabled = currentPage>=total-1;
}

function buildRowCard(row) {
  const card=document.createElement("div"); card.className="row-card"; card.tabIndex=0;
  const top=document.createElement("div"); top.className="row-card-top";
  const mod=document.createElement("div"); mod.className="row-module";
  mod.textContent=displayValue(row[CARD_FIELDS.primary])||"(no module)";
  const bu=document.createElement("div"); bu.className="row-bu";
  bu.textContent=displayValue(row[CARD_FIELDS.group]);
  top.append(mod,bu);

  const mid=document.createElement("div"); mid.className="row-card-mid";
  const hosp=displayValue(row["Hospital"]), loanset=displayValue(row["Loanset"]);
  const desc=displayValue(row[CARD_FIELDS.secondary]);
  let html = desc?escapeHtml(desc):"";
  if (hosp) html += (html?" · ":"")+`<span class='row-hospital'>${escapeHtml(hosp)}</span>`;
  if (loanset) html += (html?" · ":"")+"LS "+escapeHtml(loanset);
  const allocated = rowIsAllocated(row);
  html += " " + (allocated
    ? `<span class='row-alloc-badge alloc-yes'>Allocated</span>`
    : `<span class='row-alloc-badge alloc-no'>${escapeHtml(getAllocReason(row))}</span>`);
  mid.innerHTML = html;

  const bot=document.createElement("div"); bot.className="row-card-bot";
  const stat=displayValue(row["Status"]), fol=displayValue(row["FOL Status (Auto)"]);
  if (stat) bot.appendChild(buildChip(stat, classifyStatus(stat)));
  if (fol&&fol!=="Please enter FOL ID"&&fol!=="#REF!") bot.appendChild(buildChip(fol,"info"));

  card.append(top,mid,bot);
  card.addEventListener("click", () => openDrawer(row));
  card.addEventListener("keydown", e => { if(e.key==="Enter") openDrawer(row); });
  return card;
}
function getAllocReason(row) {
  const h=String(row["Hospital"]||"").trim();
  if (h==="FORECAST ORDER") return "Forecast";
  if (h==="ORTHOKIT") return "OrthoKit";
  if (h==="Enter Ship To") return "Pending";
  if (h==="Not in Hospital Master") return "Not in master";
  if (h==="") return "No hospital";
  return "Not allocated";
}
function buildChip(text, kind) {
  const c=document.createElement("span"); c.className="chip"+(kind?" chip-"+kind:"");
  c.textContent=text; return c;
}
function classifyStatus(v) {
  const s=v.toLowerCase();
  if (s==="equipment shipped"||s==="equipment built") return "good";
  if (s.includes("issue")||s.includes("fail")) return "bad";
  if (s.includes("await")||s.includes("orthokit")||s.includes("topup")) return "warn";
  if (s.includes("dissolved")) return "info";
  return "";
}

// ── LOAD DATA ──
async function loadData() {
  showEmpty("Initialising…");
  allRows=[]; allHeaders=[]; headerIndex={};
  try {
    showEmpty("Reading headers…");
    await Excel.run(async ctx => {
      const sheet=ctx.workbook.worksheets.getItem(SHEET_NAME);
      const r=sheet.getRange(DATA_FIRST_COL+HEADER_ROW+":"+DATA_LAST_COL+HEADER_ROW);
      r.load("values"); await ctx.sync();
      allHeaders = r.values[0].map(h => h==null?"":String(h));
      allHeaders.forEach((h,i) => { headerIndex[h]=i; });
    });
    const first=HEADER_ROW+1; let offset=0, emptyStreak=0;
    while (offset < MAX_DATA_ROW-first) {
      const size=Math.min(CHUNK_ROWS, MAX_DATA_ROW-first-offset);
      const a=first+offset, b=a+size-1; let nonEmpty=0;
      await Excel.run(async ctx => {
        const sheet=ctx.workbook.worksheets.getItem(SHEET_NAME);
        const r=sheet.getRange(DATA_FIRST_COL+a+":"+DATA_LAST_COL+b);
        r.load("values"); await ctx.sync();
        for (let i=0; i<r.values.length; i++) {
          const row=r.values[i];
          if (!row.some(v => v!==""&&v!==null&&v!==undefined)) continue;
          nonEmpty++;
          const obj={ _row:a+i };
          for (let c=0; c<allHeaders.length; c++) obj[allHeaders[c]] = normalizeCellValue(allHeaders[c], row[c]);
          allRows.push(obj);
        }
      });
      offset+=size;
      if (nonEmpty===0) { if (++emptyStreak>=3) break; } else emptyStreak=0;
      const pct=Math.min(99, Math.round(allRows.length/16000*100));
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
    showToast("Load failed","error");
  }
}

function switchView(name) {
  currentView=name;
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.view===name));
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("active", v.id==="view-"+name));
  if (name==="dashboard") renderDashboard();
  if (name==="week") renderWeekView();
}

// ── DASHBOARD ──
function renderDashboard() {
  const root=document.getElementById("dashboard-content");
  if (!allRows.length) { root.innerHTML="<p class='empty-state'>Load data first.</p>"; return; }
  const active=allRows.filter(r => !HIDDEN_STATUSES.includes(String(r["Status"]||"")));
  const allocated=active.filter(rowIsAllocated);
  const unallocated=active.filter(r => !rowIsAllocated(r));
  const issueCount=active.filter(rowIsIssue).length;
  const bekCounts=countBy(active,"BEK Status (Auto)", v => v&&v!=="NOT IN BEK");
  const folCounts=countBy(active,"FOL Status (Auto)", v => v&&v!=="#REF!"&&v!=="Please enter FOL ID");
  const weekInCounts=countBy(active,"Week To Raise Inbound", v => v&&v!=="#REF!");
  const forecastCounts=countBy(active,"ForecastMonth", v => v&&v!=="#REF!");
  const hospitalCounts=countBy(allocated,"Hospital", v => v&&!["Enter Ship To","ORTHOKIT","FORECAST ORDER","Not in Hospital Master",""].includes(v));
  const reasonCounts=countBy(unallocated,"Hospital", v => v!=="");

  // Default-collapse the larger lists, default-expand overview/allocation
  const sections = [
    { key:"overview", title:"Overview", default:"open", html:`<div class="dash-grid">
        ${tile("Active", active.length.toLocaleString(), "non-shipped rows", "info")}
        ${tile("Allocated", allocated.length.toLocaleString(), "real hospital", "good")}
        ${tile("Not allocated", unallocated.length.toLocaleString(), "no real hospital", unallocated.length>0?"warn":"")}
        ${tile("Issues", issueCount.toLocaleString(), "BEK / KB / flag", issueCount>0?"bad":"good")}
      </div>` },
    { key:"alloc", title:"Allocation breakdown", default:"open", html:`<div class="dash-grid">
        ${tile("Forecast", (reasonCounts["FORECAST ORDER"]||0).toLocaleString(), "future builds", "info")}
        ${tile("OrthoKit", (reasonCounts["ORTHOKIT"]||0).toLocaleString(), "in OrthoKit", "warn")}
        ${tile("Pending", (reasonCounts["Enter Ship To"]||0).toLocaleString(), "no ship-to", "warn")}
        ${tile("Not in master", (reasonCounts["Not in Hospital Master"]||0).toLocaleString(), "hospital missing", "bad")}
      </div>` },
    { key:"bek",      title:"BEK status",                default:"open",     html:listTile(bekCounts,6,true) },
    { key:"fol",      title:"FOL status",                default:"closed",   html:listTile(folCounts,6,true) },
    { key:"weekin",   title:"Upcoming inbound weeks",    default:"closed",   html:listTile(sortCountsForWeek(weekInCounts),10,false) },
    { key:"forecast", title:"Upcoming forecast months",  default:"closed",   html:listTile(sortCountsByMonthYear(forecastCounts),10,false) },
    { key:"hospitals",title:"Top hospitals",             default:"closed",   html:listTile(hospitalCounts,10,true) },
  ];

  root.innerHTML = sections.map(s => {
    const isCollapsed = collapseState.dash[s.key] !== undefined
      ? collapseState.dash[s.key]
      : (s.default==="closed");
    return `<div class="dash-section${isCollapsed?" collapsed":""}" data-key="${s.key}">
      <div class="dash-section-header">
        <span class="dash-section-arrow"></span>
        <span class="dash-section-title">${escapeHtml(s.title)}</span>
      </div>
      <div class="dash-section-body">${s.html}</div>
    </div>`;
  }).join("");

  // Wire collapse-headers
  root.querySelectorAll(".dash-section").forEach(sec => {
    const key = sec.dataset.key;
    sec.querySelector(".dash-section-header").addEventListener("click", () => {
      sec.classList.toggle("collapsed");
      collapseState.dash[key] = sec.classList.contains("collapsed");
    });
  });

  // Tile / list-row click → drill into Browse
  root.querySelectorAll("[data-dash-val]").forEach(el => {
    el.addEventListener("click", e => { e.stopPropagation(); dispatchDashboardAction(el.dataset.dashVal); });
  });
}

function tile(label, value, sub, kind) {
  const cls=kind?" "+kind:"";
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
  if (!entries.length) return "<div class='dash-list-row'>No data</div>";
  return entries.map(([label,n])=>`<div class="dash-list-row" data-dash-val="${escapeAttr(label)}">
    <span class="dash-list-name" title="${escapeAttr(label)}">${escapeHtml(label)}</span>
    <span class="dash-list-count">${n.toLocaleString()}</span>
  </div>`).join("");
}
function sortCountsForWeek(counts) {
  // Values are already normalized to "WW_YYYY" form
  return Object.entries(counts).sort((a,b)=>{
    const ka = parseWeekYear(a[0]), kb = parseWeekYear(b[0]);
    if (ka && kb) {
      if (ka.year !== kb.year) return ka.year - kb.year;
      return ka.week - kb.week;
    }
    if (ka) return -1; if (kb) return 1;
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
  const candidates=["filter-status","filter-week-in","filter-bu"];
  for (const id of candidates) {
    const sel=document.getElementById(id);
    if (!sel) continue;
    for (const opt of sel.options) {
      if (opt.value===value) {
        switchView("browse"); clearFilters();
        sel.value=value; applyFilters(); return;
      }
    }
  }
  switchView("browse"); clearFilters();
  document.getElementById("search").value=value;
  applyFilters();
}

// ── WEEK VIEW ──
function renderWeekView() {
  const root=document.getElementById("week-columns");
  if (!allRows.length) { root.innerHTML="<p class='empty-state'>Load data first.</p>"; return; }
  const today=new Date(), cw=isoWeek(today), cy=today.getFullYear();
  document.getElementById("week-current-label").textContent="Current week: W"+String(cw).padStart(2,"0")+" ("+cy+")";
  document.getElementById("week-config-info").textContent="";

  // Convert (week, year) to a comparable absolute week number for ordering
  const absWeek = (wk, yr) => yr * 53 + wk;
  const cwAbs = absWeek(cw, cy);

  const active=allRows.filter(r => !HIDDEN_STATUSES.includes(String(r["Status"]||"")));
  const buckets={ overdue:[], thisweek:[], nextweek:[], later:[], noweek:[] };
  for (const r of active) {
    const wy = parseWeekYear(r["Final Shipment Week"]);
    if (!wy) { buckets.noweek.push(r); continue; }
    const a = absWeek(wy.week, wy.year);
    if (a < cwAbs) buckets.overdue.push(r);
    else if (a === cwAbs) buckets.thisweek.push(r);
    else if (a === cwAbs + 1) buckets.nextweek.push(r);
    else buckets.later.push(r);
  }
  // Sort overdue oldest-first; later newest-soonest first
  const sortBy = (a,b) => {
    const wa=parseWeekYear(a["Final Shipment Week"]), wb=parseWeekYear(b["Final Shipment Week"]);
    return absWeek(wa.week,wa.year) - absWeek(wb.week,wb.year);
  };
  buckets.overdue.sort(sortBy);
  buckets.later.sort(sortBy);

  const cols = [
    { key:"overdue",  label:"Overdue", cls:"overdue", rows:buckets.overdue, defaultOpen:true },
    { key:"thisweek", label:"This week (W"+String(cw).padStart(2,"0")+")", cls:"thisweek", rows:buckets.thisweek, defaultOpen:true },
    { key:"nextweek", label:"Next week (W"+String(cw+1).padStart(2,"0")+")", cls:"nextweek", rows:buckets.nextweek, defaultOpen:true },
    { key:"later",    label:"Later", cls:"", rows:buckets.later, defaultOpen:false },
    { key:"noweek",   label:"No week set", cls:"", rows:buckets.noweek, defaultOpen:false },
  ];

  root.innerHTML = cols.map(c => {
    const isCollapsed = collapseState.week[c.key] !== undefined
      ? collapseState.week[c.key]
      : (!c.defaultOpen || c.rows.length===0);
    return weekColHtml(c, isCollapsed);
  }).join("");

  root.querySelectorAll(".week-col").forEach(col => {
    const key = col.dataset.key;
    col.querySelector(".week-col-header").addEventListener("click", () => {
      col.classList.toggle("collapsed");
      collapseState.week[key] = col.classList.contains("collapsed");
    });
  });
  root.querySelectorAll(".week-card").forEach(c => {
    c.addEventListener("click", e => {
      e.stopPropagation();
      const row = allRows.find(r => r._row===parseInt(c.dataset.row,10));
      if (row) openDrawer(row);
    });
  });
}

function weekColHtml(col, isCollapsed) {
  const cards = col.rows.slice(0, 200).map(r => {
    const hosp = displayValue(r["Hospital"]) || "—";
    const allocReason = rowIsAllocated(r) ? "" : `<span class='row-alloc-badge alloc-no' style='margin-left:6px;'>${escapeHtml(getAllocReason(r))}</span>`;
    return `<div class="week-card" data-row="${r._row}">
      <div class="week-card-top">
        <span>${escapeHtml(displayValue(r["JDE Module"])||"(no module)")}</span>
        <span class="week-card-bu">${escapeHtml(displayValue(r["Business Unit"])||"")}</span>
      </div>
      <div class="week-card-bot">${escapeHtml(hosp)}${allocReason}</div>
    </div>`;
  }).join("");
  const more = col.rows.length > 200
    ? `<div class="week-overflow">+ ${col.rows.length-200} more (use Browse to see all)</div>` : "";
  const body = col.rows.length
    ? cards + more
    : "<div class='empty-state' style='padding:10px;'>Nothing here.</div>";
  return `<div class="week-col${isCollapsed?" collapsed":""}" data-key="${col.key}">
    <div class="week-col-header ${col.cls}">
      <span class="week-col-arrow"></span>
      <strong>${escapeHtml(col.label)}</strong>
      <span class="week-col-count">${col.rows.length}</span>
    </div>
    <div class="week-col-body">${body}</div>
  </div>`;
}

// ── DRAWER ──
function openDrawer(row) {
  isCreating=false; editingRow=row; editedValues={};
  document.getElementById("save-edit").textContent="Save changes";
  document.getElementById("detail-title").textContent=displayValue(row[CARD_FIELDS.primary])||"Module details";
  const el=document.getElementById("detail-fields"); el.innerHTML="";
  DETAIL_GROUPS.forEach(g => {
    const div=document.createElement("div"); div.className="field-group";
    const t=document.createElement("div"); t.className="field-group-title"; t.textContent=g.title;
    div.appendChild(t);
    g.fields.forEach(f => { if (f.name in headerIndex) div.appendChild(buildField(f, row[f.name])); });
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
  const t=document.createElement("div"); t.className="field-group-title"; t.textContent="New row — * = required";
  g.appendChild(t);
  NEW_ROW_FIELDS.forEach(f => {
    if (!(f.name in headerIndex)) return;
    const wrap=document.createElement("div"); wrap.className="field";
    const lbl=document.createElement("label"); lbl.className="field-label"; lbl.textContent=f.name.trim();
    if (f.required){ const s=document.createElement("span"); s.className="field-required-marker"; s.textContent="*"; lbl.appendChild(s); }
    wrap.appendChild(lbl);
    const inp = f.multiline ? document.createElement("textarea") : document.createElement("input");
    inp.className = f.multiline?"field-textarea":"field-input";
    if (!f.multiline) inp.type="text"; if (f.multiline) inp.rows=3;
    inp.dataset.fieldName = f.name; if (f.required) inp.dataset.required="true";
    inp.addEventListener("input",()=>{ editedValues[f.name]=inp.value; inp.classList.remove("field-error"); });
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
function onSaveClick(){ isCreating ? createNewRow() : saveEdits(); }

// ── FULLSCREEN DIALOG ─────────────────────
// Transport strategy (works without DialogApi 1.2):
//   pane → dialog:  localStorage (same GitHub Pages origin, shared across windows)
//   dialog → pane:  Office messageParent (DialogApi 1.1, universally supported)
//   pane → dialog results: localStorage + the browser 'storage' event
let fsDialog = null;
// Light fields streamed to the dialog as arrays (field names sent once, not per row).
// Heavy comment fields are fetched per-row on demand to stay under localStorage quota.
const FS_LIGHT_FIELDS = ["_row","JDE Module","JDE Description","JDE Build number","Business Unit",
  "Status","Hospital","Loanset","SAP Module","SAP Serial Number","SAP Description",
  "Final Shipment Week","Week To Raise Inbound","9SE1 order number","PO","FID ","Set Reference"];
const FS_HEAVY_FIELDS = ["Ordering Comment ","S&OP COMMENTS","Kit build comments"];
const LS_PREFIX = "kf_";
const LS_CHUNK_ROWS = 2000;

function openFullscreen() {
  if (!allRows.length) { showToast("Load data first.","error"); return; }
  // 1. Write slimmed data to localStorage in chunks (array format — no repeated keys)
  try {
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const k = localStorage.key(i);
      if (k && k.indexOf(LS_PREFIX) === 0) localStorage.removeItem(k);
    }
    const arrays = allRows.map(r => FS_LIGHT_FIELDS.map(f => {
      const v = r[f];
      return (v === "" || v == null) ? 0 : v;   // 0 = empty marker (1 char in JSON)
    }));
    const nChunks = Math.ceil(arrays.length / LS_CHUNK_ROWS);
    for (let i = 0; i < nChunks; i++) {
      localStorage.setItem(LS_PREFIX + "chunk_" + i,
        JSON.stringify(arrays.slice(i*LS_CHUNK_ROWS, (i+1)*LS_CHUNK_ROWS)));
    }
    localStorage.setItem(LS_PREFIX + "meta",
      JSON.stringify({ fields: FS_LIGHT_FIELDS, chunks:nChunks, total:arrays.length, ts:Date.now() }));
  } catch(e) {
    showToast("Couldn't stage data: " + e.message, "error");
    return;
  }
  // 2. Open the dialog — it reads localStorage on load
  const url = window.location.href.replace(/taskpane\.html.*$/, "fullscreen.html");
  Office.context.ui.displayDialogAsync(url,
    { width: 96, height: 94, displayInIframe: false },
    (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        showToast("Couldn't open full screen: " + res.error.message, "error");
        return;
      }
      fsDialog = res.value;
      fsDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
      fsDialog.addEventHandler(Office.EventType.DialogEventReceived, () => { fsDialog = null; });
    }
  );
}

function onDialogMessage(arg) {
  let m;
  try { m = JSON.parse(arg.message); } catch { return; }
  if (m.type === "save") {
    saveFromDialog(m.reqId, m.row, m.edits);
  } else if (m.type === "create") {
    createFromDialog(m.reqId, m.values);
  } else if (m.type === "getRow") {
    // On-demand fetch of heavy fields for one row
    const r = allRows.find(x => x._row === m.row);
    const values = {};
    if (r) FS_HEAVY_FIELDS.forEach(f => { values[f] = r[f] == null ? "" : r[f]; });
    postResultToDialog({ type:"rowDetail", reqId:m.reqId, row:m.row, values:values });
  }
}

// Results go back through localStorage; the dialog listens for 'storage' events.
function postResultToDialog(payload) {
  try {
    localStorage.setItem(LS_PREFIX + "result", JSON.stringify(Object.assign({ ts:Date.now() }, payload)));
  } catch(e) { /* dialog will show its own timeout */ }
}

async function createFromDialog(reqId, values) {
  try {
    const nr = await createRowInSheet(values);
    renderRows();
    const full = allRows.find(r => r._row === nr);
    const slim = {}; FS_LIGHT_FIELDS.forEach(f => { if (full && full[f] != null) slim[f] = full[f]; });
    postResultToDialog({ type:"createResult", reqId:reqId, ok:true, row:slim });
  } catch(err) {
    postResultToDialog({ type:"createResult", reqId:reqId, ok:false, error:String(err.message||err) });
  }
}

async function saveFromDialog(reqId, sheetRow, edits) {
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      Object.keys(edits).forEach(col => {
        const idx = headerIndex[col];
        if (idx === undefined) return;
        sheet.getRange(colLetter(idx) + sheetRow).values = [[ edits[col] ]];
      });
      await ctx.sync();
    });
    const r = allRows.find(x => x._row === sheetRow);
    if (r) Object.keys(edits).forEach(k => { r[k] = edits[k]; });
    renderRows();
    postResultToDialog({ type:"saveResult", reqId:reqId, ok:true, row:sheetRow, edits:edits });
  } catch(err) {
    postResultToDialog({ type:"saveResult", reqId:reqId, ok:false, row:sheetRow, error:String(err.message||err) });
  }
}

async function saveEdits() {
  if (!editingRow) return;
  const changed=Object.keys(editedValues);
  if (!changed.length) { closeDrawer(); return; }
  const btn=document.getElementById("save-edit"); btn.disabled=true; btn.textContent="Saving…";
  try {
    await Excel.run(async ctx => {
      const sheet=ctx.workbook.worksheets.getItem(SHEET_NAME);
      changed.forEach(col => {
        const idx=headerIndex[col]; if (idx===undefined) return;
        sheet.getRange(colLetter(idx)+editingRow._row).values=[[ editedValues[col] ]];
      });
      await ctx.sync();
    });
    changed.forEach(col => { editingRow[col]=editedValues[col]; });
    showToast("Saved "+changed.length+" change"+(changed.length===1?"":"s"),"success");
    renderRows(); closeDrawer();
  } catch(err){ showToast("Save failed: "+(err.message||err),"error"); }
  finally{ btn.disabled=false; btn.textContent="Save changes"; }
}
// Creates a row in the sheet, preserving formulas from the row above.
// copyFrom(...formulas) auto-adjusts relative references like dragging a formula
// down in Excel. Then: user values overwrite their cells; copied constants are cleared.
async function createRowInSheet(values) {
  let nr = HEADER_ROW + 1;
  for (const r of allRows) if (r._row >= nr) nr = r._row + 1;
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
    const lastCol = colLetter(allHeaders.length - 1);
    const prevAddr = "A" + (nr-1) + ":" + lastCol + (nr-1);
    const newAddr  = "A" + nr     + ":" + lastCol + nr;
    const newRange = sheet.getRange(newAddr);
    // Step 1: copy formulas from row above (relative refs adjust automatically)
    newRange.copyFrom(prevAddr, Excel.RangeCopyType.formulas);
    newRange.load("formulas");
    await ctx.sync();
    // Step 2: keep real formulas, insert user values, clear copied constants
    const copied = newRange.formulas[0];
    const out = copied.map((cell, i) => {
      const col = allHeaders[i];
      if (values[col] !== undefined && values[col] !== "") return values[col];
      if (typeof cell === "string" && cell.charAt(0) === "=") return cell;
      return "";
    });
    newRange.formulas = [out];
    sheet.getRange("A" + nr).select();
    await ctx.sync();
  });
  // Local cache: user values now; formula cells show after next Refresh
  const obj = { _row: nr };
  allHeaders.forEach(h => { obj[h] = values[h] !== undefined ? values[h] : ""; });
  allRows.push(obj);
  return nr;
}

async function createNewRow() {
  const missing=[];
  document.querySelectorAll("#detail-fields [data-field-name]").forEach(inp => {
    inp.classList.remove("field-error");
    if (inp.dataset.required==="true" && !inp.value.trim()) {
      missing.push(inp.dataset.fieldName.trim());
      inp.classList.add("field-error");
    }
  });
  const vm=document.getElementById("validation-msg");
  if (missing.length){ vm.textContent="Please fill in: "+missing.join(", "); vm.classList.remove("hidden"); return; }
  vm.classList.add("hidden");
  const btn=document.getElementById("save-edit"); btn.disabled=true; btn.textContent="Creating…";
  try {
    const nr = await createRowInSheet(editedValues);
    showToast("Row created (row "+nr+") — formulas copied from row above","success");
    populateFilterDropdowns(); applyFilters(); closeDrawer();
  } catch(err){ showToast("Create failed: "+(err.message||err),"error"); }
  finally{ btn.disabled=false; btn.textContent="Create row"; }
}

// ── DATE / VALUE HELPERS ──
const MONTH_ABBR=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_YEAR_COLS=new Set(["ForecastMonth"]);
const DATE_COLS=new Set(["RequestedShipDate","Approved to bring kit in incomplete (Date & initial)",
  "Date 9SE1 actually raised","Shipment Date (Also look at auto column)",
  "Anchor Shipment Date (Auto)","JDE Requested Date 06.06.24 (Auto)",
  "TECA Shipment Date ","TECA BO FULFILMRNT DATE  "]);
function excelSerialToDate(n){ if(typeof n!=="number"||!isFinite(n)||n<1||n>200000)return null;
  const d=new Date((n-25569)*86400*1000); return isNaN(d.getTime())?null:d; }
function serialToMonthYear(n){ const d=excelSerialToDate(n); if(!d)return null;
  return MONTH_ABBR[d.getUTCMonth()]+"-"+String(d.getUTCFullYear()).slice(-2); }
function serialToDateStr(n){ const d=excelSerialToDate(n); if(!d)return null;
  return String(d.getUTCDate()).padStart(2,"0")+"-"+MONTH_ABBR[d.getUTCMonth()]+"-"+String(d.getUTCFullYear()).slice(-2); }
function normalizeCellValue(col,raw){
  if (raw===null||raw===undefined||raw==="") return raw;
  if (typeof raw==="number") {
    if (MONTH_YEAR_COLS.has(col)){ const s=serialToMonthYear(raw); if(s)return s; }
    if (DATE_COLS.has(col)){ const s=serialToDateStr(raw); if(s)return s; }
  }
  // Normalize week-style values: "5", "05" → "05_2026" (current year).
  // "05_2026" stays as "05_2026".
  if (col === "Week To Raise Inbound" || col === "Final Shipment Week" || col === "Week Inbound Raised") {
    const s = String(raw).trim();
    if (s === "" || s === "#REF!") return raw;
    // Already in WW_YYYY form? leave it.
    if (/^\d{1,2}_\d{4}$/.test(s)) {
      // Pad week to 2 digits for consistency: "5_2026" → "05_2026"
      const m = s.match(/^(\d{1,2})_(\d{4})$/);
      return String(m[1]).padStart(2,"0") + "_" + m[2];
    }
    // Bare integer like "5" or "05" — treat as current year
    const m = s.match(/^(\d{1,2})$/);
    if (m) {
      const wk = parseInt(m[1],10);
      if (wk >= 1 && wk <= 53) {
        return String(wk).padStart(2,"0") + "_" + new Date().getFullYear();
      }
    }
    // Anything else (like "2024") leave alone
  }
  return raw;
}
function parseMonthYear(s){
  const m=String(s).trim().match(/^([A-Za-z]{3})-(\d{2,4})$/); if(!m)return null;
  const mo={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}[m[1].toLowerCase()];
  if(mo===undefined)return null;
  let y=parseInt(m[2],10); if(y<100)y+=2000;
  return new Date(y,mo,1).getTime();
}
function parseWeek(raw){
  if (raw===null||raw===undefined||raw==="") return null;
  const m=String(raw).trim().match(/^(\d+)(?:_\d{4})?$/); if(!m)return null;
  const wk=parseInt(m[1],10); return (wk>=1&&wk<=53)?wk:null;
}
// Returns { week, year } from a value like "5", "05", "05_2026", or "W05 (2026)".
// Bare week numbers are assumed to belong to the current year.
function parseWeekYear(raw){
  if (raw===null||raw===undefined||raw==="") return null;
  const s=String(raw).trim();
  // Display form: "W05 (2026)"
  let m = s.match(/^W(\d{1,2})\s*\((\d{4})\)$/);
  if (m) return { week: parseInt(m[1],10), year: parseInt(m[2],10) };
  // Raw forms: "5", "05", "05_2026"
  m = s.match(/^(\d+)(?:_(\d{4}))?$/);
  if (!m) return null;
  const wk=parseInt(m[1],10);
  if (wk<1||wk>53) return null;
  const year = m[2] ? parseInt(m[2],10) : new Date().getFullYear();
  return { week: wk, year: year };
}
// Canonical string form: always "WW_YYYY" so "5" (in 2026) === "05_2026"
function canonicalWeekKey(raw){
  const wy = parseWeekYear(raw);
  if (!wy) return null;
  return String(wy.week).padStart(2,"0") + "_" + wy.year;
}
// Human display: "W05 (2026)"
function displayWeekKey(canonical){
  const m = canonical.match(/^(\d{2})_(\d{4})$/);
  if (!m) return canonical;
  return "W" + m[1] + " (" + m[2] + ")";
}
function isoWeek(d){
  const date=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  const day=date.getUTCDay()||7;
  date.setUTCDate(date.getUTCDate()+4-day);
  const y=new Date(Date.UTC(date.getUTCFullYear(),0,1));
  return Math.ceil((((date-y)/86400000)+1)/7);
}
function countBy(rows, col, filterFn) {
  const c={};
  for (const r of rows) {
    const v=r[col]; if (v==null||v==="") continue;
    const s=String(v); if (filterFn&&!filterFn(s)) continue;
    c[s]=(c[s]||0)+1;
  }
  return c;
}
function colLetter(idx){
  let s="",n=idx;
  while (n>=0){ s=String.fromCharCode((n%26)+65)+s; n=Math.floor(n/26)-1; }
  return s;
}
function displayValue(v){
  if (v===undefined||v===null) return "";
  if (typeof v==="number"&&!Number.isInteger(v)) return Math.abs(v)<1?v.toFixed(3):v.toFixed(2);
  return String(v);
}
function escapeHtml(s){ return String(s).replace(/[&<>"']/g, c => ({'&':"&amp;",'<':"&lt;",'>':"&gt;",'"':"&quot;","'":'&#39;'}[c])); }
function escapeAttr(s){ return escapeHtml(s); }
function debounce(fn,ms){ let t; return function(){ clearTimeout(t); const args=arguments, ctx=this; t=setTimeout(()=>fn.apply(ctx,args),ms); }; }
function showEmpty(msg){
  document.getElementById("empty-state").classList.remove("hidden");
  document.getElementById("empty-state").innerHTML="<p>"+escapeHtml(msg)+"</p>";
  document.getElementById("rows-container").classList.add("hidden");
  document.getElementById("pagination").classList.add("hidden");
}
function showToast(msg,kind){
  const t=document.getElementById("toast");
  t.textContent=msg; t.className="toast "+(kind||"");
  setTimeout(()=>{ t.className="toast hidden"; },2400);
}
