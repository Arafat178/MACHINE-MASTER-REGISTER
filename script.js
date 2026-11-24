const EXCEL_FILE = "MACHINE MASTER REGISTER EXCEL.xlsx";
const INSTRUMENT_EXCEL_FILE = "MASTER INSTRUMENT LIST.xlsx";

let allMachines = [];
let filteredMachines = [];
let selectedDept = "ALL";
let activeView = "register";
let activeAnalyzePanel = "overview";
let currentModalMachine = null;

// ইনস্ট্রুমেন্ট স্টেট
let allInstruments = [];
let filteredInstruments = [];
let selectedInstDept = "ALL";
let currentModalInstrument = null;

// মেশিন <-> ইনস্ট্রুমেন্ট অটো লিংক ম্যাপ
// coreKey (যেমন ENG-ACR-001) -> Machine অবজেক্ট
let machineByCoreId = {};
// coreKey -> ওই মেশিনের অধীনে থাকা ইনস্ট্রুমেন্টের লিস্ট
let instrumentsByCoreId = {};

// লোকাল স্টোরেজ থেকে Favorites ও Notes লোড করা
let favorites = new Set(JSON.parse(localStorage.getItem("mm_favorites") || "[]"));
let machineNotes = JSON.parse(localStorage.getItem("mm_notes") || "{}");
let instFavorites = new Set(JSON.parse(localStorage.getItem("mm_inst_favorites") || "[]"));
let instNotes = JSON.parse(localStorage.getItem("mm_inst_notes") || "{}");
let compareList = [];

// DOM এলিমেন্ট (হেডার ও মেশিন)
const headerTitle = document.getElementById("headerTitle");
const headerSubtitle = document.getElementById("headerSubtitle");
const searchInput = document.getElementById("searchInput");
const clearSearchBtn = document.getElementById("clearSearch");
const deptTabs = document.getElementById("deptTabs");
const locationFilter = document.getElementById("locationFilter");
const resetFiltersBtn = document.getElementById("resetFilters");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const machineList = document.getElementById("machineList");
const statusMessage = document.getElementById("statusMessage");
const resultCount = document.getElementById("resultCount");
const registerControls = document.getElementById("registerControls");
const themeToggle = document.getElementById("themeToggle");

// DOM এলিমেন্ট (ইনস্ট্রুমেন্ট)
const instrumentControls = document.getElementById("instrumentControls");
const instSearchInput = document.getElementById("instSearchInput");
const clearInstSearchBtn = document.getElementById("clearInstSearch");
const instDeptTabs = document.getElementById("instDeptTabs");
const instParentFilter = document.getElementById("instParentFilter");
const instCodeFilter = document.getElementById("instCodeFilter");
const exportInstCsvBtn = document.getElementById("exportInstCsvBtn");
const resetInstFiltersBtn = document.getElementById("resetInstFilters");
const instrumentsView = document.getElementById("instrumentsView");
const instrumentList = document.getElementById("instrumentList");
const instStatusMessage = document.getElementById("instStatusMessage");

// ভিউ ও অ্যানালাইসিস
const registerView = document.getElementById("registerView");
const analyzeView = document.getElementById("analyzeView");
const analyzeContent = document.getElementById("analyzeContent");
const analyzeSubTabs = document.getElementById("analyzeSubTabs");

// কম্পেয়ার বার
const compareBar = document.getElementById("compareBar");
const compareCount = document.getElementById("compareCount");
const openCompareBtn = document.getElementById("openCompareBtn");
const clearCompareBtn = document.getElementById("clearCompareBtn");

// মেশিন ডিটেইল মডাল
const detailModal = document.getElementById("detailModal");
const modalDeptBadge = document.getElementById("modalDeptBadge");
const modalTitle = document.getElementById("modalTitle");
const modalId = document.getElementById("modalId");
const modalBody = document.getElementById("modalBody");
const modalFavBtn = document.getElementById("modalFavBtn");
const closeModalBtn = document.getElementById("closeModal");
const noteInput = document.getElementById("noteInput");
const saveNoteBtn = document.getElementById("saveNoteBtn");
const copyBtn = document.getElementById("copyBtn");
const waShareBtn = document.getElementById("waShareBtn");

// ইনস্ট্রুমেন্ট ডিটেইল মডাল
const instDetailModal = document.getElementById("instDetailModal");
const instModalBadge = document.getElementById("instModalBadge");
const instModalTitle = document.getElementById("instModalTitle");
const instModalId = document.getElementById("instModalId");
const instModalBody = document.getElementById("instModalBody");
const instModalFavBtn = document.getElementById("instModalFavBtn");
const closeInstModalBtn = document.getElementById("closeInstModal");
const instNoteInput = document.getElementById("instNoteInput");
const saveInstNoteBtn = document.getElementById("saveInstNoteBtn");
const copyInstBtn = document.getElementById("copyInstBtn");
const waShareInstBtn = document.getElementById("waShareInstBtn");

// কম্পেয়ার ও অডিট মডাল
const compareModal = document.getElementById("compareModal");
const compareTableWrap = document.getElementById("compareTableWrap");
const closeCompareModal = document.getElementById("closeCompareModal");

const auditModal = document.getElementById("auditModal");
const auditModalTitle = document.getElementById("auditModalTitle");
const auditModalSub = document.getElementById("auditModalSub");
const auditModalList = document.getElementById("auditModalList");
const closeAuditModal = document.getElementById("closeAuditModal");

const toast = document.getElementById("toast");

// ডার্ক মোড ইনিশিয়ালাইজ
if (localStorage.getItem("mm_theme") === "dark") {
  document.body.classList.add("dark");
  themeToggle.textContent = "☀️";
}

themeToggle.addEventListener("click", () => {
  const isDark = document.body.classList.toggle("dark");
  localStorage.setItem("mm_theme", isDark ? "dark" : "light");
  themeToggle.textContent = isDark ? "☀️" : "🌙";
});

// প্রতিটি মেশিন ও ইনস্ট্রুমেন্টের ইউনিক কী
function getMachineKey(m) {
  return m.id || `SL-${m.sl}-${m.name}`;
}

function getInstrumentKey(inst) {
  return inst.id || `INST-SL-${inst.sl}-${inst.description}-${inst.parentMc}`;
}

// মেশিন আইডি (যেমন 1-ENG-ACR-001) থেকে কোর আইডি (ENG-ACR-001) বের করা
function extractMachineCoreId(machineId) {
  if (!machineId) return "";
  const parts = machineId.trim().split("-").map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 4) {
    return `${parts[1]}-${parts[2]}-${parts[3]}`.toUpperCase();
  }
  if (parts.length === 3) {
    return `${parts[0]}-${parts[1]}-${parts[2]}`.toUpperCase();
  }
  return "";
}

// ইনস্ট্রুমেন্ট আইডি (যেমন ENG-ACR-001-PG-01) পার্স করা
// ফরম্যাট: dept-machineCode-machineSeq-instCode-instSeq
function parseInstrumentId(rawId) {
  const clean = String(rawId || "").trim().toUpperCase();
  if (!clean) {
    return {
      deptCode: "",
      machineCode: "",
      machineSeq: "",
      parentCoreId: "",
      instCode: "",
      instSeq: "",
      isValidFormat: false,
    };
  }
  const parts = clean.split("-").map((p) => p.trim()).filter(Boolean);
  const deptCode = parts[0] || "";
  const machineCode = parts[1] || "";
  const machineSeq = parts[2] || "";
  const instCode = parts[3] || "";
  const instSeq = parts[4] || "";
  const parentCoreId = parts.length >= 3 ? `${deptCode}-${machineCode}-${machineSeq}` : "";

  return {
    deptCode,
    machineCode,
    machineSeq,
    parentCoreId,
    instCode,
    instSeq,
    isValidFormat: parts.length === 5,
  };
}

// মেশিন ও ইনস্ট্রুমেন্টের মধ্যে অটোমেটিক ক্রস-লিংক তৈরি
function rebuildCrossLinks() {
  machineByCoreId = {};
  instrumentsByCoreId = {};

  allMachines.forEach((m) => {
    const core = extractMachineCoreId(m.id);
    if (core) machineByCoreId[core] = m;
  });

  allInstruments.forEach((inst) => {
    inst.linkedMachine = inst.parentCoreId ? (machineByCoreId[inst.parentCoreId] || null) : null;
    if (inst.parentCoreId) {
      if (!instrumentsByCoreId[inst.parentCoreId]) {
        instrumentsByCoreId[inst.parentCoreId] = [];
      }
      instrumentsByCoreId[inst.parentCoreId].push(inst);
    }
  });
}

// ১. দুটি Excel ফাইল একসাথে লোড করা
async function loadAllData() {
  await Promise.all([loadMachineExcel(), loadInstrumentExcel()]);
  rebuildCrossLinks();
  renderMachines();
  renderInstruments();
  updateHeaderBadge();
  renderAnalyzePanel();
}

async function loadMachineExcel() {
  try {
    const response = await fetch(encodeURI(`./${EXCEL_FILE}`));
    if (!response.ok) throw new Error("Machine Excel file not found!");

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    allMachines = rawRows
      .filter((row) => String(row["Machinery/Equipment Name"] || "").trim() || String(row["ID No."] || "").trim())
      .map((row, idx) => ({
        sl: String(row["Sl. No."] || idx + 1).trim(),
        name: String(row["Machinery/Equipment Name"] || "").trim(),
        id: String(row["ID No."] || "").trim(),
        model: String(row["Model"] || "").trim(),
        brand: String(row["Brand"] || "").trim(),
        capacity: String(row["Capacity"] || "").trim(),
        origin: String(row["Country of Origin"] || "").trim(),
        department: String(row["Department"] || "Other").trim(),
        location: String(row["Location & Room"] || "").trim(),
        namePlate: String(row["NamePlate Name"] || "").trim(),
        receivingDate: String(row["Date of Receiving"] || "").trim(),
        logBy: String(row["Log Maintained By"] || "").trim(),
      }));

    buildDeptTabs();
    populateLocations();
    statusMessage.style.display = "none";
  } catch (error) {
    console.error(error);
    statusMessage.innerText =
      "Failed to load Excel file. Make sure 'MACHINE MASTER REGISTER EXCEL.xlsx' is in the same folder.";
  }
}

async function loadInstrumentExcel() {
  try {
    const response = await fetch(encodeURI(`./${INSTRUMENT_EXCEL_FILE}`));
    if (!response.ok) throw new Error("Instrument Excel file not found!");

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    // শুধু যেসব রো-তে ডেসক্রিপশন, আইডি বা প্যারেন্ট মেশিন আছে সেগুলো নেওয়া হবে
    allInstruments = rawRows
      .filter(
        (row) =>
          String(row["Instrument Description"] || "").trim() ||
          String(row["Instrument ID"] || "").trim() ||
          String(row["Parent M/C"] || "").trim()
      )
      .map((row, idx) => {
        const rawId = String(row["Instrument ID"] || "").trim();
        const parsed = parseInstrumentId(rawId);
        return {
          sl: String(row["Sl."] || idx + 1).trim(),
          description: String(row["Instrument Description"] || "").trim(),
          id: rawId,
          parentMc: String(row["Parent M/C"] || "").trim(),
          location: String(row["Location"] || "").trim(),
          range: String(row["Range"] || "").trim(),
          unit: String(row["Measuring Unit"] || "").trim(),
          remarks: String(row["Remarks"] || "").trim(),
          deptCode: parsed.deptCode,
          machineCode: parsed.machineCode,
          machineSeq: parsed.machineSeq,
          parentCoreId: parsed.parentCoreId,
          instCode: parsed.instCode,
          instSeq: parsed.instSeq,
          isValidFormat: parsed.isValidFormat,
          linkedMachine: null,
        };
      });

    buildInstDeptTabs();
    populateInstFilters();
    instStatusMessage.style.display = "none";
  } catch (error) {
    console.error(error);
    instStatusMessage.innerText =
      "Failed to load Instrument Excel. Make sure 'MASTER INSTRUMENT LIST.xlsx' is in the same folder.";
  }
}

// হেডারের কাউন্ট ব্যাজ আপডেট
function updateHeaderBadge() {
  if (activeView === "register") {
    resultCount.textContent = `${filteredMachines.length} / ${allMachines.length}`;
  } else if (activeView === "instruments") {
    resultCount.textContent = `${filteredInstruments.length} / ${allInstruments.length}`;
  } else {
    resultCount.textContent = `${allMachines.length} M | ${allInstruments.length} I`;
  }
}

// ==========================================
// ২. মেশিন রেজিস্টার ফিল্টার ও রেন্ডার
// ==========================================
function buildDeptTabs() {
  const depts = [...new Set(allMachines.map((m) => m.department).filter(Boolean))].sort();
  let html = `<button class="tab-btn active" data-dept="ALL">All</button>`;
  html += `<button class="tab-btn" data-dept="FAV">★ Pinned</button>`;
  depts.forEach((d) => {
    html += `<button class="tab-btn" data-dept="${d}">${d}</button>`;
  });
  deptTabs.innerHTML = html;
}

function populateLocations() {
  const uniqueLocations = [
    ...new Set(allMachines.map((m) => m.location).filter(Boolean)),
  ].sort();

  locationFilter.innerHTML = `<option value="ALL">All Locations (${uniqueLocations.length})</option>`;
  uniqueLocations.forEach((loc) => {
    const opt = document.createElement("option");
    opt.value = loc;
    opt.textContent = loc;
    locationFilter.appendChild(opt);
  });
}

function getDeptBadgeClass(dept) {
  const d = (dept || "").toLowerCase();
  if (d === "engineering" || d === "eng") return "badge-engineering";
  if (d === "production" || d === "prd") return "badge-production";
  if (d === "warehouse" || d === "whs") return "badge-warehouse";
  return "badge-default";
}

function renderMachines() {
  const query = searchInput.value.trim().toLowerCase();
  const selectedLoc = locationFilter.value;

  filteredMachines = allMachines.filter((m) => {
    const key = getMachineKey(m);
    const matchDept =
      selectedDept === "ALL" ||
      (selectedDept === "FAV" && favorites.has(key)) ||
      m.department === selectedDept;

    const matchLoc = selectedLoc === "ALL" || m.location === selectedLoc;

    const matchSearch =
      !query ||
      m.name.toLowerCase().includes(query) ||
      m.id.toLowerCase().includes(query) ||
      m.model.toLowerCase().includes(query) ||
      m.brand.toLowerCase().includes(query) ||
      m.location.toLowerCase().includes(query) ||
      m.namePlate.toLowerCase().includes(query);

    return matchDept && matchLoc && matchSearch;
  });

  if (activeView === "register") updateHeaderBadge();
  machineList.innerHTML = "";

  if (filteredMachines.length === 0) {
    statusMessage.style.display = "block";
    statusMessage.textContent =
      selectedDept === "FAV"
        ? "No pinned machines yet. Tap ☆ inside any machine to pin it."
        : "No matching machinery found.";
    return;
  }

  statusMessage.style.display = "none";
  const fragment = document.createDocumentFragment();

  filteredMachines.forEach((m) => {
    const key = getMachineKey(m);
    const isFav = favorites.has(key);
    const isCmp = compareList.some((x) => getMachineKey(x) === key);
    const note = machineNotes[key];
    const core = extractMachineCoreId(m.id);
    const linkedInstCount = core && instrumentsByCoreId[core] ? instrumentsByCoreId[core].length : 0;

    const card = document.createElement("div");
    card.className = "machine-card";
    card.innerHTML = `
      <div class="card-top">
        <div class="machine-name">${isFav ? '<span class="star-mark">★</span> ' : ""}${m.name || "Unnamed"}</div>
        <span class="badge ${getDeptBadgeClass(m.department)}">${m.department}</span>
      </div>
      <div class="card-id-row">
        <span class="machine-id">${m.id || "NO ID"}</span>
        <div class="card-actions">
          ${linkedInstCount ? `<span class="inst-count-pill">🎛️ ${linkedInstCount}</span>` : ""}
          <button class="cmp-chip ${isCmp ? "selected" : ""}" data-cmp="1">
            ${isCmp ? "✓ Added" : "+ Compare"}
          </button>
        </div>
      </div>
      <div class="card-meta">
        <span><strong>Brand:</strong> ${m.brand || "—"}</span>
        <span><strong>Model:</strong> ${m.model || "—"}</span>
        <span><strong>Loc:</strong> ${m.location || "—"}</span>
      </div>
      ${note ? `<div class="note-pill">📝 ${note}</div>` : ""}
    `;

    card.querySelector("[data-cmp]").addEventListener("click", (e) => {
      e.stopPropagation();
      toggleCompare(m);
    });

    card.addEventListener("click", () => openModal(m));
    fragment.appendChild(card);
  });

  machineList.appendChild(fragment);
}

// ==========================================
// ৩. মেশিন ইনস্ট্রুমেন্ট লিস্ট ফিল্টার ও রেন্ডার
// ==========================================
function buildInstDeptTabs() {
  const deptCodes = [...new Set(allInstruments.map((i) => i.deptCode).filter(Boolean))].sort();
  const pendingCount = allInstruments.filter((i) => !i.id).length;

  let html = `<button class="tab-btn active" data-inst-dept="ALL">All</button>`;
  html += `<button class="tab-btn" data-inst-dept="FAV">★ Pinned</button>`;
  html += `<button class="tab-btn warn-tab" data-inst-dept="PENDING">⚠ No ID (${pendingCount})</button>`;
  deptCodes.forEach((d) => {
    html += `<button class="tab-btn" data-inst-dept="${d}">${d}</button>`;
  });
  instDeptTabs.innerHTML = html;
}

function populateInstFilters() {
  const uniqueParents = [
    ...new Set(allInstruments.map((i) => i.parentMc).filter(Boolean)),
  ].sort();

  instParentFilter.innerHTML = `<option value="ALL">All Parent M/C (${uniqueParents.length})</option>`;
  uniqueParents.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    instParentFilter.appendChild(opt);
  });

  const uniqueCodes = [
    ...new Set(allInstruments.map((i) => i.instCode).filter(Boolean)),
  ].sort();

  instCodeFilter.innerHTML = `<option value="ALL">All Codes (${uniqueCodes.length})</option>`;
  uniqueCodes.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    instCodeFilter.appendChild(opt);
  });
}

function formatRangeUnit(range, unit) {
  if (!range && !unit) return "—";
  if (range && unit) return `${range} ${unit}`;
  return range || unit;
}

function renderInstruments() {
  const query = instSearchInput.value.trim().toLowerCase();
  const selectedParent = instParentFilter.value;
  const selectedCode = instCodeFilter.value;

  filteredInstruments = allInstruments.filter((inst) => {
    const key = getInstrumentKey(inst);
    const matchDept =
      selectedInstDept === "ALL" ||
      (selectedInstDept === "FAV" && instFavorites.has(key)) ||
      (selectedInstDept === "PENDING" && !inst.id) ||
      inst.deptCode === selectedInstDept;

    const matchParent = selectedParent === "ALL" || inst.parentMc === selectedParent;
    const matchCode = selectedCode === "ALL" || inst.instCode === selectedCode;

    const linkedName = inst.linkedMachine ? inst.linkedMachine.name.toLowerCase() : "";
    const linkedId = inst.linkedMachine ? inst.linkedMachine.id.toLowerCase() : "";

    const matchSearch =
      !query ||
      inst.description.toLowerCase().includes(query) ||
      inst.id.toLowerCase().includes(query) ||
      inst.parentMc.toLowerCase().includes(query) ||
      inst.location.toLowerCase().includes(query) ||
      inst.range.toLowerCase().includes(query) ||
      inst.unit.toLowerCase().includes(query) ||
      inst.instCode.toLowerCase().includes(query) ||
      inst.remarks.toLowerCase().includes(query) ||
      linkedName.includes(query) ||
      linkedId.includes(query);

    return matchDept && matchParent && matchCode && matchSearch;
  });

  if (activeView === "instruments") updateHeaderBadge();
  instrumentList.innerHTML = "";

  if (filteredInstruments.length === 0) {
    instStatusMessage.style.display = "block";
    instStatusMessage.textContent =
      selectedInstDept === "FAV"
        ? "No pinned instruments yet. Tap ☆ inside any instrument to pin it."
        : "No matching instruments found.";
    return;
  }

  instStatusMessage.style.display = "none";
  const fragment = document.createDocumentFragment();

  filteredInstruments.forEach((inst) => {
    const key = getInstrumentKey(inst);
    const isFav = instFavorites.has(key);
    const note = instNotes[key];
    const rangeText = formatRangeUnit(inst.range, inst.unit);
    const badgeClass = inst.deptCode
      ? getDeptBadgeClass(inst.deptCode)
      : "badge-pending";
    const badgeText = inst.instCode
      ? `${inst.deptCode} • ${inst.instCode}`
      : inst.deptCode || "Pending ID";

    const card = document.createElement("div");
    card.className = "machine-card";
    card.innerHTML = `
      <div class="card-top">
        <div class="machine-name">${isFav ? '<span class="star-mark">★</span> ' : ""}${inst.description || "Unnamed Instrument"}</div>
        <span class="badge ${badgeClass}">${badgeText}</span>
      </div>
      <div class="card-id-row">
        <span class="machine-id ${inst.id ? "" : "pending-id"}">${inst.id || "ID NOT ASSIGNED"}</span>
        ${rangeText !== "—" ? `<span class="range-pill">${rangeText}</span>` : ""}
      </div>
      <div class="card-meta">
        <span><strong>Parent M/C:</strong> ${inst.parentMc || "—"}</span>
        <span><strong>Location:</strong> ${inst.location || "—"}</span>
      </div>
      ${
        inst.linkedMachine
          ? `<div class="linked-pill">🔗 ${inst.linkedMachine.id} (${inst.linkedMachine.name})</div>`
          : ""
      }
      ${note ? `<div class="note-pill">📝 ${note}</div>` : ""}
    `;

    card.addEventListener("click", () => openInstModal(inst));
    fragment.appendChild(card);
  });

  instrumentList.appendChild(fragment);
}

// ==========================================
// ৪. কম্পেয়ার ফিচার (সর্বোচ্চ ৩টি মেশিন)
// ==========================================
function toggleCompare(machine) {
  const key = getMachineKey(machine);
  const idx = compareList.findIndex((x) => getMachineKey(x) === key);
  if (idx >= 0) {
    compareList.splice(idx, 1);
  } else {
    if (compareList.length >= 3) {
      showToast("Maximum 3 machines can be compared!");
      return;
    }
    compareList.push(machine);
  }
  updateCompareBar();
  renderMachines();
}

function updateCompareBar() {
  if (compareList.length > 0 && activeView === "register") {
    compareBar.classList.add("show");
    compareCount.textContent = `${compareList.length} machine(s) selected`;
  } else {
    compareBar.classList.remove("show");
  }
}

openCompareBtn.addEventListener("click", () => {
  if (compareList.length < 2) {
    showToast("Select at least 2 machines to compare");
    return;
  }
  const fields = [
    ["ID No.", "id"],
    ["Department", "department"],
    ["Brand", "brand"],
    ["Model", "model"],
    ["Capacity", "capacity"],
    ["Origin", "origin"],
    ["Location", "location"],
    ["NamePlate", "namePlate"],
  ];

  let html = `<table class="cmp-table"><thead><tr><th>Spec</th>`;
  compareList.forEach((m) => (html += `<th>${m.name || "Unnamed"}</th>`));
  html += `</tr></thead><tbody>`;

  fields.forEach(([label, prop]) => {
    html += `<tr><td><strong>${label}</strong></td>`;
    compareList.forEach((m) => (html += `<td>${m[prop] || "—"}</td>`));
    html += `</tr>`;
  });
  html += `</tbody></table>`;

  compareTableWrap.innerHTML = html;
  compareModal.classList.add("open");
});

clearCompareBtn.addEventListener("click", () => {
  compareList = [];
  updateCompareBar();
  renderMachines();
});

closeCompareModal.addEventListener("click", () => compareModal.classList.remove("open"));

// ==========================================
// ৫. মেশিন ডিটেইল মডাল (লিংকড ইনস্ট্রুমেন্টসহ)
// ==========================================
function openModal(m) {
  currentModalMachine = m;
  const key = getMachineKey(m);

  modalDeptBadge.textContent = m.department;
  modalDeptBadge.className = `badge ${getDeptBadgeClass(m.department)}`;
  modalTitle.textContent = m.name || "Unnamed Equipment";
  modalId.textContent = `ID: ${m.id || "N/A"}  |  Sl. #${m.sl}`;

  modalFavBtn.textContent = favorites.has(key) ? "★" : "☆";
  modalFavBtn.classList.toggle("active", favorites.has(key));
  noteInput.value = machineNotes[key] || "";

  const specs = [
    ["Model", m.model],
    ["Brand", m.brand],
    ["Capacity", m.capacity],
    ["Country of Origin", m.origin],
    ["Location & Room", m.location],
    ["NamePlate Name", m.namePlate],
    ["Date of Receiving", m.receivingDate],
    ["Log Maintained By", m.logBy],
  ];

  let bodyHtml = specs
    .map(
      ([label, val]) => `
      <div class="spec-row">
        <span class="spec-label">${label}</span>
        <span class="spec-value">${val || "—"}</span>
      </div>`
    )
    .join("");

  // এই মেশিনের অধীনে থাকা সব ইনস্ট্রুমেন্ট দেখানো
  const core = extractMachineCoreId(m.id);
  const linkedInsts = core && instrumentsByCoreId[core] ? instrumentsByCoreId[core] : [];

  if (linkedInsts.length > 0) {
    bodyHtml += `
      <div class="linked-section">
        <h4>Installed Instruments (${linkedInsts.length})</h4>
        ${linkedInsts
          .map(
            (inst, idx) => `
            <div class="linked-item" data-linked-inst-idx="${idx}">
              <div>
                <div class="linked-item-title">${inst.description || "Instrument"}</div>
                <div class="linked-item-sub">${inst.id || "Pending ID"} • ${inst.location || "—"} • ${formatRangeUnit(inst.range, inst.unit)}</div>
              </div>
              <span class="linked-arrow">→</span>
            </div>`
          )
          .join("")}
      </div>`;
  }

  modalBody.innerHTML = bodyHtml;

  modalBody.querySelectorAll("[data-linked-inst-idx]").forEach((el) => {
    el.addEventListener("click", () => {
      const targetInst = linkedInsts[Number(el.dataset.linkedInstIdx)];
      detailModal.classList.remove("open");
      openInstModal(targetInst);
    });
  });

  detailModal.classList.add("open");
}

modalFavBtn.addEventListener("click", () => {
  if (!currentModalMachine) return;
  const key = getMachineKey(currentModalMachine);
  if (favorites.has(key)) {
    favorites.delete(key);
    showToast("Removed from Pinned");
  } else {
    favorites.add(key);
    showToast("Pinned to Favorites!");
  }
  localStorage.setItem("mm_favorites", JSON.stringify([...favorites]));
  modalFavBtn.textContent = favorites.has(key) ? "★" : "☆";
  modalFavBtn.classList.toggle("active", favorites.has(key));
  renderMachines();
});

saveNoteBtn.addEventListener("click", () => {
  if (!currentModalMachine) return;
  const key = getMachineKey(currentModalMachine);
  const val = noteInput.value.trim();
  if (val) {
    machineNotes[key] = val;
    showToast("Note saved on device!");
  } else {
    delete machineNotes[key];
    showToast("Note cleared!");
  }
  localStorage.setItem("mm_notes", JSON.stringify(machineNotes));
  renderMachines();
});

function buildShareText(m) {
  const note = machineNotes[getMachineKey(m)];
  return `Machine: ${m.name}
ID No.: ${m.id || "N/A"}
Department: ${m.department}
Location: ${m.location || "N/A"}
Brand: ${m.brand || "N/A"}
Model: ${m.model || "N/A"}
Capacity: ${m.capacity || "N/A"}
Origin: ${m.origin || "N/A"}${note ? `\nNote: ${note}` : ""}`;
}

copyBtn.addEventListener("click", async () => {
  if (!currentModalMachine) return;
  try {
    await navigator.clipboard.writeText(buildShareText(currentModalMachine));
    showToast("Machine info copied!");
  } catch {
    showToast("Copy failed!");
  }
});

waShareBtn.addEventListener("click", () => {
  if (!currentModalMachine) return;
  const url = `https://wa.me/?text=${encodeURIComponent(buildShareText(currentModalMachine))}`;
  window.open(url, "_blank");
});

closeModalBtn.addEventListener("click", () => detailModal.classList.remove("open"));
detailModal.addEventListener("click", (e) => {
  if (e.target === detailModal) detailModal.classList.remove("open");
});

// ==========================================
// ৬. ইনস্ট্রুমেন্ট ডিটেইল মডাল (প্যারেন্ট মেশিন লিংকসহ)
// ==========================================
function openInstModal(inst) {
  currentModalInstrument = inst;
  const key = getInstrumentKey(inst);

  instModalBadge.textContent = inst.instCode
    ? `${inst.deptCode} • ${inst.instCode}`
    : inst.deptCode || "Pending ID";
  instModalBadge.className = `badge ${inst.deptCode ? getDeptBadgeClass(inst.deptCode) : "badge-pending"}`;
  instModalTitle.textContent = inst.description || "Unnamed Instrument";
  instModalId.textContent = `ID: ${inst.id || "NOT ASSIGNED"}  |  Sl. #${inst.sl}`;

  instModalFavBtn.textContent = instFavorites.has(key) ? "★" : "☆";
  instModalFavBtn.classList.toggle("active", instFavorites.has(key));
  instNoteInput.value = instNotes[key] || "";

  let html = "";

  // ৫-পার্ট আইডি ব্রেকডাউন কার্ড (আইডি থাকলে দেখাবে)
  if (inst.id) {
    html += `
      <div class="id-breakdown">
        <div class="id-part"><div class="id-part-val">${inst.deptCode || "—"}</div><div class="id-part-lbl">Dept</div></div>
        <div class="id-part"><div class="id-part-val">${inst.machineCode || "—"}</div><div class="id-part-lbl">M/C Code</div></div>
        <div class="id-part"><div class="id-part-val">${inst.machineSeq || "—"}</div><div class="id-part-lbl">M/C Seq</div></div>
        <div class="id-part"><div class="id-part-val">${inst.instCode || "—"}</div><div class="id-part-lbl">Inst Code</div></div>
        <div class="id-part"><div class="id-part-val">${inst.instSeq || "—"}</div><div class="id-part-lbl">Inst Seq</div></div>
      </div>`;
  }

  const specs = [
    ["Parent M/C (Excel)", inst.parentMc],
    ["Mounting / Location", inst.location],
    ["Operating Range", inst.range],
    ["Measuring Unit", inst.unit],
    ["Remarks", inst.remarks],
  ];

  html += specs
    .map(
      ([label, val]) => `
      <div class="spec-row">
        <span class="spec-label">${label}</span>
        <span class="spec-value">${val || "—"}</span>
      </div>`
    )
    .join("");

  // মেশিন মাস্টার রেজিস্টারের সাথে অটো-ম্যাচ হলে ক্লিকেবল কার্ড দেখাবে
  if (inst.linkedMachine) {
    const lm = inst.linkedMachine;
    html += `
      <div class="linked-section">
        <h4>Linked Parent Machine (Master Register)</h4>
        <div class="linked-item" id="jumpToParentMachine">
          <div>
            <div class="linked-item-title">${lm.name}</div>
            <div class="linked-item-sub">${lm.id} • ${lm.brand || "—"} • ${lm.location || "—"}</div>
          </div>
          <span class="linked-arrow">→</span>
        </div>
      </div>`;
  }

  instModalBody.innerHTML = html;

  const jumpBtn = document.getElementById("jumpToParentMachine");
  if (jumpBtn && inst.linkedMachine) {
    jumpBtn.addEventListener("click", () => {
      instDetailModal.classList.remove("open");
      openModal(inst.linkedMachine);
    });
  }

  instDetailModal.classList.add("open");
}

instModalFavBtn.addEventListener("click", () => {
  if (!currentModalInstrument) return;
  const key = getInstrumentKey(currentModalInstrument);
  if (instFavorites.has(key)) {
    instFavorites.delete(key);
    showToast("Removed from Pinned");
  } else {
    instFavorites.add(key);
    showToast("Instrument pinned!");
  }
  localStorage.setItem("mm_inst_favorites", JSON.stringify([...instFavorites]));
  instModalFavBtn.textContent = instFavorites.has(key) ? "★" : "☆";
  instModalFavBtn.classList.toggle("active", instFavorites.has(key));
  renderInstruments();
});

saveInstNoteBtn.addEventListener("click", () => {
  if (!currentModalInstrument) return;
  const key = getInstrumentKey(currentModalInstrument);
  const val = instNoteInput.value.trim();
  if (val) {
    instNotes[key] = val;
    showToast("Instrument note saved!");
  } else {
    delete instNotes[key];
    showToast("Note cleared!");
  }
  localStorage.setItem("mm_inst_notes", JSON.stringify(instNotes));
  renderInstruments();
});

function buildInstShareText(inst) {
  const note = instNotes[getInstrumentKey(inst)];
  return `Instrument: ${inst.description}
Instrument ID: ${inst.id || "Not Assigned"}
Parent M/C: ${inst.parentMc || "N/A"}${inst.linkedMachine ? ` (${inst.linkedMachine.id})` : ""}
Location: ${inst.location || "N/A"}
Range: ${formatRangeUnit(inst.range, inst.unit)}${inst.remarks ? `\nRemarks: ${inst.remarks}` : ""}${note ? `\nNote: ${note}` : ""}`;
}

copyInstBtn.addEventListener("click", async () => {
  if (!currentModalInstrument) return;
  try {
    await navigator.clipboard.writeText(buildInstShareText(currentModalInstrument));
    showToast("Instrument info copied!");
  } catch {
    showToast("Copy failed!");
  }
});

waShareInstBtn.addEventListener("click", () => {
  if (!currentModalInstrument) return;
  const url = `https://wa.me/?text=${encodeURIComponent(buildInstShareText(currentModalInstrument))}`;
  window.open(url, "_blank");
});

closeInstModalBtn.addEventListener("click", () => instDetailModal.classList.remove("open"));
instDetailModal.addEventListener("click", (e) => {
  if (e.target === instDetailModal) instDetailModal.classList.remove("open");
});

// ==========================================
// ৭. CSV এক্সপোর্ট (মেশিন ও ইনস্ট্রুমেন্ট)
// ==========================================
function downloadCsv(filename, headers, rows) {
  const csv = [headers, ...rows]
    .map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  showToast("CSV downloaded!");
}

exportCsvBtn.addEventListener("click", () => {
  if (!filteredMachines.length) return showToast("Nothing to export");
  const headers = ["Sl", "Name", "ID", "Model", "Brand", "Capacity", "Origin", "Department", "Location", "Note"];
  const rows = filteredMachines.map((m) => [
    m.sl, m.name, m.id, m.model, m.brand, m.capacity, m.origin, m.department, m.location,
    machineNotes[getMachineKey(m)] || "",
  ]);
  downloadCsv(`Machine_Register_${selectedDept}.csv`, headers, rows);
});

exportInstCsvBtn.addEventListener("click", () => {
  if (!filteredInstruments.length) return showToast("Nothing to export");
  const headers = ["Sl", "Instrument Description", "Instrument ID", "Parent M/C", "Linked Machine ID", "Location", "Range", "Unit", "Remarks", "Note"];
  const rows = filteredInstruments.map((i) => [
    i.sl, i.description, i.id, i.parentMc, i.linkedMachine ? i.linkedMachine.id : "",
    i.location, i.range, i.unit, i.remarks, instNotes[getInstrumentKey(i)] || "",
  ]);
  downloadCsv(`Instrument_List_${selectedInstDept}.csv`, headers, rows);
});

// ==========================================
// ৮. অ্যানালাইসিস ড্যাশবোর্ড লজিক (ডাইনামিক)
// ==========================================
function renderAnalyzePanel() {
  if (activeAnalyzePanel === "instruments") {
    renderInstrumentAnalytics();
    return;
  }
  if (!allMachines.length) return;

  if (activeAnalyzePanel === "overview") renderOverview();
  else if (activeAnalyzePanel === "audit") renderDataAudit();
  else if (activeAnalyzePanel === "models") renderCommonModels();
  else if (activeAnalyzePanel === "codes") renderIdCodes();
  else if (activeAnalyzePanel === "density") renderBuildingDensity();
}

function makeBarsHtml(entries, total, colorClass = "", clickType = "") {
  const maxVal = entries.length ? entries[0][1] : 1;
  return entries
    .map(([label, count]) => {
      const pct = Math.round((count / maxVal) * 100);
      const share = total ? ((count / total) * 100).toFixed(1) : "0.0";
      return `
        <div class="bar-item" data-click-type="${clickType}" data-click-val="${String(label).replace(/"/g, "&quot;")}">
          <div class="bar-head">
            <span><strong>${label}</strong></span>
            <span>${count} (${share}%)</span>
          </div>
          <div class="bar-track">
            <div class="bar-fill ${colorClass}" style="width:${pct}%"></div>
          </div>
        </div>`;
    })
    .join("");
}

// ক) Overview
function renderOverview() {
  const total = allMachines.length;
  const uniqueBrands = new Set(allMachines.map((m) => m.brand.toLowerCase()).filter(Boolean)).size;
  const uniqueLocs = new Set(allMachines.map((m) => m.location).filter(Boolean)).size;
  const missingCritical = allMachines.filter((m) => !m.id || !m.model || !m.brand || !m.location).length;

  const countBy = (fn) => {
    const map = {};
    allMachines.forEach((m) => {
      const k = fn(m);
      if (k) map[k] = (map[k] || 0) + 1;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  };

  const brandCanonical = {};
  const brandCounts = {};
  allMachines.forEach((m) => {
    if (!m.brand) return;
    const low = m.brand.toLowerCase();
    if (!brandCanonical[low]) brandCanonical[low] = m.brand;
    brandCounts[low] = (brandCounts[low] || 0) + 1;
  });
  const topBrands = Object.entries(brandCounts)
    .map(([low, c]) => [brandCanonical[low], c])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const deptEntries = countBy((m) => m.department);
  const originEntries = countBy((m) => m.origin).slice(0, 8);

  analyzeContent.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-val">${total}</div><div class="kpi-label">Total Machines</div></div>
      <div class="kpi-card"><div class="kpi-val">${allInstruments.length}</div><div class="kpi-label">Total Instruments</div></div>
      <div class="kpi-card"><div class="kpi-val">${uniqueBrands}</div><div class="kpi-label">Unique Brands</div></div>
      <div class="kpi-card"><div class="kpi-val" style="color:var(--danger)">${missingCritical}</div><div class="kpi-label">Incomplete Machines</div></div>
    </div>

    <div class="panel-card">
      <h3>Department Distribution</h3>
      <p class="panel-desc">Total machinery count by department</p>
      ${makeBarsHtml(deptEntries, total, "success")}
    </div>

    <div class="panel-card">
      <h3>Top 10 Brands (Case-Merged)</h3>
      <p class="panel-desc">Tap any brand to see its machines</p>
      ${makeBarsHtml(topBrands, total, "", "brand")}
    </div>

    <div class="panel-card">
      <h3>Top Countries of Origin</h3>
      <p class="panel-desc">Equipment breakdown by manufacturing country</p>
      ${makeBarsHtml(originEntries, total, "warn", "origin")}
    </div>
  `;
  attachAnalyzeClicks();
}

// নতুন: ইনস্ট্রুমেন্ট অ্যানালিটিক্স ও অডিট
function renderInstrumentAnalytics() {
  if (!allInstruments.length) {
    analyzeContent.innerHTML = `<p class="status-msg">No instrument data loaded.</p>`;
    return;
  }

  const total = allInstruments.length;
  const assigned = allInstruments.filter((i) => i.id).length;
  const pending = total - assigned;
  const linked = allInstruments.filter((i) => i.linkedMachine).length;

  // ডুপ্লিকেট ইনস্ট্রুমেন্ট আইডি চেক
  const idMap = {};
  allInstruments.forEach((i) => {
    if (i.id) idMap[i.id.toUpperCase()] = (idMap[i.id.toUpperCase()] || 0) + 1;
  });
  const dupIds = Object.entries(idMap).filter(([, c]) => c > 1);

  // ইনস্ট্রুমেন্ট কোড অনুযায়ী কাউন্ট (PG, PS, TS...)
  const codeMap = {};
  allInstruments.forEach((i) => {
    const c = i.instCode || "NO ID";
    codeMap[c] = (codeMap[c] || 0) + 1;
  });
  const topCodes = Object.entries(codeMap).sort((a, b) => b[1] - a[1]).slice(0, 12);

  // প্যারেন্ট মেশিন অনুযায়ী কাউন্ট
  const parentMap = {};
  allInstruments.forEach((i) => {
    if (!i.parentMc) return;
    parentMap[i.parentMc] = (parentMap[i.parentMc] || 0) + 1;
  });
  const topParents = Object.entries(parentMap).sort((a, b) => b[1] - a[1]).slice(0, 10);

  const instChecks = [
    ["Missing Instrument ID", (i) => !i.id],
    ["Invalid ID Format (Not 5 parts)", (i) => i.id && !i.isValidFormat],
    ["Unlinked to Machine Register", (i) => i.id && !i.linkedMachine],
    ["Missing Mounting Location", (i) => !i.location],
    ["Missing Range", (i) => !i.range],
    ["Missing Measuring Unit", (i) => !i.unit],
  ];

  let healthHtml = dupIds.length
    ? `<div class="alert-item danger"><strong>Duplicate Instrument IDs:</strong> ${dupIds.map(([id, c]) => `${id} (${c}x)`).join(", ")}</div>`
    : `<div class="alert-item" style="border-left-color:var(--success)">✓ No duplicate Instrument IDs found.</div>`;

  const auditRowsHtml = instChecks
    .map(([label, fn], idx) => {
      const count = allInstruments.filter(fn).length;
      return `
        <div class="audit-row" data-inst-audit-idx="${idx}">
          <span>${label}</span>
          <span class="pill-count ${count === 0 ? "ok" : ""}">${count}</span>
        </div>`;
    })
    .join("");

  analyzeContent.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-val">${total}</div><div class="kpi-label">Total Instruments</div></div>
      <div class="kpi-card"><div class="kpi-val" style="color:var(--success)">${assigned}</div><div class="kpi-label">ID Assigned</div></div>
      <div class="kpi-card clickable" id="kpiPendingInst"><div class="kpi-val" style="color:var(--warning)">${pending}</div><div class="kpi-label">Pending ID (Tap)</div></div>
      <div class="kpi-card"><div class="kpi-val">${linked}</div><div class="kpi-label">Linked to Machines</div></div>
    </div>

    <div class="panel-card">
      <h3>Instrument Data Health</h3>
      <p class="panel-desc">Update MASTER INSTRUMENT LIST.xlsx to resolve pending items</p>
      ${healthHtml}
      ${auditRowsHtml}
    </div>

    <div class="panel-card">
      <h3>Instrument Type by Code</h3>
      <p class="panel-desc">Auto-parsed from 4th segment (e.g. ENG-ACR-001-<strong>PG</strong>-01). Tap to inspect.</p>
      ${makeBarsHtml(topCodes, total, "success", "instcode")}
    </div>

    <div class="panel-card">
      <h3>Top 10 Parent Machines (by Instruments)</h3>
      <p class="panel-desc">Tap any machine to see its instruments</p>
      ${makeBarsHtml(topParents, total, "", "instparent")}
    </div>
  `;

  const kpiPending = document.getElementById("kpiPendingInst");
  if (kpiPending) {
    kpiPending.addEventListener("click", () => {
      const list = allInstruments.filter((i) => !i.id);
      openInstAuditModal("Pending Instrument IDs", `${list.length} instruments need ID in Excel`, list);
    });
  }

  analyzeContent.querySelectorAll("[data-inst-audit-idx]").forEach((el) => {
    el.addEventListener("click", () => {
      const [label, fn] = instChecks[Number(el.dataset.instAuditIdx)];
      const list = allInstruments.filter(fn);
      openInstAuditModal(label, `${list.length} instruments`, list);
    });
  });

  analyzeContent.querySelectorAll('[data-click-type="instcode"]').forEach((el) => {
    el.addEventListener("click", () => {
      const code = el.dataset.clickVal;
      const list = allInstruments.filter((i) => (i.instCode || "NO ID") === code);
      openInstAuditModal(`Code: ${code}`, `${list.length} instruments`, list);
    });
  });

  analyzeContent.querySelectorAll('[data-click-type="instparent"]').forEach((el) => {
    el.addEventListener("click", () => {
      const parent = el.dataset.clickVal;
      const list = allInstruments.filter((i) => i.parentMc === parent);
      openInstAuditModal(parent, `${list.length} instruments`, list);
    });
  });
}

// খ) Data Audit (Machines)
function renderDataAudit() {
  const checks = [
    ["Missing ID No.", (m) => !m.id],
    ["Missing Model", (m) => !m.model],
    ["Missing Brand", (m) => !m.brand],
    ["Missing Capacity", (m) => !m.capacity],
    ["Missing Location & Room", (m) => !m.location],
    ["Missing Country of Origin", (m) => !m.origin],
    ["Missing NamePlate Name", (m) => !m.namePlate],
    ["Missing Date of Receiving", (m) => !m.receivingDate],
    ["Missing Log Maintained By", (m) => !m.logBy],
  ];

  const idMap = {};
  allMachines.forEach((m) => {
    if (m.id) idMap[m.id] = (idMap[m.id] || 0) + 1;
  });
  const dupIds = Object.entries(idMap).filter(([, c]) => c > 1);

  const findCaseMismatches = (field) => {
    const groups = {};
    allMachines.forEach((m) => {
      const val = m[field];
      if (!val) return;
      const key = val.toLowerCase();
      if (!groups[key]) groups[key] = new Set();
      groups[key].add(val);
    });
    return Object.values(groups).filter((set) => set.size > 1).map((s) => [...s]);
  };

  const brandMismatches = findCaseMismatches("brand");
  const originMismatches = findCaseMismatches("origin");

  let auditRowsHtml = checks
    .map(([label, fn], idx) => {
      const count = allMachines.filter(fn).length;
      return `
        <div class="audit-row" data-audit-idx="${idx}">
          <span>${label}</span>
          <span class="pill-count ${count === 0 ? "ok" : ""}">${count}</span>
        </div>`;
    })
    .join("");

  let healthHtml = "";
  if (dupIds.length) {
    healthHtml += `<div class="alert-item danger"><strong>Duplicate IDs found:</strong> ${dupIds.map(([id, c]) => `${id} (${c}x)`).join(", ")}</div>`;
  } else {
    healthHtml += `<div class="alert-item" style="border-left-color:var(--success)">✓ No duplicate Machine IDs found.</div>`;
  }

  brandMismatches.forEach((pair) => {
    healthHtml += `<div class="alert-item"><strong>Brand Case Mismatch:</strong> ${pair.join("  vs  ")}</div>`;
  });
  originMismatches.forEach((pair) => {
    healthHtml += `<div class="alert-item"><strong>Origin Case Mismatch:</strong> ${pair.join("  vs  ")}</div>`;
  });

  analyzeContent.innerHTML = `
    <div class="panel-card">
      <h3>Data Health &amp; Inconsistency</h3>
      <p class="panel-desc">Fix these in Excel to keep reports clean</p>
      ${healthHtml}
    </div>

    <div class="panel-card">
      <h3>Missing Data Audit</h3>
      <p class="panel-desc">Tap any row to view the exact incomplete machines</p>
      ${auditRowsHtml}
    </div>
  `;

  analyzeContent.querySelectorAll("[data-audit-idx]").forEach((el) => {
    el.addEventListener("click", () => {
      const [label, fn] = checks[Number(el.dataset.auditIdx)];
      const list = allMachines.filter(fn);
      openAuditModal(label, `${list.length} machines need update in Excel`, list);
    });
  });
}

// গ) Common Models
function renderCommonModels() {
  const modelMap = {};
  allMachines.forEach((m) => {
    if (!m.model) return;
    if (!modelMap[m.model]) modelMap[m.model] = [];
    modelMap[m.model].push(m);
  });

  const repeated = Object.entries(modelMap)
    .filter(([, arr]) => arr.length > 1)
    .sort((a, b) => b[1].length - a[1].length);

  const entries = repeated.map(([model, arr]) => [
    `${model} (${arr[0].brand || "N/A"})`,
    arr.length,
    model,
  ]);

  const maxVal = entries.length ? entries[0][1] : 1;

  analyzeContent.innerHTML = `
    <div class="panel-card">
      <h3>Spare Parts Commonality (${repeated.length} Models)</h3>
      <p class="panel-desc">Identical models shared by multiple machines. Tap to see locations.</p>
      ${
        entries.length
          ? entries
              .map(([label, count, rawModel]) => {
                const pct = Math.round((count / maxVal) * 100);
                return `
                  <div class="bar-item" data-click-type="model" data-click-val="${rawModel.replace(/"/g, "&quot;")}">
                    <div class="bar-head">
                      <span><strong>${label}</strong></span>
                      <span>${count} units</span>
                    </div>
                    <div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div>
                  </div>`;
              })
              .join("")
          : "<p>No repeated models found.</p>"
      }
    </div>
  `;
  attachAnalyzeClicks();
}

// ঘ) ID Code Category Parser
function renderIdCodes() {
  const codeMap = {};
  allMachines.forEach((m) => {
    if (!m.id) return;
    const parts = m.id.split("-");
    const code = parts.length >= 3 ? parts[2].toUpperCase() : "OTHER";
    if (!codeMap[code]) codeMap[code] = [];
    codeMap[code].push(m);
  });

  const entries = Object.entries(codeMap)
    .map(([code, arr]) => [code, arr.length])
    .sort((a, b) => b[1] - a[1]);

  analyzeContent.innerHTML = `
    <div class="panel-card">
      <h3>Equipment Type by ID Code</h3>
      <p class="panel-desc">Auto-parsed from ID pattern (e.g. 4-ENG-<strong>DPG</strong>-001). Tap any code to inspect.</p>
      ${makeBarsHtml(entries, allMachines.length, "success", "idcode")}
    </div>
  `;
  attachAnalyzeClicks();
}

// ঙ) Building & Floor Density
function renderBuildingDensity() {
  const groupLocation = (loc) => {
    if (!loc) return "Unspecified Location";
    const l = loc.trim();
    if (/osd\s*building\s*level\s*\d+/i.test(l)) return l.match(/osd\s*building\s*level\s*\d+/i)[0];
    if (/^osd\s*\d+/i.test(l)) return "OSD Rooms (Numbered)";
    if (/generator/i.test(l)) return "Generator Building";
    if (/warehouse/i.test(l)) return "Warehouse Building";
    if (/engineering\s*&\s*admin/i.test(l)) return "Engineering & Admin Building";
    if (/sub-?station/i.test(l)) return "Sub-Station";
    return l;
  };

  const locMap = {};
  allMachines.forEach((m) => {
    const grp = groupLocation(m.location);
    if (!locMap[grp]) locMap[grp] = [];
    locMap[grp].push(m);
  });

  const entries = Object.entries(locMap)
    .map(([grp, arr]) => [grp, arr.length])
    .sort((a, b) => b[1] - a[1]);

  analyzeContent.innerHTML = `
    <div class="panel-card">
      <h3>Building &amp; Floor Asset Density</h3>
      <p class="panel-desc">Machines grouped by building/floor area. Tap any zone to view list.</p>
      ${makeBarsHtml(entries, allMachines.length, "warn", "zone")}
    </div>
  `;

  analyzeContent.querySelectorAll('[data-click-type="zone"]').forEach((el) => {
    el.addEventListener("click", () => {
      const zone = el.dataset.clickVal;
      openAuditModal(zone, `${locMap[zone].length} machines in this area`, locMap[zone]);
    });
  });
}

function attachAnalyzeClicks() {
  analyzeContent.querySelectorAll("[data-click-type]").forEach((el) => {
    const type = el.dataset.clickType;
    const val = el.dataset.clickVal;
    if (!type || type === "zone" || type === "instcode" || type === "instparent") return;

    el.addEventListener("click", () => {
      let list = [];
      if (type === "brand") {
        list = allMachines.filter((m) => m.brand.toLowerCase() === val.toLowerCase());
      } else if (type === "origin") {
        list = allMachines.filter((m) => m.origin === val);
      } else if (type === "model") {
        list = allMachines.filter((m) => m.model === val);
      } else if (type === "idcode") {
        list = allMachines.filter((m) => {
          const p = (m.id || "").split("-");
          const c = p.length >= 3 ? p[2].toUpperCase() : "OTHER";
          return m.id && c === val;
        });
      }
      openAuditModal(val, `${list.length} matching machines`, list);
    });
  });
}

function openAuditModal(title, sub, list) {
  auditModalTitle.textContent = title;
  auditModalSub.textContent = sub;

  if (!list.length) {
    auditModalList.innerHTML = `<p class="status-msg">All records are complete!</p>`;
  } else {
    auditModalList.innerHTML = list
      .map(
        (m, i) => `
        <div class="audit-row" data-list-idx="${i}">
          <div>
            <div><strong>${m.name || "Unnamed"}</strong></div>
            <div class="subtitle">${m.id || "NO ID"} • ${m.location || "No Location"}</div>
          </div>
          <span class="badge ${getDeptBadgeClass(m.department)}">${m.department}</span>
        </div>`
      )
      .join("");

    auditModalList.querySelectorAll("[data-list-idx]").forEach((row) => {
      row.addEventListener("click", () => {
        const machine = list[Number(row.dataset.listIdx)];
        auditModal.classList.remove("open");
        openModal(machine);
      });
    });
  }
  auditModal.classList.add("open");
}

function openInstAuditModal(title, sub, list) {
  auditModalTitle.textContent = title;
  auditModalSub.textContent = sub;

  if (!list.length) {
    auditModalList.innerHTML = `<p class="status-msg">All instrument records are complete!</p>`;
  } else {
    auditModalList.innerHTML = list
      .map(
        (inst, i) => `
        <div class="audit-row" data-inst-list-idx="${i}">
          <div>
            <div><strong>${inst.description || "Unnamed"}</strong></div>
            <div class="subtitle">${inst.id || "NO ID"} • ${inst.parentMc || "—"}</div>
          </div>
          <span class="badge ${inst.deptCode ? getDeptBadgeClass(inst.deptCode) : "badge-pending"}">${inst.instCode || "Pending"}</span>
        </div>`
      )
      .join("");

    auditModalList.querySelectorAll("[data-inst-list-idx]").forEach((row) => {
      row.addEventListener("click", () => {
        const inst = list[Number(row.dataset.instListIdx)];
        auditModal.classList.remove("open");
        openInstModal(inst);
      });
    });
  }
  auditModal.classList.add("open");
}

closeAuditModal.addEventListener("click", () => auditModal.classList.remove("open"));
auditModal.addEventListener("click", (e) => {
  if (e.target === auditModal) auditModal.classList.remove("open");
});

// ==========================================
// ৯. নেভিগেশন ও ইভেন্ট লিসেনার
// ==========================================
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    activeView = btn.dataset.view;

    registerView.classList.toggle("active", activeView === "register");
    instrumentsView.classList.toggle("active", activeView === "instruments");
    analyzeView.classList.toggle("active", activeView === "analyze");

    registerControls.style.display = activeView === "register" ? "block" : "none";
    instrumentControls.style.display = activeView === "instruments" ? "block" : "none";

    if (activeView === "register") {
      headerTitle.textContent = "Machine Master";
      headerSubtitle.textContent = "Register & Engineering Analytics";
    } else if (activeView === "instruments") {
      headerTitle.textContent = "Instrument List";
      headerSubtitle.textContent = "Sensors, Gauges & Machine Links";
    } else {
      headerTitle.textContent = "Analytics";
      headerSubtitle.textContent = "Machines & Instruments Audit";
      renderAnalyzePanel();
    }

    updateCompareBar();
    updateHeaderBadge();
  });
});

analyzeSubTabs.addEventListener("click", (e) => {
  if (!e.target.classList.contains("sub-tab")) return;
  document.querySelectorAll(".sub-tab").forEach((t) => t.classList.remove("active"));
  e.target.classList.add("active");
  activeAnalyzePanel = e.target.dataset.panel;
  renderAnalyzePanel();
});

// মেশিন ফিল্টার লিসেনার
searchInput.addEventListener("input", () => {
  clearSearchBtn.style.display = searchInput.value ? "block" : "none";
  renderMachines();
});

clearSearchBtn.addEventListener("click", () => {
  searchInput.value = "";
  clearSearchBtn.style.display = "none";
  renderMachines();
});

deptTabs.addEventListener("click", (e) => {
  if (!e.target.classList.contains("tab-btn")) return;
  deptTabs.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  e.target.classList.add("active");
  selectedDept = e.target.dataset.dept;
  renderMachines();
});

locationFilter.addEventListener("change", renderMachines);

resetFiltersBtn.addEventListener("click", () => {
  searchInput.value = "";
  clearSearchBtn.style.display = "none";
  selectedDept = "ALL";
  locationFilter.value = "ALL";
  deptTabs.querySelectorAll(".tab-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.dept === "ALL");
  });
  renderMachines();
});

// ইনস্ট্রুমেন্ট ফিল্টার লিসেনার
instSearchInput.addEventListener("input", () => {
  clearInstSearchBtn.style.display = instSearchInput.value ? "block" : "none";
  renderInstruments();
});

clearInstSearchBtn.addEventListener("click", () => {
  instSearchInput.value = "";
  clearInstSearchBtn.style.display = "none";
  renderInstruments();
});

instDeptTabs.addEventListener("click", (e) => {
  if (!e.target.classList.contains("tab-btn")) return;
  instDeptTabs.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  e.target.classList.add("active");
  selectedInstDept = e.target.dataset.instDept;
  renderInstruments();
});

instParentFilter.addEventListener("change", renderInstruments);
instCodeFilter.addEventListener("change", renderInstruments);

resetInstFiltersBtn.addEventListener("click", () => {
  instSearchInput.value = "";
  clearInstSearchBtn.style.display = "none";
  selectedInstDept = "ALL";
  instParentFilter.value = "ALL";
  instCodeFilter.value = "ALL";
  instDeptTabs.querySelectorAll(".tab-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.instDept === "ALL");
  });
  renderInstruments();
});

function showToast(msg) {
  toast.textContent = msg;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2000);
}

// Service Worker রেজিস্ট্রেশন
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(console.log);
  });
}

loadAllData();
