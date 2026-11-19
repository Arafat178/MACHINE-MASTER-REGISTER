const EXCEL_FILE = "MACHINE MASTER REGISTER EXCEL.xlsx";

let allMachines = [];
let filteredMachines = [];
let selectedDept = "ALL";
let activeAnalyzePanel = "overview";
let currentModalMachine = null;

// লোকাল স্টোরেজ থেকে Favorites ও Notes লোড করা
let favorites = new Set(JSON.parse(localStorage.getItem("mm_favorites") || "[]"));
let machineNotes = JSON.parse(localStorage.getItem("mm_notes") || "{}");
let compareList = [];

// DOM এলিমেন্ট
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

const registerView = document.getElementById("registerView");
const analyzeView = document.getElementById("analyzeView");
const analyzeContent = document.getElementById("analyzeContent");
const analyzeSubTabs = document.getElementById("analyzeSubTabs");

const compareBar = document.getElementById("compareBar");
const compareCount = document.getElementById("compareCount");
const openCompareBtn = document.getElementById("openCompareBtn");
const clearCompareBtn = document.getElementById("clearCompareBtn");

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

// প্রতিটি মেশিনের ইউনিক কী (ID না থাকলে Sl ব্যবহার করবে)
function getMachineKey(m) {
  return m.id || `SL-${m.sl}-${m.name}`;
}

// ১. Excel ফাইল লোড করা
async function loadExcelData() {
  try {
    const response = await fetch(encodeURI(`./${EXCEL_FILE}`));
    if (!response.ok) throw new Error("Excel file not found!");

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    allMachines = rawRows.map((row, idx) => ({
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
    renderMachines();
    renderAnalyzePanel();
  } catch (error) {
    console.error(error);
    statusMessage.innerText =
      "Failed to load Excel file. Make sure 'MACHINE MASTER REGISTER EXCEL.xlsx' is in the same folder.";
  }
}

// ডিপার্টমেন্ট ট্যাব অটোমেটিক তৈরি করা
function buildDeptTabs() {
  const depts = [...new Set(allMachines.map((m) => m.department).filter(Boolean))].sort();
  let html = `<button class="tab-btn active" data-dept="ALL">All</button>`;
  html += `<button class="tab-btn" data-dept="FAV">★ Pinned</button>`;
  depts.forEach((d) => {
    html += `<button class="tab-btn" data-dept="${d}">${d}</button>`;
  });
  deptTabs.innerHTML = html;
}

// লোকেশন ড্রপডাউন তৈরি
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
  if (d === "engineering") return "badge-engineering";
  if (d === "production") return "badge-production";
  if (d === "warehouse") return "badge-warehouse";
  return "badge-default";
}

// ২. মেশিন লিস্ট রেন্ডার করা
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

  resultCount.textContent = `${filteredMachines.length} / ${allMachines.length}`;
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

// ৩. কম্পেয়ার ফিচার (সর্বোচ্চ ৩টি মেশিন)
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
  if (compareList.length > 0) {
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

// ৪. ডিটেইল মডাল, ফেভারিট ও নোট
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

  modalBody.innerHTML = specs
    .map(
      ([label, val]) => `
      <div class="spec-row">
        <span class="spec-label">${label}</span>
        <span class="spec-value">${val || "—"}</span>
      </div>`
    )
    .join("");

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

// ৫. ফিল্টার্ড CSV এক্সপোর্ট
exportCsvBtn.addEventListener("click", () => {
  if (!filteredMachines.length) return showToast("Nothing to export");
  const headers = ["Sl", "Name", "ID", "Model", "Brand", "Capacity", "Origin", "Department", "Location", "Note"];
  const rows = filteredMachines.map((m) => [
    m.sl, m.name, m.id, m.model, m.brand, m.capacity, m.origin, m.department, m.location,
    machineNotes[getMachineKey(m)] || ""
  ]);
  const csv = [headers, ...rows]
    .map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Machine_Register_${selectedDept}.csv`;
  a.click();
  showToast("CSV downloaded!");
});

// ==========================================
// ৬. অ্যানালাইসিস ড্যাশবোর্ড লজিক (ডাইনামিক)
// ==========================================
function renderAnalyzePanel() {
  if (!allMachines.length) return;

  if (activeAnalyzePanel === "overview") renderOverview();
  else if (activeAnalyzePanel === "audit") renderDataAudit();
  else if (activeAnalyzePanel === "models") renderCommonModels();
  else if (activeAnalyzePanel === "codes") renderIdCodes();
  else if (activeAnalyzePanel === "density") renderBuildingDensity();
}

// হেল্পার: বার চার্ট তৈরি
function makeBarsHtml(entries, total, colorClass = "", clickType = "") {
  const maxVal = entries.length ? entries[0][1] : 1;
  return entries
    .map(([label, count]) => {
      const pct = Math.round((count / maxVal) * 100);
      const share = ((count / total) * 100).toFixed(1);
      return `
        <div class="bar-item" data-click-type="${clickType}" data-click-val="${label.replace(/"/g, "&quot;")}">
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

// ক) Overview: KPI + Dept + Normalized Brands + Countries
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

  // কেস-ইনসেন্সিটিভ ব্র্যান্ড মার্জ (যেমন DAIKIN + Daikin)
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
      <div class="kpi-card"><div class="kpi-val">${uniqueBrands}</div><div class="kpi-label">Unique Brands</div></div>
      <div class="kpi-card"><div class="kpi-val">${uniqueLocs}</div><div class="kpi-label">Locations / Rooms</div></div>
      <div class="kpi-card"><div class="kpi-val" style="color:var(--danger)">${missingCritical}</div><div class="kpi-label">Incomplete Records</div></div>
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

// খ) Data Audit: মিসিং ডাটা + বানানের গরমিল + ডুপ্লিকেট ID
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

  // ডুপ্লিকেট ID চেক
  const idMap = {};
  allMachines.forEach((m) => {
    if (m.id) idMap[m.id] = (idMap[m.id] || 0) + 1;
  });
  const dupIds = Object.entries(idMap).filter(([, c]) => c > 1);

  // বানান/কেস গরমিল চেক (Brand ও Origin)
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

// গ) Common Models: স্পেয়ার পার্টস কমনালিটি
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

// ঘ) ID Code Category Parser (যেমন 4-ENG-DPG-001 -> DPG)
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

// অ্যানালাইসিস বার ক্লিক করলে মেশিন লিস্ট পপ-আপ খোলা
function attachAnalyzeClicks() {
  analyzeContent.querySelectorAll("[data-click-type]").forEach((el) => {
    const type = el.dataset.clickType;
    const val = el.dataset.clickVal;
    if (!type || type === "zone") return;

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

closeAuditModal.addEventListener("click", () => auditModal.classList.remove("open"));
auditModal.addEventListener("click", (e) => {
  if (e.target === auditModal) auditModal.classList.remove("open");
});

// নেভিগেশন ও ইভেন্ট লিসেনার
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    const view = btn.dataset.view;
    registerView.classList.toggle("active", view === "register");
    analyzeView.classList.toggle("active", view === "analyze");
    registerControls.style.display = view === "register" ? "block" : "none";
    if (view === "analyze") renderAnalyzePanel();
  });
});

analyzeSubTabs.addEventListener("click", (e) => {
  if (!e.target.classList.contains("sub-tab")) return;
  document.querySelectorAll(".sub-tab").forEach((t) => t.classList.remove("active"));
  e.target.classList.add("active");
  activeAnalyzePanel = e.target.dataset.panel;
  renderAnalyzePanel();
});

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
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
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
  document.querySelectorAll(".tab-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.dept === "ALL");
  });
  renderMachines();
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

loadExcelData();