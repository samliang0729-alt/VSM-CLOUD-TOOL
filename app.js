const sampleState = {
  lineName: "RX59 流線",
  monthlyDemand: 2100,
  workDays: 30,
  shifts: 2,
  hoursPerShift: 12,
  breakMinutes: 0,
  availability: 85,
  stations: [
    { id: "OP10", name: "基礎研磨", ct: 15.1, mct: 12.4, octa: 2.7, octb: 3.8, octc: 90, people: 1, machines: 1, currentWip: 36, bufferMin: 240, yieldRate: 100 },
    { id: "OP20", name: "外觀研磨", ct: 12.1, mct: 9, octa: 3.1, octb: 1.3, octc: 55, people: 1, machines: 1, currentWip: 55, bufferMin: 180, yieldRate: 100 },
    { id: "OP30", name: "整修研磨", ct: 13.2, mct: 10.1, octa: 2.2, octb: 3.6, octc: 90, people: 1, machines: 1, currentWip: 20, bufferMin: 240, yieldRate: 100 },
    { id: "OP40", name: "精修研磨", ct: 19, mct: 16.8, octa: 2.2, octb: 3.6, octc: 90, people: 1, machines: 1, currentWip: 40, bufferMin: 240, yieldRate: 100 }
  ]
};

let state = structuredClone(sampleState);

const MAX_STATIONS = 20;

const globalFields = ["lineName", "monthlyDemand", "workDays", "shifts", "hoursPerShift", "breakMinutes", "availability"];
const stationFields = [
  ["id", "站別", "text", 0.1], ["name", "製程名稱", "text", 0.1], ["ct", "CT", "number", 0.1],
  ["mct", "MCT", "number", 0.1], ["octa", "OCTa", "number", 0.1], ["octb", "OCTb", "number", 0.1],
  ["octc", "C/O 換線／切換時間", "number", 0.1], ["people", "人數", "number", 1], ["machines", "並行機台", "number", 1],
  ["currentWip", "現況 WIP", "number", 1], ["bufferMin", "緩衝分鐘", "number", 1], ["yieldRate", "良率 %", "number", 0.1]
];

function n(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function format(value, digits = 1) {
  if (!Number.isFinite(value)) return "—";
  return value.toLocaleString("zh-TW", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

function formatInt(value) {
  if (!Number.isFinite(value)) return "—";
  return Math.round(value).toLocaleString("zh-TW");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]);
}

function calculate() {
  const monthlyDemand = n(state.monthlyDemand);
  const workDays = n(state.workDays);
  const shifts = n(state.shifts);
  const hours = n(state.hoursPerShift);
  const breaks = n(state.breakMinutes);
  const availability = n(state.availability) / 100;
  const netPerShift = hours * 60 - breaks;
  const valid = monthlyDemand > 0 && workDays > 0 && shifts > 0 && netPerShift > 0 && availability > 0 && availability <= 1;
  const dailyDemand = valid ? monthlyDemand / workDays : NaN;
  const netAvailableDay = valid ? shifts * netPerShift * availability : NaN;
  const takt = valid && dailyDemand > 0 ? netAvailableDay / dailyDemand : NaN;
  const monthlyMinutes = valid ? netAvailableDay * workDays : NaN;

  const stations = state.stations.map((station, index) => {
    const ct = n(station.ct);
    const yieldRate = n(station.yieldRate) / 100;
    const machines = Math.max(1, n(station.machines, 1));
    const capacity = valid && ct > 0 && yieldRate > 0 ? monthlyMinutes / ct * yieldRate * machines : NaN;
    const load = Number.isFinite(capacity) && capacity > 0 ? monthlyDemand / capacity : NaN;
    const targetWip = Number.isFinite(takt) && takt > 0 ? Math.ceil(Math.max(0, n(station.bufferMin)) / takt) : NaN;
    const currentWait = Number.isFinite(takt) ? Math.max(0, n(station.currentWip)) * takt / 60 : NaN;
    const targetWait = Number.isFinite(takt) && Number.isFinite(targetWip) ? targetWip * takt / 60 : NaN;
    const status = !Number.isFinite(load) ? "資料不足" : load > 1 ? "產能不足" : load >= .85 ? "接近滿載" : "正常";
    return { ...station, index, capacity, load, targetWip, currentWait, targetWait, status };
  });

  const validStations = stations.filter((s) => Number.isFinite(s.capacity));
  const bottleneck = validStations.length ? validStations.reduce((a, b) => b.load > a.load ? b : a) : null;
  const currentWip = stations.reduce((sum, s) => sum + Math.max(0, n(s.currentWip)), 0);
  const targetWip = stations.reduce((sum, s) => sum + (Number.isFinite(s.targetWip) ? s.targetWip : 0), 0);
  const vaMinutes = stations.reduce((sum, s) => sum + Math.max(0, n(s.ct)), 0);
  const nvaHours = stations.reduce((sum, s) => sum + (Number.isFinite(s.currentWait) ? s.currentWait : 0), 0);
  const leadTime = nvaHours + vaMinutes / 60;

  return { valid, dailyDemand, netAvailableDay, takt, monthlyMinutes, stations, bottleneck, currentWip, targetWip, vaMinutes, nvaHours, leadTime };
}

function statusClass(status) {
  return status === "產能不足" ? "status-over" : status === "接近滿載" ? "status-near" : "status-normal";
}

function getKpiCards(calc) {
  const status = calc.bottleneck?.load > 1 ? "產能不足" : calc.bottleneck?.load >= .85 ? "接近滿載" : calc.bottleneck ? "產能正常" : "資料不足";
  return [
    ["日需求", format(calc.dailyDemand, 1), "pcs", ""],
    ["需求 Takt", format(calc.takt, 1), "min/pcs", ""],
    ["瓶頸站", calc.bottleneck?.id || "—", calc.bottleneck?.name || "", calc.bottleneck?.load > 1 ? "danger" : calc.bottleneck?.load >= .85 ? "warning" : ""],
    ["瓶頸月產能", formatInt(calc.bottleneck?.capacity), "pcs", calc.bottleneck?.load > 1 ? "danger" : ""],
    ["總現況 WIP", formatInt(calc.currentWip), "pcs", ""],
    ["總目標 WIP", formatInt(calc.targetWip), "pcs", calc.currentWip > calc.targetWip ? "warning" : ""],
    ["現況 Lead Time", format(calc.leadTime, 1), "hr", ""],
    ["產能判定", status, calc.bottleneck && Number.isFinite(calc.bottleneck.load) ? `${format(calc.bottleneck.load * 100, 1)}% 負荷` : "", status === "產能不足" ? "danger" : status === "接近滿載" ? "warning" : ""]
  ];
}

function renderKpis(calc) {
  const cards = getKpiCards(calc);
  document.getElementById("kpiGrid").innerHTML = cards.map(([label, value, unit, tone]) => `
    <article class="kpi ${tone}"><small>${escapeHtml(label)}</small><strong>${escapeHtml(value)} <span>${escapeHtml(unit)}</span></strong></article>
  `).join("");
}

function processCard(station, isBottleneck) {
  return `
    <div class="process-node">
      <div class="flow-arrow" aria-hidden="true">→</div>
      <article class="process-card ${isBottleneck ? "bottleneck" : ""}">
        <div class="process-title"><strong>${escapeHtml(station.id || `站 ${station.index + 1}`)}</strong><span>${escapeHtml(station.name || "未命名製程")}</span></div>
        <dl class="process-metrics">
          <div><dt>CT</dt><dd>${format(n(station.ct), 1)} min</dd></div>
          <div><dt>MCT</dt><dd>${format(n(station.mct), 1)} min</dd></div>
          <div><dt>OCT a / b</dt><dd>${format(n(station.octa), 1)} / ${format(n(station.octb), 1)}</dd></div>
          <div><dt>C/O Changeover Time</dt><dd>${format(n(station.octc), 1)} min</dd></div>
          <div><dt>人數／機台</dt><dd>${formatInt(n(station.people))}／${formatInt(n(station.machines))}</dd></div>
          <div><dt>月產能</dt><dd>${formatInt(station.capacity)} pcs</dd></div>
          <div><dt>負荷率</dt><dd class="load-value">${format(station.load * 100, 1)}%</dd></div>
        </dl>
        <div class="status-badge ${statusClass(station.status)}">${escapeHtml(station.status)}${isBottleneck ? " · 瓶頸" : ""}</div>
      </article>
    </div>`;
}

function inventoryNode(station) {
  return `
    <div class="inventory-node">
      <div class="flow-arrow" aria-hidden="true">→</div>
      <div class="inventory-body" aria-label="${escapeHtml(station.id)} 站後 WIP">
        <svg class="triangle-icon" viewBox="0 0 64 48" role="img" aria-label="庫存三角形"><path d="M32 4 60 44H4Z" fill="none" stroke="currentColor" stroke-width="3"/></svg>
        <div class="wip-current">${formatInt(n(station.currentWip))} pcs</div>
        <div class="wip-target">目標 ${formatInt(station.targetWip)} pcs</div>
        <div class="inventory-label">${format(station.currentWait, 1)} hr 等待</div>
      </div>
    </div>`;
}

function renderVsm(calc) {
  const stationFlow = calc.stations.map((station, index) => {
    const isBottleneck = calc.bottleneck?.index === station.index;
    return processCard(station, isBottleneck) + (index < calc.stations.length - 1 ? inventoryNode(station) : "");
  }).join("");
  const timeline = calc.stations.map((station, index) => `
    <div class="time-va">VA ${format(n(station.ct), 1)} min</div>
    ${index < calc.stations.length - 1 ? `<div class="time-nva">NVA ${format(station.currentWait, 1)} hr</div>` : ""}
  `).join("");

  document.getElementById("vsmCanvas").innerHTML = `
    <div class="info-flow">
      <div class="entity">供應商</div><div class="signal" aria-hidden="true"></div><div class="entity control">生產管制</div><div class="signal" aria-hidden="true"></div><div class="entity">客戶</div>
    </div>
    <div class="process-flow">${stationFlow || '<p class="empty-state">請新增至少一個站別</p>'}</div>
    <div class="timeline">
      <p class="timeline-title">價值流時間軸</p>
      <div class="timeline-track">${timeline}</div>
      <div class="timeline-total">
        <div><span>總 VA 時間</span><strong>${format(calc.vaMinutes, 1)} min</strong></div>
        <div><span>總 NVA 等待</span><strong>${format(calc.nvaHours, 1)} hr</strong></div>
        <div><span>Lead Time</span><strong>${format(calc.leadTime, 1)} hr</strong></div>
      </div>
    </div>`;
}

function renderStationEditors() {
  const container = document.getElementById("stationEditors");
  container.innerHTML = state.stations.map((station, index) => `
    <article class="station-editor" data-index="${index}">
      <div class="station-editor-head">
        <span class="station-number"><i>${index + 1}</i>${escapeHtml(station.id || "新站別")}</span>
        <button class="button danger remove-station" type="button" data-index="${index}" aria-label="刪除第 ${index + 1} 站 ${escapeHtml(station.id || "新站別")}" ${state.stations.length === 1 ? "disabled title=\"至少保留 1 個站別\"" : ""}>刪除</button>
      </div>
      <div class="station-input-grid">
        ${stationFields.map(([key, label, type, step]) => `<label class="station-field"><span>${label}</span><input data-field="${key}" type="${type}" ${type === "number" ? `min="0" step="${step}" inputmode="decimal"` : ""} value="${escapeHtml(station[key])}" aria-label="第 ${index + 1} 站 ${label}"></label>`).join("")}
      </div>
    </article>`).join("");

  container.querySelectorAll("input").forEach((input) => input.addEventListener("input", (event) => {
    const editor = event.target.closest(".station-editor");
    const index = Number(editor.dataset.index);
    const key = event.target.dataset.field;
    state.stations[index][key] = event.target.type === "number" ? n(event.target.value) : event.target.value;
    updateResults();
    editor.querySelector(".station-number").lastChild.textContent = state.stations[index].id || "新站別";
  }));
  container.querySelectorAll(".remove-station").forEach((button) => button.addEventListener("click", () => {
    state.stations.splice(Number(button.dataset.index), 1);
    renderAll();
  }));
}

function validate(calc) {
  let message = "";
  if (n(state.monthlyDemand) <= 0) message = "月需求必須大於 0。";
  else if (n(state.workDays) <= 0) message = "工作天必須大於 0。";
  else if (n(state.hoursPerShift) * 60 <= n(state.breakMinutes)) message = "每班休息時間不可大於或等於每班工時。";
  else if (!state.stations.length) message = "請至少新增一個站別。";
  else if (state.stations.length > MAX_STATIONS) message = `站別最多 ${MAX_STATIONS} 個。`;
  else if (state.stations.some((s) => n(s.ct) <= 0)) message = "每個站別的 CT 必須大於 0。";
  document.getElementById("validationMessage").textContent = message;
  document.getElementById("globalInputs").toggleAttribute("data-invalid", !calc.valid);
  return !message;
}

function updateResults() {
  const calc = calculate();
  document.getElementById("lineTitle").textContent = state.lineName || "未命名產線";
  validate(calc);
  renderKpis(calc);
  renderVsm(calc);
  const addButton = document.getElementById("addStationButton");
  const atStationLimit = state.stations.length >= MAX_STATIONS;
  addButton.disabled = atStationLimit;
  addButton.textContent = atStationLimit ? `已達 ${MAX_STATIONS} 站` : "新增站別";
  document.getElementById("stationCountStatus").textContent = `目前 ${state.stations.length} / ${MAX_STATIONS} 站（至少保留 1 站）`;
}

function bindGlobals() {
  globalFields.forEach((field) => {
    const element = document.getElementById(field);
    element.value = state[field];
    element.addEventListener("input", () => {
      state[field] = element.type === "number" ? n(element.value) : element.value;
      updateResults();
    });
  });
}

function renderAll() {
  globalFields.forEach((field) => { document.getElementById(field).value = state[field]; });
  renderStationEditors();
  updateResults();
}

let importMessageTimer;
function setStationImportStatus(message, isError = false) {
  const status = document.getElementById("stationImportStatus");
  clearTimeout(importMessageTimer);
  status.textContent = message;
  status.classList.toggle("error", isError);
  importMessageTimer = setTimeout(() => {
    status.textContent = "";
    status.classList.remove("error");
  }, 8000);
}

function normalizeHeader(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/[\s\\/()（）%％._:：-]+/g, "");
}

const stationImportAliases = {
  id: ["Station / 站別", "站別", "站別代碼", "工站", "工站代碼", "station", "stationid"],
  name: ["Process / 製程名稱", "製程名稱", "製程", "站別名稱", "工站名稱", "process", "processname"],
  ct: ["CT (min)", "CT", "C/T", "cycle time", "cycletime"],
  mct: ["MCT (min)", "MCT"],
  octa: ["OCTa (min)", "OCTa"],
  octb: ["OCTb (min)", "OCTb"],
  octc: ["C/O Changeover Time (min) / 換線／切換時間", "C/O Changeover Time (min) / 換線時間", "C/O Changeover Time (min)", "C/O Changeover Time", "C/O (min)", "C/O", "Changeover Time", "換線時間", "切換時間", "OCTc (min)", "OCTc"],
  people: ["Operators / 人數", "人數", "作業人數", "operators", "people"],
  machines: ["Machines / 並行機台", "並行機台", "機台數", "設備數", "machines"],
  currentWip: ["Current WIP (pcs) / 現況 WIP", "現況 WIP (pcs)", "現況WIP", "WIP", "現場WIP", "currentwip"],
  bufferMin: ["Buffer (min) / 緩衝", "緩衝 (min)", "緩衝分鐘", "緩衝時間", "buffer", "buffermin"],
  yieldRate: ["Yield (%) / 良率", "良率 (%)", "良率", "良率%", "yield", "yieldrate"]
};

function findStationHeaderMap(row) {
  const normalizedRow = row.map(normalizeHeader);
  const map = {};
  for (const [field, aliases] of Object.entries(stationImportAliases)) {
    const normalizedAliases = aliases.map(normalizeHeader);
    map[field] = normalizedRow.findIndex((header) => normalizedAliases.includes(header));
  }
  return map.id >= 0 && map.name >= 0 && map.ct >= 0 ? map : null;
}

function isBlankCell(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function readImportedNumber(row, columnIndex, label, excelRow, defaultValue) {
  if (columnIndex < 0 || isBlankCell(row[columnIndex])) return defaultValue;
  const raw = row[columnIndex];
  const cleaned = typeof raw === "string" ? raw.trim().replace(/,/g, "").replace(/%$/, "") : raw;
  const value = Number(cleaned);
  if (!Number.isFinite(value)) throw new Error(`第 ${excelRow} 列「${label}」必須是數字。`);
  return value;
}

function parseStationWorkbook(workbook) {
  const sheetNames = ["站別資料", ...workbook.SheetNames.filter((name) => name !== "站別資料")];
  let selected = null;

  for (const sheetName of sheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const scanLimit = Math.min(20, rows.length);
    for (let headerIndex = 0; headerIndex < scanLimit; headerIndex += 1) {
      const headerMap = findStationHeaderMap(rows[headerIndex] || []);
      if (headerMap) {
        selected = { sheetName, rows, headerIndex, headerMap };
        break;
      }
    }
    if (selected) break;
  }

  if (!selected) throw new Error("找不到欄位標題。請使用下載的範本，且保留「站別、製程名稱、CT」欄位。");

  const stations = [];
  const { rows, headerIndex, headerMap } = selected;
  const dataFields = Object.keys(stationImportAliases);

  for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const hasStationData = dataFields.some((field) => headerMap[field] >= 0 && !isBlankCell(row[headerMap[field]]));
    if (!hasStationData) continue;

    const excelRow = rowIndex + 1;
    const id = String(row[headerMap.id] ?? "").trim();
    const name = String(row[headerMap.name] ?? "").trim();
    if (!id) throw new Error(`第 ${excelRow} 列缺少「站別」。`);
    if (!name) throw new Error(`第 ${excelRow} 列缺少「製程名稱」。`);

    const ct = readImportedNumber(row, headerMap.ct, "CT", excelRow, NaN);
    const mct = readImportedNumber(row, headerMap.mct, "MCT", excelRow, 0);
    const octa = readImportedNumber(row, headerMap.octa, "OCTa", excelRow, 0);
    const octb = readImportedNumber(row, headerMap.octb, "OCTb", excelRow, 0);
    const octc = readImportedNumber(row, headerMap.octc, "C/O 換線／切換時間", excelRow, 0);
    const people = readImportedNumber(row, headerMap.people, "人數", excelRow, 1);
    const machines = readImportedNumber(row, headerMap.machines, "並行機台", excelRow, 1);
    const currentWip = readImportedNumber(row, headerMap.currentWip, "現況 WIP", excelRow, 0);
    const bufferMin = readImportedNumber(row, headerMap.bufferMin, "緩衝分鐘", excelRow, 0);
    let yieldRate = readImportedNumber(row, headerMap.yieldRate, "良率", excelRow, 100);
    if (yieldRate > 0 && yieldRate <= 1) yieldRate *= 100;

    if (!(ct > 0)) throw new Error(`第 ${excelRow} 列「CT」必須大於 0。`);
    if ([mct, octa, octb, octc, currentWip, bufferMin].some((value) => value < 0)) throw new Error(`第 ${excelRow} 列的時間、WIP 與緩衝不可為負數。`);
    if (!Number.isInteger(people) || people < 1) throw new Error(`第 ${excelRow} 列「人數」必須是大於或等於 1 的整數。`);
    if (!Number.isInteger(machines) || machines < 1) throw new Error(`第 ${excelRow} 列「並行機台」必須是大於或等於 1 的整數。`);
    if (!(yieldRate > 0 && yieldRate <= 100)) throw new Error(`第 ${excelRow} 列「良率」必須大於 0 且不超過 100%。`);

    stations.push({ id, name, ct, mct, octa, octb, octc, people, machines, currentWip, bufferMin, yieldRate });
  }

  if (!stations.length) throw new Error("沒有找到可匯入的站別資料。請至少填寫 1 個站別。");
  if (stations.length > MAX_STATIONS) throw new Error(`範本共有 ${stations.length} 個站別，目前最多可匯入 ${MAX_STATIONS} 個。`);
  const ids = stations.map((station) => station.id.toLocaleLowerCase("zh-TW"));
  if (new Set(ids).size !== ids.length) throw new Error("站別代碼不可重複，請檢查「站別」欄位。");
  return stations;
}

async function importStationExcel(file) {
  if (!file) return;
  if (typeof XLSX === "undefined") {
    setStationImportStatus("Excel 元件尚未載入，請確認網路後再試一次。", true);
    return;
  }
  if (file.size > 5 * 1024 * 1024) {
    setStationImportStatus("檔案超過 5 MB，請刪除不必要的圖片或工作表後再匯入。", true);
    return;
  }

  setStationImportStatus("正在讀取 Excel…");
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const stations = parseStationWorkbook(workbook);
    if (!window.confirm(`將以 Excel 的 ${stations.length} 個站別取代目前資料，是否繼續？`)) {
      setStationImportStatus("已取消匯入。");
      return;
    }
    state.stations = stations;
    renderAll();
    setStationImportStatus(`已匯入 ${stations.length} 個站別並更新 VSM。`);
    document.getElementById("stationEditors").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    console.error("Station import failed", error);
    setStationImportStatus(error instanceof Error ? error.message : "Excel 匯入失敗，請確認範本格式。", true);
  }
}

let exportMessageTimer;
function setExportStatus(message, isError = false) {
  const status = document.getElementById("exportStatus");
  clearTimeout(exportMessageTimer);
  status.textContent = message;
  status.classList.toggle("error", isError);
  exportMessageTimer = setTimeout(() => {
    status.textContent = "";
    status.classList.remove("error");
  }, 5000);
}

function safeFileName(value) {
  const cleaned = String(value || "VSM產線").trim().replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_");
  return cleaned || "VSM產線";
}

function dateStamp() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

const printStationsPerPage = 4;

function groupStationsForPrint(stations) {
  const groups = [];
  for (let index = 0; index < stations.length; index += printStationsPerPage) {
    groups.push(stations.slice(index, index + printStationsPerPage));
  }
  return groups;
}

function printKpiMarkup(calc) {
  return getKpiCards(calc).map(([label, value, unit, tone]) => `
    <article class="print-kpi ${tone}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}<small>${escapeHtml(unit)}</small></strong>
    </article>
  `).join("");
}

function printStationFlowMarkup(stations, calc) {
  return stations.map((station) => {
    const isBottleneck = calc.bottleneck?.index === station.index;
    const hasNextStation = station.index < calc.stations.length - 1;
    return processCard(station, isBottleneck) + (hasNextStation ? inventoryNode(station) : "");
  }).join("");
}

function printTimelineMarkup(stations, calc) {
  return stations.map((station) => {
    const hasNextStation = station.index < calc.stations.length - 1;
    return `
      <div class="time-va">VA ${format(n(station.ct), 1)} min</div>
      ${hasNextStation ? `<div class="time-nva">NVA ${format(station.currentWait, 1)} hr</div>` : ""}
    `;
  }).join("");
}

function buildPrintReport(calc) {
  const groups = groupStationsForPrint(calc.stations);
  const totalPages = groups.length;
  const report = document.getElementById("printReport");
  report.innerHTML = groups.map((stations, pageIndex) => {
    const firstStation = stations[0]?.index + 1;
    const lastStation = stations.at(-1)?.index + 1;
    const isFirstPage = pageIndex === 0;
    const isLastPage = pageIndex === totalPages - 1;
    const nextStation = !isLastPage ? groups[pageIndex + 1][0] : null;
    return `
      <section class="print-page" data-page="${pageIndex + 1}">
        <header class="print-report-header">
          <div class="print-brand" aria-hidden="true">V</div>
          <div><p>VALUE STREAM MANAGEMENT</p><h1>${escapeHtml(state.lineName || "未命名產線")}</h1></div>
          <div class="print-page-meta">第 ${pageIndex + 1} / ${totalPages} 頁<br>${dateStamp()}</div>
        </header>
        ${isFirstPage ? `
          <section class="print-summary">
            <div class="print-section-title"><h2>即時結果</h2><span>生產條件與現況計算</span></div>
            <div class="print-kpi-grid">${printKpiMarkup(calc)}</div>
          </section>
        ` : ""}
        <section class="print-vsm-section">
          <div class="print-section-title"><h2>現況 VSM</h2><span>站別 ${firstStation}–${lastStation}，共 ${calc.stations.length} 站</span></div>
          <div class="print-info-flow">
            <div class="entity">供應商</div><div class="signal" aria-hidden="true"></div><div class="entity control">生產管制</div><div class="signal" aria-hidden="true"></div><div class="entity">客戶</div>
          </div>
          <div class="print-process-flow">${printStationFlowMarkup(stations, calc)}</div>
          <div class="timeline print-timeline">
            <p class="timeline-title">價值流時間軸</p>
            <div class="timeline-track">${printTimelineMarkup(stations, calc)}</div>
            ${isLastPage ? `
              <div class="timeline-total">
                <div><span>總 VA 時間</span><strong>${format(calc.vaMinutes, 1)} min</strong></div>
                <div><span>總 NVA 等待</span><strong>${format(calc.nvaHours, 1)} hr</strong></div>
                <div><span>Lead Time</span><strong>${format(calc.leadTime, 1)} hr</strong></div>
              </div>
            ` : `<div class="print-continuation">流程接續下一頁：${escapeHtml(nextStation?.id || "下一站")}</div>`}
          </div>
        </section>
      </section>
    `;
  }).join("");
  return totalPages;
}

function exportPdf() {
  const calc = calculate();
  if (!validate(calc)) {
    setExportStatus("請先修正輸入資料，再匯出 PDF。", true);
    return;
  }
  const totalPages = buildPrintReport(calc);
  const originalTitle = document.title;
  document.title = `VSM_${safeFileName(state.lineName)}_${dateStamp()}`;
  setExportStatus(`已整理為 ${totalPages} 頁，請在列印視窗選擇「另存為 PDF」。`);
  window.print();
  document.title = originalTitle;
}

function setCellFormat(sheet, address, numberFormat) {
  if (sheet[address]) sheet[address].z = numberFormat;
}

function exportExcel() {
  const calc = calculate();
  if (!validate(calc)) {
    setExportStatus("請先修正輸入資料，再匯出 Excel。", true);
    return;
  }
  if (typeof XLSX === "undefined") {
    setExportStatus("Excel 匯出元件尚未載入，請確認網路後再試一次。", true);
    return;
  }

  const status = calc.bottleneck?.load > 1 ? "產能不足" : calc.bottleneck?.load >= .85 ? "接近滿載" : "產能正常";
  const summaryRows = [
    ["VSM 產線分析", state.lineName],
    ["匯出日期", new Date()],
    [],
    ["生產條件", "數值", "單位"],
    ["月需求", n(state.monthlyDemand), "pcs"],
    ["工作天", n(state.workDays), "天／月"],
    ["班別", n(state.shifts), "班／日"],
    ["每班工時", n(state.hoursPerShift), "hr"],
    ["每班休息", n(state.breakMinutes), "min"],
    ["計畫稼動率", n(state.availability) / 100, "%"],
    ["淨可用時間", calc.netAvailableDay, "min／日"],
    [],
    ["計算結果", "數值", "單位"],
    ["日需求", calc.dailyDemand, "pcs"],
    ["需求 Takt", calc.takt, "min／pcs"],
    ["瓶頸站", calc.bottleneck?.id || "—", calc.bottleneck?.name || ""],
    ["瓶頸月產能", calc.bottleneck?.capacity ?? null, "pcs"],
    ["瓶頸負荷率", calc.bottleneck?.load ?? null, "%"],
    ["總現況 WIP", calc.currentWip, "pcs"],
    ["總目標 WIP", calc.targetWip, "pcs"],
    ["總 VA 時間", calc.vaMinutes, "min"],
    ["總 NVA 等待", calc.nvaHours, "hr"],
    ["Lead Time", calc.leadTime, "hr"],
    ["產能判定", status, ""]
  ];

  const stationHeaders = ["順序", "站別", "製程名稱", "CT (min)", "MCT (min)", "OCTa (min)", "OCTb (min)", "C/O Changeover Time (min) / 換線／切換時間", "人數", "並行機台", "現況 WIP (pcs)", "緩衝 (min)", "良率", "月產能 (pcs)", "負荷率", "目標 WIP (pcs)", "現況等待 (hr)", "目標等待 (hr)", "狀態"];
  const stationRows = calc.stations.map((station, index) => [
    index + 1, station.id, station.name, n(station.ct), n(station.mct), n(station.octa), n(station.octb), n(station.octc),
    n(station.people), n(station.machines), n(station.currentWip), n(station.bufferMin), n(station.yieldRate) / 100,
    station.capacity, station.load, station.targetWip, station.currentWait, station.targetWait, station.status
  ]);
  const timelineRows = [["順序", "站別", "VA 加工時間 (min)", "站後 WIP (pcs)", "NVA 等待 (hr)"], ...calc.stations.map((station, index) => [index + 1, station.id, n(station.ct), n(station.currentWip), station.currentWait])];

  const workbook = XLSX.utils.book_new();
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  summarySheet["!cols"] = [{ wch: 21 }, { wch: 20 }, { wch: 14 }];
  setCellFormat(summarySheet, "B2", "yyyy-mm-dd hh:mm");
  setCellFormat(summarySheet, "B10", "0.0%");
  setCellFormat(summarySheet, "B18", "0.0%");
  ["B11", "B14", "B15", "B17", "B21", "B22", "B23"].forEach((cell) => setCellFormat(summarySheet, cell, "#,##0.0"));
  ["B5", "B6", "B7", "B19", "B20"].forEach((cell) => setCellFormat(summarySheet, cell, "#,##0"));

  const stationSheet = XLSX.utils.aoa_to_sheet([stationHeaders, ...stationRows]);
  stationSheet["!cols"] = stationHeaders.map((header, index) => ({ wch: index === 2 ? 18 : Math.max(10, Math.min(16, header.length + 3)) }));
  stationSheet["!autofilter"] = { ref: `A1:S${stationRows.length + 1}` };
  for (let row = 2; row <= stationRows.length + 1; row++) {
    setCellFormat(stationSheet, `M${row}`, "0.0%");
    setCellFormat(stationSheet, `O${row}`, "0.0%");
    ["D", "E", "F", "G", "H", "N", "Q", "R"].forEach((column) => setCellFormat(stationSheet, `${column}${row}`, "#,##0.0"));
  }

  const timelineSheet = XLSX.utils.aoa_to_sheet(timelineRows);
  timelineSheet["!cols"] = [{ wch: 9 }, { wch: 13 }, { wch: 20 }, { wch: 18 }, { wch: 18 }];
  for (let row = 2; row <= calc.stations.length + 1; row++) {
    setCellFormat(timelineSheet, `C${row}`, "#,##0.0");
    setCellFormat(timelineSheet, `E${row}`, "#,##0.0");
  }

  XLSX.utils.book_append_sheet(workbook, summarySheet, "生產摘要");
  XLSX.utils.book_append_sheet(workbook, stationSheet, "站別資料");
  XLSX.utils.book_append_sheet(workbook, timelineSheet, "價值流時間軸");
  try {
    XLSX.writeFile(workbook, `VSM_${safeFileName(state.lineName)}_${dateStamp()}.xlsx`, { compression: true });
    setExportStatus("Excel 已開始下載。");
  } catch (error) {
    console.error("Excel export failed", error);
    setExportStatus("Excel 匯出失敗，請重新整理後再試。", true);
  }
}

document.getElementById("addStationButton").addEventListener("click", () => {
  if (state.stations.length >= MAX_STATIONS) return;
  const usedIds = new Set(state.stations.map((station) => String(station.id || "").trim().toLocaleLowerCase("zh-TW")));
  let next = 1;
  while (usedIds.has(`op${next * 10}`)) next += 1;
  state.stations.push({ id: `OP${next * 10}`, name: "新製程", ct: 1, mct: 1, octa: 0, octb: 0, octc: 0, people: 1, machines: 1, currentWip: 0, bufferMin: 60, yieldRate: 100 });
  renderAll();
  document.querySelector(".station-editor:last-child")?.scrollIntoView({ behavior: "smooth", block: "center" });
});

document.getElementById("resetButton").addEventListener("click", () => {
  state = structuredClone(sampleState);
  renderAll();
});

document.getElementById("exportExcelButton").addEventListener("click", exportExcel);
document.getElementById("exportPdfButton").addEventListener("click", exportPdf);
const stationImportFile = document.getElementById("stationImportFile");
document.getElementById("importExcelButton").addEventListener("click", () => stationImportFile.click());
stationImportFile.addEventListener("change", async () => {
  const [file] = stationImportFile.files || [];
  await importStationExcel(file);
  stationImportFile.value = "";
});

function registerWebMcp() {
  const context = typeof document === "undefined" ? undefined : document.modelContext;
  if (!context?.registerTool) return;
  const setInputs = {
    name: "set_vsm_inputs",
    title: "設定 VSM 生產條件",
    description: "更新月需求、工作天、班別、工時、休息時間與稼動率，並重新計算畫面上的 VSM 結果。",
    inputSchema: {
      type: "object",
      properties: {
        monthlyDemand: { type: "number", exclusiveMinimum: 0 },
        workDays: { type: "number", exclusiveMinimum: 0 },
        shifts: { type: "number", exclusiveMinimum: 0 },
        hoursPerShift: { type: "number", exclusiveMinimum: 0 },
        breakMinutes: { type: "number", minimum: 0 },
        availability: { type: "number", exclusiveMinimum: 0, maximum: 100 }
      },
      additionalProperties: false
    },
    annotations: { readOnlyHint: false, untrustedContentHint: false },
    execute(input) {
      if (!input || typeof input !== "object") throw new Error("輸入格式不正確");
      const nextState = { ...state };
      for (const [key, value] of Object.entries(input)) {
        if (!(key in state) || !Number.isFinite(value)) throw new Error(`無效欄位：${key}`);
        nextState[key] = value;
      }
      if (nextState.monthlyDemand <= 0 || nextState.workDays <= 0 || nextState.shifts <= 0 || nextState.hoursPerShift <= 0) throw new Error("需求與工時必須大於 0");
      if (nextState.availability <= 0 || nextState.availability > 100) throw new Error("計畫稼動率必須介於 0% 與 100% 之間");
      if (nextState.hoursPerShift * 60 <= nextState.breakMinutes) throw new Error("休息時間不可大於或等於每班工時");
      Object.assign(state, nextState);
      renderAll();
      const result = calculate();
      return { taktMinutes: result.takt, bottleneck: result.bottleneck?.id ?? null, leadTimeHours: result.leadTime };
    }
  };
  const readResults = {
    name: "read_vsm_results",
    title: "讀取 VSM 結果",
    description: "讀取目前畫面上的 Takt、瓶頸、月產能、WIP 與 Lead Time。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { readOnlyHint: true, untrustedContentHint: false },
    execute() {
      const result = calculate();
      return {
        taktMinutes: result.takt,
        bottleneck: result.bottleneck?.id ?? null,
        bottleneckMonthlyCapacity: result.bottleneck?.capacity ?? null,
        currentWip: result.currentWip,
        targetWip: result.targetWip,
        leadTimeHours: result.leadTime
      };
    }
  };
  try {
    void Promise.resolve(context.registerTool(setInputs)).catch(() => {});
    void Promise.resolve(context.registerTool(readResults)).catch(() => {});
  } catch (_) {}
}

bindGlobals();
renderAll();
registerWebMcp();
