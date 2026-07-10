const DB_NAME = "voz-db-v1";
const DB_VERSION = 1;
const STORE_NAME = "state";
const LAST_NUMBER_KEY = "voz-last-number-used";
const CHECKLIST_TEMPLATE_URL = "./checklist-template.xlsx";

const defectOptions = [
  "Extintor caducado.",
  "Hay un obstáculo.",
  "Extintor descargado.",
  "Extintor sin presión.",
  "Extintor en el suelo.",
  "Cristal armario roto o sin cristal.",
  "Sin señal.",
  "Señal caducada.",
  "Extintor en mal estado.",
];

const reportDefectLabels = {
  "Cristal armario roto o sin cristal.": "Cristal del extintor ausente o roto.",
};

const fields = [
  "cliente",
  "edificio",
  "cantidad",
  "ubicacion",
  "modelo",
  "numeroSerie",
  "fechaFabricacion",
  "fechaProximoRetimbrado",
  "observaciones",
  "senal",
];

let records = [];
let currentPhotos = ["", ""];
let voiceRecognition = null;
let voiceStep = "numero";
let voiceActive = false;

const $ = (id) => document.getElementById(id);

function safeText(value) {
  return value === undefined || value === null ? "" : String(value);
}

function createId() {
  return `rec-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeDefects(defects) {
  const normalized = [];
  for (const defect of Array.isArray(defects) ? defects : []) {
    const cleanDefect = normalizeDefectName(defect);
    if (cleanDefect && !normalized.includes(cleanDefect)) normalized.push(cleanDefect);
  }
  return normalized;
}

function cleanRecord(record = {}) {
  return {
    id: record.id || createId(),
    cliente: safeText(record.cliente),
    edificio: safeText(record.edificio ?? record.edificioCodigo),
    cantidad: safeText(record.cantidad),
    ubicacion: safeText(record.ubicacion),
    modelo: safeText(record.modelo),
    numeroSerie: safeText(record.numeroSerie),
    fechaFabricacion: safeText(record.fechaFabricacion),
    fechaProximoRetimbrado: safeText(record.fechaProximoRetimbrado),
    observaciones: safeText(record.observaciones),
    senal: safeText(record.senal),
    defectos: normalizeDefects(record.defectos),
    photos: Array.isArray(record.photos) ? [safeText(record.photos[0]), safeText(record.photos[1])] : ["", ""],
    visto: Boolean(record.visto),
    origen: record.origen || "excel",
  };
}

function normalizeKeyPart(value) {
  return safeText(value).trim().toLowerCase().replace(/\s+/g, " ");
}

function recordKey(record) {
  return [
    normalizeKeyPart(record.cantidad),
    normalizeKeyPart(record.numeroSerie),
    normalizeKeyPart(record.ubicacion),
  ].join("|");
}

function excelCellToText(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-ES");
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return excelCellToText(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
  }
  return String(value);
}

function normalizeHeader(value) {
  return safeText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[º°]/g, "o")
    .replace(/[^a-z0-9]/g, "");
}

function buildHeaderMap(rowValues) {
  const map = {};
  for (let col = 1; col < rowValues.length; col += 1) {
    const key = normalizeHeader(excelCellToText(rowValues[col]));
    if (key) map[key] = col;
  }
  return map;
}

function importedValue(rowValues, headerMap, keys, fallbackCol) {
  const col = keys.map((key) => headerMap[key]).find(Boolean) || fallbackCol;
  return excelCellToText(rowValues[col]);
}

function truthyExcelValue(value) {
  const text = normalizeHeader(excelCellToText(value));
  return ["si", "x", "true", "1", "ok", "visto"].includes(text);
}

function normalizeDefectName(value) {
  const text = safeText(value).trim();
  if (!text || text === "-") return "";
  const aliases = new Map([
    ["cristaldelextintorausenteoroto", "Cristal armario roto o sin cristal."],
    ["cristalarmariorotoosincristal", "Cristal armario roto o sin cristal."],
    ...defectOptions.map((defect) => [normalizeHeader(defect), defect]),
  ]);
  return aliases.get(normalizeHeader(text)) || text;
}

function splitImportedDefects(value) {
  return safeText(value)
    .split(/\s*\/\s*|\r?\n|;/)
    .map(normalizeDefectName)
    .filter(Boolean);
}

function importedDefects(rowValues, headerMap) {
  const defects = splitImportedDefects(importedValue(rowValues, headerMap, ["defectosencontrados", "defectos"], 0));
  const flagColumns = [
    ["extintorcaducado", "Extintor caducado."],
    ["hayunobstaculo", "Hay un obstáculo."],
    ["extintordescargado", "Extintor descargado."],
    ["extintorsinpresion", "Extintor sin presión."],
    ["extintorenelsuelo", "Extintor en el suelo."],
    ["cristalarmariorotoosincristal", "Cristal armario roto o sin cristal."],
    ["cristaldelextintorausenteoroto", "Cristal armario roto o sin cristal."],
    ["sinsenal", "Sin señal."],
    ["senalcaducada", "Señal caducada."],
    ["extintorenmalestado", "Extintor en mal estado."],
  ];
  for (const [key, defect] of flagColumns) {
    const col = headerMap[key];
    if (col && truthyExcelValue(rowValues[col])) defects.push(defect);
  }
  return normalizeDefects(defects);
}

function rowToImportedRecord(rowValues, index, headerMap = {}) {
  return cleanRecord({
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    cliente: importedValue(rowValues, headerMap, ["cliente"], 1),
    edificio: importedValue(rowValues, headerMap, ["edificio"], 2),
    cantidad: importedValue(rowValues, headerMap, ["numerosyco", "numero", "num"], 3),
    ubicacion: importedValue(rowValues, headerMap, ["ubicacion"], 4),
    modelo: importedValue(rowValues, headerMap, ["modelo"], 5),
    numeroSerie: importedValue(rowValues, headerMap, ["noserie", "numeroserie", "serie"], 6),
    fechaFabricacion: importedValue(rowValues, headerMap, ["fechaanofabricacion", "fechafabricacion", "fabricacion"], 7),
    fechaProximoRetimbrado: importedValue(rowValues, headerMap, ["fecharetimbrado", "retimbrado"], 8),
    observaciones: importedValue(rowValues, headerMap, ["observaciones", "observacion"], 9),
    senal: importedValue(rowValues, headerMap, ["senal"], 10),
    defectos: importedDefects(rowValues, headerMap),
    visto: truthyExcelValue(importedValue(rowValues, headerMap, ["visto"], 0)),
    origen: "importado",
  });
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readState() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get("records");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function writeState(value) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(value, "records");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadRecords() {
  try {
    const saved = await readState();
    if (Array.isArray(saved)) {
      records = saved.map(cleanRecord);
      return;
    }
  } catch {}
  records = (window.INITIAL_EXTINTORES_LISTADOS || []).map(cleanRecord);
  await saveRecords();
}

async function saveRecords() {
  records = records.map(cleanRecord);
  updateStats();
  await writeState(records);
}

function updateStats() {
  const total = records.length;
  const seen = records.filter((record) => record.visto).length;
  $("totalCount").textContent = total;
  $("seenCount").textContent = seen;
  $("pendingCount").textContent = total - seen;
}

function showView(name) {
  $("homeView").classList.toggle("hidden", name !== "home");
  $("listView").classList.toggle("hidden", name !== "list");
  $("correctivosView").classList.toggle("hidden", name !== "correctivos");
  $("checklistView").classList.toggle("hidden", name !== "checklist");
  $("formView").classList.toggle("hidden", name !== "form");
  if (name === "list") renderTable();
  if (name === "correctivos") renderCorrectivos();
  if (name === "checklist") renderChecklist();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function compareText(a, b) {
  return safeText(a).localeCompare(safeText(b), "es", { numeric: true, sensitivity: "base" });
}

function filteredRecords() {
  const filterEdificio = $("filterEdificio").value.trim().toLowerCase();
  const filterNumero = $("filterNumero").value.trim().toLowerCase();
  const filterSerie = $("filterSerie").value.trim().toLowerCase();
  const seenFilter = $("seenFilter").value;
  const sortOrder = $("sortOrder").value;

  const rows = records.filter((record) => {
    if (seenFilter === "seen" && !record.visto) return false;
    if (seenFilter === "pending" && record.visto) return false;
    if (filterEdificio && ![record.edificio, record.ubicacion].join(" ").toLowerCase().includes(filterEdificio)) return false;
    if (filterNumero && !safeText(record.cantidad).toLowerCase().includes(filterNumero)) return false;
    if (filterSerie && !safeText(record.numeroSerie).toLowerCase().includes(filterSerie)) return false;
    return true;
  });

  if (sortOrder === "edificio") rows.sort((a, b) => compareText(a.edificio, b.edificio) || compareText(a.cantidad, b.cantidad));
  if (sortOrder === "numero") rows.sort((a, b) => compareText(a.cantidad, b.cantidad) || compareText(a.edificio, b.edificio));
  return rows;
}

function renderTable() {
  const body = $("recordsBody");
  const rows = filteredRecords();
  body.innerHTML = "";
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="15">No hay registros con ese filtro.</td></tr>`;
    return;
  }
  for (const record of rows) {
    const defects = record.defectos.length ? record.defectos.join(" / ") : "-";
    const photo1 = record.photos[0] ? `<img class="tablePhoto" src="${record.photos[0]}" alt="Foto 1">` : `<span class="noPhoto">—</span>`;
    const photo2 = record.photos[1] ? `<img class="tablePhoto" src="${record.photos[1]}" alt="Foto 2">` : `<span class="noPhoto">—</span>`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${safeText(record.cliente) || "-"}</td>
      <td>${safeText(record.edificio) || "-"}</td>
      <td><strong>${safeText(record.cantidad) || "-"}</strong></td>
      <td>${safeText(record.ubicacion) || "-"}</td>
      <td>${safeText(record.modelo) || "-"}</td>
      <td>${safeText(record.numeroSerie) || "-"}</td>
      <td>${safeText(record.fechaFabricacion) || "-"}</td>
      <td>${safeText(record.fechaProximoRetimbrado) || "-"}</td>
      <td>${safeText(record.observaciones) || "-"}</td>
      <td>${safeText(record.senal) || "-"}</td>
      <td>${defects}</td>
      <td>${photo1}</td>
      <td>${photo2}</td>
      <td><span class="${record.visto ? "ok" : "pending"}">${record.visto ? "Sí" : "No"}</span></td>
      <td><button class="editBtn" data-edit="${record.id}">Ver / corregir</button></td>
    `;
    body.appendChild(tr);
  }
  body.querySelectorAll("[data-edit]").forEach((button) => {
    button.addEventListener("click", () => openForm(button.dataset.edit));
  });
}

function reportDefectText(defect) {
  return reportDefectLabels[defect] || defect;
}

function reportDefectKey(record) {
  return (record.defectos || []).map(reportDefectText).join(" / ");
}

function recordsWithCorrectivos() {
  return records
    .filter((record) => Array.isArray(record.defectos) && record.defectos.length)
    .map(cleanRecord)
    .sort((a, b) =>
      compareText(a.cliente, b.cliente) ||
      compareText(reportDefectKey(a), reportDefectKey(b)) ||
      compareText(a.edificio, b.edificio) ||
      compareText(a.cantidad, b.cantidad)
    );
}

function buildCorrectivosAnalysis() {
  const correctivos = recordsWithCorrectivos();
  const clients = new Map();
  for (const record of correctivos) {
    const clientName = safeText(record.cliente).trim() || "Correctivos";
    if (!clients.has(clientName)) clients.set(clientName, new Map());
    const groups = clients.get(clientName);
    const key = reportDefectKey(record);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(record);
  }
  return { correctivos, clients };
}

function countCorrectivosGroups(clients) {
  return Array.from(clients.values()).reduce((total, groups) => total + groups.size, 0);
}

function renderCorrectivos() {
  const { correctivos, clients } = buildCorrectivosAnalysis();
  const groupsCount = countCorrectivosGroups(clients);
  $("correctivosCount").textContent = correctivos.length;
  $("correctivosGroupsCount").textContent = groupsCount;
  $("correctivosClientsCount").textContent = clients.size;
  $("correctivosSummary").textContent = correctivos.length
    ? `${correctivos.length} registros con defectos agrupados en ${groupsCount} bloques.`
    : "No hay defectos registrados.";
  $("downloadCorrectivosWordBtn").disabled = !correctivos.length;

  const preview = $("correctivosPreview");
  if (!correctivos.length) {
    preview.innerHTML = `<p class="emptyReport">Cuando marques defectos en registros, aparecerá aquí el estudio de correctivos.</p>`;
    return;
  }

  preview.innerHTML = Array.from(clients, ([clientName, groups]) => {
    const groupRows = Array.from(groups, ([defects, rows]) => `
      <div class="previewGroup">
        <strong>${escapeHtml(defects)}</strong>
        <span>${rows.length} registros</span>
      </div>
    `).join("");
    return `
      <div class="previewClient">
        <h3>${escapeHtml(clientName)}</h3>
        ${groupRows}
      </div>
    `;
  }).join("");
}

function escapeHtml(value) {
  return safeText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function wordCell(value) {
  return escapeHtml(value || "");
}

function correctivosWordHtml() {
  const { correctivos, clients } = buildCorrectivosAnalysis();
  const date = new Date().toLocaleDateString("es-ES");
  const body = Array.from(clients, ([clientName, groups], clientIndex) => {
    const blocks = Array.from(groups, ([defects, rows]) => {
      const tableRows = rows.map((record) => `
        <tr>
          <td>${wordCell(record.edificio)}</td>
          <td>${wordCell(record.cantidad)}</td>
          <td>${wordCell(record.numeroSerie)}</td>
          <td>${wordCell(record.modelo)}</td>
        </tr>
      `).join("");
      return `
        <h2>${escapeHtml(defects)}</h2>
        <table>
          <thead>
            <tr>
              <th>Edificio</th>
              <th>Número</th>
              <th>Nº serie</th>
              <th>Modelo</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
      `;
    }).join("");
    return `
      <section class="${clientIndex ? "pageBreak" : ""}">
        <h1>Listado de incidencias de extintores</h1>
        <p class="client">${escapeHtml(clientName)}</p>
        ${blocks}
      </section>
    `;
  }).join("");

  return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8">
  <title>Listado de incidencias de extintores</title>
  <style>
    @page { size: 8.5in 11in; margin: .55in .55in .55in .55in; }
    body { font-family: Arial, Helvetica, sans-serif; color: #000; font-size: 10.5pt; }
    h1 { text-align: center; font-size: 20pt; margin: 18pt 0 18pt; font-weight: 700; }
    .client { text-align: center; font-size: 16pt; font-weight: 700; margin: 0 0 18pt; }
    .meta { color: #555; font-size: 9pt; text-align: right; margin: 0 0 10pt; }
    h2 { color: #2f70c0; font-size: 14pt; margin: 22pt 0 6pt .2in; font-weight: 700; }
    table { border-collapse: collapse; width: 100%; margin: 0 0 22pt; table-layout: fixed; }
    th, td { border: 1px solid #000; padding: 2pt 5pt; vertical-align: top; line-height: 1.05; }
    th { background: #d9eaf7; text-align: left; font-weight: 700; }
    th:nth-child(1), td:nth-child(1) { width: 25%; }
    th:nth-child(2), td:nth-child(2) { width: 25%; }
    th:nth-child(3), td:nth-child(3) { width: 25%; }
    th:nth-child(4), td:nth-child(4) { width: 25%; }
    .pageBreak { page-break-before: always; }
  </style>
</head>
<body>
  <p class="meta">Generado el ${escapeHtml(date)}. Registros correctivos: ${correctivos.length}.</p>
  ${body}
</body>
</html>`;
}

function downloadCorrectivosWord() {
  const { correctivos } = buildCorrectivosAnalysis();
  if (!correctivos.length) {
    alert("No hay registros con defectos para generar el Word de correctivos.");
    return;
  }
  const blob = new Blob(["\ufeff", correctivosWordHtml()], { type: "application/msword;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Correctivos_extintores_${new Date().toISOString().slice(0, 10)}.doc`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function checklistRecords() {
  return records
    .map(cleanRecord)
    .filter((record) => [record.cliente, record.edificio, record.cantidad, record.modelo, record.numeroSerie].some((value) => safeText(value).trim()))
    .sort((a, b) =>
      compareText(a.cliente, b.cliente) ||
      compareText(a.edificio, b.edificio) ||
      compareText(a.cantidad, b.cantidad)
    );
}

function checklistGroupKey(record) {
  return `${normalizeKeyPart(record.cliente)}|${normalizeKeyPart(record.edificio)}`;
}

function checklistHeader(record) {
  const cliente = safeText(record.cliente).trim();
  const edificio = safeText(record.edificio).trim();
  return [cliente, edificio].filter(Boolean).join(". ");
}

function buildChecklistGroups() {
  const groups = new Map();
  for (const record of checklistRecords()) {
    const key = checklistGroupKey(record);
    if (!groups.has(key)) groups.set(key, { title: checklistHeader(record) || "Checklist", records: [] });
    groups.get(key).records.push(record);
  }
  return Array.from(groups.values());
}

function checklistPageCount(groups = buildChecklistGroups()) {
  return groups.reduce((total, group) => total + Math.max(1, Math.ceil(group.records.length / 20)), 0);
}

function renderChecklist() {
  const groups = buildChecklistGroups();
  const totalRecords = groups.reduce((total, group) => total + group.records.length, 0);
  $("checklistCount").textContent = totalRecords;
  $("checklistPagesCount").textContent = totalRecords ? checklistPageCount(groups) : 0;
  $("checklistClientsCount").textContent = groups.length;
  $("checklistSummary").textContent = totalRecords
    ? `${totalRecords} registros preparados en ${groups.length} grupos de cliente y edificio.`
    : "No hay registros para rellenar.";
  $("downloadChecklistBtn").disabled = !totalRecords;
}

function parseYear(value) {
  const match = safeText(value).match(/\b(19|20)\d{2}\b/);
  if (match) return Number(match[0]);
  const shortYear = safeText(value).match(/\b\d{2}\b/);
  if (shortYear) return 2000 + Number(shortYear[0]);
  return null;
}

function checklistOperation(record) {
  const text = normalizeSpeechText([record.observaciones, ...(record.defectos || [])].join(" "));
  return text.includes("recarg") || text.includes("retimbrad") ? "Revisión y recarga" : "Revisión";
}

function copyCell(sourceCell, targetCell) {
  targetCell.value = sourceCell.value;
  targetCell.style = JSON.parse(JSON.stringify(sourceCell.style || {}));
  if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
}

function copyTemplateSheet(workbook, sourceSheet, name) {
  const sheet = workbook.addWorksheet(name);
  sheet.pageSetup = JSON.parse(JSON.stringify(sourceSheet.pageSetup || {}));
  sheet.pageMargins = JSON.parse(JSON.stringify(sourceSheet.pageMargins || {}));
  sheet.headerFooter = JSON.parse(JSON.stringify(sourceSheet.headerFooter || {}));
  sheet.views = JSON.parse(JSON.stringify(sourceSheet.views || []));
  sheet.properties = JSON.parse(JSON.stringify(sourceSheet.properties || {}));

  for (let col = 1; col <= sourceSheet.columnCount; col += 1) {
    const sourceCol = sourceSheet.getColumn(col);
    const targetCol = sheet.getColumn(col);
    targetCol.width = sourceCol.width;
    targetCol.hidden = sourceCol.hidden;
    targetCol.outlineLevel = sourceCol.outlineLevel;
  }

  for (let rowNumber = 1; rowNumber <= sourceSheet.rowCount; rowNumber += 1) {
    const sourceRow = sourceSheet.getRow(rowNumber);
    const targetRow = sheet.getRow(rowNumber);
    targetRow.height = sourceRow.height;
    for (let col = 1; col <= sourceSheet.columnCount; col += 1) {
      copyCell(sourceRow.getCell(col), targetRow.getCell(col));
    }
  }

  (sourceSheet.model?.merges || []).forEach((range) => sheet.mergeCells(range));
  return sheet;
}

function clearChecklistRows(sheet) {
  for (let row = 8; row <= 27; row += 1) {
    for (let col = 1; col <= 8; col += 1) sheet.getRow(row).getCell(col).value = "";
    sheet.getRow(row).getCell(38).value = "";
    sheet.getRow(row).getCell(39).value = "";
    sheet.getRow(row).getCell(40).value = "";
  }
}

function checklistDefectsText(record) {
  return Array.isArray(record.defectos) && record.defectos.length
    ? record.defectos.map(reportDefectText).join(" / ")
    : "";
}

function clearChecklistSystemObservations(sheet) {
  sheet.getCell("A50").value = "";
  sheet.getCell("A51").value = "";
  sheet.getCell("A52").value = "";
  sheet.getCell("E50").value = "";
  sheet.getCell("E51").value = "";
}

function prepareChecklistObservationCells(sheet, rowNumber, text) {
  const row = sheet.getRow(rowNumber);
  const observation = row.getCell(39);
  const extraObservation = row.getCell(40);
  if (!observation.isMerged && !extraObservation.isMerged) sheet.mergeCells(rowNumber, 39, rowNumber, 40);
  observation.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
  observation.font = { ...(observation.font || {}), size: 8 };
  if (text) {
    row.height = Math.max(row.height || 20.25, text.length > 70 ? 42 : text.length > 38 ? 32 : 24);
  }
}

function fillChecklistSheet(sheet, title, rows) {
  sheet.getCell("C2").value = title;
  clearChecklistRows(sheet);
  clearChecklistSystemObservations(sheet);
  rows.forEach((record, index) => {
    const row = sheet.getRow(8 + index);
    const rowNumber = 8 + index;
    const fabricationYear = parseYear(record.fechaFabricacion);
    const defectsText = checklistDefectsText(record);
    row.getCell(1).value = safeText(record.cantidad);
    row.getCell(2).value = safeText(record.numeroSerie);
    row.getCell(3).value = safeText(record.modelo);
    row.getCell(4).value = "";
    row.getCell(5).value = safeText(record.fechaFabricacion);
    row.getCell(6).value = safeText(record.fechaProximoRetimbrado);
    row.getCell(7).value = checklistOperation(record);
    row.getCell(8).value = fabricationYear ? String(fabricationYear + 20) : "";
    row.getCell(38).value = safeText(record.ubicacion);
    row.getCell(39).value = defectsText;
    prepareChecklistObservationCells(sheet, rowNumber, defectsText);
  });
}

function safeSheetName(value, index) {
  const name = safeText(value).replace(/[\\/*?:[\]]/g, " ").replace(/\s+/g, " ").trim() || `Checklist ${index}`;
  return name.slice(0, 25) + (index > 1 ? ` ${index}` : "");
}

async function downloadChecklist() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const groups = buildChecklistGroups();
  if (!groups.length) return alert("No hay registros para generar el checklist.");

  const response = await fetch(CHECKLIST_TEMPLATE_URL);
  if (!response.ok) return alert("No se ha podido cargar la plantilla del checklist.");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const templateSheet = workbook.worksheets[0];
  const pages = [];

  groups.forEach((group) => {
    for (let start = 0; start < group.records.length; start += 20) {
      pages.push({ title: group.title, records: group.records.slice(start, start + 20) });
    }
  });

  const sheets = pages.map((page, index) => {
    const sheet = index === 0 ? templateSheet : copyTemplateSheet(workbook, templateSheet, safeSheetName(page.title, index + 1));
    sheet.name = safeSheetName(page.title, index + 1);
    return { sheet, page };
  });

  sheets.forEach(({ sheet, page }) => fillChecklistSheet(sheet, page.title, page.records));

  const blob = new Blob([await workbook.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Checklist_extintores_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function renderDefects(selected = []) {
  const box = $("defectsList");
  box.innerHTML = "";
  for (const option of defectOptions) {
    const label = document.createElement("label");
    label.className = "checkItem";
    label.innerHTML = `<input type="checkbox" value="${option}"><span>${option}</span>`;
    label.querySelector("input").checked = selected.includes(option);
    box.appendChild(label);
  }
}

function setPhotoPreview(index, dataUrl) {
  const photo = safeText(dataUrl);
  const img = $(`photoPreview${index + 1}`);
  const text = $(`photoBox${index + 1}`).querySelector("span");
  currentPhotos[index] = photo;
  img.src = photo;
  img.classList.toggle("hidden", !photo);
  text.classList.toggle("hidden", Boolean(photo));
  $(`deletePhoto${index + 1}`).disabled = !photo;
}

function updateLastNumberUsed(currentId = "") {
  const savedNumber = localStorage.getItem(LAST_NUMBER_KEY);
  const lastRecord = records.find((record) => record.id !== currentId && safeText(record.cantidad).trim());
  const value = safeText(savedNumber).trim() || safeText(lastRecord?.cantidad).trim() || "-";
  $("lastNumberUsed").textContent = value;
}

function normalizeSpeechText(text) {
  return safeText(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,;:!?¿¡]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function spokenDigit(token) {
  const digits = {
    cero: "0",
    uno: "1",
    un: "1",
    una: "1",
    dos: "2",
    tres: "3",
    cuatro: "4",
    cinco: "5",
    seis: "6",
    siete: "7",
    ocho: "8",
    nueve: "9",
  };
  return digits[token] || (/^\d+$/.test(token) ? token : "");
}

function speechToNumberValue(text) {
  const normalized = normalizeSpeechText(text);
  const numericParts = normalized.match(/\d+/g);
  if (numericParts?.length) return numericParts.join("");

  const tokens = normalized.split(" ").filter((token) => token && token !== "y");
  const compactDigits = tokens.map(spokenDigit).join("");
  if (compactDigits && tokens.every((token) => spokenDigit(token))) return compactDigits;

  const values = {
    cero: 0,
    uno: 1,
    un: 1,
    una: 1,
    dos: 2,
    tres: 3,
    cuatro: 4,
    cinco: 5,
    seis: 6,
    siete: 7,
    ocho: 8,
    nueve: 9,
    diez: 10,
    once: 11,
    doce: 12,
    trece: 13,
    catorce: 14,
    quince: 15,
    dieciseis: 16,
    diecisiete: 17,
    dieciocho: 18,
    diecinueve: 19,
    veinte: 20,
    veintiuno: 21,
    veintidos: 22,
    veintitres: 23,
    veinticuatro: 24,
    veinticinco: 25,
    veintiseis: 26,
    veintisiete: 27,
    veintiocho: 28,
    veintinueve: 29,
    treinta: 30,
    cuarenta: 40,
    cincuenta: 50,
    sesenta: 60,
    setenta: 70,
    ochenta: 80,
    noventa: 90,
    cien: 100,
    ciento: 100,
    doscientos: 200,
    trescientos: 300,
    cuatrocientos: 400,
    quinientos: 500,
    seiscientos: 600,
    setecientos: 700,
    ochocientos: 800,
    novecientos: 900,
  };

  let total = 0;
  let current = 0;
  let found = false;
  for (const token of tokens) {
    if (token === "mil") {
      total += (current || 1) * 1000;
      current = 0;
      found = true;
      continue;
    }
    if (values[token] === undefined) continue;
    current += values[token];
    found = true;
  }

  return found ? String(total + current) : speechToPlainValue(text).replace(/\s+/g, "");
}

function speechToSerial(text) {
  const tokens = normalizeSpeechText(text).split(" ").filter(Boolean);
  const parts = tokens.map((token) => spokenDigit(token) || token.toUpperCase());
  return parts.join("").replace(/[^A-Z0-9-]/g, "");
}

function speechToPlainValue(text) {
  const digitWords = {
    cero: "0",
    uno: "1",
    un: "1",
    una: "1",
    dos: "2",
    tres: "3",
    cuatro: "4",
    cinco: "5",
    seis: "6",
    siete: "7",
    ocho: "8",
    nueve: "9",
  };
  return normalizeSpeechText(text)
    .split(" ")
    .filter(Boolean)
    .map((token) => digitWords[token] || token.toUpperCase())
    .join(" ")
    .replace(/\bKG\b/g, "KG")
    .trim();
}

function speechToModel(text) {
  const normalized = normalizeSpeechText(text);
  const compactDigits = normalized
    .split(" ")
    .map(spokenDigit)
    .join("");
  const numeric = normalized.match(/\b\d+\b/)?.[0] || compactDigits || speechToNumberValue(text);
  const models = {
    1: "ABC 1 KG",
    2: "CO2 2 KG",
    3: "ABC 3 KG",
    5: "CO2 5 KG",
    6: "ABC 6 KG",
    9: "ABC 9 KG",
    10: "CO2 10 KG",
    25: "ABC 25 KG",
    50: "ABC 50 KG",
  };
  return models[numeric] || speechToPlainValue(text);
}

function speechToYear(text) {
  const normalized = normalizeSpeechText(text);
  const numeric = normalized.match(/\b(19|20)\d{2}\b/);
  if (numeric) return numeric[0];

  const parsedNumber = Number(speechToNumberValue(text));
  if (Number.isFinite(parsedNumber)) {
    if (parsedNumber >= 1900 && parsedNumber <= 2099) return String(parsedNumber);
    if (parsedNumber >= 0 && parsedNumber <= 99) return String(2000 + parsedNumber);
  }

  const compactDigits = normalized
    .split(" ")
    .map(spokenDigit)
    .join("");
  if (/^(19|20)\d{2}$/.test(compactDigits)) return compactDigits;

  const yearWords = {
    diez: 2010,
    once: 2011,
    doce: 2012,
    trece: 2013,
    catorce: 2014,
    quince: 2015,
    dieciseis: 2016,
    diecisiete: 2017,
    dieciocho: 2018,
    diecinueve: 2019,
    veinte: 2020,
    veintiuno: 2021,
    veintidos: 2022,
    veintitres: 2023,
    veinticuatro: 2024,
    veinticinco: 2025,
    veintiseis: 2026,
  };
  for (const [word, year] of Object.entries(yearWords)) {
    if (normalized.includes(word)) return String(year);
  }
  return speechToPlainValue(text);
}

function setSelectValue(id, value) {
  const select = $(id);
  const cleanValue = safeText(value).trim();
  const option = Array.from(select.options).find((item) => item.value === cleanValue);
  if (option) {
    select.value = cleanValue;
    return;
  }
  if (/^\d{4}$/.test(cleanValue)) {
    select.add(new Option(cleanValue, cleanValue));
    select.value = cleanValue;
    return;
  }
  select.value = "";
}

function setVoiceStatus(message) {
  const status = $("voiceStatus");
  if (status) status.textContent = message;
}

function appendSerial(value) {
  const serial = speechToSerial(value);
  if (!serial) return;
  $("numeroSerie").value = `${$("numeroSerie").value}${serial}`.trim();
}

function captureAfterKeyword(text, keyword) {
  const normalized = normalizeSpeechText(text);
  const index = normalized.indexOf(keyword);
  if (index < 0) return "";
  return normalized.slice(index + keyword.length).trim();
}

function handleVoiceText(text) {
  const normalized = normalizeSpeechText(text);
  if (!normalized) return;

  if (voiceStep === "numero") {
    const numberValue = captureAfterKeyword(text, "numero");
    if (!numberValue && !normalized.includes("numero")) {
      setVoiceStatus('Di "numero" y despues el numero del extintor.');
      return;
    }
    if (numberValue) $("cantidad").value = speechToNumberValue(numberValue);
    voiceStep = "modelo";
    setVoiceStatus('Numero anotado. Ahora di "modelo" y el modelo del extintor.');
    return;
  }

  if (voiceStep === "modelo") {
    const modelValue = captureAfterKeyword(text, "modelo");
    if (!modelValue && !normalized.includes("modelo")) {
      setVoiceStatus('Di "modelo" y el modelo del extintor.');
      return;
    }
    if (modelValue) {
      $("modelo").value = speechToModel(modelValue);
      voiceStep = "serie";
      setVoiceStatus('Modelo anotado. Ahora di "serie" y dicta el numero de serie, numero a numero.');
    } else {
      voiceStep = "modeloValor";
      setVoiceStatus("Ahora di el modelo del extintor.");
    }
    return;
  }

  if (voiceStep === "modeloValor") {
    $("modelo").value = speechToModel(text);
    voiceStep = "serie";
    setVoiceStatus('Modelo anotado. Ahora di "serie" y dicta el numero de serie, numero a numero.');
    return;
  }

  if (voiceStep === "serie") {
    const serieValue = captureAfterKeyword(text, "serie");
    if (!serieValue && !normalized.includes("serie")) {
      setVoiceStatus('Di "serie" y despues el numero de serie.');
      return;
    }
    if (serieValue) {
      appendSerial(serieValue);
      voiceStep = "numeroSerie";
      setVoiceStatus('Numero de serie anotado. Cuando termines di "fabricacion" y la fecha.');
    } else {
      voiceStep = "numeroSerie";
      setVoiceStatus("Ahora dicta el numero de serie, numero a numero.");
    }
    return;
  }

  if (voiceStep === "numeroSerie") {
    const fabricationIndex = normalized.indexOf("fabricacion");
    if (fabricationIndex >= 0) {
      const before = normalized.slice(0, fabricationIndex).trim();
      const after = normalized.slice(fabricationIndex + "fabricacion".length).trim();
      if (before) appendSerial(before);
      if (after) setSelectValue("fechaFabricacion", speechToYear(after));
      voiceStep = after ? "retimbrado" : "fabricacionValor";
      setVoiceStatus(after ? 'Fabricacion anotada. Ahora di "retimbrado" y la fecha.' : "Ahora di la fecha de fabricacion.");
      return;
    }
    appendSerial(text);
    setVoiceStatus('Numero de serie anotado. Cuando termines di "fabricacion" y la fecha.');
    return;
  }

  if (voiceStep === "fabricacionValor") {
    setSelectValue("fechaFabricacion", speechToYear(text));
    voiceStep = "retimbrado";
    setVoiceStatus('Fabricacion anotada. Ahora di "retimbrado" y la fecha.');
    return;
  }

  if (voiceStep === "retimbrado") {
    const retimbradoValue = captureAfterKeyword(text, "retimbrado");
    if (!retimbradoValue && !normalized.includes("retimbrado")) {
      setVoiceStatus('Di "retimbrado" y la fecha de retimbrado.');
      return;
    }
    if (!retimbradoValue) {
      voiceStep = "retimbradoValor";
      setVoiceStatus("Ahora di la fecha de retimbrado.");
      return;
    }
    setSelectValue("fechaProximoRetimbrado", speechToYear(retimbradoValue));
    voiceStep = "completo";
    stopVoiceInput(false);
    setVoiceStatus("Datos de voz anotados. Revisa o completa manualmente y pulsa Guardar.");
    return;
  }

  if (voiceStep === "retimbradoValor") {
    setSelectValue("fechaProximoRetimbrado", speechToYear(text));
    voiceStep = "completo";
    stopVoiceInput(false);
    setVoiceStatus("Datos de voz anotados. Revisa o completa manualmente y pulsa Guardar.");
  }
}

function getSpeechRecognition() {
  return window.SpeechRecognition || window.webkitSpeechRecognition || null;
}

function startVoiceInput() {
  const SpeechRecognition = getSpeechRecognition();
  if (!SpeechRecognition) {
    setVoiceStatus("Este navegador no permite reconocimiento de voz. Prueba con Chrome o Edge.");
    return;
  }

  if (voiceRecognition) voiceRecognition.stop();
  voiceStep = "numero";
  voiceActive = true;
  voiceRecognition = new SpeechRecognition();
  voiceRecognition.lang = "es-ES";
  voiceRecognition.continuous = true;
  voiceRecognition.interimResults = false;
  voiceRecognition.onresult = (event) => {
    for (let index = event.resultIndex; index < event.results.length; index += 1) {
      if (event.results[index].isFinal) handleVoiceText(event.results[index][0].transcript);
    }
  };
  voiceRecognition.onerror = () => setVoiceStatus("No he podido escuchar bien. Puedes parar e iniciar voz de nuevo.");
  voiceRecognition.onend = () => {
    if (voiceActive && voiceStep !== "completo") {
      try {
        voiceRecognition.start();
      } catch {}
      return;
    }
    $("voiceStartBtn").disabled = false;
    $("voiceStopBtn").disabled = true;
    $("recordForm").classList.remove("voiceListening");
  };
  $("voiceStartBtn").disabled = true;
  $("voiceStopBtn").disabled = false;
  $("recordForm").classList.add("voiceListening");
  setVoiceStatus('Escuchando. Empieza diciendo "numero" y el dato.');
  voiceRecognition.start();
}

function stopVoiceInput(showMessage = true) {
  voiceActive = false;
  if (voiceRecognition) {
    try {
      voiceRecognition.stop();
    } catch {}
  }
  $("voiceStartBtn").disabled = false;
  $("voiceStopBtn").disabled = true;
  $("recordForm").classList.remove("voiceListening");
  if (showMessage) setVoiceStatus("Voz parada. Puedes revisar o completar los campos manualmente.");
}

function resizePhoto(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const maxSide = 1200;
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.72));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function openForm(id = null) {
  const record = id ? records.find((item) => item.id === id) : null;
  $("recordId").value = record?.id || "";
  updateLastNumberUsed(record?.id || "");
  $("formTitle").textContent = record ? "Ver y corregir extintor" : "Meter dato nuevo";
  $("formKicker").textContent = record ? "REGISTRO EXISTENTE" : "NUEVO REGISTRO";
  $("deleteBtn").classList.toggle("hidden", !record);
  for (const key of fields) $(key).value = safeText(record?.[key]);
  $("visto").checked = Boolean(record?.visto);
  renderDefects(record?.defectos || []);
  const photos = Array.isArray(record?.photos) ? record.photos : ["", ""];
  setPhotoPreview(0, photos[0]);
  setPhotoPreview(1, photos[1]);
  showView("form");
}

function collectForm() {
  const record = { id: $("recordId").value || createId(), origen: $("recordId").value ? "editado" : "manual" };
  for (const key of fields) record[key] = $(key).value.trim();
  record.defectos = Array.from($("defectsList").querySelectorAll("input:checked")).map((input) => input.value);
  record.photos = [currentPhotos[0] || "", currentPhotos[1] || ""];
  record.visto = $("visto").checked;
  return cleanRecord(record);
}

async function saveForm(event) {
  event.preventDefault();
  const record = collectForm();
  const isNewRecord = !$("recordId").value;
  const index = records.findIndex((item) => item.id === record.id);
  if (index >= 0) records[index] = record;
  else records.unshift(record);
  if (record.cantidad) localStorage.setItem(LAST_NUMBER_KEY, record.cantidad);
  await saveRecords();
  if (isNewRecord) openForm();
  else showView("list");
}

async function deleteCurrent() {
  const id = $("recordId").value;
  if (!id) return;
  if (!confirm("¿Seguro que quieres eliminar este registro?")) return;
  records = records.filter((record) => record.id !== id);
  await saveRecords();
  showView("list");
}

async function clearAllRecords() {
  if (!records.length) {
    alert("No hay registros para eliminar.");
    return;
  }
  if (!confirm("¿Seguro que quieres eliminar todos los registros guardados en este dispositivo?")) return;
  records = [];
  localStorage.removeItem(LAST_NUMBER_KEY);
  await saveRecords();
  renderTable();
  showView("home");
  alert("Registros eliminados. Ya puedes importar otro cliente.");
}

function mergeImportedIntoExisting(existing, imported) {
  let changed = false;
  for (const key of fields) {
    if (!safeText(existing[key]).trim() && safeText(imported[key]).trim()) {
      existing[key] = imported[key];
      changed = true;
    }
  }
  const mergedDefects = normalizeDefects([...(existing.defectos || []), ...(imported.defectos || [])]);
  if (mergedDefects.length !== (existing.defectos || []).length) {
    existing.defectos = mergedDefects;
    changed = true;
  }
  if (imported.visto && !existing.visto) {
    existing.visto = true;
    changed = true;
  }
  return changed;
}

async function importExcelFile(file) {
  if (!window.ExcelJS) return alert("No se ha cargado el lector de Excel.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) return alert("No encuentro ninguna hoja en ese Excel.");
  const imported = [];
  const knownRecords = new Map(records.map((record) => [recordKey(record), record]).filter(([key]) => key !== "||"));
  let skippedDuplicates = 0;
  let updatedDuplicates = 0;
  const headerMap = buildHeaderMap(sheet.getRow(1).values);
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = rowToImportedRecord(row.values, rowNumber, headerMap);
    const hasData = [record.edificio, record.cantidad, record.ubicacion, record.modelo, record.numeroSerie].some((value) => safeText(value).trim());
    if (!hasData) return;
    const key = recordKey(record);
    if (key !== "||" && knownRecords.has(key)) {
      skippedDuplicates += 1;
      if (mergeImportedIntoExisting(knownRecords.get(key), record)) updatedDuplicates += 1;
      return;
    }
    if (key !== "||") knownRecords.set(key, record);
    imported.push(record);
  });
  if (!imported.length && !updatedDuplicates) {
    const message = skippedDuplicates
      ? `No se importaron registros nuevos. Ya existían ${skippedDuplicates} registros.`
      : "No se encontraron registros para importar.";
    $("importStatus").textContent = message;
    return alert(message);
  }
  records = [...imported, ...records];
  await saveRecords();
  $("importStatus").textContent = `Importados ${imported.length} registros nuevos. Actualizados ${updatedDuplicates} existentes. Ya existían ${skippedDuplicates}.`;
  alert(`Importación correcta.\nRegistros nuevos importados: ${imported.length}\nRegistros existentes actualizados: ${updatedDuplicates}\nRegistros ya existentes: ${skippedDuplicates}`);
}

function defectFlag(selected, defect) {
  return selected.includes(defect) ? "Sí" : "";
}

async function downloadExcel() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Correctivos";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Extintores");
  const columns = [
    ["cliente", "Cliente", 22],
    ["edificio", "Edificio", 14],
    ["cantidad", "Número SYCo", 18],
    ["ubicacion", "Ubicación", 42],
    ["modelo", "Modelo", 20],
    ["numeroSerie", "Nº serie", 18],
    ["fechaFabricacion", "Fecha / año fabricación", 22],
    ["fechaProximoRetimbrado", "Fecha retimbrado", 20],
    ["observaciones", "Observaciones", 34],
    ["senal", "Señal", 14],
    ["defectos", "Defectos encontrados", 42],
    ["defectoCaducado", "Extintor caducado", 20],
    ["defectoObstaculo", "Hay un obstáculo", 20],
    ["defectoDescargado", "Extintor descargado", 22],
    ["defectoSinPresion", "Extintor sin presión", 22],
    ["defectoSuelo", "Extintor en el suelo", 22],
    ["defectoCristal", "Cristal armario roto o sin cristal", 32],
    ["defectoSinSenal", "Sin señal", 16],
    ["defectoSenalCaducada", "Señal caducada", 20],
    ["defectoMalEstado", "Extintor en mal estado", 24],
    ["foto1", "Foto 1", 22],
    ["foto2", "Foto 2", 22],
    ["visto", "Visto", 10],
  ];
  sheet.columns = columns.map(([key, header, width]) => ({ key, header, width }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FF3A1028" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9A8D4" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(1).height = 30;

  for (const record of filteredRecords()) {
    const selected = record.defectos || [];
    const row = sheet.addRow({
      ...record,
      defectos: selected.join(" / "),
      defectoCaducado: defectFlag(selected, "Extintor caducado."),
      defectoObstaculo: defectFlag(selected, "Hay un obstáculo."),
      defectoDescargado: defectFlag(selected, "Extintor descargado."),
      defectoSinPresion: defectFlag(selected, "Extintor sin presión."),
      defectoSuelo: defectFlag(selected, "Extintor en el suelo."),
      defectoCristal: defectFlag(selected, "Cristal armario roto o sin cristal."),
      defectoSinSenal: defectFlag(selected, "Sin señal."),
      defectoSenalCaducada: defectFlag(selected, "Señal caducada."),
      defectoMalEstado: defectFlag(selected, "Extintor en mal estado."),
      foto1: record.photos[0] ? "Foto 1" : "",
      foto2: record.photos[1] ? "Foto 2" : "",
      visto: record.visto ? "Sí" : "No",
    });
    if (record.photos[0] || record.photos[1]) row.height = 92;
    [0, 1].forEach((photoIndex) => {
      const photo = record.photos[photoIndex];
      if (!photo) return;
      const imageId = workbook.addImage({ base64: photo, extension: "jpeg" });
      const col = photoIndex === 0 ? 20 : 21;
      sheet.addImage(imageId, { tl: { col, row: row.number - 1 }, ext: { width: 120, height: 85 }, editAs: "oneCell" });
    });
  }
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE6E0DA" } },
        left: { style: "thin", color: { argb: "FFE6E0DA" } },
        bottom: { style: "thin", color: { argb: "FFE6E0DA" } },
        right: { style: "thin", color: { argb: "FFE6E0DA" } },
      };
      cell.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && rowNumber % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFAF8F5" } };
    });
  });
  const blob = new Blob([await workbook.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Correctivos_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function bindEvents() {
  $("openListBtn").addEventListener("click", () => showView("list"));
  $("openCorrectivosBtn").addEventListener("click", () => showView("correctivos"));
  $("openChecklistBtn").addEventListener("click", () => showView("checklist"));
  $("newRecordBtn").addEventListener("click", () => openForm());
  $("newRecordFromListBtn").addEventListener("click", () => openForm());
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
  $("downloadCorrectivosWordBtn").addEventListener("click", downloadCorrectivosWord);
  $("downloadChecklistBtn").addEventListener("click", downloadChecklist);
  $("clearRecordsBtn").addEventListener("click", clearAllRecords);
  $("viewTableFromFormBtn").addEventListener("click", () => showView("list"));
  $("voiceStartBtn").addEventListener("click", startVoiceInput);
  $("voiceStopBtn").addEventListener("click", () => stopVoiceInput());
  $("importExcelBtn").addEventListener("click", () => $("importExcelInput").click());
  $("importExcelInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      $("importStatus").textContent = "Importando Excel...";
      await importExcelFile(file);
      renderTable();
    } catch (error) {
      console.error(error);
      $("importStatus").textContent = "No se ha podido importar el Excel.";
      alert("No se ha podido importar el Excel. Revisa que tenga el mismo formato.");
    } finally {
      event.target.value = "";
    }
  });
  ["filterEdificio", "filterNumero", "filterSerie", "sortOrder", "seenFilter"].forEach((id) => {
    $(id).addEventListener("input", renderTable);
    $(id).addEventListener("change", renderTable);
  });
  $("recordForm").addEventListener("submit", saveForm);
  $("deleteBtn").addEventListener("click", deleteCurrent);
  [0, 1].forEach((index) => {
    $(`photoInput${index + 1}`).addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        setPhotoPreview(index, await resizePhoto(file));
      } catch {
        alert("No he podido cargar esa foto. Prueba con otra imagen.");
      } finally {
        event.target.value = "";
      }
    });
    $(`deletePhoto${index + 1}`).addEventListener("click", () => setPhotoPreview(index, ""));
  });
  document.querySelectorAll("[data-back]").forEach((button) => button.addEventListener("click", () => showView(button.dataset.back)));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./sw.js").catch(() => {}));
}

async function init() {
  await loadRecords();
  bindEvents();
  updateStats();
}

init().catch((error) => {
  console.error(error);
  const main = document.querySelector("main");
  if (!main) return;
  const box = document.createElement("div");
  box.className = "appError";
  box.textContent = "No se ha podido iniciar la app. Actualiza la página o vuelve a subir todos los archivos.";
  main.prepend(box);
});
