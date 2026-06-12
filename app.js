const PRODUCTS = [
  { article: "7041328", name: "Pot 9 cm", livrerCol: 1, venduCol: 2 },
  { article: "7041331", name: "Pot 13 cm", livrerCol: 4, venduCol: 5 },
  { article: "7042371", name: "Fraise", livrerCol: 7, venduCol: 8 },
  { article: "7042372", name: "Balconnières", livrerCol: 10, venduCol: 11 },
  { article: "7043183", name: "Ail/Asperge", livrerCol: 13, venduCol: 14 },
  { article: "7043184", name: "Concombre", livrerCol: 16, venduCol: 17 },
  { article: "77010798", name: "Basilic", livrerCol: 19, venduCol: 20 },
  { article: "77011496", name: "Framb Pim", livrerCol: 22, venduCol: 23 },
];

const DUMMY_DATE = "2000-01-01";
const RECAP_SHEET_NAME = "Récapitulatif";

const state = {
  canac: {
    name: "",
    buffer: null,
    workbook: null,
    stores: [],
  },
  royaumes: {
    name: "",
    workbook: null,
    totals: {},
  },
  stores: [],
  selectedStoreId: null,
  targetDate: todayStr(),
  lastValidation: null,
};

const $ = (id) => document.getElementById(id);

function todayStr() {
  const d = new Date();
  return [d.getFullYear(), String(d.getMonth() + 1).padStart(2, "0"), String(d.getDate()).padStart(2, "0")].join("-");
}

function normalizeSpaces(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeArticle(value) {
  return String(value ?? "").replace(/\D/g, "");
}

function normalizeDivision(value) {
  const m = String(value ?? "").match(/\d+/);
  return m ? String(parseInt(m[0], 10)) : "";
}

function colLetterToNumber(letters) {
  return String(letters || "")
    .toUpperCase()
    .split("")
    .reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function numberToColLetter(n) {
  let result = "";
  let value = Number(n) || 0;
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function toNumber(value) {
  if (value == null || value === "") return 0;
  const n = Number(String(value).replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function isBlankValue(value) {
  return normalizeSpaces(value) === "";
}

function formatDateParts(year, month, day) {
  return [String(year).padStart(4, "0"), String(month).padStart(2, "0"), String(day).padStart(2, "0")].join("-");
}

function addDaysToNormalizedDate(dateValue, days) {
  const normalized = normalizeDateValue(dateValue);
  if (!normalized) return "";

  const [year, month, day] = normalized.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  date.setUTCDate(date.getUTCDate() + days);
  return formatDateParts(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
}

function normalizedDateToJsDate(dateValue) {
  const normalized = normalizeDateValue(dateValue);
  if (!normalized) return null;

  const [year, month, day] = normalized.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function normalizeDateValue(value) {
  if (value == null || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDateParts(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return formatDateParts(parsed.y, parsed.m, parsed.d);
  }

  const raw = normalizeSpaces(value).replace(/\./g, "/").replace(/-/g, "/");

  let match = raw.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (match) {
    return formatDateParts(match[1], match[2], match[3]);
  }

  match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (match) {
    let year = match[3];
    if (year.length === 2) year = `20${year}`;
    return formatDateParts(year, match[2], match[1]);
  }

  return raw.length === 10 ? raw : "";
}

function showStatus(message) {
  $("status").textContent = message;
}

function setSelectedStore(id) {
  state.selectedStoreId = id;
  render();
}

function getSelectedStoreIndex() {
  return state.stores.findIndex((store) => store.id === state.selectedStoreId);
}

function currentStore() {
  return state.stores.find((store) => store.id === state.selectedStoreId) || null;
}

function createStore({ magNum, sheetName, source = "canac" }) {
  const values = {};
  PRODUCTS.forEach((product) => {
    values[product.article] = {
      livrer: "",
      vendu: "",
    };
  });

  return {
    id: `store-${source}-${magNum}-${Math.random().toString(36).slice(2, 8)}`,
    magNum: String(magNum),
    sheetName: sheetName || `Magasin ${magNum}`,
    source,
    removed: false,
    values,
  };
}

function syncFromRoyaumes() {
  for (const store of state.stores) {
    const totals = state.royaumes.totals[store.magNum] || {};
    PRODUCTS.forEach((product) => {
      store.values[product.article].vendu = totals[product.article] ? String(Math.round(totals[product.article])) : "";
    });
  }
}

function syncStoreListFromCanac() {
  if (!state.canac.stores.length) return;

  const existing = new Map(state.stores.map((store) => [store.magNum, store]));
  const nextStores = [];

  for (const info of state.canac.stores) {
    const match = existing.get(info.magNum);
    if (match) {
      match.sheetName = info.sheetName;
      match.source = "canac";
      nextStores.push(match);
      continue;
    }
    nextStores.push(createStore(info));
  }

  for (const store of state.stores) {
    if (!state.canac.stores.some((info) => info.magNum === store.magNum) && store.source === "manual") {
      nextStores.push(store);
    }
  }

  state.stores = nextStores;
  if (!state.selectedStoreId && state.stores[0]) {
    state.selectedStoreId = state.stores[0].id;
  }
}

function parseCanacWorkbook(workbook) {
  const stores = workbook.SheetNames.map((sheetName) => {
    const match = normalizeSpaces(sheetName).match(/^(?:magasin|mag)\s*0*(\d+)\b/i);
    if (!match) return null;
    return {
      magNum: String(parseInt(match[1], 10)),
      sheetName,
    };
  })
    .filter(Boolean)
    .sort((a, b) => Number(a.magNum) - Number(b.magNum));

  return { stores };
}

function parseRoyaumesWorkbook(workbook) {
  const sheet = workbook.Sheets.Data || workbook.Sheets.data || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return { error: "Feuille source introuvable." };

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });
  let headerRow = -1;

  for (let i = 0; i < Math.min(rows.length, 50); i++) {
    const row = (rows[i] || []).map((cell) => normalizeSpaces(cell).toLowerCase());
    if (row.includes("division") && row.includes("article")) {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    return { error: "Colonnes Division et Article introuvables." };
  }

  const headers = (rows[headerRow] || []).map((cell) => normalizeSpaces(cell).toLowerCase());
  const divCol = headers.indexOf("division");
  const artCol = headers.indexOf("article");
  const qtyCol = headers.findIndex((cell) => cell.includes("qt") || cell.includes("qte") || cell.includes("quant"));

  if (divCol === -1 || artCol === -1 || qtyCol === -1) {
    return { error: "Colonnes requises introuvables." };
  }

  const totals = {};

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const division = normalizeDivision(row[divCol]);
    const article = normalizeArticle(row[artCol]);
    const qty = toNumber(row[qtyCol]);

    if (!division || !article || qty <= 0) continue;
    if (!PRODUCTS.some((product) => product.article === article)) continue;

    totals[division] ||= {};
    totals[division][article] = (totals[division][article] || 0) + qty;
  }

  return { totals };
}

function cloneWorkbook(buffer) {
  return XLSX.read(buffer, {
    type: "array",
    cellStyles: true,
    cellDates: true,
    cellNF: true,
  });
}

function findRowWithDate(sheet, targetDate) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });

  let lastDataRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const value = rows[i]?.[0];
    if (value != null && String(value).trim() !== "") {
      lastDataRow = i;
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const value = rows[i]?.[0];
    if (normalizeDateValue(value) === targetDate) {
      return { row: i, isNew: false, lastDataRow };
    }
  }

  return { row: Math.max(lastDataRow + 1, 0), isNew: true, lastDataRow };
}

function ensureSheet(workbook, sheetName) {
  if (!workbook.Sheets[sheetName]) {
    workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([["Date"]]);
    workbook.SheetNames.push(sheetName);
  }
  return workbook.Sheets[sheetName];
}

const XML_NS = "http://www.w3.org/1999/xhtml";
const XML_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function dateToExcelSerial(dateStr) {
  const match = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const utc = Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  const excelEpoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - excelEpoch) / 86400000);
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, "application/xml");
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function getSheetDataEl(doc) {
  return doc.getElementsByTagNameNS(XML_SHEET_NS, "sheetData")[0];
}

function getRowEl(doc, rowNumber) {
  const rows = getSheetDataEl(doc)?.getElementsByTagNameNS(XML_SHEET_NS, "row") || [];
  for (const row of rows) {
    if (Number(row.getAttribute("r")) === rowNumber) return row;
  }
  return null;
}

function getCellEl(rowEl, cellRef) {
  if (!rowEl) return null;
  const cells = rowEl.getElementsByTagNameNS(XML_SHEET_NS, "c") || [];
  for (const cell of cells) {
    if (cell.getAttribute("r") === cellRef) return cell;
  }
  return null;
}

function getCellByColumn(rowEl, colLetter) {
  if (!rowEl) return null;
  const cells = rowEl.getElementsByTagNameNS(XML_SHEET_NS, "c") || [];
  const wanted = String(colLetter || "").toUpperCase();
  for (const cell of cells) {
    const match = String(cell.getAttribute("r") || "").match(/^([A-Z]+)/i);
    if (match && match[1].toUpperCase() === wanted) return cell;
  }
  return null;
}

function cloneCellForRow(doc, sourceCell, newCellRef) {
  const clone = sourceCell.cloneNode(true);
  clone.setAttribute("r", newCellRef);
  return clone;
}

function ensureRowXml(doc, rowNumber, templateRowEl = null) {
  const sheetData = getSheetDataEl(doc);
  let rowEl = getRowEl(doc, rowNumber);
  if (rowEl) return rowEl;

  rowEl = doc.createElementNS(XML_SHEET_NS, "row");
  rowEl.setAttribute("r", String(rowNumber));

  if (templateRowEl) {
    ["spans", "ht", "customHeight", "s", "customFormat", "hidden", "outlineLevel"].forEach((attr) => {
      if (templateRowEl.hasAttribute(attr)) {
        rowEl.setAttribute(attr, templateRowEl.getAttribute(attr));
      }
    });

    const templateCells = Array.from(templateRowEl.getElementsByTagNameNS(XML_SHEET_NS, "c"));
    templateCells.forEach((templateCell) => {
      const ref = templateCell.getAttribute("r") || "";
      const col = ref.match(/^[A-Z]+/i)?.[0];
      if (!col) return;
      const newRef = `${col}${rowNumber}`;
      const newCell = cloneCellForRow(doc, templateCell, newRef);
      rowEl.appendChild(newCell);
    });
  }

  const rows = Array.from(sheetData.getElementsByTagNameNS(XML_SHEET_NS, "row"));
  const insertBefore = rows.find((r) => Number(r.getAttribute("r")) > rowNumber);
  if (insertBefore) {
    sheetData.insertBefore(rowEl, insertBefore);
  } else {
    sheetData.appendChild(rowEl);
  }
  return rowEl;
}

function ensureCellXml(doc, rowEl, cellRef, templateCellEl = null) {
  let cellEl = getCellEl(rowEl, cellRef);
  if (cellEl) return cellEl;

  cellEl = doc.createElementNS(XML_SHEET_NS, "c");
  cellEl.setAttribute("r", cellRef);
  if (templateCellEl && templateCellEl.hasAttribute("s")) {
    cellEl.setAttribute("s", templateCellEl.getAttribute("s"));
  }
  if (templateCellEl && templateCellEl.hasAttribute("z")) {
    cellEl.setAttribute("z", templateCellEl.getAttribute("z"));
  }

  const cells = Array.from(rowEl.getElementsByTagNameNS(XML_SHEET_NS, "c"));
  const newCol = cellRef.match(/^[A-Z]+/i)?.[0] || "";
  const newColNum = colLetterToNumber(newCol);
  const insertBefore = cells.find((c) => colLetterToNumber((c.getAttribute("r") || "").match(/^[A-Z]+/i)?.[0] || "") > newColNum);
  if (insertBefore) {
    rowEl.insertBefore(cellEl, insertBefore);
  } else {
    rowEl.appendChild(cellEl);
  }
  return cellEl;
}

function clearCellChildren(cellEl) {
  Array.from(cellEl.children).forEach((child) => child.remove());
}

function setXmlCellNumber(doc, rowEl, cellRef, value, templateCellEl = null) {
  const cellEl = ensureCellXml(doc, rowEl, cellRef, templateCellEl);
  cellEl.setAttribute("t", "n");
  clearCellChildren(cellEl);
  const vEl = doc.createElementNS(XML_SHEET_NS, "v");
  vEl.textContent = String(value);
  cellEl.appendChild(vEl);
}

function setXmlCellText(doc, rowEl, cellRef, value, templateCellEl = null) {
  const cellEl = ensureCellXml(doc, rowEl, cellRef, templateCellEl);
  cellEl.setAttribute("t", "inlineStr");
  clearCellChildren(cellEl);
  const isEl = doc.createElementNS(XML_SHEET_NS, "is");
  const tEl = doc.createElementNS(XML_SHEET_NS, "t");
  tEl.textContent = String(value ?? "");
  isEl.appendChild(tEl);
  cellEl.appendChild(isEl);
}

function getCellStyleFromRef(doc, cellRef) {
  const rowNumber = Number(cellRef.match(/\d+$/)?.[0] || "0");
  const rowEl = getRowEl(doc, rowNumber);
  if (!rowEl) return null;
  const cellEl = getCellEl(rowEl, cellRef);
  if (!cellEl) return null;
  return {
    s: cellEl.getAttribute("s"),
    z: cellEl.getAttribute("z"),
  };
}

function getTemplateRowForStyle(doc, preferredRows = []) {
  for (const rowNum of preferredRows) {
    const rowEl = getRowEl(doc, rowNum);
    if (rowEl) return rowEl;
  }
  const rows = Array.from(getSheetDataEl(doc)?.getElementsByTagNameNS(XML_SHEET_NS, "row") || []);
  return rows[0] || null;
}

async function readZipText(zip, path) {
  const file = zip.file(path);
  return file ? file.async("string") : "";
}

function buildWorkbookSheetMaps(wbDoc, relsDoc) {
  const relMap = {};
  Array.from(relsDoc.getElementsByTagNameNS(XML_PACKAGE_REL_NS, "Relationship")).forEach((rel) => {
    relMap[rel.getAttribute("Id")] = rel.getAttribute("Target");
  });

  const sheetMap = new Map();
  const sheets = Array.from(wbDoc.getElementsByTagNameNS(XML_SHEET_NS, "sheet"));
  sheets.forEach((sheet) => {
    const name = sheet.getAttribute("name") || "";
    const relId = sheet.getAttributeNS(XML_REL_NS, "id") || sheet.getAttribute("r:id");
    const target = relMap[relId] || "";
    if (!name || !target) return;
    sheetMap.set(name, target.startsWith("xl/") ? target : `xl/${target}`);
  });

  return { relMap, sheetMap };
}

function getSheetNameToRowInfo(sheetObj, exportDate) {
  const rows = XLSX.utils.sheet_to_json(sheetObj, { header: 1, defval: null, raw: false });
  let lastDataRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const value = rows[i]?.[0];
    if (value != null && String(value).trim() !== "") lastDataRow = i;
  }

  for (let i = 0; i < rows.length; i++) {
    if (normalizeDateValue(rows[i]?.[0]) === exportDate) {
      return { row: i, isNew: false, lastDataRow };
    }
  }

  return { row: Math.max(lastDataRow + 1, 0), isNew: true, lastDataRow };
}

function updateStoreSheetXml(sheetDoc, sheetObj, store, exportDate) {
  const sheetData = getSheetDataEl(sheetDoc);
  const { row, isNew, lastDataRow } = getSheetNameToRowInfo(sheetObj, exportDate);
  const templateRowNumber = isNew && lastDataRow >= 0 ? lastDataRow + 1 : row + 1;
  const templateRowEl =
    getRowEl(sheetDoc, templateRowNumber) || getTemplateRowForStyle(sheetDoc, [templateRowNumber, Math.max(1, templateRowNumber - 1), 1]);
  const targetRowEl = ensureRowXml(sheetDoc, row + 1, templateRowEl);

  if (isNew && templateRowEl && targetRowEl !== templateRowEl) {
    const existingCells = Array.from(targetRowEl.getElementsByTagNameNS(XML_SHEET_NS, "c"));
    existingCells.forEach((cell) => cell.remove());
    Array.from(templateRowEl.getElementsByTagNameNS(XML_SHEET_NS, "c")).forEach((templateCell) => {
      const ref = templateCell.getAttribute("r") || "";
      const col = ref.match(/^[A-Z]+/i)?.[0];
      if (!col) return;
      const newRef = `${col}${row + 1}`;
      const clonedCell = cloneCellForRow(sheetDoc, templateCell, newRef);
      targetRowEl.appendChild(clonedCell);
    });
  }

  if (isNew) {
    Array.from(targetRowEl.getElementsByTagNameNS(XML_SHEET_NS, "c")).forEach((cell) => {
      clearCellChildren(cell);
      cell.removeAttribute("t");
    });
  }

  const dateTemplate = getCellEl(targetRowEl, `A${row + 1}`) || getCellByColumn(templateRowEl, "A") || getCellEl(templateRowEl, "A1");
  const exportSerial = dateToExcelSerial(exportDate);
  if (exportSerial == null) throw new Error(`Date invalide: ${exportDate}`);
  setXmlCellNumber(sheetDoc, targetRowEl, `A${row + 1}`, exportSerial, dateTemplate);

  PRODUCTS.forEach((product) => {
    const livrer = toNumber(store.values[product.article].livrer);
    const vendu = toNumber(store.values[product.article].vendu);

    const livrerRef = `${numberToColLetter(product.livrerCol + 1)}${row + 1}`;
    const venduRef = `${numberToColLetter(product.venduCol + 1)}${row + 1}`;
    const livrerTemplate = getCellEl(targetRowEl, livrerRef) || getCellByColumn(templateRowEl, numberToColLetter(product.livrerCol + 1));
    const venduTemplate = getCellEl(targetRowEl, venduRef) || getCellByColumn(templateRowEl, numberToColLetter(product.venduCol + 1));

    if (isBlankValue(store.values[product.article].livrer)) {
      setXmlCellText(sheetDoc, targetRowEl, livrerRef, "", livrerTemplate);
    } else {
      setXmlCellNumber(sheetDoc, targetRowEl, livrerRef, livrer, livrerTemplate);
    }
    setXmlCellNumber(sheetDoc, targetRowEl, venduRef, vendu || 0, venduTemplate);
  });

  const refs = Array.from(sheetData.getElementsByTagNameNS(XML_SHEET_NS, "row")).map((rowEl) => Number(rowEl.getAttribute("r")));
  const maxRow = refs.length ? Math.max(...refs) : row + 1;
  const dimensionEl = sheetDoc.getElementsByTagNameNS(XML_SHEET_NS, "dimension")[0];
  if (dimensionEl) {
    const ref = dimensionEl.getAttribute("ref") || "";
    const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (m) {
      dimensionEl.setAttribute("ref", `${m[1]}${m[2]}:${m[3]}${Math.max(Number(m[4]), maxRow)}`);
    }
  }

  return { updated: true, isNew };
}

function buildRecapSheetXml(templateXml) {
  const doc = parseXml(templateXml);
  const sheetData = getSheetDataEl(doc);
  while (sheetData.firstChild) sheetData.removeChild(sheetData.firstChild);

  const titleTemplateRow = getTemplateRowForStyle(doc, [1]);
  const headerTemplateRow = getTemplateRowForStyle(doc, [3, 2, 1]);
  const dataTemplateRow = getTemplateRowForStyle(doc, [7, 6, 5, 4, 3]);

  const titleRow = ensureRowXml(doc, 1, titleTemplateRow);
  while (titleRow.firstChild) titleRow.removeChild(titleRow.firstChild);
  setXmlCellText(doc, titleRow, "A1", "Récapitulatif", getCellByColumn(titleTemplateRow, "A"));
  setXmlCellText(doc, titleRow, "B1", getExportDate(), getCellByColumn(titleTemplateRow, "B"));

  const headerRow = ensureRowXml(doc, 3, headerTemplateRow);
  while (headerRow.firstChild) headerRow.removeChild(headerRow.firstChild);
  ["Magasin", "Article", "Produit", "Livré", "Vendu", "Total"].forEach((label, index) => {
    const ref = `${numberToColLetter(index + 1)}3`;
    setXmlCellText(
      doc,
      headerRow,
      ref,
      label,
      getCellByColumn(headerTemplateRow, numberToColLetter(index + 1)) || getCellByColumn(titleTemplateRow, numberToColLetter(index + 1)),
    );
  });

  let rowNumber = 4;
  for (const store of state.stores) {
    if (store.removed) continue;

    let firstLine = true;
    let storeLivrer = 0;
    let storeVendu = 0;

    for (const product of PRODUCTS) {
      const dataRow = ensureRowXml(doc, rowNumber, dataTemplateRow);
      while (dataRow.firstChild) dataRow.removeChild(dataRow.firstChild);

      const livrer = toNumber(store.values[product.article].livrer);
      const vendu = toNumber(store.values[product.article].vendu);
      storeLivrer += livrer;
      storeVendu += vendu;

      setXmlCellText(doc, dataRow, `A${rowNumber}`, firstLine ? `Mag ${store.magNum}` : "", getCellByColumn(dataTemplateRow, "A"));
      setXmlCellText(doc, dataRow, `B${rowNumber}`, product.article, getCellByColumn(dataTemplateRow, "B"));
      setXmlCellText(doc, dataRow, `C${rowNumber}`, product.name, getCellByColumn(dataTemplateRow, "C"));
      setXmlCellNumber(doc, dataRow, `D${rowNumber}`, livrer || 0, getCellByColumn(dataTemplateRow, "D"));
      setXmlCellNumber(doc, dataRow, `E${rowNumber}`, vendu || 0, getCellByColumn(dataTemplateRow, "E"));
      setXmlCellNumber(doc, dataRow, `F${rowNumber}`, livrer + vendu, getCellByColumn(dataTemplateRow, "F"));

      rowNumber += 1;
      firstLine = false;
    }

    const totalRow = ensureRowXml(doc, rowNumber, dataTemplateRow);
    while (totalRow.firstChild) totalRow.removeChild(totalRow.firstChild);
    setXmlCellText(doc, totalRow, `A${rowNumber}`, `Total Mag ${store.magNum}`, getCellByColumn(dataTemplateRow, "A"));
    setXmlCellNumber(doc, totalRow, `D${rowNumber}`, storeLivrer, getCellByColumn(dataTemplateRow, "D"));
    setXmlCellNumber(doc, totalRow, `E${rowNumber}`, storeVendu, getCellByColumn(dataTemplateRow, "E"));
    setXmlCellNumber(doc, totalRow, `F${rowNumber}`, storeLivrer + storeVendu, getCellByColumn(dataTemplateRow, "F"));

    rowNumber += 2;
  }

  const dimension = `A1:F${Math.max(4, rowNumber - 1)}`;
  const dimensionEl = doc.getElementsByTagNameNS(XML_SHEET_NS, "dimension")[0];
  if (dimensionEl) dimensionEl.setAttribute("ref", dimension);
  return serializeXml(doc);
}

function makeWorksheetContentTypesEntry(sheetPath) {
  return `  <Override PartName="/${sheetPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
}

function addContentTypeOverride(ctDoc, sheetPath) {
  const overrides = Array.from(ctDoc.getElementsByTagName("Override"));
  const already = overrides.some((el) => el.getAttribute("PartName") === `/${sheetPath}`);
  if (already) return;
  const root = ctDoc.documentElement;
  const override = ctDoc.createElement("Override");
  override.setAttribute("PartName", `/${sheetPath}`);
  override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
  root.appendChild(override);
}

function addWorkbookSheetEntry(wbDoc, sheetName, sheetId, relId, position = "end") {
  const sheetsEl = wbDoc.getElementsByTagNameNS(XML_SHEET_NS, "sheets")[0];
  if (!sheetsEl) throw new Error("Balise sheets introuvable dans workbook.xml");

  const existing = Array.from(sheetsEl.getElementsByTagNameNS(XML_SHEET_NS, "sheet")).find((sheet) => sheet.getAttribute("name") === sheetName);
  if (existing) existing.remove();

  const sheetEl = wbDoc.createElementNS(XML_SHEET_NS, "sheet");
  sheetEl.setAttribute("name", sheetName);
  sheetEl.setAttribute("sheetId", String(sheetId));
  sheetEl.setAttributeNS(XML_REL_NS, "r:id", relId);
  sheetEl.setAttribute("r:id", relId);

  if (position === "start" && sheetsEl.firstChild) {
    sheetsEl.insertBefore(sheetEl, sheetsEl.firstChild);
  } else {
    sheetsEl.appendChild(sheetEl);
  }
}

function cellRefToParts(ref) {
  const match = String(ref || "").match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    col: match[1],
    row: Number(match[2]),
    rowIndex: Number(match[2]) - 1,
  };
}

function cloneRowFormatting(sheet, fromRowIndex, toRowIndex) {
  const clonedRows = [];
  const fromRow = fromRowIndex + 1;
  const toRow = toRowIndex + 1;

  Object.keys(sheet).forEach((key) => {
    if (key.startsWith("!")) return;
    const parts = cellRefToParts(key);
    if (!parts || parts.row !== fromRow) return;

    const targetRef = key.replace(/\d+$/, String(toRow));
    const src = sheet[key];
    const copy = { ...src };

    delete copy.f;
    delete copy.v;
    delete copy.w;

    sheet[targetRef] = copy;
    clonedRows.push(targetRef);
  });

  if (sheet["!rows"] && sheet["!rows"][fromRowIndex]) {
    sheet["!rows"][toRowIndex] = { ...sheet["!rows"][fromRowIndex] };
  }

  return clonedRows;
}

function setStyledNumberCell(sheet, rowIndex, colIndex, value, templateCell) {
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const existing = sheet[addr] || {};
  const cell = sheet[addr] || {};

  if (templateCell && !existing.s && templateCell.s != null) {
    cell.s = templateCell.s;
  }
  if (templateCell && !existing.z && templateCell.z != null) {
    cell.z = templateCell.z;
  }

  cell.t = "n";
  cell.v = value;
  cell.w = String(value);
  sheet[addr] = cell;
}

function setStyledTextCell(sheet, rowIndex, colIndex, value, templateCell) {
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const existing = sheet[addr] || {};
  const cell = sheet[addr] || {};

  if (templateCell && !existing.s && templateCell.s != null) {
    cell.s = templateCell.s;
  }
  if (templateCell && !existing.z && templateCell.z != null) {
    cell.z = templateCell.z;
  }

  cell.t = "s";
  cell.v = value;
  cell.w = value;
  sheet[addr] = cell;
}

function copyExcelJsRowStyle(sourceRow, targetRow, maxCol = 24) {
  targetRow.height = sourceRow.height;
  targetRow.hidden = sourceRow.hidden;
  targetRow.outlineLevel = sourceRow.outlineLevel;

  for (let c = 1; c <= maxCol; c++) {
    const src = sourceRow.getCell(c);
    const dst = targetRow.getCell(c);
    dst.style = JSON.parse(JSON.stringify(src.style || {}));
    if (src.numFmt) dst.numFmt = src.numFmt;
  }
}

function applyUniformSheetFormats(ws, startRow = 5, endRow = 47) {
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    const dateCell = row.getCell(1);
    const normalizedDate = normalizeDateValue(dateCell.value || dateCell.text);
    if (normalizedDate) {
      const jsDate = normalizedDateToJsDate(normalizedDate);
      dateCell.value = jsDate || normalizedDate;
      dateCell.numFmt = "yyyy-mm-dd";
      dateCell.style = { ...(dateCell.style || {}), numFmt: "yyyy-mm-dd" };
    }

    for (let c = 2; c <= 24; c++) {
      const cell = row.getCell(c);
      cell.alignment = { ...(cell.alignment || {}), horizontal: "center" };
      if (cell.value != null && cell.value !== "") {
        if (typeof cell.value === "number") {
          cell.numFmt = "0";
        } else {
          const n = Number(String(cell.value).replace(/\s/g, "").replace(",", "."));
          if (Number.isFinite(n)) {
            cell.value = n;
            cell.numFmt = "0";
          }
        }
      } else {
        cell.numFmt = "0";
      }
    }
  }
}

function setCellFont(cell, overrides) {
  const current = cell.font || {};
  cell.font = { ...current, ...overrides };
}

function applyUniformWorkbookFont(workbook) {
  workbook.eachSheet((ws) => {
    ws.getRow(1).height = 24;
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        setCellFont(cell, { name: "Calibri", size: 11 });
        if (rowNumber === 1 && (colNumber === 5 || colNumber === 6)) {
          setCellFont(cell, { name: "Arial", size: 14, bold: true });
        }
      });
    });
  });
}

function excelJsRowHasDate(row, exportDate) {
  const cell = row.getCell(1);
  return normalizeDateValue(cell.value) === exportDate || normalizeDateValue(cell.text) === exportDate;
}

function findDateSectionRow(ws, exportDate) {
  let headerRow = null;

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const label = normalizeSpaces(row.getCell(1).text || row.getCell(1).value).toLowerCase();
    if (!headerRow && label === "date") {
      headerRow = rowNumber;
    }
  });

  if (!headerRow) {
    return { targetRow: 1, templateRow: 1, isNew: true };
  }

  let lastDateRow = headerRow;
  let foundDateRow = null;
  let sawDateBlock = false;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const aVal = normalizeDateValue(row.getCell(1).value || row.getCell(1).text);
    const hasAnyData = row.actualCellCount > 0 || row.values.some((v, idx) => idx > 0 && v != null);

    if (aVal) {
      sawDateBlock = true;
      lastDateRow = r;
      if (aVal === exportDate) {
        foundDateRow = r;
        break;
      }
      continue;
    }

    if (sawDateBlock && !hasAnyData) {
      break;
    }
  }

  if (foundDateRow) {
    return { targetRow: foundDateRow, templateRow: foundDateRow, isNew: false };
  }

  return { targetRow: lastDateRow + 1, templateRow: lastDateRow, isNew: true };
}

function inspectDateSection(ws) {
  let headerRow = null;

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const label = normalizeSpaces(row.getCell(1).text || row.getCell(1).value).toLowerCase();
    if (!headerRow && label === "date") {
      headerRow = rowNumber;
    }
  });

  if (!headerRow) {
    return null;
  }

  let lastDateRow = headerRow;
  let sawDateBlock = false;
  const rows = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const aVal = normalizeDateValue(row.getCell(1).value || row.getCell(1).text);
    const hasAnyData = row.actualCellCount > 0 || row.values.some((v, idx) => idx > 0 && v != null);

    if (aVal || hasAnyData) {
      sawDateBlock = true;
      lastDateRow = r;
      rows.push(r);
      continue;
    }

    if (sawDateBlock && !hasAnyData) {
      break;
    }
  }

  return { headerRow, lastDateRow, rows };
}

function validateWorkbookBeforeExport(workbook, exportDate) {
  const report = {
    exportDate,
    sheetsChecked: 0,
    sheetsRenamed: 0,
    dateFilled: 0,
    zeroFilled: 0,
    missingDateRows: [],
    sheets: [],
  };

  for (const storeInfo of state.canac.stores) {
    const ws = workbook.getWorksheet(storeInfo.sheetName);
    if (!ws) continue;

    report.sheetsChecked += 1;
    const section = inspectDateSection(ws);
    const sheetReport = {
      sheetName: storeInfo.sheetName,
      magNum: storeInfo.magNum,
      dateFilled: 0,
      zeroFilled: 0,
      missingDateRows: [],
    };

    if (section) {
      let previousDate = "";

      for (let r = section.headerRow + 1; r <= section.lastDateRow; r++) {
        const row = ws.getRow(r);
        let aVal = normalizeDateValue(row.getCell(1).value || row.getCell(1).text);
        const hasAnyData = row.actualCellCount > 0 || row.values.some((v, idx) => idx > 0 && v != null);

        if (!aVal && hasAnyData && previousDate) {
          const filledDate = addDaysToNormalizedDate(previousDate, 3);
          if (filledDate) {
            const dateCell = row.getCell(1);
            const jsDate = normalizedDateToJsDate(filledDate);
            dateCell.value = jsDate || filledDate;
            dateCell.numFmt = "yyyy-mm-dd";
            dateCell.style = { ...(dateCell.style || {}), numFmt: "yyyy-mm-dd" };
            aVal = filledDate;
            sheetReport.dateFilled += 1;
            report.dateFilled += 1;
          }
        }

        if (aVal) {
          previousDate = aVal;
        }

        if (!aVal && hasAnyData) {
          sheetReport.missingDateRows.push(r);
          report.missingDateRows.push({
            sheetName: storeInfo.sheetName,
            magNum: storeInfo.magNum,
            rowNumber: r,
          });
        }

        if (!hasAnyData && !aVal) {
          continue;
        }

        PRODUCTS.forEach((product) => {
          const venduCell = row.getCell(product.venduCol + 1);
          const isBlank = venduCell.value == null || venduCell.value === "";
          if (isBlank) {
            venduCell.value = 0;
            sheetReport.zeroFilled += 1;
            report.zeroFilled += 1;
          }
        });
      }
    }

    applyUniformSheetFormats(ws, 5, 47);

    report.sheets.push(sheetReport);
  }

  return report;
}

function normalizeWorkbookSheetNames(workbook) {
  let renamed = 0;
  const seen = new Set();

  for (const storeInfo of state.canac.stores) {
    const desiredName = `Magasin ${storeInfo.magNum}`;
    const worksheet = workbook.getWorksheet(storeInfo.sheetName) || workbook.getWorksheet(desiredName);
    if (!worksheet) continue;

    seen.add(worksheet.name);
    if (worksheet.name !== desiredName) {
      worksheet.name = desiredName;
      renamed += 1;
    }
    storeInfo.sheetName = desiredName;
  }

  return { renamed, seen: Array.from(seen) };
}

async function buildWorkbookForExport() {
  if (!state.canac.buffer) {
    throw new Error("Charge d'abord un fichier CANAC.");
  }

  const exportDate = getExportDate();
  if (typeof ExcelJS === "undefined") {
    throw new Error("ExcelJS est manquant.");
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(state.canac.buffer);

  const renamedInfo = normalizeWorkbookSheetNames(workbook);
  const validation = validateWorkbookBeforeExport(workbook, exportDate);
  validation.sheetsRenamed = renamedInfo.renamed;
  state.lastValidation = validation;
  return { workbook, exportDate, validation };
}

function getExportDate() {
  const normalized = normalizeDateValue(state.targetDate);
  if (normalized) return normalized;
  state.targetDate = DUMMY_DATE;
  $("target-date").value = DUMMY_DATE;
  return DUMMY_DATE;
}

function updateWorkbookWithState() {
  if (!state.canac.buffer) {
    throw new Error("Charge d'abord un fichier CANAC.");
  }

  const exportDate = getExportDate();
  const workbook = cloneWorkbook(state.canac.buffer);

  for (const store of state.stores) {
    if (store.removed) continue;

    const sheetName = store.sheetName || `Magasin ${store.magNum}`;
    const sheet = ensureSheet(workbook, sheetName);
    const { row, isNew } = findRowWithDate(sheet, exportDate);
    const templateRow = isNew && row > 0 ? row - 1 : row;

    if (isNew && templateRow >= 0) {
      cloneRowFormatting(sheet, templateRow, row);
    }

    const dateTemplate = sheet[XLSX.utils.encode_cell({ r: templateRow, c: 0 })];
    setStyledTextCell(sheet, row, 0, exportDate, dateTemplate);

    PRODUCTS.forEach((product) => {
      const livrer = toNumber(store.values[product.article].livrer);
      const vendu = toNumber(store.values[product.article].vendu);

      const livrerTemplate = sheet[XLSX.utils.encode_cell({ r: templateRow, c: product.livrerCol })];
      const venduTemplate = sheet[XLSX.utils.encode_cell({ r: templateRow, c: product.venduCol })];

      if (isBlankValue(store.values[product.article].livrer)) {
        setStyledTextCell(sheet, row, product.livrerCol, "", livrerTemplate);
      } else {
        setStyledNumberCell(sheet, row, product.livrerCol, livrer, livrerTemplate);
      }
      setStyledNumberCell(sheet, row, product.venduCol, vendu || 0, venduTemplate);
    });

    const ref = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    ref.e.r = Math.max(ref.e.r, row + 1);
    ref.e.c = Math.max(ref.e.c, 24);
    sheet["!ref"] = XLSX.utils.encode_range(ref);

    if (isNew) {
      showStatus(`Nouvelle ligne ajoutée dans ${sheetName}.`);
    }
  }

  return workbook;
}

function addManualStore() {
  const magNum = normalizeDivision($("store-input").value);
  if (!magNum) {
    showStatus("Entre un numéro de magasin valide.");
    return;
  }

  if (state.stores.some((store) => store.magNum === magNum)) {
    setSelectedStore(state.stores.find((store) => store.magNum === magNum).id);
    return;
  }

  const store = createStore({ magNum, sheetName: `Magasin ${magNum}`, source: "manual" });
  if (state.royaumes.totals[magNum]) {
    const totals = state.royaumes.totals[magNum];
    PRODUCTS.forEach((product) => {
      store.values[product.article].vendu = totals[product.article] ? String(Math.round(totals[product.article])) : "";
    });
  }

  state.stores.push(store);
  state.selectedStoreId = store.id;
  $("store-input").value = "";
  render();
}

function removeStore(id) {
  const idx = state.stores.findIndex((store) => store.id === id);
  if (idx === -1) return;
  state.stores.splice(idx, 1);
  if (state.selectedStoreId === id) {
    state.selectedStoreId = state.stores[0]?.id || null;
  }
  render();
}

function updateCell(storeId, article, field, value) {
  const store = state.stores.find((item) => item.id === storeId);
  if (!store) return;
  store.values[article][field] = value;
  renderSummaryOnly();
}

function renderStores() {
  const list = $("stores-list");
  if (!list) return;

  if (!state.stores.length) {
    list.innerHTML = '<div class="status">Aucun magasin pour l’instant.</div>';
    return;
  }

  list.innerHTML = state.stores
    .map((store) => {
      const isActive = store.id === state.selectedStoreId ? "active" : "";
      const totalVendu = PRODUCTS.reduce((sum, product) => sum + toNumber(store.values[product.article].vendu), 0);
      const totalLivrer = PRODUCTS.reduce((sum, product) => sum + toNumber(store.values[product.article].livrer), 0);

      return `
        <div class="store-item ${isActive}" data-store="${store.id}">
          <div>
            <div><strong>Mag ${store.magNum}</strong></div>
            <small>${store.sheetName || "Sans feuille"} · Vendu ${Math.round(totalVendu)} · Livré ${Math.round(totalLivrer)}</small>
          </div>
          <div class="store-actions">
            <button class="icon-btn" data-action="select" data-store="${store.id}" type="button">›</button>
            <button class="icon-btn" data-action="remove" data-store="${store.id}" type="button">×</button>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderStoreSelect() {
  const select = $("store-select");
  if (!select) return;

  if (!state.stores.length) {
    select.innerHTML = '<option value="">Aucun magasin</option>';
    select.disabled = true;
    return;
  }

  select.disabled = false;
  select.innerHTML = state.stores
    .map((store) => {
      const label = `${store.sheetName ? ` ${store.sheetName}` : ""}`;
      return `<option value="${store.id}"${store.id === state.selectedStoreId ? " selected" : ""}>${label}</option>`;
    })
    .join("");

  if (!state.selectedStoreId || !state.stores.some((store) => store.id === state.selectedStoreId)) {
    state.selectedStoreId = state.stores[0].id;
    select.value = state.selectedStoreId;
  }
}

function renderEditor() {
  const editor = $("editor");
  const store = currentStore();

  if (!store) {
    $("editor-title").textContent = "Aucun magasin sélectionné";
    $("editor-subtitle").textContent = "Charge un fichier CANAC pour ajouter des livraisons.";
    editor.innerHTML = '<div class="status">Sélectionne un magasin pour éditer ses quantités.</div>';
    return;
  }

  $("editor-title").textContent = `Mag ${store.magNum}`;
  $("editor-subtitle").textContent = store.sheetName || "Magasin manuel";

  const rows = PRODUCTS.map((product) => {
    const values = store.values[product.article];
    return `
      <tr>
        <td>
          <div class="product-name">${product.name}</div>
          <div class="article">${product.article}</div>
        </td>
        <td><input class="qty" type="number" min="0" step="1" data-field="livrer" data-article="${product.article}" value="${values.livrer}" placeholder="0" /></td>
        <td><input class="qty" type="number" min="0" step="1" data-field="vendu" data-article="${product.article}" value="${values.vendu}" placeholder="0" /></td>
      </tr>
    `;
  }).join("");

  editor.innerHTML = `
    <table class="product-table">
      <thead>
        <tr>
          <th>Produit</th>
          <th>Livré</th>
          <th>Vendu</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderSummaryOnly() {
  const summary = $("summary");
  if (!summary) return;

  const storeCount = state.stores.length;
  const storesWithData = state.stores.filter((store) =>
    PRODUCTS.some((product) => toNumber(store.values[product.article].livrer) > 0 || toNumber(store.values[product.article].vendu) > 0),
  ).length;
  const totalLivrer = state.stores.reduce(
    (sum, store) => sum + PRODUCTS.reduce((s, product) => s + toNumber(store.values[product.article].livrer), 0),
    0,
  );
  const totalVendu = state.stores.reduce(
    (sum, store) => sum + PRODUCTS.reduce((s, product) => s + toNumber(store.values[product.article].vendu), 0),
    0,
  );
  const missingLivrer = state.stores.reduce(
    (sum, store) => sum + PRODUCTS.filter((product) => !toNumber(store.values[product.article].livrer)).length,
    0,
  );
  const validation = state.lastValidation;
  const validationSummary = validation
    ? `
      <div class="metric">
        <span>Onglets renommés</span>
        <strong>${validation.sheetsRenamed || 0}</strong>
      </div>
      <div class="metric">
        <span>Dates comblées</span>
        <strong>${validation.dateFilled}</strong>
      </div>
      <div class="metric ${validation.missingDateRows.length ? "warn" : ""}">
        <span>Dates à revoir</span>
        <strong>${validation.missingDateRows.length}</strong>
      </div>
      <div class="metric">
        <span>Zéros ajoutés</span>
        <strong>${validation.zeroFilled}</strong>
      </div>
      <div class="metric">
        <span>Onglets vérifiés</span>
        <strong>${validation.sheetsChecked}</strong>
      </div>
    `
    : "";

  summary.innerHTML = `
    <div class="metric">
      <span>Magasins</span>
      <strong>${storeCount}</strong>
    </div>
    <div class="metric">
      <span>Magasins remplis</span>
      <strong>${storesWithData}</strong>
    </div>
    <div class="metric">
      <span>Total livré</span>
      <strong>${Math.round(totalLivrer).toLocaleString()}</strong>
    </div>
    <div class="metric">
      <span>Total vendu</span>
      <strong>${Math.round(totalVendu).toLocaleString()}</strong>
    </div>
    <div class="metric warn">
      <span>Champs livrés vides</span>
      <strong>${missingLivrer}</strong>
    </div>
    ${validationSummary}
  `;
}

function render() {
  $("target-date").value = state.targetDate;
  renderStores();
  renderStoreSelect();
  renderEditor();
  renderSummaryOnly();
}

async function exportFile() {
  try {
    showStatus("Vérification du fichier...");
    const { workbook, exportDate, validation } = await buildWorkbookForExport();
    const details = [];
    if (validation.zeroFilled) details.push(`${validation.zeroFilled} zéro(s) ajouté(s)`);
    if (validation.missingDateRows.length) details.push(`${validation.missingDateRows.length} date(s) à revoir`);
    showStatus(details.length ? `Vérif faite: ${details.join(", ")}.` : "Vérif faite: rien à corriger.");
    renderSummaryOnly();

    showStatus("Mise à jour des onglets...");
    for (const storeInfo of state.canac.stores) {
      const store = state.stores.find((s) => s.magNum === storeInfo.magNum && !s.removed);
      if (!store) continue;

      const ws = workbook.getWorksheet(storeInfo.sheetName);
      if (!ws) continue;

      const { targetRow, templateRow, isNew } = findDateSectionRow(ws, exportDate);
      if (isNew) {
        const sourceRow = ws.getRow(Math.max(templateRow, 1));
        const newRow = ws.getRow(targetRow);
        copyExcelJsRowStyle(sourceRow, newRow, 24);
      }

      const row = ws.getRow(targetRow);
      const dateCell = row.getCell(1);
      const exportJsDate = normalizedDateToJsDate(exportDate);
      dateCell.value = exportJsDate || exportDate;
      dateCell.numFmt = "yyyy-mm-dd";
      dateCell.style = { ...(dateCell.style || {}), numFmt: "yyyy-mm-dd" };

      PRODUCTS.forEach((product) => {
        const livrer = toNumber(store.values[product.article].livrer);
        const vendu = toNumber(store.values[product.article].vendu);
        const livrerCell = row.getCell(product.livrerCol + 1);
        const venduCell = row.getCell(product.venduCol + 1);
        if (isBlankValue(store.values[product.article].livrer)) {
          livrerCell.value = null;
        } else {
          livrerCell.value = livrer;
        }
        venduCell.value = vendu || 0;
        livrerCell.numFmt = "0";
        venduCell.numFmt = "0";
        livrerCell.alignment = { ...(livrerCell.alignment || {}), horizontal: "center" };
        venduCell.alignment = { ...(venduCell.alignment || {}), horizontal: "center" };
      });
    }

    applyUniformWorkbookFont(workbook);
    showStatus("Compression du fichier...");
    const out = await workbook.xlsx.writeBuffer();
    const blob = new Blob([out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `CANAC_MAJ_${exportDate}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showStatus(`Fichier généré. Date ${getExportDate()}.`);
  } catch (error) {
    console.error(error);
    showStatus(`Erreur export: ${error.message}`);
  }
}

async function verifyFile() {
  try {
    showStatus("Vérification du fichier...");
    const { validation } = await buildWorkbookForExport();
    const details = [];
    if (validation.zeroFilled) details.push(`${validation.zeroFilled} zéro(s) ajouté(s)`);
    if (validation.missingDateRows.length) details.push(`${validation.missingDateRows.length} date(s) à revoir`);
    showStatus(details.length ? `Vérif faite: ${details.join(", ")}.` : "Vérif faite: rien à corriger.");
    renderSummaryOnly();
  } catch (error) {
    console.error(error);
    showStatus(`Erreur vérif: ${error.message}`);
  }
}

async function loadCanac(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  state.canac.name = file.name;
  state.canac.buffer = buffer;
  state.canac.workbook = workbook;
  state.canac.stores = parseCanacWorkbook(workbook).stores;

  syncStoreListFromCanac();
  if (!state.selectedStoreId && state.stores[0]) {
    state.selectedStoreId = state.stores[0].id;
  }

  if (state.royaumes.totals && Object.keys(state.royaumes.totals).length) {
    syncFromRoyaumes();
  }

  showStatus(`CANAC chargé: ${file.name}`);
  render();
}

async function loadRoyaumes(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const parsed = parseRoyaumesWorkbook(workbook);

  if (parsed.error) {
    showStatus(parsed.error);
    return;
  }

  state.royaumes.name = file.name;
  state.royaumes.workbook = workbook;
  state.royaumes.totals = parsed.totals;

  if (state.stores.length) {
    syncFromRoyaumes();
  }

  showStatus(`Royaumes chargé: ${file.name}`);
  render();
}

function wireEvents() {
  $("canac-file").addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) loadCanac(file);
  });

  $("royaumes-file").addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) loadRoyaumes(file);
  });

  $("target-date").value = state.targetDate;
  $("target-date").addEventListener("change", (event) => {
    state.targetDate = event.target.value;
  });

  $("add-store-btn").addEventListener("click", addManualStore);
  $("add-store-submit").addEventListener("click", addManualStore);
  $("store-input").addEventListener("keydown", (event) => {
    if (event.key === "Enter") addManualStore();
  });

  $("stores-list").addEventListener("click", (event) => {
    const button = event.target.closest("button");
    const item = event.target.closest(".store-item");
    if (!item) return;

    const storeId = item.dataset.store;
    const action = button?.dataset.action;

    if (action === "remove") {
      removeStore(storeId);
      return;
    }

    setSelectedStore(storeId);
  });

  $("editor").addEventListener("input", (event) => {
    const input = event.target.closest("input[data-article][data-field]");
    if (!input) return;
    const store = currentStore();
    if (!store) return;
    updateCell(store.id, input.dataset.article, input.dataset.field, input.value);
  });

  $("store-select").addEventListener("change", (event) => {
    const id = event.target.value;
    if (id) setSelectedStore(id);
  });
  const verifyBtn = $("verify-btn");
  if (verifyBtn) {
    verifyBtn.addEventListener("click", verifyFile);
  }
  $("export-btn").addEventListener("click", exportFile);
}

function boot() {
  $("target-date").value = state.targetDate;
  wireEvents();
  render();
  showStatus("Prêt. Charge le CANAC, puis Royaumes.");
}

boot();
