"use strict";

const QUALITY_HEADERS = ["ПІБ", "ПІБ КС"];
const QUALITY_SCORE_HEADERS = ["Середній бал"];
const UA_MONTHS = [
  "Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень",
  "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"
];

const ZONES = [
  { boundary: 0.10, name: "Зелена", css: "zone-green", color: "FF70AD47", label: "Верхні 10%" },
  { boundary: 0.30, name: "Салатова", css: "zone-light-green", color: "FFC6E0B4", label: "Наступні 20%" },
  { boundary: 0.70, name: "Жовта", css: "zone-yellow", color: "FFFFD966", label: "Середні 40%" },
  { boundary: 0.90, name: "Рожева", css: "zone-pink", color: "FFF4CCCC", label: "Наступні 20%" },
  { boundary: 1.00, name: "Червона", css: "zone-red", color: "FFE06666", label: "Нижні 10%" }
];

const dom = {
  reviewsFile: document.getElementById("reviewsFile"),
  reviewsCard: document.getElementById("reviewsCard"),
  reviewsName: document.getElementById("reviewsName"),
  reviewsState: document.getElementById("reviewsState"),
  removeReviews: document.getElementById("removeReviews"),
  reviewOptions: document.getElementById("reviewOptions"),
  reviewPeriod: document.getElementById("reviewPeriod"),
  reviewSummary: document.getElementById("reviewSummary"),
  reviewDetails: document.getElementById("reviewDetails"),
  reviewAudit: document.getElementById("reviewAudit"),
  reviewMapping: document.getElementById("reviewMapping"),
  reviewMappingList: document.getElementById("reviewMappingList"),
  ratingHead: document.getElementById("ratingHead"),
  qualityFile: document.getElementById("qualityFile"),
  ratingFile: document.getElementById("ratingFile"),
  qualityCard: document.getElementById("qualityCard"),
  ratingCard: document.getElementById("ratingCard"),
  qualityName: document.getElementById("qualityName"),
  ratingName: document.getElementById("ratingName"),
  qualityState: document.getElementById("qualityState"),
  ratingState: document.getElementById("ratingState"),
  sheetRow: document.getElementById("sheetRow"),
  sheetSelect: document.getElementById("sheetSelect"),
  calculateButton: document.getElementById("calculateButton"),
  readiness: document.getElementById("readiness"),
  progress: document.getElementById("progress"),
  results: document.getElementById("results"),
  totalCount: document.getElementById("totalCount"),
  foundCount: document.getElementById("foundCount"),
  missingCount: document.getElementById("missingCount"),
  sheetSummary: document.getElementById("sheetSummary"),
  missingBox: document.getElementById("missingBox"),
  missingText: document.getElementById("missingText"),
  ratingBody: document.getElementById("ratingBody"),
  downloadExcel: document.getElementById("downloadExcel"),
  downloadScores: document.getElementById("downloadScores"),
  toast: document.getElementById("toast")
};

const state = {
  busy: false,
  ratingLoading: false,
  reviewsLoading: false,
  ratingVersion: 0,
  reviewsVersion: 0,
  reviewsFile: null,
  reviewsData: null,
  reviewOverrides: Object.create(null),
  report: null,
  columns: [],
  sourceSheet: "",
  qualityFile: null,
  ratingFile: null,
  ratingWorkbook: null,
  originalResults: [],
  rankedResults: [],
  missing: []
};

function normalizeText(value) {
  return RatingCore.norm(value);
}

function cellValue(cell) {
  return RatingCore.value(cell);
}

function toNumber(value) {
  return RatingCore.number(value);
}

function formatNumber(value) {
  if (value === null || value === undefined) return "—";
  return new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 }).format(value);
}

function personKey(fio) {
  const parts = String(fio ?? "").trim().split(/\s+/).filter(Boolean);
  if (parts.length < 2) return null;
  return [normalizeText(parts[0]), normalizeText(parts[1])];
}

function keyId(key) {
  return key ? `${key[0]}\u0000${key[1]}` : "";
}

function findColumn(worksheet, headers, maxHeaderRows = 5) {
  const wanted = new Set(headers.map(normalizeText));
  const rowLimit = Math.min(maxHeaderRows, worksheet.rowCount || 0);
  const columnLimit = worksheet.columnCount || 0;

  for (let row = 1; row <= rowLimit; row += 1) {
    for (let column = 1; column <= columnLimit; column += 1) {
      if (wanted.has(normalizeText(cellValue(worksheet.getCell(row, column))))) {
        return column;
      }
    }
  }
  throw new Error(`Не знайдено колонку: ${headers.join(" / ")}`);
}

function findHeaderRow(worksheet, column, headers, maxHeaderRows = 5) {
  const wanted = new Set(headers.map(normalizeText));
  const rowLimit = Math.min(maxHeaderRows, worksheet.rowCount || 0);
  for (let row = 1; row <= rowLimit; row += 1) {
    if (wanted.has(normalizeText(cellValue(worksheet.getCell(row, column))))) return row;
  }
  return 1;
}

function buildSurnameIndex(scores) {
  const index = new Map();
  for (const entry of scores.values()) {
    const surname = entry.key[0];
    if (!index.has(surname)) index.set(surname, []);
    index.get(surname).push(entry);
  }
  return index;
}

function findWorksheet(workbook, wantedName) {
  const wanted = normalizeText(wantedName);
  return workbook.worksheets.find(sheet => normalizeText(sheet.name) === wanted) || null;
}

function inferPersonFromPrivateSheet(workbook, sheetName, candidates) {
  const worksheet = findWorksheet(workbook, sheetName);
  if (!worksheet) return null;
  const counts = new Map(candidates.map(candidate => [keyId(candidate.key), 0]));

  worksheet.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      const text = normalizeText(cellValue(cell));
      if (!text) return;
      for (const candidate of candidates) {
        const needle = `мене звати ${candidate.key[1]}`;
        const occurrences = text.split(needle).length - 1;
        if (occurrences) counts.set(keyId(candidate.key), counts.get(keyId(candidate.key)) + occurrences);
      }
    });
  });

  const ranked = candidates
    .map(candidate => ({ candidate, count: counts.get(keyId(candidate.key)) }))
    .sort((a, b) => b.count - a.count);

  if (!ranked.length || ranked[0].count === 0) return null;
  if (ranked.length > 1 && ranked[0].count === ranked[1].count) return null;
  return ranked[0].candidate;
}

async function loadWorkbook(file) {
  if (!window.ExcelJS) throw new Error("Не вдалося відкрити модуль Excel. Перевірте наявність exceljs-4.4.0.min.js та шлях до нього в index.html.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return workbook;
}

async function readQualityScores(file) {
  const workbook = await loadWorkbook(file);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("statistics.xlsx не містить листів.");

  const fioColumn = findColumn(worksheet, QUALITY_HEADERS);
  const scoreColumn = findColumn(worksheet, QUALITY_SCORE_HEADERS);
  const startRow = Math.max(
    findHeaderRow(worksheet, fioColumn, QUALITY_HEADERS),
    findHeaderRow(worksheet, scoreColumn, QUALITY_SCORE_HEADERS)
  ) + 1;

  const scores = new Map();
  const duplicates = [];
  for (let row = startRow; row <= worksheet.rowCount; row += 1) {
    const fio = String(cellValue(worksheet.getCell(row, fioColumn)) ?? "").trim();
    const key = personKey(fio);
    if (!key) continue;
    const id = keyId(key);
    if (scores.has(id)) {
      duplicates.push(fio);
      scores.get(id).score = null;
      continue;
    }
    scores.set(id, {
      key,
      fio,
      score: toNumber(cellValue(worksheet.getCell(row, scoreColumn)))
    });
  }
  return { scores, duplicates };
}

function readRating(workbook, sheetName, scores) {
  const worksheet = findWorksheet(workbook, sheetName);
  if (!worksheet) throw new Error(`Не знайдено лист «${sheetName}».`);

  const layout = RatingCore.ratingLayout(worksheet);
  const surnameIndex = buildSurnameIndex(scores);
  const results = RatingCore.readRows(worksheet, layout);

  for (const item of results) {
    const fio = item.fio;
    const parts = fio.split(/\s+/).filter(Boolean);
    if (!parts.length) continue;

    const candidates = surnameIndex.get(normalizeText(parts[0])) || [];
    let matched = null;

    if (parts.length >= 2) {
      matched = scores.get(keyId([normalizeText(parts[0]), normalizeText(parts[1])])) || null;
    } else if (candidates.length === 1) {
      matched = candidates[0];
    } else if (candidates.length > 1) {
      matched = inferPersonFromPrivateSheet(workbook, fio, candidates);
    }

    item.qualityScore = matched?.score ?? null;
    item.qualityFio = matched?.fio || null;
  }
  state.report = null;
  // No third file: no period validation, no review parser and no review replacement.
  if (state.reviewsFile) {
    if (!state.reviewsData) throw new Error("Перегляди ще не прочитано. Дочекайтеся завантаження або натисніть «Прибрати перегляди».");
    const inferredPeriod = RatingCore.inferPeriod(worksheet);
    if (inferredPeriod && inferredPeriod !== dom.reviewPeriod.value) {
      throw new Error(`Місяць переглядів не збігається з листом рейтингу (${inferredPeriod}). Оберіть відповідний лист або місяць.`);
    }
    state.report = RatingCore.applyReviews(results, state.reviewsData, dom.reviewPeriod.value, state.reviewOverrides);
  }
  const base = RatingCore.normalization(worksheet, layout);
  state.maximum = RatingCore.recalculate(results, base);
  state.baseLabel = base.label;
  state.columns = [...layout.columns,
    { key: "place", title: "Місце", width: 10, format: "0" },
    { key: "status", title: "Статус КЯ / розрахунку", width: 38 },
  ];
  return results;
}

function rankResults(results) {
  const ranked = results.map(item => ({ ...item }))
    .sort((a, b) => (b.finalScore ?? -Infinity) - (a.finalScore ?? -Infinity) || a.fio.localeCompare(b.fio, "uk"));
  const eligible = ranked.filter(item => item.finalScore !== null).length;
  let previous = null;
  let place = 0;

  ranked.forEach((item, index) => {
    if (item.finalScore === null) {
      item.place = null;
      item.zone = { css: "zone-unrated", color: "FFD9D9D9" };
      return;
    }
    const score = Math.round(item.finalScore * 1e10) / 1e10;
    if (score !== previous) {
      place = index + 1;
      previous = score;
    }
    const percentile = (place - 0.5) / Math.max(eligible, 1);
    const zone = ZONES.find(candidate => percentile <= candidate.boundary) || ZONES.at(-1);
    item.place = place;
    item.zone = zone;
  });
  return ranked;
}

function renderResults(sheetName) {
  const found = state.originalResults.filter(item => item.qualityScore !== null).length;
  state.missing = state.originalResults.filter(item => item.qualityScore === null);

  dom.totalCount.textContent = state.originalResults.length;
  dom.foundCount.textContent = found;
  dom.missingCount.textContent = state.missing.length;
  dom.sheetSummary.textContent = sheetName;
  dom.ratingBody.replaceChildren();
  dom.ratingHead.replaceChildren();
  for (const column of state.columns) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = column.title;
    dom.ratingHead.appendChild(cell);
  }

  for (const item of state.rankedResults) {
    const row = document.createElement("tr");
    row.className = item.zone.css;
    state.columns.forEach(column => {
      const cell = document.createElement("td");
      const value = item[column.key];
      cell.textContent = displayValue(value, column);
      if (column.key === "status" && item.issues.length) cell.className = "missing-label";
      if (column.key === "reviews" && Math.abs(item.previousReviews - item.reviews) > 1e-9) {
        cell.classList.add("changed-review");
        cell.title = `Було: ${item.previousReviews}; стало: ${item.reviews}`;
      }
      row.appendChild(cell);
    });
    dom.ratingBody.appendChild(row);
  }
  renderReviews();
  renderReviewMapping();

  if (state.missing.length) {
    dom.missingText.textContent = `${state.missing.map(item => item.fio).join(", ")}. Для них КЯ тимчасово дорівнює 0.`;
    dom.missingBox.hidden = false;
  } else {
    dom.missingBox.hidden = true;
  }
  dom.results.hidden = false;
  (state.report?.unmatched ? dom.reviewMapping : dom.results).scrollIntoView({ behavior: "smooth", block: "start" });
}

function calculationBlockReason() {
  if (state.busy) return "Триває обробка…";
  if (state.ratingLoading) return "Читаю великий файл рейтингу. Дочекайтеся позначки «Готово» — третій файл для запуску не потрібен.";
  if (!state.qualityFile || !state.ratingFile) return "Для запуску потрібні лише statistics.xlsx та rating.xlsx. Перегляди — необов’язкові.";
  if (!state.ratingWorkbook) return "Не вдалося прочитати рейтинг. Перевірте повідомлення про помилку нижче та оберіть файл повторно.";
  if (!dom.sheetSelect.value) return "Оберіть лист фінального рейтингу.";
  if (!state.reviewsFile) return "";
  if (state.reviewsLoading) return "Читаю доданий файл переглядів…";
  if (!state.reviewsData) return "Доданий файл переглядів не прочитано. Оберіть правильний файл або натисніть «Прибрати перегляди» для запуску з двома файлами.";
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(dom.reviewPeriod.value)) return "Для доданого файлу переглядів оберіть місяць і рік.";
  return "";
}

function updateReadyState() {
  const reason = calculationBlockReason();
  dom.calculateButton.disabled = Boolean(reason);
  dom.reviewPeriod.disabled = state.busy || !state.reviewsFile;
  dom.reviewPeriod.required = false;
  dom.calculateButton.querySelector("span").textContent = state.busy ? "Обробляю…"
    : state.ratingLoading ? "Читаю файл рейтингу…" : "Розрахувати фінальний рейтинг";
  dom.readiness.textContent = reason || (state.reviewsFile
    ? "Готово до розрахунку з файлом переглядів. РЗ не враховується."
    : "Готово до розрахунку з двома файлами. Третій файл не потрібен. РЗ не враховується.");
}

function displayValue(value, column) {
  if (value == null) return "—";
  if (typeof value !== "number") return String(value);
  const percent = column.format === "0.00%";
  const digits = column.format === "0.000" ? 3 : column.format === "0" ? 0 : 2;
  return new Intl.NumberFormat("uk-UA", { style: percent ? "percent" : "decimal",
    minimumFractionDigits: digits, maximumFractionDigits: digits }).format(value);
}

function invalidateResults() {
  dom.results.hidden = true;
  state.originalResults = [];
  state.rankedResults = [];
  state.report = null;
  setProgress("Файли або налаштування змінено. Натисніть «Розрахувати».");
}

function setBusy(busy) {
  state.busy = busy;
  [dom.qualityFile, dom.ratingFile, dom.reviewsFile, dom.sheetSelect, dom.reviewPeriod,
    dom.removeReviews, dom.downloadExcel, dom.downloadScores].forEach(input => { input.disabled = busy; });
  dom.reviewMappingList.querySelectorAll("select").forEach(input => { input.disabled = busy; });
  dom.calculateButton.classList.toggle("loading", busy);
  updateReadyState();
}

function syncPeriod() {
  try {
    dom.reviewPeriod.value = RatingCore.inferPeriod(findWorksheet(state.ratingWorkbook, dom.sheetSelect.value));
  } catch { dom.reviewPeriod.value = ""; }
}

function renderReviews() {
  dom.reviewAudit.replaceChildren();
  dom.reviewDetails.hidden = !state.report;
  dom.reviewSummary.classList.toggle("review-warning", Boolean(state.report &&
    (state.report.unmatched || state.report.invalid || state.report.applied === 0)));
  if (!state.report) {
    dom.reviewSummary.textContent = "Без третього файлу: колонку «перегляд» збережено з вихідного рейтингу.";
    return;
  }
  const r = state.report;
  dom.reviewSummary.textContent = `${r.period}: записів за датою замовлення ${r.selected}; враховано ${r.applied} на ${formatNumber(r.points)} балів. `
    + `Без підтвердження: ${r.pending}; відхилено: ${r.rejected}; не зіставлено: ${r.unmatched}; без числових балів: ${r.invalid - r.unknownDates}. `
    + `Записів із невідомою датою в усьому файлі: ${r.unknownDates}. `
    + `Змін у колонці «перегляд»: ${r.changes.length}. Інші місяці пропущено: ${r.outside}.`
    + (r.applied === 0 ? " Увага: додаткових балів не враховано, у колонці «перегляд» тепер нулі." : "");
  const lines = r.changes.map(item => `${item.fio}: було ${item.before}, стало ${item.after}.`);
  for (const item of r.audit.filter(item => item.status !== "Враховано" && item.status !== "Відхилено")) {
    lines.push(`Рядок ${item.row} · ${item.fio || "ПІБ відсутній"} · ${item.date?.iso || item.rawDate} · ${item.status}`);
  }
  dom.reviewAudit.textContent = lines.join("\n") || "Розбіжностей немає. Повний журнал врахування — на окремому листі завантаженого Excel.";
}

function resetReviewMapping() {
  state.reviewOverrides = Object.create(null);
  dom.reviewMappingList.replaceChildren();
  dom.reviewMapping.hidden = true;
}

function renderReviewMapping() {
  const uncertain = new Map();
  for (const entry of state.report?.audit || []) {
    if (entry.status.startsWith("Не знайдено") || entry.manual) uncertain.set(normalizeText(entry.fio), entry.fio);
  }
  dom.reviewMappingList.replaceChildren();
  dom.reviewMapping.hidden = uncertain.size === 0;
  for (const [identity, fio] of uncertain) {
    const label = document.createElement("label");
    label.className = "mapping-item";
    const text = document.createElement("span");
    text.textContent = fio || "ПІБ відсутній";
    const select = document.createElement("select");
    select.setAttribute("aria-label", `Консьєрж для переглядів: ${fio || "ПІБ відсутній"}`);
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "Не враховувати — оберіть працівника, якщо впевнені";
    select.appendChild(empty);
    for (const row of state.originalResults) {
      const option = document.createElement("option");
      option.value = String(row.sourceRow);
      option.textContent = row.qualityFio || row.fio;
      select.appendChild(option);
    }
    select.value = String(state.reviewOverrides[identity] || "");
    select.addEventListener("change", () => {
      if (select.value) state.reviewOverrides[identity] = Number(select.value);
      else delete state.reviewOverrides[identity];
      invalidateResults();
      updateReadyState();
    });
    label.append(text, select);
    dom.reviewMappingList.appendChild(label);
  }
}

function setFileCard(card, nameNode, stateNode, file) {
  card.classList.toggle("ready", Boolean(file));
  nameNode.textContent = file ? file.name : "Файл не обрано";
  stateNode.textContent = file ? "Готово ✓" : "Обрати";
}

function setProgress(message, isError = false) {
  dom.progress.textContent = message;
  dom.progress.classList.toggle("error", isError);
}

function showToast(message) {
  dom.toast.textContent = message;
  dom.toast.classList.add("show");
  window.setTimeout(() => dom.toast.classList.remove("show"), 2200);
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadFinalExcel() {
  if (!state.rankedResults.length || state.busy) return;
  setBusy(true);
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Quality Rating";
    const worksheet = workbook.addWorksheet("Фінальний рейтинг", {
      views: [{ state: "frozen", ySplit: 1, xSplit: 1 }],
      pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    worksheet.columns = state.columns.map(column => ({ header: column.title, key: column.key, width: column.width }));

    const header = worksheet.getRow(1);
    header.height = 44;
    header.font = { color: { argb: "FF17251B" }, bold: true, size: 10 };
    header.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDF5EF" } };

    for (const item of state.rankedResults) {
      const row = worksheet.addRow(state.columns.map(column => item[column.key] ?? null));
      row.height = 18;
      row.font = { size: 10 };
      row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.zone.color } };
      state.columns.forEach((column, index) => {
        const cell = row.getCell(index + 1);
        cell.numFmt = column.format || "0.00";
        cell.alignment = { horizontal: typeof item[column.key] === "number" ? "center" : "left", vertical: "middle" };
        cell.border = { bottom: { style: "hair", color: { argb: "55808080" } }, right: { style: "hair", color: { argb: "55808080" } } };
        if (column.key === "status" && item.issues.length) cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 10 };
        if (column.key === "reviews" && Math.abs(item.previousReviews - item.reviews) > 1e-9) {
          cell.note = `У вихідному рейтингу: ${item.previousReviews}. Із файлу переглядів: ${item.reviews}.`;
        }
      });
    }

    worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: worksheet.rowCount, column: state.columns.length } };
    const legend = workbook.addWorksheet("Легенда");
    legend.columns = [{ header: "Зона", width: 18 }, { header: "Частина рейтингу", width: 22 }];
    legend.getRow(1).font = { color: { argb: "FFFFFFFF" }, bold: true };
    legend.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
    ZONES.forEach(zone => {
      const row = legend.addRow([zone.name, zone.label]);
      row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: zone.color } };
    });
    legend.addRow([]);
    legend.addRow(["Джерело", state.sourceSheet]);
    legend.addRow(["Перегляди", state.report ? state.report.period : "Збережено з рейтингу"]);
    legend.addRow(["База ефективності", `${state.baseLabel} вихідного листа; максимум ${state.maximum}`]);
    legend.addRow(["Формула фіналу", "КК в рейтингу + РЗ в рейтингу + рейтингу (M + K + I)"]);
    legend.addRow(["РЗ", "РЗ = 0; РЗ в рейтингу = 0. Дані з іншого джерела не імпортуються"]);
    legend.addRow(["Внески", "КК × 0,4 + 0 + ефективність × 0,4. Вага РЗ не перерозподіляється"]);
    legend.addRow(["КК / КЯ", "У колонку КК підставлено середню оцінку КЯ зі statistics.xlsx"]);
    legend.addRow(["Однакові бали", "Однакове місце й зона; через округлення та нічиї частки зон приблизні"]);
    legend.addRow(["Немає КЯ", "Попередній результат: відсутній внесок КЯ дорівнює 0"]);
    legend.addRow(["Формат результату", "Значення на момент розрахунку; щоб оновити, перерахуйте сайт із новими файлами"]);
    if (state.report) {
      const report = state.report;
      const audit = workbook.addWorksheet("Перегляди — журнал");
      audit.columns = [
        { header: "Рядок джерела", width: 16 }, { header: "ПІБ у переглядах", width: 40 },
        { header: "Дата замовлення", width: 22 }, { header: "Рішення", width: 18 },
        { header: "Додаткові бали", width: 18 }, { header: "ПІБ у рейтингу", width: 28 },
        { header: "Результат врахування", width: 65 }, { header: "Зіставлення", width: 24 },
      ];
      report.audit.forEach(entry => audit.addRow([entry.row, entry.fio, entry.date?.iso || entry.rawDate,
        entry.decision, entry.points, entry.target, entry.status, entry.manual ? "Вибрано користувачем" : ""]));
      audit.getRow(1).font = { bold: true };
      audit.views = [{ state: "frozen", ySplit: 1 }];
      audit.autoFilter = `A1:H${audit.rowCount}`;
      audit.getColumn(5).numFmt = "0.000";
      const changes = workbook.addWorksheet("Зміни переглядів");
      changes.columns = [{ header: "ПІБ КС", width: 28 }, { header: "Було в рейтингу", width: 22 },
        { header: "Із переглядів", width: 22 }, { header: "Різниця", width: 18 }];
      report.changes.forEach(entry => changes.addRow([entry.fio, entry.before, entry.after, entry.after - entry.before]));
      changes.getRow(1).font = { bold: true };
      [2, 3, 4].forEach(c => { changes.getColumn(c).numFmt = "0.000"; });
      legend.addRow(["Записів за місяць", report.selected]);
      legend.addRow(["Враховано записів", report.applied]);
      legend.addRow(["Враховано балів", report.points]);
      legend.addRow(["Не зіставлено", report.unmatched]);
      legend.addRow(["Без підтвердження", report.pending]);
      legend.addRow(["Помилки дати / балів", report.invalid]);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "current_final_rating.xlsx");
    showToast("Excel завантажено ✓");
  } catch (error) {
    setProgress(error.message || String(error), true);
  } finally {
    setBusy(false);
  }
}

function downloadScoresTxt() {
  if (!state.originalResults.length || state.busy) return;
  const text = state.originalResults
    .map(item => item.qualityScore === null ? "" : String(item.qualityScore).replace(".", ","))
    .join("\r\n");
  saveBlob(new Blob([`\uFEFF${text}\r\n`], { type: "text/plain;charset=utf-8" }), "result_only_scores.txt");
  showToast("TXT завантажено ✓");
}

dom.qualityFile.addEventListener("change", event => {
  invalidateResults();
  resetReviewMapping();
  state.qualityFile = event.target.files?.[0] || null;
  setFileCard(dom.qualityCard, dom.qualityName, dom.qualityState, state.qualityFile);
  updateReadyState();
});

dom.ratingFile.addEventListener("change", async event => {
  invalidateResults();
  resetReviewMapping();
  const version = ++state.ratingVersion;
  state.ratingFile = event.target.files?.[0] || null;
  state.ratingLoading = Boolean(state.ratingFile);
  state.ratingWorkbook = null;
  dom.sheetRow.hidden = true;
  setFileCard(dom.ratingCard, dom.ratingName, dom.ratingState, state.ratingFile);
  dom.ratingCard.classList.remove("ready");
  updateReadyState();
  if (!state.ratingFile) return;

  try {
    setProgress("Читаю листи великого файлу rating.xlsx…");
    dom.ratingState.textContent = "Читаю…";
    const workbook = await loadWorkbook(state.ratingFile);
    if (version !== state.ratingVersion) return;
    state.ratingWorkbook = workbook;
    const finalSheets = state.ratingWorkbook.worksheets
      .map(sheet => sheet.name)
      .filter(name => normalizeText(name).startsWith("рейтинг фінал"));
    if (!finalSheets.length) throw new Error("Не знайдено жодного листа «Рейтинг фінал (…)». ");

    dom.sheetSelect.replaceChildren();
    finalSheets.forEach(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      dom.sheetSelect.appendChild(option);
    });
    const currentMonth = UA_MONTHS[new Date().getMonth()];
    const preferred = finalSheets.find(name => normalizeText(name).includes(normalizeText(`(${currentMonth})`)));
    dom.sheetSelect.value = preferred || finalSheets.at(-1);
    syncPeriod();
    dom.sheetRow.hidden = false;
    dom.ratingState.textContent = "Готово ✓";
    dom.ratingCard.classList.add("ready");
    setProgress(`Знайдено фінальних листів: ${finalSheets.length}`);
  } catch (error) {
    if (version !== state.ratingVersion) return;
    state.ratingWorkbook = null;
    dom.ratingCard.classList.remove("ready");
    dom.ratingState.textContent = "Помилка";
    setProgress(error.message || String(error), true);
  }
  if (version === state.ratingVersion) {
    state.ratingLoading = false;
    updateReadyState();
  }
});

dom.sheetSelect.addEventListener("change", () => {
  invalidateResults();
  resetReviewMapping();
  syncPeriod();
  updateReadyState();
});

dom.reviewPeriod.addEventListener("change", () => { invalidateResults(); resetReviewMapping(); updateReadyState(); });

dom.reviewsFile.addEventListener("change", async event => {
  invalidateResults();
  resetReviewMapping();
  const version = ++state.reviewsVersion;
  state.reviewsFile = event.target.files?.[0] || null;
  state.reviewsLoading = Boolean(state.reviewsFile);
  state.reviewsData = null;
  dom.reviewOptions.hidden = !state.reviewsFile;
  dom.removeReviews.hidden = !state.reviewsFile;
  setFileCard(dom.reviewsCard, dom.reviewsName, dom.reviewsState, state.reviewsFile);
  dom.reviewsCard.classList.remove("ready");
  updateReadyState();
  if (!state.reviewsFile) return;
  try {
    dom.reviewsState.textContent = "Читаю…";
    setProgress("Читаю файл переглядів…");
    const workbook = await loadWorkbook(state.reviewsFile);
    if (version !== state.reviewsVersion) return;
    state.reviewsData = RatingCore.readReviews(workbook);
    dom.reviewsState.textContent = "Готово ✓";
    dom.reviewsCard.classList.add("ready");
    setProgress(`Перегляди прочитано. Перевірте місяць і натисніть «Розрахувати».`);
  } catch (error) {
    if (version !== state.reviewsVersion) return;
    dom.reviewsCard.classList.remove("ready");
    dom.reviewsState.textContent = "Помилка";
    setProgress(error.message || String(error), true);
  }
  if (version === state.reviewsVersion) {
    state.reviewsLoading = false;
    updateReadyState();
  }
});

dom.removeReviews.addEventListener("click", () => {
  resetReviewMapping();
  state.reviewsVersion++;
  state.reviewsFile = null;
  state.reviewsData = null;
  state.reviewsLoading = false;
  dom.reviewsFile.value = "";
  dom.reviewOptions.hidden = true;
  dom.removeReviews.hidden = true;
  setFileCard(dom.reviewsCard, dom.reviewsName, dom.reviewsState, null);
  invalidateResults();
  updateReadyState();
});

dom.calculateButton.addEventListener("click", async () => {
  const reason = calculationBlockReason();
  if (reason) { updateReadyState(); setProgress(reason, true); return; }
  setBusy(true);
  dom.results.hidden = true;

  try {
    setProgress("Читаю statistics.xlsx та зіставляю працівників…");
    await new Promise(resolve => window.setTimeout(resolve, 30));
    const { scores, duplicates } = await readQualityScores(state.qualityFile);
    state.originalResults = readRating(state.ratingWorkbook, dom.sheetSelect.value, scores);
    state.rankedResults = rankResults(state.originalResults);
    state.sourceSheet = dom.sheetSelect.value;
    renderResults(dom.sheetSelect.value);
    const duplicateNote = duplicates.length ? ` Дублікатів ПІБ у статистиці: ${duplicates.length}; їхні КЯ не підставлено.` : "";
    setProgress(`Готово. Оброблено ${state.originalResults.length} працівників.${duplicateNote}`);
  } catch (error) {
    setProgress(error.message || String(error), true);
  } finally {
    setBusy(false);
  }
});

dom.downloadExcel.addEventListener("click", downloadFinalExcel);
dom.downloadScores.addEventListener("click", downloadScoresTxt);
updateReadyState();
