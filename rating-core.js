"use strict";

// Only the known rating arithmetic is evaluated. Excel formulas are never executed.
window.RatingCore = (() => {
  const norm = value => String(value ?? "").trim().toLocaleLowerCase("uk-UA")
    .replace(/[’'`ʼ‘´]/g, "").replace(/ё/g, "е").replace(/\s+/g, " ");
  function value(cell) {
    // ExcelJS 4.4's cell.value copy omits falsy formula results (notably 0).
    // The public result getter retains them. Do not turn actual missing caches into zero.
    if (cell?.type === ExcelJS.ValueType.Formula) return cell.result ?? null;
    const v = cell?.value;
    if (v == null) return null;
    if (v instanceof Date || typeof v !== "object") return v;
    if (Object.hasOwn(v, "result")) return v.result;
    if (v.richText) return v.richText.map(part => part.text).join("");
    return v.text ?? null;
  }
  function number(v) {
    if (v == null || typeof v === "boolean" || String(v).trim() === "") return null;
    const n = typeof v === "number" ? v : Number(String(v).replace(/\s/g, "").replace(",", "."));
    return Number.isFinite(n) ? n : null;
  }
  const columns = [
    { key: "fio", title: "ПІБ КС", aliases: ["ПІБ КС", "ПІБ"], width: 27 },
    { key: "earnedPoints", title: "набрано балів", width: 13 },
    { key: "reviews", title: "перегляд", aliases: ["перегляд", "перегляди"], format: "0.000", width: 13 },
    { key: "hours", title: "години (графік)", width: 13 },
    { key: "outsideHours", title: "поза рейтингом", width: 13 },
    { key: "ratingHours", title: "години в рейтинг", width: 13 },
    { key: "pointsPerHour", title: "бали на годину", format: "0.00%", width: 13 },
    { key: "efficiency", title: "ефективність", width: 14 },
    { key: "pointsPart", title: "рейтингу", aliases: ["рейтингу", "бали в рейтингу"], width: 13 },
    { key: "rz", title: "РЗ", width: 12 },
    { key: "rzPart", title: "РЗ в рейтингу", width: 13 },
    { key: "qualityScore", title: "КК", aliases: ["КК", "КЯ"], width: 12 },
    { key: "qualityPart", title: "КК в рейтингу", aliases: ["КК в рейтингу", "КЯ в рейтингу"], width: 13 },
    { key: "finalScore", title: "Фінал", format: "0.000", width: 14 },
  ];

  function header(ws, definitions) {
    for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
      const found = {};
      ws.getRow(r).eachCell({ includeEmpty: false }, (cell, c) => {
        for (const def of definitions) {
          if ((def.aliases || [def.title]).some(name => norm(name) === norm(value(cell)))) found[def.key] = c;
        }
      });
      if (definitions.every(def => found[def.key])) return { row: r, cols: found };
    }
    return null;
  }

  function ratingLayout(ws) {
    // Older sheets sometimes have no label in N; all other columns are required.
    const layout = header(ws, columns.slice(0, 13));
    if (!layout) throw new Error("Не знайдено повну таблицю рейтингу (ПІБ, бали, перегляд, години, РЗ, КК). Оберіть лист «Рейтинг фінал».");
    layout.cols.finalScore = layout.cols.qualityPart + 1;
    layout.columns = columns.map(def => ({ ...def,
      title: String(value(ws.getCell(layout.row, layout.cols[def.key])) || def.title).trim(),
    }));
    ws.getRow(layout.row).eachCell({ includeEmpty: false }, (cell, c) => {
      if (c <= layout.cols.finalScore || ["місце", "статус кя"].includes(norm(value(cell)))) return;
      const key = `extra${c}`;
      layout.cols[key] = c;
      layout.columns.push({ key, title: String(value(cell)), width: 18 });
    });
    return layout;
  }

  function inferPeriod(ws) {
    if (!ws) return "";
    const layout = ratingLayout(ws);
    let period = "";
    ws.eachRow({ includeEmpty: false }, row => {
      if (period || row.number <= layout.row) return;
      const f = row.getCell(layout.cols.outsideHours).formula || "";
      const match = f.match(/>=\s*0?1\.(\d{1,2})\.(\d{4})/);
      if (match) period = `${match[2]}-${match[1].padStart(2, "0")}`;
    });
    return period;
  }

  function readRows(ws, layout) {
    const rows = [];
    ws.eachRow({ includeEmpty: false }, row => {
      if (row.number <= layout.row) return;
      const fio = String(value(row.getCell(layout.cols.fio)) || "").trim();
      if (!fio) return;
      const item = { sourceRow: row.number, fio, issues: [] };
      for (const def of layout.columns.slice(1)) {
        const raw = value(row.getCell(layout.cols[def.key]));
        item[def.key] = def.key.startsWith("extra") ? raw : number(raw);
      }
      item.previousReviews = item.reviews ?? 0;
      rows.push(item);
    });
    if (!rows.length) throw new Error("У листі рейтингу немає працівників.");
    return rows;
  }

  function normalization(ws, layout) {
    const efficiencyCell = ws.getCell(layout.row + 1, layout.cols.efficiency);
    const f = (efficiencyCell.formula || "").replace(/\s/g, "");
    // Example H3=G3/$K$1*100. Preserve the original reference group, NOT global MAX.
    const ref = f.match(/^=?\$?[A-Z]+\$?\d+\/(\$?[A-Z]+\$?\d+)\*100$/i);
    if (!ref) throw new Error("Не вдалося визначити формулу ефективності. Потрібна формула на кшталт G3/$K$1*100 у вихідному Excel.");
    const base = ws.getCell(ref[1].replace(/\$/g, ""));
    const max = (base.formula || "").replace(/\s/g, "").match(/^=?MAX\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)$/i);
    if (max) {
      const col = ws.getColumn(layout.cols.pointsPerHour).letter;
      if (max[1].toUpperCase() !== col || max[3].toUpperCase() !== col) throw new Error("Невідомий діапазон MAX для ефективності.");
      return { first: Number(max[2]), last: Number(max[4]), label: `${col}${max[2]}:${col}${max[4]}` };
    }
    if (!base.formula && number(value(base)) > 0) return { constant: number(value(base)), label: base.address };
    throw new Error("Невідома формула бази ефективності. Розрахунок зупинено, щоб не підмінити правила рейтингу.");
  }

  function recalculate(rows, base) {
    for (const item of rows) {
      item.ratingHours = item.hours === null ? null : item.hours - (item.outsideHours ?? 0);
      item.pointsPerHour = item.ratingHours > 0 && item.earnedPoints !== null
        ? (item.earnedPoints + (item.reviews ?? 0)) / item.ratingHours : null;
    }
    const reference = rows.filter(item => item.sourceRow >= base.first && item.sourceRow <= base.last);
    if (!base.constant && reference.some(item => item.pointsPerHour === null)) {
      throw new Error("У базовій групі ефективності є порожні бали або непозитивні години. Перевірте вихідний рейтинг.");
    }
    const maximum = base.constant ?? Math.max(0, ...reference.map(item => item.pointsPerHour));
    if (!(maximum > 0)) throw new Error("База ефективності дорівнює нулю: перевірте бали й години рейтингу.");
    for (const item of rows) {
      item.efficiency = item.pointsPerHour === null ? null : item.pointsPerHour / maximum * 100;
      item.pointsPart = item.efficiency === null ? null : item.efficiency * 0.4;
      item.rzPart = (item.rz ?? 0) * 0.2;
      item.qualityPart = (item.qualityScore ?? 0) * 0.4;
      item.finalScore = item.pointsPart === null ? null : item.pointsPart + item.rzPart + item.qualityPart;
      if (item.pointsPerHour === null) item.issues.push("Перевірте бали / години; місце не визначено");
      if (item.qualityScore === null) item.issues.push("НЕМАЄ КЯ (внесок 0)");
      item.status = item.issues.join("; ") || "Знайдено";
    }
    return maximum;
  }

  function dateParts(raw, date1904 = false) {
    let y, m, d;
    if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
      [y, m, d] = [raw.getUTCFullYear(), raw.getUTCMonth() + 1, raw.getUTCDate()];
    } else if (typeof raw === "number" && Number.isFinite(raw)) {
      return dateParts(new Date(Date.UTC(1899, 11, 30) + (Math.floor(raw) + (date1904 ? 1462 : 0)) * 86400000));
    } else {
      const str = String(raw ?? "").trim();
      const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
      const ua = str.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:\s.*)?$/);
      if (iso) [, y, m, d] = iso.map(Number);
      else if (ua) [, d, m, y] = ua.map(Number);
      else return null;
    }
    const check = new Date(Date.UTC(y, m - 1, d));
    if (check.getUTCFullYear() !== y || check.getUTCMonth() + 1 !== m || check.getUTCDate() !== d) return null;
    return { iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`, period: `${y}-${String(m).padStart(2, "0")}` };
  }

  const reviewHeaders = [
    { key: "fio", aliases: ["ПІБ консьерж", "ПІБ консьєржа", "ПІБ консьєрж", "ПІБ КС", "ПІБ"] },
    { key: "date", aliases: ["Дата замовлення (НЕ БІЛЬШ 5 ДНІВ від дати рішення)", "Дата замовлення"] },
    { key: "decision", aliases: ["Решение", "Рішення"] },
    { key: "points", aliases: ["Количество добавленных баллов", "Кількість додаткових балів", "Додаткові бали"] },
  ];
  function readReviews(workbook) {
    const matches = workbook.worksheets.map(ws => ({ ws, layout: header(ws, reviewHeaders) })).filter(x => x.layout);
    const preferred = matches.filter(x => ["пересмотры", "перегляди"].includes(norm(x.ws.name)));
    const candidates = preferred.length ? preferred : matches;
    if (candidates.length !== 1) throw new Error("У файлі переглядів потрібен один лист «Пересмотры» / «Перегляди» з ПІБ, датою замовлення, рішенням і додатковими балами.");
    const { ws, layout } = candidates[0];
    const entries = [];
    ws.eachRow({ includeEmpty: false }, row => {
      if (row.number <= layout.row) return;
      const get = key => value(row.getCell(layout.cols[key]));
      if (!["fio", "date", "decision", "points"].some(key => get(key) !== null)) return;
      entries.push({ row: row.number, fio: String(get("fio") ?? "").trim(),
        date: dateParts(get("date"), workbook.properties.date1904),
        rawDate: String(get("date") ?? ""), decision: norm(get("decision")), points: number(get("points")),
      });
    });
    return { sheet: ws.name, entries };
  }

  function matchIdentity(fio, identities) {
    const parts = norm(fio).split(" ");
    const exact = identities.filter(name => norm(name) === norm(fio));
    if (exact.length === 1) return exact[0];
    // Never discard a supplied patronymic or guess among namesakes.
    if (parts.length >= 3) return null;
    const matches = identities.filter(name => {
      const candidate = norm(name).split(" ");
      return parts.every((part, index) => candidate[index] === part);
    });
    return matches.length === 1 ? matches[0] : null;
  }

  function applyReviews(rows, reviews, period, overrides = {}) {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(period)) throw new Error("Оберіть місяць і рік переглядів.");
    const identities = [...new Map(reviews.entries.filter(e => e.fio).map(e => [norm(e.fio), e.fio])).values()];
    const targets = new Map();
    for (const item of rows) {
      item.reviews = 0; // Replace this month's existing amount; never double-count it.
      const identity = matchIdentity(item.qualityFio || item.fio, identities);
      if (identity) {
        const key = norm(identity);
        if (!targets.has(key)) targets.set(key, []);
        targets.get(key).push(item);
      }
    }
    for (const [identity, sourceRow] of Object.entries(overrides)) {
      const target = rows.find(item => item.sourceRow === Number(sourceRow));
      if (target) targets.set(identity, [target]);
    }
    const report = { period, sheet: reviews.sheet, selected: 0, approved: 0, applied: 0, pending: 0,
      rejected: 0, outside: 0, invalid: 0, unknownDates: 0, unmatched: 0, points: 0, audit: [], changes: [] };
    for (const entry of reviews.entries) {
      if (entry.date && entry.date.period !== period) { report.outside++; continue; }
      const audit = { ...entry, target: "", status: "" };
      report.audit.push(audit);
      if (!entry.date) { report.invalid++; report.unknownDates++; audit.status = "Невідома дата — не враховано"; continue; }
      report.selected++;
      if (["нет", "ні", "no"].includes(entry.decision)) { report.rejected++; audit.status = "Відхилено"; continue; }
      if (!["да", "так", "yes"].includes(entry.decision)) { report.pending++; audit.status = "Немає підтвердження"; continue; }
      report.approved++;
      if (entry.points === null) { report.invalid++; audit.status = "Немає числових балів — не враховано"; continue; }
      const matches = targets.get(norm(entry.fio)) || [];
      if (matches.length !== 1) {
        report.unmatched++;
        audit.status = "Не знайдено однозначного ПІБ у рейтингу — не враховано";
        continue;
      }
      matches[0].reviews += entry.points;
      report.points += entry.points;
      report.applied++;
      audit.target = matches[0].fio;
      audit.status = "Враховано";
      audit.manual = Object.hasOwn(overrides, norm(entry.fio));
    }
    for (const item of rows) {
      item.reviews = Math.round(item.reviews * 1e9) / 1e9;
      if (Math.abs(item.reviews - item.previousReviews) > 1e-9) report.changes.push({
        fio: item.fio, before: item.previousReviews, after: item.reviews,
      });
    }
    report.points = Math.round(report.points * 1e9) / 1e9;
    return report;
  }
  return { norm, value, number, columns, header, ratingLayout, inferPeriod, readRows, normalization,
    recalculate, dateParts, readReviews, matchIdentity, applyReviews };
})();
