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

  const themeColors = new WeakMap();
  function isGrayFill(fill, workbook) {
    if (fill?.type !== "pattern" || fill.pattern !== "solid") return false;
    const color = fill.fgColor || fill.bgColor;
    if (!color) return false;
    let rgb = color.argb?.slice(-6);
    if (!rgb && color.indexed !== undefined) {
      // Grayscale entries of Excel's standard indexed palette. System colors
      // and unfilled/white cells are not exclusion markers.
      rgb = { 0: "000000", 1: "FFFFFF", 8: "000000", 9: "FFFFFF",
        22: "C0C0C0", 23: "808080", 55: "969696", 63: "333333" }[color.indexed];
    }
    if (!rgb && color.theme !== undefined) {
      if (!themeColors.has(workbook)) {
        const colors = ["FFFFFF", "000000", "EEECE1", "1F497D"];
        // ExcelJS 4.4 preserves the original theme XML on the workbook.
        const xml = workbook._themes?.theme1;
        if (xml) {
          const doc = new DOMParser().parseFromString(xml, "application/xml");
          const scheme = doc.getElementsByTagNameNS("*", "clrScheme")[0];
          ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink"].forEach((name, index) => {
            const node = scheme?.getElementsByTagNameNS("*", name)[0]?.firstElementChild;
            if (node) colors[index] = node.getAttribute("lastClr") || node.getAttribute("val");
          });
        }
        themeColors.set(workbook, colors);
      }
      rgb = themeColors.get(workbook)[color.theme];
    }
    if (!/^[0-9a-f]{6}$/i.test(rgb || "")) return false;
    const tint = Number(color.tint) || 0;
    const channels = rgb.match(/../g).map(hex => {
      const channel = parseInt(hex, 16);
      return Math.round(tint < 0 ? channel * (1 + tint) : channel + (255 - channel) * tint);
    });
    const darkest = Math.min(...channels), lightest = Math.max(...channels);
    // Includes B7B7B7, CCCCCC and D9D9D9 used in the source books;
    // excludes black, white and near-white alternating table backgrounds.
    return darkest >= 48 && lightest <= 230 && lightest - darkest <= 8;
  }

  function columnNumber(letters) {
    return [...letters.toUpperCase()].reduce((result, letter) => result * 26 + letter.charCodeAt(0) - 64, 0);
  }

  function formattingRange(ref) {
    const match = ref.match(/^\$?([A-Z]+)(?:\$?(\d+))?(?::\$?([A-Z]+)(?:\$?(\d+))?)?$/i);
    if (!match) return null;
    return { left: columnNumber(match[1]), top: Number(match[2] || 1),
      right: columnNumber(match[3] || match[1]), bottom: Number(match[4] || match[2] || 1048576) };
  }

  function formattingOperand(text, ws, row, column, anchor) {
    const token = text.trim();
    if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(token)) return Number(token);
    if (/^"(?:[^"]|"")*"$/.test(token)) return token.slice(1, -1).replace(/""/g, '"');
    if (/^(TRUE|FALSE)(\(\))?$/i.test(token)) return /^TRUE/i.test(token);
    // Google Sheets exports row-color rules as $O:$O=57. They refer to
    // the current row's cached O value, before the calculator sorts anything.
    const whole = token.match(/^(\$?)([A-Z]+):\$?\2$/i);
    const cell = token.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
    if (!whole && !cell) return undefined;
    const ref = whole || cell;
    const c = columnNumber(ref[2]) + (ref[1] ? 0 : column - anchor.left);
    const r = whole ? row : Number(cell[4]) + (cell[3] ? 0 : row - anchor.top);
    if (c < 1 || c > 16384 || r < 1 || r > 1048576) return undefined;
    const source = ws.getCell(r, c);
    const cached = value(source);
    if (source.type === ExcelJS.ValueType.Formula && cached === null) return undefined;
    return cached ?? "";
  }

  function formattingComparison(left, operator, right) {
    if (left === undefined || right === undefined) return null;
    const a = number(left), b = number(right);
    const x = a !== null && b !== null ? a : norm(left);
    const y = a !== null && b !== null ? b : norm(right);
    switch (operator) {
      case "=": return x === y;
      case "<>": return x !== y;
      case "<": return x < y;
      case ">": return x > y;
      case "<=": return x <= y;
      case ">=": return x >= y;
      default: return null;
    }
  }

  function formattingMatches(rule, ws, row, column, anchor) {
    const operand = text => formattingOperand(String(text ?? ""), ws, row, column, anchor);
    if (rule.type === "expression") {
      const formula = String(rule.formulae?.[0] ?? "").trim().replace(/^=/, "");
      const comparison = formula.match(/^(.+?)(<=|>=|<>|=|<|>)(.+)$/);
      if (comparison) return formattingComparison(operand(comparison[1]), comparison[2], operand(comparison[3]));
      const constant = operand(formula);
      return typeof constant === "boolean" ? constant : typeof constant === "number" ? constant !== 0 : null;
    }
    if (rule.type === "cellIs") {
      const operator = { equal: "=", notEqual: "<>", lessThan: "<", greaterThan: ">",
        lessThanOrEqual: "<=", greaterThanOrEqual: ">=" }[rule.operator];
      return formattingComparison(value(ws.getCell(row, column)) ?? "", operator, operand(rule.formulae?.[0]));
    }
    if (["containsBlanks", "notContainsBlanks"].includes(rule.type)) {
      const blank = !String(value(ws.getCell(row, column)) ?? "").trim();
      return rule.type === "containsBlanks" ? blank : !blank;
    }
    return null;
  }

  function grayRowReader(ws, layout) {
    const column = layout.cols.fio;
    const rules = (ws.conditionalFormattings || []).flatMap(format => {
      const ranges = String(format.ref || "").split(/\s+/).map(formattingRange).filter(Boolean);
      return (format.rules || []).filter(rule => rule.style?.fill || rule.stopIfTrue)
        .map(rule => ({ rule, ranges, anchor: ranges[0] }));
    }).sort((a, b) => (a.rule.priority ?? Infinity) - (b.rule.priority ?? Infinity));
    return row => {
      const baseFill = ws.getCell(row, column).fill;
      const matching = rules.filter(({ ranges }) => ranges.some(range =>
        row >= range.top && row <= range.bottom && column >= range.left && column <= range.right));
      if (!isGrayFill(baseFill, ws.workbook) && !matching.some(({ rule }) => isGrayFill(rule.style?.fill, ws.workbook))) return false;
      for (const { rule, anchor } of matching) {
        const applies = formattingMatches(rule, ws, row, column, anchor);
        if (applies === null) {
          throw new Error(`Не вдалося визначити сіре позначення в рядку ${row}. Перерахуйте та збережіть Excel. Якщо правило форматування не підтримується, використайте копію зі звичайною заливкою рядків замість умовного форматування.`);
        }
        if (!applies) continue;
        if (rule.style?.fill) return isGrayFill(rule.style.fill, ws.workbook);
        if (rule.stopIfTrue) break;
      }
      return isGrayFill(baseFill, ws.workbook);
    };
  }

  function readRows(ws, layout) {
    const rows = [];
    const isGrayRow = grayRowReader(ws, layout);
    ws.eachRow({ includeEmpty: false }, row => {
      if (row.number <= layout.row) return;
      const fio = String(value(row.getCell(layout.cols.fio)) || "").trim();
      if (!fio) return;
      const item = { sourceRow: row.number, fio, issues: [], excludedFromRating: isGrayRow(row.number) };
      for (const def of layout.columns.slice(1)) {
        // Knowledge level comes from another source and is out of scope here.
        // Keep both columns, but never import old RZ values or their cached contribution.
        if (def.key === "rz" || def.key === "rzPart") { item[def.key] = 0; continue; }
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
    const participants = rows.filter(item => !item.excludedFromRating);
    const reference = participants.filter(item => item.sourceRow >= base.first && item.sourceRow <= base.last);
    if (!base.constant && reference.some(item => item.pointsPerHour === null)) {
      throw new Error("У базовій групі ефективності є порожні бали або непозитивні години. Перевірте вихідний рейтинг.");
    }
    const maximum = participants.length ? base.constant ?? Math.max(0, ...reference.map(item => item.pointsPerHour)) : null;
    if (participants.length && !(maximum > 0)) throw new Error("Немає додатної бази ефективності серед учасників вихідної групи MAX. Перевірте бали, години та сірі позначення рейтингу.");
    for (const item of rows) {
      item.efficiency = item.pointsPerHour === null || maximum === null ? null : item.pointsPerHour / maximum * 100;
      item.pointsPart = item.efficiency === null ? null : item.efficiency * 0.4;
      item.rz = 0;
      item.rzPart = 0;
      item.qualityPart = (item.qualityScore ?? 0) * 0.4;
      // Final = "КК в рейтингу" + "РЗ в рейтингу" + "рейтингу" (M + K + I).
      // The missing 20% RZ contribution is NOT redistributed to the other components.
      item.finalScore = item.pointsPart === null ? null : item.qualityPart + item.rzPart + item.pointsPart;
      if (item.pointsPerHour === null) item.issues.push("Перевірте бали / години; місце не визначено");
      if (item.qualityScore === null) item.issues.push("НЕМАЄ КЯ (внесок 0)");
      item.status = [item.excludedFromRating ? "Не рейтингується — сірий рядок у вихідному файлі" : "", ...item.issues]
        .filter(Boolean).join("; ") || "Знайдено";
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
