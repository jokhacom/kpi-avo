/*
 * Логика надстройки Excel без привязки к Office: разбор листа и сборка листа «Сумма KPI».
 * Сам расчёт — window.KpiCore (kpi-core.js), та же формула, что на сайте.
 * Распознавание столбцов повторяет импорт сайта (index.html → buildRowsFromAoa),
 * поэтому файл, который сайт принимает на импорт, надстройка тоже понимает.
 */
(function (global) {
  "use strict";

  const DEFAULT_TAX_RATE = 7.5;
  const HOURS_ALIASES = ["часы", "отработанныечасы", "отработаночасов", "часовотработано", "hours", "hoursworked", "soat", "ishlagansoat"];

  function normalizeHeader(s) {
    return String(s || "").toLowerCase().trim().replace(/[^a-zа-яё0-9]+/gi, "");
  }
  function parseFactValue(raw) {
    if (raw === null || raw === undefined || raw === "") return null;
    if (typeof raw === "number") return raw;
    const s = String(raw).replace(",", ".");
    const m = s.match(/-?\d+(\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }
  // Ячейка в процентном формате хранит долю (0.975 вместо 97.5).
  function readCell(value, numberFormat) {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "number" && String(numberFormat || "").includes("%")) return value * 100;
    return value;
  }
  // values/numberFormats — то, что отдаёт Excel для используемого диапазона листа.
  function toAoa(values, numberFormats) {
    return (values || []).map((row, r) => row.map((v, c) => (r === 0 ? String(v ?? "") : readCell(v, numberFormats?.[r]?.[c]))));
  }
  function findHoursColumn(headerRow) {
    for (let c = 1; c < headerRow.length; c++) if (HOURS_ALIASES.includes(normalizeHeader(headerRow[c]))) return c;
    return -1;
  }
  function matchColumnsToIndicators(headerRow, indicators, excludeIdx) {
    const columns = []; const unmatched = [];
    for (let c = 1; c < headerRow.length; c++) {
      if (c === excludeIdx) continue;
      const norm = normalizeHeader(headerRow[c]);
      if (!norm) continue;
      let match = indicators.find((ind) => normalizeHeader(ind.name) === norm);
      if (!match) match = indicators.find((ind) => normalizeHeader(ind.name).includes(norm) || norm.includes(normalizeHeader(ind.name)));
      if (match) columns.push({ colIdx: c, indicatorId: match.id, indicatorName: match.name });
      else unmatched.push(headerRow[c]);
    }
    const matchedIds = new Set(columns.map((c) => c.indicatorId));
    const missingIndicators = indicators.filter((ind) => !matchedIds.has(ind.id)).map((ind) => ind.name);
    return { columns, unmatched, missingIndicators };
  }
  function buildRows(aoa, indicators) {
    if (!aoa || aoa.length === 0) return { rows: [], unmatchedColumns: [], missingIndicators: indicators.map((i) => i.name), warnings: [] };
    const [headerRow, ...dataRows] = aoa;
    const hoursColIdx = findHoursColumn(headerRow);
    const { columns, unmatched, missingIndicators } = matchColumnsToIndicators(headerRow, indicators, hoursColIdx);
    const warnings = [];
    const rows = [];
    dataRows.forEach((row, i) => {
      const name = String(row[0] ?? "").trim();
      if (!name) return;
      const facts = {};
      columns.forEach(({ colIdx, indicatorId, indicatorName }) => {
        const val = parseFactValue(row[colIdx]);
        if (val === null) { warnings.push(`Строка ${i + 2}, «${name}»: не удалось прочитать «${indicatorName}», поставлено 0`); facts[indicatorId] = 0; }
        else facts[indicatorId] = val;
      });
      const hoursWorked = hoursColIdx !== -1 ? parseFactValue(row[hoursColIdx]) : null;
      rows.push({ name, facts, hoursWorked });
    });
    return { rows, unmatchedColumns: unmatched, missingIndicators, warnings, matchedColumns: columns.length };
  }

  const round1 = (v) => Math.round((Number(v) || 0) * 10) / 10;
  const taxRateOf = (settings) => { const v = Number(settings?.tax_rate); return Number.isFinite(v) && v >= 0 && v <= 100 ? v : DEFAULT_TAX_RATE; };

  // settings — строка kpi_settings с сайта; allowOver — режим «перевыполнение в премии»
  // (на сайте он действует только для Super Admin, сюда передаётся уже проверенным).
  function calculate(rows, settings, allowOver) {
    const core = global.KpiCore;
    const indicators = (settings.indicators || []).filter((i) => i.enabled !== false);
    const taxRate = taxRateOf(settings);
    return rows.map((row) => {
      const { perIndicator, overall, payOverall } = core.computeOverall(indicators, row.facts, settings.formula_type);
      const grade = core.gradeFor(overall, settings.grades || []);
      const hoursWorked = row.hoursWorked ?? settings.norm_hours;
      const premium = core.computePremium(allowOver ? overall : payOverall, hoursWorked, settings.norm_hours, settings.base_amount, allowOver);
      return { name: row.name, perIndicator, overall, payOverall: allowOver ? overall : payOverall, grade, hoursWorked, premium, taxRate, afterTax: premium * (1 - taxRate / 100) };
    });
  }

  // Таблица для листа «Сумма KPI»: по строке на сотрудника + итог.
  function buildResultTable(results, settings) {
    const indicators = (settings.indicators || []).filter((i) => i.enabled !== false);
    const header = ["Сотрудник", ...indicators.map((i) => `${i.name}, %`), "Итоговый KPI, %", "KPI для премии, %", "Оценка", "Часы", "Премия, сум", "Налог, %", "Премия после налога, сум"];
    const body = results.map((r) => [
      r.name,
      ...indicators.map((ind) => round1(r.perIndicator.find((p) => p.id === ind.id)?.percent)),
      round1(r.overall), round1(r.payOverall), r.grade, Number(r.hoursWorked) || 0,
      Math.round(r.premium), r.taxRate, Math.round(r.afterTax),
    ]);
    const total = new Array(header.length).fill("");
    total[0] = "Итого";
    total[header.length - 3] = results.reduce((s, r) => s + Math.round(r.premium), 0);
    total[header.length - 1] = results.reduce((s, r) => s + Math.round(r.afterTax), 0);
    return [header, ...body, total];
  }

  global.KpiExcel = { normalizeHeader, parseFactValue, toAoa, buildRows, calculate, buildResultTable, taxRateOf };
})(typeof window !== "undefined" ? window : globalThis);
