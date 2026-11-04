/*
 * KpiCore — единственное место, где описана логика расчёта KPI.
 * Его подключают и сайт (index.html), и надстройка Excel (excel/taskpane.html).
 * Это гарантирует, что процент по каждому показателю, итоговый KPI и премия
 * всегда считаются одной и той же формулой — без расхождений между разделами.
 */
(function (global) {
  "use strict";

  function clamp01to100(v) {
    return Math.max(0, Math.min(100, v));
  }

  // Потолок для процента ОТДЕЛЬНОГО показателя — 120%, а не 100%. Это специально
  // сделано так, чтобы сильное перевыполнение одной цели могло компенсировать
  // недовыполнение другой при расчёте ОБЩЕГО KPI и буквенной оценки. На сумму
  // премии это не влияет — там используется отдельный потолок в 100%
  // (см. computePremium), деньги от этой компенсации не растут.
  const INDICATOR_PERCENT_CAP = 120;
  function clampIndicatorPercent(v) {
    return Math.max(0, Math.min(INDICATOR_PERCENT_CAP, v));
  }

  // Направление показателя ("больше — лучше" / "меньше — лучше") хранится явно
  // в самом показателе (indicator.direction), а не выводится из его названия —
  // переименование показателя никогда не влияет на расчёт. Для показателей,
  // сохранённых до появления этого поля, направление определяется от чисел
  // (норматив меньше минимума), как раньше — только как запасной вариант.
  function indicatorDirection(indicator) {
    if (indicator.direction === "lower" || indicator.direction === "higher") return indicator.direction;
    return Number(indicator.norm) < Number(indicator.min) ? "lower" : "higher";
  }

  // Процент выполнения ОДНОГО показателя. Результат в диапазоне [0, 120] —
  // это единственное место, где происходит ограничение, поэтому дальнейшие расчёты
  // (общий KPI, буквенная оценка) не должны и не будут повторно "подрезать" значения
  // другой логикой. Премия отдельно ограничивается 100% внутри computePremium.
  function computeIndicatorPercent(indicator, actualRaw, formulaType) {
    const min = Number(indicator.min);
    const norm = Number(indicator.norm);
    const actual = Number(actualRaw) || 0;
    const lowerIsBetter = indicatorDirection(indicator) === "lower";

    if (formulaType === "threshold") {
      const reachedNorm = lowerIsBetter ? actual <= norm : actual >= norm;
      const reachedMin = lowerIsBetter ? actual <= min : actual >= min;
      if (reachedNorm) return 100;
      if (reachedMin) return 50;
      return 0;
    }

    if (min === norm) {
      const reached = lowerIsBetter ? actual <= norm : actual >= norm;
      return reached ? 100 : 0;
    }
    if (lowerIsBetter) {
      if (actual >= min) return 0;
      return clampIndicatorPercent(((min - actual) / (min - norm)) * 100);
    }
    if (actual <= min) return 0;
    return clampIndicatorPercent(((actual - min) / (norm - min)) * 100);
  }

  // Общий KPI считается ИСКЛЮЧИТЕЛЬНО как средневзвешенное уже посчитанных
  // (и уже ограниченных до 100%) процентов показателей — никакой отдельной
  // формулы или повторного применения других правил здесь нет.
  function computeOverall(indicators, facts, formulaType) {
    // ВАЖНО: вес каждого показателя — это его доля из фиксированных 100, а не доля
    // среди только выбранных/активных показателей. Если считается всего один
    // показатель с весом 20%, его вклад в итог — 20%, а не 100%. Веса НИКОГДА
    // не перераспределяются и не нормализуются под то, что сейчас выбрано.
    //
    // Считаются ДВА итога:
    //  • overall    — KPI для буквенной оценки: показатель учитывается до 120%,
    //                 перевыполнение одной цели компенсирует недовыполнение другой;
    //  • payOverall — «денежный» KPI для премии: каждый показатель не выше 100%,
    //                 перевыполнение НЕ добавляет денег и не перекрывает другие показатели.
    let weightedSum = 0;
    let paySum = 0;
    const perIndicator = [];
    indicators.forEach((ind) => {
      if (ind.enabled === false) return; // выключенные показатели не участвуют в расчёте
      const actual = Number(facts[ind.id] ?? 0);
      const percent = computeIndicatorPercent(ind, actual, formulaType); // в [0,120]
      const weight = Number(ind.weight) || 0;
      const contribution = (percent * weight) / 100; // вклад в KPI для оценки
      const payPercent = Math.min(percent, 100); // для денег показатель не выше 100%
      const payContribution = (payPercent * weight) / 100; // вклад в денежный KPI
      weightedSum += contribution;
      paySum += payContribution;
      perIndicator.push({ id: ind.id, name: ind.name, min: ind.min, norm: ind.norm, weight: ind.weight, actual, percent, contribution, payPercent, payContribution });
    });
    const overall = weightedSum; // KPI для оценки (из исходных весов, без перенормировки)
    const payOverall = paySum; // KPI для премии
    return { perIndicator, overall, payOverall };
  }

  // Премия = база × KPI для премии × коэффициент по отработанным часам.
  // Два режима (переключает только Super Admin в настройках/карточке):
  //  • allowOverachievement = false (по умолчанию) — передаётся payOverall из
  //    computeOverall (каждый показатель не выше 100%), итог не больше 100% базы;
  //  • allowOverachievement = true — передаётся overall (показатели до 120%),
  //    перевыполнение увеличивает премию, потолок 100% базы не применяется.
  function computePremium(kpiPercent, hoursWorked, normHours, baseAmount, allowOverachievement) {
    const hours = Number(hoursWorked) || 0;
    const norm = Number(normHours) || 0;
    const hoursCoef = norm > 0 ? Math.min(1, hours / norm) : 1;
    const base = Number(baseAmount || 0);
    const v = Number(kpiPercent) || 0;
    const percent = allowOverachievement ? Math.max(0, Math.min(INDICATOR_PERCENT_CAP, v)) : clamp01to100(v);
    const raw = (percent / 100) * base * hoursCoef;
    return allowOverachievement ? Math.max(0, raw) : Math.max(0, Math.min(raw, base));
  }

  function gradeFor(percent, grades) {
    const sorted = [...grades].sort((a, b) => b.min - a.min);
    for (const g of sorted) if (percent >= Number(g.min)) return g.letter;
    return sorted.length ? sorted[sorted.length - 1].letter : "F";
  }

  global.KpiCore = {
    indicatorDirection,
    computeIndicatorPercent,
    computeOverall,
    computePremium,
    gradeFor,
    INDICATOR_PERCENT_CAP,
  };
})(typeof window !== "undefined" ? window : globalThis);

