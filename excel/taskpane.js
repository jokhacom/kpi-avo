/*
 * Панель надстройки Unoom KPI в Excel.
 * Вход — тот же Supabase, что у сайта; настройки KPI берутся оттуда же
 * (своя строка kpi_settings, иначе общий шаблон, как на сайте).
 * «Рассчитать» читает активный лист и пишет результат на лист «Сумма KPI».
 */
(function () {
  "use strict";

  const SUPABASE_URL = "https://icigxnjpjnsxnouettvx.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_QD_RtawVbBpEkc9iGt-oYQ_MGEJNkd9";
  const DEFAULT_TEMPLATE_ID = "00000000-0000-0000-0000-000000000000";
  const RESULT_SHEET = "Сумма KPI";

  const $ = (id) => document.getElementById(id);
  const show = (id, on) => $(id).classList.toggle("hidden", !on);
  const fmt = (n) => Math.round(Number(n) || 0).toLocaleString("ru-RU").replace(/,/g, " ");

  let sb = null;
  let state = { session: null, profile: null, settings: null };

  async function loadContext(session) {
    state.session = session;
    if (!session) { state.profile = null; state.settings = null; return render(); }
    try {
      const { data: profile } = await sb.from("profiles").select("*").eq("id", session.user.id).maybeSingle();
      let { data: settings } = await sb.from("kpi_settings").select("*").eq("user_id", session.user.id).maybeSingle();
      if (!settings) ({ data: settings } = await sb.from("kpi_settings").select("*").eq("user_id", DEFAULT_TEMPLATE_ID).maybeSingle());
      state.profile = profile || null;
      state.settings = settings || null;
    } catch (e) {
      state.settings = null;
      $("calcErr").textContent = "Не удалось загрузить настройки: " + (e.message || e);
    }
    render();
  }

  function render() {
    show("boot", false);
    show("login", !state.session);
    show("main", !!state.session);
    if (!state.session) return;
    const u = state.session.user;
    $("who").textContent = u.is_anonymous ? "Гостевой режим" : "Вы вошли как " + (u.email || "");
    const s = state.settings;
    if (!s) {
      $("settingsInfo").textContent = "Настройки KPI не найдены. Задайте их на сайте в разделе «Настройки KPI».";
      $("btnCalc").disabled = true;
      return;
    }
    const inds = (s.indicators || []).filter((i) => i.enabled !== false);
    $("settingsInfo").innerHTML = "";
    const p = document.createElement("div");
    p.textContent = `Показатели: ${inds.map((i) => i.name).join(", ")}. База премии ${fmt(s.base_amount)} сум, норма ${s.norm_hours} ч.`;
    $("settingsInfo").appendChild(p);
    $("btnCalc").disabled = false;
  }

  async function login() {
    $("loginErr").textContent = "";
    $("btnLogin").disabled = true;
    const { error } = await sb.auth.signInWithPassword({ email: $("email").value.trim(), password: $("password").value });
    $("btnLogin").disabled = false;
    if (error) $("loginErr").textContent = "Не удалось войти: " + error.message;
  }
  async function guest() {
    $("loginErr").textContent = "";
    const { error } = await sb.auth.signInAnonymously();
    if (error) $("loginErr").textContent = "Гостевой вход недоступен: " + error.message;
  }

  async function calculate() {
    $("calcErr").textContent = ""; $("calcOk").textContent = ""; $("calcWarn").innerHTML = "";
    const settings = state.settings;
    if (!settings) return;
    const allowOver = state.profile?.role === "super_admin" && settings.pay_overachievement === true;
    const indicators = (settings.indicators || []).filter((i) => i.enabled !== false);
    $("btnCalc").disabled = true;
    try {
      await Excel.run(async (ctx) => {
        const src = ctx.workbook.worksheets.getActiveWorksheet();
        src.load("name");
        const used = src.getUsedRangeOrNullObject(true);
        used.load(["values", "numberFormat"]);
        await ctx.sync();
        if (src.name === RESULT_SHEET) throw new Error(`Откройте лист с исходными данными, а не «${RESULT_SHEET}».`);
        if (used.isNullObject) throw new Error("Лист пустой.");

        const aoa = KpiExcel.toAoa(used.values, used.numberFormat);
        const parsed = KpiExcel.buildRows(aoa, indicators);
        if (!parsed.matchedColumns) throw new Error("В первой строке не найдено ни одного показателя. Ожидаются столбцы: " + indicators.map((i) => i.name).join(", ") + ".");
        if (!parsed.rows.length) throw new Error("Не найдено ни одного сотрудника в столбце A.");

        const results = KpiExcel.calculate(parsed.rows, settings, allowOver);
        const table = KpiExcel.buildResultTable(results, settings);

        const old = ctx.workbook.worksheets.getItemOrNullObject(RESULT_SHEET);
        await ctx.sync();
        if (!old.isNullObject) old.delete();
        const out = ctx.workbook.worksheets.add(RESULT_SHEET);
        const cols = table[0].length;
        const range = out.getRangeByIndexes(0, 0, table.length, cols);
        range.values = table;
        const head = out.getRangeByIndexes(0, 0, 1, cols);
        head.format.font.bold = true;
        head.format.fill.color = "#E8EEFF";
        out.getRangeByIndexes(table.length - 1, 0, 1, cols).format.font.bold = true;
        const moneyFmt = table.slice(1).map(() => ["#,##0"]);
        out.getRangeByIndexes(1, cols - 3, table.length - 1, 1).numberFormat = moneyFmt;
        out.getRangeByIndexes(1, cols - 1, table.length - 1, 1).numberFormat = moneyFmt;
        out.freezePanes.freezeRows(1);
        range.format.autofitColumns();
        out.activate();
        await ctx.sync();

        const totalPremium = table[table.length - 1][cols - 3];
        $("calcOk").textContent = `Готово: ${results.length} сотрудн., премии ${fmt(totalPremium)} сум. Результат на листе «${RESULT_SHEET}».`;
        const notes = [];
        if (parsed.missingIndicators.length) notes.push("Нет столбцов (посчитаны как 0): " + parsed.missingIndicators.join(", "));
        if (parsed.unmatchedColumns.length) notes.push("Не распознаны и пропущены: " + parsed.unmatchedColumns.join(", "));
        notes.push(...parsed.warnings.slice(0, 10));
        if (parsed.warnings.length > 10) notes.push(`…и ещё ${parsed.warnings.length - 10} предупреждений`);
        if (notes.length) {
          const ul = document.createElement("ul");
          notes.forEach((n) => { const li = document.createElement("li"); li.textContent = n; ul.appendChild(li); });
          $("calcWarn").appendChild(ul);
        }
      });
    } catch (e) {
      $("calcErr").textContent = e.message || String(e);
    } finally {
      $("btnCalc").disabled = false;
    }
  }

  if (typeof Office === "undefined") {
    document.getElementById("boot").textContent = "Эта страница работает внутри Excel: откройте её кнопкой «Unoom KPI» на вкладке «Главная».";
    return;
  }
  Office.onReady(async () => {
    if (!window.supabase || !window.KpiCore) { $("boot").textContent = "Не удалось загрузить модули. Проверьте интернет."; return; }
    sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    $("btnLogin").onclick = login;
    $("password").onkeydown = (e) => { if (e.key === "Enter") login(); };
    $("btnGuest").onclick = guest;
    $("btnLogout").onclick = () => sb.auth.signOut();
    $("btnCalc").onclick = calculate;
    sb.auth.onAuthStateChange((_event, session) => {
      if (session?.user?.id !== state.session?.user?.id) setTimeout(() => loadContext(session), 0);
    });
    const { data } = await sb.auth.getSession();
    await loadContext(data.session);
  });
})();
