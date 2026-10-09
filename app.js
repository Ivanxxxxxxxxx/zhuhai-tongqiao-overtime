/* 统侨科加班统计 — 前端逻辑（依赖 xlsx.js 暴露的 window.XLSXGen 与 vendor/xlsx.full.min.js 的 window.XLSX）
 * v3：个人登录(部门默认统战科、标题随科室) + 起止时间自动换算时长 + 导出/导入 Excel(按模板) + 移动端响应式
 * 补贴规则：不满1小时不计，满1小时后按小时向上取整；工作日15(封顶80)/双休日20(封顶160)/法定30(封顶240)
 */
(function () {
  'use strict';
  const G = window.XLSXGen;
  const XLSX = window.XLSX; // SheetJS，用于导入读取
  const LS_CURRENT = 'tq_current_user';
  const TYPE_LABEL = { weekday: '工作日', weekend: '双休日', holiday: '法定节假日' };

  // ---------- 持久化（按登录人隔离）----------
  function userKey(name) { return 'tq_entries_' + encodeURIComponent(name); }
  function settingsKey(name) { return 'tq_settings_' + encodeURIComponent(name); }

  let currentUser = null;
  let entries = [];
  let settings = {};
  let editingId = null;

  function loadUser(name) {
    try { entries = JSON.parse(localStorage.getItem(userKey(name)) || '[]'); } catch (e) { entries = []; }
    try { settings = JSON.parse(localStorage.getItem(settingsKey(name)) || '{}'); } catch (e) { settings = {}; }
  }
  function saveEntries() { if (currentUser) localStorage.setItem(userKey(currentUser.name), JSON.stringify(entries)); }
  function saveSettings() { if (currentUser) localStorage.setItem(settingsKey(currentUser.name), JSON.stringify(settings)); }

  // ---------- 工具 ----------
  function $(id) { return document.getElementById(id); }
  function toast(msg) {
    const t = $('toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 1800);
  }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function todayCN() { const d = new Date(); return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`; }
  function thisMonth() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; }
  function monthKey(dateStr) { return dateStr.slice(0, 7); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function esc(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
  function fmt(n) { return (Math.round(n * 100) / 100).toString(); }
  function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }

  // Excel 日期序列号 / JS Date → YYYY-MM-DD（SheetJS 默认把日期格转成 JS Date 对象）
  function excelSerialToDate(serial) {
    if (serial == null || serial === '') return '';
    if (serial instanceof Date) { const d = serial; return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
    const s = Number(serial);
    if (!isFinite(s)) {
      const d = new Date(String(serial));
      return isNaN(d) ? '' : d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    }
    const d = new Date((s - 25569) * 86400000);
    if (isNaN(d)) return '';
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }
  // Excel 时间（小数 0.5833=14:00，或 SheetJS 转成的 Date 对象）→ HH:MM；也兼容 "14:00" 文本
  function fracTime(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return pad(v.getHours()) + ':' + pad(v.getMinutes());
    const s = String(v).trim();
    if (/^\d{1,2}:\d{2}/.test(s)) return s.slice(0, 5);
    const f = Number(v);
    if (!isFinite(f) || f <= 0 || f >= 1) return '';
    const totalMin = Math.round(f * 24 * 60);
    return pad(Math.floor(totalMin / 60)) + ':' + pad(totalMin % 60);
  }

  // 起止时间 → 小时（整数，支持跨天）。规则：不满1小时不计；满1小时后按整小时向上取整
  function calcHours(start, end) {
    if (!start || !end) return 0;
    const [sh, sm] = start.split(':').map(Number);
    const [eh, em] = end.split(':').map(Number);
    let mins = (eh * 60 + em) - (sh * 60 + sm);
    if (mins < 0) mins += 24 * 60;
    if (mins < 60) return 0;             // 不满1小时不计
    return Math.ceil(mins / 60);         // 向上取整到整小时（1.1h→2h、2.5h→3h）
  }
  // ===== 原生时间选择器（type=time，移动端滚轮式，记录/导出均为 24 小时制）=====
  function getTimeVal(prefix) {
    const v = $(prefix).value;
    return v || '';
  }
  function setTimeVal(prefix, val) {
    $(prefix).value = val || '';
  }
  function updateHoursPreview() {
    const h = calcHours(getTimeVal('fStart'), getTimeVal('fEnd'));
    let suffix = '小时（自动按起止时间换算，不满1小时不计、满1小时后按整小时向上取整）';
    if (h === 0) suffix = '小时（不满1小时的不计加班，无需录入）';
    $('hoursPreview').innerHTML = `本次时长：<span class="num">${h}</span> ${suffix}`;
    return h;
  }

  // ---------- 设置（仅保留填报单位 + 统计月份）----------
  function applySettingsToUI() {
    if (settings.unit) $('setUnit').value = settings.unit;
    if (settings.month) $('setMonth').value = settings.month;
  }
  function collectSettings() {
    const s = {
      unit: $('setUnit').value.trim() || '高新区党群工作部',
      month: $('setMonth').value || thisMonth(),
    };
    settings = s; saveSettings(); return s;
  }

  // ---------- 登录 / 登出 ----------
  function enterApp(user) {
    currentUser = user;
    localStorage.setItem(LS_CURRENT, JSON.stringify(user));
    loadUser(user.name);
    if (!settings.month) { settings.month = thisMonth(); saveSettings(); }
    $('whoName').textContent = user.name;
    $('whoDept').textContent = user.dept ? '（' + user.dept + '）' : '';
    $('appTitle').textContent = '珠海高新区' + (user.dept || '') + '加班统计与补贴核算';
    $('loginScreen').hidden = true;
    $('appScreen').hidden = false;
    applySettingsToUI();
    renderAll();
  }
  function logout() {
    saveEntries(); saveSettings();
    currentUser = null; entries = []; settings = {};
    localStorage.removeItem(LS_CURRENT);
    $('appScreen').hidden = true;
    $('loginScreen').hidden = false;
    $('loginName').value = ''; $('loginDept').value = '统战科';
    $('loginName').focus();
  }
  function doLogin() {
    const name = $('loginName').value.trim();
    const dept = $('loginDept').value.trim();
    if (!name) { $('loginName').focus(); toast('请填写姓名'); return; }
    enterApp({ name, dept });
    toast('已进入：' + name);
  }

  // ---------- 渲染 ----------
  function currentMonthEntries() {
    const m = ($('setMonth').value || thisMonth());
    return entries.filter(e => monthKey(e.date) === m).sort((a, b) => a.date.localeCompare(b.date));
  }
  function renderEntries() {
    const m = $('setMonth').value || thisMonth();
    $('monthLabel').textContent = m.replace('-', '年') + '月';
    const list = currentMonthEntries();
    const box = $('entryTable');
    if (!list.length) { box.innerHTML = '<div class="empty">本月暂无记录，请在上方录入或从模板导入。</div>'; return; }
    let html = '<table><thead><tr><th>序号</th><th>姓名</th><th>部门</th><th>日期</th><th>类型</th><th>时长(h)</th><th>起</th><th>止</th><th>事由</th><th>操作</th></tr></thead><tbody>';
    list.forEach((e, i) => {
      html += `<tr>
        <td>${i + 1}</td>
        <td class="name">${esc(e.name)}</td>
        <td>${esc(e.dept || '')}</td>
        <td>${e.date}</td>
        <td><span class="tag ${e.type}">${TYPE_LABEL[e.type]}</span></td>
        <td>${fmt(e.hours)}</td>
        <td>${esc(e.start || '')}</td>
        <td>${esc(e.end || '')}</td>
        <td class="reason">${esc(e.reason || '')}</td>
        <td><a class="dl" data-edit="${e.id}">编辑</a> &nbsp; <a class="dl" data-del="${e.id}" style="color:#d23b3b">删除</a></td>
      </tr>`;
    });
    html += '</tbody></table>';
    box.innerHTML = html;
    box.querySelectorAll('[data-del]').forEach(a => a.onclick = () => delEntry(a.getAttribute('data-del')));
    box.querySelectorAll('[data-edit]').forEach(a => a.onclick = () => startEdit(a.getAttribute('data-edit')));
  }
  function renderSummary() {
    const list = currentMonthEntries();
    const box = $('summaryTable');
    if (!list.length) { box.innerHTML = '<div class="empty">本月暂无记录，无法汇总。</div>'; return; }
    const people = {};
    let wkT = 0, weT = 0, hoT = 0, subT = 0;
    for (const e of list) {
      const p = people[e.name] = people[e.name] || { name: e.name, wk: 0, we: 0, ho: 0, sub: 0 };
      if (e.type === 'weekday') p.wk += e.hours; else if (e.type === 'weekend') p.we += e.hours; else p.ho += e.hours;
      const s = G.entrySubsidy(e.type, e.hours);
      p.sub += s; wkT += p.wk; weT += p.we; hoT += p.ho; subT += s;
    }
    const names = Object.keys(people);
    let html = '<table><thead><tr><th>序号</th><th>姓名</th><th>工作日加班(h)</th><th>双休日加班(h)</th><th>法定节假日加班(h)</th><th>累计时长(h)</th><th>加班补贴(元)</th></tr></thead><tbody>';
    names.forEach((nm, i) => {
      const p = people[nm];
      const total = (p.wk + p.we + p.ho);
      html += `<tr><td>${i + 1}</td><td class="name">${esc(p.name)}</td><td>${fmt(p.wk)}</td><td>${fmt(p.we)}</td><td>${fmt(p.ho)}</td><td>${fmt(total)}</td><td>${fmt(p.sub)}</td></tr>`;
    });
    html += `<tr class="summary-total"><td colspan="2">合计</td><td>${fmt(wkT)}</td><td>${fmt(weT)}</td><td>${fmt(hoT)}</td><td>${fmt(wkT + weT + hoT)}</td><td>${fmt(subT)}</td></tr>`;
    html += '</tbody></table>';
    box.innerHTML = html;
  }
  function renderAll() { applySettingsToUI(); renderEntries(); renderSummary(); }

  // ---------- 增改删 ----------
  function addOrUpdate() {
    if (!currentUser) return;
    const date = $('fDate').value;
    const type = $('fType').value;
    const start = getTimeVal('fStart');
    const end = getTimeVal('fEnd');
    const hours = calcHours(start, end);
    const tip = $('formTip');
    if (!date) { tip.textContent = '请选择加班日期'; tip.style.color = '#d23b3b'; return; }
    if (!start || !end) { tip.textContent = '请填写起始时间和终止时间'; tip.style.color = '#d23b3b'; return; }
    if (!(hours > 0)) { tip.textContent = '起止时间换算后时长需大于 0'; tip.style.color = '#d23b3b'; return; }
    tip.textContent = '';
    const rec = {
      id: editingId || uid(),
      name: currentUser.name, dept: currentUser.dept,
      date, type, hours, start, end,
      reason: $('fReason').value.trim(),
    };
    if (editingId) {
      const idx = entries.findIndex(e => e.id === editingId);
      if (idx >= 0) entries[idx] = rec;
      editingId = null; $('btnAdd').textContent = '添加记录';
      toast('已更新记录');
    } else {
      entries.push(rec);
      toast('已添加记录');
    }
    saveEntries();
    const mk = monthKey(date);
    if (mk !== $('setMonth').value) { $('setMonth').value = mk; collectSettings(); }
    clearForm(); renderAll();
  }
  function clearForm() {
    $('fDate').value = ''; setTimeVal('fStart', ''); setTimeVal('fEnd', ''); $('fReason').value = '';
    $('fType').value = 'weekday'; updateHoursPreview();
  }
  function startEdit(id) {
    const e = entries.find(x => x.id === id); if (!e) return;
    editingId = id;
    $('fDate').value = e.date; $('fType').value = e.type;
    setTimeVal('fStart', e.start || ''); setTimeVal('fEnd', e.end || '');
    $('fReason').value = e.reason || '';
    updateHoursPreview();
    $('btnAdd').textContent = '保存修改';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  function delEntry(id) {
    if (!confirm('确定删除这条记录？')) return;
    entries = entries.filter(e => e.id !== id);
    saveEntries(); renderAll(); toast('已删除');
  }

  // ---------- 导出 Excel（基于内置统侨模板填数据）----------
  function b64ToBuf(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const arr = new Uint8Array(len);
    for (let i = 0; i < len; i++) arr[i] = bin.charCodeAt(i);
    return arr.buffer;
  }
  function buildCtx() {
    const s = collectSettings();
    const [y, m] = s.month.split('-').map(Number);
    const today = todayCN();
    return { year: y, month: m, unit: s.unit, pubStart: today, pubEnd: today, makeDate: today, timeStr: today };
  }
  async function exportXlsx() {
    collectSettings();
    const ctx = buildCtx();
    const list = currentMonthEntries();
    if (!list.length) { toast('本月暂无记录，无法导出'); return; }
    if (!window.TQ_TEMPLATE_B64) { alert('模板未加载，请刷新页面后重试'); return; }
    try {
      const tplBuf = b64ToBuf(window.TQ_TEMPLATE_B64);
      const res = await G.fillTemplate(ctx, list, tplBuf);
      G.downloadXlsx(res.buf, `统侨加班_${ctx.year}-${pad(ctx.month)}月_公示表+发放表.xlsx`);
      const ym = ctx.year + '年' + ctx.month + '月';
      toast(`已导出 ${ym} 表格（${list.length} 条记录，补贴合计 ${res.subT} 元）`);
    } catch (err) {
      alert('导出失败：' + (err && err.message ? err.message : err));
    }
  }

  // ---------- 导入 Excel（按统侨模板，自动识别公示表）----------
  function mkEntry(name, dept, date, type, hours, start, end, reason) {
    return { id: uid(), name, dept, date, type, hours: Math.round(hours * 100) / 100, start, end, reason };
  }
  // 解析公示表二维数组 → entries[]
  function parseGongshiRows(rows) {
    let headRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      const joined = r.map(c => String(c)).join(' ');
      if (joined.indexOf('姓名') >= 0 && joined.indexOf('加班日期') >= 0) { headRow = i; break; }
    }
    if (headRow < 0) return [];
    const out = [];
    for (let i = headRow + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.length) continue;
      const name = String(r[2] != null ? r[2] : '').trim();
      if (!name) continue; // 跳过子表头/空行
      const first = String(r[0] != null ? r[0] : '');
      if (name === '小计' || name === '合计' || first.indexOf('制表') >= 0) break; // 到小计/制表人结束
      const dept = String(r[1] != null ? r[1] : '').trim();
      const date = excelSerialToDate(r[3]);
      if (!date) continue;
      const wk = num(r[4]), we = num(r[5]), ho = num(r[6]);
      const start = fracTime(r[7]), end = fracTime(r[8]);
      const reason = String(r[9] != null ? r[9] : '').trim();
      if (wk > 0) out.push(mkEntry(name, dept, date, 'weekday', wk, start, end, reason));
      if (we > 0) out.push(mkEntry(name, dept, date, 'weekend', we, start, end, reason));
      if (ho > 0) out.push(mkEntry(name, dept, date, 'holiday', ho, start, end, reason));
    }
    return out;
  }
  // 按人+月覆盖式合并写入（幂等，不串月、不影响其他人）
  function mergeImported(list) {
    const monthsByName = {};
    for (const e of list) {
      const m = monthKey(e.date);
      (monthsByName[e.name] = monthsByName[e.name] || new Set()).add(m);
    }
    for (const name of Object.keys(monthsByName)) {
      const months = monthsByName[name];
      let arr = [];
      try { arr = JSON.parse(localStorage.getItem(userKey(name)) || '[]'); } catch (e) { arr = []; }
      arr = arr.filter(x => !months.has(monthKey(x.date)));
      arr = arr.concat(list.filter(e => e.name === name));
      localStorage.setItem(userKey(name), JSON.stringify(arr));
    }
    if (currentUser && monthsByName[currentUser.name]) loadUser(currentUser.name);
    return Object.keys(monthsByName).length;
  }
  function importXlsx(file) {
    if (!XLSX) { alert('导入组件未加载，请刷新页面后重试'); return; }
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const wsName = wb.SheetNames.find(n => String(n).indexOf('公示') >= 0) || wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
        const parsed = parseGongshiRows(rows);
        if (!parsed.length) { alert('未在表格中识别到公示表加班记录，请确认文件为统侨模板（含“公示表”工作表）。'); return; }
        const people = mergeImported(parsed);
        toast(`已导入 ${parsed.length} 条记录（${people} 人）`);
        renderAll();
      } catch (err) {
        alert('导入失败：' + (err && err.message ? err.message : err));
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- 事件 ----------
  $('btnLogin').onclick = doLogin;
  $('loginName').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
  $('loginDept').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
  $('btnLogout').onclick = logout;
  $('btnAdd').onclick = addOrUpdate;
  $('btnClearForm').onclick = () => { editingId = null; $('btnAdd').textContent = '添加记录'; clearForm(); $('formTip').textContent = ''; };
  $('fStart').addEventListener('change', updateHoursPreview);
  $('fEnd').addEventListener('change', updateHoursPreview);
  $('setMonth').onchange = () => { collectSettings(); renderAll(); };
  $('setUnit').onchange = collectSettings;
  $('btnExport').onclick = exportXlsx;
  $('btnImportXlsx').onclick = () => $('fileImportXls').click();
  $('fileImportXls').onchange = (e) => { if (e.target.files[0]) importXlsx(e.target.files[0]); e.target.value = ''; };
  $('btnClearMonth').onclick = () => {
    const m = $('setMonth').value;
    if (!confirm(`确定清空 ${m} 本月所有记录？`)) return;
    entries = entries.filter(e => monthKey(e.date) !== m);
    saveEntries(); renderAll(); toast('已清空本月记录');
  };
  $('btnClearAll').onclick = () => {
    if (!confirm('确定清空本人全部数据？此操作不可恢复。')) return;
    entries = []; saveEntries(); renderAll(); toast('已清空本人全部数据');
  };

  // ---------- 初始化 24 小时制时间选择器 ----------
  // 原生 time 输入无需初始化

  // ---------- 启动 ----------
  try {
    const saved = JSON.parse(localStorage.getItem(LS_CURRENT) || 'null');
    if (saved && saved.name) { enterApp(saved); }
    else { $('loginScreen').hidden = false; $('appScreen').hidden = true; $('loginName').focus(); }
  } catch (e) {
    $('loginScreen').hidden = false; $('appScreen').hidden = true;
  }
})();
