/* 统侨科加班统计 — 前端逻辑（依赖 xlsx.js 暴露的 window.XLSXGen）
 * v2：个人登录（按姓名隔离数据）+ 起止时间自动换算时长 + 移动端响应式
 */
(function () {
  'use strict';
  const G = window.XLSXGen;
  const LS_CURRENT = 'tq_current_user';
  const TYPE_LABEL = { weekday: '工作日', weekend: '双休日', holiday: '法定节假日' };

  // ---------- 持久化（按登录人隔离）----------
  function userKey(name) { return 'tq_entries_' + encodeURIComponent(name); }
  function settingsKey(name) { return 'tq_settings_' + encodeURIComponent(name); }

  let currentUser = null;      // {name, dept}
  let entries = [];            // 当前用户的记录
  let settings = {};           // 当前用户的设置
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

  // 起止时间 → 小时（支持跨天：结束<=开始视为跨过午夜）
  function calcHours(start, end) {
    if (!start || !end) return 0;
    const [sh, sm] = start.split(':').map(Number);
    const [eh, em] = end.split(':').map(Number);
    let mins = (eh * 60 + em) - (sh * 60 + sm);
    if (mins < 0) mins += 24 * 60; // 仅结束早于开始时按跨天处理；相等则为 0（无效时长）
    return Math.round((mins / 60) * 100) / 100;
  }
  function updateHoursPreview() {
    const h = calcHours($('fStart').value, $('fEnd').value);
    $('hoursPreview').innerHTML = `本次时长：<span class="num">${fmt(h)}</span> 小时（自动按起止时间换算）`;
    return h;
  }

  // ---------- 设置 ----------
  function applySettingsToUI() {
    if (settings.unit) $('setUnit').value = settings.unit;
    if (settings.month) $('setMonth').value = settings.month;
    if (settings.pubStart) $('setPubStart').value = settings.pubStart;
    if (settings.pubEnd) $('setPubEnd').value = settings.pubEnd;
    if (settings.makeDate) $('setMakeDate').value = settings.makeDate;
  }
  function collectSettings() {
    const s = {
      unit: $('setUnit').value.trim() || '高新区党群工作部',
      month: $('setMonth').value || thisMonth(),
      pubStart: $('setPubStart').value.trim(),
      pubEnd: $('setPubEnd').value.trim(),
      makeDate: $('setMakeDate').value.trim() || todayCN(),
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
    $('loginName').value = ''; $('loginDept').value = '统战';
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
    if (!list.length) { box.innerHTML = '<div class="empty">本月暂无记录，请在上方录入。</div>'; return; }
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
    const start = $('fStart').value;
    const end = $('fEnd').value;
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
    $('fDate').value = ''; $('fStart').value = ''; $('fEnd').value = ''; $('fReason').value = '';
    $('fType').value = 'weekday'; updateHoursPreview();
  }
  function startEdit(id) {
    const e = entries.find(x => x.id === id); if (!e) return;
    editingId = id;
    $('fDate').value = e.date; $('fType').value = e.type;
    $('fStart').value = e.start || ''; $('fEnd').value = e.end || '';
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

  // ---------- 导出 ----------
  function buildCtx() {
    const s = collectSettings();
    const [y, m] = s.month.split('-').map(Number);
    return { year: y, month: m, unit: s.unit, pubStart: s.pubStart || '', pubEnd: s.pubEnd || '', makeDate: s.makeDate, timeStr: s.makeDate };
  }
  function exportXlsx() {
    collectSettings();
    const ctx = buildCtx();
    const list = currentMonthEntries();
    const ym = ctx.year + '年' + ctx.month + '月';
    const res = G.buildMonthWorkbook(ctx, list);
    G.downloadXlsx(res.buf, `统侨加班_${ctx.year}-${pad(ctx.month)}月_公示表+发放表.xlsx`);
    toast(`已导出 ${ym} 表格（${list.length} 条记录）`);
  }
  function exportJson() {
    const data = { version: 2, user: currentUser, exportedAt: new Date().toISOString(), settings, entries };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `统侨加班_${currentUser.name}_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('已导出数据备份');
  }
  function importJson(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (!Array.isArray(data.entries)) throw new Error('格式错误');
        if (data.settings) { settings = data.settings; saveSettings(); }
        const map = {};
        entries.forEach(e => map[e.id] = e);
        data.entries.forEach(e => map[e.id || uid()] = e);
        entries = Object.values(map);
        saveEntries(); renderAll();
        toast(`已导入 ${data.entries.length} 条记录`);
      } catch (e) { alert('导入失败：' + e.message); }
    };
    reader.readAsText(file);
  }

  // ---------- 事件 ----------
  $('btnLogin').onclick = doLogin;
  $('loginName').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
  $('loginDept').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
  $('btnLogout').onclick = logout;
  $('btnAdd').onclick = addOrUpdate;
  $('btnClearForm').onclick = () => { editingId = null; $('btnAdd').textContent = '添加记录'; clearForm(); $('formTip').textContent = ''; };
  $('fStart').addEventListener('input', updateHoursPreview);
  $('fEnd').addEventListener('input', updateHoursPreview);
  $('setMonth').onchange = () => { collectSettings(); renderAll(); };
  $('setUnit').onchange = $('setPubStart').onchange = $('setPubEnd').onchange = $('setMakeDate').onchange = collectSettings;
  $('btnExport').onclick = exportXlsx;
  $('btnExportJson').onclick = exportJson;
  $('btnImportJson').onclick = () => $('fileImport').click();
  $('fileImport').onchange = (e) => { if (e.target.files[0]) importJson(e.target.files[0]); e.target.value = ''; };
  $('btnClearMonth').onclick = () => {
    const m = $('setMonth').value;
    if (!confirm(`确定清空 ${m} 本月所有记录？`)) return;
    entries = entries.filter(e => monthKey(e.date) !== m);
    saveEntries(); renderAll(); toast('已清空本月记录');
  };
  $('btnClearAll').onclick = () => {
    if (!confirm('确定清空本人全部数据？此操作不可恢复，建议先导出备份。')) return;
    entries = []; saveEntries(); renderAll(); toast('已清空本人全部数据');
  };

  // ---------- 启动 ----------
  try {
    const saved = JSON.parse(localStorage.getItem(LS_CURRENT) || 'null');
    if (saved && saved.name) { enterApp(saved); }
    else { $('loginScreen').hidden = false; $('appScreen').hidden = true; $('loginName').focus(); }
  } catch (e) {
    $('loginScreen').hidden = false; $('appScreen').hidden = true;
  }
})();
