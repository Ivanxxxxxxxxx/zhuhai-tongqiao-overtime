/* 统侨科加班统计 — 前端逻辑（依赖 xlsx.js 暴露的 window.XLSXGen） */
(function () {
  'use strict';
  const G = window.XLSXGen;
  const LS_ENTRIES = 'tq_overtime_entries';
  const LS_SETTINGS = 'tq_overtime_settings';
  const TYPE_LABEL = { weekday: '工作日', weekend: '双休日', holiday: '法定节假日' };

  // ---------- 存储 ----------
  function loadEntries() {
    try { return JSON.parse(localStorage.getItem(LS_ENTRIES)) || []; } catch (e) { return []; }
  }
  function saveEntries(arr) { localStorage.setItem(LS_ENTRIES, JSON.stringify(arr)); }
  function loadSettings() {
    try { return JSON.parse(localStorage.getItem(LS_SETTINGS)) || {}; } catch (e) { return {}; }
  }
  function saveSettings(s) { localStorage.setItem(LS_SETTINGS, JSON.stringify(s)); }

  let entries = loadEntries();
  let settings = loadSettings();
  let editingId = null;

  // ---------- 工具 ----------
  function $(id) { return document.getElementById(id); }
  function toast(msg) {
    const t = $('toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 1800);
  }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function todayCN() {
    const d = new Date();
    return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
  }
  function monthKey(dateStr) { return dateStr.slice(0, 7); } // YYYY-MM
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  // ---------- 设置回填 ----------
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
      month: $('setMonth').value || '2026-08',
      pubStart: $('setPubStart').value.trim(),
      pubEnd: $('setPubEnd').value.trim(),
      makeDate: $('setMakeDate').value.trim() || todayCN(),
    };
    settings = s; saveSettings(s); return s;
  }

  // ---------- 渲染 ----------
  function currentMonthEntries() {
    const m = ($('setMonth').value || '2026-08');
    return entries.filter(e => monthKey(e.date) === m).sort((a, b) => a.date.localeCompare(b.date));
  }

  function renderEntries() {
    const m = $('setMonth').value || '2026-08';
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
        <td>${e.hours}</td>
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

  function esc(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
  function fmt(n) { return (Math.round(n * 100) / 100).toString(); }

  // ---------- 增改删 ----------
  function addOrUpdate() {
    const name = $('fName').value.trim();
    const date = $('fDate').value;
    const type = $('fType').value;
    const hours = parseFloat($('fHours').value);
    const tip = $('formTip');
    if (!name) { tip.textContent = '请填写姓名'; tip.style.color = '#d23b3b'; return; }
    if (!date) { tip.textContent = '请选择加班日期'; tip.style.color = '#d23b3b'; return; }
    if (!(hours > 0)) { tip.textContent = '请填写大于0的加班时长'; tip.style.color = '#d23b3b'; return; }
    tip.textContent = '';
    const rec = {
      id: editingId || uid(),
      name, dept: $('fDept').value.trim(),
      date, type, hours,
      start: $('fStart').value, end: $('fEnd').value,
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
    saveEntries(entries);
    // 若录入日期不在当前统计月份，自动跳到该月份
    const mk = monthKey(date);
    if (mk !== $('setMonth').value) { $('setMonth').value = mk; collectSettings(); }
    clearForm(); renderAll();
  }
  function clearForm() {
    $('fName').value = ''; $('fDate').value = ''; $('fHours').value = '';
    $('fStart').value = ''; $('fEnd').value = ''; $('fReason').value = '';
    $('fType').value = 'weekday'; $('fDept').value = '统战';
  }
  function startEdit(id) {
    const e = entries.find(x => x.id === id); if (!e) return;
    editingId = id;
    $('fName').value = e.name; $('fDept').value = e.dept || '统战'; $('fDate').value = e.date;
    $('fType').value = e.type; $('fHours').value = e.hours; $('fStart').value = e.start || '';
    $('fEnd').value = e.end || ''; $('fReason').value = e.reason || '';
    $('btnAdd').textContent = '保存修改';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  function delEntry(id) {
    if (!confirm('确定删除这条记录？')) return;
    entries = entries.filter(e => e.id !== id);
    saveEntries(entries); renderAll(); toast('已删除');
  }

  // ---------- 导出 ----------
  function buildCtx() {
    const s = collectSettings();
    const [y, m] = s.month.split('-').map(Number);
    return {
      year: y, month: m, unit: s.unit,
      pubStart: s.pubStart || '', pubEnd: s.pubEnd || '',
      makeDate: s.makeDate, timeStr: s.makeDate,
    };
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
    const data = { version: 1, exportedAt: new Date().toISOString(), settings, entries };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `统侨加班数据备份_${new Date().toISOString().slice(0, 10)}.json`;
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
        if (data.settings) { settings = data.settings; saveSettings(settings); applySettingsToUI(); }
        // 合并（按 id 去重，存在则覆盖）
        const map = {};
        entries.forEach(e => map[e.id] = e);
        data.entries.forEach(e => map[e.id || uid()] = e);
        entries = Object.values(map);
        saveEntries(entries); renderAll();
        toast(`已导入 ${data.entries.length} 条记录`);
      } catch (e) { alert('导入失败：' + e.message); }
    };
    reader.readAsText(file);
  }

  function renderAll() { applySettingsToUI(); renderEntries(); renderSummary(); }

  // ---------- 事件 ----------
  $('btnAdd').onclick = addOrUpdate;
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
    saveEntries(entries); renderAll(); toast('已清空本月记录');
  };
  $('btnClearAll').onclick = () => {
    if (!confirm('确定清空全部数据？此操作不可恢复，建议先导出备份。')) return;
    entries = []; saveEntries(entries); renderAll(); toast('已清空全部数据');
  };

  // 首次进入：若设置未填制表日期则补默认
  if (!settings.makeDate) { settings.makeDate = todayCN(); saveSettings(settings); }
  renderAll();
})();
