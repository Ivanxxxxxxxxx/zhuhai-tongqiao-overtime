/* 统侨科加班统计 — 导出组件
 * 方案：直接基于内置的「统侨模板」底座（vendor/template.js 提供的原 .xls），
 *        把当月数据逐格填入对应位置，再另存为 .xlsx。
 *        —— 不自行重建表格、不自行添加合并，彻底避免表头跨列/合并错位。
 * 依赖：window.XLSX（vendor/xlsx.full.min.js 的 SheetJS）
 * 浏览器端：window.XLSXGen；Node 端：module.exports
 */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.XLSXGen = api;
})(this, function () {
  // ===== 补贴计算规则（与模板真实数字吻合）=====
  // 不满1小时不计；满1小时后按小时向上取整（1.1h按2h、2.5h按3h…）
  // 工作日：15元/小时，封顶80元/天；双休日：20元/小时，封顶160元/天；法定节假日：30元/小时，封顶240元/天
  const RATES = { weekday: 15, weekend: 20, holiday: 30 };
  const CAPS = { weekday: 80, weekend: 160, holiday: 240 };
  function effHours(hours) { return hours < 1 ? 0 : Math.ceil(hours); }
  function entrySubsidy(type, hours) {
    const h = effHours(hours);
    if (h <= 0) return 0;
    return Math.min(h * RATES[type], CAPS[type]);
  }

  function XLSX() { return (typeof window !== 'undefined' && window.XLSX) || (typeof global !== 'undefined' && global.window && global.window.XLSX); }

  // 在某个 sheet 中找「同时含所有关键词」的行（0-based 行号）
  function findRow(ws, keywords) {
    const XU = XLSX().utils;
    const range = XU.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      const vals = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XU.encode_cell({ r, c })];
        if (cell && cell.v != null) vals.push(String(cell.v));
      }
      const joined = vals.join(' ');
      if (keywords.every(k => joined.indexOf(k) >= 0)) return r;
    }
    return -1;
  }
  // 取某列在数据区的首个样式（s），用于写入新行时沿用边框/字体
  function colStyle(ws, col, r0, r1) {
    const XU = XLSX().utils;
    for (let r = r0; r <= r1; r++) {
      const cell = ws[XU.encode_cell({ r, c: col })];
      if (cell && cell.s != null) return cell.s;
    }
    return undefined;
  }
  function setCell(ws, r, c, value, style, type) {
    const XU = XLSX().utils;
    const addr = XU.encode_cell({ r, c });
    if (value === '' || value == null) {
      ws[addr] = { t: 's', v: '' };
    } else {
      ws[addr] = { t: type || (typeof value === 'number' ? 'n' : 's'), v: value };
    }
    if (style != null) ws[addr].s = style;
  }
  // 把 beforeRow（0-based，含）及之后的所有行整体下移 count 行，并处理合并区
  function insertRows(ws, beforeRow, count) {
    if (count <= 0) return;
    const XU = XLSX().utils;
    const moves = [];
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const cc = XU.decode_cell(addr);
      if (cc.r >= beforeRow) moves.push(addr);
    }
    moves.sort((a, b) => XU.decode_cell(b).r - XU.decode_cell(a).r);
    for (const addr of moves) {
      const cc = XU.decode_cell(addr);
      const cell = ws[addr];
      delete ws[addr];
      ws[XU.encode_cell({ r: cc.r + count, c: cc.c })] = cell;
    }
    const range = XU.decode_range(ws['!ref']);
    range.e.r += count;
    ws['!ref'] = XU.encode_range(range);
    if (ws['!merges']) {
      ws['!merges'] = ws['!merges'].map(m => ({
        s: { r: m.s.r >= beforeRow ? m.s.r + count : m.s.r, c: m.s.c },
        e: { r: m.e.r >= beforeRow ? m.e.r + count : m.e.r, c: m.e.c },
      }));
    }
  }

  // 把「向上取整后的整数小时」写入模板三类时长列
  function effInt(h) { const v = effHours(h); return v; }

  /* 核心：基于模板底座填数据
   * ctx: {year, month, unit, pubStart, pubEnd, makeDate, timeStr}
   * entries: 当月记录数组 [{date,type,name,dept,hours,start,end,reason}]
   * tplBuf: 模板文件 ArrayBuffer/Uint8Array
   * 返回 { buf: Uint8Array(xlsx), subT }
   */
  function fillTemplate(ctx, entries, tplBuf) {
    const X = XLSX(); const XU = X.utils;
    const wb = X.read(tplBuf, { type: 'array', cellDates: true, cellStyles: true });
    const ymLabel = `${ctx.year}年${ctx.month}月`;

    const gName = wb.SheetNames.find(n => String(n).indexOf('公示') >= 0) || wb.SheetNames[0];
    const fName = wb.SheetNames.find(n => String(n).indexOf('发放') >= 0) || wb.SheetNames[1] || wb.SheetNames[0];
    const gws = wb.Sheets[gName];
    const fws = wb.Sheets[fName];

    // ===================== 公示表 =====================
    const gHead = findRow(gws, ['姓名', '加班日期']);
    const gSub = findRow(gws, ['工作日']) >= 0 ? findRow(gws, ['工作日']) : gHead + 1;
    const gDataStart = gSub + 1;
    const gTotalRow = findRow(gws, ['小计']);
    const gTplRows = gTotalRow - gDataStart;
    const gColStyle = [];
    for (let c = 0; c < 10; c++) gColStyle[c] = colStyle(gws, c, gDataStart, gTotalRow - 1);

    const gNeed = entries.length;
    if (gNeed > gTplRows) insertRows(gws, gTotalRow, gNeed - gTplRows);
    const gTotalNew = findRow(gws, ['小计']); // 下移后重新定位

    let wkSum = 0, weSum = 0, hoSum = 0;
    const gFinal = Math.max(gNeed, gTplRows);
    for (let i = 0; i < gFinal; i++) {
      const r = gDataStart + i;
      if (i < gNeed) {
        const e = entries[i];
        const col = e.type === 'weekday' ? 4 : e.type === 'weekend' ? 5 : 6;
        const h = effInt(e.hours);
        setCell(gws, r, 0, i + 1, gColStyle[0], 'n');
        setCell(gws, r, 1, e.dept || '', gColStyle[1]);
        setCell(gws, r, 2, e.name, gColStyle[2]);
        setCell(gws, r, 3, e.date, gColStyle[3]);
        setCell(gws, r, 4, e.type === 'weekday' ? h : '', gColStyle[4], 'n');
        setCell(gws, r, 5, e.type === 'weekend' ? h : '', gColStyle[5], 'n');
        setCell(gws, r, 6, e.type === 'holiday' ? h : '', gColStyle[6], 'n');
        setCell(gws, r, 7, e.start || '', gColStyle[7]);
        setCell(gws, r, 8, e.end || '', gColStyle[8]);
        setCell(gws, r, 9, e.reason || '', gColStyle[9]);
        if (e.type === 'weekday') wkSum += h; else if (e.type === 'weekend') weSum += h; else hoSum += h;
      } else {
        for (let c = 0; c < 10; c++) setCell(gws, r, c, '', gColStyle[c]);
      }
    }
    setCell(gws, gTotalNew, 1, '小计', gColStyle[1]);
    setCell(gws, gTotalNew, 4, wkSum, gColStyle[4], 'n');
    setCell(gws, gTotalNew, 5, weSum, gColStyle[5], 'n');
    setCell(gws, gTotalNew, 6, hoSum, gColStyle[6], 'n');
    // 标题与日期
    setCell(gws, 0, 0, `珠海高新区合同制职员加班情况公示表（${ymLabel}）`, colStyle(gws, 0, 0, 0));
    setCell(gws, 2, 0, `填报单位：${ctx.unit}`, colStyle(gws, 0, 2, 2));
    setCell(gws, 2, 5, `公示时间：${ctx.pubStart}至${ctx.pubEnd}`, colStyle(gws, 5, 2, 2));
    setCell(gws, 2, 9, `制表日期：${ctx.makeDate}`, colStyle(gws, 9, 2, 2));

    // ===================== 发放表（按人汇总）=====================
    const fHead = findRow(fws, ['序号', '姓名', '累计加班时间']) >= 0 ? findRow(fws, ['序号', '姓名', '累计加班时间']) : findRow(fws, ['姓名']);
    const fSub = findRow(fws, ['工作日加班']) >= 0 ? findRow(fws, ['工作日加班']) : fHead + 1;
    const fDataStart = fSub + 1;
    const fTotalRow = findRow(fws, ['合计']);
    const fTplRows = fTotalRow - fDataStart;
    const fColStyle = [];
    for (let c = 0; c < 7; c++) fColStyle[c] = colStyle(fws, c, fDataStart, fTotalRow - 1);

    const people = {};
    for (const e of entries) {
      const p = people[e.name] = people[e.name] || { name: e.name, wk: 0, we: 0, ho: 0, sub: 0 };
      const h = effInt(e.hours);
      if (e.type === 'weekday') p.wk += h; else if (e.type === 'weekend') p.we += h; else p.ho += h;
      p.sub += entrySubsidy(e.type, e.hours);
    }
    const names = Object.keys(people);
    const fNeed = names.length;
    if (fNeed > fTplRows) insertRows(fws, fTotalRow, fNeed - fTplRows);
    const fTotalNew = findRow(fws, ['合计']);

    let wkT = 0, weT = 0, hoT = 0, subT = 0;
    const fFinal = Math.max(fNeed, fTplRows);
    for (let i = 0; i < fFinal; i++) {
      const r = fDataStart + i;
      if (i < fNeed) {
        const p = people[names[i]];
        setCell(fws, r, 0, i + 1, fColStyle[0], 'n');
        setCell(fws, r, 1, p.name, fColStyle[1]);
        setCell(fws, r, 2, p.wk, fColStyle[2], 'n');
        setCell(fws, r, 3, p.we, fColStyle[3], 'n');
        setCell(fws, r, 4, p.ho, fColStyle[4], 'n');
        setCell(fws, r, 5, Math.round(p.sub * 100) / 100, fColStyle[5], 'n');
        setCell(fws, r, 6, '', fColStyle[6]);
        wkT += p.wk; weT += p.we; hoT += p.ho; subT += p.sub;
      } else {
        for (let c = 0; c < 7; c++) setCell(fws, r, c, '', fColStyle[c]);
      }
    }
    setCell(fws, fTotalNew, 0, '合计', fColStyle[0]);
    setCell(fws, fTotalNew, 2, wkT, fColStyle[2], 'n');
    setCell(fws, fTotalNew, 3, weT, fColStyle[3], 'n');
    setCell(fws, fTotalNew, 4, hoT, fColStyle[4], 'n');
    setCell(fws, fTotalNew, 5, Math.round(subT * 100) / 100, fColStyle[5], 'n');
    // 标题与日期
    setCell(fws, 0, 0, `${ymLabel}珠海高新区合同制职员加班补贴发放表`, colStyle(fws, 0, 0, 0));
    setCell(fws, 2, 0, `填报单位：${ctx.unit}`, colStyle(fws, 0, 2, 2));
    setCell(fws, 2, 5, `时间：${ctx.timeStr}`, colStyle(fws, 5, 2, 2));

    const buf = X.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
    return { buf: buf, subT: Math.round(subT * 100) / 100, wkSum, weSum, hoSum, gFinal, fFinal };
  }

  // 浏览器端下载
  function downloadXlsx(buf, filename) {
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  return { entrySubsidy, effHours, fillTemplate, downloadXlsx };
});
