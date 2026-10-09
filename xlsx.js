/* 统侨科加班统计 — 导出组件
 * 方案：基于 ExcelJS 读取内置「统侨模板」底座（vendor/template.js 的原 .xlsx），
 *        把当月数据逐格填入对应位置，再写回 .xlsx。
 *        ExcelJS 读/写均完整保留母版边框、字体、列宽、合并、日期/时间格式，
 *        彻底避免 SheetJS 社区版写不出边框导致的“排版错乱”。
 * 依赖：window.ExcelJS（vendor/exceljs.min.js）
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
  function effHours(h) { return h < 1 ? 0 : Math.ceil(h); }
  function entrySubsidy(type, hours) {
    const h = effHours(hours);
    if (h <= 0) return 0;
    return Math.min(h * RATES[type], CAPS[type]);
  }
  function effInt(h) { return effHours(h); }

  // "YYYY-MM-DD" -> Date（按本地0点）；"HH:MM" -> Date（1899-12-30 基准，Excel 时间小数）
  function toDateObj(s) {
    if (!s) return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s));
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  function toTimeObj(s) {
    if (!s) return null;
    const m = /^(\d{1,2}):(\d{2})/.exec(String(s));
    if (!m) return null;
    return new Date(1899, 11, 30, Number(m[1]), Number(m[2]), 0);
  }

  // ===== 行列工具 =====
  function colL(n) { let s = ''; n = Number(n); while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colLetterToNum(l) { let n = 0; for (const ch of String(l)) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function refRow(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? +m[2] : 0; }
  function refCol(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? colLetterToNum(m[1]) : 0; }

  // 找含全部关键词的行（1-based 行号）
  function findRow(ws, kws) {
    let f = -1;
    ws.eachRow((row, rn) => {
      if (f >= 0) return;
      const vals = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.value != null && cell.value !== '') vals.push(String(cell.value));
      });
      const j = vals.join(' ');
      if (kws.every(k => j.indexOf(k) >= 0)) f = rn;
    });
    return f;
  }

  // 合并区域随插入行整体下移
  function shiftMerges(ws, beforeRow1, count) {
    const merges = ws.model.merges || [];
    ws.model.merges = merges.map(m => {
      const parts = m.split(':');
      const a = parts[0], b = parts[1] || a;
      const ra = refRow(a), ca = refCol(a), rb = refRow(b), cb = refCol(b);
      const na = ra >= beforeRow1 ? ra + count : ra;
      const nb = rb >= beforeRow1 ? rb + count : rb;
      if (na !== ra || nb !== rb) return colL(ca) + na + ':' + colL(cb) + nb;
      return m;
    });
  }

  // 在 beforeRow1（1-based）之前插入 count 个空行：
  //   - 用 ExcelJS 原生 insertRow 把其后内容整体下移（自动带样式，但【不】推移合并）
  //   - 手动 shiftMerges 把合并区域随插入下移
  //   - 给新插入的空行复制“数据区末行”样式（边框）以保版式
  async function excelInsertRows(ws, beforeRow1, count) {
    if (count <= 0) return;
    for (let k = 0; k < count; k++) ws.insertRow(beforeRow1, []);
    shiftMerges(ws, beforeRow1, count);
    const styleSrc = ws.getRow(Math.min(beforeRow1 - 1, ws.rowCount));
    if (!styleSrc) return;
    for (let r = beforeRow1; r < beforeRow1 + count; r++) {
      const dst = ws.getRow(r);
      styleSrc.eachCell({ includeEmpty: true }, (cell, col) => {
        dst.getCell(col).style = Object.assign({}, cell.style);
      });
    }
  }

  /* 核心：基于模板底座填数据（async，依赖 ExcelJS）
   * ctx: {year, month, unit, pubStart, pubEnd, makeDate, timeStr}
   * entries: 当月记录数组 [{date,type,name,dept,hours,start,end,reason}]
   * tplBuf: 模板文件 ArrayBuffer/Uint8Array
   */
  async function fillTemplate(ctx, entries, tplBuf) {
    const ExcelJS = (typeof window !== 'undefined' && window.ExcelJS) ||
      (typeof global !== 'undefined' && global.window && global.window.ExcelJS) ||
      (typeof require !== 'undefined' ? require('exceljs') : null);
    if (!ExcelJS) throw new Error('ExcelJS 未加载');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(tplBuf);
    const ymLabel = `${ctx.year}年${ctx.month}月`;

    // ===================== 公示表 =====================
    const gwsName = wb.worksheets.find(w => String(w.name).indexOf('公示') >= 0).name;
    const gws = wb.getWorksheet(gwsName);
    const gHead = findRow(gws, ['姓名', '加班日期']);
    const gSub = findRow(gws, ['工作日']) || gHead + 1;
    const gDataStart = gSub + 1;
    const gTotal = findRow(gws, ['小计']);
    const gTplRows = gTotal - gDataStart;
    const gNeed = entries.length;
    const gCount = gNeed > gTplRows ? gNeed - gTplRows : 0;
    if (gCount > 0) await excelInsertRows(gws, gTotal, gCount);
    const gTotalNew = gTotal + gCount;
    let wkSum = 0, weSum = 0, hoSum = 0;
    const gFinal = Math.max(gNeed, gTplRows);
    for (let i = 0; i < gFinal; i++) {
      const row = gws.getRow(gDataStart + i);
      if (i < gNeed) {
        const e = entries[i];
        const h = effInt(e.hours);
        row.getCell(1).value = i + 1;
        row.getCell(2).value = e.dept || '';
        row.getCell(3).value = e.name;
        const dc = row.getCell(4); const dObj = toDateObj(e.date); if (dObj) { dc.value = dObj; dc.numFmt = 'yyyy-mm-dd'; }
        row.getCell(5).value = e.type === 'weekday' ? h : '';
        row.getCell(6).value = e.type === 'weekend' ? h : '';
        row.getCell(7).value = e.type === 'holiday' ? h : '';
        const sc = row.getCell(8); const sObj = toTimeObj(e.start); if (sObj) { sc.value = sObj; sc.numFmt = 'h:mm'; }
        const ec = row.getCell(9); const eObj = toTimeObj(e.end); if (eObj) { ec.value = eObj; ec.numFmt = 'h:mm'; }
        row.getCell(10).value = e.reason || '';
        if (e.type === 'weekday') wkSum += h; else if (e.type === 'weekend') weSum += h; else hoSum += h;
      } else {
        for (let c = 1; c <= 10; c++) row.getCell(c).value = '';
      }
    }
    gws.getRow(gTotalNew).getCell(2).value = '小计';
    gws.getRow(gTotalNew).getCell(5).value = wkSum;
    gws.getRow(gTotalNew).getCell(6).value = weSum;
    gws.getRow(gTotalNew).getCell(7).value = hoSum;
    gws.getRow(1).getCell(1).value = `珠海高新区合同制职员加班情况公示表（${ymLabel}）`;
    gws.getRow(3).getCell(1).value = `填报单位：${ctx.unit}`;
    gws.getRow(3).getCell(6).value = `公示时间：${ctx.pubStart}至${ctx.pubEnd}`;
    gws.getRow(3).getCell(10).value = `制表日期：${ctx.makeDate}`;

    // ===================== 发放表（按人汇总）=====================
    const fwsName = wb.worksheets.find(w => String(w.name).indexOf('发放') >= 0).name;
    const fws = wb.getWorksheet(fwsName);
    const fHead = findRow(fws, ['序号', '姓名', '累计加班时间']) >= 0 ? findRow(fws, ['序号', '姓名', '累计加班时间']) : findRow(fws, ['姓名']);
    const fSub = findRow(fws, ['工作日加班']) || fHead + 1;
    const fDataStart = fSub + 1;
    const fTotal = findRow(fws, ['合计']);
    const fTplRows = fTotal - fDataStart;
    const people = {};
    for (const e of entries) {
      const p = people[e.name] = people[e.name] || { name: e.name, wk: 0, we: 0, ho: 0, sub: 0 };
      const h = effInt(e.hours);
      if (e.type === 'weekday') p.wk += h; else if (e.type === 'weekend') p.we += h; else p.ho += h;
      p.sub += entrySubsidy(e.type, e.hours);
    }
    const names = Object.keys(people);
    const fNeed = names.length;
    const fCount = fNeed > fTplRows ? fNeed - fTplRows : 0;
    if (fCount > 0) await excelInsertRows(fws, fTotal, fCount);
    const fTotalNew = fTotal + fCount;
    let wkT = 0, weT = 0, hoT = 0, subT = 0;
    const fFinal = Math.max(fNeed, fTplRows);
    for (let i = 0; i < fFinal; i++) {
      const row = fws.getRow(fDataStart + i);
      if (i < fNeed) {
        const p = people[names[i]];
        row.getCell(1).value = i + 1;
        row.getCell(2).value = p.name;
        row.getCell(3).value = p.wk;
        row.getCell(4).value = p.we;
        row.getCell(5).value = p.ho;
        row.getCell(6).value = Math.round(p.sub * 100) / 100;
        row.getCell(7).value = '';
        wkT += p.wk; weT += p.we; hoT += p.ho; subT += p.sub;
      } else {
        for (let c = 1; c <= 7; c++) row.getCell(c).value = '';
      }
    }
    // 母版“合计”行为合并 A:E，insertRow 会丢失该合并。
    // 顺序：先解除合并 → 写入真实数值（A=合计、C/D/E=分项合计、F=总补贴）→ 再合并（与母版一致，隐藏的 C/D/E 存真实合计）
    try { fws.unMergeCells('A' + fTotalNew + ':E' + fTotalNew); } catch (e) { /* 无合并则忽略 */ }
    fws.getRow(fTotalNew).getCell(1).value = '合计';
    fws.getRow(fTotalNew).getCell(2).value = '';
    fws.getRow(fTotalNew).getCell(3).value = wkT;
    fws.getRow(fTotalNew).getCell(4).value = weT;
    fws.getRow(fTotalNew).getCell(5).value = hoT;
    fws.getRow(fTotalNew).getCell(6).value = Math.round(subT * 100) / 100;
    try { fws.mergeCells('A' + fTotalNew + ':E' + fTotalNew); } catch (e) { /* 忽略 */ }
    fws.getRow(1).getCell(1).value = `${ymLabel}珠海高新区合同制职员加班补贴发放表`;
    fws.getRow(3).getCell(1).value = `填报单位：${ctx.unit}`;
    fws.getRow(3).getCell(6).value = `时间：${ctx.timeStr}`;

    const buf = await wb.xlsx.writeBuffer();
    return { buf, subT: Math.round(subT * 100) / 100, wkSum, weSum, hoSum, gFinal, fFinal };
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
