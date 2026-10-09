/* 纯 JS 的 XLSX 生成器（无外部依赖，stored zip + CRC32）
 * 浏览器端：window.XLSXGen；Node 端：module.exports
 * 已用真实模板样本校验：徐梦华 工作日7h 双休日5h 补贴=205 元，与模板发放表数字完全吻合。
 */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.XLSXGen = api;
})(this, function () {
  const enc = new TextEncoder();

  function crc32(buf) {
    let table = crc32.table;
    if (!table) {
      table = crc32.table = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
      }
    }
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xFF];
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function zip(files) {
    const parts = [];
    const central = [];
    let offset = 0;
    for (const f of files) {
      const nameBytes = enc.encode(f.name);
      const data = f.data;
      const crc = crc32(data);
      const size = data.length;
      const lh = new Uint8Array(30 + nameBytes.length);
      const dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0, true);
      dv.setUint16(8, 0, true);
      dv.setUint16(10, 0, true);
      dv.setUint16(12, 0, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, size, true);
      dv.setUint32(22, size, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      parts.push(lh, data);
      const ch = new Uint8Array(46 + nameBytes.length);
      const cdv = new DataView(ch.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0, true);
      cdv.setUint16(10, 0, true);
      cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, size, true);
      cdv.setUint32(24, size, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint16(38, 0, true);
      cdv.setUint32(42, offset, true); // 本地文件头偏移（中央目录第42字节）
      ch.set(nameBytes, 46);
      central.push(ch);
      offset += lh.length + data.length;
    }
    const centralSize = central.reduce((s, c) => s + c.length, 0);
    const centralOffset = offset;
    const end = new Uint8Array(22);
    const edv = new DataView(end.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(4, 0, true);
    edv.setUint16(6, 0, true);
    edv.setUint16(8, files.length, true);
    edv.setUint16(10, files.length, true);
    edv.setUint32(12, centralSize, true);
    edv.setUint32(16, centralOffset, true);
    edv.setUint16(20, 0, true);
    const all = [...parts, ...central, end];
    const total = all.reduce((s, a) => s + a.length, 0);
    const out = new Uint8Array(total);
    let p = 0;
    for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  function xmlEsc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function colLetter(c) {
    let s = ''; c++;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function ref(c, r) { return colLetter(c) + (r + 1); }

  // ===== 补贴计算规则（与模板真实数字吻合）=====
  // 不满1小时不算；满1小时后按小时向上取整（1.1h按2h、2.5h按3h…）
  // 工作日：15元/小时，封顶80元/天；双休日：20元/小时，封顶160元/天；法定节假日：30元/小时，封顶240元/天
  const RATES = { weekday: 15, weekend: 20, holiday: 30 };
  const CAPS = { weekday: 80, weekend: 160, holiday: 240 };
  function effHours(hours) { return hours < 1 ? 0 : Math.ceil(hours); }
  function entrySubsidy(type, hours) {
    const h = effHours(hours);
    if (h <= 0) return 0;
    return Math.min(h * RATES[type], CAPS[type]);
  }

  function buildSheetXML(cols, cells) {
    const byRow = {};
    for (const cell of cells) (byRow[cell.r] = byRow[cell.r] || []).push(cell);
    let colXml = '';
    if (cols && cols.length) {
      colXml = '<cols>' + cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    }
    let rowsXml = '';
    const rowIdxs = Object.keys(byRow).map(Number).sort((a, b) => a - b);
    for (const r of rowIdxs) {
      const rowCells = byRow[r].slice().sort((a, b) => a.c - b.c);
      let cXml = '';
      for (const cell of rowCells) {
        const R = ref(cell.c, r);
        if (cell.t === 'n') {
          cXml += `<c r="${R}" s="${cell.s}"><v>${cell.v}</v></c>`;
        } else {
          cXml += `<c r="${R}" s="${cell.s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(cell.v)}</t></is></c>`;
        }
      }
      rowsXml += `<row r="${r + 1}">${cXml}</row>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${colXml}<sheetData>${rowsXml}</sheetData></worksheet>`;
  }

  function buildWorkbook(sheets) {
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('\n')}
</sheets>
</workbook>`;
    const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n')}
<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
<font><sz val="11"/><name val="宋体"/></font>
<font><b/><sz val="11"/><name val="宋体"/></font>
<font><b/><sz val="16"/><name val="宋体"/></font>
</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
</cellXfs>
</styleSheet>`;
    const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>统侨科加班统计</dc:creator><cp:lastModifiedBy>统侨科加班统计</cp:lastModifiedBy></cp:coreProperties>`;
    const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>统侨科加班统计</Application></Properties>`;

    const files = [
      { name: '[Content_Types].xml', data: enc.encode(contentTypes) },
      { name: '_rels/.rels', data: enc.encode(rootRels) },
      { name: 'docProps/core.xml', data: enc.encode(coreXml) },
      { name: 'docProps/app.xml', data: enc.encode(appXml) },
      { name: 'xl/workbook.xml', data: enc.encode(workbookXml) },
      { name: 'xl/_rels/workbook.xml.rels', data: enc.encode(wbRels) },
      { name: 'xl/styles.xml', data: enc.encode(stylesXml) },
    ];
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc.encode(buildSheetXML(s.cols, s.cells)) }));
    return zip(files);
  }

  // 按模板构建「公示表 + 发放表」。返回 {buf, people, wkSum, weSum, hoSum, subT}
  function buildMonthWorkbook(ctx, entries) {
    const { year, month, unit, pubStart, pubEnd, makeDate, timeStr } = ctx;
    const ymLabel = `${year}年${month}月`;
    const S = { title: 1, header: 2, data: 3, dataLeft: 4, bold: 5, note: 6 };

    // ===== 公示表 =====
    const gCells = [];
    gCells.push({ c: 0, r: 0, t: 's', v: `珠海高新区合同制职员加班情况公示表（${ymLabel}）`, s: S.title });
    gCells.push({ c: 0, r: 2, t: 's', v: `填报单位：${unit}`, s: S.note });
    gCells.push({ c: 5, r: 2, t: 's', v: `公示时间：${pubStart}至${pubEnd}`, s: S.note });
    gCells.push({ c: 9, r: 2, t: 's', v: `制表日期：${makeDate}`, s: S.note });
    const gHead = ['序号', '部门', '姓名', '加班日期', '加班时间（小时）', '', '', '加时时段', '', '加班事由'];
    gHead.forEach((v, c) => { if (v) gCells.push({ c, r: 3, t: 's', v, s: S.header }); });
    const gSub = { 4: '工作日', 5: '双休日', 6: '法定\n节假日', 7: '起始时间', 8: '终止时间' };
    Object.entries(gSub).forEach(([c, v]) => gCells.push({ c: Number(c), r: 4, t: 's', v, s: S.header }));

    let wkSum = 0, weSum = 0, hoSum = 0;
    entries.forEach((e, i) => {
      const r = 5 + i;
      gCells.push({ c: 0, r, t: 'n', v: i + 1, s: S.data });
      gCells.push({ c: 1, r, t: 's', v: e.dept || '', s: S.data });
      gCells.push({ c: 2, r, t: 's', v: e.name, s: S.data });
      gCells.push({ c: 3, r, t: 's', v: e.date, s: S.data });
      const col = e.type === 'weekday' ? 4 : e.type === 'weekend' ? 5 : 6;
      gCells.push({ c: col, r, t: 'n', v: e.hours, s: S.data });
      if (e.type === 'weekday') wkSum += e.hours; else if (e.type === 'weekend') weSum += e.hours; else hoSum += e.hours;
      gCells.push({ c: 7, r, t: 's', v: e.start || '', s: S.data });
      gCells.push({ c: 8, r, t: 's', v: e.end || '', s: S.data });
      gCells.push({ c: 9, r, t: 's', v: e.reason || '', s: S.dataLeft });
    });
    const gEnd = 5 + entries.length;
    gCells.push({ c: 1, r: gEnd, t: 's', v: '小计', s: S.bold });
    gCells.push({ c: 4, r: gEnd, t: 'n', v: wkSum, s: S.data });
    gCells.push({ c: 5, r: gEnd, t: 'n', v: weSum, s: S.data });
    gCells.push({ c: 6, r: gEnd, t: 'n', v: hoSum, s: S.data });
    gCells.push({ c: 0, r: gEnd + 1, t: 's', v: '制表人：                   科室负责人意见：                                分管领导意见：', s: S.note });
    const gSheet = { name: '公示表', cols: [6, 10, 10, 14, 10, 10, 12, 10, 10, 40], cells: gCells };

    // ===== 发放表（按人汇总）=====
    const people = {};
    for (const e of entries) {
      const p = people[e.name] = people[e.name] || { name: e.name, wk: 0, we: 0, ho: 0, sub: 0 };
      if (e.type === 'weekday') p.wk += e.hours; else if (e.type === 'weekend') p.we += e.hours; else p.ho += e.hours;
      p.sub += entrySubsidy(e.type, e.hours);
    }
    const names = Object.keys(people);
    const fCells = [];
    fCells.push({ c: 0, r: 0, t: 's', v: `${ymLabel}珠海高新区合同制职员加班补贴发放表`, s: S.title });
    fCells.push({ c: 0, r: 2, t: 's', v: `填报单位：${unit}`, s: S.note });
    fCells.push({ c: 5, r: 2, t: 's', v: `时间：${timeStr}`, s: S.note });
    const fHead = ['序号', '姓名', '累计加班时间（小时）', '', '', '加班补贴金额\n（元）', '备注'];
    fHead.forEach((v, c) => { if (v) fCells.push({ c, r: 3, t: 's', v, s: S.header }); });
    const fSub = { 2: '工作日加班', 3: '双休日加班', 4: '法定节假日\n加班' };
    Object.entries(fSub).forEach(([c, v]) => fCells.push({ c: Number(c), r: 4, t: 's', v, s: S.header }));
    let wkT = 0, weT = 0, hoT = 0, subT = 0;
    names.forEach((nm, i) => {
      const p = people[nm]; const r = 5 + i;
      fCells.push({ c: 0, r, t: 'n', v: i + 1, s: S.data });
      fCells.push({ c: 1, r, t: 's', v: p.name, s: S.data });
      fCells.push({ c: 2, r, t: 'n', v: p.wk, s: S.data });
      fCells.push({ c: 3, r, t: 'n', v: p.we, s: S.data });
      fCells.push({ c: 4, r, t: 'n', v: p.ho, s: S.data });
      fCells.push({ c: 5, r, t: 'n', v: Math.round(p.sub * 100) / 100, s: S.data });
      fCells.push({ c: 6, r, t: 's', v: '', s: S.dataLeft });
      wkT += p.wk; weT += p.we; hoT += p.ho; subT += p.sub;
    });
    const fEnd = 5 + names.length;
    fCells.push({ c: 0, r: fEnd, t: 's', v: '合计', s: S.bold });
    fCells.push({ c: 2, r: fEnd, t: 'n', v: wkT, s: S.data });
    fCells.push({ c: 3, r: fEnd, t: 'n', v: weT, s: S.data });
    fCells.push({ c: 4, r: fEnd, t: 'n', v: hoT, s: S.data });
    fCells.push({ c: 5, r: fEnd, t: 'n', v: Math.round(subT * 100) / 100, s: S.data });
    const rules = '备注：加班补贴按小时向上取整核算（不满1小时不计，满1小时后1.1小时按2小时算，以此类推）；工作日加班以15元/小时标准计算，当天加班补贴最高为80元；双休日加班以20元/小时标准计算，当天加班补贴最高为160元；法定节假日加班以30元/小时标准计算，当天加班补贴最高为240元。';
    fCells.push({ c: 0, r: fEnd + 1, t: 's', v: rules, s: S.note });
    fCells.push({ c: 0, r: fEnd + 2, t: 's', v: '制表人：                                   单位意见：                 ', s: S.note });
    fCells.push({ c: 0, r: fEnd + 3, t: 's', v: '党群工作部意见：                               发改财政局意见：', s: S.note });
    const fSheet = { name: '发放表', cols: [6, 12, 14, 14, 16, 16, 20], cells: fCells };

    return { buf: buildWorkbook([gSheet, fSheet]), people, wkSum, weSum, hoSum, subT: Math.round(subT * 100) / 100 };
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

  return { crc32, zip, entrySubsidy, buildWorkbook, buildMonthWorkbook, downloadXlsx };
});
