(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TPExport = factory();
})(this, function () {
  'use strict';

  const MAX_CELL = 32767;
  const BAD_XML = /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g;
  const SCALE_DEC = { '': 0, K: 3, M: 6, B: 9 };

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function isoFromMs(t, withTime) {
    const d = new Date(t);
    const s = d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    return withTime ? s + ' ' + pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes()) + ':' + pad(d.getUTCSeconds()) : s;
  }

  function cleanText(s) {
    const t = s == null ? '' : String(s).replace(BAD_XML, '');
    return t.length > MAX_CELL ? t.slice(0, MAX_CELL) : t;
  }

  // Raw number rounded to what the table displays, so CSV values match the screen without symbols.
  function roundAsShown(v, spec) {
    if (!spec) return Number(v.toPrecision(12));
    const dec = spec.dec + (spec.pct ? 2 : 0) + SCALE_DEC[spec.scale || ''];
    return Number(v.toFixed(Math.min(dec, 15)));
  }

  // snap.value(col, rowIndex) -> { kind: 'num'|'date'|'text'|'null', v, time }
  function csvCell(cell, col) {
    switch (cell.kind) {
      case 'num': return String(roundAsShown(cell.v, col.spec));
      case 'date': return isoFromMs(cell.v, cell.time);
      case 'text': return noFormula(cleanText(cell.v));
      case 'link': return noFormula(cleanText(cell.csv != null ? cell.csv : cell.v));
      default: return '';
    }
  }

  // Text starting with = + - @ would run as a formula when the CSV is opened in Excel; a leading apostrophe stops that.
  function noFormula(s) { return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; }

  function csvEscape(s) { return /[",\r\n]/.test(s) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }

  // CSV is raw data only: a header line and every data row, never totals or subtotals.
  function buildCsv(snap) {
    const lines = [];
    const cols = snap.columns;
    const row = cells => lines.push(cells.map(csvEscape).join(','));
    row(cols.map(c => noFormula(cleanText(c.label))));
    for (const r of snap.rows) {
      if (r.type !== 'data') continue;
      row(cols.map(c => csvCell(snap.value(c, r.i), c)));
    }
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  function sheetName(s) {
    const t = String(s || 'Data').replace(/[\\\/\?\*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '');
    // "History" is reserved by Excel.
    return (t && t.toLowerCase() !== 'history' ? t : t ? t + ' data' : 'Data').slice(0, 31);
  }

  async function buildXlsx(snap, ExcelJS) {
    const cols = snap.columns;
    const wb = new ExcelJS.Workbook();
    const top = snap.includeTotals && snap.totalsLocation === 'top';
    const bottom = snap.includeTotals && snap.totalsLocation === 'bottom';
    const groups = (snap.groupHeaders || []).filter(g => g.label);
    const headRow = groups.length ? 2 : 1;
    const ySplit = headRow + (top ? 1 : 0);
    const ws = wb.addWorksheet(sheetName(snap.sheetName), {
      views: [{ state: 'frozen', xSplit: Math.min(snap.frozenCols || 0, cols.length), ySplit }]
    });
    ws.columns = cols.map(c => ({ key: c.key, width: Math.max(8, Math.min(60, Math.round((c.width || 120) / 7))) }));
    cols.forEach((c, i) => { if (c.excelFmt) ws.getColumn(i + 1).numFmt = c.excelFmt; });

    const xlCell = (cell, col) => {
      switch (cell.kind) {
        case 'num': return cell.v;
        case 'date': return new Date(cell.v);
        case 'text': return cleanText(cell.v);
        case 'link': return { text: cleanText(cell.v), hyperlink: cell.url };
        default: return null;
      }
    };
    const totalsValues = (tot, labelText) => cols.map(c => {
      if (c.key === (tot.__label || snap.totalsLabelKey)) return labelText;
      const v = tot[c.key];
      return v == null || !isFinite(v) ? null : v;
    });
    // Header, totals and subtotal colors and boldness follow the table's style settings.
    const st = snap.style || {};
    const argbOf = hex => 'FF' + String(hex).replace('#', '').toUpperCase();
    const fill = hex => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: argbOf(hex) } });
    const styleRow = (row, kind) => {
      if (kind === 'header') {
        row.font = Object.assign({ bold: st.headBold !== false }, st.headText ? { color: { argb: argbOf(st.headText) } } : {});
        row.fill = fill(st.headBg || '#F2F2F2');
        row.border = { bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } } };
      } else if (kind === 'group') {
        row.font = { bold: st.subBold !== false };
        row.fill = fill(st.subBg || '#F7F7F7');
      } else {
        row.font = { bold: st.totBold !== false };
        if (st.totBg) row.fill = fill(st.totBg);
      }
    };

    // Conditional formatting: fills and font colors are copied as static colors; symbols go into the number
    // format, so the cell is still a real number that sums and sorts in Excel.
    const argb = hex => 'FF' + hex.replace('#', '').toUpperCase();
    // The symbol goes after any [Red]-style color code, which Excel requires to come first in a section.
    const withSymbol = (fmt, sym) => (fmt || 'General').split(';').map(sec => {
      const lead = (sec.match(/^(\[[^\]]*\])*/) || [''])[0];
      const rest = sec.slice(lead.length);
      return lead + (rest.startsWith('"') ? '"' + sym + ' ' + rest.slice(1) : '"' + sym + ' "' + rest);
    }).join(';');
    const applyFormat = (row, values, isTotal) => {
      if (!snap.format) return;
      cols.forEach((c, k) => {
        if (!c.isNum) return;
        const v = values[k];
        if (typeof v !== 'number') return;
        const f = snap.format(c, v, isTotal);
        if (!f) return;
        const cell = row.getCell(k + 1);
        if (f.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(f.fill) } };
        if (f.font) cell.font = Object.assign({}, cell.font, { color: { argb: argb(f.font) } });
        if (f.sym) cell.numFmt = withSymbol(c.excelFmt, f.sym);
      });
    };

    // Subtotal rows sit above their detail rows, which are grouped with Excel's outline (+/- in the margin).
    // Other rows (Top N) are kept even when subtotal rows are left out, so the file still matches the screen.
    const keepGroup = r => r.type === 'group' && (snap.includeSubtotals || r.isOther);
    const outline = snap.rows.some(keepGroup);
    if (outline) {
      ws.properties.outlineProperties = { summaryBelow: false, summaryRight: false };
      ws.properties.outlineLevelRow = Math.max(...snap.rows.map(r => r.level || 0));
    }

    // Group headers: one merged, centered cell over each run of member columns, as on screen.
    if (groups.length) {
      const gr = ws.addRow(cols.map(() => null));
      styleRow(gr, 'header');
      for (const g of groups) {
        gr.getCell(g.start + 1).value = cleanText(g.label);
        if (g.end > g.start) ws.mergeCells(1, g.start + 1, 1, g.end + 1);
        gr.getCell(g.start + 1).alignment = { horizontal: 'center' };
      }
    }
    styleRow(ws.addRow(cols.map(c => cleanText(c.label))), 'header');
    if (top) { const totVals = totalsValues(snap.totals, 'Total'); const tr = ws.addRow(totVals); styleRow(tr, 'totals'); applyFormat(tr, totVals, true); }
    const firstData = ws.rowCount + 1;
    const dataRows = [];
    for (const r of snap.rows) {
      if (r.type === 'group') {
        if (!keepGroup(r)) continue;
        const gv = totalsValues(r.totals, r.label);
        const gr = ws.addRow(gv);
        styleRow(gr, 'group');
        applyFormat(gr, gv, true);
        if (r.level) gr.outlineLevel = r.level;
        if (r.hidden) gr.hidden = true;
      } else {
        const vals = cols.map(c => xlCell(snap.value(c, r.i), c));
        const row = ws.addRow(vals);
        dataRows.push(ws.rowCount);
        vals.forEach((v, k) => { if (v && v.hyperlink) row.getCell(k + 1).font = { color: { argb: 'FF0563C1' }, underline: true }; });
        cols.forEach((c, k) => { if (c.bold) { const cell = row.getCell(k + 1); cell.font = Object.assign({}, cell.font, { bold: true }); } });
        applyFormat(row, vals, false);
        if (outline && r.level) {
          row.outlineLevel = r.level;
          if (r.hidden) row.hidden = true;
        }
      }
    }
    const lastData = ws.rowCount;
    if (bottom) { const totVals = totalsValues(snap.totals, 'Total'); const tr = ws.addRow(totVals); styleRow(tr, 'totals'); applyFormat(tr, totVals, true); }
    // Column alignment overrides, below the (centered) group header row, including a bottom totals row.
    cols.forEach((c, k) => {
      if (!c.align) return;
      const horizontal = { l: 'left', c: 'center', r: 'right' }[c.align];
      ws.getColumn(k + 1).eachCell((cell, rowNo) => { if (rowNo >= headRow) cell.alignment = Object.assign({}, cell.alignment, { horizontal }); });
    });
    // Data bars use Excel's own data bars over the data rows, so they stay live if values change in Excel.
    // Bars cover data rows only (never subtotal or Other rows), as runs of consecutive rows sharing one scale.
    const runs = [];
    for (const n of dataRows) { const last = runs[runs.length - 1]; if (last && last[1] === n - 1) last[1] = n; else runs.push([n, n]); }
    for (const b of snap.barColumns || []) {
      const k = cols.findIndex(c => c.key === b.key);
      if (k < 0 || !runs.length) continue;
      const letter = ws.getColumn(k + 1).letter;
      ws.addConditionalFormatting({
        ref: runs.map(([a, z]) => letter + a + ':' + letter + z).join(' '),
        rules: [{ type: 'dataBar', gradient: false, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: 'FFA9CDEB' } }]
      });
    }
    // Excel's filter/sort would scramble subtotal groups, so it is only added to plain tables.
    if (!top && !outline && lastData >= firstData && cols.length) {
      ws.autoFilter = { from: { row: headRow, column: 1 }, to: { row: lastData, column: cols.length } };
    }
    return wb.xlsx.writeBuffer();
  }

  function buildFileName(template, ctx, ext, now) {
    now = now || new Date();
    const date = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
    const time = pad(now.getHours()) + pad(now.getMinutes());
    let name = String(template || '{sheet} {date}')
      .replace(/\{sheet\}/gi, ctx.sheet || '')
      .replace(/\{date\}/gi, date)
      .replace(/\{time\}/gi, time)
      .replace(/\{param:([^}]+)\}/gi, (m, p) => {
        const hit = Object.keys(ctx.params || {}).find(k => k.toLowerCase() === p.trim().toLowerCase());
        return hit ? ctx.params[hit] : '';
      });
    name = name.replace(/\.(xlsx|csv)$/i, '').replace(/[\\\/:*?"<>|\x00-\x1F]/g, '-').replace(/\s+/g, ' ').trim().replace(/^\.+/, '');
    return (name || 'export').slice(0, 150) + '.' + ext;
  }

  return { buildCsv, buildXlsx, buildFileName, roundAsShown, sheetName, isoFromMs };
});
