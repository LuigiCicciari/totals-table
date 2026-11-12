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
      case 'text': return cleanText(cell.v);
      default: return '';
    }
  }

  function csvEscape(s) { return /[",\r\n]/.test(s) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }

  // CSV is raw data only: a header line and every data row, never totals or subtotals.
  function buildCsv(snap) {
    const lines = [];
    const cols = snap.columns;
    const row = cells => lines.push(cells.map(csvEscape).join(','));
    row(cols.map(c => cleanText(c.label)));
    for (const r of snap.rows) {
      if (r.type !== 'data') continue;
      row(cols.map(c => csvCell(snap.value(c, r.i), c)));
    }
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  function sheetName(s) {
    const t = String(s || 'Data').replace(/[\\\/\?\*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '');
    return (t || 'Data').slice(0, 31);
  }

  async function buildXlsx(snap, ExcelJS) {
    const cols = snap.columns;
    const wb = new ExcelJS.Workbook();
    const top = snap.includeTotals && snap.totalsLocation === 'top';
    const bottom = snap.includeTotals && snap.totalsLocation === 'bottom';
    const ySplit = top ? 2 : 1;
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
        default: return null;
      }
    };
    const totalsValues = (tot, labelText) => cols.map(c => {
      if (c.key === snap.totalsLabelKey) return labelText;
      const v = tot[c.key];
      return v == null || !isFinite(v) ? null : v;
    });
    const styleRow = (row, kind) => {
      row.font = { bold: true };
      if (kind === 'header') {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        row.border = { bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } } };
      } else if (kind === 'group') {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F7F7' } };
      }
    };

    // Subtotal rows sit above their detail rows, which are grouped with Excel's outline (+/- in the margin).
    const outline = snap.includeSubtotals && snap.rows.some(r => r.type === 'group');
    if (outline) ws.properties.outlineProperties = { summaryBelow: false, summaryRight: false };

    styleRow(ws.addRow(cols.map(c => cleanText(c.label))), 'header');
    if (top) styleRow(ws.addRow(totalsValues(snap.totals, 'Total')), 'totals');
    const firstData = ws.rowCount + 1;
    for (const r of snap.rows) {
      if (r.type === 'group') {
        if (!outline) continue;
        styleRow(ws.addRow(totalsValues(r.totals, r.label)), 'group');
      } else {
        const row = ws.addRow(cols.map(c => xlCell(snap.value(c, r.i), c)));
        if (outline && r.level) {
          row.outlineLevel = r.level;
          if (r.hidden) row.hidden = true;
        }
      }
    }
    const lastData = ws.rowCount;
    if (bottom) styleRow(ws.addRow(totalsValues(snap.totals, 'Total')), 'totals');
    // Excel's filter/sort would scramble subtotal groups, so it is only added to plain tables.
    if (!top && !outline && lastData >= firstData && cols.length) {
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: lastData, column: cols.length } };
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
