/* global TPRules */
(function (root) {
  'use strict';

  const PAGE_ROWS = 10000;
  const R = root.TPRules;

  class Store {
    constructor(columns) {
      this.columns = columns;
      this.n = 0;
      this.raw = {};
      this.fmt = {};
      this.dates = {};
      this.tuple = [];
      for (const c of columns) { this.raw[c.key] = []; this.fmt[c.key] = []; }
      this.complete = false;
    }

    addPage(page, colIndex) {
      const data = page.data;
      const marks = page.marksInfo;
      for (const c of this.columns) {
        const j = colIndex[c.key];
        const raw = this.raw[c.key];
        const fmt = this.fmt[c.key];
        for (let r = 0; r < data.length; r++) {
          const dv = data[r][j];
          if (!dv || dv.value === '%null%' || dv.nativeValue === null) { raw.push(c.isNum ? NaN : null); fmt.push(null); continue; }
          const f = dv.formattedValue != null ? String(dv.formattedValue) : String(dv.value);
          fmt.push(f);
          if (c.isNum) {
            const v = typeof dv.nativeValue === 'number' ? dv.nativeValue : Number(dv.value);
            raw.push(isFinite(v) ? v : NaN);
          } else if (c.isDate) {
            let d = R.parseIsoDate(typeof dv.value === 'string' ? dv.value : '');
            if (!d && dv.nativeValue instanceof Date) {
              const n = dv.nativeValue;
              d = { t: Date.UTC(n.getFullYear(), n.getMonth(), n.getDate(), n.getHours(), n.getMinutes(), n.getSeconds()) };
            }
            raw.push(d ? d.t : null);
          } else {
            raw.push(f);
          }
        }
      }
      // Tableau identifies a summary-data mark by its 1-based row position unless marksInfo says otherwise.
      for (let r = 0; r < data.length; r++) {
        const m = marks && marks[r] && marks[r].tupleId != null ? marks[r].tupleId : null;
        this.tuple.push(m != null ? m : this.n + r + 1);
      }
      this.n += data.length;
    }

    // Number formats and text-that-is-really-a-date detection, sampled from the loaded rows.
    analyse() {
      const limit = Math.min(this.n, 2000);
      for (const c of this.columns) {
        const raw = this.raw[c.key], fmt = this.fmt[c.key];
        if (c.isNum) {
          const samples = [];
          for (let i = 0; i < limit; i++) if (fmt[i] != null) samples.push({ v: raw[i], f: fmt[i] });
          c.spec = R.inferSpec(samples);
          c.excelFmt = R.excelNumFmt(c.spec);
        } else if (c.isDate) {
          const f = [];
          for (let i = 0; i < limit && f.length < 20; i++) if (fmt[i]) f.push(fmt[i]);
          c.excelFmt = R.inferDateFmt(f, c.dataType === 'date-time');
        } else {
          let nonNull = 0, iso = 0, time = false;
          for (let i = 0; i < limit; i++) {
            const s = fmt[i];
            if (s == null || s === '') continue;
            nonNull++;
            const d = R.parseIsoDate(s.trim());
            if (d && s.trim().length <= 19) { iso++; time = time || d.time; } else break;
          }
          c.isoText = nonNull > 0 && iso === nonNull;
          if (c.isoText) {
            c.excelFmt = time ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd';
            const arr = new Array(this.n);
            for (let i = 0; i < this.n; i++) { const d = fmt[i] ? R.parseIsoDate(fmt[i].trim()) : null; arr[i] = d ? d.t : null; }
            this.dates[c.key] = arr;
          } else {
            delete this.dates[c.key];
          }
        }
      }
    }
  }

  function makeColumns(pageCols) {
    return pageCols.map((c, i) => ({
      key: c.fieldName,
      fieldId: c.fieldId,
      name: R.normName(c.fieldName),
      dataType: c.dataType,
      isNum: c.dataType === 'int' || c.dataType === 'float',
      isDate: c.dataType === 'date' || c.dataType === 'date-time',
      index: c.index != null ? c.index : i
    }));
  }

  // Detail-shelf fields in shelf order; fields only on Tooltip are left out of the table.
  async function shelfInfo(ws) {
    try {
      const spec = await ws.getVisualSpecificationAsync();
      const ms = spec.marksSpecifications[spec.activeMarksSpecificationIndex || 0];
      const encs = (ms && ms.encodings) || [];
      const kind = e => String(e.type || e.id || '').toLowerCase();
      const detail = encs.filter(e => kind(e).indexOf('detail') >= 0 && e.field).map(e => e.field);
      const tooltip = encs.filter(e => kind(e).indexOf('tooltip') >= 0 && e.field).map(e => e.field);
      return { detail, tooltip };
    } catch (e) {
      return null;
    }
  }

  function sameField(col, f) {
    if (f.id && col.fieldId && f.id === col.fieldId) return true;
    const n = String(f.name || '');
    return n === col.key || R.normName(n).toLowerCase() === col.name.toLowerCase();
  }

  function orderColumns(columns, shelf) {
    if (!shelf || !shelf.detail.length) return columns;
    const out = [];
    const used = new Set();
    for (const f of shelf.detail) {
      const c = columns.find(x => !used.has(x.key) && sameField(x, f));
      if (c) { used.add(c.key); out.push(c); }
    }
    for (const c of columns) {
      if (used.has(c.key)) continue;
      const onlyTooltip = shelf.tooltip.some(f => sameField(c, f));
      if (!onlyTooltip) out.push(c);
    }
    return out.length ? out : columns;
  }

  async function openReader(ws) {
    const base = { ignoreSelection: true, includeDataValuesOption: 'all-values' };
    try { return await ws.getSummaryDataReaderAsync(PAGE_ROWS, Object.assign({ applyWorksheetFormatting: true }, base)); } catch (e) { /* older Tableau */ }
    return ws.getSummaryDataReaderAsync(PAGE_ROWS, base);
  }

  // Streams the worksheet's data into a Store, calling onPage after each page so the table can draw early.
  async function load(ws, isCurrent, onPage) {
    const timing = { start: performance.now() };
    const [shelf, reader] = await Promise.all([shelfInfo(ws), openReader(ws)]);
    timing.query = performance.now() - timing.start;
    let store = null;
    try {
      for (let p = 0; p < reader.pageCount; p++) {
        if (!isCurrent()) return null;
        const page = await reader.getPageAsync(p);
        if (!store) {
          const cols = orderColumns(makeColumns(page.columns), shelf);
          store = new Store(cols);
          store.colIndex = {};
          page.columns.forEach((c, i) => { store.colIndex[c.fieldName] = c.index != null ? c.index : i; });
          store.total = reader.totalRowCount;
        }
        store.addPage(page, store.colIndex);
        if (p === 0 || p === reader.pageCount - 1) store.analyse();
        timing.fetch = performance.now() - timing.start;
        if (p < reader.pageCount - 1) onPage(store, timing, false);
      }
    } finally {
      try { await reader.releaseAsync(); } catch (e) { /* already released */ }
    }
    if (!store) {
      let cols = [];
      try {
        const r = await ws.getSummaryColumnsInfoAsync();
        cols = Array.isArray(r) ? r : (r && r.columns) || [];
      } catch (e) { /* no column info available */ }
      store = new Store(orderColumns(makeColumns(cols), shelf));
      store.total = 0;
      store.analyse();
    }
    store.complete = true;
    timing.fetch = performance.now() - timing.start;
    onPage(store, timing, true);
    return store;
  }

  root.TPData = { load, Store, makeColumns, orderColumns };
})(this);
