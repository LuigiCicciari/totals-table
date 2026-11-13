/* global tableau, TPRules, TPData, TPExport, ExcelJS */
(function () {
  'use strict';

  const R = window.TPRules;
  const SETTINGS_KEY = 'tp-config';
  const ROW_H = 26;
  const OVERSCAN = 12;
  const MIN_W = 48;
  const FREEZE_MAX = 0.6;

  function defaults() {
    return {
      v: 1,
      showFilters: false,
      filterMode: 'contains',
      showSearch: false,
      showTooltips: true,
      allowViewerLayout: true,
      nullText: { string: 'Null', number: 'Null', date: 'Null' },
      totals: { show: true, location: 'top' },
      autoHideNone: true,
      noneValues: ['(none)', 'none'],
      headerLinks: {},
      totalOverrides: {},
      rules: null,
      export: { formats: 'excel', excelTotals: true, subtotals: true, filename: '{sheet} {date}' },
      layout: { order: [], hidden: [], widths: {}, labels: {}, freeze: 'auto', sort: null, groupBy: null }
    };
  }

  function mergeConfig(saved) {
    const d = defaults();
    if (!saved) return d;
    const out = Object.assign(d, saved);
    out.nullText = Object.assign(defaults().nullText, saved.nullText);
    out.totals = Object.assign(defaults().totals, saved.totals);
    out.export = Object.assign(defaults().export, saved.export);
    out.layout = Object.assign(defaults().layout, saved.layout);
    out.headerLinks = saved.headerLinks || {};
    out.totalOverrides = saved.totalOverrides || {};
    return out;
  }

  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

  class TableApp {
    constructor(rootEl) {
      this.el = rootEl;
      this.cfg = defaults();
      this.store = null;
      this.params = [];
      this.paramObjs = [];
      this.local = null;
      this.filters = {};
      this.search = '';
      this.collapsed = new Set();
      this.selected = new Set();
      this.range = null;
      this.gen = 0;
      this.timing = null;
      this.view = { cols: [], disp: [], idx: [], totals: {}, groups: 0 };
      this.buildDom();
    }

    // ---------- environment ----------

    get authoring() { return tableau.extensions.environment.mode === 'authoring'; }
    get ws() { return tableau.extensions.worksheetContent.worksheet; }

    async start() {
      await tableau.extensions.initializeAsync({ configure: () => this.openFormat() });
      this.loadConfig();
      const ws = this.ws;
      ws.addEventListener(tableau.TableauEventType.SummaryDataChanged, () => this.reload());
      tableau.extensions.settings.addEventListener(tableau.TableauEventType.SettingsChanged, () => { this.loadConfig(); this.refresh(); });
      await this.loadParams();
      this.reload();
    }

    loadConfig() {
      let saved = null;
      try { saved = JSON.parse(tableau.extensions.settings.get(SETTINGS_KEY) || 'null'); } catch (e) { saved = null; }
      this.cfg = mergeConfig(saved);
    }

    async saveConfig() {
      if (!this.authoring) return;
      tableau.extensions.settings.set(SETTINGS_KEY, JSON.stringify(this.cfg));
      clearTimeout(this.saveTimer);
      this.saveTimer = setTimeout(() => tableau.extensions.settings.saveAsync().catch(e => this.flash('Could not save: ' + e.message)), 600);
    }

    async loadParams() {
      try {
        const ps = await this.ws.getParametersAsync();
        this.paramObjs = ps;
        this.params = ps.map(p => ({ name: p.name, value: p.currentValue ? String(p.currentValue.formattedValue) : '' }));
        ps.forEach(p => p.addEventListener(tableau.TableauEventType.ParameterChanged, () => this.onParamChanged()));
      } catch (e) {
        this.params = [];
      }
    }

    async onParamChanged() {
      try {
        const ps = await this.ws.getParametersAsync();
        this.params = ps.map(p => ({ name: p.name, value: p.currentValue ? String(p.currentValue.formattedValue) : '' }));
      } catch (e) { /* keep previous values */ }
      this.refresh();
    }

    reload() {
      const gen = ++this.gen;
      this.setStatus('Loading…');
      TPData.load(this.ws, () => gen === this.gen, (store, timing, done) => {
        if (gen !== this.gen) return;
        const first = !this.store || this.store !== store;
        this.store = store;
        this.timing = timing;
        if (first) this.loadLocalLayout();
        this.refresh();
        if (done) this.setStatus(null);
      }).catch(e => {
        if (gen === this.gen) this.setStatus('Could not load data: ' + ((e && e.message) || e), true);
      });
    }

    // ---------- layout ----------

    layoutKey() {
      const keys = this.store ? this.store.columns.map(c => c.key).sort().join('|') : '';
      let h = 0;
      for (let i = 0; i < keys.length; i++) h = (h * 31 + keys.charCodeAt(i)) | 0;
      return 'tp-layout:' + this.ws.name + ':' + (h >>> 0).toString(36);
    }

    loadLocalLayout() {
      this.local = null;
      if (this.authoring || !this.cfg.allowViewerLayout) return;
      try { this.local = JSON.parse(localStorage.getItem(this.layoutKey()) || 'null'); } catch (e) { this.local = null; }
    }

    get layout() {
      const base = this.cfg.layout;
      const out = !this.local ? base : Object.assign({}, base, this.local, {
        widths: Object.assign({}, base.widths, this.local.widths),
        labels: Object.assign({}, base.labels, this.local.labels)
      });
      return this.sessionSort === undefined ? out : Object.assign({}, out, { sort: this.sessionSort });
    }

    canEditLayout() { return this.authoring || this.cfg.allowViewerLayout; }

    updateLayout(patch) {
      if (!this.canEditLayout()) return;
      if (this.authoring) {
        this.cfg.layout = Object.assign({}, this.cfg.layout, patch);
        this.saveConfig();
      } else {
        this.local = Object.assign({}, this.local || {}, patch);
        try { localStorage.setItem(this.layoutKey(), JSON.stringify(this.local)); } catch (e) { /* storage blocked */ }
      }
      this.refresh();
    }

    resetLocal() {
      this.local = null;
      try { localStorage.removeItem(this.layoutKey()); } catch (e) { /* storage blocked */ }
      this.refresh();
    }

    orderedKeys() {
      const all = this.store.columns.map(c => c.key);
      const saved = (this.layout.order || []).filter(k => all.includes(k));
      return saved.concat(all.filter(k => !saved.includes(k)));
    }

    colDef(key) { return this.store.columns.find(c => c.key === key); }

    paramFor(col) {
      const link = this.cfg.headerLinks[col.key];
      if (link === '') return null;
      const name = link || col.name;
      return this.params.find(p => p.name.toLowerCase() === String(name).toLowerCase()) || null;
    }

    labelFor(col) {
      const l = this.layout.labels && this.layout.labels[col.key];
      if (l) return l;
      const p = this.paramFor(col);
      return p ? p.value : col.name;
    }

    autoHidden(col) {
      if (!this.cfg.autoHideNone) return false;
      const p = this.paramFor(col);
      if (!p) return false;
      const v = p.value.trim().toLowerCase();
      return v === '' || this.cfg.noneValues.some(n => n.toLowerCase() === v);
    }

    // ---------- pipeline ----------

    refresh() {
      if (!this.store) return;
      const t0 = performance.now();
      this.computeView();
      this.renderHead();
      this.renderBody(true);
      this.renderBar();
      if (this.timing) this.timing.draw = performance.now() - t0;
      this.renderBar();
    }

    computeView() {
      const st = this.store;
      const lay = this.layout;
      const hidden = new Set(lay.hidden || []);
      const cols = [];
      for (const k of this.orderedKeys()) {
        const c = this.colDef(k);
        if (!c || hidden.has(k) || this.autoHidden(c)) continue;
        cols.push(c);
      }
      const widths = lay.widths || {};
      let left = 0;
      const nFrozen = this.frozenCount(cols);
      cols.forEach((c, i) => {
        c.w = Math.max(MIN_W, widths[c.key] || (c.isNum ? 110 : 140));
        c.label = this.labelFor(c);
        c.frozen = i < nFrozen;
        c.left = left;
        if (c.frozen) left += c.w;
      });
      this.frozenWidth = left;

      const idx = this.filterRows(cols);
      this.sortRows(idx, lay.sort, cols);
      const totals = this.computeTotals(cols, idx);
      const disp = this.groupRows(cols, idx, lay.groupBy);
      this.view = { cols, idx, disp, totals, nFrozen };
    }

    frozenCount(cols) {
      const f = this.layout.freeze;
      let n;
      if (typeof f === 'number') n = f;
      else {
        n = 0;
        while (n < cols.length && this.paramFor(cols[n])) n++;
        if (!n) while (n < cols.length && !cols[n].isNum) n++;
      }
      n = Math.max(0, Math.min(n, cols.length));
      const width = cols.slice(0, n).reduce((s, c) => s + Math.max(MIN_W, (this.layout.widths || {})[c.key] || (c.isNum ? 110 : 140)), 0);
      this.freezePaused = n > 0 && width > this.scroller.clientWidth * FREEZE_MAX;
      return this.freezePaused ? 0 : n;
    }

    filterRows(cols) {
      const st = this.store;
      const n = st.n;
      const tests = [];
      const mode = this.cfg.filterMode;
      for (const c of cols) {
        const f = this.filters[c.key];
        if (!f) continue;
        if (c.isNum && (f.min !== '' && f.min != null || f.max !== '' && f.max != null)) {
          const lo = f.min === '' || f.min == null ? -Infinity : Number(f.min);
          const hi = f.max === '' || f.max == null ? Infinity : Number(f.max);
          const raw = st.raw[c.key];
          const scale = c.spec && c.spec.pct ? 100 : 1;
          tests.push(i => { const v = raw[i] * scale; return v >= lo && v <= hi; });
        } else if (!c.isNum && f.text) {
          const q = f.text.toLowerCase();
          const fmt = st.fmt[c.key];
          if (mode === 'startswith') tests.push(i => fmt[i] != null && fmt[i].toLowerCase().startsWith(q));
          else if (mode === 'exact') tests.push(i => fmt[i] != null && fmt[i].toLowerCase() === q);
          else tests.push(i => fmt[i] != null && fmt[i].toLowerCase().indexOf(q) >= 0);
        }
      }
      if (this.search) {
        const q = this.search.toLowerCase();
        const fmts = cols.map(c => st.fmt[c.key]);
        tests.push(i => fmts.some(f => f[i] != null && f[i].toLowerCase().indexOf(q) >= 0));
      }
      const idx = [];
      outer: for (let i = 0; i < n; i++) {
        for (let t = 0; t < tests.length; t++) if (!tests[t](i)) continue outer;
        idx.push(i);
      }
      return idx;
    }

    sortKeyArray(c) {
      const st = this.store;
      if (c.isNum || c.isDate) return { arr: st.raw[c.key], num: true };
      if (st.dates[c.key]) return { arr: st.dates[c.key], num: true };
      return { arr: st.fmt[c.key], num: false };
    }

    // With no sort chosen, order rows A to Z by the text columns on the left, like a crosstab.
    defaultSort(idx, cols) {
      const keys = [];
      for (const c of cols) { if (c.isNum) break; keys.push(this.sortKeyArray(c)); }
      if (!keys.length) return;
      idx.sort((a, b) => {
        for (const { arr, num } of keys) {
          const x = arr[a], y = arr[b];
          const xn = x == null || x !== x, yn = y == null || y !== y;
          if (xn || yn) { if (xn !== yn) return xn ? 1 : -1; continue; }
          const d = num ? (x < y ? -1 : x > y ? 1 : 0) : collator.compare(x, y);
          if (d) return d;
        }
        return a - b;
      });
    }

    sortRows(idx, sort, cols) {
      if (!sort || !sort.key) return this.defaultSort(idx, cols);
      const c = this.colDef(sort.key);
      if (!c) return;
      const { arr, num } = this.sortKeyArray(c);
      const dir = sort.dir === 'desc' ? -1 : 1;
      if (num) {
        idx.sort((a, b) => {
          const x = arr[a], y = arr[b];
          const xn = x == null || x !== x, yn = y == null || y !== y;
          if (xn || yn) return xn === yn ? a - b : xn ? 1 : -1;
          return x === y ? a - b : (x < y ? -dir : dir);
        });
      } else {
        idx.sort((a, b) => {
          const x = arr[a], y = arr[b];
          if (x == null || y == null) return x == null && y == null ? a - b : x == null ? 1 : -1;
          return dir * collator.compare(x, y) || a - b;
        });
      }
    }

    totalSpec(c) {
      const t = R.resolveTotal(c, this.store.columns, this.cfg.rules && this.cfg.rules.doc, this.cfg.totalOverrides[c.key]);
      t.text = R.describeTotal(t, this.store.columns);
      return t;
    }

    computeTotals(cols, idx) {
      const out = {};
      const raw = k => this.store.raw[k];
      for (const c of cols) {
        if (!c.isNum) continue;
        c.total = this.totalSpec(c);
        out[c.key] = R.computeTotal(c.total, c.key, idx, raw);
      }
      out.__label = (cols.find(c => !c.isNum) || {}).key || null;
      return out;
    }

    groupRows(cols, idx, groupBy) {
      const gc = groupBy && cols.find(c => c.key === groupBy);
      this.groups = null;
      if (!gc) return idx;
      const fmt = this.store.fmt[gc.key];
      const groups = new Map();
      for (const i of idx) {
        const k = fmt[i] == null ? '\u0000null' : fmt[i];
        let g = groups.get(k);
        if (!g) groups.set(k, g = { key: k, label: fmt[i] == null ? this.nullFor(gc) : fmt[i], rows: [] });
        g.rows.push(i);
      }
      const list = Array.from(groups.values());
      this.groups = list;
      const raw = k => this.store.raw[k];
      for (const g of list) {
        g.type = 'group';
        g.totals = {};
        for (const c of cols) if (c.isNum) g.totals[c.key] = R.computeTotal(c.total, c.key, g.rows, raw);
      }
      // Sorting a number column orders the groups by their subtotal as well as the rows inside them.
      const sort = this.layout.sort;
      const sc = sort && cols.find(c => c.key === sort.key);
      const byLabel = (a, b) => collator.compare(a.label, b.label);
      if (sc && sc.isNum && sc.key !== gc.key && list.some(g => g.totals[sc.key] != null)) {
        const dir = sort.dir === 'desc' ? -1 : 1;
        list.sort((a, b) => {
          const x = a.totals[sc.key], y = b.totals[sc.key];
          if (x == null || y == null) return x == null && y == null ? byLabel(a, b) : x == null ? 1 : -1;
          return x === y ? byLabel(a, b) : (x < y ? -dir : dir);
        });
      } else if (!sort || sort.key !== gc.key) {
        list.sort(byLabel);
      }
      const disp = [];
      for (const g of list) {
        disp.push(g);
        if (!this.collapsed.has(g.key)) for (const i of g.rows) disp.push(i);
      }
      return disp;
    }

    nullFor(c) {
      const t = this.cfg.nullText;
      return c.isNum ? t.number : c.isDate ? t.date : t.string;
    }

    // ---------- DOM ----------

    buildDom() {
      this.el.innerHTML = `
        <div class="tp-bar">
          <input class="tp-search" type="search" placeholder="Search" aria-label="Search table" hidden>
          <span class="tp-status" role="status" aria-live="polite"></span>
          <span class="tp-spacer"></span>
          <button class="tp-chip" type="button" hidden title="Your column changes are saved in this browser only">Custom layout · Reset</button>
          <button class="tp-icon tp-cols" type="button" title="Show and hide columns" aria-label="Show and hide columns">${ICONS.cols}</button>
          <button class="tp-icon tp-dl" type="button" title="Download" aria-label="Download">${ICONS.dl}</button>
          <button class="tp-icon tp-gear" type="button" title="Format table" aria-label="Format table" hidden>${ICONS.gear}</button>
          <button class="tp-icon tp-info" type="button" aria-label="Load timing">${ICONS.info}</button>
        </div>
        <div class="tp-scroll" tabindex="0">
          <div class="tp-grid">
            <div class="tp-top"><div class="tp-head"></div><div class="tp-tot tp-tot-top"></div></div>
            <div class="tp-body"><div class="tp-pad"></div><div class="tp-rows"></div></div>
            <div class="tp-tot tp-tot-bottom"></div>
          </div>
        </div>
        <div class="tp-pop" hidden></div>`;
      const q = s => this.el.querySelector(s);
      this.scroller = q('.tp-scroll');
      this.grid = q('.tp-grid');
      this.head = q('.tp-head');
      this.totTop = q('.tp-tot-top');
      this.totBottom = q('.tp-tot-bottom');
      this.body = q('.tp-body');
      this.pad = q('.tp-pad');
      this.rowsEl = q('.tp-rows');
      this.pop = q('.tp-pop');
      this.statusEl = q('.tp-status');
      this.searchEl = q('.tp-search');

      this.scroller.addEventListener('scroll', () => { if (!this.raf) this.raf = requestAnimationFrame(() => { this.raf = 0; this.renderBody(false); }); });
      window.addEventListener('resize', () => this.refresh());
      this.searchEl.addEventListener('input', () => { clearTimeout(this.searchT); this.searchT = setTimeout(() => { this.search = this.searchEl.value.trim(); this.refresh(); }, 150); });
      q('.tp-chip').addEventListener('click', () => this.resetLocal());
      q('.tp-cols').addEventListener('click', e => this.openColumnsPicker(e.currentTarget));
      q('.tp-dl').addEventListener('click', e => this.onDownloadClick(e.currentTarget));
      q('.tp-gear').addEventListener('click', () => this.openFormat());
      document.addEventListener('mousedown', e => { if (!this.pop.hidden && !this.pop.contains(e.target)) this.closePop(); });
      document.addEventListener('keydown', e => this.onKey(e));
      // Any native drag would hand text to Tableau, which tries to open it as a new data source.
      document.addEventListener('dragstart', e => e.preventDefault());
      this.wireHead();
      this.wireBody();
    }

    setStatus(msg, isError) {
      this.statusMsg = msg;
      this.statusErr = !!isError;
      this.renderBar();
    }

    flash(msg) {
      this.setStatus(msg, true);
      clearTimeout(this.flashT);
      this.flashT = setTimeout(() => this.setStatus(null), 5000);
    }

    renderBar() {
      const q = s => this.el.querySelector(s);
      const st = this.store;
      let text = this.statusMsg;
      if (!text && st) {
        const shown = this.view.idx.length;
        const rows = n => n.toLocaleString() + (n === 1 ? ' row' : ' rows');
        text = shown === st.n ? rows(st.n) : shown.toLocaleString() + ' of ' + rows(st.n);
        if (!st.complete) text = 'Loading ' + st.n.toLocaleString() + ' of ' + (st.total || 0).toLocaleString() + '…';
      }
      this.statusEl.textContent = text || '';
      this.statusEl.classList.toggle('err', !!this.statusErr);
      this.searchEl.hidden = !this.cfg.showSearch;
      q('.tp-gear').hidden = !this.authoring;
      q('.tp-cols').hidden = !this.canEditLayout();
      q('.tp-chip').hidden = !(this.local && Object.keys(this.local).length) || this.authoring;
      const t = this.timing;
      q('.tp-info').title = t ? 'Tableau query ' + (t.query / 1000).toFixed(2) + 's · all rows loaded ' + ((t.fetch || 0) / 1000).toFixed(2) + 's · last draw ' + Math.round(t.draw || 0) + ' ms' + (this.freezePaused ? '\nFreeze paused: frozen columns are wider than 60% of the table.' : '') : '';
    }

    template() { return this.view.cols.map(c => c.w + 'px').join(' '); }

    cellAttrs(c, extra) {
      let cls = 'tp-c' + (c.isNum ? ' num' : '') + (extra ? ' ' + extra : '');
      let style = '';
      if (c.frozen) {
        cls += ' fz';
        style = ` style="left:${c.left}px"`;
        if (c === this.view.cols[this.view.nFrozen - 1]) cls += ' fz-edge';
      }
      return `class="${cls}"${style}`;
    }

    renderHead() {
      const { cols, totals } = this.view;
      this.grid.style.setProperty('--cols', this.template());
      this.grid.style.width = cols.reduce((s, c) => s + c.w, 0) + 'px';
      const sort = this.layout.sort || {};
      const canEdit = this.canEditLayout();
      let h = '<div class="tp-hrow">';
      cols.forEach((c, i) => {
        const arrow = sort.key === c.key ? (sort.dir === 'desc' ? ' ▼' : ' ▲') : '';
        const sortTip = sort.key === c.key ? '\nSorted ' + (c.isNum ? (sort.dir === 'desc' ? 'largest to smallest' : 'smallest to largest') : (sort.dir === 'desc' ? 'Z to A' : 'A to Z')) : '';
        const p = this.paramFor(c);
        h += `<div ${this.cellAttrs(c, 'tp-h')} data-i="${i}" title="${esc(c.label + (p ? '  (from parameter ' + p.name + ')' : '') + (c.label !== c.name ? '\nField: ' + c.name : '') + sortTip)}">` +
          `<span class="tp-hl">${esc(c.label)}${arrow}</span>` +
          `<button class="tp-hm" type="button" data-i="${i}" aria-label="Column menu for ${esc(c.label)}">▾</button>` +
          (canEdit ? `<span class="tp-rz" data-i="${i}"></span>` : '') + '</div>';
      });
      h += '</div>';
      if (this.cfg.showFilters) {
        h += '<div class="tp-frow">';
        cols.forEach((c, i) => {
          const f = this.filters[c.key] || {};
          h += `<div ${this.cellAttrs(c, 'tp-f')}>` + (c.isNum
            ? `<input data-i="${i}" data-k="min" type="number" step="any" placeholder="min" value="${esc(f.min || '')}" aria-label="Minimum ${esc(c.label)}"><input data-i="${i}" data-k="max" type="number" step="any" placeholder="max" value="${esc(f.max || '')}" aria-label="Maximum ${esc(c.label)}">`
            : `<input data-i="${i}" data-k="text" type="search" placeholder="Filter" value="${esc(f.text || '')}" aria-label="Filter ${esc(c.label)}">`) + '</div>';
        });
        h += '</div>';
      }
      this.head.innerHTML = h;
      if (canEdit && cols.length) {
        const n = this.view.nFrozen;
        const host = this.head.querySelectorAll('.tp-h')[n ? n - 1 : 0];
        const fz = document.createElement('span');
        fz.className = 'tp-fzh' + (n ? '' : ' at-start');
        fz.title = 'Drag to change frozen columns';
        host.appendChild(fz);
      }

      const show = this.cfg.totals.show && cols.some(c => c.isNum);
      const loc = this.cfg.totals.location;
      const totHtml = show ? this.totalsHtml(totals, 'Total', 'tp-trow') : '';
      this.totTop.innerHTML = loc === 'top' ? totHtml : '';
      this.totBottom.innerHTML = loc === 'bottom' ? totHtml : '';
    }

    totalsHtml(tot, label, cls) {
      const { cols } = this.view;
      const loading = !this.store.complete;
      let h = `<div class="${cls}">`;
      cols.forEach(c => {
        if (c.key === tot.__label || (!tot.__label && c === cols[0] && !c.isNum)) {
          h += `<div ${this.cellAttrs(c)}>${esc(label)}</div>`;
        } else if (c.isNum) {
          const v = tot[c.key];
          const text = loading ? '…' : v == null ? '' : R.formatNumber(v, c.spec);
          const title = (c.total ? c.total.text : '') + (c.total && c.total.problem ? ' — ' + c.total.problem : '') + (c.total && c.total.source === 'override' ? ' (set for this column)' : c.total && c.total.source === 'rule' ? ' (' + c.total.label + ')' : '');
          h += `<div ${this.cellAttrs(c, c.total && c.total.problem ? 'warn' : '')} title="${esc(title)}">${esc(text)}</div>`;
        } else {
          h += `<div ${this.cellAttrs(c)}></div>`;
        }
      });
      return h + '</div>';
    }

    renderBody(force) {
      const { disp, cols } = this.view;
      const total = disp.length;
      const viewH = this.scroller.clientHeight;
      const top = Math.max(0, this.scroller.scrollTop - this.body.offsetTop);
      const start = Math.max(0, Math.floor(top / ROW_H) - OVERSCAN);
      const end = Math.min(total, Math.ceil((top + viewH) / ROW_H) + OVERSCAN);
      if (!force && start === this.rStart && end === this.rEnd) return;
      this.rStart = start;
      this.rEnd = end;
      this.body.style.height = total * ROW_H + 'px';
      this.pad.style.height = start * ROW_H + 'px';
      const st = this.store;
      const nulls = cols.map(c => esc(this.nullFor(c)));
      const fmts = cols.map(c => st.fmt[c.key]);
      const rg = this.normRange();
      let h = '';
      for (let r = start; r < end; r++) {
        const d = disp[r];
        if (typeof d === 'object') {
          const open = !this.collapsed.has(d.key);
          h += `<div class="tp-row tp-grp" data-r="${r}">`;
          cols.forEach((c, j) => {
            if (c.key === this.layout.groupBy) h += `<div ${this.cellAttrs(c)} data-c="${j}"><span class="tp-caret">${open ? '▾' : '▸'}</span>${esc(d.label)} <span class="tp-cnt">(${d.rows.length.toLocaleString()})</span></div>`;
            else if (c.isNum) h += `<div ${this.cellAttrs(c)} data-c="${j}">${esc(d.totals[c.key] == null ? '' : R.formatNumber(d.totals[c.key], c.spec))}</div>`;
            else h += `<div ${this.cellAttrs(c)} data-c="${j}"></div>`;
          });
          h += '</div>';
          continue;
        }
        const sel = st.tuple[d] != null && this.selected.has(st.tuple[d]);
        h += `<div class="tp-row${sel ? ' sel' : ''}${r % 2 ? ' odd' : ''}" data-r="${r}">`;
        for (let j = 0; j < cols.length; j++) {
          const v = fmts[j][d];
          const inRange = rg && r >= rg.r0 && r <= rg.r1 && j >= rg.c0 && j <= rg.c1;
          h += `<div ${this.cellAttrs(cols[j], (v == null ? 'nul' : '') + (inRange ? ' rg' : ''))} data-c="${j}">${v == null ? nulls[j] : esc(v)}</div>`;
        }
        h += '</div>';
      }
      this.rowsEl.innerHTML = h;
    }

    // ---------- header interactions ----------

    wireHead() {
      this.head.addEventListener('click', e => {
        const menu = e.target.closest('.tp-hm');
        if (menu) { this.openColumnMenu(+menu.dataset.i, menu); return; }
        if (this.suppressClick || e.target.closest('.tp-rz,.tp-fzh,input')) return;
        const cell = e.target.closest('.tp-h');
        if (!cell) return;
        const c = this.view.cols[+cell.dataset.i];
        const s = this.layout.sort;
        const next = !s || s.key !== c.key ? { key: c.key, dir: c.isNum ? 'desc' : 'asc' } : s.dir === (c.isNum ? 'desc' : 'asc') ? { key: c.key, dir: c.isNum ? 'asc' : 'desc' } : null;
        this.setSort(next);
      });
      this.head.addEventListener('input', e => {
        const inp = e.target.closest('input[data-i]');
        if (!inp) return;
        const c = this.view.cols[+inp.dataset.i];
        const f = this.filters[c.key] = Object.assign({}, this.filters[c.key]);
        f[inp.dataset.k] = inp.value;
        clearTimeout(this.filterT);
        this.filterT = setTimeout(() => {
          const focusKey = inp.dataset.i + ':' + inp.dataset.k;
          this.computeView();
          this.renderBody(true);
          this.renderTotalsOnly();
          this.renderBar();
          const again = this.head.querySelector(`input[data-i="${focusKey.split(':')[0]}"][data-k="${focusKey.split(':')[1]}"]`);
          if (again && document.activeElement !== again) again.focus();
        }, 200);
      });
      this.head.addEventListener('mousedown', e => {
        if (e.button !== 0) return;
        const rz = e.target.closest('.tp-rz');
        if (rz) return this.startResize(e, +rz.dataset.i);
        if (e.target.closest('.tp-fzh')) return this.startFreezeDrag(e);
        if (e.target.closest('.tp-hm,input')) return;
        const cell = e.target.closest('.tp-h');
        if (cell && this.canEditLayout()) this.startColumnDrag(e, +cell.dataset.i);
      });
    }

    // Pointer-based dragging; native HTML drag-and-drop hands text to Tableau, which tries to open it as a data file.
    startColumnDrag(e, from) {
      e.preventDefault();
      const x0 = e.clientX;
      const gridLeft = this.grid.getBoundingClientRect().left;
      const mids = [];
      let acc = 0;
      this.view.cols.forEach(c => { mids.push(acc + c.w / 2); acc += c.w; });
      const bounds = [0];
      this.view.cols.forEach(c => bounds.push(bounds[bounds.length - 1] + c.w));
      let guide = null, gap = from;
      const move = ev => {
        if (!guide) {
          if (Math.abs(ev.clientX - x0) < 6) return;
          guide = document.createElement('div');
          guide.className = 'tp-guide';
          this.grid.appendChild(guide);
          this.grid.classList.add('tp-dragging');
        }
        const x = ev.clientX - gridLeft;
        gap = mids.filter(m => m < x).length;
        guide.style.left = bounds[gap] + 'px';
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        if (!guide) return;
        guide.remove();
        this.grid.classList.remove('tp-dragging');
        this.suppressClick = true;
        setTimeout(() => { this.suppressClick = false; }, 0);
        const to = gap > from ? gap - 1 : gap;
        if (to !== from) this.moveColumn(from, to);
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    }

    renderTotalsOnly() {
      const show = this.cfg.totals.show && this.view.cols.some(c => c.isNum);
      const html = show ? this.totalsHtml(this.view.totals, 'Total', 'tp-trow') : '';
      this.totTop.innerHTML = this.cfg.totals.location === 'top' ? html : '';
      this.totBottom.innerHTML = this.cfg.totals.location === 'bottom' ? html : '';
    }

    // Sorting always works; when viewers can't change the layout it lasts only for this session.
    setSort(sort) {
      if (this.canEditLayout()) return this.updateLayout({ sort });
      this.sessionSort = sort;
      this.refresh();
    }

    moveColumn(fromVis, toVis) {
      const cols = this.view.cols;
      if (fromVis === toVis || !cols[fromVis] || !cols[toVis]) return;
      const order = this.orderedKeys();
      const k = cols[fromVis].key;
      const target = cols[toVis].key;
      order.splice(order.indexOf(k), 1);
      const ti = order.indexOf(target);
      order.splice(fromVis < toVis ? ti + 1 : ti, 0, k);
      this.updateLayout({ order });
    }

    startResize(e, i) {
      e.preventDefault();
      const c = this.view.cols[i];
      const x0 = e.clientX, w0 = c.w;
      const move = ev => {
        c.w = Math.max(MIN_W, w0 + ev.clientX - x0);
        this.grid.style.setProperty('--cols', this.template());
        this.grid.style.width = this.view.cols.reduce((s, x) => s + x.w, 0) + 'px';
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        this.updateLayout({ widths: Object.assign({}, this.layout.widths, { [c.key]: c.w }) });
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    }

    startFreezeDrag(e) {
      e.preventDefault();
      e.stopPropagation();
      const gridLeft = this.grid.getBoundingClientRect().left;
      const bounds = [0];
      this.view.cols.forEach(c => bounds.push(bounds[bounds.length - 1] + c.w));
      let n = this.view.nFrozen;
      const guide = document.createElement('div');
      guide.className = 'tp-guide';
      guide.style.left = bounds[n] + 'px';
      this.grid.appendChild(guide);
      const move = ev => {
        const x = ev.clientX - gridLeft;
        let best = 0;
        bounds.forEach((b, k) => { if (Math.abs(b - x) < Math.abs(bounds[best] - x)) best = k; });
        n = best;
        guide.style.left = bounds[best] + 'px';
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        guide.remove();
        this.updateLayout({ freeze: n });
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    }

    // ---------- body interactions ----------

    rowAt(target) {
      const row = target.closest('.tp-row');
      return row ? +row.dataset.r : -1;
    }

    wireBody() {
      this.rowsEl.addEventListener('mousedown', e => {
        if (e.button !== 0) return;
        const r = this.rowAt(e.target);
        const cell = e.target.closest('[data-c]');
        if (r < 0 || !cell) return;
        this.dragStart = { r, c: +cell.dataset.c, moved: false };
        this.range = e.shiftKey && this.range ? { r0: this.range.r0, c0: this.range.c0, r1: r, c1: +cell.dataset.c } : null;
      });
      this.rowsEl.addEventListener('mousemove', e => {
        if (this.dragStart && e.buttons === 1) {
          const r = this.rowAt(e.target);
          const cell = e.target.closest('[data-c]');
          if (r >= 0 && cell && (r !== this.dragStart.r || +cell.dataset.c !== this.dragStart.c)) {
            this.dragStart.moved = true;
            this.range = { r0: this.dragStart.r, c0: this.dragStart.c, r1: r, c1: +cell.dataset.c };
            this.renderBody(true);
          }
        }
        this.hover(e);
      });
      this.rowsEl.addEventListener('mouseup', e => {
        const ds = this.dragStart;
        this.dragStart = null;
        if (!ds || ds.moved || e.shiftKey) { this.renderBody(true); return; }
        this.range = { r0: ds.r, c0: ds.c, r1: ds.r, c1: ds.c };
        this.clickRow(ds.r, e.ctrlKey || e.metaKey, { x: e.clientX, y: e.clientY });
      });
      this.rowsEl.addEventListener('mouseleave', () => this.unhover());
      // Right-click copy: Tableau Desktop keeps Ctrl+C for itself, so this is the reliable way to copy.
      this.rowsEl.addEventListener('contextmenu', e => {
        e.preventDefault();
        const r = this.rowAt(e.target);
        const cell = e.target.closest('[data-c]');
        const rg = this.normRange();
        if (r >= 0 && cell && !(rg && r >= rg.r0 && r <= rg.r1 && +cell.dataset.c >= rg.c0 && +cell.dataset.c <= rg.c1)) {
          this.range = { r0: r, c0: +cell.dataset.c, r1: r, c1: +cell.dataset.c };
          this.renderBody(true);
        }
        const g = this.normRange();
        const n = g ? (g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) : 0;
        const what = n > 1 ? 'Copy ' + n.toLocaleString() + ' cells' : 'Copy cell';
        const x = e.clientX, y = e.clientY;
        const html = '<div class="tp-menu" role="menu">' +
          `<button type="button" data-m="cells">${what}</button>` +
          '<button type="button" data-m="headers">Copy with headers</button><hr>' +
          `<button type="button" data-m="all">Copy whole table (${this.view.idx.length.toLocaleString()} rows)</button></div>`;
        this.openPop({ getBoundingClientRect: () => ({ left: x, bottom: y }) }, html, pop => pop.addEventListener('click', ev => {
          const b = ev.target.closest('button[data-m]');
          if (!b) return;
          this.closePop();
          const text = this.copyText(b.dataset.m);
          if (text) this.copyToClipboard(text, b.dataset.m === 'all' ? this.view.idx.length : null);
        }));
      });
      this.el.addEventListener('contextmenu', e => e.preventDefault());
    }

    normRange() {
      const g = this.range;
      if (!g) return null;
      return { r0: Math.min(g.r0, g.r1), r1: Math.max(g.r0, g.r1), c0: Math.min(g.c0, g.c1), c1: Math.max(g.c0, g.c1) };
    }

    clickRow(r, toggle, point) {
      const d = this.view.disp[r];
      if (typeof d === 'object') {
        if (this.collapsed.has(d.key)) this.collapsed.delete(d.key); else this.collapsed.add(d.key);
        this.computeView();
        this.renderBody(true);
        return;
      }
      const tid = this.store.tuple[d];
      if (tid == null) { this.renderBody(true); return; }
      if (toggle) { if (this.selected.has(tid)) this.selected.delete(tid); else this.selected.add(tid); }
      else if (this.selected.size === 1 && this.selected.has(tid)) this.selected.clear();
      else { this.selected.clear(); this.selected.add(tid); }
      this.renderBody(true);
      const ids = Array.from(this.selected);
      // Same shape as Tableau's samples: a real anchor point for the tooltip, or none at all (an empty object errors).
      const tip = this.cfg.showTooltips && point && ids.length ? { tooltipAnchorPoint: { x: point.x, y: point.y } } : undefined;
      const sel = this.ws.selectTuplesAsync(ids, tableau.SelectOptions.Simple, tip);
      Promise.resolve(sel).catch(() => { /* selection is best effort */ });
    }

    hover(e) {
      if (!this.cfg.showTooltips) return;
      const r = this.rowAt(e.target);
      const d = this.view.disp[r];
      const tid = typeof d === 'number' ? this.store.tuple[d] : null;
      if (tid == null || tid === this.hoverTid) return;
      this.hoverTid = tid;
      clearTimeout(this.hoverT);
      const x = e.clientX, y = e.clientY;
      this.hoverT = setTimeout(() => {
        Promise.resolve(this.ws.hoverTupleAsync(tid, { tooltipAnchorPoint: { x, y } })).catch(() => { /* tooltips are best effort */ });
      }, 120);
    }

    unhover() {
      clearTimeout(this.hoverT);
      if (this.hoverTid == null) return;
      this.hoverTid = null;
      Promise.resolve(this.ws.hoverTupleAsync(null)).catch(() => { /* ignore */ });
    }

    onKey(e) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c' && !e.target.closest('input')) {
        const text = this.copyText();
        if (!text) return;
        e.preventDefault();
        this.copyToClipboard(text);
      } else if (e.key === 'Escape') {
        this.closePop();
      }
    }

    // mode: 'auto' adds headers for multi-cell copies, 'cells' never, 'headers' always, 'all' copies every visible row.
    copyText(mode) {
      const rg = this.normRange();
      const { disp, cols } = this.view;
      const st = this.store;
      let rows = [], c0 = 0, c1 = cols.length - 1;
      if (mode === 'all') {
        // Includes rows inside collapsed subtotal groups.
        const all = [];
        for (const d of disp) {
          if (typeof d === 'number') all.push(d);
          else if (this.collapsed.has(d.key)) all.push(...d.rows);
        }
        rows = all.map(i => ({ i }));
      } else if (rg && !(rg.r0 === rg.r1 && rg.c0 === rg.c1 && this.selected.size > 1)) {
        for (let r = rg.r0; r <= rg.r1; r++) rows.push(r);
        c0 = rg.c0; c1 = rg.c1;
      } else if (this.selected.size) {
        disp.forEach((d, r) => { if (typeof d === 'number' && this.selected.has(st.tuple[d])) rows.push(r); });
      }
      if (!rows.length) return '';
      const lines = [];
      const headers = mode === 'headers' || mode === 'all' || (mode !== 'cells' && (rows.length > 1 || c1 > c0));
      if (headers) lines.push(cols.slice(c0, c1 + 1).map(c => c.label).join('\t'));
      for (const r of rows) {
        const d = typeof r === 'object' ? r.i : disp[r];
        if (typeof d !== 'number') continue;
        const vals = [];
        for (let j = c0; j <= c1; j++) { const v = st.fmt[cols[j].key][d]; vals.push(v == null ? '' : v.replace(/[\t\r\n]+/g, ' ')); }
        lines.push(vals.join('\t'));
      }
      return lines.join('\n');
    }

    // execCommand first: it works inside embedded frames where the newer clipboard API is often blocked.
    copyToClipboard(text, rowCount) {
      const lines = text.split('\n').length;
      const done = () => this.flashOk('Copied ' + (rowCount != null ? rowCount.toLocaleString() + ' rows' : lines > 1 ? lines + ' lines' : 'cell'));
      if (this.copyViaTextarea(text)) return done();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(done, () => this.flash('Copy is blocked here'));
      } else this.flash('Copy is blocked here');
    }

    copyViaTextarea(text) {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0;user-select:text';
      document.body.appendChild(ta);
      ta.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      ta.remove();
      return ok;
    }

    flashOk(msg) {
      this.setStatus(msg, false);
      clearTimeout(this.flashT);
      this.flashT = setTimeout(() => this.setStatus(null), 1500);
    }

    // ---------- popups ----------

    openPop(anchor, html, wire) {
      // A fresh element each time, so listeners from earlier menus can't fire again.
      const fresh = this.pop.cloneNode(false);
      this.pop.replaceWith(fresh);
      this.pop = fresh;
      this.pop.innerHTML = html;
      this.pop.hidden = false;
      const r = anchor.getBoundingClientRect();
      const pw = this.pop.offsetWidth, ph = this.pop.offsetHeight;
      this.pop.style.left = Math.max(4, Math.min(r.left, window.innerWidth - pw - 4)) + 'px';
      this.pop.style.top = Math.max(4, Math.min(r.bottom + 2, window.innerHeight - ph - 4)) + 'px';
      wire(this.pop);
      const first = this.pop.querySelector('button,input');
      if (first) first.focus();
    }

    closePop() { this.pop.hidden = true; this.pop.innerHTML = ''; }

    openColumnMenu(i, anchor) {
      const c = this.view.cols[i];
      const lay = this.layout;
      const edit = this.canEditLayout();
      const last = this.view.cols.length - 1;
      const words = c.isNum ? ['smallest to largest', 'largest to smallest'] : (c.isDate || this.store.dates[c.key]) ? ['oldest to newest', 'newest to oldest'] : ['A to Z', 'Z to A'];
      const items = [
        ['sort-asc', 'Sort ' + words[0]], ['sort-desc', 'Sort ' + words[1]],
        lay.sort && lay.sort.key === c.key ? ['sort-clear', 'Clear sort'] : null
      ];
      if (edit) {
        items.push('-',
          i > 0 ? ['first', 'Move to first'] : null, i > 0 ? ['left', 'Move left'] : null,
          i < last ? ['right', 'Move right'] : null, i < last ? ['last', 'Move to last'] : null,
          ['hide', 'Hide column'], '-',
          ['freeze', 'Freeze up to here'], this.view.nFrozen ? ['unfreeze', 'Unfreeze columns'] : null,
          !c.isNum ? (lay.groupBy === c.key ? ['ungroup', 'Remove subtotals'] : ['group', 'Subtotal by this column']) : null,
          lay.groupBy === c.key ? ['collapse-all', 'Collapse all'] : null,
          lay.groupBy === c.key ? ['expand-all', 'Expand all'] : null, '-',
          ['rename', 'Rename…'], lay.labels && lay.labels[c.key] ? ['unrename', 'Reset name'] : null);
      }
      if (c.isNum && c.total) items.push('-', ['info', 'Total: ' + c.total.text + (c.total.problem ? ' (' + c.total.problem + ')' : ''), true]);
      if (this.authoring) items.push(['format', 'Format table…']);
      const html = '<div class="tp-menu" role="menu">' + items.filter(Boolean).map(it => it === '-' ? '<hr>' :
        `<button type="button" role="menuitem" data-a="${it[0]}"${it[2] ? ' disabled' : ''}>${esc(it[1])}</button>`).join('') + '</div>';
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        const b = e.target.closest('button[data-a]');
        if (!b) return;
        this.closePop();
        this.menuAction(b.dataset.a, i, c);
      }));
    }

    menuAction(a, i, c) {
      const cols = this.view.cols;
      const lay = this.layout;
      switch (a) {
        case 'sort-asc': return this.setSort({ key: c.key, dir: 'asc' });
        case 'sort-desc': return this.setSort({ key: c.key, dir: 'desc' });
        case 'sort-clear': return this.setSort(null);
        case 'first': return this.moveColumn(i, 0);
        case 'left': return this.moveColumn(i, i - 1);
        case 'right': return this.moveColumn(i, i + 1);
        case 'last': return this.moveColumn(i, cols.length - 1);
        case 'hide': return this.updateLayout({ hidden: (lay.hidden || []).concat(c.key) });
        case 'freeze': return this.updateLayout({ freeze: i + 1 });
        case 'unfreeze': return this.updateLayout({ freeze: 0 });
        case 'group': return this.updateLayout({ groupBy: c.key });
        case 'ungroup': return this.updateLayout({ groupBy: null });
        case 'collapse-all':
        case 'expand-all':
          this.collapsed = a === 'collapse-all' ? new Set((this.groups || []).map(g => g.key)) : new Set();
          this.computeView();
          this.renderBody(true);
          return;
        case 'rename': {
          const name = window.prompt('Column name', c.label);
          if (name == null) return;
          return this.updateLayout({ labels: Object.assign({}, lay.labels, { [c.key]: name.trim() }) });
        }
        case 'unrename': {
          const labels = Object.assign({}, lay.labels);
          delete labels[c.key];
          if (!this.authoring) {
            const base = this.cfg.layout.labels || {};
            if (base[c.key]) labels[c.key] = '';
          }
          return this.updateLayout({ labels });
        }
        case 'format': return this.openFormat();
        default:
      }
    }

    openColumnsPicker(anchor) {
      const hidden = new Set(this.layout.hidden || []);
      const keys = this.orderedKeys();
      const html = '<div class="tp-menu tp-picker"><div class="tp-ph">Show and hide columns</div>' + keys.map(k => {
        const c = this.colDef(k);
        const auto = this.autoHidden(c);
        return `<label${auto ? ' class="muted" title="Hidden because its parameter is set to none"' : ''}><input type="checkbox" data-k="${esc(k)}"${!hidden.has(k) ? ' checked' : ''}${auto ? ' disabled' : ''}> ${esc(this.labelFor(c))}</label>`;
      }).join('') + '<hr><button type="button" data-a="all">Show all</button></div>';
      this.openPop(anchor, html, pop => {
        pop.addEventListener('change', e => {
          const k = e.target.dataset.k;
          if (!k) return;
          const h = new Set(this.layout.hidden || []);
          if (e.target.checked) h.delete(k); else h.add(k);
          this.updateLayout({ hidden: Array.from(h) });
        });
        pop.addEventListener('click', e => {
          if (e.target.dataset.a === 'all') { this.updateLayout({ hidden: [] }); this.closePop(); }
        });
      });
    }

    // ---------- format dialog ----------

    openFormat() {
      if (!this.store) return;
      const payload = JSON.stringify({
        config: this.cfg,
        sheet: this.ws.name,
        params: this.params.map(p => p.name),
        columns: this.store.columns.map(c => ({ key: c.key, name: c.name, isNum: c.isNum, dataType: c.dataType }))
      });
      const url = new URL('format.html' + window.location.search, window.location.href).href;
      return tableau.extensions.ui.displayDialogAsync(url, payload, { width: 860, height: 680 })
        .then(() => { this.loadConfig(); this.refresh(); })
        .catch(err => {
          if (!(tableau.ErrorCodes && err.errorCode === tableau.ErrorCodes.DialogClosedByUser)) this.flash(err.message);
        });
    }

    // ---------- export ----------

    onDownloadClick(anchor) {
      const f = this.cfg.export.formats;
      if (f === 'csv') return this.download('csv');
      if (f === 'excel') return this.download('xlsx');
      this.openPop(anchor, '<div class="tp-menu"><button type="button" data-f="xlsx">Excel (.xlsx)</button><button type="button" data-f="csv">CSV (.csv)</button></div>', pop =>
        pop.addEventListener('click', e => { const b = e.target.closest('button[data-f]'); if (b) { this.closePop(); this.download(b.dataset.f); } }));
    }

    snapshot(includeTotals) {
      const { cols, disp, totals, nFrozen } = this.view;
      const st = this.store;
      const value = (sc, i) => {
        const c = this.colDef(sc.key);
        if (c.isNum) { const v = st.raw[c.key][i]; return v === v ? { kind: 'num', v } : { kind: 'null' }; }
        if (c.isDate) { const v = st.raw[c.key][i]; return v == null ? { kind: 'null' } : { kind: 'date', v, time: c.dataType === 'date-time' }; }
        const d = st.dates[c.key];
        if (d) { const v = d[i]; return v == null ? { kind: 'null' } : { kind: 'date', v, time: /hh/.test(c.excelFmt || '') }; }
        const f = st.fmt[c.key][i];
        return f == null ? { kind: 'null' } : { kind: 'text', v: f };
      };
      const groupKey = this.layout.groupBy;
      // Every row is exported whatever is collapsed on screen; collapsed groups carry a flag for Excel's outline.
      let rows;
      if (this.groups) {
        rows = [];
        for (const g of this.groups) {
          const collapsed = this.collapsed.has(g.key);
          rows.push({ type: 'group', label: g.label + ' subtotal', totals: Object.assign({ __label: groupKey }, g.totals), collapsed });
          for (const i of g.rows) rows.push({ type: 'data', i, level: 1, hidden: collapsed });
        }
      } else {
        rows = this.view.idx.map(i => ({ type: 'data', i }));
      }
      return {
        sheetName: this.ws.name,
        columns: cols.map(c => ({ key: c.key, label: c.label, spec: c.spec, excelFmt: c.excelFmt, width: c.w, isNum: c.isNum })),
        rows,
        value,
        totals,
        totalsLabelKey: totals.__label,
        includeTotals: includeTotals && this.cfg.totals.show,
        includeSubtotals: this.cfg.export.subtotals,
        totalsLocation: this.cfg.totals.location,
        frozenCols: nFrozen
      };
    }

    async download(kind) {
      if (!this.store || !this.store.complete) { this.flash('Wait for all rows to load, then download.'); return; }
      const t0 = performance.now();
      const params = {};
      this.params.forEach(p => { params[p.name] = p.value; });
      const name = TPExport.buildFileName(this.cfg.export.filename, { sheet: this.ws.name, params }, kind);
      try {
        this.setStatus('Preparing ' + (kind === 'csv' ? 'CSV' : 'Excel') + '…');
        await new Promise(r => setTimeout(r, 0));
        let blob;
        if (kind === 'csv') {
          blob = new Blob([TPExport.buildCsv(this.snapshot(false))], { type: 'text/csv;charset=utf-8' });
        } else {
          await loadScript('lib/exceljs.min.js?v=4.4.0', () => window.ExcelJS);
          const buf = await TPExport.buildXlsx(this.snapshot(this.cfg.export.excelTotals), window.ExcelJS);
          blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        }
        saveBlob(blob, name);
        this.flashOk('Downloaded ' + name + ' (' + ((performance.now() - t0) / 1000).toFixed(1) + 's)');
      } catch (e) {
        this.flash('Download failed: ' + ((e && e.message) || e));
      }
    }
  }

  function loadScript(src, ready) {
    if (ready()) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Could not load ' + src));
      document.head.appendChild(s);
    });
  }

  function saveBlob(blob, name) {
    if (window.__tpCapture) { window.__tpCapture(blob, name); return; }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  const ICONS = {
    dl: '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M8 1a.75.75 0 0 1 .75.75v6.69l2.22-2.22a.75.75 0 1 1 1.06 1.06l-3.5 3.5a.75.75 0 0 1-1.06 0l-3.5-3.5a.75.75 0 1 1 1.06-1.06l2.22 2.22V1.75A.75.75 0 0 1 8 1ZM2.75 12a.75.75 0 0 1 .75.75v.75h9v-.75a.75.75 0 0 1 1.5 0v1.5a.75.75 0 0 1-.75.75H2.75a.75.75 0 0 1-.75-.75v-1.5a.75.75 0 0 1 .75-.75Z"/></svg>',
    cols: '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M1.75 2h12.5a.75.75 0 0 1 .75.75v10.5a.75.75 0 0 1-.75.75H1.75a.75.75 0 0 1-.75-.75V2.75A.75.75 0 0 1 1.75 2Zm.75 1.5v9h3v-9Zm4.5 0v9h2v-9Zm3.5 0v9h3v-9Z"/></svg>',
    gear: '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M8 5.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5ZM7 1h2l.4 1.9 1.3.6 1.7-1 1.4 1.4-1 1.7.6 1.3L15 7v2l-1.9.4-.6 1.3 1 1.7-1.4 1.4-1.7-1-1.3.6L9 15H7l-.4-1.9-1.3-.6-1.7 1-1.4-1.4 1-1.7-.6-1.3L1 9V7l1.9-.4.6-1.3-1-1.7 1.4-1.4 1.7 1 1.3-.6Z"/></svg>',
    info: '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M8 1a7 7 0 1 1 0 14A7 7 0 0 1 8 1Zm0 1.5a5.5 5.5 0 1 0 0 11 5.5 5.5 0 0 0 0-11ZM8.75 7v4.25h-1.5V7ZM8 4.25a.9.9 0 1 1 0 1.8.9.9 0 0 1 0-1.8Z"/></svg>'
  };

  window.TPDefaults = defaults;
  window.TPMergeConfig = mergeConfig;
  if (document.getElementById('tp-root')) {
    const app = new TableApp(document.getElementById('tp-root'));
    window.__tpApp = app;
    app.start().catch(e => app.setStatus('Could not start: ' + ((e && e.message) || e), true));
  }
})();
