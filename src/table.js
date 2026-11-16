/* global tableau, TPRules, TPData, TPExport, ExcelJS */
(function () {
  'use strict';

  const R = window.TPRules;
  const SETTINGS_KEY = 'tp-config';
  const OVERSCAN = 12;
  const MIN_W = 48;
  const FREEZE_MAX = 0.6;
  const MAX_LEVELS = 3;
  const NUL = '\u0000null';
  // Semantic version (MAJOR.MINOR.PATCH); see CHANGELOG.md. The ?v= tags in the HTML are only for browser caches.
  const VERSION = '1.3.1';
  const calcKey = d => d.kind === 'rank' ? '__rank:' + (d.dir || 'desc') + ':' + d.src : d.kind === 'run' ? '__run:' + d.src : '__' + d.kind + ':' + d.src + '|' + d.other;

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
      // Catch-all values such as "Other" that sort after everything else in their column in the default order.
      lastRows: { values: '', always: false },
      links: {},
      cf: {},
      colGroups: {},
      colFormat: {},
      // Per column: hide it while a parameter has one of these values, e.g. { param: 'Show Details', values: 'False' }.
      hideWhen: {},
      // Whole-table look. The defaults are the original design, so nothing changes until an author sets something.
      style: { size: 'normal', density: 'normal', text: '#1f2328', headBg: '#f6f8fa', headText: '#1f2328', headBold: true, headWrap: false, lines: 'both', banding: true, totBg: '#fbfbfc', totBold: true, subBg: '#f3f5f8', subBold: true, neg: '', accent: '#1f6feb' },
      layout: { order: [], hidden: [], widths: {}, labels: {}, freeze: 'auto', sort: null, groupBy: null, pct: [], calc: [], cfOff: [], hideGroups: false, topN: null }
    };
  }

  function mergeConfig(saved) {
    const d = defaults();
    if (!saved) return d;
    const out = Object.assign(d, saved);
    out.nullText = Object.assign(defaults().nullText, saved.nullText);
    out.totals = Object.assign(defaults().totals, saved.totals);
    out.export = Object.assign(defaults().export, saved.export);
    out.lastRows = Object.assign(defaults().lastRows, saved.lastRows);
    out.layout = Object.assign(defaults().layout, saved.layout);
    out.headerLinks = saved.headerLinks || {};
    out.totalOverrides = saved.totalOverrides || {};
    out.links = saved.links || {};
    out.cf = saved.cf || {};
    out.colGroups = saved.colGroups || {};
    out.colFormat = saved.colFormat || {};
    out.hideWhen = saved.hideWhen || {};
    out.style = Object.assign(defaults().style, saved.style);
    return out;
  }

  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

  // Colorblind-safe good/warning/bad colors. Every pair stays distinct under protan, deutan and tritan simulation,
  // dark text on each tint is above 9:1 contrast, and each text color is above 5:1 on white.
  const CF = {
    bg: { bad: '#F6CDB6', warn: '#FBE7A8', good: '#BFE6D8' },
    text: { bad: '#9E2A2B', warn: '#8A6100', good: '#00704F' },
    sym: { bad: '✕', warn: '!', good: '✓' },
    word: { bad: 'Below target', warn: 'Watch', good: 'On target' },
    bar: '#A9CDEB'
  };

  function mix(a, b, t) {
    const p = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));
    const x = p(a), y = p(b);
    t = Math.max(0, Math.min(1, t));
    return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
  }

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
      this.otherOpen = new Set();
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
      await this.loadSourceNames();
      this.reload();
    }

    loadConfig() {
      let saved = null;
      try { saved = JSON.parse(tableau.extensions.settings.get(SETTINGS_KEY) || 'null'); } catch (e) { saved = null; }
      this.cfg = mergeConfig(saved);
      this.applyStyle();
    }

    // ---------- style ----------

    applyStyle() {
      const s = this.cfg.style;
      const set = (el, k, v) => el.style.setProperty(k, v);
      this.rowH = ({ compact: 22, normal: 26, roomy: 32 }[s.density] || 26) + (s.size === 'large' ? 2 : 0);
      set(this.el, '--tp-row-h', this.rowH + 'px');
      set(this.el, '--tp-text', s.text);
      set(this.el, '--tp-accent', s.accent);
      set(this.el, '--tp-sel', mix(s.accent, '#FFFFFF', 0.84));
      this.el.style.fontSize = ({ small: 11, normal: 12, large: 14 }[s.size] || 12) + 'px';
      // Header and total colors apply inside the table only, so the toolbar keeps its look.
      set(this.scroller, '--tp-head-bg', s.headBg);
      set(this.scroller, '--tp-head-text', s.headText);
      set(this.scroller, '--tp-tot-bg', s.totBg);
      set(this.scroller, '--tp-grp-bg', s.subBg);
      set(this.scroller, '--tp-grp-bg1', mix(s.subBg, '#FFFFFF', 0.4));
      set(this.scroller, '--tp-grp-bg2', mix(s.subBg, '#FFFFFF', 0.65));
      const cl = this.scroller.classList;
      ['tp-lines-h', 'tp-lines-v', 'tp-lines-none'].forEach(c => cl.remove(c));
      if (s.lines === 'horizontal') cl.add('tp-lines-h'); else if (s.lines === 'vertical') cl.add('tp-lines-v'); else if (s.lines === 'none') cl.add('tp-lines-none');
      cl.toggle('tp-noband', !s.banding);
      cl.toggle('tp-hwrap', !!s.headWrap);
      cl.toggle('tp-hnb', !s.headBold);
      cl.toggle('tp-totnb', !s.totBold);
      cl.toggle('tp-subnb', !s.subBold);
    }

    // Effective number format: Tableau's own, then the table-wide negative style, then the column's overrides.
    effSpec(c) {
      const o = (!c.virtual && this.cfg.colFormat[c.key]) || {};
      const neg = o.neg || this.cfg.style.neg;
      const dec = o.dec === undefined || o.dec === '' ? null : Number(o.dec);
      if (!neg && dec == null && !o.units && !o.kind) return null;
      const base = c.virtual ? c.ownSpec : c.baseSpec !== undefined ? c.baseSpec : c.spec;
      const s = Object.assign({ cur: '', pct: false, dec: 2, group: true, scale: '', neg: 'minus' }, base || {});
      if (o.kind === 'number') { s.cur = ''; s.pct = false; }
      else if (o.kind === 'currency') { s.cur = (base && base.cur) || '$'; s.pct = false; s.group = true; }
      else if (o.kind === 'percent') { s.pct = true; s.cur = ''; }
      if (o.units === 'none') s.scale = '';
      else if (o.units === 'K' || o.units === 'M' || o.units === 'B') s.scale = o.units;
      else if (o.units === 'auto') {
        let m = 0;
        for (const v of this.store.raw[c.key]) { const a = Math.abs(v); if (a > m) m = a; }
        s.scale = s.pct ? '' : m >= 1e9 ? 'B' : m >= 1e6 ? 'M' : m >= 1e3 ? 'K' : '';
      }
      if (dec != null) s.dec = dec;
      if (neg) { s.neg = neg === 'parens' || neg === 'redparens' ? 'paren' : 'minus'; s.red = neg === 'red' || neg === 'redparens'; }
      return s;
    }

    excelFor(spec) {
      const f = R.excelNumFmt(spec);
      if (!f || !spec.red) return f;
      const parts = f.split(';');
      parts[1] = '[Red]' + parts[1];
      return parts.join(';');
    }

    // Applies number overrides once all rows are in (rows still arriving keep Tableau's format for a moment),
    // keeping Tableau's own strings so an override can be undone. Also sets each column's alignment and bold classes.
    applyFormats() {
      const st = this.store;
      st.fmtBase = st.fmtBase || {};
      for (const c of st.columns) {
        const o = (!c.virtual && this.cfg.colFormat[c.key]) || {};
        c.cls = (o.align ? ' al-' + o.align : '') + (o.bold ? ' b' : '');
        if (!c.isNum) continue;
        if (c.virtual) {
          // Differences and running totals are in the source column's units, so they share its format.
          const src = (c.vkind === 'run' || c.vkind === 'diff') && this.colDef(c.srcKey);
          if (src) { c.spec = src.spec; c.excelFmt = src.excelFmt; continue; }
          const eff = this.effSpec(c);
          c.spec = eff || c.ownSpec;
          c.excelFmt = eff ? this.excelFor(eff) : c.ownExcel;
          continue;
        }
        if (!st.complete) continue;
        const eff = this.effSpec(c);
        const sig = eff ? JSON.stringify(eff) : '';
        if (sig === (c.fmtSig || '')) continue;
        if (c.baseSpec === undefined) { c.baseSpec = c.spec; c.baseExcel = c.excelFmt; st.fmtBase[c.key] = st.fmt[c.key]; }
        if (eff) {
          const raw = st.raw[c.key];
          st.fmt[c.key] = raw.map(v => v === v && v != null ? R.formatNumber(v, eff) : null);
          c.spec = eff;
          c.excelFmt = this.excelFor(eff);
        } else {
          st.fmt[c.key] = st.fmtBase[c.key];
          c.spec = c.baseSpec;
          c.excelFmt = c.baseExcel;
        }
        c.fmtSig = sig;
      }
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
        this.params = ps.map(p => ({ name: p.name, value: p.currentValue ? String(p.currentValue.formattedValue) : '', raw: p.currentValue ? p.currentValue.value : null }));
        ps.forEach(p => p.addEventListener(tableau.TableauEventType.ParameterChanged, () => this.onParamChanged()));
      } catch (e) {
        this.params = [];
      }
    }

    async onParamChanged() {
      try {
        const ps = await this.ws.getParametersAsync();
        this.params = ps.map(p => ({ name: p.name, value: p.currentValue ? String(p.currentValue.formattedValue) : '', raw: p.currentValue ? p.currentValue.value : null }));
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
        if (first) { this.loadLocalLayout(); this.loadPresets(); }
        this.refresh();
        if (done) this.setStatus(null);
      }).catch(e => {
        if (gen === this.gen) this.setStatus('Could not load data: ' + ((e && e.message) || e), true);
      });
    }

    // ---------- layout ----------

    // Viewer layouts belong to a worksheet and its data sources, so adding or removing a field on the sheet keeps
    // them (missing fields are skipped, new ones show at the end). If the data source names can't be read, the
    // older key based on the exact field list is used.
    async loadSourceNames() {
      try {
        const ds = await this.ws.getDataSourcesAsync();
        this.dsSig = ds.map(d => d.name).sort().join('|') || null;
      } catch (e) { this.dsSig = null; }
    }

    layoutKey() {
      return 'tp-layout:' + this.ws.name + ':' + (this.keyBase || (this.dsSig ? 'ds' + hashStr(this.dsSig) : this.legacyHash()));
    }

    // Different workbooks can share a sheet name and generic data source names ("Extract"), and all of them share
    // this browser storage. Each slot remembers its field list; a table whose fields mostly differ from the slot's
    // gets a slot of its own, so one workbook's layouts never show up in another. Small field changes keep the slot.
    resolveKey() {
      const fields = this.store.columns.filter(c => !c.virtual).map(c => c.key);
      const legacy = this.legacyHash();
      let base = this.dsSig ? 'ds' + hashStr(this.dsSig) : legacy;
      if (this.dsSig) {
        const metaKey = k => 'tp-fields:' + this.ws.name + ':' + k;
        try {
          const saved = JSON.parse(localStorage.getItem(metaKey(base)) || 'null');
          if (Array.isArray(saved) && fieldOverlap(saved, fields) < 0.5) base += '-' + legacy;
          localStorage.setItem(metaKey(base), JSON.stringify(fields));
        } catch (e) { /* storage blocked */ }
      }
      this.keyBase = base;
    }

    legacyHash() { return hashStr(this.store ? this.store.columns.filter(c => !c.virtual).map(c => c.key).sort().join('|') : ''); }

    // Moves layouts saved under the older field-list key to the current key, once.
    migrateKey(prefix) {
      const key = this.layoutKey().replace('tp-layout:', prefix);
      const old = prefix + this.ws.name + ':' + this.legacyHash();
      if (key === old) return;
      try {
        if (localStorage.getItem(key) == null && localStorage.getItem(old) != null) localStorage.setItem(key, localStorage.getItem(old));
      } catch (e) { /* storage blocked */ }
    }

    loadLocalLayout() {
      this.local = null;
      if (this.authoring || !this.cfg.allowViewerLayout) return;
      this.resolveKey();
      this.migrateKey('tp-layout:');
      this.migrateKey('tp-presets:');
      try {
        const l = cleanLayout(JSON.parse(localStorage.getItem(this.layoutKey()) || 'null'));
        this.local = Object.keys(l).length ? l : null;
      } catch (e) { this.local = null; }
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
      if (this.presets && this.presets.active) { this.presets.active = null; this.savePresets(); }
      try { localStorage.removeItem(this.layoutKey()); } catch (e) { /* storage blocked */ }
      this.refresh();
    }

    orderedKeys() {
      const all = this.store.columns.map(c => c.key);
      const out = (this.layout.order || []).filter(k => all.includes(k));
      for (const c of this.store.columns) {
        if (out.includes(c.key)) continue;
        // A new % of total column goes right after the column it is based on.
        let at = c.virtual ? out.indexOf(c.srcKey) : -1;
        while (at >= 0 && at + 1 < out.length && (this.colDef(out[at + 1]) || {}).srcKey === c.srcKey) at++;
        if (at >= 0) out.splice(at + 1, 0, c.key); else out.push(c.key);
      }
      return out;
    }

    colDef(key) { return this.store.columns.find(c => c.key === key); }

    // Sorts are a list of { key, dir }, first one wins; older layouts stored a single object.
    sortList() { const s = this.layout.sort; return !s ? [] : Array.isArray(s) ? s : s.key ? [s] : []; }

    // Subtotal levels, outermost first; older layouts stored a single column key.
    groupKeys() { const g = this.layout.groupBy; return !g ? [] : Array.isArray(g) ? g.slice(0, MAX_LEVELS) : [g]; }

    // ---------- % of total columns ----------

    // Table calculation columns: % of total (layout.pct), plus rank, running total, difference and
    // % difference (layout.calc). Each sits right after the column it is based on.
    virtualDefs() {
      const lay = this.layout;
      const num = k => { const c = k && this.colDef(k); return !!(c && c.isNum && !c.virtual); };
      const out = [];
      for (const k of lay.pct || []) if (num(k)) out.push({ kind: 'pct', src: k, key: '__pct:' + k });
      for (const d of lay.calc || []) {
        if (!num(d.src) || ((d.kind === 'diff' || d.kind === 'pdiff') && !num(d.other))) continue;
        out.push(Object.assign({ key: calcKey(d) }, d));
      }
      return out;
    }

    syncVirtual() {
      const st = this.store;
      const defs = this.virtualDefs();
      const want = new Set(defs.map(d => d.key));
      st.columns = st.columns.filter(c => !c.virtual || want.has(c.key));
      for (const d of defs) {
        if (st.columns.some(c => c.key === d.key)) continue;
        const src = this.colDef(d.src);
        const spec = d.kind === 'pct' ? { cur: '', pct: true, dec: 2, group: false, scale: '', neg: 'minus' }
          : d.kind === 'pdiff' ? { cur: '', pct: true, dec: 1, group: false, scale: '', neg: 'minus' }
          : d.kind === 'rank' ? { cur: '', pct: false, dec: 0, group: true, scale: '', neg: 'minus' } : src.spec;
        const excelFmt = d.kind === 'pct' ? '0.00%;-0.00%' : d.kind === 'pdiff' ? '0.0%;-0.0%' : d.kind === 'rank' ? '#,##0' : src.excelFmt;
        const col = { key: d.key, name: '', virtual: true, vkind: d.kind, srcKey: d.src, otherKey: d.other, dir: d.dir, isNum: true, dataType: 'float', spec, excelFmt, ownSpec: spec, ownExcel: excelFmt };
        st.columns.push(col);
        col.name = this.labelFor(col);
        st.raw[d.key] = new Array(st.n).fill(NaN);
        st.fmt[d.key] = new Array(st.n).fill(null);
      }
    }

    // Shares are of the filtered total, so they always add up to 100% of what is on screen. Ranks cover the
    // filtered rows. Running totals are filled later, once the final row order is known.
    fillVirtual(idx) {
      const st = this.store;
      for (const c of st.columns) {
        if (!c.virtual || c.vkind === 'run') continue;
        const src = st.raw[c.srcKey], raw = st.raw[c.key], fmt = st.fmt[c.key];
        if (raw.length !== st.n) { raw.length = st.n; fmt.length = st.n; }
        raw.fill(NaN);
        if (c.vkind === 'rank') {
          const dir = c.dir === 'asc' ? 1 : -1;
          // A zero usually means no activity (a cost-per of $0 when nothing came in), so zeros always rank last,
          // tied with each other, whichever direction is best. Blanks get no rank.
          const rows = idx.filter(i => src[i] === src[i]).sort((a, b) => {
            const za = src[a] === 0, zb = src[b] === 0;
            return za !== zb ? (za ? 1 : -1) : dir * (src[a] - src[b]);
          });
          rows.forEach((i, k) => { raw[i] = k && src[i] === src[rows[k - 1]] ? raw[rows[k - 1]] : k + 1; });
        } else if (c.vkind === 'diff' || c.vkind === 'pdiff') {
          const oth = st.raw[c.otherKey];
          for (let i = 0; i < st.n; i++) {
            const a = src[i], b = oth[i];
            if (a !== a || b !== b) continue;
            raw[i] = c.vkind === 'diff' ? a - b : b ? (a - b) / b : NaN;
          }
        } else {
          let total = 0;
          for (const i of idx) { const v = src[i]; if (v === v) total += v; }
          for (let i = 0; i < st.n; i++) { const v = src[i]; raw[i] = v === v && total ? v / total : NaN; }
        }
        for (let i = 0; i < st.n; i++) fmt[i] = raw[i] === raw[i] ? R.formatNumber(raw[i], c.spec) : null;
      }
    }

    // Running totals follow the rows in the order shown (inside subtotal groups, in group order).
    fillRunning(cols, idx) {
      const st = this.store;
      const runs = cols.filter(c => c.vkind === 'run');
      if (!runs.length) return;
      const order = [];
      if (this.groups) {
        const walk = list => { for (const g of list) { if (g.children) walk(g.children); else order.push(...g.rows); } };
        walk(this.groups);
      } else order.push(...idx);
      for (const c of runs) {
        const src = st.raw[c.srcKey], raw = st.raw[c.key], fmt = st.fmt[c.key];
        raw.fill(NaN); fmt.fill(null);
        let acc = 0;
        for (const i of order) { const v = src[i]; if (v === v) acc += v; raw[i] = acc; fmt[i] = R.formatNumber(acc, c.spec); }
      }
    }

    // ---------- links ----------

    safeUrl(s) {
      if (!s) return null;
      try {
        const u = new URL(String(s).trim());
        return u.protocol === 'http:' || u.protocol === 'https:' ? u.href : null;
      } catch (e) { return null; }
    }

    linkFor(c) {
      const l = this.cfg.links[c.key] || {};
      if (c.isUrl) return l.mode === 'off' ? null : { urlKey: c.key, label: l.label || '' };
      if (l.linkTo && this.colDef(l.linkTo)) return { urlKey: l.linkTo, label: '' };
      return null;
    }

    paramFor(col) {
      const link = this.cfg.headerLinks[col.key];
      if (link === '') return null;
      const name = link || col.name;
      return this.params.find(p => p.name.toLowerCase() === String(name).toLowerCase()) || null;
    }

    labelFor(col) {
      const l = this.layout.labels && this.layout.labels[col.key];
      if (l) return l;
      if (col.virtual) {
        const src = this.colDef(col.srcKey), oth = col.otherKey && this.colDef(col.otherKey);
        const s = src ? this.labelFor(src) : col.srcKey, o = oth ? this.labelFor(oth) : '';
        return { pct: s + ' % of total', rank: s + ' rank', run: s + ' running total', diff: s + ' − ' + o, pdiff: s + ' vs ' + o + ' %' }[col.vkind || 'pct'];
      }
      const p = this.paramFor(col);
      return p ? p.value : col.name;
    }

    // A column hides itself when its header parameter is set to "(none)", or when its "Hide when" rule matches the
    // chosen parameter's current value. Table calculation columns follow the column they're based on.
    autoHidden(col) {
      if (col.virtual) { const src = this.colDef(col.srcKey); return !!src && this.autoHidden(src); }
      const rule = this.cfg.hideWhen && this.cfg.hideWhen[col.key];
      if (rule && rule.param && rule.values) {
        const p = this.params.find(x => x.name.toLowerCase() === String(rule.param).toLowerCase());
        if (p) {
          const cur = [String(p.value).trim().toLowerCase(), String(p.raw).trim().toLowerCase()];
          if (String(rule.values).split(',').map(v => v.trim().toLowerCase()).filter(Boolean).some(v => cur.includes(v))) return true;
        }
      }
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
      this.syncVirtual();
      this.applyFormats();
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
      this.fillVirtual(idx);
      this.sortRows(idx, this.sortList(), cols);
      const totals = this.computeTotals(cols, idx);
      this.computeCfStats(cols, idx);
      this.viewCols = cols;
      const disp = this.groupRows(cols, idx, this.groupKeys());
      this.fillRunning(cols, idx);
      this.view = { cols, idx, disp, totals, nFrozen, spans: this.groupSpans(cols) };
    }

    // ---------- column groups ----------

    // Membership is set by the author and never changes when columns move; the header spans each run of
    // adjacent members, so a member moved away shows the label again in its new spot.
    groupOf(c) { return (c.virtual ? this.cfg.colGroups[c.srcKey] : this.cfg.colGroups[c.key]) || ''; }

    groupNames() {
      const seen = [];
      for (const k of this.orderedKeys()) {
        const c = this.colDef(k);
        const g = c && this.groupOf(c);
        if (g && !seen.includes(g)) seen.push(g);
      }
      return seen;
    }

    groupSpans(cols) {
      if (this.layout.hideGroups || !cols.some(c => this.groupOf(c))) return null;
      const spans = [];
      cols.forEach((c, i) => {
        const g = this.groupOf(c), last = spans[spans.length - 1];
        if (last && last.label === g && last.frozen === c.frozen) last.end = i;
        else spans.push({ label: g, start: i, end: i, frozen: c.frozen });
      });
      return spans;
    }

    // The full run of adjacent columns in the same group (a header split by the freeze line is one group).
    groupRun(start, end) {
      const cols = this.view.cols;
      const g = this.groupOf(cols[start]);
      while (start > 0 && this.groupOf(cols[start - 1]) === g) start--;
      while (end < cols.length - 1 && this.groupOf(cols[end + 1]) === g) end++;
      return { start, end };
    }

    setGroups(change) {
      if (!this.authoring) return;
      const g = Object.assign({}, this.cfg.colGroups);
      change(g);
      Object.keys(g).forEach(k => { if (!g[k]) delete g[k]; });
      this.cfg.colGroups = g;
      this.saveConfig();
      this.refresh();
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
        const f = !c.virtual && this.filters[c.key];
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
        if (f.keep || f.exclude) {
          const fmt = st.fmt[c.key], raw = st.raw[c.key];
          const v = c.isNum ? i => (raw[i] === raw[i] && raw[i] != null ? raw[i] : NUL) : i => fmt[i] == null ? NUL : fmt[i];
          if (f.keep) { const s = new Set(f.keep); tests.push(i => s.has(v(i))); }
          if (f.exclude) { const s = new Set(f.exclude); tests.push(i => !s.has(v(i))); }
        }
      }
      if (this.search) {
        const q = this.search.toLowerCase();
        const fmts = cols.filter(c => !c.virtual).map(c => st.fmt[c.key]);
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
      for (const c of cols) { if (c.isNum) break; keys.push(Object.assign(this.sortKeyArray(c), { last: this.lastFlags(c) })); }
      if (!keys.length) return;
      // "Shown last" values sort after the other values of their own column, so a city called "Texas Other"
      // ends up last among the Texas cities rather than at the bottom of the table.
      idx.sort((a, b) => {
        for (const { arr, num, last } of keys) {
          if (last && last[a] !== last[b]) return last[a] ? 1 : -1;
          const x = arr[a], y = arr[b];
          const xn = x == null || x !== x, yn = y == null || y !== y;
          if (xn || yn) { if (xn !== yn) return xn ? 1 : -1; continue; }
          const d = num ? (x < y ? -1 : x > y ? 1 : 0) : collator.compare(x, y);
          if (d) return d;
        }
        return a - b;
      });
    }

    // A test for the "shown last" values as whole words, ignoring case: "Other" matches "Other" and
    // "Texas Other" but not "Mother". Null when the setting is empty.
    lastMatcher() {
      const vals = String((this.cfg.lastRows && this.cfg.lastRows.values) || '').split(',').map(v => v.trim()).filter(Boolean);
      if (!vals.length) return null;
      const sig = vals.join('\u0001');
      if (this.lastSig !== sig) {
        const words = vals.map(v => v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');
        this.lastRe = new RegExp('(?:^|[^\\p{L}\\p{N}])(?:' + words + ')(?=$|[^\\p{L}\\p{N}])', 'iu');
        this.lastSig = sig;
      }
      const re = this.lastRe;
      return s => s != null && re.test(String(s));
    }

    // Per text column, which rows hold a "shown last" value (null for columns with none).
    lastFlags(c) {
      const isLast = this.lastMatcher();
      if (!isLast || c.isNum || c.isDate) return null;
      const fmt = this.store.fmt[c.key];
      let any = false;
      const flags = fmt.map(v => { const f = isLast(v); if (f) any = true; return f; });
      return any ? flags : null;
    }

    // Each sort breaks ties left by the one before it; empty values always go last.
    sortRows(idx, sorts, cols) {
      const keys = sorts.map(s => {
        const c = this.colDef(s.key);
        return c && c.vkind !== 'run' && Object.assign(this.sortKeyArray(c), { dir: s.dir === 'desc' ? -1 : 1 });
      }).filter(Boolean);
      if (!keys.length) return this.defaultSort(idx, cols);
      // When the author keeps them last even when sorted, rows holding a "shown last" value in any visible text
      // column go after the rest, which are sorted as usual.
      let lastRow = null;
      if (this.cfg.lastRows.always) {
        const flags = cols.map(c => this.lastFlags(c)).filter(Boolean);
        if (flags.length) lastRow = i => flags.some(f => f[i]);
      }
      idx.sort((a, b) => {
        if (lastRow) { const la = lastRow(a), lb = lastRow(b); if (la !== lb) return la ? 1 : -1; }
        for (const { arr, num, dir } of keys) {
          const x = arr[a], y = arr[b];
          const xn = x == null || x !== x, yn = y == null || y !== y;
          if (xn || yn) { if (xn !== yn) return xn ? 1 : -1; continue; }
          const d = num ? (x < y ? -1 : x > y ? 1 : 0) : collator.compare(x, y);
          if (d) return dir * d;
        }
        return a - b;
      });
    }

    totalSpec(c) {
      if (c.virtual) {
        if (c.vkind === 'rank') return { type: 'none', source: 'calc', text: 'No total for a rank; zeros rank last and blanks are not ranked' };
        if (c.vkind === 'run') return { type: 'none', source: 'calc', text: 'No total for a running total' };
        if (c.vkind === 'diff' || c.vkind === 'pdiff') {
          const a = this.colDef(c.srcKey), b = this.colDef(c.otherKey);
          a.total = this.totalSpec(a); b.total = this.totalSpec(b);
          return { type: 'diff', mode: c.vkind, a, b, source: 'calc', text: c.vkind === 'diff' ? 'Total of ' + this.labelFor(a) + ' − total of ' + this.labelFor(b) : '(Total of ' + this.labelFor(a) + ' − total of ' + this.labelFor(b) + ') ÷ total of ' + this.labelFor(b) };
        }
        return { type: 'sum', source: 'pct', text: 'Share of the filtered total' };
      }
      const t = R.resolveTotal(c, this.store.columns, this.cfg.rules && this.cfg.rules.doc, this.cfg.totalOverrides[c.key]);
      t.text = R.describeTotal(t, this.store.columns);
      return t;
    }

    computeTotals(cols, idx) {
      const out = {};
      for (const c of cols) {
        if (!c.isNum) continue;
        c.total = this.totalSpec(c);
        out[c.key] = this.totalOf(c, idx);
      }
      out.__label = (cols.find(c => !c.isNum) || {}).key || null;
      return out;
    }

    // Difference columns total from their two columns' own totals, so ratio columns compare correctly.
    totalOf(c, rows) {
      const t = c.total;
      if (t && t.type === 'diff') {
        const A = this.totalOf(t.a, rows), B = this.totalOf(t.b, rows);
        if (A == null || B == null) return null;
        return t.mode === 'diff' ? A - B : B ? (A - B) / B : null;
      }
      return R.computeTotal(t, c.key, rows, k => this.store.raw[k]);
    }

    // Builds up to three nested subtotal levels. Group keys are paths, so the same city under two states
    // collapses separately.
    groupRows(cols, idx, keys) {
      const gcs = keys.map(k => cols.find(c => c.key === k)).filter(Boolean);
      this.groups = null;
      this.groupCols = gcs;
      this.topRoot = null;
      if (!gcs.length) {
        if (!this.topNActive()) return idx;
        this.topRoot = this.topFor(idx, '', 0, cols);
        const flat = [];
        this.emitTop(flat, this.topRoot, idx);
        return flat;
      }
      const sorts = this.sortList();
      const firstNum = sorts.map(s => cols.find(c => c.key === s.key)).find(c => c && c.isNum);
      const numDir = firstNum && sorts.find(s => s.key === firstNum.key).dir === 'desc' ? -1 : 1;
      // Date columns order their groups by date, not by the text of the date.
      const byLabel = (a, b) => a.sv != null && b.sv != null ? a.sv - b.sv : collator.compare(a.label, b.label);
      const build = (rows, level, prefix) => {
        const gc = gcs[level];
        const fmt = this.store.fmt[gc.key];
        const dateArr = gc.isDate ? this.store.raw[gc.key] : this.store.dates[gc.key] || null;
        const map = new Map();
        for (const i of rows) {
          const k = fmt[i] == null ? '\u0000null' : fmt[i];
          let g = map.get(k);
          if (!g) map.set(k, g = { type: 'group', key: prefix + k, label: fmt[i] == null ? this.nullFor(gc) : fmt[i], rows: [], level, colKey: gc.key, sv: dateArr ? dateArr[i] : null });
          g.rows.push(i);
        }
        const list = Array.from(map.values());
        for (const g of list) {
          g.totals = {};
          for (const c of cols) if (c.isNum) g.totals[c.key] = this.totalOf(c, g.rows);
          if (level + 1 < gcs.length) g.children = build(g.rows, level + 1, g.key + '\u0001');
          else g.top = this.topFor(g.rows, g.key, level + 1, cols);
        }
        // A sort on this level's own column orders its groups by name. Otherwise the first number sort orders
        // the groups at every level by their subtotal, as well as the rows inside them.
        const own = sorts.find(s => s.key === gc.key);
        if (own) {
          const d = own.dir === 'desc' ? -1 : 1;
          list.sort((a, b) => d * byLabel(a, b));
        } else if (firstNum && list.some(g => g.totals[firstNum.key] != null)) {
          list.sort((a, b) => {
            const x = a.totals[firstNum.key], y = b.totals[firstNum.key];
            if (x == null || y == null) return x == null && y == null ? byLabel(a, b) : x == null ? 1 : -1;
            return x === y ? byLabel(a, b) : (x < y ? -numDir : numDir);
          });
        } else list.sort(byLabel);
        // Groups named like "Other" go after the others at their level (only in the default order unless
        // the author keeps them last when sorted).
        const isLast = this.lastMatcher();
        if (isLast && (!sorts.length || this.cfg.lastRows.always)) {
          const keep = list.filter(g => !isLast(g.label));
          if (keep.length !== list.length) return keep.concat(list.filter(g => isLast(g.label)));
        }
        return list;
      };
      this.groups = build(idx, 0, '');
      const disp = [];
      const walk = list => {
        for (const g of list) {
          disp.push(g);
          if (this.collapsed.has(g.key)) continue;
          if (g.children) walk(g.children);
          else this.emitTop(disp, g.top, g.rows);
        }
      };
      walk(this.groups);
      return disp;
    }

    // ---------- top N ----------

    // The active Top N, if its column is on screen.
    topNActive() {
      const t = this.layout.topN;
      if (!t || !t.key || !(t.n > 0)) return null;
      const c = this.colDef(t.key);
      return c && c.isNum && this.view && (this.viewCols || []).some(x => x.key === t.key) ? t : null;
    }

    // Splits rows (already in display order) into the top N by the chosen column and the rest. Zeros and blanks
    // never make the top; ties at the cutoff follow the display order.
    topSplit(rows, t) {
      const v = this.store.raw[t.key];
      const dir = t.dir === 'bottom' ? 1 : -1;
      const cand = rows.map((i, pos) => ({ i, pos })).filter(o => v[o.i] === v[o.i] && v[o.i] != null && v[o.i] !== 0);
      cand.sort((a, b) => dir * (v[a.i] - v[b.i]) || a.pos - b.pos);
      const keep = new Set(cand.slice(0, t.n).map(o => o.i));
      return { shown: rows.filter(i => keep.has(i)), other: rows.filter(i => !keep.has(i)) };
    }

    // Works out the top rows and the "Other" row (which totals the rest with the usual rules) for a group or the
    // whole table. Null when Top N is off.
    topFor(rows, key, level, cols) {
      const t = this.topNActive();
      if (!t) return null;
      const { shown, other } = this.topSplit(rows, t);
      if (!other.length) return { shown, node: null };
      const node = { type: 'group', isOther: true, key: '\u0002other\u0001' + key, label: 'Other', rows: other, level, colKey: this.otherLabelKey(cols), totals: {} };
      for (const c of cols) if (c.isNum) node.totals[c.key] = this.totalOf(c, other);
      return { shown, node };
    }

    // Adds the rows to the display: all of them, or the top rows, then Other (and its rows when it is open).
    emitTop(disp, top, rows) {
      if (!top) { for (const i of rows) disp.push(i); return; }
      for (const i of top.shown) disp.push(i);
      if (!top.node) return;
      disp.push(top.node);
      if (this.otherOpen.has(top.node.key)) for (const i of top.node.rows) disp.push(i);
    }

    // The Other label goes in the first text column that isn't a subtotal column.
    otherLabelKey(cols) {
      const g = new Set((this.groupCols || []).map(c => c.key));
      const c = cols.find(x => !x.isNum && !g.has(x.key)) || cols.find(x => !x.isNum) || cols[0];
      return c ? c.key : null;
    }

    groupsAtLevel(level) {
      const out = [];
      const walk = list => { for (const g of list || []) { if (g.level === level) out.push(g); else if (g.level < level) walk(g.children); } };
      walk(this.groups);
      return out;
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
          <span class="tp-vfs"></span>
          <span class="tp-spacer"></span>
          <button class="tp-chip tp-lchip" type="button" hidden title="Your column changes are saved in this browser only">Custom layout · Reset</button>
          <button class="tp-icon tp-lay" type="button" title="Layouts" aria-label="Layouts" hidden>${ICONS.lay}</button>
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
      q('.tp-lchip').addEventListener('click', e => { if (e.currentTarget.dataset.mode === 'menu') this.openPresets(e.currentTarget); else this.resetLocal(); });
      q('.tp-lay').addEventListener('click', e => this.openPresets(e.currentTarget));
      q('.tp-vfs').addEventListener('click', e => {
        if (e.target.closest('button[data-top]')) return this.updateLayout({ topN: null });
        const b = e.target.closest('button[data-k]');
        if (b) this.clearValueFilter(b.dataset.k);
      });
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
      // One chip per column filtered with Keep only / Exclude; clicking it removes that filter.
      const chips = [];
      if (st) {
        for (const c of this.view.cols) {
          const f = this.filters[c.key];
          if (!f || !(f.keep || f.exclude)) continue;
          const name = v => this.valueText(c, v);
          const list = (f.keep || f.exclude).map(name);
          const text = c.label + (f.keep ? ': only ' : ': not ') + (list.length > 2 ? list.slice(0, 2).join(', ') + ' +' + (list.length - 2) : list.join(', '));
          chips.push(`<button type="button" class="tp-chip tp-vchip" data-k="${esc(c.key)}" title="${esc(c.label + (f.keep ? ' kept: ' : ' excluded: ') + list.join(', ') + '\nClick to remove this filter')}">${esc(text.length > 48 ? text.slice(0, 47) + '…' : text)} ✕</button>`);
        }
      }
      const top = st && this.topNActive();
      if (top) {
        const tc = this.colDef(top.key);
        chips.push(`<button type="button" class="tp-chip tp-vchip" data-top="1" title="Showing the ${top.dir === 'bottom' ? 'smallest' : 'largest'} ${top.n} rows${this.groups ? ' in each group' : ''}; the rest are in Other.\nClick to show all rows">${esc((top.dir === 'bottom' ? 'Bottom ' : 'Top ') + top.n + ' by ' + this.labelFor(tc))} ✕</button>`);
      }
      const vfs = q('.tp-vfs');
      const html = chips.join('');
      if (vfs.innerHTML !== html) vfs.innerHTML = html;
      this.statusEl.classList.toggle('err', !!this.statusErr);
      this.searchEl.hidden = !this.cfg.showSearch;
      q('.tp-gear').hidden = !this.authoring;
      q('.tp-cols').hidden = !this.canEditLayout();
      const chip = q('.tp-lchip');
      const act = this.presets && this.activePreset();
      q('.tp-lay').hidden = this.authoring || !this.cfg.allowViewerLayout || !this.presets;
      if (act && !this.authoring) {
        chip.hidden = false;
        chip.dataset.mode = 'menu';
        chip.textContent = 'Layout: ' + act.name + (this.presetEdited(act) ? ' (edited)' : '');
        chip.title = 'Your saved layout. Click for layouts.';
      } else {
        chip.hidden = !(this.local && Object.keys(this.local).length) || this.authoring;
        chip.dataset.mode = 'reset';
        chip.textContent = 'Custom layout · Reset';
        chip.title = 'Your column changes are saved in this browser only';
      }
      const t = this.timing;
      q('.tp-info').title = 'Totals Table ' + VERSION + (t ? '\n' : '') + (t ? 'Tableau query ' + (t.query / 1000).toFixed(2) + 's · all rows loaded ' + ((t.fetch || 0) / 1000).toFixed(2) + 's · last draw ' + Math.round(t.draw || 0) + ' ms' + (this.freezePaused ? '\nFreeze paused: frozen columns are wider than 60% of the table.' : '') : '');
    }

    template() { return this.view.cols.map(c => c.w + 'px').join(' '); }

    negCls(c, v) { return c.spec && c.spec.red && v < 0 ? ' neg' : ''; }

    cellAttrs(c, extra, extraStyle) {
      let cls = 'tp-c' + (c.isNum ? ' num' : '') + (c.cls || '') + (extra ? ' ' + extra : '');
      let style = extraStyle || '';
      if (c.frozen) {
        cls += ' fz';
        style = 'left:' + c.left + 'px;' + style;
        if (c === this.view.cols[this.view.nFrozen - 1]) cls += ' fz-edge';
      }
      return `class="${cls}"` + (style ? ` style="${style}"` : '');
    }

    // ---------- conditional formatting ----------

    paramNumber(name) {
      const p = this.params.find(x => x.name.toLowerCase() === String(name).toLowerCase());
      if (!p) return null;
      const parsed = R.parseFormatted(p.value);
      if (parsed) return parsed.spec.pct ? parsed.value * 100 : parsed.value;
      const n = Number(p.raw);
      return isFinite(n) ? n : null;
    }

    cutValue(def) {
      if (!def) return null;
      if (def.param) return this.paramNumber(def.param);
      if (def.value === '' || def.value == null) return null;
      const n = Number(def.value);
      return isFinite(n) ? n : null;
    }

    // Thresholds compare in the units shown on screen, so 5 means 5% on a percent column.
    cfLevel(rule, c, v) {
      const x = v * (c.spec && c.spec.pct ? 100 : 1);
      const lo = this.cutValue(rule.low), hi = this.cutValue(rule.high);
      if (rule.better === 'lower') {
        if (lo != null && x <= lo) return 'good';
        if (hi != null && x > hi) return 'bad';
      } else {
        if (lo != null && x < lo) return 'bad';
        if (hi != null && x >= hi) return 'good';
      }
      return lo != null && hi != null ? 'warn' : null;
    }

    computeCfStats(cols, idx) {
      this.cfStats = {};
      for (const c of cols) {
        const rule = c.isNum && this.cfg.cf[c.key];
        if (!rule || (rule.type !== 'scale' && rule.type !== 'bars')) continue;
        const raw = this.store.raw[c.key];
        const vals = [];
        for (const i of idx) { const v = raw[i]; if (v === v) vals.push(v); }
        if (!vals.length) continue;
        vals.sort((x, y) => x - y);
        const min = vals[0], max = vals[vals.length - 1];
        // Color scales span the 5th to 95th percentile, so a few extreme rows don't wash out everything else.
        const p = q => vals[Math.min(vals.length - 1, Math.max(0, Math.round(q * (vals.length - 1))))];
        this.cfStats[c.key] = { min, max, lo: p(0.05), hi: p(0.95), maxAbs: Math.max(Math.abs(min), Math.abs(max)) };
      }
    }

    // Returns { style, cls, sym, word } for one numeric cell, or null when no formatting applies.
    cellFormat(c, v, isTotal) {
      const rule = this.cfg.cf[c.key];
      if (!rule || !rule.type || v == null || v !== v) return null;
      if ((this.layout.cfOff || []).includes(c.key)) return null;
      if (isTotal && !rule.totals) return null;
      if (rule.type === 'threshold') {
        const lvl = this.cfLevel(rule, c, v);
        if (!lvl) return null;
        const style = rule.target === 'text' ? 'color:' + CF.text[lvl] + ';' : 'background-color:' + CF.bg[lvl] + ';';
        return { style, cls: 'cf', sym: rule.symbols !== false ? CF.sym[lvl] : '', word: CF.word[lvl], level: lvl, fill: rule.target === 'text' ? null : CF.bg[lvl], font: rule.target === 'text' ? CF.text[lvl] : null };
      }
      const st = this.cfStats[c.key];
      if (!st) return null;
      if (rule.type === 'scale') {
        const t = st.hi === st.lo ? 0.5 : Math.max(0, Math.min(1, (v - st.lo) / (st.hi - st.lo)));
        let color;
        if (rule.better === 'none') color = mix('#FFFFFF', CF.bar, t);
        else {
          const u = rule.better === 'lower' ? 1 - t : t;
          color = u < 0.5 ? mix(CF.bg.bad, '#FFFFFF', u * 2) : mix('#FFFFFF', CF.bg.good, (u - 0.5) * 2);
        }
        return { style: 'background-color:' + color + ';', cls: 'cf', fill: color };
      }
      if (rule.type === 'bars') {
        const pct = st.maxAbs ? Math.min(100, Math.abs(v) / st.maxAbs * 100) : 0;
        const col = v < 0 ? CF.bg.bad : CF.bar;
        return { style: `background-image:linear-gradient(to right,${col} ${pct.toFixed(1)}%,transparent ${pct.toFixed(1)}%);`, cls: 'cf', bar: pct };
      }
      return null;
    }

    cfInner(f, text) {
      if (!f || !f.word) return text;
      return (f.sym ? `<span class="tp-sym tp-sym-${f.level}" aria-hidden="true">${f.sym}</span>` : '') + `<span class="tp-sr">${f.word}: </span>` + text;
    }

    cfDescribe(c) {
      const r = this.cfg.cf[c.key];
      if (!r || !r.type) return '';
      if ((this.layout.cfOff || []).includes(c.key)) return '\nFormatting: hidden (column menu → Show formatting)';
      if (r.type === 'bars') return '\nFormatting: data bars';
      if (r.type === 'scale') return '\nFormatting: color scale' + (r.better === 'none' ? '' : r.better === 'lower' ? ', lower is better' : ', higher is better');
      const name = d => d && d.param ? d.param : d && d.value !== '' && d.value != null ? d.value : null;
      const lo = name(r.low), hi = name(r.high);
      return '\nFormatting: ' + (r.better === 'lower'
        ? [lo != null ? '✓ at or below ' + lo : '', hi != null ? '✕ above ' + hi : ''].filter(Boolean).join(', ')
        : [lo != null ? '✕ below ' + lo : '', hi != null ? '✓ at or above ' + hi : ''].filter(Boolean).join(', ')) + (lo != null && hi != null ? ', ! in between' : '');
    }

    // How a number column totals, in words, for the header tooltip. Plain sums say nothing.
    formulaLine(c) {
      const t = c.isNum && c.total;
      if (!t) return '';
      if (c.vkind === 'pct') return 'Share of the filtered total';
      if (c.vkind === 'rank' || c.vkind === 'run') return '';
      const problem = t.problem ? ' (total blank: ' + t.problem + ')' : '';
      if (t.type === 'ratio' || t.type === 'diff') return t.text + problem;
      if (t.type === 'avg' || t.type === 'min' || t.type === 'max') return 'Totals as the ' + { avg: 'average', min: 'minimum', max: 'maximum' }[t.type];
      if (t.type === 'none') return 'No total';
      return '';
    }

    renderHead() {
      const { cols, totals } = this.view;
      this.grid.style.setProperty('--cols', this.template());
      this.grid.style.width = cols.reduce((s, c) => s + c.w, 0) + 'px';
      const sorts = this.sortList().filter(s => cols.some(c => c.key === s.key));
      const canEdit = this.canEditLayout();
      let h = '';
      if (this.view.spans) {
        h += '<div class="tp-grow">';
        for (const s of this.view.spans) {
          const c0 = cols[s.start];
          const cls = 'tp-c tp-g' + (s.label ? (canEdit ? ' mv' : '') : ' empty') + (s.frozen ? ' fz' : '') + (s.frozen && s.end === this.view.nFrozen - 1 ? ' fz-edge' : '');
          const style = `grid-column:${s.start + 1} / ${s.end + 2};` + (s.frozen ? 'left:' + c0.left + 'px;' : '');
          h += `<div class="${cls}" style="${style}" data-s="${s.start}" data-e="${s.end}"` +
            (s.label ? ` title="${esc(s.label + (canEdit ? '\nDrag to move this group' : ''))}">` +
              `<span class="tp-gl">${esc(s.label)}</span><button class="tp-hm tp-gm" type="button" data-s="${s.start}" aria-label="Group menu for ${esc(s.label)}">▾</button>` : '>') + '</div>';
        }
        h += '</div>';
      }
      h += '<div class="tp-hrow">';
      cols.forEach((c, i) => {
        const si = sorts.findIndex(s => s.key === c.key);
        const sort = sorts[si];
        const arrow = sort ? (sort.dir === 'desc' ? ' ▼' : ' ▲') + (sorts.length > 1 ? `<span class="tp-sn">${si + 1}</span>` : '') : '';
        const sortTip = c.vkind === 'run' ? '\nFollows the row order' : sort
          ? '\nSorted ' + (c.isNum ? (sort.dir === 'desc' ? 'largest to smallest' : 'smallest to largest') : (sort.dir === 'desc' ? 'Z to A' : 'A to Z')) + (sorts.length > 1 ? ' (sort ' + (si + 1) + ' of ' + sorts.length + ')' : '') + ' · Shift-click to add a sort'
          : '\nClick to sort';
        const how = this.formulaLine(c);
        const p = this.paramFor(c);
        h += `<div ${this.cellAttrs(c, 'tp-h')} data-i="${i}" title="${esc(c.label + (p && this.authoring ? ' (from parameter ' + p.name + ')' : '') + (how ? '\n' + how : '') + sortTip)}">` +
          `<span class="tp-hl">${esc(c.label)}${arrow}</span>` +
          `<button class="tp-hm" type="button" data-i="${i}" aria-label="Column menu for ${esc(c.label)}">▾</button>` +
          (canEdit ? `<span class="tp-rz" data-i="${i}"></span>` : '') + '</div>';
      });
      h += '</div>';
      if (this.cfg.showFilters) {
        h += '<div class="tp-frow">';
        cols.forEach((c, i) => {
          const f = this.filters[c.key] || {};
          if (c.virtual) { h += `<div ${this.cellAttrs(c, 'tp-f')}></div>`; return; }
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
          const f = loading ? null : this.cellFormat(c, v, true);
          h += `<div ${this.cellAttrs(c, (c.total && c.total.problem ? 'warn' : '') + (f ? ' cf' : '') + (loading ? '' : this.negCls(c, v)), f && f.style)} title="${esc(title)}">${this.cfInner(f, esc(text))}</div>`;
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
      const rowH = this.rowH || 26;
      const start = Math.max(0, Math.floor(top / rowH) - OVERSCAN);
      const end = Math.min(total, Math.ceil((top + viewH) / rowH) + OVERSCAN);
      if (!force && start === this.rStart && end === this.rEnd) return;
      this.rStart = start;
      this.rEnd = end;
      this.body.style.height = total * rowH + 'px';
      this.pad.style.height = start * rowH + 'px';
      const st = this.store;
      const nulls = cols.map(c => esc(this.nullFor(c)));
      const fmts = cols.map(c => st.fmt[c.key]);
      const links = cols.map(c => { const l = this.linkFor(c); return l ? { urls: st.fmt[l.urlKey], label: l.label } : null; });
      const cfCols = cols.map(c => !!(c.isNum && this.cfg.cf[c.key] && this.cfg.cf[c.key].type));
      const reds = cols.map(c => !!(c.isNum && c.spec && c.spec.red));
      const rg = this.normRange();
      let h = '';
      for (let r = start; r < end; r++) {
        const d = disp[r];
        if (typeof d === 'object') {
          const open = d.isOther ? this.otherOpen.has(d.key) : !this.collapsed.has(d.key);
          h += `<div class="tp-row tp-grp l${d.level}${d.isOther ? ' tp-other' : ''}" data-r="${r}">`;
          cols.forEach((c, j) => {
            if (c.key === d.colKey) h += `<div ${this.cellAttrs(c)} data-c="${j}"><span class="tp-caret">${open ? '▾' : '▸'}</span>${esc(d.label)} <span class="tp-cnt">(${d.rows.length.toLocaleString()})</span></div>`;
            else if (c.isNum) {
              const gv = d.totals[c.key];
              const f = this.cellFormat(c, gv, true);
              h += `<div ${this.cellAttrs(c, (f ? 'cf' : '') + this.negCls(c, gv), f && f.style)} data-c="${j}">${this.cfInner(f, esc(gv == null ? '' : R.formatNumber(gv, c.spec)))}</div>`;
            }
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
          const f = cfCols[j] ? this.cellFormat(cols[j], st.raw[cols[j].key][d], false) : null;
          let inner = v == null ? nulls[j] : this.cfInner(f, esc(v));
          const lk = links[j];
          const url = lk && this.safeUrl(lk.urls[d]);
          if (url && (v != null || lk.label)) inner = `<a class="tp-link" href="${esc(url)}" target="_blank" rel="noopener noreferrer" title="${esc(url)}">${esc(lk.label || v)}</a>`;
          h += `<div ${this.cellAttrs(cols[j], (v == null ? 'nul' : '') + (inRange ? ' rg' : '') + (f ? ' cf' : '') + (reds[j] ? this.negCls(cols[j], st.raw[cols[j].key][d]) : ''), f && f.style)} data-c="${j}">${inner}</div>`;
        }
        h += '</div>';
      }
      if (!total && this.store.complete) {
        h = `<div class="tp-empty">${this.store.n ? 'No rows match these filters. <button type="button" data-a="clear-filters">Clear filters</button>' : 'No rows to show.'}</div>`;
      }
      this.rowsEl.innerHTML = h;
    }

    // ---------- header interactions ----------

    wireHead() {
      this.head.addEventListener('click', e => {
        const gm = e.target.closest('.tp-gm');
        if (gm) { this.openGroupMenu(+gm.dataset.s, gm); return; }
        const menu = e.target.closest('.tp-hm');
        if (menu) { this.openColumnMenu(+menu.dataset.i, menu); return; }
        if (this.suppressClick || e.target.closest('.tp-rz,.tp-fzh,input')) return;
        const cell = e.target.closest('.tp-h');
        if (!cell) return;
        const c = this.view.cols[+cell.dataset.i];
        if (c.vkind === 'run') return this.flash('Running totals follow the row order, so they can\'t be sorted.');
        const sorts = this.sortList();
        const def = c.isNum ? 'desc' : 'asc', alt = c.isNum ? 'asc' : 'desc';
        // Shift-click adds the column as the next sort, flips it, then takes it out again.
        if (e.shiftKey && sorts.length) {
          const next = sorts.slice();
          const i = next.findIndex(s => s.key === c.key);
          if (i < 0) next.push({ key: c.key, dir: def });
          else if (next[i].dir === def) next[i] = { key: c.key, dir: alt };
          else next.splice(i, 1);
          return this.setSort(next.length ? next : null);
        }
        const s = sorts.length === 1 ? sorts[0] : null;
        const next = !s || s.key !== c.key ? { key: c.key, dir: def } : s.dir === def ? { key: c.key, dir: alt } : null;
        this.setSort(next ? [next] : null);
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
        const grp = e.target.closest('.tp-g.mv');
        if (grp) { const r = this.groupRun(+grp.dataset.s, +grp.dataset.e); return this.startColumnDrag(e, r.start, r.end); }
        const cell = e.target.closest('.tp-h');
        if (cell && this.canEditLayout()) this.startColumnDrag(e, +cell.dataset.i);
      });
    }

    // Pointer-based dragging; native HTML drag-and-drop hands text to Tableau, which tries to open it as a data file.
    // A group header drags its whole run of columns (from..last) as one block.
    startColumnDrag(e, from, last) {
      if (last == null) last = from;
      e.preventDefault();
      const x0 = e.clientX;
      const cols = this.view.cols;
      const nF = this.view.nFrozen;
      const frozenW = this.frozenWidth || 0;
      const sc = this.scroller;
      const mids = [];
      let acc = 0;
      cols.forEach(c => { mids.push(acc + c.w / 2); acc += c.w; });
      const bounds = [0];
      cols.forEach(c => bounds.push(bounds[bounds.length - 1] + c.w));
      let guide = null, gap = from, lastX = x0, raf = 0;
      const ok = g => this.dropOk(g, from, last);
      // Frozen columns stay pinned on the left, so the pointer maps to them there and to the scrolled columns elsewhere.
      const visible = g => g <= nF || bounds[g] >= frozenW + sc.scrollLeft - 1;
      const place = () => {
        const vx = lastX - sc.getBoundingClientRect().left;
        const x = nF && vx < frozenW ? vx : vx + sc.scrollLeft;
        gap = mids.filter(m => m < x).length;
        if (!ok(gap) || !visible(gap)) {
          let best = -1;
          bounds.forEach((b, g) => { if (ok(g) && visible(g) && (best < 0 || Math.abs(b - x) < Math.abs(bounds[best] - x))) best = g; });
          if (best >= 0) gap = best;
        }
        guide.style.left = (gap <= nF ? bounds[gap] + sc.scrollLeft : bounds[gap]) + 'px';
      };
      // Holding the pointer near either edge scrolls the table, so far-off positions can be reached.
      const EDGE = 48;
      const tick = () => {
        raf = 0;
        if (!guide) return;
        const r = sc.getBoundingClientRect();
        const vx = lastX - r.left;
        const leftEdge = nF ? frozenW : 0;
        let dx = 0;
        if (vx > sc.clientWidth - EDGE) dx = Math.ceil((vx - (sc.clientWidth - EDGE)) / 3);
        else if (vx >= leftEdge && vx < leftEdge + EDGE && sc.scrollLeft > 0) dx = -Math.ceil((leftEdge + EDGE - vx) / 3);
        if (dx) {
          sc.scrollLeft += dx;
          place();
          raf = setTimeout(tick, 16);
        }
      };
      const move = ev => {
        lastX = ev.clientX;
        if (!guide) {
          if (Math.abs(ev.clientX - x0) < 6) return;
          guide = document.createElement('div');
          guide.className = 'tp-guide';
          this.grid.appendChild(guide);
          this.grid.classList.add('tp-dragging');
        }
        place();
        if (!raf) raf = setTimeout(tick, 16);
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        if (raf) clearTimeout(raf);
        if (!guide) return;
        guide.remove();
        this.grid.classList.remove('tp-dragging');
        this.suppressClick = true;
        setTimeout(() => { this.suppressClick = false; }, 0);
        if (last !== from) return this.moveBlock(from, last, gap);
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

    // A group (from..last) lands only between groups or among ungrouped columns, never inside another group.
    dropOk(g, from, last) {
      const cols = this.view.cols;
      return last === from || g <= 0 || g >= cols.length || (g > from && g <= last) ||
        !this.groupOf(cols[g - 1]) || this.groupOf(cols[g - 1]) !== this.groupOf(cols[g]);
    }

    // Moves visible columns from..last so they start at visible gap position `gap` (0..cols.length).
    moveBlock(from, last, gap) {
      const cols = this.view.cols;
      if (gap >= from && gap <= last + 1) return;
      const keys = cols.slice(from, last + 1).map(c => c.key);
      const order = this.orderedKeys().filter(k => !keys.includes(k));
      const at = gap < cols.length ? order.indexOf(cols[gap].key) : order.indexOf(cols[cols.length - 1].key) + 1;
      order.splice(at, 0, ...keys);
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
        if (c.w !== w0) this.updateLayout({ widths: Object.assign({}, this.layout.widths, { [c.key]: c.w }) });
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
      const n0 = n;
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
        if (n !== n0) this.updateLayout({ freeze: n });
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
        // A link click opens the link and nothing else: no selection, no redraw that would swallow the click.
        if (e.button !== 0 || e.target.closest('a.tp-link')) return;
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
        if (e.target.closest('a.tp-link')) { this.dragStart = null; return; }
        const ds = this.dragStart;
        this.dragStart = null;
        if (!ds || ds.moved || e.shiftKey) { this.renderBody(true); return; }
        this.range = { r0: ds.r, c0: ds.c, r1: ds.r, c1: ds.c };
        this.clickRow(ds.r, e.ctrlKey || e.metaKey, { x: e.clientX, y: e.clientY });
      });
      this.rowsEl.addEventListener('mouseleave', () => this.unhover());
      this.rowsEl.addEventListener('click', e => { if (e.target.closest('button[data-a="clear-filters"]')) this.clearFilters(); });
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
        const pick = this.pickedValues();
        const short = s => s.length > 30 ? s.slice(0, 29) + '…' : s;
        const shown = pick && (pick.values.length === 1 ? '"' + short(this.valueText(pick.col, pick.values[0])) + '"' : pick.values.length + ' values');
        const html = '<div class="tp-menu" role="menu">' +
          (pick ? `<button type="button" data-v="keep">Keep only ${esc(shown)}</button><button type="button" data-v="exclude">Exclude ${esc(shown)}</button><hr>` : '') +
          `<button type="button" data-m="cells">${what}</button>` +
          '<button type="button" data-m="headers">Copy with headers</button><hr>' +
          `<button type="button" data-m="all">Copy whole table (${this.view.idx.length.toLocaleString()} rows)</button></div>`;
        this.openPop({ getBoundingClientRect: () => ({ left: x, bottom: y }) }, html, pop => pop.addEventListener('click', ev => {
          const v = ev.target.closest('button[data-v]');
          if (v) { this.closePop(); return this.valueFilter(pick, v.dataset.v); }
          const b = ev.target.closest('button[data-m]');
          if (!b) return;
          this.closePop();
          const text = this.copyText(b.dataset.m);
          if (text) this.copyToClipboard(text, b.dataset.m === 'all' ? this.view.idx.length : null);
        }));
      });
      this.el.addEventListener('contextmenu', e => e.preventDefault());
    }

    // ---------- keep only / exclude ----------

    // The distinct shown values in the selected cells of one column (subtotal rows count as their group's value).
    pickedValues() {
      const g = this.normRange();
      if (!g || g.c0 !== g.c1) return null;
      const col = this.view.cols[g.c0];
      if (!col || col.virtual) return null;
      const fmt = this.store.fmt[col.key];
      const vals = new Set();
      for (let r = g.r0; r <= g.r1 && vals.size <= 50; r++) {
        const d = this.view.disp[r];
        let i = d;
        if (typeof d === 'object') { if (d.colKey !== col.key) continue; i = d.rows[0]; }
        if (typeof i !== 'number') continue;
        if (col.isNum) { const v = this.store.raw[col.key][i]; vals.add(v === v && v != null ? v : NUL); }
        else vals.add(fmt[i] == null ? NUL : fmt[i]);
      }
      return vals.size ? { col, values: Array.from(vals) } : null;
    }

    valueText(c, v) { return v === NUL ? this.nullFor(c) : typeof v === 'number' ? R.formatNumber(v, c.spec) : String(v); }

    // Table-only filters, like the filter row: they last for this session and never change Tableau's filters.
    valueFilter(pick, mode) {
      const f = this.filters[pick.col.key] = Object.assign({}, this.filters[pick.col.key]);
      if (mode === 'keep') { f.keep = pick.values.slice(); delete f.exclude; }
      else f.exclude = Array.from(new Set((f.exclude || []).concat(pick.values)));
      this.range = null;
      this.refresh();
    }

    // Clears every table-only filter: the filter row, search, and Keep only / Exclude.
    clearFilters() {
      this.filters = {};
      this.search = '';
      this.searchEl.value = '';
      this.refresh();
    }

    clearValueFilter(key) {
      const f = this.filters[key];
      if (f) { delete f.keep; delete f.exclude; }
      this.refresh();
    }

    normRange() {
      const g = this.range;
      if (!g) return null;
      return { r0: Math.min(g.r0, g.r1), r1: Math.max(g.r0, g.r1), c0: Math.min(g.c0, g.c1), c1: Math.max(g.c0, g.c1) };
    }

    clickRow(r, toggle, point) {
      const d = this.view.disp[r];
      if (typeof d === 'object') {
        const set = d.isOther ? this.otherOpen : this.collapsed;
        if (set.has(d.key)) set.delete(d.key); else set.add(d.key);
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
          else if (d.isOther ? !this.otherOpen.has(d.key) : this.collapsed.has(d.key)) all.push(...d.rows);
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

    note(msg, ms) {
      this.setStatus(msg, false);
      clearTimeout(this.flashT);
      this.flashT = setTimeout(() => this.setStatus(null), ms || 6000);
    }

    // ---------- viewer layouts (presets) ----------

    // Named layouts a viewer saves in their own browser, per worksheet and field list, like the working layout.
    presetKey() { return this.layoutKey().replace('tp-layout:', 'tp-presets:'); }

    loadPresets() {
      this.presets = { list: [], active: null };
      this.storageOk = true;
      try { localStorage.setItem('tp-test', '1'); localStorage.removeItem('tp-test'); } catch (e) { this.storageOk = false; }
      if (this.authoring || !this.cfg.allowViewerLayout) return;
      try {
        const p = JSON.parse(localStorage.getItem(this.presetKey()) || 'null');
        if (p && Array.isArray(p.list)) {
          const list = p.list.filter(x => x && typeof x.name === 'string' && x.name.trim()).map(x => ({ name: x.name, layout: cleanLayout(x.layout), savedAt: x.savedAt }));
          this.presets = { list, active: list.some(x => x.name === p.active) ? p.active : null };
        }
      } catch (e) { /* keep empty */ }
    }

    // Presets stay in memory for the session even when the browser won't store them.
    savePresets() {
      try { localStorage.setItem(this.presetKey(), JSON.stringify(this.presets)); return true; } catch (e) { return false; }
    }

    activePreset() { return this.presets && this.presets.active ? this.presets.list.find(p => p.name === this.presets.active) || null : null; }

    presetEdited(p) { return JSON.stringify(this.local || {}) !== JSON.stringify(p.layout || {}); }

    setLocal(layout) {
      this.local = layout && Object.keys(layout).length ? JSON.parse(JSON.stringify(layout)) : null;
      try {
        if (this.local) localStorage.setItem(this.layoutKey(), JSON.stringify(this.local)); else localStorage.removeItem(this.layoutKey());
      } catch (e) { /* storage blocked */ }
    }

    applyPreset(name) {
      const p = this.presets.list.find(x => x.name === name);
      if (!p) return;
      this.setLocal(p.layout);
      this.presets.active = name;
      this.savePresets();
      this.collapsed = new Set();
      this.refresh();
    }

    savePresetAs(name) {
      name = String(name || '').trim().slice(0, 60);
      if (!name) return;
      const layout = JSON.parse(JSON.stringify(this.local || {}));
      const at = this.presets.list.findIndex(x => x.name === name);
      const entry = { name, layout, savedAt: new Date().toISOString() };
      if (at >= 0) this.presets.list[at] = entry; else this.presets.list.push(entry);
      this.presets.active = name;
      this.savePresets();
      this.refresh();
      let tipped = false;
      try { tipped = !!localStorage.getItem('tp-presets-tip'); if (!tipped) localStorage.setItem('tp-presets-tip', '1'); } catch (e) { tipped = true; }
      if (!tipped) this.note('Saved "' + name + '". Tip: layouts live in this browser only, so export a backup if you\'d like to keep them safe.', 9000);
      else this.flashOk('Saved "' + name + '"');
    }

    exportList() {
      if (this.presets.list.length) return this.presets.list;
      return this.local && Object.keys(this.local).length ? [{ name: 'My layout', layout: this.local }] : [];
    }

    exportPresets() {
      const list = this.exportList();
      if (!list.length) return;
      const doc = { type: 'totals-table-layouts', v: 1, sheet: this.ws.name, exportedAt: new Date().toISOString(), layouts: list };
      saveBlob(new Blob([JSON.stringify(doc, null, 2)], { type: 'application/json' }), TPExport.buildFileName('Table layouts {sheet}', { sheet: this.ws.name }, 'json'));
    }

    importPresets(text) {
      let doc;
      try { doc = JSON.parse(text); } catch (e) { doc = null; }
      if (!doc || doc.type !== 'totals-table-layouts' || !Array.isArray(doc.layouts)) { this.flash('That file isn\'t a table layouts file.'); return; }
      let n = 0;
      for (const p of doc.layouts) {
        if (!p || typeof p.name !== 'string' || !p.name.trim() || typeof p.layout !== 'object') continue;
        const entry = { name: p.name.trim().slice(0, 60), layout: cleanLayout(p.layout), savedAt: typeof p.savedAt === 'string' ? p.savedAt : new Date().toISOString() };
        const at = this.presets.list.findIndex(x => x.name === entry.name);
        if (at >= 0) this.presets.list[at] = entry; else this.presets.list.push(entry);
        n++;
      }
      this.savePresets();
      this.refresh();
      this.note(n + (n === 1 ? ' layout' : ' layouts') + ' restored.', 4000);
    }

    openPresets(anchor, mode, confirmDelete) {
      const L = this.presets.list;
      const act = this.activePreset();
      const edited = act && this.presetEdited(act);
      const hasLocal = !!(this.local && Object.keys(this.local).length);
      const btn = (a, label, extra) => `<button type="button" data-a="${a}"${extra || ''}>${esc(label)}</button>`;
      let h = '<div class="tp-menu tp-lays" role="menu"><div class="tp-ph">Layouts</div>';
      h += btn('pub', (!act && !hasLocal ? '✓ ' : '') + 'Published layout');
      L.forEach((p, i) => { h += btn('apply', (p === act ? '✓ ' : '') + p.name, ` data-i="${i}"`); });
      h += '<hr>';
      if (mode === 'save' || mode === 'rename') {
        const val = mode === 'rename' && act ? act.name : '';
        h += `<div class="tp-lname"><input type="text" maxlength="60" value="${esc(val)}" placeholder="Layout name" aria-label="Layout name"><button type="button" data-a="${mode}-ok">${mode === 'save' ? 'Save' : 'Rename'}</button></div>`;
      } else h += btn('save', 'Save current layout as…');
      if (act && edited) h += btn('update', 'Update "' + act.name + '" with your changes') + btn('undo', 'Undo changes to "' + act.name + '"');
      if (act && mode !== 'rename') h += btn('rename', 'Rename "' + act.name + '"…') + btn('delete', (confirmDelete ? 'Click again to delete "' : 'Delete "') + act.name + '"');
      h += '<hr>' + btn('export', 'Export layouts', this.exportList().length ? '' : ' disabled') + btn('import', 'Import layouts…') +
        '<input type="file" accept=".json,application/json" hidden>';
      h += '<div class="tp-lnote">' + (this.storageOk
        ? 'Your layouts are saved in this browser. Private windows, other computers, and clearing cookies and site data won\'t have them. <button type="button" class="linkish" data-a="export">Export a backup</button> anytime.'
        : 'This browser isn\'t letting the table save layouts right now, often because of a privacy setting. You can still arrange the table, and Export will keep a copy.') + '</div></div>';
      this.openPop(anchor, h, pop => {
        const input = pop.querySelector('.tp-lname input');
        if (input) {
          input.focus();
          input.select();
          input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); pop.querySelector('.tp-lname button').click(); } });
        }
        pop.querySelector('input[type="file"]').addEventListener('change', e => {
          const f = e.target.files && e.target.files[0];
          if (!f) return;
          this.closePop();
          f.text().then(t => this.importPresets(t));
        });
        pop.addEventListener('click', e => {
          const b = e.target.closest('button[data-a]');
          if (!b || b.disabled) return;
          const a = b.dataset.a;
          if (a === 'save' || a === 'rename') return this.openPresets(anchor, a);
          if (a === 'delete' && !confirmDelete) return this.openPresets(anchor, null, true);
          if (a === 'import') return pop.querySelector('input[type="file"]').click();
          this.closePop();
          if (a === 'pub') return this.resetLocal();
          if (a === 'apply') return this.applyPreset(L[+b.dataset.i].name);
          if (a === 'save-ok') return this.savePresetAs(input.value);
          if (a === 'rename-ok') {
            const name = input.value.trim().slice(0, 60);
            if (!name || !act) return;
            if (L.some(x => x !== act && x.name === name)) return this.flash('There is already a layout named "' + name + '". Pick another name.');
            act.name = name;
            this.presets.active = name;
            this.savePresets();
            return this.refresh();
          }
          if (a === 'update') { act.layout = JSON.parse(JSON.stringify(this.local || {})); act.savedAt = new Date().toISOString(); this.savePresets(); this.refresh(); return this.flashOk('Updated "' + act.name + '"'); }
          if (a === 'undo') return this.applyPreset(act.name);
          if (a === 'delete') {
            this.presets.list = L.filter(x => x !== act);
            this.presets.active = null;
            this.savePresets();
            return this.refresh();
          }
          if (a === 'export') return this.exportPresets();
        });
      });
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
      const sorts = this.sortList();
      const inSort = sorts.some(s => s.key === c.key);
      const others = sorts.some(s => s.key !== c.key);
      const canSort = c.vkind !== 'run';
      const items = [
        canSort ? ['sort-asc', 'Sort ' + words[0]] : null, canSort ? ['sort-desc', 'Sort ' + words[1]] : null,
        canSort && others ? ['then-asc', 'Then sort ' + words[0]] : null, canSort && others ? ['then-desc', 'Then sort ' + words[1]] : null,
        inSort && sorts.length > 1 ? ['sort-remove', 'Remove from sort'] : null,
        sorts.length ? ['sort-clear', sorts.length > 1 ? 'Clear all sorting' : 'Clear sort'] : null
      ];
      const gk = this.groupKeys().filter(k => this.view.cols.some(x => x.key === k));
      const gl = gk.indexOf(c.key);
      if (edit) {
        items.push('-',
          i > 0 ? ['first', 'Move to first'] : null, i > 0 ? ['left', 'Move left'] : null,
          i < last ? ['right', 'Move right'] : null, i < last ? ['last', 'Move to last'] : null,
          ['hide', 'Hide column'], '-',
          ['freeze', 'Freeze up to here'], this.view.nFrozen ? ['unfreeze', 'Unfreeze columns'] : null,
          !c.isNum && gl < 0 && gk.length < MAX_LEVELS ? ['group', gk.length ? 'Add subtotal level ' + (gk.length + 1) : 'Subtotal by this column'] : null,
          gl >= 0 ? ['ungroup', gk.length > 1 ? 'Remove subtotal level ' + (gl + 1) : 'Remove subtotals'] : null,
          gl >= 0 ? ['collapse-all', 'Collapse all'] : null,
          gl >= 0 ? ['expand-all', 'Expand all'] : null, '-',
          ['rename', 'Rename…'], lay.labels && lay.labels[c.key] ? ['unrename', 'Reset name'] : null);
        // Shares only make sense for columns that add up, so ratio columns don't offer it.
        // Viewers can hide formatting they find distracting; it's saved with their layout.
        if (c.isNum && this.cfg.cf[c.key] && this.cfg.cf[c.key].type) {
          const off = (lay.cfOff || []).includes(c.key);
          items.push('-', [off ? 'cf-show' : 'cf-hide', off ? 'Show formatting' : 'Hide formatting']);
        }
        if (this.authoring && !c.virtual) {
          const cur = this.cfg.colGroups[c.key];
          items.push('-', ...this.groupNames().filter(n => n !== cur).slice(0, 8).map(n => ['g-add:' + n, 'Add to group "' + n + '"']),
            ['g-new', 'New group…'], cur ? ['g-out', 'Remove from group "' + cur + '"'] : null);
        }
        const hasPct = (lay.pct || []).includes(c.key);
        if (c.virtual) items.push('-', ['calc-remove', c.vkind === 'pct' ? 'Remove % of total column' : 'Remove this column']);
        else if (c.isNum) {
          // Shares and running totals only make sense for columns that add up; ranks and comparisons work for any number.
          const calc = lay.calc || [];
          const has = d => calc.some(x => calcKey(x) === calcKey(d));
          const sums = c.total && c.total.type === 'sum';
          items.push('-',
            hasPct || sums ? [hasPct ? 'pct-remove' : 'pct-add', hasPct ? 'Remove % of total column' : 'Add % of total column'] : null,
            !has({ kind: 'rank', src: c.key, dir: 'desc' }) ? ['rank-desc', 'Add rank (largest = 1)'] : null,
            !has({ kind: 'rank', src: c.key, dir: 'asc' }) ? ['rank-asc', 'Add rank (smallest = 1)'] : null,
            sums && !has({ kind: 'run', src: c.key }) ? ['run', 'Add running total'] : null,
            ['compare', 'Compare with another column…'],
            '-', ['topn', 'Show top or bottom…'],
            lay.topN && lay.topN.key === c.key ? ['topn-off', 'Show all rows'] : null);
        }
      }
      items.push('-', ['stats', 'Column stats…']);
      const infos = [];
      if (c.isNum && c.total) infos.push(['info', 'Total: ' + c.total.text + (c.total.problem ? ' (' + c.total.problem + ')' : ''), true]);
      const cfText = this.cfDescribe(c).trim();
      if (cfText) infos.push(['info', cfText, true]);
      // Field and parameter names are how the dashboard is built; only authors see them.
      if (this.authoring && !c.virtual && c.label !== c.name) infos.push(['info', 'Field: ' + c.name, true]);
      if (infos.length) items.push('-', ...infos);
      if (this.authoring) items.push(['format', 'Format table…']);
      const html = '<div class="tp-menu" role="menu">' + items.filter(Boolean).map(it => it === '-' ? '<hr>' :
        `<button type="button" role="menuitem" data-a="${it[0]}"${it[2] ? ' disabled' : ''}>${esc(it[1])}</button>`).join('') + '</div>';
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        const b = e.target.closest('button[data-a]');
        if (!b) return;
        this.closePop();
        if (b.dataset.a === 'stats') return this.openStats(i, anchor);
        if (b.dataset.a === 'rename') {
          return this.askName(anchor, 'Column name', c.label, name => this.updateLayout({ labels: Object.assign({}, this.layout.labels, { [c.key]: name }) }));
        }
        if (b.dataset.a === 'g-new') return this.askName(anchor, 'Group name (shown above this column)', '', name => this.setGroups(g => { g[c.key] = name; }));
        if (b.dataset.a === 'compare') return this.openCompare(i, anchor);
        if (b.dataset.a === 'topn') return this.openTopN(i, anchor);
        this.menuAction(b.dataset.a, i, c);
      }));
    }

    // ---------- column stats ----------

    // Summary of one column over the rows passing the current filters; data rows only, never totals or subtotals.
    columnStats(c) {
      const st = this.store;
      const idx = this.view.idx;
      const n = idx.length;
      const lines = [['Rows', n === st.n ? n.toLocaleString() : n.toLocaleString() + ' of ' + st.n.toLocaleString() + ' (filtered)']];
      if (!st.complete) lines.push(['Note', 'Still loading rows, so these numbers will change']);
      const textCols = this.view.cols.filter(x => !x.isNum && x !== c).slice(0, 2);
      const rowName = i => { const s = textCols.map(x => st.fmt[x.key][i]).filter(v => v != null).join(' · '); return s ? ', ' + s : ''; };
      const count = k => k.toLocaleString();
      if (c.isNum) {
        const raw = st.raw[c.key];
        const vals = [];
        let sum = 0, lo = -1, hi = -1;
        for (const i of idx) {
          const v = raw[i];
          if (v !== v) continue;
          vals.push(v); sum += v;
          if (lo < 0 || v < raw[lo]) lo = i;
          if (hi < 0 || v > raw[hi]) hi = i;
        }
        const f = v => R.formatNumber(v, c.spec);
        lines.push(['With a value', count(vals.length)], ['Blank', count(n - vals.length)]);
        if (vals.length) {
          vals.sort((a, b) => a - b);
          const m = vals.length >> 1;
          const med = vals.length % 2 ? vals[m] : (vals[m - 1] + vals[m]) / 2;
          // A sum of ratios or averages means nothing, so Sum only shows for columns that total by summing.
          if (!c.total || c.total.type === 'sum') lines.push(['Sum', f(sum)]);
          lines.push(['Average', f(sum / vals.length)], ['Median', f(med)],
            ['Min', f(raw[lo]) + rowName(lo)], ['Max', f(raw[hi]) + rowName(hi)]);
          const t = this.view.totals[c.key];
          if (c.total && t != null) lines.push(['Table total', f(t) + ' (' + c.total.text + ')']);
        }
      } else if (c.isDate || st.dates[c.key]) {
        const arr = c.isDate ? st.raw[c.key] : st.dates[c.key];
        const fmt = st.fmt[c.key];
        let lo = -1, hi = -1, blank = 0;
        const days = new Set();
        for (const i of idx) {
          const v = arr[i];
          if (v == null || v !== v) { blank++; continue; }
          days.add(Math.floor(v / 86400000));
          if (lo < 0 || v < arr[lo]) lo = i;
          if (hi < 0 || v > arr[hi]) hi = i;
        }
        if (lo >= 0) lines.push(['Earliest', fmt[lo]], ['Latest', fmt[hi]]);
        lines.push(['Distinct days', count(days.size)], ['Blank', count(blank)]);
      } else {
        const fmt = st.fmt[c.key];
        const counts = new Map();
        let blank = 0;
        for (const i of idx) {
          const v = fmt[i];
          if (v == null || v === '') { blank++; continue; }
          counts.set(v, (counts.get(v) || 0) + 1);
        }
        lines.push(['Distinct values', count(counts.size)], ['Blank', count(blank)]);
        // Only worth showing when some value repeats; otherwise it would list five arbitrary names.
        const top = Array.from(counts).filter(x => x[1] > 1).sort((a, b) => b[1] - a[1] || collator.compare(a[0], b[0])).slice(0, 5);
        top.forEach(([v, k], j) => lines.push([j ? '' : 'Most common', v + ' (' + count(k) + ' rows)']));
      }
      return lines;
    }

    openTopN(i, anchor) {
      const c = this.view.cols[i];
      const cur = this.layout.topN && this.layout.topN.key === c.key ? this.layout.topN : null;
      const btn = (dir, n) => `<button type="button" data-dir="${dir}" data-n="${n}"${cur && cur.dir === dir && cur.n === n ? ' aria-pressed="true"' : ''}>${n}</button>`;
      const html = `<div class="tp-menu tp-topn"><div class="tp-ph">Show by ${esc(c.label)}</div>` +
        `<div class="tp-tnrow"><span>Top</span>${[5, 10, 20, 50].map(n => btn('top', n)).join('')}</div>` +
        `<div class="tp-tnrow"><span>Bottom</span>${[5, 10, 20, 50].map(n => btn('bottom', n)).join('')}</div>` +
        `<div class="tp-lname"><input type="number" min="1" step="1" placeholder="Other number" aria-label="Number of rows"><button type="button" data-custom="top">Top</button><button type="button" data-custom="bottom">Bottom</button></div>` +
        `<div class="tp-lnote">The rest of the rows${this.groups ? ' in each group' : ''} are combined into an Other row. Totals still include every row.</div></div>`;
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        const b = e.target.closest('button');
        if (!b) return;
        let dir = b.dataset.dir, n = +b.dataset.n;
        if (b.dataset.custom) { dir = b.dataset.custom; n = Math.floor(+pop.querySelector('input').value); }
        if (!dir || !(n >= 1)) return;
        this.closePop();
        // Turning it on also sorts by the same column, largest (or smallest) first, so the rows read as a ranking.
        // Any other sort can be chosen afterwards.
        this.updateLayout({ topN: { key: c.key, n: Math.min(n, 100000), dir }, sort: [{ key: c.key, dir: dir === 'bottom' ? 'asc' : 'desc' }] });
      }));
    }

    openCompare(i, anchor) {
      const c = this.view.cols[i];
      const others = this.store.columns.filter(x => x.isNum && !x.virtual && x.key !== c.key);
      if (!others.length) return this.flash('There is no other number column to compare with.');
      const html = `<div class="tp-menu tp-picker tp-compare"><div class="tp-ph">Compare ${esc(c.label)} with</div>` +
        `<select aria-label="Column to compare with">${others.map(x => `<option value="${esc(x.key)}">${esc(this.labelFor(x))}</option>`).join('')}</select>` +
        '<label><input type="radio" name="tp-cm" value="diff" checked> Difference (this − that)</label>' +
        '<label><input type="radio" name="tp-cm" value="pdiff"> % difference ((this − that) ÷ that)</label>' +
        '<hr><button type="button" data-a="add">Add column</button></div>';
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        if (!e.target.closest('button[data-a="add"]')) return;
        const other = pop.querySelector('select').value;
        const kind = pop.querySelector('input[name="tp-cm"]:checked').value;
        this.closePop();
        const d = { kind, src: c.key, other };
        const calc = (this.layout.calc || []).filter(x => calcKey(x) !== calcKey(d));
        this.updateLayout({ calc: calc.concat(d) });
      }));
    }

    openStats(i, anchor) {
      const c = this.view.cols[i];
      if (!c || !this.store) return;
      const lines = this.columnStats(c);
      const html = `<div class="tp-menu tp-stats"><div class="tp-ph">${esc(c.label)}</div><table>` +
        lines.map(([k, v]) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`).join('') +
        '</table><hr><button type="button" data-a="copy">Copy stats</button></div>';
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        if (!e.target.closest('button[data-a="copy"]')) return;
        this.closePop();
        this.copyToClipboard([c.label].concat(lines.map(([k, v]) => k + '\t' + v)).join('\n'));
      }));
    }

    menuAction(a, i, c) {
      const cols = this.view.cols;
      const lay = this.layout;
      switch (a) {
        case 'sort-asc': return this.setSort([{ key: c.key, dir: 'asc' }]);
        case 'sort-desc': return this.setSort([{ key: c.key, dir: 'desc' }]);
        case 'then-asc':
        case 'then-desc': {
          const dir = a === 'then-asc' ? 'asc' : 'desc';
          const next = this.sortList().slice();
          const k = next.findIndex(s => s.key === c.key);
          if (k >= 0) next[k] = { key: c.key, dir }; else next.push({ key: c.key, dir });
          return this.setSort(next);
        }
        case 'sort-remove': return this.setSort(this.sortList().filter(s => s.key !== c.key));
        case 'sort-clear': return this.setSort(null);
        case 'first': return this.moveColumn(i, 0);
        case 'left': return this.moveColumn(i, i - 1);
        case 'right': return this.moveColumn(i, i + 1);
        case 'last': return this.moveColumn(i, cols.length - 1);
        case 'hide': return this.updateLayout({ hidden: (lay.hidden || []).concat(c.key) });
        case 'freeze': return this.updateLayout({ freeze: i + 1 });
        case 'unfreeze': return this.updateLayout({ freeze: 0 });
        case 'group': return this.updateLayout({ groupBy: this.groupKeys().filter(k => cols.some(x => x.key === k)).concat(c.key).slice(0, MAX_LEVELS) });
        case 'ungroup': { const g = this.groupKeys().filter(k => k !== c.key); return this.updateLayout({ groupBy: g.length ? g : null }); }
        // Collapse all on a level shows that level's subtotals (opening the levels above it);
        // Expand all opens that level and everything above it.
        case 'collapse-all':
        case 'expand-all': {
          const level = (this.groupCols || []).findIndex(g => g.key === c.key);
          if (level < 0) return;
          for (let l = 0; l <= level; l++) {
            for (const g of this.groupsAtLevel(l)) {
              if (a === 'collapse-all' && l === level) this.collapsed.add(g.key); else this.collapsed.delete(g.key);
            }
          }
          this.computeView();
          this.renderBody(true);
          return;
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
        case 'cf-hide': return this.updateLayout({ cfOff: (lay.cfOff || []).concat(c.key) });
        case 'cf-show': return this.updateLayout({ cfOff: (lay.cfOff || []).filter(k => k !== c.key) });
        case 'g-out': return this.setGroups(g => { delete g[c.key]; });
        case 'rank-desc':
        case 'rank-asc':
        case 'run': {
          const d = a === 'run' ? { kind: 'run', src: c.key } : { kind: 'rank', src: c.key, dir: a === 'rank-asc' ? 'asc' : 'desc' };
          return this.updateLayout({ calc: (lay.calc || []).concat(d) });
        }
        case 'calc-remove':
          if (c.vkind === 'pct') return this.updateLayout({ pct: (lay.pct || []).filter(k => k !== c.srcKey) });
          return this.updateLayout({ calc: (lay.calc || []).filter(d => calcKey(d) !== c.key) });
        case 'topn-off': return this.updateLayout({ topN: null });
        case 'pct-add': return this.updateLayout({ pct: (lay.pct || []).concat(c.key) });
        case 'pct-remove': {
          const src = c.virtual ? c.srcKey : c.key;
          return this.updateLayout({ pct: (lay.pct || []).filter(k => k !== src) });
        }
        default:
          if (a.startsWith('g-add:')) { const n = a.slice(6); return this.setGroups(g => { g[c.key] = n; }); }
      }
    }

    // A name box inside the menu; browser pop-up prompts are often blocked inside Tableau.
    askName(anchor, title, initial, onOk) {
      const html = `<div class="tp-menu tp-ask"><div class="tp-ph">${esc(title)}</div><div class="tp-lname">` +
        `<input type="text" maxlength="80" value="${esc(initial || '')}" aria-label="${esc(title)}"><button type="button" data-a="ok">OK</button></div></div>`;
      this.openPop(anchor, html, pop => {
        const input = pop.querySelector('input');
        const done = () => { const v = input.value.trim(); this.closePop(); if (v) onOk(v); };
        input.focus();
        input.select();
        input.addEventListener('keydown', e => {
          if (e.key === 'Enter') { e.preventDefault(); done(); }
          else if (e.key === 'Escape') { e.preventDefault(); this.closePop(); }
        });
        pop.querySelector('button[data-a="ok"]').addEventListener('click', done);
      });
    }

    openGroupMenu(start, anchor) {
      const sp = (this.view.spans || []).find(x => x.start === start);
      if (!sp || !sp.label) return;
      // Moves act on the whole group, even where the freeze line splits its header.
      const run = this.groupRun(sp.start, sp.end);
      const s = { label: sp.label, frozen: sp.frozen, start: run.start, end: run.end };
      const items = [];
      const n = this.view.cols.length;
      const okGap = g => this.dropOk(g, s.start, s.end);
      const gaps = [];
      for (let g = 0; g <= n; g++) if (okGap(g) && (g < s.start || g > s.end + 1)) gaps.push(g);
      const first = s.frozen ? 0 : this.view.nFrozen;
      const moves = {
        'mv-start': gaps.find(g => g >= first && g < s.start),
        'mv-left': gaps.filter(g => g >= first && g < s.start).pop(),
        'mv-right': gaps.find(g => g > s.end + 1),
        'mv-end': s.end < n - 1 ? n : undefined
      };
      if (this.canEditLayout()) {
        items.push(moves['mv-start'] != null ? ['mv-start', 'Move group to start'] : null, moves['mv-left'] != null ? ['mv-left', 'Move group left'] : null,
          moves['mv-right'] != null ? ['mv-right', 'Move group right'] : null, moves['mv-end'] != null ? ['mv-end', 'Move group to end'] : null, '-');
        items.push(['hide', 'Hide group headers']);
      }
      if (this.authoring) items.push('-', ['rename', 'Rename group…'], ['remove', 'Remove group "' + s.label + '"']);
      if (!items.length) items.push(['none', s.label, true]);
      const html = '<div class="tp-menu" role="menu">' + items.filter(Boolean).map(it => it === '-' ? '<hr>' :
        `<button type="button" role="menuitem" data-a="${it[0]}"${it[2] ? ' disabled' : ''}>${esc(it[1])}</button>`).join('') + '</div>';
      this.openPop(anchor, html, pop => pop.addEventListener('click', e => {
        const b = e.target.closest('button[data-a]');
        if (!b) return;
        this.closePop();
        const a = b.dataset.a;
        if (a in moves) return this.moveBlock(s.start, s.end, moves[a]);
        if (a === 'hide') return this.updateLayout({ hideGroups: true });
        if (a === 'remove') return this.setGroups(g => { Object.keys(g).forEach(k => { if (g[k] === s.label) delete g[k]; }); });
        if (a === 'rename') {
          return this.askName(anchor, 'Group name', s.label, name => this.setGroups(g => { Object.keys(g).forEach(k => { if (g[k] === s.label) g[k] = name; }); }));
        }
      }));
    }

    openColumnsPicker(anchor) {
      const hidden = new Set(this.layout.hidden || []);
      const keys = this.orderedKeys();
      // Columns hidden because their parameter is set to "(none)" aren't listed for viewers at all.
      const html = '<div class="tp-menu tp-picker"><div class="tp-ph">Show and hide columns</div>' + keys.filter(k => this.authoring || !this.autoHidden(this.colDef(k))).map(k => {
        const c = this.colDef(k);
        const auto = this.autoHidden(c);
        return `<label${auto ? ' class="muted" title="Hidden by a parameter setting"' : ''}><input type="checkbox" data-k="${esc(k)}"${!hidden.has(k) ? ' checked' : ''}${auto ? ' disabled' : ''}> ${esc(this.labelFor(c))}</label>`;
      }).join('') + (this.groupNames().length ? `<hr><label><input type="checkbox" data-g="1"${this.layout.hideGroups ? '' : ' checked'}> Show group headers</label>` : '') +
        '<hr><button type="button" data-a="all">Show all</button></div>';
      this.openPop(anchor, html, pop => {
        pop.addEventListener('change', e => {
          if (e.target.dataset.g) return this.updateLayout({ hideGroups: !e.target.checked });
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
        columns: this.store.columns.filter(c => !c.virtual).map(c => ({ key: c.key, name: c.name, isNum: c.isNum, dataType: c.dataType, isUrl: !!c.isUrl }))
      });
      const url = new URL('format.html' + window.location.search, window.location.href).href;
      return tableau.extensions.ui.displayDialogAsync(url, payload, { width: 1100, height: 760 })
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
        const lk = this.linkFor(c);
        const url = lk && this.safeUrl(st.fmt[lk.urlKey][i]);
        // Excel gets a real hyperlink; CSV keeps the plain address in URL columns and the plain text elsewhere.
        if (url && (f != null || lk.label)) return { kind: 'link', v: lk.label || f, url, csv: f };
        return f == null ? { kind: 'null' } : { kind: 'text', v: f };
      };
      // Every row is exported whatever is collapsed on screen; outline levels and hidden flags let Excel
      // start with the same groups collapsed. With Top N, the top rows come first, then the Other row with the
      // rest grouped under it (closed unless it is open on screen).
      let rows;
      const pushTop = (top, level, hid) => {
        for (const i of top.shown) rows.push({ type: 'data', i, level, hidden: hid });
        const n = top.node;
        const open = this.otherOpen.has(n.key);
        rows.push({ type: 'group', label: 'Other (' + n.rows.length.toLocaleString() + (n.rows.length === 1 ? ' row)' : ' rows)'), totals: Object.assign({ __label: n.colKey }, n.totals), collapsed: !open, level, hidden: hid });
        for (const i of n.rows) rows.push({ type: 'data', i, level: level + 1, hidden: hid || !open });
      };
      if (this.groups) {
        rows = [];
        const walk = (list, hiddenAbove) => {
          for (const g of list) {
            const collapsed = this.collapsed.has(g.key);
            rows.push({ type: 'group', label: g.label + ' subtotal', totals: Object.assign({ __label: g.colKey }, g.totals), collapsed, level: g.level, hidden: hiddenAbove });
            const hid = hiddenAbove || collapsed;
            if (g.children) walk(g.children, hid);
            else if (g.top && g.top.node) pushTop(g.top, g.level + 1, hid);
            else for (const i of g.rows) rows.push({ type: 'data', i, level: g.level + 1, hidden: hid });
          }
        };
        walk(this.groups, false);
      } else if (this.topRoot && this.topRoot.node) {
        rows = [];
        pushTop(this.topRoot, 0, false);
      } else {
        rows = this.view.idx.map(i => ({ type: 'data', i }));
      }
      return {
        sheetName: this.ws.name,
        columns: cols.map(c => { const o = (!c.virtual && this.cfg.colFormat[c.key]) || {}; return { key: c.key, label: c.label, spec: c.spec, excelFmt: c.excelFmt, width: c.w, isNum: c.isNum, align: o.align || null, bold: !!o.bold }; }),
        style: this.cfg.style,
        rows,
        value,
        totals,
        totalsLabelKey: totals.__label,
        includeTotals: includeTotals && this.cfg.totals.show,
        includeSubtotals: this.cfg.export.subtotals,
        totalsLocation: this.cfg.totals.location,
        frozenCols: nFrozen,
        groupHeaders: this.view.spans ? this.view.spans.reduce((out, s) => {
          const last = out[out.length - 1];
          if (last && last.label === s.label && last.end + 1 === s.start) last.end = s.end;
          else out.push({ label: s.label, start: s.start, end: s.end });
          return out;
        }, []).filter(s => s.label) : null,
        format: (sc, v, isTotal) => this.cellFormat(this.colDef(sc.key), v, isTotal),
        barColumns: cols.filter(c => c.isNum && this.cfg.cf[c.key] && this.cfg.cf[c.key].type === 'bars').map(c => ({ key: c.key, totals: !!this.cfg.cf[c.key].totals }))
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

  function fieldOverlap(a, b) {
    const s = new Set(a);
    const shared = b.filter(k => s.has(k)).length;
    return shared / Math.max(a.length, b.length, 1);
  }

  // Keeps only well-formed layout settings, so a damaged or hand-edited layout can't break the table.
  function cleanLayout(l) {
    if (!l || typeof l !== 'object' || Array.isArray(l)) return {};
    const out = {};
    const strs = v => Array.isArray(v) ? v.filter(x => typeof x === 'string') : null;
    const map = (v, ok) => {
      if (!v || typeof v !== 'object' || Array.isArray(v)) return null;
      const o = {};
      for (const k of Object.keys(v)) if (ok(v[k])) o[k] = v[k];
      return o;
    };
    for (const k of ['order', 'hidden', 'pct', 'cfOff']) if (k in l) { const v = strs(l[k]); if (v) out[k] = v; }
    if ('widths' in l) { const v = map(l.widths, x => typeof x === 'number' && isFinite(x) && x > 0 && x < 5000); if (v) out.widths = v; }
    if ('labels' in l) { const v = map(l.labels, x => typeof x === 'string'); if (v) out.labels = v; }
    if (l.freeze === 'auto') out.freeze = 'auto';
    else if (typeof l.freeze === 'number' && isFinite(l.freeze) && l.freeze >= 0) out.freeze = Math.floor(l.freeze);
    const okSort = x => x && typeof x.key === 'string' && (x.dir === 'asc' || x.dir === 'desc');
    if ('sort' in l) {
      if (l.sort === null) out.sort = null;
      else if (Array.isArray(l.sort)) out.sort = l.sort.filter(okSort).map(x => ({ key: x.key, dir: x.dir }));
      else if (okSort(l.sort)) out.sort = [{ key: l.sort.key, dir: l.sort.dir }];
    }
    if ('groupBy' in l) {
      if (l.groupBy === null) out.groupBy = null;
      else if (typeof l.groupBy === 'string') out.groupBy = [l.groupBy];
      else { const v = strs(l.groupBy); if (v) out.groupBy = v.slice(0, MAX_LEVELS); }
    }
    if (Array.isArray(l.calc)) {
      out.calc = l.calc.filter(d => d && typeof d.src === 'string' && (
        (d.kind === 'rank' && (d.dir === undefined || d.dir === 'asc' || d.dir === 'desc')) || d.kind === 'run' ||
        ((d.kind === 'diff' || d.kind === 'pdiff') && typeof d.other === 'string')
      )).map(d => { const o = { kind: d.kind, src: d.src }; if (d.dir) o.dir = d.dir; if (d.other) o.other = d.other; return o; });
    }
    if ('hideGroups' in l) out.hideGroups = !!l.hideGroups;
    if (l.topN === null) out.topN = null;
    else if (l.topN && typeof l.topN.key === 'string' && (l.topN.dir === 'top' || l.topN.dir === 'bottom') && Number.isInteger(l.topN.n) && l.topN.n > 0) {
      out.topN = { key: l.topN.key, n: Math.min(l.topN.n, 100000), dir: l.topN.dir };
    }
    return out;
  }

  function hashStr(t) {
    let h = 0;
    for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) | 0;
    return (h >>> 0).toString(36);
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
    lay: '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M4 1.5h8A1.5 1.5 0 0 1 13.5 3v11.2a.6.6 0 0 1-.95.49L8 11.5l-4.55 3.19a.6.6 0 0 1-.95-.49V3A1.5 1.5 0 0 1 4 1.5Zm0 1.5v9.9l3.57-2.5a.75.75 0 0 1 .86 0L12 12.9V3Z"/></svg>',
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
