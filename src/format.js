/* global tableau, TPRules, TPExport, TPMergeConfig, TPDefaults */
(function () {
  'use strict';

  const R = window.TPRules;
  const $ = id => document.getElementById(id);
  let cfg, cols, params, sheet;

  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

  function showTab(name) {
    document.querySelectorAll('[role="tab"]').forEach(t => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
    document.querySelectorAll('[role="tabpanel"]').forEach(p => { p.hidden = p.id !== 'tab-' + name; });
    if (name === 'columns') renderColumns();
    if (name === 'rules') renderRules();
    if (name === 'formatting') renderFormatting();
    if (name === 'style') renderStyle();
  }

  // ---------- display + download ----------

  function fillSimple() {
    $('showFilters').checked = cfg.showFilters;
    $('filterMode').value = cfg.filterMode;
    $('showSearch').checked = cfg.showSearch;
    $('showTooltips').checked = cfg.showTooltips;
    $('allowViewerLayout').checked = cfg.allowViewerLayout;
    $('viewerWeights').checked = !!cfg.viewerWeights;
    $('totalsShow').checked = cfg.totals.show;
    $('totalsLocation').value = cfg.totals.location;
    // "Freeze up to here" can freeze more columns than the list offers; add that choice so Save keeps it.
    const fz = String(cfg.layout.freeze);
    if (!Array.from($('freeze').options).some(o => o.value === fz)) $('freeze').add(new Option('First ' + fz + ' columns', fz));
    $('freeze').value = fz;
    $('autoHideNone').checked = cfg.autoHideNone;
    $('lastValues').value = cfg.lastRows.values || '';
    $('lastAlways').checked = !!cfg.lastRows.always;
    $('nullString').value = cfg.nullText.string;
    $('nullNumber').value = cfg.nullText.number;
    $('nullDate').value = cfg.nullText.date;
    document.querySelectorAll('input[name="formats"]').forEach(r => { r.checked = r.value === cfg.export.formats; });
    $('excelTotals').checked = cfg.export.excelTotals;
    $('subtotals').checked = cfg.export.subtotals;
    $('filename').value = cfg.export.filename;
    preview();
  }

  function readSimple() {
    readStyle();
    cfg.showFilters = $('showFilters').checked;
    cfg.filterMode = $('filterMode').value;
    cfg.showSearch = $('showSearch').checked;
    cfg.showTooltips = $('showTooltips').checked;
    cfg.allowViewerLayout = $('allowViewerLayout').checked;
    cfg.viewerWeights = $('viewerWeights').checked;
    cfg.totals = { show: $('totalsShow').checked, location: $('totalsLocation').value };
    const fz = $('freeze').value;
    cfg.layout.freeze = fz === 'auto' ? 'auto' : Number(fz);
    cfg.autoHideNone = $('autoHideNone').checked;
    cfg.lastRows = { values: $('lastValues').value.trim(), always: $('lastAlways').checked };
    cfg.nullText = { string: $('nullString').value, number: $('nullNumber').value, date: $('nullDate').value };
    const f = document.querySelector('input[name="formats"]:checked');
    cfg.export = {
      formats: f ? f.value : 'excel',
      excelTotals: $('excelTotals').checked,
      subtotals: $('subtotals').checked,
      filename: $('filename').value.trim() || '{sheet} {date}'
    };
  }

  function preview() {
    const p = {};
    params.forEach(n => { p[n] = '…'; });
    $('filePreview').textContent = TPExport.buildFileName($('filename').value, { sheet, params: p }, 'xlsx');
  }

  // ---------- columns ----------

  function rulesDoc() { return cfg.rules && cfg.rules.doc; }

  function resolvesHtml(c) {
    if (!c.isNum) return '<span class="src">Text column, no total</span>';
    const t = R.resolveTotal(c, cols, rulesDoc(), cfg.totalOverrides[c.key]);
    const src = t.source === 'override' ? 'set here' : t.source === 'rule' ? t.label : 'default';
    const named = cols.map(x => ({ key: x.key, name: pickName(x) }));
    return esc(R.describeTotal(t, named)) + ' <span class="src">(' + esc(src) + ')</span>' + (t.problem || t.note ? '<div class="prob">' + esc(t.problem || t.note) + '</div>' : '');
  }

  // Number columns that share a name (SUM(Sales) and AVG(Sales) both show as Sales) are told apart by their field.
  function pickName(c) {
    return c.isNum && cols.filter(x => x.isNum && x.name.toLowerCase() === c.name.toLowerCase()).length > 1 ? c.key : c.name;
  }

  function numOptions(selected) {
    return '<option value="">Choose…</option>' + cols.filter(c => c.isNum).map(c => `<option value="${esc(c.key)}"${c.key === selected ? ' selected' : ''}>${esc(pickName(c))}</option>`).join('');
  }

  // Web-address columns are clickable by default; any text column can link to one of them instead.
  function linkCell(c, i) {
    const l = cfg.links[c.key] || {};
    if (c.isUrl) {
      return `<select data-i="${i}" data-k="urlMode" aria-label="Link display for ${esc(c.name)}">
          <option value="on"${l.mode !== 'off' ? ' selected' : ''}>Clickable</option>
          <option value="off"${l.mode === 'off' ? ' selected' : ''}>Not clickable</option>
        </select>` + (l.mode !== 'off' ? `<label class="linktext">Link text <input type="text" data-i="${i}" data-k="urlLabel" value="${esc(l.label || '')}" placeholder="Optional, e.g. View"></label><div class="hint">Blank shows the full address.</div>` : '');
    }
    if (c.isNum) return '<span class="muted">—</span>';
    const urlCols = cols.filter(x => x.isUrl);
    if (!urlCols.length) return '<span class="muted">No web-address fields</span>';
    return `<select data-i="${i}" data-k="linkTo" aria-label="Link ${esc(c.name)} to">
        <option value="">No link</option>
        ${urlCols.map(u => `<option value="${esc(u.key)}"${l.linkTo === u.key ? ' selected' : ''}>Link to ${esc(u.name)}</option>`).join('')}
      </select>`;
  }

  function renderColumns() {
    const hidden = new Set(cfg.layout.hidden || []);
    const labels = cfg.layout.labels || {};
    $('colRows').innerHTML = cols.map((c, i) => {
      const link = cfg.headerLinks[c.key];
      const linkSel = link === undefined ? '__auto__' : link === '' ? '__none__' : link;
      const autoParam = params.find(p => p.toLowerCase() === c.name.toLowerCase());
      const o = cfg.totalOverrides[c.key] || { type: 'auto' };
      const totalCell = c.isNum
        ? `<select data-i="${i}" data-k="total">${R.TOTAL_TYPES.map(t => `<option value="${t}"${o.type === t ? ' selected' : ''}>${{ auto: 'Auto', sum: 'Sum', avg: 'Average', min: 'Minimum', max: 'Maximum', ratio: 'Ratio (A ÷ B)', none: 'None' }[t]}</option>`).join('')}</select>` +
          (o.type === 'ratio' ? `<div class="ratio"><select data-i="${i}" data-k="numKey" aria-label="Numerator">${numOptions(o.numKey)}</select> ÷ <select data-i="${i}" data-k="denKey" aria-label="Denominator">${numOptions(o.denKey)}</select> × <input type="text" data-i="${i}" data-k="multiplier" value="${esc(o.multiplier || 1)}" aria-label="Multiplier"></div>` : '')
        : '<span class="muted">—</span>';
      const totalBlock = c.isNum ? totalCell + `<div class="resolves" id="res-${i}">${resolvesHtml(c)}</div>` : totalCell;
      return `<tr>
        <td class="name" title="${esc(c.key)}">${esc(c.name)}</td>
        <td><input type="text" data-i="${i}" data-k="label" value="${esc(labels[c.key] || '')}" placeholder="${esc(autoParam ? 'Parameter value' : c.name)}"></td>
        <td><input type="text" class="grp" data-i="${i}" data-k="group" list="groupList" value="${esc(cfg.colGroups[c.key] || '')}" placeholder="None" aria-label="Group for ${esc(c.name)}"></td>
        <td><select data-i="${i}" data-k="link">
          <option value="__auto__"${linkSel === '__auto__' ? ' selected' : ''}>Auto${autoParam ? ' (' + esc(autoParam) + ')' : ' (none found)'}</option>
          <option value="__none__"${linkSel === '__none__' ? ' selected' : ''}>Don't link</option>
          ${params.map(p => `<option value="${esc(p)}"${linkSel === p ? ' selected' : ''}>${esc(p)}</option>`).join('')}
        </select></td>
        <td>${totalBlock}</td>
        <td>${linkCell(c, i)}</td>
        <td>${hideWhenCell(c, i)}</td>
        <td><input type="checkbox" data-i="${i}" data-k="hidden"${hidden.has(c.key) ? ' checked' : ''} aria-label="Hide ${esc(c.name)} by default"></td>
      </tr>`;
    }).join('');
    fillGroupList();
  }

  function fillGroupList() {
    const names = Array.from(new Set(Object.values(cfg.colGroups).filter(Boolean)));
    $('groupList').innerHTML = names.map(n => `<option value="${esc(n)}"></option>`).join('');
  }

  // Hide the column while a parameter has one of the listed values (for example a True/False switch set to False).
  function hideWhenCell(c, i) {
    if (!params.length) return '<span class="muted">No parameters</span>';
    const r = cfg.hideWhen[c.key] || {};
    return `<select data-i="${i}" data-k="hwParam" aria-label="Parameter that hides ${esc(c.name)}"><option value="">Never</option>${params.map(p => `<option value="${esc(p)}"${r.param === p ? ' selected' : ''}>${esc(p)}</option>`).join('')}</select>` +
      (r.param ? `<label class="linktext">is <input type="text" data-i="${i}" data-k="hwValues" value="${esc(r.values || '')}" placeholder="e.g. False"></label><div class="hint">Separate several values with commas.</div>` : '');
  }

  function onColumnChange(e) {
    const el = e.target;
    if (!el.dataset || el.dataset.i == null) return;
    const c = cols[+el.dataset.i];
    const k = el.dataset.k;
    if (k === 'hwParam' || k === 'hwValues') {
      const r = Object.assign({}, cfg.hideWhen[c.key]);
      if (k === 'hwParam') { if (el.value) r.param = el.value; else { delete r.param; delete r.values; } }
      else r.values = el.value.trim();
      if (r.param) cfg.hideWhen[c.key] = r; else delete cfg.hideWhen[c.key];
      if (k === 'hwParam') renderColumns();
      return;
    }
    if (k === 'group') {
      cfg.colGroups = Object.assign({}, cfg.colGroups);
      if (el.value.trim()) cfg.colGroups[c.key] = el.value.trim(); else delete cfg.colGroups[c.key];
      if (e.type === 'change') fillGroupList();
      return;
    }
    if (k === 'label') {
      cfg.layout.labels = Object.assign({}, cfg.layout.labels);
      if (el.value.trim()) cfg.layout.labels[c.key] = el.value.trim(); else delete cfg.layout.labels[c.key];
      return;
    }
    if (k === 'link') {
      if (el.value === '__auto__') delete cfg.headerLinks[c.key];
      else cfg.headerLinks[c.key] = el.value === '__none__' ? '' : el.value;
      return;
    }
    if (k === 'urlMode' || k === 'urlLabel' || k === 'linkTo') {
      const l = Object.assign({}, cfg.links[c.key]);
      if (k === 'urlMode') { if (el.value === 'off') l.mode = 'off'; else delete l.mode; }
      if (k === 'urlLabel') { if (el.value.trim()) l.label = el.value.trim(); else delete l.label; }
      if (k === 'linkTo') { if (el.value) l.linkTo = el.value; else delete l.linkTo; }
      if (Object.keys(l).length) cfg.links[c.key] = l; else delete cfg.links[c.key];
      if (k === 'urlMode') renderColumns();
      return;
    }
    if (k === 'hidden') {
      const h = new Set(cfg.layout.hidden || []);
      if (el.checked) h.add(c.key); else h.delete(c.key);
      cfg.layout.hidden = Array.from(h);
      return;
    }
    const o = Object.assign({ type: 'auto' }, cfg.totalOverrides[c.key]);
    if (k === 'total') o.type = el.value;
    else if (k === 'multiplier') o.multiplier = Number(el.value) || 1;
    else o[k] = el.value;
    if (o.type === 'auto') delete cfg.totalOverrides[c.key]; else cfg.totalOverrides[c.key] = o;
    if (k === 'total') renderColumns();
    else $('res-' + el.dataset.i).innerHTML = resolvesHtml(c);
  }

  // ---------- conditional formatting ----------

  function cutInputs(i, which, def) {
    def = def || {};
    return `<input type="number" step="any" data-i="${i}" data-k="${which}Value" value="${esc(def.param ? '' : (def.value != null ? def.value : ''))}" aria-label="${which === 'low' ? 'Lower' : 'Upper'} threshold value"${def.param ? ' disabled' : ''}>` +
      (params.length ? ` or <select data-i="${i}" data-k="${which}Param" aria-label="${which === 'low' ? 'Lower' : 'Upper'} threshold parameter"><option value="">fixed value</option>${params.map(p => `<option value="${esc(p)}"${def.param === p ? ' selected' : ''}>${esc(p)}</option>`).join('')}</select>` : '');
  }

  function cfSettings(c, i, r) {
    const totals = `<label><input type="checkbox" data-i="${i}" data-k="totals"${r.totals ? ' checked' : ''}> Also color totals and subtotals</label>`;
    if (r.type === 'bars') return `<div class="cfset">${totals}</div>`;
    if (r.type === 'scale') {
      return `<div class="cfset"><label>Colors <select data-i="${i}" data-k="better">
          <option value="higher"${r.better !== 'lower' && r.better !== 'none' ? ' selected' : ''}>Higher is better</option>
          <option value="lower"${r.better === 'lower' ? ' selected' : ''}>Lower is better</option>
          <option value="none"${r.better === 'none' ? ' selected' : ''}>No good or bad (one color)</option>
        </select></label>${totals}</div>`;
    }
    const lower = r.better === 'lower';
    return `<div class="cfset">
        <label><select data-i="${i}" data-k="better" aria-label="Which direction is good">
          <option value="higher"${!lower ? ' selected' : ''}>Higher is better</option>
          <option value="lower"${lower ? ' selected' : ''}>Lower is better</option>
        </select></label>
        <label>${lower ? '✓ at or below' : '✕ below'} ${cutInputs(i, 'low', r.low)}</label>
        <label>${lower ? '✕ above' : '✓ at or above'} ${cutInputs(i, 'high', r.high)}</label>
        <label>Color the <select data-i="${i}" data-k="target"><option value="bg"${r.target !== 'text' ? ' selected' : ''}>background</option><option value="text"${r.target === 'text' ? ' selected' : ''}>text</option></select></label>
        <label><input type="checkbox" data-i="${i}" data-k="symbols"${r.symbols !== false ? ' checked' : ''}> Show ✓ ! ✕ symbols</label>
        ${totals}
      </div><div class="hint">With one value, only that side is colored (useful for flagging just the problems). With both, values in between get ! Watch. Values are in the units shown, so 5 means 5% on a percent column.</div>`;
  }

  function renderFormatting() {
    const nums = cols.map((c, i) => ({ c, i })).filter(x => x.c.isNum);
    $('cfRows').innerHTML = nums.map(({ c, i }) => {
      const r = cfg.cf[c.key] || {};
      return `<tr>
        <td class="name" title="${esc(c.key)}">${esc(c.name)}</td>
        <td><select data-i="${i}" data-k="type" aria-label="Formatting style for ${esc(c.name)}">
          <option value=""${!r.type ? ' selected' : ''}>None</option>
          <option value="threshold"${r.type === 'threshold' ? ' selected' : ''}>Thresholds</option>
          <option value="scale"${r.type === 'scale' ? ' selected' : ''}>Color scale</option>
          <option value="bars"${r.type === 'bars' ? ' selected' : ''}>Data bars</option>
        </select></td>
        <td>${r.type ? cfSettings(c, i, r) : '<span class="muted">—</span>'}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="3" class="muted">No number columns on this table.</td></tr>';
  }

  function onCfChange(e) {
    const el = e.target;
    if (!el.dataset || el.dataset.i == null) return;
    const c = cols[+el.dataset.i];
    const k = el.dataset.k;
    const r = Object.assign({}, cfg.cf[c.key]);
    let rerender = false;
    if (k === 'type') {
      if (!el.value) { delete cfg.cf[c.key]; renderFormatting(); return; }
      r.type = el.value;
      if (!r.better) r.better = 'higher';
      rerender = true;
    } else if (k === 'better') { r.better = el.value; rerender = true; }
    else if (k === 'target') r.target = el.value;
    else if (k === 'symbols' || k === 'totals') r[k] = el.checked;
    else if (/^(low|high)(Value|Param)$/.test(k)) {
      const which = k.startsWith('low') ? 'low' : 'high';
      const d = Object.assign({}, r[which]);
      if (k.endsWith('Param')) { if (el.value) { d.param = el.value; delete d.value; } else delete d.param; rerender = true; }
      else d.value = el.value === '' ? null : Number(el.value);
      r[which] = d;
    }
    cfg.cf[c.key] = r;
    if (rerender) renderFormatting();
  }

  // ---------- style ----------

  // Palettes picked for readable pairings; any other color can be typed as a #RRGGBB code.
  const PALETTES = {
    bg: ['#FFFFFF', '#F6F8FA', '#EEF1F5', '#E3EAF4', '#E6F2EC', '#F4EFE4', '#EDE9F5', '#1F2937', '#0B3D63', '#14532D'],
    text: ['#1F2328', '#000000', '#374151', '#4B5563', '#0B3D63', '#FFFFFF'],
    accent: ['#1F6FEB', '#0B5CAD', '#0F766E', '#6D28D9', '#B42318', '#9A3412', '#374151']
  };
  const STYLE_LABELS = { text: 'Text color', headBg: 'Header background', headText: 'Header text', totBg: 'Totals background', subBg: 'Subtotal background', accent: 'Accent color' };

  function lum(hex) {
    const c = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255).map(v => v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  }
  function contrast(a, b) {
    const x = lum(a), y = lum(b);
    return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
  }
  const isHex = v => /^#[0-9a-f]{6}$/i.test(v);

  // The pairings each color is read against; the weakest one is shown.
  function contrastFor(k) {
    const st = cfg.style;
    const pairs = {
      text: [[st.text, '#FFFFFF'], [st.text, st.totBg], [st.text, st.subBg]],
      headBg: [[st.headText, st.headBg]],
      headText: [[st.headText, st.headBg]],
      totBg: [[st.text, st.totBg]],
      subBg: [[st.text, st.subBg]],
      accent: [[st.accent, '#FFFFFF']]
    }[k];
    return Math.min(...pairs.filter(p => isHex(p[0]) && isHex(p[1])).map(p => contrast(p[0], p[1])));
  }

  function colorCtl(el) {
    const k = el.dataset.k;
    const cur = String(cfg.style[k] || '').toUpperCase();
    const ratio = contrastFor(k);
    const low = ratio < 4.5;
    el.innerHTML = PALETTES[el.dataset.pal].map(c => `<button type="button" class="swatch" data-c="${c}" style="background:${c}" aria-pressed="${c === cur}" aria-label="${STYLE_LABELS[k]} ${c}" title="${c}"></button>`).join('') +
      `<input type="text" value="${esc(cur)}" aria-label="${STYLE_LABELS[k]} code" spellcheck="false">` +
      `<span class="cnote${low ? ' low' : ''}">${isFinite(ratio) ? 'Contrast ' + ratio.toFixed(1) + ':1' + (low ? ': hard to read, aim for 4.5:1 or more' : '') : ''}</span>`;
  }

  function fmtSelect(i, k, val, opts, label) {
    return `<select data-i="${i}" data-k="${k}" aria-label="${esc(label)}">` + opts.map(([v, t]) => `<option value="${v}"${String(val || '') === v ? ' selected' : ''}>${t}</option>`).join('') + '</select>';
  }

  // The style controls are only read once they have been filled in, so an unopened tab never overwrites settings.
  let styleShown = false;

  function renderStyle() {
    styleShown = true;
    const st = cfg.style;
    $('stSize').value = st.size;
    $('stDensity').value = st.density;
    $('stHeadBold').checked = st.headBold;
    $('stHeadWrap').checked = st.headWrap;
    $('stLines').value = st.lines;
    $('stBanding').checked = st.banding;
    $('stTotBold').checked = st.totBold;
    $('stSubBold').checked = st.subBold;
    $('stNeg').value = st.neg || '';
    document.querySelectorAll('.colorctl').forEach(colorCtl);
    $('fmtRows').innerHTML = cols.map((c, i) => {
      const o = cfg.colFormat[c.key] || {};
      const num = c.isNum;
      return `<tr>
        <td class="name" title="${esc(c.key)}">${esc(c.name)}</td>
        <td>${fmtSelect(i, 'align', o.align, [['', 'Default'], ['l', 'Left'], ['c', 'Center'], ['r', 'Right']], 'Alignment for ' + c.name)}</td>
        <td><input type="checkbox" data-i="${i}" data-k="bold"${o.bold ? ' checked' : ''} aria-label="Bold ${esc(c.name)}"></td>
        <td>${num ? fmtSelect(i, 'kind', o.kind, [['', 'As Tableau'], ['number', 'Number'], ['currency', 'Currency'], ['percent', 'Percent']], 'Number style for ' + c.name) : '<span class="muted">—</span>'}</td>
        <td>${num ? fmtSelect(i, 'dec', o.dec, [['', 'As Tableau'], ['0', '0'], ['1', '1'], ['2', '2'], ['3', '3'], ['4', '4']], 'Decimals for ' + c.name) : '<span class="muted">—</span>'}</td>
        <td>${num ? fmtSelect(i, 'units', o.units, [['', 'As Tableau'], ['none', 'None'], ['K', 'K'], ['M', 'M'], ['B', 'B'], ['auto', 'Auto']], 'Units for ' + c.name) : '<span class="muted">—</span>'}</td>
        <td>${num ? fmtSelect(i, 'neg', o.neg, [['', 'Table default'], ['minus', '-1,234'], ['parens', '(1,234)'], ['red', '-1,234 red'], ['redparens', '(1,234) red']], 'Negatives for ' + c.name) : '<span class="muted">—</span>'}</td>
      </tr>`;
    }).join('');
  }

  function readStyle() {
    if (!styleShown) return;
    const st = cfg.style;
    st.size = $('stSize').value;
    st.density = $('stDensity').value;
    st.headBold = $('stHeadBold').checked;
    st.headWrap = $('stHeadWrap').checked;
    st.lines = $('stLines').value;
    st.banding = $('stBanding').checked;
    st.totBold = $('stTotBold').checked;
    st.subBold = $('stSubBold').checked;
    st.neg = $('stNeg').value;
  }

  function setColor(k, v) {
    if (!isHex(v)) return false;
    cfg.style[k] = v.toUpperCase();
    document.querySelectorAll('.colorctl').forEach(colorCtl);
    return true;
  }

  function onFmtChange(e) {
    const el = e.target;
    if (!el.dataset || el.dataset.i == null) return;
    const c = cols[+el.dataset.i];
    const k = el.dataset.k;
    const o = Object.assign({}, cfg.colFormat[c.key]);
    const v = el.type === 'checkbox' ? el.checked : el.value;
    if (v === '' || v === false) delete o[k]; else o[k] = v;
    if (Object.keys(o).length) cfg.colFormat[c.key] = o; else delete cfg.colFormat[c.key];
  }

  // ---------- rules ----------

  function describeRule(r) {
    const w = r.when;
    const cond = w.equals != null ? 'is "' + w.equals + '"' : w.startsWith != null ? 'starts with "' + w.startsWith + '"' : w.endsWith != null ? 'ends with "' + w.endsWith + '"' : 'contains "' + w.contains + '"';
    const t = r.total;
    const res = t.type === 'ratio' ? t.numerator + ' ÷ ' + t.denominator + (t.multiplier && t.multiplier !== 1 ? ' × ' + t.multiplier : '') : { sum: 'Sum', avg: 'Average', min: 'Minimum', max: 'Maximum', none: 'No total' }[t.type];
    return 'If the name ' + cond + ': ' + res;
  }

  function renderRules() {
    const r = cfg.rules;
    $('rulesError').hidden = true;
    if (!r || !r.doc) {
      $('rulesStatus').className = 'note warn';
      $('rulesStatus').textContent = 'No rules imported. Every number column sums unless you set its total on the Columns & totals tab.';
      $('rulesList').innerHTML = '';
      return;
    }
    $('rulesStatus').className = 'note';
    $('rulesStatus').textContent = r.doc.rules.length + ' rules' + (r.source ? ' from ' + r.source : '') + (r.importedAt ? ', imported ' + new Date(r.importedAt).toLocaleString() : '') + '. First match wins; anything unmatched uses ' + ((r.doc.default && r.doc.default.type) || 'sum') + '. Rule {rest} means the rest of the column name.';
    $('rulesList').innerHTML = r.doc.rules.map(x => '<li>' + esc(describeRule(x)) + '</li>').join('');
  }

  function useRules(text, source) {
    try {
      const doc = R.validateRules(JSON.parse(text));
      cfg.rules = { doc, source, importedAt: new Date().toISOString() };
      renderRules();
      $('msg').textContent = 'Rules loaded. Click Save to keep them in this workbook.';
    } catch (e) {
      $('rulesError').hidden = false;
      $('rulesError').textContent = 'Could not use these rules: ' + e.message;
    }
  }

  function downloadRules() {
    if (!cfg.rules) return;
    const blob = new Blob([JSON.stringify(cfg.rules.doc, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'totals-rules.json';
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  // ---------- save ----------

  async function save() {
    readSimple();
    $('save').disabled = true;
    $('msg').textContent = 'Saving…';
    try {
      tableau.extensions.settings.set('tp-config', JSON.stringify(cfg));
      await tableau.extensions.settings.saveAsync();
      tableau.extensions.ui.closeDialog('saved');
    } catch (e) {
      $('msg').textContent = 'Could not save: ' + ((e && e.message) || e);
      $('save').disabled = false;
    }
  }

  function wire() {
    document.querySelectorAll('[role="tab"]').forEach(t => t.addEventListener('click', () => { readSimple(); showTab(t.dataset.tab); }));
    $('colRows').addEventListener('change', onColumnChange);
    $('cfRows').addEventListener('change', onCfChange);
    $('cfRows').addEventListener('input', e => { if (/Value$/.test(e.target.dataset.k || '')) onCfChange(e); });
    $('colRows').addEventListener('input', e => { if (['label', 'group', 'multiplier', 'urlLabel', 'hwValues'].includes(e.target.dataset.k)) onColumnChange(e); });
    $('filename').addEventListener('input', preview);
    $('rulesFile').addEventListener('change', e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      f.text().then(t => useRules(t, f.name));
      e.target.value = '';
    });
    $('rulesApply').addEventListener('click', () => useRules($('rulesPaste').value, 'pasted text'));
    $('rulesExport').addEventListener('click', downloadRules);
    $('rulesClear').addEventListener('click', () => { cfg.rules = null; renderRules(); });
    $('tab-style').addEventListener('click', e => {
      const sw = e.target.closest('.swatch');
      if (sw) setColor(sw.closest('.colorctl').dataset.k, sw.dataset.c);
    });
    $('tab-style').addEventListener('change', e => {
      const ctl = e.target.closest('.colorctl');
      if (ctl && e.target.type === 'text') {
        const v = e.target.value.trim();
        if (!setColor(ctl.dataset.k, v.startsWith('#') ? v : '#' + v)) $('msg').textContent = 'Colors need a 6-digit code like #1F2328.';
      }
    });
    $('fmtRows').addEventListener('change', onFmtChange);
    $('stReset').addEventListener('click', () => { cfg.style = TPDefaults().style; renderStyle(); });
    $('save').addEventListener('click', save);
    $('cancel').addEventListener('click', () => tableau.extensions.ui.closeDialog(''));
  }

  tableau.extensions.initializeDialogAsync().then(payload => {
    const p = JSON.parse(payload || '{}');
    cfg = TPMergeConfig(p.config);
    cols = p.columns || [];
    params = p.params || [];
    sheet = p.sheet || '';
    wire();
    fillSimple();
    showTab('display');
  }).catch(e => { $('msg').textContent = 'Could not open: ' + ((e && e.message) || e); });
})();
