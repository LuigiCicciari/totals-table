/* global tableau, TPRules, TPExport, TPMergeConfig */
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
  }

  // ---------- display + download ----------

  function fillSimple() {
    $('showFilters').checked = cfg.showFilters;
    $('filterMode').value = cfg.filterMode;
    $('showSearch').checked = cfg.showSearch;
    $('showTooltips').checked = cfg.showTooltips;
    $('allowViewerLayout').checked = cfg.allowViewerLayout;
    $('totalsShow').checked = cfg.totals.show;
    $('totalsLocation').value = cfg.totals.location;
    $('freeze').value = String(cfg.layout.freeze);
    $('autoHideNone').checked = cfg.autoHideNone;
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
    cfg.showFilters = $('showFilters').checked;
    cfg.filterMode = $('filterMode').value;
    cfg.showSearch = $('showSearch').checked;
    cfg.showTooltips = $('showTooltips').checked;
    cfg.allowViewerLayout = $('allowViewerLayout').checked;
    cfg.totals = { show: $('totalsShow').checked, location: $('totalsLocation').value };
    const fz = $('freeze').value;
    cfg.layout.freeze = fz === 'auto' ? 'auto' : Number(fz);
    cfg.autoHideNone = $('autoHideNone').checked;
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
    return esc(R.describeTotal(t, cols)) + ' <span class="src">(' + esc(src) + ')</span>' + (t.problem ? '<div class="prob">' + esc(t.problem) + '</div>' : '');
  }

  function numOptions(selected) {
    return '<option value="">Choose…</option>' + cols.filter(c => c.isNum).map(c => `<option value="${esc(c.key)}"${c.key === selected ? ' selected' : ''}>${esc(c.name)}</option>`).join('');
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
      const linkSel = link === undefined ? 'auto' : link === '' ? 'none' : link;
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
        <td><select data-i="${i}" data-k="link">
          <option value="auto"${linkSel === 'auto' ? ' selected' : ''}>Auto${autoParam ? ' (' + esc(autoParam) + ')' : ' (none found)'}</option>
          <option value="none"${linkSel === 'none' ? ' selected' : ''}>Don't link</option>
          ${params.map(p => `<option value="${esc(p)}"${linkSel === p ? ' selected' : ''}>${esc(p)}</option>`).join('')}
        </select></td>
        <td>${totalBlock}</td>
        <td>${linkCell(c, i)}</td>
        <td><input type="checkbox" data-i="${i}" data-k="hidden"${hidden.has(c.key) ? ' checked' : ''} aria-label="Hide ${esc(c.name)} by default"></td>
      </tr>`;
    }).join('');
  }

  function onColumnChange(e) {
    const el = e.target;
    if (!el.dataset || el.dataset.i == null) return;
    const c = cols[+el.dataset.i];
    const k = el.dataset.k;
    if (k === 'label') {
      cfg.layout.labels = Object.assign({}, cfg.layout.labels);
      if (el.value.trim()) cfg.layout.labels[c.key] = el.value.trim(); else delete cfg.layout.labels[c.key];
      return;
    }
    if (k === 'link') {
      if (el.value === 'auto') delete cfg.headerLinks[c.key];
      else cfg.headerLinks[c.key] = el.value === 'none' ? '' : el.value;
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
    $('colRows').addEventListener('input', e => { if (['label', 'multiplier', 'urlLabel'].includes(e.target.dataset.k)) onColumnChange(e); });
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
