(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TPRules = factory();
})(this, function () {
  'use strict';

  const TOTAL_TYPES = ['auto', 'sum', 'avg', 'min', 'max', 'ratio', 'none'];

  // "AGG(Cost per Visit)" -> "Cost per Visit"; plain names pass through. Only Tableau's own aggregation and date
  // prefixes are removed, so a field named "Profit(Adj)" keeps its name.
  const AGG_PREFIX = /^(AGG|SUM|AVG|MIN|MAX|MEDIAN|COUNT|COUNTD|CNT|CNTD|ATTR|STDEV|STDEVP|VAR|VARP|PCT\d*|YEAR|QUARTER|MONTH|WEEK|DAY|WEEKDAY|HOUR|MINUTE|SECOND|MDY|MY|YR|QR|MN|DY|WK|TYR|TQR|TMN|TWK|TDY|USR|NONE)\((.*)\)$/i;
  function normName(fieldName) {
    const s = String(fieldName == null ? '' : fieldName).trim();
    const m = s.match(AGG_PREFIX);
    return (m ? m[2] : s).trim();
  }

  function eq(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }

  function nameVariants(n) {
    const s = String(n).trim();
    const out = [s, s + 's', s + 'es'];
    if (/ies$/i.test(s)) out.push(s.slice(0, -3) + 'y');
    if (/y$/i.test(s)) out.push(s.slice(0, -1) + 'ies');
    if (/es$/i.test(s)) out.push(s.slice(0, -2));
    if (/s$/i.test(s)) out.push(s.slice(0, -1));
    return out;
  }

  // Finds a column by display name, tolerating singular/plural differences.
  function findColumn(wanted, columns) {
    if (!wanted) return null;
    const exact = columns.find(c => eq(c.name, wanted) || eq(c.key, wanted));
    if (exact) return exact;
    for (const v of nameVariants(wanted)) {
      const hit = columns.find(c => eq(c.name, v));
      if (hit) return hit;
    }
    return null;
  }

  function validateRules(obj) {
    if (!obj || typeof obj !== 'object' || !Array.isArray(obj.rules)) throw new Error('Rules file must be JSON with a "rules" list.');
    obj.rules.forEach((r, i) => {
      const w = r && r.when;
      if (!w || !(w.equals || w.startsWith || w.endsWith || w.contains)) throw new Error('Rule ' + (i + 1) + ' needs "when" with equals, startsWith, endsWith or contains.');
      const t = r.total;
      if (!t || !TOTAL_TYPES.includes(t.type) || t.type === 'auto') throw new Error('Rule ' + (i + 1) + ' has an invalid total type.');
      if (t.type === 'ratio' && (!t.numerator || !t.denominator)) throw new Error('Rule ' + (i + 1) + ' is a ratio but is missing numerator or denominator.');
    });
    const d = obj.default;
    if (d != null && (typeof d !== 'object' || !['sum', 'avg', 'min', 'max', 'none'].includes(d.type))) throw new Error('"default" must have a type of sum, avg, min, max or none.');
    return obj;
  }

  function matchRule(name, rule) {
    const w = rule.when;
    const n = String(name).trim();
    const lower = n.toLowerCase();
    if (w.equals != null) return eq(n, w.equals) ? '' : null;
    if (w.startsWith != null) {
      const p = String(w.startsWith).toLowerCase();
      return lower.startsWith(p) && n.length > p.length ? n.slice(p.length).trim() : null;
    }
    if (w.endsWith != null) {
      const p = String(w.endsWith).toLowerCase();
      return lower.endsWith(p) ? n.slice(0, n.length - p.length).trim() : null;
    }
    if (w.contains != null) return lower.indexOf(String(w.contains).toLowerCase()) >= 0 ? '' : null;
    return null;
  }

  function fill(template, rest) {
    return String(template).replace(/\{rest\}/gi, rest || '').replace(/\s+/g, ' ').trim();
  }

  // Works out how one column totals: per-column override first, then the first matching rule, then the default.
  function resolveTotal(col, columns, rulesDoc, override) {
    if (!col.isNum) return { type: 'none', source: 'text' };
    const fromSpec = (t, source, label) => {
      if (t.type !== 'ratio') return { type: t.type, source, label };
      const num = findColumn(t.numerator, columns);
      const den = findColumn(t.denominator, columns);
      const out = { type: 'ratio', source, label, numKey: num && num.key, denKey: den && den.key, multiplier: Number(t.multiplier) || 1, numerator: t.numerator, denominator: t.denominator };
      if (!num || !den) out.problem = 'Needs ' + (!num ? '"' + t.numerator + '"' : '') + (!num && !den ? ' and ' : '') + (!den ? '"' + t.denominator + '"' : '') + ' on the Detail shelf';
      return out;
    };
    if (override && override.type && override.type !== 'auto') {
      if (override.type === 'ratio') {
        const num = columns.find(c => c.key === override.numKey);
        const den = columns.find(c => c.key === override.denKey);
        return fromSpec({ type: 'ratio', numerator: num ? num.name : override.numKey, denominator: den ? den.name : override.denKey, multiplier: override.multiplier }, 'override');
      }
      return { type: override.type, source: 'override' };
    }
    const rules = (rulesDoc && rulesDoc.rules) || [];
    for (let i = 0; i < rules.length; i++) {
      const rest = matchRule(col.name, rules[i]);
      if (rest === null) continue;
      const t = rules[i].total;
      const spec = t.type === 'ratio'
        ? { type: 'ratio', numerator: fill(t.numerator, rest), denominator: fill(t.denominator, rest), multiplier: t.multiplier }
        : t;
      return fromSpec(spec, 'rule', 'Rule ' + (i + 1));
    }
    const def = rulesDoc && rulesDoc.default && TOTAL_TYPES.includes(rulesDoc.default.type) ? rulesDoc.default : { type: 'sum' };
    return fromSpec(def, 'default');
  }

  function describeTotal(t, columns) {
    const nameOf = k => { const c = columns.find(x => x.key === k); return c ? c.name : k; };
    switch (t.type) {
      case 'sum': return 'Sum';
      case 'avg': return 'Average';
      case 'min': return 'Minimum';
      case 'max': return 'Maximum';
      case 'none': return 'No total';
      case 'ratio': {
        const base = (t.numKey ? nameOf(t.numKey) : t.numerator) + ' ÷ ' + (t.denKey ? nameOf(t.denKey) : t.denominator);
        return base + (t.multiplier && t.multiplier !== 1 ? ' × ' + t.multiplier : '');
      }
      default: return '';
    }
  }

  // raw(key) returns the column's numeric array (NaN for nulls); idx is the list of row indices to total.
  function computeTotal(t, colKey, idx, raw) {
    if (!t || t.type === 'none') return null;
    if (t.type === 'ratio') {
      if (!t.numKey || !t.denKey) return null;
      const a = raw(t.numKey), b = raw(t.denKey);
      let sa = 0, sb = 0;
      for (let k = 0; k < idx.length; k++) {
        const i = idx[k];
        const x = a[i], y = b[i];
        if (x === x) sa += x;
        if (y === y) sb += y;
      }
      if (!sb) return null;
      return (sa / sb) * (t.multiplier || 1);
    }
    const v = raw(colKey);
    let s = 0, n = 0, mn = Infinity, mx = -Infinity;
    for (let k = 0; k < idx.length; k++) {
      const x = v[idx[k]];
      if (x !== x) continue;
      s += x; n++;
      if (x < mn) mn = x;
      if (x > mx) mx = x;
    }
    if (!n) return null;
    switch (t.type) {
      case 'sum': return s;
      case 'avg': return s / n;
      case 'min': return mn;
      case 'max': return mx;
      default: return null;
    }
  }

  // ---- number formats ----

  const NUM_RE = /^(\()?\s*(-)?\s*([$€£¥])?\s*(-)?\s*(\d{1,3}(?:,\d{3})+|\d+)?(?:\.(\d+))?\s*([KMB])?\s*(%)?\s*(\))?$/;
  const SCALE = { K: 1e3, M: 1e6, B: 1e9 };

  function parseFormatted(fv) {
    if (typeof fv !== 'string') return null;
    const s = fv.trim();
    if (!s) return null;
    const m = s.match(NUM_RE);
    if (!m || (m[5] == null && m[6] == null)) return null;
    if (!!m[1] !== !!m[9]) return null;
    const intPart = (m[5] || '0').replace(/,/g, '');
    let v = Number(intPart + (m[6] ? '.' + m[6] : ''));
    const neg = !!(m[1] || m[2] || m[4]);
    if (m[7]) v *= SCALE[m[7]];
    if (m[8]) v /= 100;
    if (neg) v = -v;
    return {
      value: v,
      spec: { cur: m[3] || '', pct: !!m[8], dec: m[6] ? m[6].length : 0, group: !!(m[5] && m[5].indexOf(',') >= 0), scale: m[7] || '', neg: m[1] ? 'paren' : 'minus' }
    };
  }

  function closeEnough(a, b, spec) {
    if (!isFinite(a) || !isFinite(b)) return false;
    let unit = Math.pow(10, -spec.dec);
    if (spec.pct) unit /= 100;
    if (spec.scale) unit *= SCALE[spec.scale];
    return Math.abs(a - b) <= unit * 0.51 + 1e-9 * Math.max(1, Math.abs(b));
  }

  // Infers one display format per column from Tableau's formatted strings, checked against the raw numbers.
  function inferSpec(samples) {
    let best = null;
    let seen = 0;
    for (const s of samples) {
      if (seen >= 300) break;
      const p = parseFormatted(s.f);
      if (!p || !isFinite(s.v)) continue;
      if (!closeEnough(p.value, s.v, p.spec)) continue;
      seen++;
      if (!best) { best = Object.assign({}, p.spec); continue; }
      if (p.spec.cur !== best.cur || p.spec.pct !== best.pct || p.spec.scale !== best.scale) continue;
      best.dec = Math.max(best.dec, p.spec.dec);
      best.group = best.group || p.spec.group;
      if (p.spec.neg === 'paren') best.neg = 'paren';
    }
    // Small values never show a separator, so only treat grouping as off if a big value appeared without one.
    if (best && !best.group && !best.pct) {
      const bigPlain = samples.some(s => Math.abs(s.v) >= 1000 && typeof s.f === 'string' && /\d{4}/.test(s.f));
      best.group = !bigPlain;
    }
    return best;
  }

  function formatNumber(v, spec) {
    if (v == null || !isFinite(v)) return '';
    if (!spec) return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
    let x = Math.abs(v);
    if (spec.pct) x *= 100;
    if (spec.scale) x /= SCALE[spec.scale];
    let body = x.toLocaleString('en-US', { minimumFractionDigits: spec.dec, maximumFractionDigits: spec.dec, useGrouping: spec.group || !!spec.cur });
    body = spec.cur + body + (spec.scale || '') + (spec.pct ? '%' : '');
    const isNeg = v < 0 && Number(x.toFixed(spec.dec)) !== 0;
    if (!isNeg) return body;
    return spec.neg === 'paren' ? '(' + body + ')' : '-' + body;
  }

  function excelNumFmt(spec) {
    if (!spec) return null;
    const dec = spec.dec ? '.' + '0'.repeat(spec.dec) : '';
    let core = (spec.group || spec.cur ? '#,##0' : '0') + dec;
    if (spec.scale) core += ','.repeat({ K: 1, M: 2, B: 3 }[spec.scale]) + '"' + spec.scale + '"';
    if (spec.pct) core += '%';
    const pos = (spec.cur ? '"' + spec.cur + '"' : '') + core;
    const neg = spec.neg === 'paren' ? '(' + pos + ')' : '-' + pos;
    return pos + ';' + neg;
  }

  // ---- dates ----

  // The whole value must be a date (or date and time), so codes like "2024-03-15_Spring" stay text.
  const ISO_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?Z?)?$/;

  function parseIsoDate(s) {
    const m = typeof s === 'string' && s.match(ISO_RE);
    if (!m) return null;
    const t = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    if (isNaN(t)) return null;
    const d = new Date(t);
    if (d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) return null;
    return { t, time: m[4] !== undefined };
  }

  function inferDateFmt(formattedSamples, hasTime) {
    const f = formattedSamples.find(x => typeof x === 'string' && x.trim());
    let d = 'yyyy-mm-dd';
    if (f && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(f)) d = 'm/d/yyyy';
    else if (f && /^\d{1,2}\/\d{1,2}\/\d{2}\b/.test(f)) d = 'm/d/yy';
    if (!hasTime) return d;
    return d + (f && /[AP]M/i.test(f) ? ' h:mm:ss AM/PM' : ' hh:mm:ss');
  }

  return {
    TOTAL_TYPES,
    normName,
    findColumn,
    validateRules,
    resolveTotal,
    describeTotal,
    computeTotal,
    parseFormatted,
    inferSpec,
    formatNumber,
    excelNumFmt,
    parseIsoDate,
    inferDateFmt
  };
});
