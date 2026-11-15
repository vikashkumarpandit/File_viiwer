/* readers.js — PDF, Word (.docx), Excel (.xlsx), CSV, text and image readers.
   Everything except PDF is plain JavaScript (+ JSZip). PDF uses pdf.js from ./vendor/ (added by the build step). */
(function (global) {
'use strict';

/* ------------------------------------------------------------ tiny XML helpers */
function kid(el, n) { if (!el) return null; for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === n) return c; return null; }
function kids(el, n) { const o = []; if (!el) return o; for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!n || c.localName === n) o.push(c); return o; }
function attr(el, n, d) {
  if (!el) return d;
  let v = el.getAttribute(n);
  if (v === null) v = el.getAttribute('w:' + n);
  if (v === null) v = el.getAttribute('r:' + n);
  return v === null ? d : v;
}
function parseXml(t) {
  const d = new DOMParser().parseFromString(t, 'application/xml');
  return d.getElementsByTagName('parsererror').length ? null : d.documentElement;
}
function mk(tag, css, text) { const e = document.createElement(tag); if (css) e.style.cssText = css; if (text != null) e.textContent = text; return e; }
function dirname(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }
function resolve(base, target) {
  if (target.charAt(0) === '/') return target.slice(1);
  const out = [];
  for (const s of (dirname(base) + target).split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); }
  return out.join('/');
}
async function zipXml(zip, p) { const f = zip.file(p); return f ? parseXml(await f.async('string')) : null; }
async function zipRels(zip, part) {
  const p = dirname(part) + '_rels/' + part.slice(part.lastIndexOf('/') + 1) + '.rels';
  const root = await zipXml(zip, p);
  const m = {};
  for (const r of kids(root, 'Relationship')) if (r.getAttribute('TargetMode') !== 'External') m[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: resolve(part, r.getAttribute('Target') || '') };
  return m;
}
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

/* ------------------------------------------------------------ DOCX (reflowing view) */
async function loadDocx(buf) {
  const zip = await global.JSZip.loadAsync(buf);
  const doc = await zipXml(zip, 'word/document.xml');
  if (!doc) throw new Error('This is not a valid .docx file.');
  const rels = await zipRels(zip, 'word/document.xml');
  const styles = await zipXml(zip, 'word/styles.xml');
  const numbering = await zipXml(zip, 'word/numbering.xml');
  const urls = [];

  const styleName = {}, styleNum = {}, styleBase = {};
  for (const s of kids(styles, 'style')) {
    const sid = s.getAttribute('w:styleId') || s.getAttribute('styleId');
    styleName[sid] = (attr(kid(s, 'name'), 'val', '') || '').toLowerCase();
    styleBase[sid] = attr(kid(s, 'basedOn'), 'val', '');
    const sn = kid(kid(s, 'pPr'), 'numPr');
    if (sn) styleNum[sid] = sn;
  }
  const numFromStyle = (sid) => { for (let i = 0; i < 4 && sid; i++) { if (styleNum[sid]) return styleNum[sid]; sid = styleBase[sid]; } return null; };

  // numbering: numId -> {ilvl -> fmt}
  const absFmt = {}, numToAbs = {};
  for (const a of kids(numbering, 'abstractNum')) {
    const id = a.getAttribute('w:abstractNumId') || a.getAttribute('abstractNumId');
    absFmt[id] = {};
    for (const l of kids(a, 'lvl')) absFmt[id][l.getAttribute('w:ilvl') || l.getAttribute('ilvl')] = attr(kid(l, 'numFmt'), 'val', 'bullet');
  }
  for (const n of kids(numbering, 'num')) numToAbs[n.getAttribute('w:numId') || n.getAttribute('numId')] = attr(kid(n, 'abstractNumId'), 'val', '');
  const counters = {};
  const alpha = (n) => { let t = ''; while (n > 0) { n--; t = String.fromCharCode(97 + n % 26) + t; n = Math.floor(n / 26); } return t; };
  const roman = (k) => { let t = ''; for (const [v, c] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (k >= v) { t += c; k -= v; } return t; };
  function marker(numId, ilvl) {
    const fmt = (absFmt[numToAbs[numId]] || {})[ilvl] || 'bullet';
    if (fmt === 'bullet' || fmt === 'none') return fmt === 'none' ? '' : '•';
    const key = numId + ':' + ilvl;
    counters[key] = (counters[key] || 0) + 1;
    for (const k of Object.keys(counters)) { const [nid, lv] = k.split(':'); if (nid === numId && +lv > +ilvl) delete counters[k]; }
    const n = counters[key];
    if (/lowerLetter/.test(fmt)) return alpha(n) + '.';
    if (/upperLetter/.test(fmt)) return alpha(n).toUpperCase() + '.';
    if (/lowerRoman/.test(fmt)) return roman(n) + '.';
    if (/upperRoman/.test(fmt)) return roman(n).toUpperCase() + '.';
    return n + '.';
  }

  async function imageFor(drawing) {
    const bl = drawing.getElementsByTagNameNS('*', 'blip')[0];
    if (!bl) return null;
    const id = bl.getAttribute('r:embed') || bl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
    const rel = rels[id];
    if (!rel) return null;
    const ext = rel.target.split('.').pop().toLowerCase();
    const f = zip.file(rel.target);
    if (!f || !MIME[ext]) return null;
    const url = URL.createObjectURL(new Blob([await f.async('uint8array')], { type: MIME[ext] }));
    urls.push(url);
    const img = document.createElement('img');
    img.src = url;
    const ex = drawing.getElementsByTagNameNS('*', 'extent')[0];
    if (ex) { const w = parseInt(ex.getAttribute('cx') || '0', 10) / 9525; if (w) img.style.width = w + 'px'; }
    img.style.maxWidth = '100%'; img.style.height = 'auto';
    return img;
  }

  async function runs(parent, out) {
    for (const c of kids(parent)) {
      switch (c.localName) {
        case 'r': {
          const rPr = kid(c, 'rPr');
          const css = [];
          if (rPr) {
            const on = (n) => { const e = kid(rPr, n); return e && !/^(0|false|off)$/.test(attr(e, 'val', '1')); };
            if (on('b')) css.push('font-weight:700');
            if (on('i')) css.push('font-style:italic');
            const u = kid(rPr, 'u'); if (u && attr(u, 'val', 'single') !== 'none') css.push('text-decoration:underline');
            if (on('strike')) css.push('text-decoration:line-through');
            const col = attr(kid(rPr, 'color'), 'val', null); if (col && /^[0-9a-f]{6}$/i.test(col)) css.push('color:#' + col);
            const sz = attr(kid(rPr, 'sz'), 'val', null); if (sz) css.push('font-size:' + (parseInt(sz, 10) / 2 * 4 / 3).toFixed(1) + 'px');
            const hl = attr(kid(rPr, 'highlight'), 'val', null); if (hl && hl !== 'none') css.push('background:' + hl);
            const va = attr(kid(rPr, 'vertAlign'), 'val', null); if (va === 'superscript') css.push('vertical-align:super;font-size:.75em'); else if (va === 'subscript') css.push('vertical-align:sub;font-size:.75em');
            if (on('caps')) css.push('text-transform:uppercase');
          }
          for (const k of kids(c)) {
            if (k.localName === 't') { const s = mk('span', css.join(';'), k.textContent); s.style.whiteSpace = 'pre-wrap'; out.appendChild(s); }
            else if (k.localName === 'tab') out.appendChild(mk('span', 'display:inline-block;width:2em'));
            else if (k.localName === 'br' || k.localName === 'cr') { if (attr(k, 'type', '') === 'page') out.appendChild(mk('hr', 'border:0;border-top:1px dashed #bbb;margin:14px 0')); else out.appendChild(document.createElement('br')); }
            else if (k.localName === 'drawing') { const im = await imageFor(k); if (im) out.appendChild(im); }
            else if (k.localName === 'pict') { const im = k.getElementsByTagNameNS('*', 'imagedata')[0]; if (im) { const id = im.getAttribute('r:id'); const rel = rels[id]; const f = rel && zip.file(rel.target); const ext = rel ? rel.target.split('.').pop().toLowerCase() : ''; if (f && MIME[ext]) { const url = URL.createObjectURL(new Blob([await f.async('uint8array')], { type: MIME[ext] })); urls.push(url); const img = mk('img', 'max-width:100%;height:auto'); img.src = url; out.appendChild(img); } } }
          }
          break;
        }
        case 'hyperlink': case 'fldSimple': case 'ins': case 'smartTag': case 'sdt': case 'sdtContent': case 'customXml': {
          await runs(c, out);
          break;
        }
        default: break;
      }
    }
  }

  async function para(p) {
    const pPr = kid(p, 'pPr');
    const sid = attr(kid(pPr, 'pStyle'), 'val', '');
    const sname = styleName[sid] || sid.toLowerCase();
    let tag = 'p';
    const hm = /heading\s*([1-6])/.exec(sname);
    if (hm) tag = 'h' + hm[1]; else if (/^title$/.test(sname)) tag = 'h1'; else if (/^subtitle$/.test(sname)) tag = 'h2';
    const el = document.createElement(tag);
    el.className = 'dx-p';
    const jc = attr(kid(pPr, 'jc'), 'val', '');
    if (jc === 'center') el.style.textAlign = 'center'; else if (jc === 'right' || jc === 'end') el.style.textAlign = 'right'; else if (jc === 'both') el.style.textAlign = 'justify';
    const ind = kid(pPr, 'ind');
    const left = parseInt(attr(ind, 'left', attr(ind, 'start', '0')), 10) || 0;
    if (left) el.style.marginLeft = Math.min(left / 15, 300) + 'px';
    const numPr = kid(pPr, 'numPr') || numFromStyle(sid);
    if (numPr) {
      const numId = attr(kid(numPr, 'numId'), 'val', '0');
      const lv = attr(kid(numPr, 'ilvl'), 'val', '0');
      if (numId !== '0') {
        const m = marker(numId, lv);
        if (m) {
          el.style.display = 'flex'; el.style.gap = '.5em';
          el.style.marginLeft = (24 + (parseInt(lv, 10) || 0) * 24) + 'px';
          const body = mk('span', 'flex:1;min-width:0');
          await runs(p, body);
          el.appendChild(mk('span', 'flex:none;min-width:1.2em', m));
          el.appendChild(body);
          return el;
        }
      }
    }
    await runs(p, el);
    if (!el.firstChild) el.appendChild(document.createElement('br'));
    return el;
  }

  async function table(t) {
    const tbl = mk('table', 'border-collapse:collapse;margin:10px 0;max-width:100%');
    tbl.className = 'dx-t';
    const hasBorders = !!kid(kid(t, 'tblPr'), 'tblBorders');
    for (const tr of kids(t, 'tr')) {
      const row = document.createElement('tr');
      for (const tc of kids(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr');
        if (kid(tcPr, 'vMerge') && attr(kid(tcPr, 'vMerge'), 'val', 'continue') === 'continue') continue;
        const td = document.createElement('td');
        td.style.cssText = 'vertical-align:top;padding:4px 7px;' + (hasBorders ? 'border:1px solid #888;' : 'border:1px solid #ddd;');
        const gs = parseInt(attr(kid(tcPr, 'gridSpan'), 'val', '1'), 10); if (gs > 1) td.colSpan = gs;
        const shd = attr(kid(tcPr, 'shd'), 'fill', null); if (shd && /^[0-9a-f]{6}$/i.test(shd)) td.style.background = '#' + shd;
        await block(tc, td);
        row.appendChild(td);
      }
      tbl.appendChild(row);
    }
    const wrap = mk('div', 'overflow-x:auto;max-width:100%');
    wrap.appendChild(tbl);
    return wrap;
  }

  async function block(parent, out) {
    for (const c of kids(parent)) {
      if (c.localName === 'p') out.appendChild(await para(c));
      else if (c.localName === 'tbl') out.appendChild(await table(c));
      else if (c.localName === 'sdt' || c.localName === 'sdtContent') await block(c, out);
    }
  }

  return {
    mount(container) {
      container.classList.add('rd-paper');
      const page = mk('div', '', null);
      page.className = 'dx';
      container.appendChild(page);
      return block(kid(doc, 'body'), page).then(() => ({ el: page }));
    },
    zoomable: true,
    destroy() { urls.forEach((u) => URL.revokeObjectURL(u)); }
  };
}

/* ------------------------------------------------------------ spreadsheets */
function colIndex(ref) { const m = /^([A-Z]+)/.exec(ref); let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
function colName(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
const BUILTIN_DATE = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
function isDateCode(code) { return /[dmyhs]/i.test(code.replace(/\[[^\]]*\]|"[^"]*"|\\./g, '')) && !/^general$/i.test(code); }
function pad(n) { return n < 10 ? '0' + n : '' + n; }
function fmtSerial(v, code) {
  const d = new Date(Math.round((v - 25569) * 86400000));
  if (isNaN(d)) return String(v);
  const hasTime = /h|s/i.test(code.replace(/\[[^\]]*\]|"[^"]*"/g, '')) && !/d|y/i.test(code.replace(/\[[^\]]*\]|"[^"]*"/g, ''));
  const date = d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  const time = pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes());
  if (hasTime) return time;
  return (v % 1 && /h/i.test(code)) ? date + ' ' + time : date;
}
function fmtNum(v, code) {
  if (!code || /^general$/i.test(code)) return String(Number(v.toPrecision(12)));
  const sec = code.split(';')[0].replace(/"[^"]*"/g, (m) => m.slice(1, -1)).replace(/\[[^\]]*\]/g, '').replace(/_./g, '').replace(/\\/g, '').replace(/\*./g, '');
  const pct = /%/.test(sec);
  const m = /0\.(0+)/.exec(sec) || /#\.(0+)/.exec(sec);
  const dec = m ? m[1].length : 0;
  let n = pct ? v * 100 : v;
  let str = Math.abs(n).toFixed(dec);
  if (/,/.test(sec)) { const [i, f] = str.split('.'); str = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? '.' + f : ''); }
  const pre = (/^[^0#?.,]*/.exec(sec) || [''])[0].replace(/[%]/g, '');
  return (n < 0 ? '-' : '') + pre + str + (pct ? '%' : '');
}

async function loadXlsx(buf) {
  const zip = await global.JSZip.loadAsync(buf);
  const wb = await zipXml(zip, 'xl/workbook.xml');
  if (!wb) throw new Error('This is not a valid .xlsx file.');
  const wrels = await zipRels(zip, 'xl/workbook.xml');
  const sst = await zipXml(zip, 'xl/sharedStrings.xml');
  const strings = kids(sst, 'si').map((si) => {
    let t = '';
    const walk = (e) => { for (const c of kids(e)) { if (c.localName === 't') t += c.textContent; else if (c.localName === 'r') walk(c); } };
    walk(si);
    return t;
  });
  const st = await zipXml(zip, 'xl/styles.xml');
  const codes = {};
  for (const nf of kids(kid(st, 'numFmts'), 'numFmt')) codes[nf.getAttribute('numFmtId')] = nf.getAttribute('formatCode') || '';
  const xfs = kids(kid(st, 'cellXfs'), 'xf').map((x) => parseInt(x.getAttribute('numFmtId') || '0', 10));
  const sheets = [];
  for (const s of kids(kid(wb, 'sheets'), 'sheet')) {
    const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    if (wrels[rid]) sheets.push({ name: s.getAttribute('name'), path: wrels[rid].target, hidden: s.getAttribute('state') === 'hidden' });
  }
  const vis = sheets.filter((x) => !x.hidden);
  const list = vis.length ? vis : sheets;

  async function readSheet(sh) {
    const root = await zipXml(zip, sh.path);
    const grid = [];
    let maxC = 0;
    const MAXR = 5000, MAXC = 120;
    for (const row of kids(kid(root, 'sheetData'), 'row')) {
      const r = parseInt(row.getAttribute('r') || '0', 10) - 1;
      if (r < 0 || r >= MAXR) continue;
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r') || '';
        const ci = colIndex(ref);
        if (ci >= MAXC) continue;
        const t = c.getAttribute('t');
        const v = kid(c, 'v');
        let val = '';
        if (t === 's' && v) val = strings[parseInt(v.textContent, 10)] || '';
        else if (t === 'inlineStr') val = kids(kid(c, 'is'), 't').map((x) => x.textContent).join('') || (kid(c, 'is') ? kid(c, 'is').textContent : '');
        else if (t === 'b') val = v && v.textContent === '1' ? 'TRUE' : 'FALSE';
        else if (v) {
          if (t === 'str' || t === 'e') val = v.textContent;
          else {
            const num = parseFloat(v.textContent);
            const sid = parseInt(c.getAttribute('s') || '0', 10);
            const nid = xfs[sid] || 0;
            const code = codes[nid] !== undefined ? codes[nid] : '';
            if (isNaN(num)) val = v.textContent;
            else if (BUILTIN_DATE.has(nid) || (code && isDateCode(code))) val = fmtSerial(num, code || (nid === 18 || nid === 20 || nid === 21 ? 'h:mm' : 'yyyy-mm-dd'));
            else if (nid === 9 || nid === 10) val = fmtNum(num, nid === 9 ? '0%' : '0.00%');
            else if (nid === 2 || nid === 4) val = fmtNum(num, nid === 2 ? '0.00' : '#,##0.00');
            else if (nid === 3) val = fmtNum(num, '#,##0');
            else val = fmtNum(num, code);
          }
        }
        (grid[r] = grid[r] || [])[ci] = { v: val, num: !t && !!v };
        if (ci + 1 > maxC) maxC = ci + 1;
      }
    }
    const merges = [];
    for (const m of kids(kid(root, 'mergeCells'), 'mergeCell')) {
      const [a, b] = (m.getAttribute('ref') || '').split(':');
      if (a && b) merges.push({ r1: parseInt(/\d+/.exec(a)[0], 10) - 1, c1: colIndex(a), r2: parseInt(/\d+/.exec(b)[0], 10) - 1, c2: colIndex(b) });
    }
    const widths = {};
    for (const c of kids(kid(root, 'cols'), 'col')) {
      const w = parseFloat(c.getAttribute('width') || '0');
      for (let i = parseInt(c.getAttribute('min'), 10) - 1; i < Math.min(parseInt(c.getAttribute('max'), 10), MAXC); i++) widths[i] = Math.round(w * 7 + 5);
    }
    return { grid, maxC, merges, widths, rows: grid.length };
  }
  return { sheetNames: list.map((x) => x.name), read: (i) => readSheet(list[i]) };
}

function gridToTable(data) {
  const { grid, maxC, merges, widths } = data;
  const wrap = mk('div', '');
  wrap.className = 'xl-wrap';
  const tbl = document.createElement('table');
  tbl.className = 'xl';
  const cg = document.createElement('colgroup');
  cg.appendChild(mk('col', 'width:42px'));
  for (let c = 0; c < maxC; c++) cg.appendChild(mk('col', 'width:' + clamp(widths[c] || 80, 40, 420) + 'px'));
  tbl.appendChild(cg);
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  hr.appendChild(mk('th', '', ''));
  for (let c = 0; c < maxC; c++) hr.appendChild(mk('th', '', colName(c)));
  thead.appendChild(hr); tbl.appendChild(thead);
  const skip = new Set();
  const span = {};
  for (const m of merges) { span[m.r1 + ',' + m.c1] = m; for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) skip.add(r + ',' + c); }
  const tb = document.createElement('tbody');
  for (let r = 0; r < grid.length; r++) {
    const tr = document.createElement('tr');
    tr.appendChild(mk('th', '', String(r + 1)));
    for (let c = 0; c < maxC; c++) {
      if (skip.has(r + ',' + c)) continue;
      const cell = grid[r] && grid[r][c];
      const td = mk('td', '', cell ? cell.v : '');
      if (cell && cell.num) td.style.textAlign = 'right';
      const m = span[r + ',' + c];
      if (m) { td.colSpan = m.c2 - m.c1 + 1; td.rowSpan = m.r2 - m.r1 + 1; }
      tr.appendChild(td);
    }
    tb.appendChild(tr);
  }
  tbl.appendChild(tb);
  wrap.appendChild(tbl);
  return wrap;
}

function csvParse(text) {
  const rows = [];
  let row = [], cur = '', q = false;
  const delim = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : (text.split('\n')[0].match(/\t/g) || []).length > 2 ? '\t' : ',';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
    else if (ch !== '\r') cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function sheetView(container, tabs, getData) {
  container.classList.add('rd-sheet');
  const bar = mk('div', '');
  bar.className = 'xl-tabs';
  const body = mk('div', '');
  container.append(bar, body);
  let busy = 0;
  async function show(i) {
    const tok = ++busy;
    [...bar.children].forEach((b, k) => b.classList.toggle('on', k === i));
    body.textContent = 'Loading…';
    const data = await getData(i);
    if (tok !== busy) return;
    body.textContent = '';
    if (!data.grid.length) { body.appendChild(mk('p', 'padding:20px;color:#888', 'This sheet is empty.')); return; }
    body.appendChild(gridToTable(data));
    if (data.rows >= 5000) body.appendChild(mk('p', 'padding:12px;color:#888', 'Showing the first 5000 rows.'));
  }
  tabs.forEach((n, i) => { const b = mk('button', '', n); b.type = 'button'; b.addEventListener('click', () => show(i)); bar.appendChild(b); });
  if (tabs.length < 2) bar.style.display = 'none';
  return show(0);
}

/* ------------------------------------------------------------ PDF */
let pdfLib = null;
async function pdfjs() {
  if (pdfLib) return pdfLib;
  try { pdfLib = await import('./vendor/pdf.min.mjs'); }
  catch (e) { throw new Error('The PDF engine is missing. Deploy the app with the GitHub Actions workflow (build.yml) so that pdf.js gets installed.'); }
  pdfLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', location.href).href;
  return pdfLib;
}
async function loadPdf(buf) {
  const lib = await pdfjs();
  let pdf = null, pw;
  for (let tries = 0; tries < 4 && !pdf; tries++) {
    try {
      pdf = await lib.getDocument({
        data: new Uint8Array(buf.slice(0)), password: pw,
        cMapUrl: './vendor/cmaps/', cMapPacked: true, standardFontDataUrl: './vendor/standard_fonts/', wasmUrl: './vendor/wasm/', iccUrl: './vendor/iccs/'
      }).promise;
    } catch (e) {
      if (e && e.name === 'PasswordException') {
        pw = global.prompt(tries ? 'Wrong password. Try again:' : 'This PDF is password protected. Enter password:');
        if (pw === null) throw new Error('Password needed to open this PDF.');
      } else throw e;
    }
  }
  if (!pdf) throw new Error('Could not open this PDF.');
  const first = await pdf.getPage(1);
  const vp1 = first.getViewport({ scale: 1 });
  const dpr = Math.min(global.devicePixelRatio || 1, 2.5);

  async function drawTo(canvas, page, cssW) {
    const vp = page.getViewport({ scale: 1 });
    let sc = cssW / vp.width * dpr;
    const maxPix = 14e6;
    if (vp.width * vp.height * sc * sc > maxPix) sc = Math.sqrt(maxPix / (vp.width * vp.height));
    const v = page.getViewport({ scale: sc });
    canvas.width = Math.floor(v.width); canvas.height = Math.floor(v.height);
    canvas.style.width = '100%'; canvas.style.height = '100%';
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, canvas, viewport: v }).promise;
  }

  const slides = {
    width: vp1.width, height: vp1.height, count: pdf.numPages,
    async renderSlide(i) {
      const page = await pdf.getPage(i + 1);
      const vp = page.getViewport({ scale: 1 });
      const el = mk('div', `position:relative;width:${vp.width}px;height:${vp.height}px;background:#fff;`);
      el._w = vp.width; el._h = vp.height;
      const cv = document.createElement('canvas');
      el.appendChild(cv);
      await drawTo(cv, page, Math.min(vp.width * Math.max(2, dpr), 2200) / Math.max(dpr, 1));
      return el;
    },
    fit() {}, destroy() {}
  };

  let cleanup = null;
  const api = {
    pages: pdf.numPages,
    scroller: true,
    slides,
    zoomable: true,
    mount(container, onPage) {
      container.classList.add('rd-pdf');
      let z = 1;
      let io = null;
      const build = () => {
        if (io) io.disconnect();
        container.textContent = '';
        const cw = Math.min(container.clientWidth - 12, 1100) * z;
        const ratio = vp1.height / vp1.width;
        const boxes = [];
        for (let i = 1; i <= pdf.numPages; i++) {
          const b = mk('div', `width:${cw}px;height:${cw * ratio}px;`);
          b.className = 'pdf-pg'; b.dataset.i = i;
          container.appendChild(b); boxes.push(b);
        }
        io = new IntersectionObserver((ents) => {
          for (const en of ents) {
            if (!en.isIntersecting) continue;
            const b = en.target;
            if (b.dataset.done) continue;
            b.dataset.done = '1';
            (async () => {
              try {
                const page = await pdf.getPage(+b.dataset.i);
                const vp = page.getViewport({ scale: 1 });
                b.style.height = (cw * vp.height / vp.width) + 'px';
                const cv = document.createElement('canvas');
                b.appendChild(cv);
                await drawTo(cv, page, cw);
              } catch (e) { b.textContent = 'Could not draw this page.'; }
            })();
          }
        }, { root: container, rootMargin: '700px 0px' });
        boxes.forEach((b) => io.observe(b));
      };
      build();
      let raf = 0;
      const onScroll = () => {
        cancelAnimationFrame(raf);
        raf = requestAnimationFrame(() => {
          const mid = container.scrollTop + container.clientHeight / 2;
          for (const b of container.children) { if (b.offsetTop + b.offsetHeight >= mid) { if (onPage) onPage(+b.dataset.i, pdf.numPages); break; } }
        });
      };
      container.addEventListener('scroll', onScroll, { passive: true });
      onScroll();
      api.zoom = (f) => {
        const frac = container.scrollTop / Math.max(1, container.scrollHeight);
        z = clamp(z * f, 0.6, 4);
        build();
        container.scrollTop = frac * container.scrollHeight;
      };
      cleanup = () => { if (io) io.disconnect(); container.removeEventListener('scroll', onScroll); };
      return Promise.resolve({});
    },
    destroy() { if (cleanup) cleanup(); try { pdf.destroy(); } catch (e) { /* ignore */ } }
  };
  return api;
}

/* ------------------------------------------------------------ public */
const TYPES = {
  pptx: 'pptx', ppsx: 'pptx', pptm: 'pptx', potx: 'pptx', ppt: 'ppt', pps: 'ppt',
  pdf: 'pdf', docx: 'docx', docm: 'docx', xlsx: 'xlsx', xlsm: 'xlsx', csv: 'csv', tsv: 'csv',
  txt: 'text', md: 'text', log: 'text', json: 'text', xml: 'text', ini: 'text', yml: 'text', yaml: 'text',
  png: 'image', jpg: 'image', jpeg: 'image', gif: 'image', webp: 'image', bmp: 'image', svg: 'image'
};
const UNSUPPORTED = { doc: 'Old Word (.doc)', xls: 'Old Excel (.xls)', odt: 'OpenDocument (.odt)', ods: 'OpenDocument (.ods)', odp: 'OpenDocument (.odp)', rtf: 'RTF' };
function extOf(name) { const m = /\.([A-Za-z0-9]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
function typeOf(name) { return TYPES[extOf(name)] || null; }

async function open(type, buf, name) {
  if (type === 'pdf') return loadPdf(buf);
  if (type === 'docx') return loadDocx(buf);
  if (type === 'xlsx') {
    const x = await loadXlsx(buf);
    return { zoomable: true, destroy() {}, mount(c) { return sheetView(c, x.sheetNames, (i) => x.read(i)).then(() => ({})); } };
  }
  if (type === 'csv') {
    const text = new TextDecoder('utf-8').decode(new Uint8Array(buf)).replace(/^﻿/, '');
    const rows = csvParse(text).slice(0, 5000);
    const maxC = Math.min(120, rows.reduce((a, r) => Math.max(a, r.length), 0));
    const grid = rows.map((r) => r.slice(0, maxC).map((v) => ({ v, num: /^-?\d+(\.\d+)?$/.test(v) })));
    return { zoomable: true, destroy() {}, mount(c) { return sheetView(c, ['Sheet'], async () => ({ grid, maxC, merges: [], widths: {}, rows: grid.length })).then(() => ({})); } };
  }
  if (type === 'text') {
    let text = new TextDecoder('utf-8').decode(new Uint8Array(buf));
    if (extOf(name) === 'json') { try { text = JSON.stringify(JSON.parse(text), null, 2); } catch (e) { /* keep raw */ } }
    return { zoomable: true, destroy() {}, mount(c) { c.classList.add('rd-text'); c.appendChild(mk('pre', '', text)); return Promise.resolve({}); } };
  }
  if (type === 'image') {
    const ext = extOf(name);
    const url = URL.createObjectURL(new Blob([buf], { type: MIME[ext] || 'image/*' }));
    return { zoomable: true, destroy() { URL.revokeObjectURL(url); }, mount(c) { c.classList.add('rd-image'); const im = document.createElement('img'); im.src = url; c.appendChild(im); return Promise.resolve({}); } };
  }
  throw new Error('This file type is not supported.');
}

global.Readers = { open, typeOf, extOf, TYPES, UNSUPPORTED };
})(typeof window !== 'undefined' ? window : globalThis);
