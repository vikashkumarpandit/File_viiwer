/* pptx.js — PowerPoint (.pptx) ko browser me hi draw karne wala chhota renderer.
   Internet ki zarurat nahi. JSZip chahiye (jszip.min.js). */
(function (global) {
'use strict';

const SVGNS = 'http://www.w3.org/2000/svg';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU_PX = 9525;
const PT = 96 / 72;
const px = (e) => e / EMU_PX;
let UID = 0;

/* ------------------------------------------------------------------ XML helpers */
function kid(el, name) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
  return null;
}
function kids(el, name) {
  const o = [];
  if (!el) return o;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) o.push(c);
  return o;
}
function at(el, ...names) {
  let c = el;
  for (const n of names) { c = kid(c, n); if (!c) return null; }
  return c;
}
function attr(el, n, d) {
  if (!el) return d;
  const v = el.getAttribute(n);
  return v === null ? d : v;
}
function numAttr(el, n, d) {
  const v = attr(el, n, null);
  if (v === null || v === '') return d;
  const f = parseFloat(v);
  return isNaN(f) ? d : f;
}
function findDeep(el, name) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === name) return c;
    const r = findDeep(c, name);
    if (r) return r;
  }
  return null;
}
function rAttr(el, n) {
  if (!el) return null;
  return el.getAttributeNS(NS_R, n) || el.getAttribute('r:' + n) || null;
}
function h(tag, css, cls) {
  const e = document.createElement(tag);
  if (css) e.style.cssText = css;
  if (cls) e.className = cls;
  return e;
}
function s(tag, attrs) {
  const e = document.createElementNS(SVGNS, tag);
  if (attrs) for (const k in attrs) e.setAttribute(k, attrs[k]);
  return e;
}

/* ------------------------------------------------------------------ colours */
const PRST = {
  black: [0, 0, 0], white: [255, 255, 255], red: [255, 0, 0], green: [0, 128, 0], blue: [0, 0, 255],
  yellow: [255, 255, 0], gray: [128, 128, 128], grey: [128, 128, 128], orange: [255, 165, 0],
  ltGray: [211, 211, 211], dkGray: [169, 169, 169], cyan: [0, 255, 255], magenta: [255, 0, 255],
  purple: [128, 0, 128], brown: [165, 42, 42], pink: [255, 192, 203], navy: [0, 0, 128],
  dkBlue: [0, 0, 139], dkRed: [139, 0, 0], dkGreen: [0, 100, 0], ltBlue: [173, 216, 230],
  ltGreen: [144, 238, 144], ltYellow: [255, 255, 224], gold: [255, 215, 0], silver: [192, 192, 192],
  dkGrey: [169, 169, 169], ltGrey: [211, 211, 211]
};
const COLOR_TAGS = new Set(['srgbClr', 'sysClr', 'schemeClr', 'prstClr', 'scrgbClr', 'hslClr']);

function hex2rgb(hx) {
  hx = (hx || '000000').replace('#', '');
  if (hx.length === 3) hx = hx.split('').map((c) => c + c).join('');
  const n = parseInt(hx, 16);
  if (isNaN(n)) return [0, 0, 0];
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function rgb2hsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let hh = 0, ss = 0;
  const l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    ss = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    if (mx === r) hh = (g - b) / d + (g < b ? 6 : 0);
    else if (mx === g) hh = (b - r) / d + 2;
    else hh = (r - g) / d + 4;
    hh /= 6;
  }
  return [hh, ss, l];
}
function hsl2rgb(hh, ss, l) {
  if (ss === 0) { const v = Math.round(l * 255); return [v, v, v]; }
  const q = l < 0.5 ? l * (1 + ss) : l + ss - l * ss;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [Math.round(f(hh + 1 / 3) * 255), Math.round(f(hh) * 255), Math.round(f(hh - 1 / 3) * 255)];
}
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

function colorChild(el) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (COLOR_TAGS.has(c.localName)) return c;
  return null;
}

/* returns {r,g,b,a} or null.  env = {theme, clrMap} */
function resolveColor(el, env, phClr) {
  if (!el) return null;
  const c = COLOR_TAGS.has(el.localName) ? el : colorChild(el);
  if (!c) return null;
  let rgb;
  let alpha = 1;
  switch (c.localName) {
    case 'srgbClr': rgb = hex2rgb(attr(c, 'val', '000000')); break;
    case 'sysClr': rgb = hex2rgb(attr(c, 'lastClr', attr(c, 'val', '000000') === 'window' ? 'FFFFFF' : '000000')); break;
    case 'prstClr': rgb = (PRST[attr(c, 'val', 'black')] || [0, 0, 0]).slice(); break;
    case 'scrgbClr': rgb = [numAttr(c, 'r', 0), numAttr(c, 'g', 0), numAttr(c, 'b', 0)].map((v) => clamp(Math.round(v / 100000 * 255), 0, 255)); break;
    case 'schemeClr': {
      const v = attr(c, 'val', 'tx1');
      if (v === 'phClr') rgb = phClr ? phClr.slice() : [0, 0, 0];
      else {
        const m = (env.clrMap && env.clrMap[v]) || v;
        const hx = env.theme.colors[m] || env.theme.colors[v];
        rgb = hex2rgb(hx || (m === 'lt1' || m === 'bg1' ? 'FFFFFF' : '000000'));
      }
      break;
    }
    default: rgb = [0, 0, 0];
  }
  for (const m of kids(c)) {
    const val = numAttr(m, 'val', 0);
    switch (m.localName) {
      case 'alpha': alpha = val / 100000; break;
      case 'lumMod': { const [a, b, l] = rgb2hsl(...rgb); rgb = hsl2rgb(a, b, clamp(l * val / 100000, 0, 1)); break; }
      case 'lumOff': { const [a, b, l] = rgb2hsl(...rgb); rgb = hsl2rgb(a, b, clamp(l + val / 100000, 0, 1)); break; }
      case 'satMod': { const [a, b, l] = rgb2hsl(...rgb); rgb = hsl2rgb(a, clamp(b * val / 100000, 0, 1), l); break; }
      case 'satOff': { const [a, b, l] = rgb2hsl(...rgb); rgb = hsl2rgb(a, clamp(b + val / 100000, 0, 1), l); break; }
      case 'hueOff': { const [a, b, l] = rgb2hsl(...rgb); rgb = hsl2rgb(((a + val / 21600000) % 1 + 1) % 1, b, l); break; }
      case 'tint': { const t = val / 100000; rgb = rgb.map((x) => Math.round(x * t + 255 * (1 - t))); break; }
      case 'shade': { const t = val / 100000; rgb = rgb.map((x) => Math.round(x * t)); break; }
      default: break;
    }
  }
  return { r: rgb[0], g: rgb[1], b: rgb[2], a: alpha };
}
function cssColor(c) {
  if (!c) return null;
  return c.a >= 1 ? `rgb(${c.r},${c.g},${c.b})` : `rgba(${c.r},${c.g},${c.b},${c.a.toFixed(3)})`;
}
function mixWhite(c, keep) {
  return { r: Math.round(c.r * keep + 255 * (1 - keep)), g: Math.round(c.g * keep + 255 * (1 - keep)), b: Math.round(c.b * keep + 255 * (1 - keep)), a: 1 };
}

/* ------------------------------------------------------------------ fonts */
function fontStack(face, theme) {
  if (!face || face.charAt(0) === '+') {
    if (face === '+mj-lt' || face === '+mj-ea' || face === '+mj-cs') face = theme.major;
    else face = theme.minor;
  }
  const f = String(face || 'Calibri').replace(/["\\]/g, '');
  const l = f.toLowerCase();
  let generic = 'sans-serif';
  if (/times|georgia|cambria|garamond|palatino|book antiqua|minion|baskerville|century|constantia|serif/.test(l) && !/sans/.test(l)) generic = 'serif';
  if (/courier|consolas|mono|lucida console/.test(l)) generic = 'monospace';
  let extra = '';
  if (/^calibri/.test(l)) extra = '"Carlito",';
  else if (/^cambria/.test(l)) extra = '"Caladea",';
  else if (/^arial/.test(l)) extra = '"Liberation Sans",Helvetica,';
  else if (/^times/.test(l)) extra = '"Liberation Serif",';
  else if (/^courier/.test(l)) extra = '"Liberation Mono",';
  return `"${f}",${extra}"Noto Sans","Noto Sans Devanagari","Mangal",Roboto,${generic}`;
}

/* ------------------------------------------------------------------ package */
function dirname(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }
function basename(p) { const i = p.lastIndexOf('/'); return i < 0 ? p : p.slice(i + 1); }
function resolvePath(base, target) {
  if (target.charAt(0) === '/') return target.slice(1);
  const parts = (dirname(base) + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

class Pkg {
  constructor(zip) { this.zip = zip; this._xml = new Map(); this._rels = new Map(); this._img = new Map(); this.urls = []; }
  async xml(p) {
    if (this._xml.has(p)) return this._xml.get(p);
    const f = this.zip.file(p);
    let root = null;
    if (f) {
      const t = await f.async('string');
      const d = new DOMParser().parseFromString(t, 'application/xml');
      if (!d.getElementsByTagName('parsererror').length) root = d.documentElement;
    }
    this._xml.set(p, root);
    return root;
  }
  async rels(p) {
    if (this._rels.has(p)) return this._rels.get(p);
    const rp = dirname(p) + '_rels/' + basename(p) + '.rels';
    const out = {};
    const root = await this.xml(rp);
    for (const r of kids(root, 'Relationship')) {
      if (attr(r, 'TargetMode') === 'External') continue;
      out[attr(r, 'Id')] = { type: attr(r, 'Type', ''), target: resolvePath(p, attr(r, 'Target', '')) };
    }
    this._rels.set(p, out);
    return out;
  }
  async imgUrl(p) {
    if (this._img.has(p)) return this._img.get(p);
    let url = null;
    const ext = (p.split('.').pop() || '').toLowerCase();
    const f = this.zip.file(p);
    if (f && MIME[ext]) {
      const u8 = await f.async('uint8array');
      url = URL.createObjectURL(new Blob([u8], { type: MIME[ext] }));
      this.urls.push(url);
    }
    this._img.set(p, url);
    return url;
  }
}

/* ------------------------------------------------------------------ theme / tables styles */
function parseTheme(root) {
  const colors = {};
  const cs = at(root, 'themeElements', 'clrScheme');
  for (const c of kids(cs)) {
    const v = c.firstElementChild;
    if (!v) continue;
    colors[c.localName] = v.localName === 'sysClr' ? attr(v, 'lastClr', 'FFFFFF') : attr(v, 'val', '000000');
  }
  const fs = at(root, 'themeElements', 'fontScheme');
  const major = attr(at(fs, 'majorFont', 'latin'), 'typeface', 'Calibri');
  const minor = attr(at(fs, 'minorFont', 'latin'), 'typeface', 'Calibri');
  const lnW = kids(at(root, 'themeElements', 'fmtScheme', 'lnStyleLst'), 'ln').map((l) => numAttr(l, 'w', 9525));
  return { colors, major, minor, lnW };
}

function lnFromEl(ln, env) {
  if (!ln) return null;
  if (kid(ln, 'noFill')) return null;
  const col = resolveColor(ln, env);
  if (!col) return null;
  return { w: px(numAttr(ln, 'w', 12700)), color: col };
}
function parseTableStyles(root, env) {
  const out = {};
  for (const st of kids(root, 'tblStyle')) {
    const parts = {};
    for (const part of kids(st)) {
      const o = { fill: null, bold: false, color: null, bdr: {} };
      const tx = kid(part, 'tcTxStyle');
      if (tx) { o.bold = attr(tx, 'b') === 'on'; o.color = resolveColor(tx, env); }
      const tc = kid(part, 'tcStyle');
      if (tc) {
        const f = kid(tc, 'fill');
        if (f && kid(f, 'solidFill')) o.fill = resolveColor(kid(f, 'solidFill'), env);
        else if (f && kid(f, 'noFill')) o.fill = 'none';
        const fr = kid(tc, 'fillRef');
        if (!o.fill && fr) o.fill = resolveColor(fr, env);
        const bd = kid(tc, 'tcBdr');
        for (const side of kids(bd)) {
          const ln = kid(side, 'ln');
          if (ln) o.bdr[side.localName] = lnFromEl(ln, env) || 'none';
          else {
            const lr = kid(side, 'lnRef');
            if (lr) { const c = resolveColor(lr, env); o.bdr[side.localName] = c ? { w: 1.33, color: c } : 'none'; }
          }
        }
      }
      parts[part.localName] = o;
    }
    out[attr(st, 'styleId')] = parts;
  }
  return out;
}
function builtinTableStyle(id, env) {
  const white = { r: 255, g: 255, b: 255, a: 1 };
  const black = { r: 0, g: 0, b: 0, a: 1 };
  const acc = resolveColor({ localName: 'schemeClr', getAttribute: (n) => (n === 'val' ? 'accent1' : null), children: [], firstElementChild: null }, env) || { r: 68, g: 114, b: 196, a: 1 };
  if (id === '{2D5ABB26-0587-4C30-8999-92F81FD0307C}') return {};
  if (id === '{5940675A-B579-460E-94D1-54222C63F5DA}') {
    const b = { w: 1.33, color: black };
    return { wholeTbl: { fill: null, bold: false, color: null, bdr: { left: b, right: b, top: b, bottom: b, insideH: b, insideV: b } } };
  }
  const wb = { w: 1.33, color: white };
  return {
    wholeTbl: { fill: mixWhite(acc, 0.2), bold: false, color: null, bdr: { left: wb, right: wb, top: wb, bottom: wb, insideH: wb, insideV: wb } },
    band1H: { fill: mixWhite(acc, 0.4), bold: false, color: null, bdr: {} },
    firstRow: { fill: acc, bold: true, color: white, bdr: { bottom: { w: 3, color: white } } },
    firstCol: { fill: acc, bold: true, color: white, bdr: {} },
    lastRow: { fill: acc, bold: true, color: white, bdr: { top: { w: 3, color: white } } },
    lastCol: { fill: acc, bold: true, color: white, bdr: {} }
  };
}

/* ------------------------------------------------------------------ geometry */
function rrect(x, y, w, hh, r) {
  r = Math.max(0, Math.min(r, w / 2, hh / 2));
  return `M${x + r},${y} H${x + w - r} A${r},${r} 0 0 1 ${x + w},${y + r} V${y + hh - r} A${r},${r} 0 0 1 ${x + w - r},${y + hh} H${x + r} A${r},${r} 0 0 1 ${x},${y + hh - r} V${y + r} A${r},${r} 0 0 1 ${x + r},${y} Z`;
}
function poly(pts) { return 'M' + pts.map((p) => p[0] + ',' + p[1]).join(' L') + ' Z'; }
function starPts(n, w, hh, ratio) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const ang = -Math.PI / 2 + i * Math.PI / n;
    const rr = i % 2 === 0 ? 1 : ratio;
    pts.push([w / 2 + Math.cos(ang) * rr * w / 2, hh / 2 + Math.sin(ang) * rr * hh / 2]);
  }
  return pts;
}
function arrowRightPts(L, T, a1, a2, ss) {
  const hd = clamp(ss * a2 / 100000, 0, L);
  const half = T * a1 / 200000;
  const y1 = T / 2 - half, y2 = T / 2 + half, xh = L - hd;
  return [[0, y1], [xh, y1], [xh, 0], [L, T / 2], [xh, T], [xh, y2], [0, y2]];
}

function presetPaths(prst, w, hh, av) {
  const adj = (n, d) => (av[n] !== undefined ? av[n] : d);
  const ss = Math.min(w, hh);
  switch (prst) {
    case 'roundRect': return [{ d: rrect(0, 0, w, hh, ss * adj('adj', 16667) / 100000) }];
    case 'flowChartAlternateProcess': return [{ d: rrect(0, 0, w, hh, ss * 0.1667) }];
    case 'flowChartTerminator': return [{ d: rrect(0, 0, w, hh, hh / 2) }];
    case 'round2SameRect': {
      const r1 = ss * adj('adj1', 16667) / 100000, r2 = ss * adj('adj2', 0) / 100000;
      return [{ d: `M${r1},0 H${w - r1} A${r1},${r1} 0 0 1 ${w},${r1} V${hh - r2} A${r2},${r2} 0 0 1 ${w - r2},${hh} H${r2} A${r2},${r2} 0 0 1 0,${hh - r2} V${r1} A${r1},${r1} 0 0 1 ${r1},0 Z` }];
    }
    case 'round1Rect': {
      const r = ss * adj('adj', 16667) / 100000;
      return [{ d: `M0,0 H${w - r} A${r},${r} 0 0 1 ${w},${r} V${hh} H0 Z` }];
    }
    case 'snip1Rect': {
      const r = ss * adj('adj', 16667) / 100000;
      return [{ d: poly([[0, 0], [w - r, 0], [w, r], [w, hh], [0, hh]]) }];
    }
    case 'snip2SameRect': {
      const r = ss * adj('adj1', 16667) / 100000;
      return [{ d: poly([[r, 0], [w - r, 0], [w, r], [w, hh], [0, hh], [0, r]]) }];
    }
    case 'ellipse': case 'flowChartConnector': case 'wedgeEllipseCallout': case 'cloud': case 'cloudCallout':
      return [{ d: `M0,${hh / 2} A${w / 2},${hh / 2} 0 1 0 ${w},${hh / 2} A${w / 2},${hh / 2} 0 1 0 0,${hh / 2} Z` }];
    case 'donut': {
      const t = ss * adj('adj', 25000) / 100000;
      const o = `M0,${hh / 2} A${w / 2},${hh / 2} 0 1 0 ${w},${hh / 2} A${w / 2},${hh / 2} 0 1 0 0,${hh / 2} Z`;
      const i = `M${t},${hh / 2} A${w / 2 - t},${hh / 2 - t} 0 1 0 ${w - t},${hh / 2} A${w / 2 - t},${hh / 2 - t} 0 1 0 ${t},${hh / 2} Z`;
      return [{ d: o + ' ' + i, evenodd: true }];
    }
    case 'triangle': { const x = w * adj('adj', 50000) / 100000; return [{ d: poly([[0, hh], [x, 0], [w, hh]]) }]; }
    case 'rtTriangle': return [{ d: poly([[0, 0], [0, hh], [w, hh]]) }];
    case 'diamond': case 'flowChartDecision': return [{ d: poly([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]) }];
    case 'parallelogram': { const x = ss * adj('adj', 25000) / 100000; return [{ d: poly([[x, 0], [w, 0], [w - x, hh], [0, hh]]) }]; }
    case 'flowChartData': case 'flowChartInputOutput': return [{ d: poly([[w * 0.2, 0], [w, 0], [w * 0.8, hh], [0, hh]]) }];
    case 'trapezoid': { const x = ss * adj('adj', 25000) / 100000; return [{ d: poly([[0, hh], [x, 0], [w - x, 0], [w, hh]]) }]; }
    case 'pentagon': return [{ d: poly(starPts(5, w, hh, 1).filter((_, i) => i % 2 === 0)) }];
    case 'hexagon': { const x = ss * adj('adj', 25000) / 100000; return [{ d: poly([[0, hh / 2], [x, 0], [w - x, 0], [w, hh / 2], [w - x, hh], [x, hh]]) }]; }
    case 'flowChartPreparation': return [{ d: poly([[0, hh / 2], [w * 0.2, 0], [w * 0.8, 0], [w, hh / 2], [w * 0.8, hh], [w * 0.2, hh]]) }];
    case 'octagon': { const x = ss * adj('adj', 29289) / 100000; return [{ d: poly([[x, 0], [w - x, 0], [w, x], [w, hh - x], [w - x, hh], [x, hh], [0, hh - x], [0, x]]) }]; }
    case 'plus': case 'mathPlus': { const x = ss * adj('adj', 25000) / 100000; return [{ d: poly([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, hh - x], [w - x, hh - x], [w - x, hh], [x, hh], [x, hh - x], [0, hh - x], [0, x], [x, x]]) }]; }
    case 'star4': return [{ d: poly(starPts(4, w, hh, adj('adj', 12500) / 50000)) }];
    case 'star5': return [{ d: poly(starPts(5, w, hh, adj('adj', 19098) / 50000)) }];
    case 'star6': return [{ d: poly(starPts(6, w, hh, adj('adj', 28868) / 50000)) }];
    case 'star7': return [{ d: poly(starPts(7, w, hh, adj('adj', 34601) / 50000)) }];
    case 'star8': return [{ d: poly(starPts(8, w, hh, adj('adj', 37500) / 50000)) }];
    case 'star10': return [{ d: poly(starPts(10, w, hh, adj('adj', 42533) / 50000)) }];
    case 'star12': return [{ d: poly(starPts(12, w, hh, adj('adj', 37500) / 50000)) }];
    case 'star16': return [{ d: poly(starPts(16, w, hh, adj('adj', 37500) / 50000)) }];
    case 'star24': return [{ d: poly(starPts(24, w, hh, adj('adj', 37500) / 50000)) }];
    case 'rightArrow': return [{ d: poly(arrowRightPts(w, hh, adj('adj1', 50000), adj('adj2', 50000), ss)) }];
    case 'leftArrow': return [{ d: poly(arrowRightPts(w, hh, adj('adj1', 50000), adj('adj2', 50000), ss).map((p) => [w - p[0], p[1]])) }];
    case 'upArrow': return [{ d: poly(arrowRightPts(hh, w, adj('adj1', 50000), adj('adj2', 50000), ss).map((p) => [p[1], hh - p[0]])) }];
    case 'downArrow': return [{ d: poly(arrowRightPts(hh, w, adj('adj1', 50000), adj('adj2', 50000), ss).map((p) => [p[1], p[0]])) }];
    case 'leftRightArrow': {
      const hd = clamp(ss * adj('adj2', 50000) / 100000, 0, w / 2), half = hh * adj('adj1', 50000) / 200000;
      const y1 = hh / 2 - half, y2 = hh / 2 + half;
      return [{ d: poly([[0, hh / 2], [hd, 0], [hd, y1], [w - hd, y1], [w - hd, 0], [w, hh / 2], [w - hd, hh], [w - hd, y2], [hd, y2], [hd, hh]]) }];
    }
    case 'chevron': { const x = ss * adj('adj', 50000) / 100000; return [{ d: poly([[0, 0], [w - x, 0], [w, hh / 2], [w - x, hh], [0, hh], [x, hh / 2]]) }]; }
    case 'homePlate': { const x = ss * adj('adj', 50000) / 100000; return [{ d: poly([[0, 0], [w - x, 0], [w, hh / 2], [w - x, hh], [0, hh]]) }]; }
    case 'heart': return [{ d: `M${w / 2},${hh} C${-w * 0.1},${hh * 0.55} ${w * 0.05},${-hh * 0.05} ${w / 2},${hh * 0.28} C${w * 0.95},${-hh * 0.05} ${w * 1.1},${hh * 0.55} ${w / 2},${hh} Z` }];
    case 'line': case 'straightConnector1': case 'lineInv': return [{ d: `M0,0 L${w},${hh}`, open: true }];
    case 'bentConnector2': return [{ d: `M0,0 L${w},0 L${w},${hh}`, open: true }];
    case 'bentConnector3': case 'bentConnector4': case 'bentConnector5': { const x = w * adj('adj1', 50000) / 100000; return [{ d: `M0,0 L${x},0 L${x},${hh} L${w},${hh}`, open: true }]; }
    case 'curvedConnector2': return [{ d: `M0,0 C${w / 2},0 ${w},${hh / 2} ${w},${hh}`, open: true }];
    case 'curvedConnector3': case 'curvedConnector4': case 'curvedConnector5': return [{ d: `M0,0 C${w / 2},0 ${w / 2},${hh} ${w},${hh}`, open: true }];
    case 'leftBracket': return [{ d: `M${w},0 L0,0 L0,${hh} L${w},${hh}`, open: true }];
    case 'rightBracket': return [{ d: `M0,0 L${w},0 L${w},${hh} L0,${hh}`, open: true }];
    case 'rightBrace': case 'leftBrace': {
      const d = `M0,0 C${w / 2},0 ${w / 2},${hh * 0.15} ${w / 2},${hh * 0.25} L${w / 2},${hh * 0.4} C${w / 2},${hh * 0.46} ${w * 0.7},${hh * 0.5} ${w},${hh * 0.5} C${w * 0.7},${hh * 0.5} ${w / 2},${hh * 0.54} ${w / 2},${hh * 0.6} L${w / 2},${hh * 0.75} C${w / 2},${hh * 0.85} ${w / 2},${hh} 0,${hh}`;
      if (prst === 'rightBrace') return [{ d, open: true }];
      return [{ d, open: true, mirror: true }];
    }
    default: return [{ d: `M0,0 H${w} V${hh} H0 Z` }];
  }
}

function custGeomPaths(cg, w, hh) {
  const out = [];
  for (const p of kids(at(cg, 'pathLst'), 'path')) {
    const pw = numAttr(p, 'w', 0) || w, ph = numAttr(p, 'h', 0) || hh;
    const sx = w / pw, sy = hh / ph;
    const fillMode = attr(p, 'fill', 'norm');
    const stroke = attr(p, 'stroke', '1') !== '0';
    let d = '';
    let cur = { x: 0, y: 0 };
    for (const cmd of kids(p)) {
      const pts = kids(cmd, 'pt').map((pt) => ({ x: numAttr(pt, 'x', 0) * sx, y: numAttr(pt, 'y', 0) * sy }));
      switch (cmd.localName) {
        case 'moveTo': if (pts[0]) { d += `M${pts[0].x},${pts[0].y} `; cur = pts[0]; } break;
        case 'lnTo': if (pts[0]) { d += `L${pts[0].x},${pts[0].y} `; cur = pts[0]; } break;
        case 'cubicBezTo': if (pts.length >= 3) { d += `C${pts[0].x},${pts[0].y} ${pts[1].x},${pts[1].y} ${pts[2].x},${pts[2].y} `; cur = pts[2]; } break;
        case 'quadBezTo': if (pts.length >= 2) { d += `Q${pts[0].x},${pts[0].y} ${pts[1].x},${pts[1].y} `; cur = pts[1]; } break;
        case 'arcTo': {
          const wR = numAttr(cmd, 'wR', 0) * sx, hR = numAttr(cmd, 'hR', 0) * sy;
          const st = numAttr(cmd, 'stAng', 0) / 60000 * Math.PI / 180;
          const sw = numAttr(cmd, 'swAng', 0) / 60000 * Math.PI / 180;
          if (!wR || !hR) break;
          const cx = cur.x - wR * Math.cos(st), cy = cur.y - hR * Math.sin(st);
          const steps = Math.abs(sw) >= Math.PI * 2 - 0.001 ? 2 : 1;
          for (let i = 1; i <= steps; i++) {
            const a = st + sw * i / steps;
            const ex = cx + wR * Math.cos(a), ey = cy + hR * Math.sin(a);
            d += `A${wR},${hR} 0 ${Math.abs(sw / steps) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${ex},${ey} `;
            cur = { x: ex, y: ey };
          }
          break;
        }
        case 'close': d += 'Z '; break;
        default: break;
      }
    }
    if (d) out.push({ d, noFill: fillMode === 'none', noStroke: !stroke });
  }
  return out;
}

/* ------------------------------------------------------------------ text */
const ALGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify', thaiDist: 'justify' };
const WING = { '§': '▪', 'q': '❑', 'Ø': '➢', 'ü': '✓', 'v': '❖', 'n': '■', 'l': '●', 'o': '○', 'þ': '☑', '¨': '◻', 'w': '◆', 'à': '➔', 'è': '➔' };

function mergePPr(acc, p) {
  if (!p) return;
  for (const n of ['algn', 'marL', 'indent', 'rtl']) { const v = p.getAttribute(n); if (v !== null) acc[n] = v; }
  for (let c = p.firstElementChild; c; c = c.nextElementSibling) {
    switch (c.localName) {
      case 'lnSpc': acc.lnSpc = c; break;
      case 'spcBef': acc.spcBef = c; break;
      case 'spcAft': acc.spcAft = c; break;
      case 'buNone': acc.bu = { t: 'none' }; break;
      case 'buChar': acc.bu = { t: 'char', ch: c.getAttribute('char') || '•' }; break;
      case 'buBlip': acc.bu = { t: 'char', ch: '•' }; break;
      case 'buAutoNum': acc.bu = { t: 'auto', type: c.getAttribute('type') || 'arabicPeriod', start: parseInt(c.getAttribute('startAt') || '1', 10) }; break;
      case 'buFont': acc.buFont = c.getAttribute('typeface'); break;
      case 'buFontTx': acc.buFont = null; break;
      case 'buClr': acc.buClr = c; break;
      case 'buClrTx': acc.buClr = null; break;
      case 'buSzPct': acc.buSz = parseFloat(c.getAttribute('val')) / 100000; break;
      case 'buSzTx': acc.buSz = null; break;
      case 'defRPr': acc.rprs.push(c); break;
      default: break;
    }
  }
}
function resolvePPr(chain, lvl, own) {
  const acc = { rprs: [] };
  for (const src of chain) { if (src) mergePPr(acc, kid(src, 'lvl' + (lvl + 1) + 'pPr')); }
  mergePPr(acc, own);
  return acc;
}
function rv(rprs, name) {
  for (let i = rprs.length - 1; i >= 0; i--) {
    const e = rprs[i];
    const v = e ? e.getAttribute(name) : null;
    if (v !== null && v !== undefined) return v;
  }
  return null;
}
function rchild(rprs, name) {
  for (let i = rprs.length - 1; i >= 0; i--) { const c = kid(rprs[i], name); if (c) return c; }
  return null;
}
function rfill(rprs) {
  for (let i = rprs.length - 1; i >= 0; i--) {
    const e = rprs[i];
    if (!e) continue;
    for (const c of kids(e)) if (c.localName === 'solidFill' || c.localName === 'noFill' || c.localName === 'gradFill') return c;
  }
  return null;
}
function autoNum(type, n) {
  const alpha = (k) => { let t = ''; while (k > 0) { k--; t = String.fromCharCode(97 + k % 26) + t; k = Math.floor(k / 26); } return t; };
  const roman = (k) => { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let t = ''; for (const [v, c] of m) while (k >= v) { t += c; k -= v; } return t; };
  switch (type) {
    case 'arabicPeriod': return n + '.';
    case 'arabicParenR': return n + ')';
    case 'arabicParenBoth': return '(' + n + ')';
    case 'arabicPlain': return '' + n;
    case 'alphaLcPeriod': return alpha(n) + '.';
    case 'alphaUcPeriod': return alpha(n).toUpperCase() + '.';
    case 'alphaLcParenR': return alpha(n) + ')';
    case 'alphaUcParenR': return alpha(n).toUpperCase() + ')';
    case 'romanLcPeriod': return roman(n) + '.';
    case 'romanUcPeriod': return roman(n).toUpperCase() + '.';
    default: return n + '.';
  }
}

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/* ek nakli lstStyle jo sirf default text colour set karta hai (shape style / table style ke liye) */
function colorStyle(c) {
  const root = document.createElementNS(NS_A, 'a:lstStyle');
  const hex = ((1 << 24) + (c.r << 16) + (c.g << 8) + c.b).toString(16).slice(1);
  for (let i = 1; i <= 9; i++) {
    const l = document.createElementNS(NS_A, 'a:lvl' + i + 'pPr');
    const d = document.createElementNS(NS_A, 'a:defRPr');
    const f = document.createElementNS(NS_A, 'a:solidFill');
    const clr = document.createElementNS(NS_A, 'a:srgbClr');
    clr.setAttribute('val', hex);
    f.appendChild(clr); d.appendChild(f); l.appendChild(d); root.appendChild(l);
  }
  return root;
}

/* txBody -> {el: .pp-in, ...}.  opts: chain[], rc, fontColor (css/obj), cellDefault, number */
function renderTextBody(txBody, rc, opts) {
  const env = rc.env;
  const inner = h('div', 'flex:0 0 auto;width:100%;', 'pp-in');
  const lst = kid(txBody, 'lstStyle');
  const chain = (opts.chain || []).concat(lst ? [lst] : []);
  const counters = [];
  let first = true;
  const S = (v) => `calc(var(--fs,1)*${v.toFixed(3)}px)`;
  let lastLvlAuto = [];

  for (const p of kids(txBody, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = clamp(numAttr(pPr, 'lvl', 0), 0, 8);
    const P = resolvePPr(chain, lvl, pPr);
    const endRPr = kid(p, 'endParaRPr');

    // gather runs
    const items = [];
    for (const c of kids(p)) {
      if (c.localName === 'r') items.push({ k: 'r', rPr: kid(c, 'rPr'), t: (kid(c, 't') || {}).textContent || '' });
      else if (c.localName === 'br') items.push({ k: 'br', rPr: kid(c, 'rPr') });
      else if (c.localName === 'fld') {
        let t = (kid(c, 't') || {}).textContent || '';
        if (/slidenum/i.test(attr(c, 'type', ''))) t = String(rc.number);
        items.push({ k: 'r', rPr: kid(c, 'rPr'), t });
      }
    }
    const hasText = items.some((it) => it.k === 'r' && it.t.length > 0);

    // base size
    let baseRprs = P.rprs.slice();
    if (!items.length && endRPr) baseRprs = baseRprs.concat([endRPr]);
    const firstR = items.find((it) => it.k === 'r');
    const sizeOf = (rprs) => {
      const v = rv(rprs, 'sz');
      return (v !== null ? parseFloat(v) / 100 : 18) * PT;
    };
    const baseSzPx = firstR ? sizeOf(P.rprs.concat([firstR.rPr])) : sizeOf(baseRprs);

    const para = h('div', '', 'pp-p');
    const st = para.style;
    st.textAlign = ALGN[P.algn] || 'left';
    const marL = px(parseFloat(P.marL || 0)), ind = px(parseFloat(P.indent || 0));
    if (P.rtl === '1') { st.direction = 'rtl'; st.paddingRight = marL + 'px'; } else { st.paddingLeft = marL + 'px'; }
    st.textIndent = ind + 'px';
    st.whiteSpace = opts.nowrap ? 'pre' : 'pre-wrap';
    st.wordBreak = 'normal';
    st.overflowWrap = 'break-word';
    // line spacing
    const lsEl = P.lnSpc;
    let lh = 1.2;
    if (lsEl) {
      const pct = kid(lsEl, 'spcPct'), pts = kid(lsEl, 'spcPts');
      if (pct) lh = numAttr(pct, 'val', 100000) / 100000 * 1.2;
      else if (pts) lh = null;
      if (pts) st.lineHeight = S(numAttr(pts, 'val', 0) / 100 * PT);
    }
    if (lh !== null) { lh = lh * (1 - (opts.lnSpcReduction || 0)); st.lineHeight = String(lh.toFixed(3)); }
    const spc = (el) => {
      if (!el) return 0;
      const pct = kid(el, 'spcPct'), pts = kid(el, 'spcPts');
      if (pct) return numAttr(pct, 'val', 0) / 100000 * baseSzPx * 1.2;
      if (pts) return numAttr(pts, 'val', 0) / 100 * PT;
      return 0;
    };
    if (!first) { const b = spc(P.spcBef); if (b) st.marginTop = S(b); }
    const a = spc(P.spcAft); if (a) st.marginBottom = S(a);

    // text colour defaults
    const defColor = opts.fontColor || null;
    const mkRunStyle = (span, rPr) => {
      const rprs = P.rprs.concat([rPr]);
      const sz = rv(rprs, 'sz');
      let px_ = (sz !== null ? parseFloat(sz) / 100 : 18) * PT;
      const base = rv(rprs, 'baseline');
      if (base !== null && parseFloat(base) !== 0) { span.style.verticalAlign = parseFloat(base) > 0 ? 'super' : 'sub'; px_ *= 0.7; }
      span.style.fontSize = S(px_);
      const b = rv(rprs, 'b');
      if (b === '1' || b === 'true' || (b === null && opts.cellDefault && opts.cellDefault.bold)) span.style.fontWeight = '700';
      const it = rv(rprs, 'i');
      if (it === '1' || it === 'true') span.style.fontStyle = 'italic';
      const u = rv(rprs, 'u'), stk = rv(rprs, 'strike');
      const deco = [];
      if (u !== null && u !== 'none') deco.push('underline');
      if (stk !== null && stk !== 'noStrike') deco.push('line-through');
      if (deco.length) span.style.textDecoration = deco.join(' ');
      const cap = rv(rprs, 'cap');
      if (cap === 'all') span.style.textTransform = 'uppercase';
      else if (cap === 'small') span.style.fontVariant = 'small-caps';
      const sp = rv(rprs, 'spc');
      if (sp !== null) span.style.letterSpacing = (parseFloat(sp) / 100 * PT) + 'px';
      const lat = rchild(rprs, 'latin');
      span.style.fontFamily = fontStack(lat ? attr(lat, 'typeface') : null, env.theme);
      const fill = rfill(rprs);
      let col = null;
      if (fill && fill.localName === 'solidFill') col = resolveColor(fill, env);
      else if (fill && fill.localName === 'gradFill') col = resolveColor(kid(kid(fill, 'gsLst'), 'gs'), env);
      else if (!fill && defColor) col = defColor;
      if (col) span.style.color = cssColor(col);
      else if (fill && fill.localName === 'noFill') span.style.color = 'transparent';
      const hl = rchild(rprs, 'highlight');
      if (hl) { const hc = resolveColor(hl, env); if (hc) span.style.backgroundColor = cssColor(hc); }
      return { px: px_, color: col };
    };

    // bullet
    let firstRunInfo = null;
    if (hasText) {
      const frag = [];
      for (const it of items) {
        if (it.k === 'br') { frag.push(document.createElement('br')); continue; }
        if (!it.t) continue;
        const span = document.createElement('span');
        const info = mkRunStyle(span, it.rPr);
        if (!firstRunInfo) firstRunInfo = info;
        span.textContent = it.t.replace(/\u000b/g, '\n');
        frag.push(span);
      }
      if (P.bu && P.bu.t !== 'none') {
        let label = '';
        if (P.bu.t === 'char') {
          label = P.bu.ch;
          const bf = (P.buFont || '').toLowerCase();
          if (bf.indexOf('wingdings') >= 0 || bf.indexOf('symbol') >= 0) label = WING[label] || (bf.indexOf('symbol') >= 0 && label === '·' ? '•' : '•');
        } else {
          for (let l = lvl + 1; l < counters.length; l++) counters[l] = undefined;
          if (counters[lvl] === undefined || !lastLvlAuto[lvl]) counters[lvl] = P.bu.start;
          else counters[lvl] += 1;
          label = autoNum(P.bu.type, counters[lvl]);
        }
        lastLvlAuto[lvl] = P.bu.t === 'auto';
        const bu = h('span', '', 'pp-bu');
        bu.textContent = label;
        bu.style.display = 'inline-block';
        bu.style.textIndent = '0';
        bu.style.width = (ind < 0 ? -ind : 0) + 'px';
        bu.style.minWidth = (ind < 0 ? 0 : 0) + 'px';
        bu.style.whiteSpace = 'pre';
        const bsz = (firstRunInfo ? firstRunInfo.px : baseSzPx) * (P.buSz || 1);
        bu.style.fontSize = S(bsz);
        bu.style.fontFamily = P.bu.t === 'char' && P.buFont && !/wingdings|symbol/i.test(P.buFont) ? fontStack(P.buFont, env.theme) : 'Arial,"Noto Sans",sans-serif';
        const bc = P.buClr ? resolveColor(P.buClr, env) : (firstRunInfo && firstRunInfo.color) || defColor;
        if (bc) bu.style.color = cssColor(bc);
        para.appendChild(bu);
        if (ind >= 0) { const g = h('span', 'display:inline-block;width:0.5em'); para.appendChild(g); }
      } else {
        lastLvlAuto[lvl] = false;
      }
      for (let l = lvl + 1; l < lastLvlAuto.length; l++) lastLvlAuto[l] = false;
      frag.forEach((n) => para.appendChild(n));
    } else {
      const span = document.createElement('span');
      mkRunStyle(span, endRPr);
      span.textContent = '\u200b';
      para.appendChild(span);
    }
    inner.appendChild(para);
    first = false;
  }
  return inner;
}

/* ------------------------------------------------------------------ fills (div / css side) */
function gradCss(gf, env, phClr) {
  const stops = kids(kid(gf, 'gsLst'), 'gs').map((g) => {
    const c = resolveColor(g, env, phClr);
    return { pos: numAttr(g, 'pos', 0) / 1000, col: cssColor(c) };
  }).filter((x) => x.col);
  if (!stops.length) return null;
  if (stops.length === 1) return stops[0].col;
  const lin = kid(gf, 'lin'), path = kid(gf, 'path');
  const list = stops.map((x) => `${x.col} ${x.pos.toFixed(1)}%`).join(',');
  if (path) return `radial-gradient(circle at 50% 50%,${list})`;
  const ang = lin ? numAttr(lin, 'ang', 0) / 60000 : 90;
  return `linear-gradient(${(ang + 90) % 360}deg,${list})`;
}
function pattMix(fe, env, phClr) {
  const fg = resolveColor(kid(fe, 'fgClr'), env, phClr) || { r: 0, g: 0, b: 0, a: 1 };
  const bg = resolveColor(kid(fe, 'bgClr'), env, phClr) || { r: 255, g: 255, b: 255, a: 1 };
  const m = /^pct(\d+)$/.exec(attr(fe, 'prst', ''));
  const k = m ? parseInt(m[1], 10) / 100 : 0.35;
  const mix = (a, b) => Math.round(a * k + b * (1 - k));
  return { r: mix(fg.r, bg.r), g: mix(fg.g, bg.g), b: mix(fg.b, bg.b), a: 1 };
}
function fillToCss(fe, env, part, phClr) {
  if (!fe) return null;
  switch (fe.localName) {
    case 'solidFill': return cssColor(resolveColor(fe, env, phClr));
    case 'gradFill': return gradCss(fe, env, phClr);
    case 'pattFill': return cssColor(pattMix(fe, env, phClr));
    case 'blipFill': {
      const url = blipUrl(fe, part);
      if (!url) return null;
      return kid(fe, 'tile') ? `url(${url}) 0 0/auto repeat` : `url(${url}) center/100% 100% no-repeat`;
    }
    default: return null;
  }
}
function blipUrl(fe, part) {
  const b = kid(fe, 'blip');
  const id = rAttr(b, 'embed');
  return id && part.imgs ? part.imgs[id] || null : null;
}

/* ------------------------------------------------------------------ placeholders */
function normType(t) {
  if (t === 'ctrTitle') return 'title';
  if (!t || t === 'subTitle' || t === 'obj' || t === 'chart' || t === 'tbl' || t === 'clipArt' || t === 'media' || t === 'dgm' || t === 'pic') return 'body';
  return t;
}
function phOf(el) {
  let nv = null;
  if (el.localName === 'sp') nv = at(el, 'nvSpPr', 'nvPr');
  else if (el.localName === 'pic') nv = at(el, 'nvPicPr', 'nvPr');
  else if (el.localName === 'graphicFrame') nv = at(el, 'nvGraphicFramePr', 'nvPr');
  return kid(nv, 'ph');
}
function findPh(tree, ph, isMaster) {
  if (!tree) return null;
  const type = attr(ph, 'type', null), idx = attr(ph, 'idx', null);
  let byIdx = null, byType = null;
  for (const el of kids(tree)) {
    const p = phOf(el);
    if (!p) continue;
    const t = attr(p, 'type', null), i = attr(p, 'idx', null);
    if (!isMaster && idx !== null && i === idx && !byIdx) byIdx = el;
    if (!byType && normType(t) === normType(type) && (isMaster || idx === null || i === null || true)) byType = byType || el;
  }
  if (!isMaster && byIdx) return byIdx;
  if (!isMaster && idx !== null && type === null) return byType;
  return byIdx || byType;
}
function firstWith(list, f) { for (const x of list) { const r = x ? f(x) : null; if (r) return r; } return null; }

/* ------------------------------------------------------------------ rendering */
function renderTree(tree, rc, container, opts) {
  for (const el of kids(tree)) {
    let node = el;
    if (el.localName === 'AlternateContent') {
      const fb = kid(el, 'Fallback'), ch = kid(el, 'Choice');
      const src = (fb && fb.firstElementChild) ? fb : ch;
      if (!src) continue;
      for (const sub of kids(src)) renderOne(sub, rc, container, opts);
      continue;
    }
    renderOne(node, rc, container, opts);
  }
}
function renderOne(el, rc, container, opts) {
  try {
    let out = null;
    switch (el.localName) {
      case 'sp': if (opts && opts.skipPh && phOf(el)) return; out = renderSp(el, rc, false); break;
      case 'cxnSp': out = renderSp(el, rc, true); break;
      case 'pic': if (opts && opts.skipPh && phOf(el)) return; out = renderPic(el, rc); break;
      case 'grpSp': out = renderGroup(el, rc, opts); break;
      case 'graphicFrame': out = renderFrame(el, rc); break;
      default: break;
    }
    if (out) container.appendChild(out);
  } catch (e) {
    if (global.console) console.warn('shape skipped:', e);
  }
}

function xfrmBox(xfrm) {
  const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
  return {
    x: px(numAttr(off, 'x', 0)), y: px(numAttr(off, 'y', 0)),
    w: px(numAttr(ext, 'cx', 0)), h: px(numAttr(ext, 'cy', 0)),
    rot: numAttr(xfrm, 'rot', 0) / 60000,
    fh: attr(xfrm, 'flipH') === '1', fv: attr(xfrm, 'flipV') === '1'
  };
}
function placeBox(box, b, applyRot) {
  box.style.left = b.x + 'px'; box.style.top = b.y + 'px';
  box.style.width = b.w + 'px'; box.style.height = b.h + 'px';
  if (applyRot && b.rot) box.style.transform = `rotate(${b.rot}deg)`;
}

function renderSp(sp, rc, isCxn) {
  const env = rc.env;
  const spPr = kid(sp, 'spPr');
  const style = kid(sp, 'style');
  const txBody = kid(sp, 'txBody');
  const ph = phOf(sp);
  let lay = null, mas = null;
  if (ph && rc.part === env.slidePart) {
    lay = findPh(at(env.layout.xml, 'cSld', 'spTree'), ph, false);
    mas = findPh(at(env.master.xml, 'cSld', 'spTree'), ph, true);
  }
  const spPrs = [spPr, lay ? kid(lay, 'spPr') : null, mas ? kid(mas, 'spPr') : null];
  const xfrm = firstWith(spPrs, (x) => kid(x, 'xfrm'));
  if (!xfrm) return null;
  const b = xfrmBox(xfrm);

  const box = h('div', 'position:absolute;', 'pp-sh');
  placeBox(box, b, true);

  /* geometry */
  let paths = null;
  const cust = firstWith(spPrs, (x) => kid(x, 'custGeom'));
  const prstEl = firstWith(spPrs, (x) => kid(x, 'prstGeom'));
  if (cust && kid(spPr, 'custGeom')) paths = custGeomPaths(cust, b.w, b.h);
  else if (prstEl && (kid(spPr, 'prstGeom') || !kid(spPr, 'custGeom'))) {
    const av = {};
    for (const g of kids(at(prstEl, 'avLst'), 'gd')) {
      const m = /val\s+(-?\d+)/.exec(attr(g, 'fmla', ''));
      if (m) av[attr(g, 'name')] = parseInt(m[1], 10);
    }
    paths = presetPaths(attr(prstEl, 'prst', 'rect'), b.w, b.h, av);
  } else if (cust) paths = custGeomPaths(cust, b.w, b.h);
  else if (isCxn) paths = presetPaths('line', b.w, b.h, {});

  /* fill + line */
  const fillRef = kid(style, 'fillRef'), lnRef = kid(style, 'lnRef'), fontRef = kid(style, 'fontRef');
  const fillEl = firstWith(spPrs, (x) => {
    for (const c of kids(x)) if (['solidFill', 'gradFill', 'blipFill', 'pattFill', 'noFill'].indexOf(c.localName) >= 0) return c;
    return null;
  });
  let phFill = null;
  if (fillRef && numAttr(fillRef, 'idx', 0) > 0) phFill = resolveColor(fillRef, env);
  const lnEl = firstWith(spPrs, (x) => kid(x, 'ln'));
  let stroke = null;
  {
    let w = null, col = null, none = false;
    if (lnRef && numAttr(lnRef, 'idx', 0) > 0) {
      col = resolveColor(lnRef, env);
      const i = numAttr(lnRef, 'idx', 1) - 1;
      w = env.theme.lnW[i] !== undefined ? env.theme.lnW[i] : 9525;
    }
    if (lnEl) {
      if (kid(lnEl, 'noFill')) none = true;
      else {
        const sf = kid(lnEl, 'solidFill'), gf = kid(lnEl, 'gradFill');
        if (sf) col = resolveColor(sf, env, col ? [col.r, col.g, col.b] : null);
        else if (gf) col = resolveColor(kid(kid(gf, 'gsLst'), 'gs'), env);
        if (lnEl.getAttribute('w') !== null) w = numAttr(lnEl, 'w', 12700);
        else if (w === null && col) w = 12700;
      }
    }
    if (!none && col && w !== null) {
      stroke = { w: Math.max(px(w), 0.5), color: col, ln: lnEl };
    }
  }

  /* svg */
  const svg = s('svg', { width: Math.max(b.w, 1), height: Math.max(b.h, 1), viewBox: `0 0 ${Math.max(b.w, 1)} ${Math.max(b.h, 1)}` });
  svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;';
  const tf = [];
  if (b.fh) tf.push('scaleX(-1)');
  if (b.fv) tf.push('scaleY(-1)');
  if (tf.length) svg.style.transform = tf.join(' ');
  const defs = s('defs');
  svg.appendChild(defs);

  let svgFill = 'none', svgFillOpacity = null;
  if (fillEl && fillEl.localName !== 'noFill') {
    if (fillEl.localName === 'solidFill' || fillEl.localName === 'pattFill') {
      const c = fillEl.localName === 'solidFill' ? resolveColor(fillEl, env, phFill ? [phFill.r, phFill.g, phFill.b] : null) : pattMix(fillEl, env, null);
      if (c) { svgFill = `rgb(${c.r},${c.g},${c.b})`; if (c.a < 1) svgFillOpacity = c.a; }
    } else if (fillEl.localName === 'gradFill') {
      const gid = 'g' + (++UID);
      const lin = kid(fillEl, 'lin'), pth = kid(fillEl, 'path');
      let g;
      if (pth) g = s('radialGradient', { id: gid, cx: 0.5, cy: 0.5, r: 0.5 });
      else {
        const th = (lin ? numAttr(lin, 'ang', 0) / 60000 : 90) * Math.PI / 180;
        g = s('linearGradient', { id: gid, x1: 0.5 - Math.cos(th) / 2, y1: 0.5 - Math.sin(th) / 2, x2: 0.5 + Math.cos(th) / 2, y2: 0.5 + Math.sin(th) / 2 });
      }
      for (const gs of kids(kid(fillEl, 'gsLst'), 'gs')) {
        const c = resolveColor(gs, env);
        if (!c) continue;
        g.appendChild(s('stop', { offset: numAttr(gs, 'pos', 0) / 100000, 'stop-color': `rgb(${c.r},${c.g},${c.b})`, 'stop-opacity': c.a }));
      }
      defs.appendChild(g);
      svgFill = `url(#${gid})`;
    } else if (fillEl.localName === 'blipFill') {
      const url = blipUrl(fillEl, rc.part);
      if (url) {
        const pid = 'p' + (++UID);
        const pat = s('pattern', { id: pid, patternUnits: 'objectBoundingBox', width: 1, height: 1 });
        pat.appendChild(s('image', { href: url, x: 0, y: 0, width: Math.max(b.w, 1), height: Math.max(b.h, 1), preserveAspectRatio: 'none' }));
        defs.appendChild(pat);
        svgFill = `url(#${pid})`;
      }
    }
  } else if (!fillEl && phFill && !isCxn) {
    svgFill = `rgb(${phFill.r},${phFill.g},${phFill.b})`;
  }

  /* arrow heads */
  let markerStart = null, markerEnd = null;
  if (stroke && stroke.ln) {
    const mk = (kindEl, start) => {
      const t = attr(kindEl, 'type', 'none');
      if (t === 'none' || !kindEl) return null;
      const id = 'm' + (++UID);
      const m = s('marker', { id, viewBox: '0 0 10 10', refX: t === 'triangle' || t === 'arrow' || t === 'stealth' ? 8 : 5, refY: 5, markerWidth: 4, markerHeight: 4, orient: start ? 'auto-start-reverse' : 'auto' });
      const col = `rgb(${stroke.color.r},${stroke.color.g},${stroke.color.b})`;
      if (t === 'oval') m.appendChild(s('circle', { cx: 5, cy: 5, r: 4, fill: col }));
      else if (t === 'diamond') m.appendChild(s('path', { d: 'M5,0 L10,5 L5,10 L0,5 Z', fill: col }));
      else m.appendChild(s('path', { d: 'M0,0 L10,5 L0,10 Z', fill: col }));
      defs.appendChild(m);
      return `url(#${id})`;
    };
    markerStart = mk(kid(stroke.ln, 'headEnd'), true);
    markerEnd = mk(kid(stroke.ln, 'tailEnd'), false);
  }

  const dashMap = { dash: '4 3', sysDash: '3 1', dot: '1 2', sysDot: '1 1', lgDash: '8 3', dashDot: '4 3 1 3', lgDashDot: '8 3 1 3', lgDashDotDot: '8 3 1 3 1 3', sysDashDot: '3 1 1 1', sysDashDotDot: '3 1 1 1 1 1' };
  (paths || []).forEach((pt, idx) => {
    const path = s('path', { d: pt.d });
    if (pt.mirror) path.setAttribute('transform', `translate(${b.w},0) scale(-1,1)`);
    const noFill = pt.open || pt.noFill || svgFill === 'none';
    if (noFill) path.setAttribute('fill', 'none');
    else {
      path.setAttribute('fill', svgFill);
      if (svgFillOpacity !== null) path.setAttribute('fill-opacity', svgFillOpacity);
      if (pt.evenodd) path.setAttribute('fill-rule', 'evenodd');
    }
    if (stroke && !pt.noStroke) {
      path.setAttribute('stroke', `rgb(${stroke.color.r},${stroke.color.g},${stroke.color.b})`);
      if (stroke.color.a < 1) path.setAttribute('stroke-opacity', stroke.color.a);
      path.setAttribute('stroke-width', stroke.w);
      path.setAttribute('stroke-linejoin', 'round');
      const dash = stroke.ln ? attr(kid(stroke.ln, 'prstDash'), 'val', 'solid') : 'solid';
      if (dashMap[dash]) path.setAttribute('stroke-dasharray', dashMap[dash].split(' ').map((n) => n * Math.max(stroke.w, 1)).join(' '));
      if (idx === 0 || pt.open) {
        if (markerStart) path.setAttribute('marker-start', markerStart);
        if (markerEnd) path.setAttribute('marker-end', markerEnd);
      }
    }
    svg.appendChild(path);
  });

  /* shadow */
  const eff = firstWith(spPrs, (x) => at(x, 'effectLst', 'outerShdw'));
  if (false && eff && paths) {
    const dist = px(numAttr(eff, 'dist', 0)), dir = numAttr(eff, 'dir', 0) / 60000 * Math.PI / 180;
    const col = resolveColor(eff, env);
    if (col) svg.style.filter = `drop-shadow(${(Math.cos(dir) * dist).toFixed(1)}px ${(Math.sin(dir) * dist).toFixed(1)}px ${px(numAttr(eff, 'blurRad', 0)).toFixed(1)}px ${cssColor(col)})`;
  }
  if (paths) box.appendChild(svg);

  /* text */
  if (txBody && !isCxn) {
    const hasAny = kids(txBody, 'p').some((p) => kids(p).some((c) => c.localName === 'r' || c.localName === 'fld' || c.localName === 'br'));
    if (hasAny) {
      const bodyPrs = [kid(txBody, 'bodyPr'), lay ? at(lay, 'txBody', 'bodyPr') : null, mas ? at(mas, 'txBody', 'bodyPr') : null];
      const bp = (n) => { for (const x of bodyPrs) { const v = x ? x.getAttribute(n) : null; if (v !== null) return v; } return null; };
      const chain = [env.defTextStyle];
      if (fontRef) { const fc = resolveColor(fontRef, env); if (fc) chain.push(colorStyle(fc)); }
      if (ph) {
        const t = attr(ph, 'type', null);
        const key = (t === 'title' || t === 'ctrTitle') ? 'titleStyle' : (t === 'dt' || t === 'ftr' || t === 'sldNum') ? 'otherStyle' : 'bodyStyle';
        chain.push(kid(env.master.txStyles, key));
        if (mas) chain.push(kid(at(mas, 'txBody'), 'lstStyle'));
        if (lay) chain.push(kid(at(lay, 'txBody'), 'lstStyle'));
      }
      const autofitEl = firstWith(bodyPrs, (x) => kid(x, 'normAutofit') || kid(x, 'spAutoFit') || kid(x, 'noAutofit'));
      const norm = autofitEl && kid(autofitEl, 'normAutofit') ? kid(autofitEl, 'normAutofit') : null;
      const fontScale = norm ? numAttr(norm, 'fontScale', 100000) / 100000 : 1;
      const lnRed = norm ? numAttr(norm, 'lnSpcReduction', 0) / 100000 : 0;
      const wrap = bp('wrap') !== 'none';
      const inner = renderTextBody(txBody, rc, { chain, fontColor: null, nowrap: !wrap, lnSpcReduction: lnRed });
      const tx = h('div', '', 'pp-tx');
      const ts = tx.style;
      ts.position = 'absolute'; ts.left = '0'; ts.top = '0'; ts.right = '0'; ts.bottom = '0';
      ts.boxSizing = 'border-box'; ts.display = 'flex'; ts.flexDirection = 'column';
      const anchor = bp('anchor') || 't';
      ts.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
      const ins = (n, d) => px(bp(n) !== null ? parseFloat(bp(n)) : d);
      ts.paddingLeft = ins('lIns', 91440) + 'px'; ts.paddingRight = ins('rIns', 91440) + 'px';
      ts.paddingTop = ins('tIns', 45720) + 'px'; ts.paddingBottom = ins('bIns', 45720) + 'px';
      ts.setProperty('--fs', String(fontScale));
      const vert = bp('vert');
      if (vert === 'vert' || vert === 'eaVert') ts.writingMode = 'vertical-rl';
      else if (vert === 'vert270') { ts.writingMode = 'vertical-rl'; ts.transform = 'rotate(180deg)'; }
      if (autofitEl && norm) tx.setAttribute('data-fit', 'norm');
      else if (autofitEl && kid(autofitEl, 'spAutoFit')) tx.setAttribute('data-fit', 'sp');
      tx.appendChild(inner);
      box.appendChild(tx);
    }
  }
  return box;
}

function renderPic(pic, rc) {
  const env = rc.env;
  const spPr = kid(pic, 'spPr');
  const ph = phOf(pic);
  let lay = null, mas = null;
  if (ph && rc.part === env.slidePart) {
    lay = findPh(at(env.layout.xml, 'cSld', 'spTree'), ph, false);
    mas = findPh(at(env.master.xml, 'cSld', 'spTree'), ph, true);
  }
  const xfrm = firstWith([spPr, lay ? kid(lay, 'spPr') : null, mas ? kid(mas, 'spPr') : null], (x) => kid(x, 'xfrm'));
  if (!xfrm) return null;
  const b = xfrmBox(xfrm);
  const bf = kid(pic, 'blipFill');
  const url = blipUrl(bf, rc.part);
  if (!url) return null;
  const box = h('div', 'position:absolute;overflow:hidden;', 'pp-pic');
  placeBox(box, b, true);
  const src = kid(bf, 'srcRect');
  const l = numAttr(src, 'l', 0) / 100000, t = numAttr(src, 't', 0) / 100000, r = numAttr(src, 'r', 0) / 100000, bt = numAttr(src, 'b', 0) / 100000;
  const wd = 1 - l - r, ht = 1 - t - bt;
  const img = h('img', '');
  img.src = url;
  img.draggable = false;
  img.style.cssText = `position:absolute;display:block;max-width:none;width:${(100 / (wd || 1)).toFixed(3)}%;height:${(100 / (ht || 1)).toFixed(3)}%;left:${(-l / (wd || 1) * 100).toFixed(3)}%;top:${(-t / (ht || 1) * 100).toFixed(3)}%;`;
  const tf = [];
  if (b.fh) tf.push('scaleX(-1)');
  if (b.fv) tf.push('scaleY(-1)');
  if (tf.length) { box.style.transform = (box.style.transform ? box.style.transform + ' ' : '') + tf.join(' '); }
  box.appendChild(img);
  const geom = at(spPr, 'prstGeom');
  const prst = attr(geom, 'prst', 'rect');
  if (prst === 'ellipse') box.style.borderRadius = '50%';
  else if (prst === 'roundRect') {
    const av = /val\s+(\d+)/.exec(attr(at(geom, 'avLst', 'gd'), 'fmla', ''));
    box.style.borderRadius = (Math.min(b.w, b.h) * (av ? parseInt(av[1], 10) : 16667) / 100000) + 'px';
  }
  const ln = kid(spPr, 'ln');
  if (ln && kid(ln, 'solidFill')) {
    const c = resolveColor(kid(ln, 'solidFill'), env);
    if (c) box.style.boxShadow = `inset 0 0 0 ${Math.max(px(numAttr(ln, 'w', 12700)), 0.5)}px ${cssColor(c)}`;
  }
  return box;
}

function renderGroup(grp, rc, opts) {
  const xfrm = at(grp, 'grpSpPr', 'xfrm');
  if (!xfrm) return null;
  const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext'), cho = kid(xfrm, 'chOff'), che = kid(xfrm, 'chExt');
  const b = xfrmBox(xfrm);
  const box = h('div', 'position:absolute;', 'pp-grp');
  placeBox(box, b, true);
  const cw = px(numAttr(che, 'cx', numAttr(ext, 'cx', 0))), ch = px(numAttr(che, 'cy', numAttr(ext, 'cy', 0)));
  const inner = h('div', 'position:absolute;left:0;top:0;transform-origin:0 0;');
  inner.style.width = cw + 'px'; inner.style.height = ch + 'px';
  const sx = cw ? b.w / cw : 1, sy = ch ? b.h / ch : 1;
  inner.style.transform = `scale(${sx},${sy}) translate(${-px(numAttr(cho, 'x', 0))}px,${-px(numAttr(cho, 'y', 0))}px)`;
  renderTree(grp, rc, inner, opts);
  box.appendChild(inner);
  return box;
}

function renderFrame(gf, rc) {
  const env = rc.env;
  const xf = kid(gf, 'xfrm');
  if (!xf) return null;
  const b = xfrmBox(xf);
  const gd = at(gf, 'graphic', 'graphicData');
  const tbl = kid(gd, 'tbl');
  if (tbl) return renderTable(tbl, b, rc);
  const pic = findDeep(gd, 'pic');
  if (pic) { const o = renderPic(pic, rc); if (o) return o; }
  const uri = attr(gd, 'uri', '');
  const label = /chart/i.test(uri) ? 'चार्ट' : /diagram/i.test(uri) ? 'SmartArt' : 'ऑब्जेक्ट';
  const box = h('div', 'position:absolute;display:flex;align-items:center;justify-content:center;border:2px dashed rgba(128,128,128,.55);color:rgba(100,100,100,.9);font:600 20px sans-serif;text-align:center;background:rgba(128,128,128,.08);box-sizing:border-box;');
  placeBox(box, b, false);
  box.textContent = label + ' (इस ऐप में नहीं दिखता)';
  return box;
}

function renderTable(tbl, b, rc) {
  const env = rc.env;
  const wrap = h('div', 'position:absolute;', 'pp-tbl');
  placeBox(wrap, b, false);
  const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => px(numAttr(c, 'w', 0)));
  const rows = kids(tbl, 'tr');
  const tblPr = kid(tbl, 'tblPr');
  const flag = (n) => attr(tblPr, n) === '1' || attr(tblPr, n) === 'true';
  const firstRow = flag('firstRow'), lastRow = flag('lastRow'), firstCol = flag('firstCol'), lastCol = flag('lastCol'), bandRow = flag('bandRow'), bandCol = flag('bandCol');
  const sidEl = kid(tblPr, 'tableStyleId');
  const sid = sidEl ? sidEl.textContent.trim() : '';
  const style = sid ? (env.tableStyles[sid] || builtinTableStyle(sid, env)) : {};

  const table = h('table', 'border-collapse:collapse;table-layout:fixed;');
  table.style.width = cols.reduce((a, c) => a + c, 0) + 'px';
  const cg = document.createElement('colgroup');
  cols.forEach((w) => { const c = document.createElement('col'); c.style.width = w + 'px'; cg.appendChild(c); });
  table.appendChild(cg);

  const nR = rows.length, nC = cols.length;
  rows.forEach((tr, ri) => {
    const trEl = document.createElement('tr');
    trEl.style.height = px(numAttr(tr, 'h', 0)) + 'px';
    let ci = 0;
    kids(tr, 'tc').forEach((tc) => {
      const myCol = ci;
      ci += 1;
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') return;
      const td = document.createElement('td');
      const gs = numAttr(tc, 'gridSpan', 1), rs = numAttr(tc, 'rowSpan', 1);
      if (gs > 1) td.colSpan = gs;
      if (rs > 1) td.rowSpan = rs;
      // style parts that apply
      const parts = [];
      const add = (n) => { if (style[n]) parts.push(style[n]); };
      add('wholeTbl');
      if (bandRow && !(firstRow && ri === 0)) { const bi = ri - (firstRow ? 1 : 0); add(bi % 2 === 0 ? 'band1H' : 'band2H'); }
      if (bandCol) add((myCol - (firstCol ? 1 : 0)) % 2 === 0 ? 'band1V' : 'band2V');
      if (firstCol && myCol === 0) add('firstCol');
      if (lastCol && myCol + gs >= nC) add('lastCol');
      if (lastRow && ri === nR - 1) add('lastRow');
      if (firstRow && ri === 0) add('firstRow');
      let fill = null, bold = false, tcol = null;
      parts.forEach((p) => { if (p.fill) fill = p.fill; if (p.bold) bold = true; if (p.color) tcol = p.color; });
      const bd = {};
      const sides = [['left', myCol === 0 ? 'left' : 'insideV'], ['right', myCol + gs >= nC ? 'right' : 'insideV'], ['top', ri === 0 ? 'top' : 'insideH'], ['bottom', ri + rs >= nR ? 'bottom' : 'insideH']];
      parts.forEach((p) => { sides.forEach(([side, key]) => { if (p.bdr[key] !== undefined) bd[side] = p.bdr[key]; }); });
      const tcPr = kid(tc, 'tcPr');
      if (tcPr) {
        const f = ['solidFill', 'gradFill', 'blipFill', 'noFill'].map((n) => kid(tcPr, n)).find(Boolean);
        if (f) {
          if (f.localName === 'noFill') fill = 'none';
          else { const css = fillToCss(f, env, rc.part); if (css) fill = { css }; }
        }
        [['left', 'lnL'], ['right', 'lnR'], ['top', 'lnT'], ['bottom', 'lnB']].forEach(([side, n]) => {
          const ln = kid(tcPr, n);
          if (ln) bd[side] = lnFromEl(ln, env) || 'none';
        });
      }
      if (fill && fill !== 'none') td.style.background = fill.css || cssColor(fill);
      ['left', 'right', 'top', 'bottom'].forEach((side) => {
        const v = bd[side];
        const prop = 'border' + side.charAt(0).toUpperCase() + side.slice(1);
        if (v && v !== 'none') td.style[prop] = `${Math.max(v.w, 0.75)}px solid ${cssColor(v.color)}`;
        else td.style[prop] = 'none';
      });
      const anchor = attr(tcPr, 'anchor', 't');
      td.style.verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';
      td.style.padding = `${px(numAttr(tcPr, 'marT', 45720))}px ${px(numAttr(tcPr, 'marR', 91440))}px ${px(numAttr(tcPr, 'marB', 45720))}px ${px(numAttr(tcPr, 'marL', 91440))}px`;
      td.style.overflow = 'hidden';
      const tx = kid(tc, 'txBody');
      if (tx) {
        const inner = renderTextBody(tx, rc, { chain: [env.defTextStyle, tcol ? colorStyle(tcol) : null], cellDefault: { bold }, fontColor: null });
        td.appendChild(inner);
      }
      trEl.appendChild(td);
    });
    table.appendChild(trEl);
  });
  wrap.appendChild(table);
  return wrap;
}

/* ------------------------------------------------------------------ slide background */
function bgCss(cSld, env, part) {
  const bg = kid(cSld, 'bg');
  if (!bg) return null;
  const pr = kid(bg, 'bgPr');
  if (pr) {
    for (const c of kids(pr)) {
      if (['solidFill', 'gradFill', 'blipFill', 'pattFill'].indexOf(c.localName) >= 0) return fillToCss(c, env, part);
      if (c.localName === 'noFill') return 'transparent';
    }
  }
  const ref = kid(bg, 'bgRef');
  if (ref) return cssColor(resolveColor(ref, env));
  return null;
}

/* ------------------------------------------------------------------ fit pass */
function fitAll(root) {
  root.querySelectorAll('.pp-tx[data-fit]').forEach((tx) => {
    const inner = tx.firstElementChild;
    if (!inner) return;
    const cs = getComputedStyle(tx);
    const avail = tx.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    const mode = tx.getAttribute('data-fit');
    if (mode === 'norm') {
      let sc = parseFloat(tx.style.getPropertyValue('--fs')) || 1;
      let n = 0;
      while (inner.offsetHeight > avail + 1 && sc > 0.3 && n < 16) {
        sc -= 0.05; n++;
        tx.style.setProperty('--fs', sc.toFixed(2));
      }
    } else if (mode === 'sp') {
      const need = inner.offsetHeight + parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
      if (need > tx.clientHeight + 1) tx.parentElement.style.height = need + 'px';
    }
  });
}

/* ------------------------------------------------------------------ loader */
async function load(buf) {
  if (!global.JSZip) throw new Error('JSZip लोड नहीं हुआ');
  const zip = await global.JSZip.loadAsync(buf);
  const pkg = new Pkg(zip);
  const pres = await pkg.xml('ppt/presentation.xml');
  if (!pres) throw new Error('यह फ़ाइल सही .pptx नहीं लगती');
  const sz = kid(pres, 'sldSz');
  const W = px(numAttr(sz, 'cx', 9144000)), H = px(numAttr(sz, 'cy', 5143500));
  const prels = await pkg.rels('ppt/presentation.xml');
  const defTextStyle = kid(pres, 'defaultTextStyle');

  const slides = [];
  const ids = kids(kid(pres, 'sldIdLst'), 'sldId');
  for (let i = 0; i < ids.length; i++) {
    const rel = prels[rAttr(ids[i], 'id')];
    if (!rel) continue;
    const x = await pkg.xml(rel.target);
    if (!x) continue;
    slides.push({ path: rel.target, number: i + 1, hidden: attr(x, 'show') === '0' });
  }
  let list = slides.filter((x) => !x.hidden);
  if (!list.length) list = slides;
  if (!list.length) throw new Error('इसमें कोई स्लाइड नहीं मिली');

  const tsRoot = await pkg.xml('ppt/tableStyles.xml');
  const tmplCache = new Map();

  async function preload(part) {
    const rels = await pkg.rels(part.path);
    part.imgs = {};
    const walk = async (root) => {
      const bl = root.getElementsByTagNameNS('*', 'blip');
      for (let i = 0; i < bl.length; i++) {
        const id = rAttr(bl[i], 'embed');
        if (id && rels[id] && !(id in part.imgs)) part.imgs[id] = await pkg.imgUrl(rels[id].target);
      }
    };
    await walk(part.xml);
  }
  async function template(slidePath, srels) {
    const empty = { theme: { colors: {}, major: 'Calibri', minor: 'Calibri', lnW: [9525] }, clrMap: {} };
    let lpath = '', lxml = null, mpath = '', mxml = null, theme = empty.theme, clrMap = {};
    try {
      const lrel = Object.values(srels).find((r) => /\/slideLayout$/.test(r.type));
      if (lrel) {
        lpath = lrel.target;
        if (tmplCache.has(lpath)) return tmplCache.get(lpath);
        lxml = await pkg.xml(lpath);
        const lrels = await pkg.rels(lpath);
        const mrel = Object.values(lrels).find((r) => /\/slideMaster$/.test(r.type));
        if (mrel) {
          mpath = mrel.target;
          mxml = await pkg.xml(mpath);
          const mrels = await pkg.rels(mpath);
          const trel = Object.values(mrels).find((r) => /\/theme$/.test(r.type));
          const txml = trel ? await pkg.xml(trel.target) : null;
          if (txml) theme = parseTheme(txml);
        }
      }
      const cm = kid(mxml, 'clrMap');
      if (cm) for (const a of cm.attributes) clrMap[a.name] = a.value;
    } catch (e) { if (global.console) console.warn('template', e); }
    if (!Object.keys(theme.colors).length) {
      theme = { colors: { dk1: '000000', lt1: 'FFFFFF', dk2: '1F497D', lt2: 'EEECE1', accent1: '4F81BD', accent2: 'C0504D', accent3: '9BBB59', accent4: '8064A2', accent5: '4BACC6', accent6: 'F79646' }, major: theme.major, minor: theme.minor, lnW: theme.lnW };
    }
    if (!Object.keys(clrMap).length) clrMap = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink' };
    const base = { theme, clrMap };
    let tstyles = {};
    try { tstyles = tsRoot ? parseTableStyles(tsRoot, base) : {}; } catch (e) { /* ignore */ }
    const t = {
      theme, clrMap,
      layout: { xml: lxml, path: lpath },
      master: { xml: mxml, path: mpath, txStyles: kid(mxml, 'txStyles') },
      tableStyles: tstyles
    };
    try { if (lxml) await preload(t.layout); } catch (e) { t.layout.imgs = {}; }
    try { if (mxml) await preload(t.master); } catch (e) { t.master.imgs = {}; }
    t.layout.imgs = t.layout.imgs || {}; t.master.imgs = t.master.imgs || {};
    if (lpath) tmplCache.set(lpath, t);
    return t;
  }

  function ovrMap(xml) {
    const o = at(xml, 'clrMapOvr', 'overrideClrMapping');
    if (!o) return null;
    const m = {};
    for (const a of o.attributes) m[a.name] = a.value;
    return m;
  }

  async function renderSlide(i) {
    const info = list[i];
    const root = h('div', `position:relative;overflow:hidden;width:${W}px;height:${H}px;background:#fff;`, 'pp-slide');
    root._w = W; root._h = H;
    try {
      const sx = await pkg.xml(info.path);
      const srels = await pkg.rels(info.path);
      const t = await template(info.path, srels);
      const slidePart = { xml: sx, path: info.path, imgs: {} };
      try { await preload(slidePart); } catch (e) { /* images just missing */ }
      const env = { theme: t.theme, clrMap: t.clrMap, layout: t.layout, master: t.master, slidePart, defTextStyle, tableStyles: t.tableStyles };
      const layOvr = ovrMap(t.layout.xml);
      const sldOvr = ovrMap(sx);
      const envFor = (m) => Object.assign({}, env, { clrMap: m || t.clrMap });
      const cSld = kid(sx, 'cSld');
      const bgOf = (cs, part, e) => { try { return bgCss(cs, e, part); } catch (er) { return null; } };
      const bg = bgOf(cSld, slidePart, envFor(sldOvr || layOvr)) || bgOf(kid(t.layout.xml, 'cSld'), t.layout, envFor(layOvr)) || bgOf(kid(t.master.xml, 'cSld'), t.master, envFor(null));
      if (bg) root.style.background = bg;

      const showLayout = attr(sx, 'showMasterSp') !== '0';
      const showMaster = showLayout && attr(t.layout.xml, 'showMasterSp') !== '0';
      const mk = (part, e) => ({ env: e, part, number: info.number });
      const safe = (fn) => { try { fn(); } catch (er) { if (global.console) console.warn(er); } };
      if (showMaster && t.master.xml) safe(() => renderTree(at(t.master.xml, 'cSld', 'spTree'), mk(t.master, envFor(null)), root, { skipPh: true }));
      if (showLayout && t.layout.xml) safe(() => renderTree(at(t.layout.xml, 'cSld', 'spTree'), mk(t.layout, envFor(layOvr)), root, { skipPh: true }));
      safe(() => renderTree(at(sx, 'cSld', 'spTree'), mk(slidePart, envFor(sldOvr || layOvr)), root, {}));
    } catch (e) {
      if (global.console) console.warn('slide failed', e);
      const m = h('div', 'position:absolute;left:60px;top:60px;right:60px;font:28px sans-serif;color:#b00;');
      m.textContent = 'Could not draw this slide: ' + (e && e.message ? e.message : e);
      root.appendChild(m);
    }
    return root;
  }

  return {
    width: W,
    height: H,
    count: list.length,
    renderSlide,
    fit: fitAll,
    destroy() { pkg.urls.forEach((u) => URL.revokeObjectURL(u)); pkg.urls = []; }
  };
}

global.PptxRender = { load };
})(typeof window !== 'undefined' ? window : globalThis);
