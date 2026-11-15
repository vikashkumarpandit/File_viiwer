/* ppt-legacy.js — old binary .ppt files: reads the text of every slide.
   (Layout/pictures of the old format are not drawn; text only.) */
(function (global) {
'use strict';

/* ---------- minimal OLE2 / CFB reader ---------- */
function readCFB(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  for (let i = 0; i < 8; i++) if (u8[i] !== sig[i]) throw new Error('Not an old PowerPoint (.ppt) file.');
  const ssz = 1 << dv.getUint16(0x1E, true);
  const mssz = 1 << dv.getUint16(0x20, true);
  const dirStart = dv.getUint32(0x30, true);
  const cutoff = dv.getUint32(0x38, true);
  const miniStart = dv.getUint32(0x3C, true);
  const difStart = dv.getUint32(0x44, true);
  const nDif = dv.getUint32(0x48, true);
  const END = 0xFFFFFFFE;
  const secOff = (n) => (n + 1) * ssz;

  const fatSecs = [];
  for (let i = 0; i < 109; i++) { const v = dv.getUint32(0x4C + i * 4, true); if (v < 0xFFFFFFFA) fatSecs.push(v); }
  let d = difStart;
  for (let k = 0; k < nDif && d < 0xFFFFFFFA; k++) {
    const o = secOff(d);
    for (let i = 0; i < ssz / 4 - 1; i++) { const v = dv.getUint32(o + i * 4, true); if (v < 0xFFFFFFFA) fatSecs.push(v); }
    d = dv.getUint32(o + ssz - 4, true);
  }
  const fat = [];
  for (const sct of fatSecs) {
    const o = secOff(sct);
    if (o + ssz > u8.length) break;
    for (let i = 0; i < ssz / 4; i++) fat.push(dv.getUint32(o + i * 4, true));
  }
  function chain(start, table) {
    const out = [];
    let n = start, guard = 0;
    while (n < 0xFFFFFFFA && n !== END && guard++ < 5000000) { out.push(n); n = table[n]; if (n === undefined) break; }
    return out;
  }
  function readBig(start, size) {
    const secs = chain(start, fat);
    const out = new Uint8Array(secs.length * ssz);
    secs.forEach((sct, i) => { const o = secOff(sct); out.set(u8.subarray(o, Math.min(o + ssz, u8.length)), i * ssz); });
    return size != null ? out.subarray(0, size) : out;
  }
  const dir = readBig(dirStart);
  const dd = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const nl = dd.getUint16(o + 0x40, true);
    let name = '';
    for (let i = 0; i < Math.max(0, nl / 2 - 1); i++) name += String.fromCharCode(dd.getUint16(o + i * 2, true));
    entries.push({ name, type: dir[o + 0x42], start: dd.getUint32(o + 0x74, true), size: dd.getUint32(o + 0x78, true) });
  }
  const root = entries[0];
  let miniStream = null, miniFat = null;
  function ensureMini() {
    if (miniStream) return;
    miniStream = readBig(root.start, root.size);
    const mf = readBig(miniStart);
    const mdv = new DataView(mf.buffer, mf.byteOffset, mf.byteLength);
    miniFat = [];
    for (let i = 0; i + 4 <= mf.length; i += 4) miniFat.push(mdv.getUint32(i, true));
  }
  function stream(name) {
    const e = entries.find((x) => x.type === 2 && x.name === name);
    if (!e) return null;
    if (e.size < cutoff) {
      ensureMini();
      const secs = chain(e.start, miniFat);
      const out = new Uint8Array(secs.length * mssz);
      secs.forEach((sct, i) => out.set(miniStream.subarray(sct * mssz, sct * mssz + mssz), i * mssz));
      return out.subarray(0, e.size);
    }
    return readBig(e.start, e.size);
  }
  return { stream };
}

/* ---------- PowerPoint records ---------- */
function decodeLatin1(b) { let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]); return s; }
function decodeUtf16(b) { let s = ''; for (let i = 0; i + 1 < b.length; i += 2) s += String.fromCharCode(b[i] | (b[i + 1] << 8)); return s; }

function parseSlides(ppt) {
  const dv = new DataView(ppt.buffer, ppt.byteOffset, ppt.byteLength);
  const slides = [];      // from SlideListWithText (placeholders)
  const boxes = [];       // text boxes inside each Slide container
  let curBoxes = null;

  function walk(start, end, ctx) {
    let o = start;
    let lastKind = 1;
    while (o + 8 <= end) {
      const verInst = dv.getUint16(o, true);
      const type = dv.getUint16(o + 2, true);
      const len = dv.getUint32(o + 4, true);
      const bodyStart = o + 8;
      const bodyEnd = Math.min(bodyStart + len, end);
      const ver = verInst & 0xF, inst = verInst >> 4;
      if (ver === 0xF) {
        if (type === 4080) walk(bodyStart, bodyEnd, { slwt: inst === 0, box: false });
        else if (type === 1006) { curBoxes = []; boxes.push(curBoxes); walk(bodyStart, bodyEnd, { slwt: false, box: true }); curBoxes = null; }
        else walk(bodyStart, bodyEnd, ctx);
      } else {
        if (type === 1011 && ctx.slwt) slides.push([]);
        else if (type === 3999) lastKind = dv.getUint32(bodyStart, true);
        else if (type === 4000 || type === 4008) {
          const raw = ppt.subarray(bodyStart, bodyEnd);
          const text = type === 4000 ? decodeUtf16(raw) : decodeLatin1(raw);
          if (ctx.slwt && slides.length) slides[slides.length - 1].push({ kind: lastKind, text });
          else if (ctx.box && curBoxes) curBoxes.push({ kind: -1, text });
        }
      }
      o = bodyStart + len;
    }
  }
  walk(0, ppt.length, { slwt: false, box: false });
  const n = Math.max(slides.length, boxes.length);
  const out = [];
  for (let i = 0; i < n; i++) out.push((slides[i] || []).concat(boxes[i] || []));
  return out;
}

/* ---------- build a viewer document ---------- */
function load(buf) {
  const u8 = new Uint8Array(buf);
  const cfb = readCFB(u8);
  const ppt = cfb.stream('PowerPoint Document');
  if (!ppt) throw new Error('No PowerPoint data found inside this file.');
  let slides = parseSlides(ppt);
  slides = slides.map((items) => items.filter((it) => it.text && it.text.replace(/[\s\u0000]/g, '').length));
  if (!slides.length) throw new Error('No readable text found in this old .ppt file. Please save it as .pptx and try again.');
  const W = 1280, H = 720;
  const clean = (t) => t.replace(/\u000b/g, '\n').replace(/\r/g, '\n').replace(/\u0000/g, '').trim();

  function renderSlide(i) {
    const root = document.createElement('div');
    root.className = 'pp-slide legacy';
    root.style.cssText = `position:relative;overflow:hidden;width:${W}px;height:${H}px;background:#fff;color:#111;font-family:Calibri,Carlito,"Noto Sans",sans-serif;`;
    root._w = W; root._h = H;
    const items = slides[i];
    let title = items.find((x) => x.kind === 0 || x.kind === 6);
    const rest = items.filter((x) => x !== title);
    if (!title && rest.length && rest[0].text.length < 90) title = rest.shift();
    const total = rest.reduce((a, x) => a + x.text.length, 0);
    const bodySize = total > 1400 ? 20 : total > 900 ? 24 : total > 500 ? 28 : 34;

    const bar = document.createElement('div');
    bar.style.cssText = 'position:absolute;left:0;top:0;width:14px;height:100%;background:#ff6b4a;';
    root.appendChild(bar);
    if (title) {
      const t = document.createElement('div');
      t.style.cssText = 'position:absolute;left:70px;right:60px;top:40px;font-size:48px;font-weight:700;line-height:1.15;';
      t.textContent = clean(title.text);
      root.appendChild(t);
    }
    const b = document.createElement('div');
    b.style.cssText = `position:absolute;left:70px;right:60px;top:${title ? 170 : 60}px;bottom:50px;overflow:hidden;font-size:${bodySize}px;line-height:1.35;white-space:pre-wrap;`;
    for (const it of rest) {
      for (const line of clean(it.text).split('\n')) {
        if (!line.trim()) continue;
        const p = document.createElement('div');
        p.style.cssText = 'margin:0 0 .45em 0;padding-left:1.1em;text-indent:-1.1em;';
        p.textContent = (it.kind === 1 || it.kind === 5 ? '• ' : '') + line;
        if (it.kind !== 1 && it.kind !== 5) { p.style.paddingLeft = '0'; p.style.textIndent = '0'; }
        b.appendChild(p);
      }
    }
    root.appendChild(b);
    const note = document.createElement('div');
    note.style.cssText = 'position:absolute;left:70px;bottom:14px;font:16px sans-serif;color:#9a9a9a;';
    note.textContent = 'Old .ppt format – text only. Save as .pptx to see the full design.';
    root.appendChild(note);
    return Promise.resolve(root);
  }
  return { width: W, height: H, count: slides.length, renderSlide, fit() {}, destroy() {}, legacy: true };
}

global.PptLegacy = { load };
})(typeof window !== 'undefined' ? window : globalThis);
