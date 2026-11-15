/* Doc Viewer — library + full-screen viewer for PPT, PDF, Word, Excel, text and images */
(() => {
'use strict';

const $ = (s) => document.querySelector(s);
const R = window.Readers;
const OLD_PPT = /\.(ppt|pps)$/i;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

/* ------------------------------------------------------------ IndexedDB */
let dbp = null;
function db() {
  if (dbp) return dbp;
  dbp = new Promise((res, rej) => {
    const r = indexedDB.open('ppt-reader', 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore('files', { keyPath: 'id', autoIncrement: true });
      d.createObjectStore('blobs');
      d.createObjectStore('meta');
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}
function tx(stores, mode, fn) {
  return db().then((d) => new Promise((res, rej) => {
    const t = d.transaction(stores, mode);
    let out;
    Promise.resolve(fn(t)).then((v) => { out = v; }, rej);
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error);
  }));
}
const req = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const dbAll = () => tx(['files'], 'readonly', (t) => req(t.objectStore('files').getAll()));
const dbBlob = (id) => tx(['blobs'], 'readonly', (t) => req(t.objectStore('blobs').get(id)));
const dbMetaGet = (k) => tx(['meta'], 'readonly', (t) => req(t.objectStore('meta').get(k)));
const dbMetaSet = (k, v) => tx(['meta'], 'readwrite', (t) => req(t.objectStore('meta').put(v, k)));
async function dbAdd(file) {
  const id = await tx(['files'], 'readwrite', (t) => req(t.objectStore('files').add({
    name: file.name, size: file.size, mod: file.lastModified || 0, added: Date.now()
  })));
  await tx(['blobs'], 'readwrite', (t) => req(t.objectStore('blobs').put(file, id)));
  return id;
}
const dbDel = (id) => tx(['files', 'blobs'], 'readwrite', (t) => {
  t.objectStore('files').delete(id);
  t.objectStore('blobs').delete(id);
});

/* ------------------------------------------------------------ helpers */
let toastTimer = 0;
function toast(msg, ms = 2800) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('on');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('on'), ms);
}
function fmtSize(b) {
  if (b < 1024 * 1024) return Math.max(1, Math.round(b / 1024)) + ' KB';
  return (b / 1024 / 1024).toFixed(1) + ' MB';
}
function fmtDate(ts) {
  try { return new Date(ts).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }); } catch (e) { return ''; }
}
const GROUPS = [
  { id: 'all', label: 'All', types: null },
  { id: 'slides', label: 'Slides', types: ['pptx', 'ppt'] },
  { id: 'pdf', label: 'PDF', types: ['pdf'] },
  { id: 'docs', label: 'Documents', types: ['docx', 'text'] },
  { id: 'sheets', label: 'Sheets', types: ['xlsx', 'csv'] },
  { id: 'images', label: 'Images', types: ['image'] }
];
const BADGE = { pptx: 'PPT', ppt: 'PPT', pdf: 'PDF', docx: 'DOC', xlsx: 'XLS', csv: 'CSV', text: 'TXT', image: 'IMG' };

/* ------------------------------------------------------------ library */
let records = [];
let group = 'all';
async function refresh() {
  records = (await dbAll()).sort((a, b) => b.added - a.added);
  render();
}
function render() {
  const q = $('#search').value.trim().toLowerCase();
  const list = $('#list');
  list.textContent = '';
  const g = GROUPS.find((x) => x.id === group) || GROUPS[0];
  const shown = records.filter((r) => (!q || r.name.toLowerCase().includes(q)) && (!g.types || g.types.includes(R.typeOf(r.name))));
  $('#empty').hidden = records.length > 0;
  $('#search').hidden = records.length < 6;
  // filter chips (only for groups that exist)
  const present = GROUPS.filter((x) => !x.types || records.some((r) => x.types.includes(R.typeOf(r.name))));
  const fb = $('#filters');
  fb.hidden = present.length < 3;
  fb.textContent = '';
  if (!present.some((x) => x.id === group)) group = 'all';
  present.forEach((x) => {
    const b = document.createElement('button');
    b.type = 'button'; b.textContent = x.label;
    if (x.id === group) b.className = 'on';
    b.addEventListener('click', () => { group = x.id; render(); });
    fb.appendChild(b);
  });
  for (const r of shown) {
    const type = R.typeOf(r.name) || 'text';
    const li = document.createElement('li');
    li.className = 'item';
    const ic = document.createElement('div');
    ic.className = 'ic t-' + type + (OLD_PPT.test(r.name) ? ' old' : '');
    ic.textContent = BADGE[type] || 'DOC';
    const b = document.createElement('button');
    b.className = 'tx'; b.type = 'button';
    const nm = document.createElement('div');
    nm.className = 'nm'; nm.textContent = r.name.replace(/\.[A-Za-z0-9]+$/, '');
    const mt = document.createElement('div');
    mt.className = 'mt'; mt.textContent = R.extOf(r.name).toUpperCase() + ' · ' + fmtSize(r.size) + ' · ' + fmtDate(r.added);
    b.append(nm, mt);
    b.addEventListener('click', () => openRecord(r, true));
    ic.addEventListener('click', () => openRecord(r, true));
    const rm = document.createElement('button');
    rm.className = 'rm'; rm.type = 'button'; rm.textContent = '🗑'; rm.setAttribute('aria-label', 'Remove');
    rm.addEventListener('click', async () => {
      if (!confirm('Remove "' + r.name + '" from this list?\n(The original file on your phone is not deleted.)')) return;
      await dbDel(r.id);
      await refresh();
    });
    li.append(ic, b, rm);
    list.appendChild(li);
  }
}
async function addFiles(fileList, quiet) {
  const all = [...fileList];
  const files = all.filter((f) => R.typeOf(f.name));
  const bad = all.filter((f) => !R.typeOf(f.name) && R.UNSUPPORTED[R.extOf(f.name)]);
  if (bad.length && !quiet) toast(bad.length + ' file(s) skipped: ' + R.UNSUPPORTED[R.extOf(bad[0].name)] + ' is not supported — save it as the newer format first.', 5000);
  if (!files.length) { if (!quiet && !bad.length) toast('No supported files found'); return []; }
  const have = new Set(records.map((r) => r.name + '|' + r.size));
  const added = [];
  for (const f of files) {
    if (have.has(f.name + '|' + f.size)) continue;
    try { added.push({ id: await dbAdd(f), name: f.name }); have.add(f.name + '|' + f.size); } catch (e) { console.warn(e); toast('Could not add "' + f.name + '" (storage full?)'); }
  }
  await refresh();
  if (!quiet && !bad.length) toast(added.length ? added.length + ' file(s) added ✓' : 'These files are already in the list');
  return added;
}

/* ---- pickers */
$('#btnAdd').addEventListener('click', () => $('#fileInput').click());
$('#fileInput').addEventListener('change', async (e) => { const f = [...e.target.files]; e.target.value = ''; if (f.length) await addFiles(f); });

let dirHandle = null;
async function scanDir(handle, out, depth) {
  for await (const [name, h] of handle.entries()) {
    if (h.kind === 'file' && R.typeOf(name)) out.push(await h.getFile());
    else if (h.kind === 'directory' && depth < 6 && name[0] !== '.') await scanDir(h, out, depth + 1);
  }
}
$('#btnFolder').addEventListener('click', async () => {
  if (window.showDirectoryPicker) {
    try {
      const h = await window.showDirectoryPicker({ mode: 'read' });
      toast('Scanning folder…', 6000);
      const out = [];
      await scanDir(h, out, 0);
      dirHandle = h;
      try { await dbMetaSet('dir', h); } catch (e) { /* ignore */ }
      showRescan();
      await addFiles(out);
    } catch (e) { if (e && e.name !== 'AbortError') { console.warn(e); $('#dirInput').click(); } }
  } else {
    $('#dirInput').click();
  }
});
$('#dirInput').addEventListener('change', async (e) => {
  const f = [...e.target.files]; e.target.value = '';
  if (f.length) await addFiles(f);
});
function showRescan() {
  const b = $('#btnRescan');
  b.hidden = !dirHandle;
  if (dirHandle) b.textContent = '🔄 Rescan "' + dirHandle.name + '"';
}
$('#btnRescan').addEventListener('click', async () => {
  if (!dirHandle) return;
  try {
    if (dirHandle.requestPermission && (await dirHandle.requestPermission({ mode: 'read' })) !== 'granted') { toast('Permission was not granted'); return; }
    toast('Scanning folder…', 6000);
    const out = [];
    await scanDir(dirHandle, out, 0);
    await addFiles(out);
  } catch (e) { console.warn(e); toast('Could not scan — please pick the folder again'); }
});
$('#search').addEventListener('input', render);

/* ---- install */
let deferredInstall = null;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredInstall = e; $('#btnInstall').hidden = false; });
$('#btnInstall').addEventListener('click', async () => {
  if (!deferredInstall) return;
  deferredInstall.prompt();
  try { await deferredInstall.userChoice; } catch (e) { /* ignore */ }
  deferredInstall = null; $('#btnInstall').hidden = true;
});
window.addEventListener('appinstalled', () => { $('#btnInstall').hidden = true; toast('App installed ✓'); });
const standalone = matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
if (!standalone) {
  const hint = $('#hint');
  hint.hidden = false;
  hint.textContent = 'Tip: in Chrome open the menu (⋮) → "Install app" / "Add to Home screen" to get the app icon on your phone.';
}

/* ------------------------------------------------------------ viewer */
const V = {
  open: false, mode: 'slides', doc: null, desc: null, idx: 0, shown: -1, wanted: 0, rendering: false, cache: new Map(),
  auto: null, idleT: 0, token: 0, fsTried: false, sheet: false, wake: null, ptr: null, type: '', z: 1
};
const viewer = $('#viewer'), stage = $('#stage'), reader = $('#reader');

function tryFullscreen() {
  const el = viewer;
  const f = el.requestFullscreen || el.webkitRequestFullscreen;
  if (f && !document.fullscreenElement) {
    try {
      const p = f.call(el, { navigationUI: 'hide' });
      if (p && p.then) p.then(lockLandscape, () => {});
    } catch (e) { /* ignore */ }
  }
}
function lockLandscape() {
  try {
    if (V.mode === 'slides' && V.doc && V.doc.width >= V.doc.height && screen.orientation && screen.orientation.lock) screen.orientation.lock('landscape').catch(() => {});
  } catch (e) { /* ignore */ }
}
function exitFullscreen() {
  try { if (screen.orientation && screen.orientation.unlock) screen.orientation.unlock(); } catch (e) { /* ignore */ }
  try { if (document.fullscreenElement) document.exitFullscreen(); } catch (e) { /* ignore */ }
}
async function wakeOn() { try { if (navigator.wakeLock) V.wake = await navigator.wakeLock.request('screen'); } catch (e) { /* ignore */ } }
function wakeOff() { try { if (V.wake) V.wake.release(); } catch (e) { /* ignore */ } V.wake = null; }
document.addEventListener('visibilitychange', () => { if (V.open && document.visibilityState === 'visible') wakeOn(); });

function poke() {
  viewer.classList.remove('idle');
  clearTimeout(V.idleT);
  V.idleT = setTimeout(() => viewer.classList.add('idle'), 2800);
}
function openUI() {
  V.open = true;
  viewer.hidden = false;
  document.body.style.overflow = 'hidden';
  history.pushState({ viewer: 1 }, '');
  poke();
}
function closeViewer(fromPop) {
  if (!V.open) return;
  V.open = false; V.token++;
  stopAuto();
  wakeOff();
  exitFullscreen();
  viewer.hidden = true;
  document.body.style.overflow = '';
  $('#vSheet').hidden = true; $('#vEnd').hidden = true; $('#vErr').hidden = true; $('#vHint').hidden = true;
  V.sheet = false;
  stage.textContent = ''; reader.textContent = ''; reader.hidden = true; stage.hidden = false;
  V.cache.clear();
  const d = V.doc, ds = V.desc;
  V.doc = null; V.desc = null;
  try { if (d && (!ds || d !== ds.slides)) d.destroy(); } catch (e) { /* ignore */ }
  try { if (ds) ds.destroy(); } catch (e) { /* ignore */ }
  V.shown = -1; V.rendering = false;
  if (!fromPop) { try { history.back(); } catch (e) { /* ignore */ } }
}
window.addEventListener('popstate', () => { if (V.open) closeViewer(true); });

function showErr(msg, detail) {
  $('#loading').hidden = true;
  $('#vErrMsg').textContent = msg;
  $('#vErrDetail').textContent = detail || '';
  $('#vErrDetail').hidden = !detail;
  $('#vErr').hidden = false;
}

function sniff(buf) {
  const m = new Uint8Array(buf.slice(0, 4));
  if (m[0] === 0xD0 && m[1] === 0xCF && m[2] === 0x11 && m[3] === 0xE0) return 'ole';
  if (m[0] === 0x50 && m[1] === 0x4B) return 'zip';
  if (m[0] === 0x25 && m[1] === 0x50 && m[2] === 0x44 && m[3] === 0x46) return 'pdf';
  return '';
}

async function openRecord(rec, fromTap) {
  if (V.open) return;
  const token = ++V.token;
  openUI();
  $('#loading').hidden = false;
  $('#loadMsg').textContent = 'Opening "' + rec.name.replace(/\.[A-Za-z0-9]+$/, '') + '"…';
  V.fsTried = false;
  if (fromTap) { tryFullscreen(); V.fsTried = true; }
  wakeOn();
  try {
    const blob = await dbBlob(rec.id);
    if (!blob) throw new Error('This file is missing. Remove it from the list and add it again.');
    const buf = await blob.arrayBuffer();
    let type = R.typeOf(rec.name);
    const kind = sniff(buf);
    if (kind === 'pdf') type = 'pdf';
    else if ((type === 'pptx' || type === 'ppt') && kind) type = kind === 'ole' ? 'ppt' : 'pptx';
    V.type = type;
    if (type === 'pptx') {
      const doc = await window.PptxRender.load(buf);
      if (token !== V.token || !V.open) { doc.destroy(); return; }
      V.desc = null;
      await startSlides(doc);
    } else if (type === 'ppt') {
      const doc = window.PptLegacy.load(buf);
      if (token !== V.token || !V.open) return;
      V.desc = null;
      await startSlides(doc);
    } else {
      const desc = await R.open(type, buf, rec.name);
      if (token !== V.token || !V.open) { desc.destroy(); return; }
      V.desc = desc;
      await startReader();
    }
    if (token !== V.token) return;
    $('#loading').hidden = true;
    if (V.mode === 'slides') {
      const h = $('#vHint');
      h.hidden = false; h.style.animation = 'none'; void h.offsetWidth; h.style.animation = '';
      setTimeout(() => { h.hidden = true; }, 3700);
    }
  } catch (e) {
    console.error(e);
    if (token === V.token) showErr(e && e.message ? e.message : 'Could not open this file.', e && e.stack ? String(e.stack).split('\n').slice(0, 4).join('\n') : '');
  }
}

/* ------------------------------------------------------------ slide mode */
async function startSlides(doc) {
  V.mode = 'slides';
  V.doc = doc; V.cache.clear(); V.shown = -1; V.wanted = 0;
  reader.hidden = true; stage.hidden = false; stage.textContent = '';
  const sl = $('#vSlider'); sl.max = String(doc.count); sl.value = '1';
  syncMenu();
  lockLandscape();
  layoutOnResize();
  await go(0);
}
async function slideWrap(i) {
  if (V.cache.has(i)) return V.cache.get(i);
  const doc = V.doc, token = V.token;
  const el = await doc.renderSlide(i);
  if (token !== V.token || V.doc !== doc) return null;
  const w = el._w || doc.width, h = el._h || doc.height;
  const wrap = document.createElement('div');
  wrap.className = 'sw';
  wrap.style.width = w + 'px'; wrap.style.height = h + 'px';
  wrap.dataset.w = w; wrap.dataset.h = h;
  wrap.appendChild(el);
  V.cache.set(i, wrap);
  for (const k of [...V.cache.keys()]) if (Math.abs(k - i) > 3) V.cache.delete(k);
  return wrap;
}
function scaleWrap(wrap) {
  const w = +wrap.dataset.w, h = +wrap.dataset.h;
  const vw = stage.clientWidth || window.innerWidth, vh = stage.clientHeight || window.innerHeight;
  const s = Math.min(vw / w, vh / h);
  wrap.style.transform = `translate(${(vw - w * s) / 2}px,${(vh - h * s) / 2}px) scale(${s})`;
}
let rzT = 0;
function layoutOnResize() {
  if (layoutOnResize.bound) return;
  layoutOnResize.bound = true;
  const f = () => { clearTimeout(rzT); rzT = setTimeout(() => { const w = stage.firstElementChild; if (w && V.mode === 'slides') scaleWrap(w); }, 60); };
  window.addEventListener('resize', f);
  window.addEventListener('orientationchange', f);
  document.addEventListener('fullscreenchange', f);
}
async function go(i) {
  if (!V.doc || V.mode !== 'slides') return;
  V.wanted = clamp(i, 0, V.doc.count - 1);
  if (V.rendering) return;
  V.rendering = true;
  try {
    while (V.open && V.doc && V.mode === 'slides' && V.shown !== V.wanted) {
      const idx = V.wanted;
      const wrap = await slideWrap(idx);
      if (!wrap || !V.open) break;
      scaleWrap(wrap);
      stage.textContent = '';
      wrap.style.animation = 'none'; void wrap.offsetWidth; wrap.style.animation = '';
      stage.appendChild(wrap);
      if (!wrap.dataset.fit) { wrap.dataset.fit = '1'; try { V.doc.fit(wrap); } catch (e) { console.warn(e); } }
      V.shown = idx;
      updateChrome();
      if (idx + 1 < V.doc.count) slideWrap(idx + 1).catch(() => {});
    }
  } catch (e) {
    console.error(e);
    toast('Could not draw this page');
  } finally { V.rendering = false; }
}
function updateChrome() {
  if (!V.doc || V.shown < 0) return;
  const n = V.doc.count, i = V.shown;
  $('#vCount').textContent = (V.auto ? '▶ ' : '') + (i + 1) + ' / ' + n;
  $('#vBar').firstElementChild.style.width = ((i + 1) / n * 100) + '%';
  $('#vSlider').value = String(i + 1);
  $('#vSliderLbl').textContent = (i + 1) + ' / ' + n;
}
function next() {
  if (!V.doc || V.mode !== 'slides') return;
  if (V.wanted >= V.doc.count - 1) { stopAuto(); $('#vEnd').hidden = false; return; }
  go(V.wanted + 1);
}
function prev() { if (V.doc && V.mode === 'slides' && V.wanted > 0) go(V.wanted - 1); }

function stopAuto() { if (V.auto) { clearInterval(V.auto); V.auto = null; } $('#vAuto').textContent = '▶ Auto-play'; updateChrome(); }
function startAuto() {
  stopAuto();
  V.auto = setInterval(() => {
    if (!V.doc || V.mode !== 'slides') return stopAuto();
    if (V.wanted >= V.doc.count - 1) { stopAuto(); return; }
    go(V.wanted + 1);
  }, 4500);
  $('#vAuto').textContent = '⏸ Stop auto-play';
  updateChrome();
}

/* ------------------------------------------------------------ reader mode */
function onPage(n, total) {
  $('#vCount').textContent = n + ' / ' + total;
  $('#vBar').firstElementChild.style.width = (n / total * 100) + '%';
}
async function startReader() {
  V.mode = 'reader';
  stage.hidden = true; stage.textContent = '';
  reader.hidden = false; reader.textContent = ''; reader.className = ''; reader.scrollTop = 0;
  reader.dataset.t = V.type; reader.style.zoom = '';
  V.z = 1;
  $('#vCount').textContent = ''; $('#vBar').firstElementChild.style.width = '0';
  syncMenu();
  let target = reader;
  if (!V.desc.scroller && !V.desc.pages) {
    target = document.createElement('div'); target.id = 'rdIn';
    reader.appendChild(target);
  }
  await V.desc.mount(target, onPage);
}
async function startPdfPages() {
  // PDF in "page mode": tap to turn pages
  V.mode = 'slides';
  reader.hidden = true; reader.textContent = '';
  stage.hidden = false; stage.textContent = '';
  const doc = V.desc.slides;
  V.doc = doc; V.cache.clear(); V.shown = -1;
  const sl = $('#vSlider'); sl.max = String(doc.count); sl.value = '1';
  syncMenu();
  layoutOnResize();
  $('#loading').hidden = false;
  await go(0);
  $('#loading').hidden = true;
  const h = $('#vHint'); h.hidden = false; h.style.animation = 'none'; void h.offsetWidth; h.style.animation = '';
  setTimeout(() => { h.hidden = true; }, 3700);
}
function syncMenu() {
  const slides = V.mode === 'slides';
  $('#rowSlider').hidden = !slides;
  $('#vAuto').hidden = !slides;
  const zoomable = !slides && V.desc && V.desc.zoomable;
  $('#vZoomIn').hidden = !zoomable; $('#vZoomOut').hidden = !zoomable;
  const m = $('#vMode');
  m.hidden = !(V.desc && V.desc.slides);
  m.textContent = slides ? 'Scroll mode' : 'Page mode';
  $('#vBar').style.display = '';
}
function zoomBy(f) {
  if (V.mode !== 'reader' || !V.desc) return;
  if (V.desc.zoom) { V.desc.zoom(f); return; }
  V.z = clamp(V.z * f, 0.5, 3);
  const t = document.getElementById('rdIn');
  if (t) t.style.zoom = V.z;
}

/* ------------------------------------------------------------ gestures */
function onDown(e) {
  if (V.mode !== 'slides' || e.target.closest('.chrome')) return;
  V.ptr = { x: e.clientX, y: e.clientY, t: Date.now(), id: e.pointerId };
}
function onUp(e) {
  const p = V.ptr; V.ptr = null;
  if (V.mode !== 'slides' || !p || e.target.closest('.chrome') || !$('#vErr').hidden || !$('#loading').hidden) return;
  poke();
  if (V.sheet) { toggleSheet(false); return; }
  const dx = e.clientX - p.x, dy = e.clientY - p.y;
  if (!V.fsTried) { V.fsTried = true; tryFullscreen(); }
  if (Math.abs(dx) > 48 && Math.abs(dx) > Math.abs(dy) * 1.3) { dx < 0 ? next() : prev(); return; }
  if (Math.abs(dx) < 14 && Math.abs(dy) < 14) {
    const w = stage.clientWidth || window.innerWidth;
    if (e.clientX < w * 0.25) prev(); else next();
  }
}
viewer.addEventListener('pointerdown', onDown);
viewer.addEventListener('pointerup', onUp);
viewer.addEventListener('pointercancel', () => { V.ptr = null; });
viewer.addEventListener('pointermove', (e) => { if (e.pointerType === 'mouse') poke(); });
reader.addEventListener('scroll', poke, { passive: true });
reader.addEventListener('pointerdown', poke);

document.addEventListener('keydown', (e) => {
  if (!V.open) return;
  if (V.mode === 'slides') {
    switch (e.key) {
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': e.preventDefault(); next(); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': e.preventDefault(); prev(); break;
      case 'Home': go(0); break;
      case 'End': if (V.doc) go(V.doc.count - 1); break;
      default: break;
    }
  }
  if (e.key === 'f' || e.key === 'F') { document.fullscreenElement ? exitFullscreen() : tryFullscreen(); }
  if (e.key === 'Escape' && !document.fullscreenElement) closeViewer(false);
  poke();
});

/* ------------------------------------------------------------ buttons */
function toggleSheet(on) {
  V.sheet = on === undefined ? !V.sheet : on;
  $('#vSheet').hidden = !V.sheet;
}
$('#vClose').addEventListener('click', () => closeViewer(false));
$('#vMenu').addEventListener('click', () => { poke(); toggleSheet(); });
$('#vSlider').addEventListener('input', (e) => { const i = parseInt(e.target.value, 10) - 1; $('#vSliderLbl').textContent = (i + 1) + ' / ' + V.doc.count; go(i); });
$('#vAuto').addEventListener('click', () => { V.auto ? stopAuto() : startAuto(); toggleSheet(false); });
$('#vFull').addEventListener('click', () => { document.fullscreenElement ? exitFullscreen() : tryFullscreen(); toggleSheet(false); });
$('#vHome').addEventListener('click', () => closeViewer(false));
$('#vZoomIn').addEventListener('click', () => zoomBy(1.25));
$('#vZoomOut').addEventListener('click', () => zoomBy(0.8));
$('#vMode').addEventListener('click', async () => {
  toggleSheet(false);
  if (!V.desc || !V.desc.slides) return;
  try { if (V.mode === 'slides') { stopAuto(); V.doc = null; await startReader(); } else await startPdfPages(); } catch (e) { console.error(e); toast('Could not switch mode'); }
});
$('#vAgain').addEventListener('click', () => { $('#vEnd').hidden = true; go(0); });
$('#vBack').addEventListener('click', () => closeViewer(false));
$('#vErrBack').addEventListener('click', () => closeViewer(false));
$('#vEnd').addEventListener('click', (e) => { if (e.target === $('#vEnd')) $('#vEnd').hidden = true; });

/* ------------------------------------------------------------ share target / file handler */
async function importShared() {
  if (!('caches' in window)) return;
  try {
    const c = await caches.open('ppt-shared-v1');
    const keys = await c.keys();
    if (!keys.length) return;
    const files = [];
    for (const k of keys) {
      const r = await c.match(k);
      if (!r) continue;
      const name = decodeURIComponent(r.headers.get('X-Name') || 'shared.pdf');
      const blob = await r.blob();
      files.push(new File([blob], name, { type: blob.type, lastModified: parseInt(r.headers.get('X-Mod') || '0', 10) || Date.now() }));
      await c.delete(k);
    }
    if (files.length) {
      const added = await addFiles(files, true);
      const id = added[0] ? added[0].id : (records.find((x) => x.name === files[0].name && x.size === files[0].size) || {}).id;
      const rec = records.find((x) => x.id === id);
      if (rec) openRecord(rec, false);
    }
  } catch (e) { console.warn('share import', e); }
  if (location.search) history.replaceState(null, '', location.pathname);
}
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (p) => {
    if (!p.files || !p.files.length) return;
    const files = [];
    for (const h of p.files) { try { files.push(await h.getFile()); } catch (e) { /* ignore */ } }
    if (!files.length) return;
    const added = await addFiles(files, true);
    const id = added[0] ? added[0].id : (records.find((x) => x.name === files[0].name && x.size === files[0].size) || {}).id;
    const rec = records.find((x) => x.id === id);
    if (rec) openRecord(rec, false);
  });
}

/* ------------------------------------------------------------ boot */
window.addEventListener('error', (e) => { console.warn('error', e.message); });
(async () => {
  if ('serviceWorker' in navigator) { navigator.serviceWorker.register('sw.js').catch((e) => console.warn('sw', e)); }
  try { dirHandle = (await dbMetaGet('dir')) || null; } catch (e) { dirHandle = null; }
  showRescan();
  await refresh();
  await importShared();
  if (navigator.storage && navigator.storage.persist) { try { navigator.storage.persist(); } catch (e) { /* ignore */ } }
})();

})();
