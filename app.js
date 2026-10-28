/* AGD Signature Log: offline signature capture for Device_Maintenance_Log_Sheet2.xlsx.
 * The workbook is never round-tripped through a spreadsheet library. Only the parts
 * signing needs are patched as text inside the OOXML package (JSZip); every other
 * zip entry is written back byte-for-byte.
 */
'use strict';
(() => {

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const SHEET_NAME = 'Maintenance Log';
const HEADER_ROW = 11, FIRST_ROW = 12, LAST_ROW = 51, SUMMARY_ROW = 53;
const LAST_COL = 13; // M
const ROLES = {
  tech:  { sigCol: 'G', methodCol: 'H', label: 'Technician' },
  staff: { sigCol: 'J', methodCol: 'K', label: 'Staff' }
};
const METHOD = 'Drawn (Pen/Touch)';
const COMMENT_AUTHOR = 'AGD Signature Log';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const OUT_PREFIX = 'Device_Maintenance_Log_Sheet2_signed_';
const EMU_PER_PX = 9525;
const CELL_PAD = 2;          // px padding inside the cell
const INK_WIDTH = 3;         // css px on the pad
const MDW = 7;               // max digit width of Calibri 11, the workbook's default font

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const REL = {
  officeDocument: NS_R + '/officeDocument',
  drawing: NS_R + '/drawing',
  image: NS_R + '/image',
  comments: NS_R + '/comments',
  vml: NS_R + '/vmlDrawing'
};
const CT = {
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  png: 'image/png',
  rels: 'application/vnd.openxmlformats-package.relationships+xml'
};
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const EMPTY_RELS = XML_DECL + `<Relationships xmlns="${NS_PKG_REL}"></Relationships>`;

// Worksheet children that must come after <drawing> / <legacyDrawing> (CT_Worksheet order).
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
const AFTER_LEGACY = AFTER_DRAWING.slice(1);

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const escHtml = esc;
const colNum = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: +m[2] }; };
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml');
const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
const relsPathFor = (p) => dirOf(p) + '_rels/' + p.slice(p.lastIndexOf('/') + 1) + '.rels';

function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir + target).split('/');
  const out = [];
  for (const part of parts) {
    if (part === '..') out.pop();
    else if (part !== '.' && part !== '') out.push(part);
  }
  return out.join('/');
}

function relativePath(fromDir, toPath) {
  const a = fromDir.split('/').filter(Boolean), b = toPath.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}

function todayStamp() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function excelDate(serial) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()];
  return `${String(d.getUTCDate()).padStart(2, '0')}-${mon}-${d.getUTCFullYear()}`;
}

let toastTimer = 0;
function toast(msg, ms = 3500) {
  const t = $('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, ms);
}

// ---------------------------------------------------------------------------
// IndexedDB: unsaved work survives a crash, screen lock or app switch
// ---------------------------------------------------------------------------
const idb = (() => {
  let dbp = null;
  const open = () => dbp || (dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open('agd-signature-log', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('kv');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
  const tx = async (mode, fn) => {
    const db = await open();
    return new Promise((resolve, reject) => {
      const t = db.transaction('kv', mode);
      const req = fn(t.objectStore('kv'));
      t.oncomplete = () => resolve(req && req.result);
      t.onerror = () => reject(t.error);
    });
  };
  return {
    get: (k) => tx('readonly', (s) => s.get(k)),
    set: (k, v) => tx('readwrite', (s) => s.put(v, k)),
    del: (k) => tx('readwrite', (s) => s.delete(k))
  };
})();

// ---------------------------------------------------------------------------
// Package reading
// ---------------------------------------------------------------------------
function parseRels(xml) {
  const out = [];
  const re = /<Relationship\b((?:[^>"']|"[^"]*"|'[^']*')*?)\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    const attr = (n) => { const a = new RegExp(`\\s${n}="([^"]*)"`).exec(m[1]); return a ? a[1] : null; };
    out.push({ id: attr('Id'), type: attr('Type'), target: attr('Target'), mode: attr('TargetMode') });
  }
  return out;
}

async function readText(zip, path) {
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

async function findSheetPath(zip) {
  const rootRels = parseRels(await readText(zip, '_rels/.rels'));
  const wbRel = rootRels.find((r) => r.type === REL.officeDocument);
  const wbPath = resolvePath('', wbRel ? wbRel.target : 'xl/workbook.xml');
  const wbDoc = parseXml(await readText(zip, wbPath));
  const sheet = [...wbDoc.getElementsByTagNameNS(NS_MAIN, 'sheet')].find((s) => s.getAttribute('name') === SHEET_NAME);
  if (!sheet) throw new Error(`This workbook has no "${SHEET_NAME}" sheet.`);
  const rid = sheet.getAttributeNS(NS_R, 'id');
  const rel = parseRels(await readText(zip, relsPathFor(wbPath))).find((r) => r.id === rid);
  return resolvePath(dirOf(wbPath), rel.target);
}

async function readSharedStrings(zip) {
  const wbRelsPath = 'xl/_rels/workbook.xml.rels';
  const rel = parseRels((await readText(zip, wbRelsPath)) || '').find((r) => /\/sharedStrings$/.test(r.type));
  if (!rel) return [];
  const doc = parseXml(await readText(zip, resolvePath('xl/', rel.target)));
  return [...doc.getElementsByTagNameNS(NS_MAIN, 'si')].map(siText);
}

// Text of an <si>/<is>, ignoring phonetic runs.
function siText(el) {
  let s = '';
  for (const t of el.getElementsByTagNameNS(NS_MAIN, 't')) {
    if (t.parentNode.localName !== 'rPh') s += t.textContent;
  }
  return s;
}

function readCells(doc, sst) {
  const cells = {};
  for (const c of doc.getElementsByTagNameNS(NS_MAIN, 'c')) {
    const ref = c.getAttribute('r');
    const t = c.getAttribute('t');
    const vEl = c.getElementsByTagNameNS(NS_MAIN, 'v')[0];
    const v = vEl ? vEl.textContent : '';
    let text = '', num = null;
    if (t === 's') text = v === '' ? '' : (sst[+v] ?? '');
    else if (t === 'inlineStr') { const is = c.getElementsByTagNameNS(NS_MAIN, 'is')[0]; text = is ? siText(is) : ''; }
    else if (t === 'b') text = v === '1' ? 'TRUE' : v === '0' ? 'FALSE' : '';
    else if (t === 'str' || t === 'e') text = v;
    else if (v !== '') { num = +v; text = v; }
    cells[ref] = { text, num, hasFormula: !!c.getElementsByTagNameNS(NS_MAIN, 'f')[0] };
  }
  return cells;
}

function sheetGeometry(doc) {
  const fmt = doc.getElementsByTagNameNS(NS_MAIN, 'sheetFormatPr')[0];
  const defW = fmt && fmt.getAttribute('defaultColWidth') ? +fmt.getAttribute('defaultColWidth') : 9.140625;
  const defH = fmt && fmt.getAttribute('defaultRowHeight') ? +fmt.getAttribute('defaultRowHeight') : 15;
  const cols = [...doc.getElementsByTagNameNS(NS_MAIN, 'col')].map((c) => ({
    min: +c.getAttribute('min'), max: +c.getAttribute('max'),
    width: c.getAttribute('width') ? +c.getAttribute('width') : defW,
    hidden: c.getAttribute('hidden') === '1' || c.getAttribute('hidden') === 'true'
  }));
  const rows = new Map();
  for (const r of doc.getElementsByTagNameNS(NS_MAIN, 'row')) {
    rows.set(+r.getAttribute('r'), {
      ht: r.getAttribute('ht') ? +r.getAttribute('ht') : null,
      hidden: r.getAttribute('hidden') === '1' || r.getAttribute('hidden') === 'true'
    });
  }
  const widthToPx = (w) => Math.trunc(((256 * w + Math.trunc(128 / MDW)) / 256) * MDW);
  const colPx = (n) => {
    const c = cols.find((x) => n >= x.min && n <= x.max);
    if (c && c.hidden) return 0;
    return widthToPx(c ? c.width : defW);
  };
  const rowPx = (r) => {
    const x = rows.get(r);
    if (x && x.hidden) return 0;
    return Math.round(((x && x.ht) || defH) * 96 / 72);
  };
  const colLeft = (n) => { let s = 0; for (let i = 1; i < n; i++) s += colPx(i); return s; };
  const rowTop = (r) => { let s = 0; for (let i = 1; i < r; i++) s += rowPx(i); return s; };
  return { colPx, rowPx, colLeft, rowTop };
}

// Pictures this app placed earlier ("Signature G12" etc.), so a re-opened file shows them.
async function readSignaturePictures(zip, sheetPath) {
  const pics = {};
  const sheetRels = parseRels((await readText(zip, relsPathFor(sheetPath))) || '');
  const dRel = sheetRels.find((r) => r.type === REL.drawing);
  if (!dRel) return pics;
  const drawingPath = resolvePath(dirOf(sheetPath), dRel.target);
  const drawingXml = await readText(zip, drawingPath);
  if (!drawingXml) return pics;
  const dRels = parseRels((await readText(zip, relsPathFor(drawingPath))) || '');
  const re = /name="Signature ([A-Z]+\d+)"[\s\S]*?:embed="([^"]+)"/g;
  let m;
  while ((m = re.exec(drawingXml))) {
    const rel = dRels.find((r) => r.id === m[2]);
    const file = rel && zip.file(resolvePath(dirOf(drawingPath), rel.target));
    if (file) pics[m[1]] = URL.createObjectURL(new Blob([await file.async('uint8array')], { type: CT.png }));
  }
  return pics;
}

async function readWorkbook(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const sheetPath = await findSheetPath(zip);
  const sst = await readSharedStrings(zip);
  const doc = parseXml(await readText(zip, sheetPath));
  if (doc.getElementsByTagName('parsererror').length) throw new Error('The log sheet could not be read.');
  return {
    sheetPath,
    cells: readCells(doc, sst),
    geo: sheetGeometry(doc),
    pictures: await readSignaturePictures(zip, sheetPath)
  };
}

// ---------------------------------------------------------------------------
// Package writing (text patches only)
// ---------------------------------------------------------------------------

// Depth-aware scan for the first direct child of the root element whose name is in `names`.
// Returns the index to insert before, or the index of the root close tag.
function childInsertIndex(xml, names) {
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let depth = 0, m;
  while ((m = re.exec(xml))) {
    const [, close, name, , selfClose] = m;
    if (close) { depth--; if (depth === 0) return m.index; continue; }
    if (depth === 1 && names.includes(name)) return m.index;
    if (!selfClose) depth++;
  }
  throw new Error('Malformed XML');
}

// Index just after a direct child of the root named `name`, or -1.
function childEndIndex(xml, name) {
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let depth = 0, m, inside = false;
  while ((m = re.exec(xml))) {
    const [, close, tag, , selfClose] = m;
    if (close) {
      depth--;
      if (inside && depth === 1 && tag === name) return re.lastIndex;
      continue;
    }
    if (depth === 1 && tag === name) {
      if (selfClose) return re.lastIndex;
      inside = true;
    }
    if (!selfClose) depth++;
  }
  return -1;
}

// Write a part without JSZip's implicit folder entries, which OOXML packages do not have.
const put = (zip, path, data) => zip.file(path, data, { createFolders: false });

function insertAt(s, i, frag) { return s.slice(0, i) + frag + s.slice(i); }

function nextRid(relsXml) {
  let max = 0;
  for (const r of parseRels(relsXml)) { const m = /^rId(\d+)$/.exec(r.id || ''); if (m) max = Math.max(max, +m[1]); }
  return 'rId' + (max + 1);
}

function addRel(relsXml, id, type, target) {
  const frag = `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/>`;
  if (/<Relationships\b[^>]*\/>/.test(relsXml)) return relsXml.replace(/<Relationships\b([^>]*)\/>/, `<Relationships$1>${frag}</Relationships>`);
  return relsXml.replace('</Relationships>', frag + '</Relationships>');
}

function removeRel(relsXml, id) {
  return relsXml.replace(new RegExp(`<Relationship\\b(?:[^>"']|"[^"]*"|'[^']*')*?\\sId="${reEsc(id)}"(?:[^>"']|"[^"]*"|'[^']*')*?\\/>`), '');
}

function ensureDefault(ct, ext, type) {
  if (new RegExp(`<Default\\b[^>]*\\sExtension="${ext}"`, 'i').test(ct)) return ct;
  return ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
}

function ensureOverride(ct, partName, type) {
  if (ct.includes(`PartName="${partName}"`)) return ct;
  return ct.replace('</Types>', `<Override PartName="${partName}" ContentType="${type}"/></Types>`);
}

function uniquePart(zip, prefix, ext) {
  let n = 1;
  while (zip.file(prefix + n + ext)) n++;
  return prefix + n + ext;
}

// Put a text value into one cell as an inline string, keeping its style.
function setCellText(sheet, ref, text) {
  const { col, row } = splitRef(ref);
  const sdStart = sheet.indexOf('<sheetData');
  const rowRe = new RegExp(`<row\\b[^>]*?\\sr="${row}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/row>)`);
  const rm = rowRe.exec(sheet.slice(sdStart));
  const newCell = (s) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t>${esc(text)}</t></is></c>`;

  if (!rm) {
    // Row absent: insert it in order.
    const rowXml = `<row r="${row}">${newCell(null)}</row>`;
    if (/<sheetData\s*\/>/.test(sheet)) return sheet.replace(/<sheetData\s*\/>/, `<sheetData>${rowXml}</sheetData>`);
    const rows = [...sheet.matchAll(/<row\b[^>]*?\sr="(\d+)"/g)];
    const after = rows.find((x) => +x[1] > row);
    return after ? insertAt(sheet, after.index, rowXml) : sheet.replace('</sheetData>', rowXml + '</sheetData>');
  }

  const rowStart = sdStart + rm.index;
  let rowXml = rm[0];
  if (rowXml.endsWith('/>') && !/<\/row>$/.test(rowXml)) rowXml = rowXml.slice(0, -2) + '></row>';
  const cellRe = new RegExp(`<c\\b[^>]*?\\sr="${ref}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const cm = cellRe.exec(rowXml);
  if (cm) {
    const s = (/\ss="(\d+)"/.exec(cm[0]) || [])[1];
    rowXml = rowXml.slice(0, cm.index) + newCell(s) + rowXml.slice(cm.index + cm[0].length);
  } else {
    const cells = [...rowXml.matchAll(/<c\b[^>]*?\sr="([A-Z]+)\d+"/g)];
    const after = cells.find((x) => colNum(x[1]) > colNum(col));
    rowXml = after ? insertAt(rowXml, after.index, newCell(null)) : rowXml.replace('</row>', newCell(null) + '</row>');
  }
  return sheet.slice(0, rowStart) + rowXml + sheet.slice(rowStart + rm[0].length);
}

// Row 53 counts a row as fully signed when B, G and J are non-blank. Refresh its cached value
// so viewers that do not recalculate on open still show the right number.
function refreshSignedCount(sheet, sst) {
  const cells = readCells(parseXml(sheet), sst);
  const filled = (ref) => !!(cells[ref] && cells[ref].text !== '');
  let n = 0;
  for (let r = FIRST_ROW; r <= LAST_ROW; r++) if (filled('B' + r) && filled('G' + r) && filled('J' + r)) n++;
  const re = new RegExp(`(<c\\b[^>]*?\\sr="[A-Z]+${SUMMARY_ROW}"[^>]*>\\s*<f>[^<]*G${FIRST_ROW}:G${LAST_ROW}[^<]*J${FIRST_ROW}:J${LAST_ROW}[^<]*<\\/f>)\\s*(?:<v>[^<]*<\\/v>|<v\\/>)?`);
  return sheet.replace(re, `$1<v>${n}</v>`);
}

async function addPictures(zip, ctx, sigs) {
  let { sheet, sheetRels, ct } = ctx;
  const sheetDir = dirOf(ctx.sheetPath);
  let dRel = parseRels(sheetRels).find((r) => r.type === REL.drawing);
  let drawingPath, drawingXml;
  if (dRel) {
    drawingPath = resolvePath(sheetDir, dRel.target);
    drawingXml = await readText(zip, drawingPath);
  } else {
    drawingPath = uniquePart(zip, 'xl/drawings/drawing', '.xml');
    drawingXml = XML_DECL + `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"></xdr:wsDr>`;
    const rid = nextRid(sheetRels);
    sheetRels = addRel(sheetRels, rid, REL.drawing, relativePath(sheetDir, drawingPath));
    sheet = insertAt(sheet, childInsertIndex(sheet, AFTER_DRAWING), `<drawing r:id="${rid}"/>`);
    ct = ensureOverride(ct, '/' + drawingPath, CT.drawing);
  }
  const drawingDir = dirOf(drawingPath);
  const dRelsPath = relsPathFor(drawingPath);
  let dRels = (await readText(zip, dRelsPath)) || EMPTY_RELS;

  const rootM = /<(?:(\w+):)?wsDr\b((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/.exec(drawingXml);
  const p = rootM[1] ? rootM[1] + ':' : '';
  // Declare a: and r: on the root once so new anchors can use them.
  let rootTag = `<${p}wsDr${rootM[2]}`;
  if (!/\sxmlns:a="/.test(rootM[2])) rootTag += ` xmlns:a="${NS_A}"`;
  if (!/\sxmlns:r="/.test(rootM[2])) rootTag += ` xmlns:r="${NS_R}"`;
  rootTag += '>' + (rootM[3] ? `</${p}wsDr>` : '');
  drawingXml = drawingXml.slice(0, rootM.index) + rootTag + drawingXml.slice(rootM.index + rootM[0].length);

  // Re-signing replaces the earlier picture instead of stacking another.
  const anchorRe = new RegExp(`<${reEsc(p)}(twoCellAnchor|oneCellAnchor|absoluteAnchor)\\b[\\s\\S]*?<\\/${reEsc(p)}\\1>`, 'g');
  const refs = new Set(sigs.map((s) => s.ref));
  drawingXml = drawingXml.replace(anchorRe, (anchor) => {
    const nm = /\sname="Signature ([A-Z]+\d+)"/.exec(anchor);
    if (!nm || !refs.has(nm[1])) return anchor;
    const emb = /:embed="([^"]+)"/.exec(anchor);
    if (emb) {
      const rel = parseRels(dRels).find((r) => r.id === emb[1]);
      dRels = removeRel(dRels, emb[1]);
      if (rel) {
        const media = resolvePath(drawingDir, rel.target);
        if (/\/signature\d+\.png$/.test(media)) zip.remove(media);
      }
    }
    return '';
  });

  let nextId = Math.max(1, ...[...drawingXml.matchAll(/cNvPr\b[^>]*?\sid="(\d+)"/g)].map((m) => +m[1])) + 1;
  let anchors = '';
  for (const sig of sigs) {
    const { col, row } = splitRef(sig.ref);
    const c = colNum(col);
    const cellW = ctx.geo.colPx(c), cellH = ctx.geo.rowPx(row);
    const boxW = cellW - 2 * CELL_PAD, boxH = cellH - 2 * CELL_PAD;
    let w = sig.w / 2, h = sig.h / 2;
    const k = Math.min(1, boxW / w, boxH / h);
    w *= k; h *= k;
    const offX = CELL_PAD + (boxW - w) / 2, offY = CELL_PAD + (boxH - h) / 2;
    const E = (px) => Math.round(px * EMU_PER_PX);

    const media = uniquePart(zip, 'xl/media/signature', '.png');
    put(zip, media, await sig.blob.arrayBuffer());
    const rid = nextRid(dRels);
    dRels = addRel(dRels, rid, REL.image, relativePath(drawingDir, media));

    anchors +=
      `<${p}twoCellAnchor editAs="oneCell">` +
      `<${p}from><${p}col>${c - 1}</${p}col><${p}colOff>${E(offX)}</${p}colOff><${p}row>${row - 1}</${p}row><${p}rowOff>${E(offY)}</${p}rowOff></${p}from>` +
      `<${p}to><${p}col>${c - 1}</${p}col><${p}colOff>${E(offX + w)}</${p}colOff><${p}row>${row - 1}</${p}row><${p}rowOff>${E(offY + h)}</${p}rowOff></${p}to>` +
      `<${p}pic><${p}nvPicPr><${p}cNvPr id="${nextId++}" name="Signature ${sig.ref}" descr="Signature of ${esc(sig.name)}"/>` +
      `<${p}cNvPicPr><a:picLocks noChangeAspect="1"/></${p}cNvPicPr></${p}nvPicPr>` +
      `<${p}blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></${p}blipFill>` +
      `<${p}spPr><a:xfrm><a:off x="${E(ctx.geo.colLeft(c) + offX)}" y="${E(ctx.geo.rowTop(row) + offY)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></${p}spPr></${p}pic><${p}clientData/></${p}twoCellAnchor>`;
  }
  drawingXml = drawingXml.replace(`</${p}wsDr>`, anchors + `</${p}wsDr>`);

  put(zip, drawingPath, drawingXml);
  put(zip, dRelsPath, dRels);
  ct = ensureDefault(ct, 'png', CT.png);
  ct = ensureDefault(ct, 'rels', CT.rels);
  return { ...ctx, sheet, sheetRels, ct };
}

const VML_HEAD = (idmap) =>
  `<xml xmlns:v="urn:schemas-microsoft-com:vml"\n xmlns:o="urn:schemas-microsoft-com:office:office"\n xmlns:x="urn:schemas-microsoft-com:office:excel">\n` +
  ` <o:shapelayout v:ext="edit">\n  <o:idmap v:ext="edit" data="${idmap}"/>\n </o:shapelayout>` + VML_SHAPETYPE;
const VML_SHAPETYPE =
  `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\n  path="m,l,21600r21600,l21600,xe">\n` +
  `  <v:stroke joinstyle="miter"/>\n  <v:path gradientshapeok="t" o:connecttype="rect"/>\n </v:shapetype>`;

function vmlNote(shapeId, c0, r0, z) {
  return `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style='position:absolute;\n` +
    `  margin-left:0;margin-top:0;width:160pt;height:48pt;z-index:${z};\n  visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">\n` +
    `  <v:fill color2="#ffffe1"/>\n  <v:shadow on="t" color="black" obscured="t"/>\n  <v:path o:connecttype="none"/>\n` +
    `  <v:textbox style='mso-direction-alt:auto'>\n   <div style='text-align:left'></div>\n  </v:textbox>\n` +
    `  <x:ClientData ObjectType="Note">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n` +
    `   <x:Anchor>\n    ${c0 + 1}, 15, ${Math.max(0, r0 - 1)}, 10, ${c0 + 3}, 15, ${r0 + 3}, 4</x:Anchor>\n` +
    `   <x:AutoFill>False</x:AutoFill>\n   <x:Row>${r0}</x:Row>\n   <x:Column>${c0}</x:Column>\n  </x:ClientData>\n </v:shape>`;
}

async function addComments(zip, ctx, sigs) {
  let { sheet, sheetRels, ct } = ctx;
  const sheetDir = dirOf(ctx.sheetPath);
  const rels = parseRels(sheetRels);
  const cRel = rels.find((r) => r.type === REL.comments);
  const legacy = /<legacyDrawing\b[^>]*?\sr:id="([^"]+)"/.exec(sheet);
  const vRel = legacy && rels.find((r) => r.id === legacy[1]);

  let commentsPath, commentsXml, vmlPath, vmlXml;
  if (cRel) {
    commentsPath = resolvePath(sheetDir, cRel.target);
    commentsXml = await readText(zip, commentsPath);
  } else {
    commentsPath = uniquePart(zip, 'xl/comments', '.xml');
    commentsXml = XML_DECL + `<comments xmlns="${NS_MAIN}"><authors></authors><commentList></commentList></comments>`;
    sheetRels = addRel(sheetRels, nextRid(sheetRels), REL.comments, relativePath(sheetDir, commentsPath));
    ct = ensureOverride(ct, '/' + commentsPath, CT.comments);
  }
  if (vRel) {
    vmlPath = resolvePath(sheetDir, vRel.target);
    vmlXml = await readText(zip, vmlPath);
    if (!vmlXml.includes('id="_x0000_t202"')) vmlXml = vmlXml.replace(/<\/o:shapelayout>/, '</o:shapelayout>' + VML_SHAPETYPE);
  } else {
    // Each VML part needs its own shape-id block across the whole workbook.
    let maxMap = 0;
    for (const name of Object.keys(zip.files).filter((n) => /\.vml$/i.test(n))) {
      const m = /<o:idmap\b[^>]*\sdata="([\d,\s]+)"/.exec(await readText(zip, name));
      if (m) maxMap = Math.max(maxMap, ...m[1].split(',').map(Number));
    }
    vmlPath = uniquePart(zip, 'xl/drawings/vmlDrawing', '.vml');
    vmlXml = VML_HEAD(maxMap + 1) + '</xml>';
    const rid = nextRid(sheetRels);
    sheetRels = addRel(sheetRels, rid, REL.vml, relativePath(sheetDir, vmlPath));
    const afterDrawing = childEndIndex(sheet, 'drawing');
    sheet = insertAt(sheet, afterDrawing >= 0 ? afterDrawing : childInsertIndex(sheet, AFTER_LEGACY), `<legacyDrawing r:id="${rid}"/>`);
    ct = ensureDefault(ct, 'vml', CT.vml);
  }

  // Author index
  if (/<authors\s*\/>/.test(commentsXml)) commentsXml = commentsXml.replace(/<authors\s*\/>/, '<authors></authors>');
  const authors = [...commentsXml.matchAll(/<author>([^<]*)<\/author>/g)].map((m) => m[1]);
  let authorId = authors.indexOf(esc(COMMENT_AUTHOR));
  if (authorId < 0) {
    authorId = authors.length;
    commentsXml = commentsXml.replace('</authors>', `<author>${esc(COMMENT_AUTHOR)}</author></authors>`);
  }
  if (/<commentList\s*\/>/.test(commentsXml)) commentsXml = commentsXml.replace(/<commentList\s*\/>/, '<commentList></commentList>');

  const idmap = +((/<o:idmap\b[^>]*\sdata="(\d+)/.exec(vmlXml) || [])[1] || 1);
  let shapeId = Math.max(idmap * 1024, ...[...vmlXml.matchAll(/id="_x0000_s(\d+)"/g)].map((m) => +m[1])) + 1;
  let z = [...vmlXml.matchAll(/<v:shape\b/g)].length + 1;

  let comments = '', shapes = '';
  for (const sig of sigs) {
    const { col, row } = splitRef(sig.ref);
    const c0 = colNum(col) - 1, r0 = row - 1;
    // Replace any earlier note on this cell.
    commentsXml = commentsXml.replace(new RegExp(`<comment\\s[^>]*?\\bref="${sig.ref}"[\\s\\S]*?<\\/comment>`, 'g'), '');
    vmlXml = vmlXml.replace(/<v:shape\b(?:(?!<\/v:shape>)[\s\S])*?<\/v:shape>/g, (shape) =>
      shape.includes('ObjectType="Note"') && shape.includes(`<x:Row>${r0}</x:Row>`) && shape.includes(`<x:Column>${c0}</x:Column>`) ? '' : shape);

    const body = `${ROLES[sig.role].label} signature\nSigned by: ${sig.name}\nSigned at: ${sig.ts}\nMethod: ${METHOD}`;
    comments += `<comment ref="${sig.ref}" authorId="${authorId}" shapeId="0"><text>` +
      `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>${esc(COMMENT_AUTHOR)}:</t></r>` +
      `<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">\n${esc(body)}</t></r>` +
      `</text></comment>`;
    shapes += vmlNote(shapeId++, c0, r0, z++);
  }
  commentsXml = commentsXml.replace('</commentList>', comments + '</commentList>');
  vmlXml = vmlXml.replace(/<\/xml>\s*$/, shapes + '</xml>');

  put(zip, commentsPath, commentsXml);
  put(zip, vmlPath, vmlXml);
  return { ...ctx, sheet, sheetRels, ct };
}

async function buildSignedWorkbook(bytes, sigMap) {
  const zip = await JSZip.loadAsync(bytes);
  const sheetPath = await findSheetPath(zip);
  const sheetRelsPath = relsPathFor(sheetPath);
  const origSheetRels = await readText(zip, sheetRelsPath);
  const origCt = await readText(zip, '[Content_Types].xml');
  let sheet = await readText(zip, sheetPath);
  const sst = await readSharedStrings(zip);
  const sigs = Object.values(sigMap).sort((a, b) => splitRef(a.ref).row - splitRef(b.ref).row || a.ref.localeCompare(b.ref));

  for (const sig of sigs) {
    const { row } = splitRef(sig.ref);
    sheet = setCellText(sheet, sig.ref, sig.name);
    sheet = setCellText(sheet, ROLES[sig.role].methodCol + row, METHOD);
  }
  sheet = refreshSignedCount(sheet, sst);

  let ctx = { sheetPath, sheet, sheetRels: origSheetRels || EMPTY_RELS, ct: origCt, geo: sheetGeometry(parseXml(sheet)) };
  ctx = await addPictures(zip, ctx, sigs);
  ctx = await addComments(zip, ctx, sigs);

  put(zip, sheetPath, ctx.sheet);
  if (ctx.sheetRels !== origSheetRels) put(zip, sheetRelsPath, ctx.sheetRels);
  if (ctx.ct !== origCt) put(zip, '[Content_Types].xml', ctx.ct);
  return zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME, compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

// ---------------------------------------------------------------------------
// Signature pad
// ---------------------------------------------------------------------------
const pad = (() => {
  const el = $('pad'), canvas = $('padCanvas'), area = $('padArea'), bar = $('padBar');
  const nameIn = $('padName'), msg = $('padMsg');
  const ctx2d = canvas.getContext('2d');
  let strokes = [], current = null, activeId = null, penSeen = false;
  let session = null; // { resolve, cellW, cellH, onChange }

  function size() {
    const r = area.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(r.width * dpr));
    canvas.height = Math.max(1, Math.round(r.height * dpr));
    ctx2d.setTransform(dpr, 0, 0, dpr, 0, 0);
    redraw();
  }

  function style(c, width) {
    c.lineCap = 'round';
    c.lineJoin = 'round';
    c.strokeStyle = '#000';
    c.fillStyle = '#000';
    c.lineWidth = width;
  }

  // Quadratic curves through the midpoints of successive samples.
  function drawStroke(c, pts, map, width) {
    style(c, width);
    const P = pts.map(map);
    if (P.length === 1) {
      c.beginPath(); c.arc(P[0].x, P[0].y, width / 2, 0, Math.PI * 2); c.fill();
      return;
    }
    c.beginPath();
    c.moveTo(P[0].x, P[0].y);
    for (let i = 1; i < P.length - 1; i++) {
      c.quadraticCurveTo(P[i].x, P[i].y, (P[i].x + P[i + 1].x) / 2, (P[i].y + P[i + 1].y) / 2);
    }
    c.lineTo(P[P.length - 1].x, P[P.length - 1].y);
    c.stroke();
  }

  const id = (p) => p;
  function redraw() {
    const r = area.getBoundingClientRect();
    ctx2d.clearRect(0, 0, r.width, r.height);
    for (const s of strokes) drawStroke(ctx2d, s.points, id, INK_WIDTH);
    if (current) drawStroke(ctx2d, current.points, id, INK_WIDTH);
  }

  // Incremental draw of the newest curve segment while the pointer moves.
  function drawTail() {
    const P = current.points, n = P.length;
    style(ctx2d, INK_WIDTH);
    ctx2d.beginPath();
    if (n === 2) {
      ctx2d.moveTo(P[0].x, P[0].y);
      ctx2d.lineTo((P[0].x + P[1].x) / 2, (P[0].y + P[1].y) / 2);
    } else {
      const a = P[n - 3], b = P[n - 2], c = P[n - 1];
      ctx2d.moveTo((a.x + b.x) / 2, (a.y + b.y) / 2);
      ctx2d.quadraticCurveTo(b.x, b.y, (b.x + c.x) / 2, (b.y + c.y) / 2);
    }
    ctx2d.stroke();
  }

  function point(e) {
    const r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  function changed() { if (session && session.onChange) session.onChange({ name: nameIn.value, strokes }); }

  canvas.addEventListener('pointerdown', (e) => {
    if (!session) return;
    // Palm rejection: once a pen has drawn, touch is ignored until the pad closes.
    if (e.pointerType === 'touch' && penSeen) return;
    if (e.pointerType === 'pen') {
      penSeen = true;
      if (current && current.type === 'touch') { current = null; activeId = null; redraw(); }
    }
    if (activeId !== null || e.button > 0) return;
    e.preventDefault();
    canvas.setPointerCapture(e.pointerId);
    activeId = e.pointerId;
    current = { type: e.pointerType, points: [point(e)] };
    msg.textContent = '';
  });

  canvas.addEventListener('pointermove', (e) => {
    if (e.pointerId !== activeId || !current) return;
    e.preventDefault();
    const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
    for (const ev of (evs.length ? evs : [e])) {
      const p = point(ev), last = current.points[current.points.length - 1];
      if (Math.abs(p.x - last.x) + Math.abs(p.y - last.y) < 0.75) continue;
      current.points.push(p);
      drawTail();
    }
  });

  function endStroke(e) {
    if (e.pointerId !== activeId) return;
    if (current) {
      strokes.push(current);
      current = null;
      redraw();
      changed();
    }
    activeId = null;
  }
  canvas.addEventListener('pointerup', endStroke);
  canvas.addEventListener('pointercancel', endStroke);
  // Belt and braces with touch-action:none: no scroll, zoom or pull-to-refresh from the pad.
  el.addEventListener('touchmove', (e) => { if (e.target === canvas) e.preventDefault(); }, { passive: false });
  nameIn.addEventListener('input', changed);

  // Keep the name bar inside the visible area when the on-screen keyboard opens.
  if (window.visualViewport) {
    const keep = () => { bar.style.transform = `translateY(${Math.max(0, window.visualViewport.offsetTop)}px)`; };
    window.visualViewport.addEventListener('resize', keep);
    window.visualViewport.addEventListener('scroll', keep);
  }
  nameIn.addEventListener('focus', () => { setTimeout(() => nameIn.scrollIntoView({ block: 'nearest' }), 300); });
  window.addEventListener('resize', () => { if (session) size(); });

  $('padClear').addEventListener('click', () => { strokes = []; current = null; redraw(); changed(); });
  $('padUndo').addEventListener('click', () => { strokes.pop(); redraw(); changed(); });
  $('padCancel').addEventListener('click', () => close(null));
  $('padAccept').addEventListener('click', async () => {
    const name = nameIn.value.trim().replace(/\s+/g, ' ');
    if (!strokes.length) { msg.textContent = 'Please sign in the box first.'; return; }
    if (!name) { msg.textContent = 'Enter the printed name.'; nameIn.focus(); return; }
    const png = await exportPng(session.cellW, session.cellH);
    close({ name, ...png });
  });

  // Trim to the ink, then render a transparent PNG at 2x the cell's inner size.
  async function exportPng(cellW, cellH) {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const s of strokes) for (const p of s.points) {
      minX = Math.min(minX, p.x); minY = Math.min(minY, p.y);
      maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y);
    }
    const rawW = Math.max(1, maxX - minX), rawH = Math.max(1, maxY - minY);
    const boxW = 2 * (cellW - 2 * CELL_PAD), boxH = 2 * (cellH - 2 * CELL_PAD);
    const s0 = Math.min(boxW / rawW, boxH / rawH);
    const lw = Math.max(2.4, INK_WIDTH * s0);
    const m = lw / 2 + 1;
    const scale = Math.min((boxW - 2 * m) / rawW, (boxH - 2 * m) / rawH);
    const w = Math.max(1, Math.ceil(rawW * scale + 2 * m)), h = Math.max(1, Math.ceil(rawH * scale + 2 * m));
    const out = document.createElement('canvas');
    out.width = w; out.height = h;
    const c = out.getContext('2d');
    const map = (p) => ({ x: (p.x - minX) * scale + m, y: (p.y - minY) * scale + m });
    for (const st of strokes) drawStroke(c, st.points, map, lw);
    const blob = await new Promise((res) => out.toBlob(res, 'image/png'));
    return { blob, w, h };
  }

  function close(result) {
    const s = session;
    session = null;
    el.hidden = true;
    nameIn.blur();
    strokes = []; current = null; activeId = null; penSeen = false;
    document.documentElement.style.overflow = '';
    if (s) s.resolve(result);
  }

  function open(opts) {
    return new Promise((resolve) => {
      session = { resolve, cellW: opts.cellW, cellH: opts.cellH, onChange: opts.onChange };
      $('padTitle').textContent = opts.title;
      $('padNameLabel').textContent = opts.nameLabel;
      nameIn.value = opts.name || '';
      msg.textContent = '';
      strokes = (opts.strokes || []).map((s) => ({ type: s.type, points: s.points.slice() }));
      current = null; activeId = null; penSeen = false;
      document.documentElement.style.overflow = 'hidden';
      el.hidden = false;
      requestAnimationFrame(size);
    });
  }

  return { open, isOpen: () => !!session };
})();

// ---------------------------------------------------------------------------
// App state and list view
// ---------------------------------------------------------------------------
const state = {
  fileName: '',
  bytes: null,   // ArrayBuffer of the workbook as picked (never written back)
  wb: null,      // parsed view of the log sheet
  sigs: {},      // ref -> { ref, role, name, blob, w, h, ts }
  output: null   // { blob, name, committed }
};

const cellText = (ref) => (state.wb && state.wb.cells[ref] ? state.wb.cells[ref].text : '');

// The picked workbook is kept with the signatures so a restart resumes where it left off.
async function persist() {
  if (state.bytes) await idb.set('session', { fileName: state.fileName, bytes: state.bytes, sigs: state.sigs });
}

function render() {
  const has = !!state.wb;
  $('emptyState').hidden = has;
  $('rowsTable').hidden = !has;
  $('printBtn').disabled = !has;
  const pending = Object.keys(state.sigs).length;
  $('saveBtn').disabled = !pending;
  $('saveBtn').textContent = pending ? `Save workbook (${pending})` : 'Save workbook';
  $('fileName').textContent = has ? state.fileName : 'No log open';
  if (!has) return;

  const status = (role, row) => {
    const ref = ROLES[role].sigCol + row;
    const sig = state.sigs[ref];
    const text = sig ? sig.name : cellText(ref);
    if (!text && !state.wb.pictures[ref]) return '<span class="status no">Not signed</span>';
    return `<span class="status yes">Signed<small>${escHtml(text)}${sig ? ' · unsaved' : ''}</small></span>`;
  };
  let html = '';
  for (let r = FIRST_ROW; r <= LAST_ROW; r++) {
    const serial = cellText('F' + r).trim();
    if (!serial) continue;
    html += `<tr data-row="${r}">` +
      `<td class="num">${escHtml(cellText('A' + r))}</td>` +
      `<td class="mono">${escHtml(serial)}</td>` +
      `<td>${escHtml(cellText('C' + r))}</td>` +
      `<td>${escHtml(cellText('D' + r))}</td>` +
      `<td>${escHtml(cellText('I' + r))}</td>` +
      `<td>${status('tech', r)}</td>` +
      `<td>${status('staff', r)}</td>` +
      `<td><div class="sign-cell"><button class="btn btn-primary" data-role="tech" data-row="${r}">Technician signs</button>` +
      `<button class="btn btn-primary" data-role="staff" data-row="${r}">Staff signs</button></div></td></tr>`;
  }
  $('rowsBody').innerHTML = html || '<tr><td colspan="8" class="empty">No rows with a serial number.</td></tr>';
}

async function loadBytes(fileName, bytes, sigs = {}) {
  const wb = await readWorkbook(bytes);
  if (state.wb) Object.values(state.wb.pictures).forEach((u) => URL.revokeObjectURL(u));
  Object.assign(state, { fileName, bytes, wb, sigs, output: null });
  render();
}

async function sign(role, row) {
  const ref = ROLES[role].sigCol + row;
  const c = colNum(ROLES[role].sigCol);
  const draftKey = { ref, role };
  const existing = state.sigs[ref];
  const result = await pad.open({
    title: `${ROLES[role].label} · Row ${row} · ${cellText('F' + row)}`,
    nameLabel: role === 'tech' ? 'Technician printed name' : 'Staff name',
    name: existing ? existing.name : role === 'staff' ? cellText('I' + row) : '',
    cellW: state.wb.geo.colPx(c),
    cellH: state.wb.geo.rowPx(row),
    onChange: (d) => { idb.set('draft', { ...draftKey, name: d.name, strokes: d.strokes }).catch(() => {}); }
  });
  await idb.del('draft').catch(() => {});
  if (!result) return;
  state.sigs[ref] = { ref, role, name: result.name, blob: result.blob, w: result.w, h: result.h, ts: new Date().toISOString() };
  await persist();
  render();
  toast(`${ROLES[role].label} signature saved for row ${row}. Tap Save workbook when done.`);
}

async function resumeDraft(draft) {
  if (!draft || !state.wb) return;
  const { row } = splitRef(draft.ref);
  const c = colNum(ROLES[draft.role].sigCol);
  const result = await pad.open({
    title: `${ROLES[draft.role].label} · Row ${row} · ${cellText('F' + row)} (restored)`,
    nameLabel: draft.role === 'tech' ? 'Technician printed name' : 'Staff name',
    name: draft.name,
    strokes: draft.strokes,
    cellW: state.wb.geo.colPx(c),
    cellH: state.wb.geo.rowPx(row),
    onChange: (d) => { idb.set('draft', { ...draft, name: d.name, strokes: d.strokes }).catch(() => {}); }
  });
  await idb.del('draft').catch(() => {});
  if (!result) return;
  state.sigs[draft.ref] = { ref: draft.ref, role: draft.role, name: result.name, blob: result.blob, w: result.w, h: result.h, ts: new Date().toISOString() };
  await persist();
  render();
}

$('rowsBody').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-role]');
  if (b) sign(b.dataset.role, +b.dataset.row);
});

$('fileInput').addEventListener('change', async (e) => {
  const file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const pending = Object.keys(state.sigs).length;
  if (pending && !confirm(`Discard ${pending} unsaved signature(s) and open ${file.name}?`)) return;
  try {
    const bytes = await file.arrayBuffer();
    await loadBytes(file.name, bytes, {});
    await persist();
    await idb.del('draft').catch(() => {});
  } catch (err) {
    console.error(err);
    alert('Could not open this file. Pick the maintenance log .xlsx.\n\n' + err.message);
  }
});

// ---------------------------------------------------------------------------
// Save, share, download
// ---------------------------------------------------------------------------
$('saveBtn').addEventListener('click', async () => {
  const btn = $('saveBtn');
  btn.disabled = true;
  btn.textContent = 'Building…';
  try {
    const blob = await buildSignedWorkbook(state.bytes, state.sigs);
    const name = OUT_PREFIX + todayStamp() + '.xlsx';
    state.output = { blob, name, committed: false };
    const file = new File([blob], name, { type: XLSX_MIME });
    $('shareBtn').hidden = !(navigator.canShare && navigator.canShare({ files: [file] }));
    $('saveName').textContent = name;
    $('saveStatus').textContent = '';
    $('saveDlg').hidden = false;
  } catch (err) {
    console.error(err);
    alert('Could not build the workbook: ' + err.message);
  } finally {
    render();
  }
});

// After the file has left the app, the signed file becomes the working copy.
async function commitSaved() {
  const out = state.output;
  if (!out || out.committed) return;
  out.committed = true;
  await loadBytes(out.name, await out.blob.arrayBuffer(), {});
  state.output = out;
  await persist(); // unsaved signatures cleared; the signed file is the new base
}

$('shareBtn').addEventListener('click', async () => {
  const out = state.output;
  const file = new File([out.blob], out.name, { type: XLSX_MIME });
  try {
    await navigator.share({ files: [file], title: out.name });
    $('saveStatus').textContent = 'Shared.';
    await commitSaved();
  } catch (err) {
    if (err.name !== 'AbortError') $('saveStatus').textContent = 'Share failed: ' + err.message + '. Use Download instead.';
  }
});

$('downloadBtn').addEventListener('click', async () => {
  const out = state.output;
  const url = URL.createObjectURL(out.blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = out.name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  $('saveStatus').textContent = 'Saved to Downloads.';
  await commitSaved();
});

$('saveClose').addEventListener('click', () => { $('saveDlg').hidden = true; });

// ---------------------------------------------------------------------------
// Print view
// ---------------------------------------------------------------------------
function buildPrintTable() {
  const wb = state.wb;
  const pending = {};
  for (const s of Object.values(state.sigs)) pending[s.ref] = URL.createObjectURL(s.blob);
  const filled = (ref) => !!(state.sigs[ref] || cellText(ref) !== '');

  let html = '<table class="ptable"><colgroup>';
  // Percent widths in the workbook's proportions, so the table always fits the landscape page.
  let total = 0;
  for (let c = 1; c <= LAST_COL; c++) total += wb.geo.colPx(c);
  for (let c = 1; c <= LAST_COL; c++) html += `<col style="width:${(100 * wb.geo.colPx(c) / total).toFixed(2)}%">`;
  html += '</colgroup><thead><tr>';
  for (let c = 1; c <= LAST_COL; c++) html += `<th>${escHtml(cellText(colName(c) + HEADER_ROW))}</th>`;
  html += '</tr></thead><tbody>';

  for (let r = FIRST_ROW; r <= LAST_ROW; r++) {
    html += '<tr class="data">';
    for (let c = 1; c <= LAST_COL; c++) {
      const col = colName(c), ref = col + r;
      if (col === 'G' || col === 'J') {
        const sig = state.sigs[ref];
        const img = sig ? pending[ref] : wb.pictures[ref];
        const name = sig ? sig.name : cellText(ref);
        html += img
          ? `<td class="sig"><img src="${img}" alt="Signature of ${escHtml(name)}"><span class="sig-name">${escHtml(name)}</span></td>`
          : `<td class="c">${escHtml(name)}</td>`;
        continue;
      }
      let v = cellText(ref);
      if (sigMethod(col, r)) v = METHOD;
      if (col === 'L' && wb.cells[ref] && wb.cells[ref].num !== null) v = excelDate(wb.cells[ref].num);
      const cls = ['A', 'E', 'H', 'K', 'L'].includes(col) ? ' class="c"' : '';
      html += `<td${cls}>${escHtml(v)}</td>`;
    }
    html += '</tr>';
  }

  let signed = 0;
  for (let r = FIRST_ROW; r <= LAST_ROW; r++) if (filled('B' + r) && filled('G' + r) && filled('J' + r)) signed++;
  const s = (ref) => escHtml(cellText(ref));
  html += `<tr><td class="nb" colspan="${LAST_COL}">&nbsp;</td></tr>` +
    `<tr class="summary"><td colspan="2" style="text-align:right">${s('A' + SUMMARY_ROW)}</td><td>${s('C' + SUMMARY_ROW)}</td>` +
    `<td style="text-align:right">${s('D' + SUMMARY_ROW)}</td><td>${signed}</td>` +
    `<td style="text-align:right">${s('F' + SUMMARY_ROW)}</td><td>${s('G' + SUMMARY_ROW)}</td><td colspan="${LAST_COL - 7}"></td></tr>`;
  html += '</tbody></table>';
  return html;
}

// H/K show the method of a pending (not yet saved) signature.
function sigMethod(col, row) {
  if (col === 'H') return !!state.sigs['G' + row];
  if (col === 'K') return !!state.sigs['J' + row];
  return false;
}

$('printBtn').addEventListener('click', () => {
  $('printContent').innerHTML = buildPrintTable();
  $('printView').hidden = false;
});
$('printNow').addEventListener('click', () => window.print());
$('printClose').addEventListener('click', () => {
  $('printView').hidden = true;
  $('printContent').querySelectorAll('img[src^="blob:"]').forEach((img) => {
    if (!Object.values(state.wb.pictures).includes(img.src)) URL.revokeObjectURL(img.src);
  });
  $('printContent').innerHTML = '';
});

// ---------------------------------------------------------------------------
// Startup
// ---------------------------------------------------------------------------
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch((err) => console.warn('Service worker not registered', err));
}

(async () => {
  try {
    const saved = await idb.get('session');
    if (saved && saved.bytes) {
      await loadBytes(saved.fileName, saved.bytes, saved.sigs || {});
      const n = Object.keys(state.sigs).length;
      if (n) toast(`Restored ${n} unsaved signature(s) for ${saved.fileName}.`);
      await resumeDraft(await idb.get('draft'));
    }
  } catch (err) {
    console.error('Could not restore session', err);
  }
})();

})();
