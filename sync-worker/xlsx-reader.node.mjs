// Node-compatible port of ../xlsx-reader.js. Same parsing rules and limits,
// using zlib + fast-xml-parser instead of browser DecompressionStream/DOMParser.
// Read-only: no macros or workbook formulas are executed.
import { inflateRawSync } from 'node:zlib';
import { XMLParser } from 'fast-xml-parser';

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  isArray: (name) => ['row', 'c', 'si', 't', 'sheet', 'Relationship', 'mergeCell', 'sheetData'].includes(name),
});

function col(ref) {
  let n = 0;
  for (const c of ref.match(/^[A-Z]+/i)?.[0] || '') n = n * 26 + c.toUpperCase().charCodeAt(0) - 64;
  return n - 1;
}

// Minimal central-directory ZIP reader (matches the browser unzip() logic).
function unzip(buffer) {
  const bytes = buffer instanceof Buffer ? buffer : Buffer.from(buffer);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (bytes.readUInt32LE(i) === 0x06054b50) { end = i; break; }
  }
  if (end < 0) throw new Error('This is not an XLSX workbook. Save older .xls files as .xlsx in Excel.');
  const entries = bytes.readUInt16LE(end + 10);
  let p = bytes.readUInt32LE(end + 16);
  const files = new Map();
  let total = 0;
  for (let i = 0; i < entries; i++) {
    if (bytes.readUInt32LE(p) !== 0x02014b50) throw new Error('Workbook ZIP directory is damaged.');
    const flags = bytes.readUInt16LE(p + 8);
    const method = bytes.readUInt16LE(p + 10);
    const size = bytes.readUInt32LE(p + 20);
    const unpacked = bytes.readUInt32LE(p + 24);
    const nameLen = bytes.readUInt16LE(p + 28);
    const extra = bytes.readUInt16LE(p + 30);
    const comment = bytes.readUInt16LE(p + 32);
    const offset = bytes.readUInt32LE(p + 42);
    if (flags & 1) throw new Error('Password-protected workbooks must be saved without a password for import.');
    total += unpacked;
    if (total > 200 * 1024 * 1024) throw new Error('Uncompressed workbook exceeds 200 MB. Remove unused sheets.');
    const name = bytes.subarray(p + 46, p + 46 + nameLen).toString('utf8');
    files.set(name, { method, size, offset });
    p += 46 + nameLen + extra + comment;
  }
  return (name) => {
    const e = files.get(name);
    if (!e) return null;
    const start = e.offset + 30 + bytes.readUInt16LE(e.offset + 26) + bytes.readUInt16LE(e.offset + 28);
    const compressed = bytes.subarray(start, start + e.size);
    if (e.method === 0) return compressed.toString('utf8');
    if (e.method !== 8) throw new Error('Unsupported workbook compression. Re-save as .xlsx.');
    return inflateRawSync(compressed).toString('utf8');
  };
}

function xml(text) {
  try { return xmlParser.parse(text); }
  catch { throw new Error('Invalid workbook XML.'); }
}
const arr = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);
const attr = (node, name) => node?.[`@_${name}`];
const text = (node) => (node == null ? '' : typeof node === 'object' ? String(node['#text'] ?? '') : String(node));

export function parseCSV(input) {
  const rows = [];
  let row = [], s = '', quoted = false;
  const sep = input.split('\n')[0].includes('\t') ? '\t' : ',';
  for (let i = 0; i < input.length; i++) {
    const c = input[i];
    if (c === '"') { if (quoted && input[i + 1] === '"') { s += '"'; i++; } else quoted = !quoted; }
    else if (c === sep && !quoted) { row.push(s); s = ''; }
    else if ((c === '\n' || c === '\r') && !quoted) { if (c === '\r' && input[i + 1] === '\n') i++; row.push(s); rows.push(row); row = []; s = ''; }
    else s += c;
  }
  if (s || row.length) { row.push(s); rows.push(row); }
  if (rows[0]?.[0]) rows[0][0] = rows[0][0].replace(/^\uFEFF/, '');
  return rows;
}

export function readWorkbookBuffer(buffer, fileName = 'workbook.xlsx') {
  if (buffer.length > 30 * 1024 * 1024) throw new Error('File exceeds 30 MB. Use a workbook containing the required tables only.');
  if (/\.(csv|tsv)$/i.test(fileName)) return [{ name: fileName, grid: parseCSV(buffer.toString('utf8')) }];
  const read = unzip(buffer);
  const wbText = read('xl/workbook.xml');
  if (!wbText) throw new Error('Excel workbook metadata is missing.');
  const wb = xml(wbText);
  const relsText = read('xl/_rels/workbook.xml.rels');
  const relmap = new Map(arr(xml(relsText || '<Relationships/>').Relationships?.Relationship).map((r) => [attr(r, 'Id'), attr(r, 'Target')]));
  const sharedText = read('xl/sharedStrings.xml');
  const strings = sharedText
    ? arr(xml(sharedText).sst?.si).map((si) => arr(si.t).map(text).join(text(si.r ? arr(si.r).map((r) => r.t) : '')) || arr(si.r).map((r) => text(r.t)).join(''))
    : [];
  const sheets = [];
  for (const sheet of arr(wb.workbook?.sheets?.sheet)) {
    const rid = attr(sheet, 'r:id') || attr(sheet, 'id');
    const target = relmap.get(rid);
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
    const sheetText = read(path);
    if (!sheetText) continue;
    const doc = xml(sheetText);
    const grid = [];
    for (const row of arr(doc.worksheet?.sheetData?.[0]?.row)) {
      const rowIndex = Number(attr(row, 'r')) - 1;
      if (rowIndex > 200000) throw new Error('Workbook has too many formatted rows. Remove unused rows and save again.');
      const line = [];
      for (const cell of arr(row.c)) {
        const i = col(attr(cell, 'r') || '');
        if (i < 0 || i > 500) continue;
        const type = attr(cell, 't');
        let val = cell.v != null ? text(cell.v) : '';
        if (type === 's') val = strings[Number(val)] ?? '';
        else if (type === 'inlineStr') val = arr(cell.is?.t).map(text).join('');
        else if (!type && val !== '') val = Number(val);
        else if (type === 'b') val = val === '1';
        line[i] = val;
      }
      grid[rowIndex] = line;
    }
    for (let i = 0; i < grid.length; i++) grid[i] ??= [];
    for (const m of arr(doc.worksheet?.mergeCells?.[0]?.mergeCell)) {
      const ref = attr(m, 'ref');
      const [a, b] = (ref || '').split(':');
      if (!b) continue;
      const r1 = +a.match(/\d+/)[0] - 1, r2 = +b.match(/\d+/)[0] - 1, c1 = col(a), c2 = col(b);
      if (c1 === c2 && r2 - r1 < 1000) for (let r = r1 + 1; r <= r2; r++) { grid[r] ??= []; grid[r][c1] = grid[r1]?.[c1] ?? ''; }
    }
    sheets.push({ name: attr(sheet, 'name'), grid });
  }
  return sheets;
}
