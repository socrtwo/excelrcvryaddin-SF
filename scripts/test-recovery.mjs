// Smoke-test the recovery engine (web/recovery.js) in Node against a tiny
// real .xlsx built with Python's zipfile module, plus deliberately corrupted
// variants (truncated tail / zeroed central directory).
//
// Run:  node scripts/test-recovery.mjs

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import url from 'node:url';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(url.fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const R = require(path.join(ROOT, 'web', 'recovery.js'));

let failures = 0;
const ok = (m) => console.log('  \x1b[32mOK\x1b[0m', m);
const bad = (m) => { console.log('  \x1b[31mFAIL\x1b[0m', m); failures++; };
const check = (cond, m) => (cond ? ok(m) : bad(m));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'excelrcvry-test-'));
const fixture = path.join(tmp, 'fixture.xlsx');

// ---------------------------------------------------------------- fixture
const PY = `
import sys, zipfile
out = sys.argv[1]
files = {
  '[Content_Types].xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    '</Types>',
  '_rels/.rels':
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml':
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    '<sheet name="Budget" sheetId="1" r:id="rId1"/>'
    '<sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '</Relationships>',
  'xl/sharedStrings.xml':
    '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">'
    '<si><t>Hello Recovery</t></si><si><t>Second sheet cell</t></si></sst>',
  'xl/worksheets/sheet1.xml':
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>'
    '<row r="2"><c r="A2" t="inlineStr"><is><t>World</t></is></c><c r="B2"><v>3.14</v></c></row>'
    '</sheetData></worksheet>',
  'xl/worksheets/sheet2.xml':
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    '<row r="1"><c r="A1" t="s"><v>1</v></c></row>'
    '</sheetData></worksheet>',
}
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in files.items():
        z.writestr(name, data)
print('wrote', out)
`;
execFileSync('python3', ['-c', PY, fixture], { stdio: 'inherit' });
const good = new Uint8Array(fs.readFileSync(fixture));
console.log('Fixture size:', good.length, 'bytes');

function expectData(files, label, { minEntries }) {
  const n = Object.keys(files).length;
  check(n >= minEntries, `${label}: recovered ${n} entries (>= ${minEntries})`);
  const { sheets } = R.extractData(files);
  const budget = sheets.find(s => s.name === 'Budget');
  const notes = sheets.find(s => s.name === 'Notes');
  check(!!budget, `${label}: sheet "Budget" resolved by name`);
  if (budget) {
    check(budget.rows[0] && budget.rows[0][0] === 'Hello Recovery', `${label}: A1 shared string = "Hello Recovery"`);
    check(budget.rows[0] && budget.rows[0][1] === '42', `${label}: B1 = 42`);
    check(budget.rows[1] && budget.rows[1][0] === 'World', `${label}: A2 inline string = "World"`);
    check(budget.rows[1] && budget.rows[1][1] === '3.14', `${label}: B2 = 3.14`);
  }
  check(!!notes && notes.rows[0] && notes.rows[0][0] === 'Second sheet cell', `${label}: sheet 2 cell recovered`);
  return sheets;
}

// [1] intact file --------------------------------------------------------
console.log('\n[1] intact .xlsx');
check(R.detectFormat(good, 'fixture.xlsx') === 'xlsx', 'format detected as xlsx');
const r1 = R.zipRepair(good);
expectData(r1.files, 'intact', { minEntries: 7 });

// CSV output sanity
{
  const { sheets } = R.extractData(r1.files);
  const csv = R.rowsToCsv(sheets[0].rows);
  check(csv.includes('Hello Recovery,42'), 'CSV row 1 rendered');
}

// Rebuild round-trip: python zipfile must read the repaired container.
{
  const rebuilt = await R.rebuildXlsx(r1.files);
  const outPath = path.join(tmp, 'rebuilt.xlsx');
  fs.writeFileSync(outPath, rebuilt);
  const PY_CHECK = `
import sys, zipfile
z = zipfile.ZipFile(sys.argv[1])
assert z.testzip() is None, 'CRC check failed'
names = z.namelist()
assert 'xl/worksheets/sheet1.xml' in names, names
body = z.read('xl/worksheets/sheet1.xml').decode()
assert '<v>42</v>' in body, body
print('entries:', len(names))
`;
  const out = execFileSync('python3', ['-c', PY_CHECK, outPath]).toString().trim();
  check(/entries: 7/.test(out), `rebuilt .xlsx verifies with python zipfile (${out})`);
}

// [2] truncated tail (central directory destroyed) -----------------------
console.log('\n[2] truncated archive (central directory cut off)');
const cut = good.slice(0, good.length - 200);
const r2 = R.zipRepair(cut);
expectData(r2.files, 'truncated', { minEntries: 6 });

// [3] zeroed central directory -------------------------------------------
console.log('\n[3] zeroed central directory');
const mangled = good.slice();
{
  // Find EOCD, read CD offset, zero the whole central directory + EOCD.
  let eocd = -1;
  for (let i = mangled.length - 22; i >= 0; i--) {
    if (mangled[i] === 0x50 && mangled[i+1] === 0x4B && mangled[i+2] === 0x05 && mangled[i+3] === 0x06) { eocd = i; break; }
  }
  check(eocd > 0, 'found EOCD to corrupt');
  const cdOff = mangled[eocd+16] | (mangled[eocd+17] << 8) | (mangled[eocd+18] << 16) | (mangled[eocd+19] << 24);
  mangled.fill(0, cdOff);
}
const r3 = R.zipRepair(mangled);
expectData(r3.files, 'zeroed-CD', { minEntries: 7 });

// [4] Salvage Cells on broken worksheet XML ------------------------------
console.log('\n[4] Salvage Cells on partially readable XML');
{
  const brokenXml =
    '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>salvage me</t></is></c>' +
    '<c r="B1"><v>7</v></c></row><row r="2"><c r="A2"><v>99</v></c><badly <<< truncated garbage';
  const res = R.salvageCells(new TextEncoder().encode(brokenXml));
  check(res.cells.length >= 3, `harvested ${res.cells.length} referenced cells from broken XML`);
  const flat = res.rows.flat();
  check(flat.includes('salvage me') && flat.includes('7') && flat.includes('99'), 'salvaged values present in rows');
}

// [5] .xls text rescue ----------------------------------------------------
console.log('\n[5] .xls text rescue (printable-string sweep)');
{
  // Synthetic OLE2-ish bytes: magic + noise + ASCII + UTF-16LE strings.
  const parts = [];
  parts.push(new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]));
  parts.push(new Uint8Array(64).fill(0x01));
  parts.push(new TextEncoder().encode('Quarterly totals'));
  parts.push(new Uint8Array(16));
  const wide = 'Wide char cell';
  const wideBytes = new Uint8Array(wide.length * 2);
  for (let i = 0; i < wide.length; i++) { wideBytes[i * 2] = wide.charCodeAt(i); wideBytes[i * 2 + 1] = 0; }
  parts.push(wideBytes);
  parts.push(new Uint8Array(16).fill(0xFF));
  const total = parts.reduce((s, p) => s + p.length, 0);
  const fakeXls = new Uint8Array(total);
  let p = 0;
  for (const part of parts) { fakeXls.set(part, p); p += part.length; }

  check(R.detectFormat(fakeXls, 'fake.xls') === 'xls', 'format detected as xls (OLE2 magic)');
  const res = R.xlsTextRescue(fakeXls);
  check(res.isOle2, 'OLE2 container flagged');
  check(res.strings.includes('Quarterly totals'), 'ASCII string rescued');
  check(res.strings.includes('Wide char cell'), 'UTF-16LE string rescued');
}

// ------------------------------------------------------------------- done
fs.rmSync(tmp, { recursive: true, force: true });
console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll checks passed.');
process.exit(failures ? 1 : 0);
