/*
 * Excel Recovery Add-in — cross-platform edition.
 *
 * Client-side remakes of the portable recovery strategies from the original
 * VB.NET Excel COM add-in (Code/Implementation.vb):
 *
 *   - "ZIP Repair — Try"        after Impl_ZipRepairTry (zip.exe -FF): scan the
 *     byte stream for PK\x03\x04 local headers, re-inflate every salvageable
 *     entry with the fault-tolerant Immortal Inflater, and repackage a fresh,
 *     well-formed .xlsx.
 *   - "Extract Data (non-MS)"   after Impl_NonMSXlsxExtractData /
 *     Impl_NonMSXlsxExtractData2 (doctotext / coffec): parse shared strings and
 *     xl/worksheets/*.xml directly — no Excel needed — and emit CSV per sheet.
 *   - "Salvage Cells"           in the spirit of Impl_ExcelFix-style extraction:
 *     regex-harvest cell references, values and text runs from partially
 *     readable XML, even when the sheet won't parse.
 *   - ".xls Text Rescue"        best-effort printable-string sweep (ASCII +
 *     UTF-16LE) of a legacy OLE2/BIFF workbook that nothing else will open.
 *
 * Everything runs locally in the browser (or Node for tests). No uploads.
 *
 * The COM-driven strategies of the original add-in (Open & Repair, Safe Mode,
 * SYLK, …) require a real Excel installation and remain in the Windows add-in.
 */

(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./immortal-inflate.js'));
  } else {
    root.ExcelAddinRecovery = factory({
      inflate: root.ImmortalInflate,
      unzipImmortal: root.unzipImmortal,
    });
  }
}(typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : this, function (Immortal) {
  'use strict';

  const NULL_LOG = { info() {}, ok() {}, warn() {}, err() {} };

  /* ---------------------------------------------------------------- utils */

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function decodeEntities(s) {
    return String(s)
      .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => safeCodePoint(parseInt(n, 16)))
      .replace(/&#(\d+);/g, (_, n) => safeCodePoint(parseInt(n, 10)))
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }

  function safeCodePoint(n) {
    try { return String.fromCodePoint(n); } catch (_) { return ''; }
  }

  function attr(attrs, name) {
    const m = attrs.match(new RegExp('(?:^|\\s)' + name + '\\s*=\\s*"([^"]*)"'));
    if (m) return m[1];
    const m2 = attrs.match(new RegExp("(?:^|\\s)" + name + "\\s*=\\s*'([^']*)'"));
    return m2 ? m2[1] : null;
  }

  function colToIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  function decodeText(bytes) {
    return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  }

  function detectFormat(bytes, fileName) {
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B &&
        (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)) return 'xlsx';
    if (bytes.length >= 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF &&
        bytes[2] === 0x11 && bytes[3] === 0xE0) return 'xls';
    const lower = String(fileName || '').toLowerCase();
    if (/\.(xlsx|xlsm|xltx|xltm)$/.test(lower)) return 'xlsx';
    if (/\.(xls|xlt|xlsb)$/.test(lower)) return 'xls';
    return 'unknown';
  }

  /* --------------------------------------------- strategy: ZIP Repair — Try
   * Honors Impl_ZipRepairTry (which shelled out to Info-ZIP `zip -FF`).
   * Scans local headers with the Immortal unzipper and rebuilds a clean zip. */

  function zipRepair(bytes, log) {
    log = log || NULL_LOG;
    const { files, stats } = Immortal.unzipImmortal(bytes, {
      onLog: (lvl, m) => (lvl === 'warn' ? log.warn(m) : log.info(m)),
    });
    log.info(`ZIP scan: ${stats.scanned} local header(s), ${stats.recovered} recovered ` +
             `(${stats.partial} partial), ${stats.skipped} skipped.`);
    return { files, stats };
  }

  /* Repair obviously broken XML text (strip invalid chars, escape bare '&',
   * truncate after last '>', close unclosed tags). */
  function repairXml(text) {
    let t = String(text)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g, '&amp;');
    const lastGt = t.lastIndexOf('>');
    if (lastGt >= 0 && lastGt < t.length - 1) t = t.slice(0, lastGt + 1);
    // Close unclosed tags (simple balance pass).
    const stack = [];
    const re = /<\s*(\/?)\s*([A-Za-z_][\w:.\-]*)\b[^>]*?(\/?)>/g;
    let m;
    while ((m = re.exec(t))) {
      if (m[3] === '/') continue;
      if (m[1] === '/') {
        for (let i = stack.length - 1; i >= 0; i--) {
          if (stack[i] === m[2]) { stack.splice(i, 1); break; }
        }
      } else if (!/^[?!]/.test(m[2])) {
        stack.push(m[2]);
      }
    }
    while (stack.length) t += `</${stack.pop()}>`;
    return t;
  }

  /* Rebuild a well-formed ZIP (xlsx) from { name -> Uint8Array }. */
  async function buildZip(entries) {
    const enc = new TextEncoder();
    const now = new Date();
    const time = ((now.getHours() & 0x1F) << 11) | ((now.getMinutes() & 0x3F) << 5) | ((now.getSeconds() >> 1) & 0x1F);
    const date = (((now.getFullYear() - 1980) & 0x7F) << 9) | (((now.getMonth() + 1) & 0xF) << 5) | (now.getDate() & 0x1F);

    async function deflateRaw(raw) {
      if (typeof CompressionStream === 'undefined') return null;
      try {
        const stream = new Response(new Blob([raw]).stream().pipeThrough(new CompressionStream('deflate-raw')));
        return new Uint8Array(await stream.arrayBuffer());
      } catch (_) { return null; }
    }

    const localChunks = [], cdChunks = [];
    let offset = 0;
    const names = Object.keys(entries).sort();
    for (const name of names) {
      const raw = entries[name];
      const nameBytes = enc.encode(name);
      const compressed = await deflateRaw(raw);
      const useStored = !compressed || compressed.length >= raw.length;
      const data = useStored ? raw : compressed;
      const method = useStored ? 0 : 8;
      const crc = crc32(raw);

      const lh = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, time, true);
      lv.setUint16(12, date, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, raw.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lh.set(nameBytes, 30);
      localChunks.push(lh, data);

      const ch = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(10, method, true);
      cv.setUint16(12, time, true);
      cv.setUint16(14, date, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, raw.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      cdChunks.push(ch);
      offset += lh.length + data.length;
    }

    const cdSize = cdChunks.reduce((s, c) => s + c.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, names.length, true);
    ev.setUint16(10, names.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);

    const all = [...localChunks, ...cdChunks, eocd];
    const total = all.reduce((s, c) => s + c.length, 0);
    const out = new Uint8Array(total);
    let p = 0;
    for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  /* Rebuild the xlsx container: heal XML entries, then repackage. */
  async function rebuildXlsx(files, log) {
    log = log || NULL_LOG;
    const enc = new TextEncoder();
    const healed = {};
    let healedCount = 0;
    for (const [name, data] of Object.entries(files)) {
      if (/\.(xml|rels)$/i.test(name)) {
        const text = decodeText(data);
        const fixed = repairXml(text);
        if (fixed !== text) healedCount++;
        healed[name] = enc.encode(fixed);
      } else {
        healed[name] = data;
      }
    }
    if (healedCount) log.warn(`Healed XML in ${healedCount} entr${healedCount === 1 ? 'y' : 'ies'}.`);
    const zipBytes = await buildZip(healed);
    log.ok(`Repackaged ${Object.keys(healed).length} entries into a fresh .xlsx (${zipBytes.length} bytes).`);
    return zipBytes;
  }

  /* ------------------------------------------ strategy: Extract Data (non-MS)
   * Honors Impl_NonMSXlsxExtractData / Impl_NonMSXlsxExtractData2, which used
   * doctotext / coffec to pull data out without Excel. Here we parse the OOXML
   * parts directly. */

  function parseSharedStrings(files) {
    const key = Object.keys(files).find(n => /(^|\/)sharedStrings\.xml$/i.test(n));
    if (!key) return [];
    const text = decodeText(files[key]);
    const out = [];
    const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = re.exec(text))) {
      const runs = [];
      const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let tm;
      while ((tm = tre.exec(m[1]))) runs.push(decodeEntities(tm[1]));
      out.push(runs.join(''));
    }
    return out;
  }

  /* Map worksheet paths to their workbook-declared names (best-effort). */
  function worksheetInventory(files) {
    const sheetPaths = Object.keys(files)
      .filter(n => /^xl\/worksheets\/[^/]+\.xml$/i.test(n))
      .sort((a, b) => {
        const an = parseInt((a.match(/(\d+)\.xml$/i) || [])[1] || '0', 10);
        const bn = parseInt((b.match(/(\d+)\.xml$/i) || [])[1] || '0', 10);
        return an - bn || a.localeCompare(b);
      });

    const names = {};   // path -> display name
    const wbKey = Object.keys(files).find(n => /(^|\/)workbook\.xml$/i.test(n));
    const relKey = Object.keys(files).find(n => /(^|\/)workbook\.xml\.rels$/i.test(n));
    if (wbKey && relKey) {
      try {
        const rels = {};
        const relText = decodeText(files[relKey]);
        const rre = /<Relationship\b([^>]*?)\/?>/g;
        let rm;
        while ((rm = rre.exec(relText))) {
          const id = attr(rm[1], 'Id');
          let target = attr(rm[1], 'Target');
          if (!id || !target) continue;
          target = target.replace(/^\.\//, '').replace(/^\//, '');
          if (!/^xl\//i.test(target)) target = 'xl/' + target;
          rels[id] = target;
        }
        const wbText = decodeText(files[wbKey]);
        const sre = /<sheet\b([^>]*?)\/?>/g;
        let sm;
        while ((sm = sre.exec(wbText))) {
          const name = attr(sm[1], 'name');
          const rid = attr(sm[1], 'r:id') || attr(sm[1], 'id');
          if (name && rid && rels[rid]) names[rels[rid]] = decodeEntities(name);
        }
      } catch (_) { /* fall back to file names */ }
    }

    return sheetPaths.map((path, i) => ({
      path,
      name: names[path] || path.replace(/^xl\/worksheets\//i, '').replace(/\.xml$/i, '') || `Sheet${i + 1}`,
    }));
  }

  /* Parse one worksheet XML into a dense 2-D array of strings. */
  function parseWorksheet(text, shared) {
    const cells = [];   // { r, c, v }
    let maxR = -1, maxC = -1;
    let autoRow = 0, autoCol = 0;

    // Track <row r="n"> so cells missing an r attribute still land somewhere sane.
    const chunks = text.split(/(?=<row\b)/);
    for (const chunk of chunks) {
      const rowAttr = chunk.match(/^<row\b([^>]*)>/);
      let rowIdx = null;
      if (rowAttr) {
        const rv = attr(rowAttr[1], 'r');
        rowIdx = rv ? parseInt(rv, 10) - 1 : autoRow;
        autoRow = (rowIdx == null ? autoRow : rowIdx) + 1;
        autoCol = 0;
      }
      const cre = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      while ((cm = cre.exec(chunk))) {
        const attrs = cm[1] || '';
        const body = cm[2] || '';
        const ref = attr(attrs, 'r');
        const type = attr(attrs, 't') || 'n';
        let r, c;
        if (ref && /^[A-Z]+\d+$/i.test(ref)) {
          const mm = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
          c = colToIndex(mm[1]);
          r = parseInt(mm[2], 10) - 1;
        } else {
          r = rowIdx != null ? rowIdx : autoRow;
          c = autoCol;
        }
        autoCol = c + 1;

        let value = null;
        if (type === 'inlineStr' || /<is\b/.test(body)) {
          const runs = [];
          const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
          let tm;
          while ((tm = tre.exec(body))) runs.push(decodeEntities(tm[1]));
          value = runs.join('');
        } else {
          const vm = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
          if (vm) {
            const raw = decodeEntities(vm[1]).trim();
            if (type === 's') {
              const idx = parseInt(raw, 10);
              value = (shared && Number.isInteger(idx) && shared[idx] !== undefined) ? shared[idx] : raw;
            } else if (type === 'b') {
              value = raw === '1' ? 'TRUE' : 'FALSE';
            } else {
              value = raw;
            }
          }
        }
        if (value === null || value === '') continue;
        cells.push({ r, c, v: value });
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }
    }

    const rows = [];
    for (let i = 0; i <= maxR; i++) rows.push(new Array(maxC + 1).fill(''));
    for (const cell of cells) {
      if (cell.r >= 0 && cell.c >= 0 && cell.r <= maxR && cell.c <= maxC) rows[cell.r][cell.c] = cell.v;
    }
    // Drop fully empty rows at the tail but keep interior blanks.
    while (rows.length && rows[rows.length - 1].every(v => v === '')) rows.pop();
    return rows;
  }

  function extractData(files, log) {
    log = log || NULL_LOG;
    const shared = parseSharedStrings(files);
    if (shared.length) log.info(`Parsed ${shared.length} shared string(s).`);
    const inventory = worksheetInventory(files);
    const sheets = [];
    for (const { path, name } of inventory) {
      try {
        const rows = parseWorksheet(decodeText(files[path]), shared);
        const cellCount = rows.reduce((s, r) => s + r.filter(v => v !== '').length, 0);
        sheets.push({ path, name, rows, cellCount });
        log.ok(`Sheet "${name}": ${rows.length} row(s), ${cellCount} cell(s).`);
      } catch (e) {
        log.warn(`Sheet "${name}" failed to parse: ${e.message || e}`);
      }
    }
    return { sheets, sharedCount: shared.length };
  }

  /* ------------------------------------------------ strategy: Salvage Cells
   * Regex-harvest whatever cell text survives in partially readable XML. */

  function salvageCells(input, log) {
    log = log || NULL_LOG;
    const texts = [];
    if (input && typeof input === 'object' && !(input instanceof Uint8Array)) {
      // { name -> Uint8Array } from a zip scan.
      for (const [name, data] of Object.entries(input)) {
        if (/\.(xml|rels)$/i.test(name) || /sharedStrings|worksheets/i.test(name)) {
          texts.push({ name, text: decodeText(data) });
        }
      }
    } else if (input instanceof Uint8Array) {
      texts.push({ name: '(raw bytes)', text: decodeText(input) });
    }

    const refCells = [];   // cells with a usable A1 reference
    const loose = [];      // values / text runs with no reference
    for (const { text } of texts) {
      const cre = /<c\b([^>]*?)>([\s\S]*?)<\/c>/g;
      let cm;
      while ((cm = cre.exec(text))) {
        const ref = attr(cm[1], 'r');
        const vm = cm[2].match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
        const tm = cm[2].match(/<t\b[^>]*>([\s\S]*?)<\/t>/);
        const val = decodeEntities((tm ? tm[1] : vm ? vm[1] : '')).trim();
        if (!val) continue;
        if (ref && /^[A-Z]+\d+$/i.test(ref)) {
          const mm = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
          refCells.push({ r: parseInt(mm[2], 10) - 1, c: colToIndex(mm[1]), v: val });
        } else {
          loose.push(val);
        }
      }
      // Text runs living outside any <c> (shared strings, damaged fragments).
      const stripped = text.replace(/<c\b[^>]*?>[\s\S]*?<\/c>/g, '');
      const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let tm2;
      while ((tm2 = tre.exec(stripped))) {
        const val = decodeEntities(tm2[1]).trim();
        if (val) loose.push(val);
      }
    }

    let rows = [];
    if (refCells.length) {
      const maxR = Math.max(...refCells.map(x => x.r));
      const maxC = Math.max(...refCells.map(x => x.c));
      for (let i = 0; i <= maxR; i++) rows.push(new Array(maxC + 1).fill(''));
      for (const cell of refCells) rows[cell.r][cell.c] = cell.v;
      rows = rows.filter(r => r.some(v => v !== ''));
    }
    log.info(`Salvage: ${refCells.length} referenced cell(s), ${loose.length} loose value(s)/text run(s).`);
    return { rows, cells: refCells, loose };
  }

  /* ------------------------------------------------ strategy: .xls Text Rescue
   * Printable-string sweep (ASCII + UTF-16LE) over a legacy OLE2/BIFF workbook.
   * Best-effort: recovers the human-readable content, not the grid. */

  function xlsTextRescue(bytes, log) {
    log = log || NULL_LOG;
    const isOle2 = bytes.length >= 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF &&
                   bytes[2] === 0x11 && bytes[3] === 0xE0;
    const strings = [];
    const seen = new Set();
    const MIN = 4;

    const push = (s) => {
      const t = s.trim();
      if (t.length < MIN) return;
      if (/^[\s\x20-\x2F\x3A-\x40\x5B-\x60\x7B-\x7E]+$/.test(t)) return; // punctuation only
      if (!seen.has(t)) { seen.add(t); strings.push(t); }
    };

    // ASCII / Latin-1 sweep.
    let run = '';
    for (let i = 0; i < bytes.length; i++) {
      const b = bytes[i];
      if ((b >= 0x20 && b <= 0x7E) || b === 0x09) run += String.fromCharCode(b);
      else { if (run.length >= MIN) push(run); run = ''; }
    }
    if (run.length >= MIN) push(run);

    // UTF-16LE sweep (both byte alignments).
    for (let start = 0; start < 2; start++) {
      let wrun = '';
      for (let i = start; i + 1 < bytes.length; i += 2) {
        const lo = bytes[i], hi = bytes[i + 1];
        if (hi === 0 && ((lo >= 0x20 && lo <= 0x7E) || lo === 0x09)) wrun += String.fromCharCode(lo);
        else { if (wrun.length >= MIN) push(wrun); wrun = ''; }
      }
      if (wrun.length >= MIN) push(wrun);
    }

    log.info(`${isOle2 ? 'OLE2 container detected. ' : ''}Rescued ${strings.length} distinct text string(s).`);
    return { isOle2, strings };
  }

  /* ------------------------------------------------------------------ CSV */

  function rowsToCsv(rows) {
    return rows.map(row => row.map(v => {
      const s = String(v == null ? '' : v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',')).join('\r\n') + '\r\n';
  }

  return {
    detectFormat,
    zipRepair,
    repairXml,
    rebuildXlsx,
    buildZip,
    extractData,
    parseSharedStrings,
    parseWorksheet,
    worksheetInventory,
    salvageCells,
    xlsTextRescue,
    rowsToCsv,
    crc32,
  };
}));

/* =============================== Browser UI =============================== */

(() => {
  'use strict';
  if (typeof document === 'undefined') return; // Node: engine-only

  const R = window.ExcelAddinRecovery;
  const $ = (id) => document.getElementById(id);
  const fmtBytes = (n) => n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(2)} MB`;

  const state = {
    file: null,
    bytes: null,
    format: 'unknown',
    files: null,       // recovered zip entries
    zipStats: null,
    sheets: null,
    repairedZip: null, // Uint8Array
  };

  /* ---------- log ---------- */
  const logEl = $('log');
  function push(cls, prefix, msg) {
    $('log-wrap').classList.remove('hidden');
    const line = document.createElement('div');
    line.className = cls;
    line.textContent = prefix + ' ' + msg;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }
  const log = {
    info: (m) => push('info', '·', m),
    ok:   (m) => push('ok', '✓', m),
    warn: (m) => push('warn', '!', m),
    err:  (m) => push('err', '✗', m),
    clear: () => { logEl.innerHTML = ''; $('log-wrap').classList.add('hidden'); },
  };

  function setStatus(msg, kind) {
    const s = $('status');
    s.classList.remove('hidden', 'error', 'success');
    if (kind) s.classList.add(kind);
    s.textContent = msg;
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    log.info(`Downloaded ${name}`);
  }

  const baseName = () => (state.file && state.file.name ? state.file.name : 'workbook').replace(/\.[^.]+$/, '');

  /* ---------- file intake ---------- */
  const drop = $('drop');
  const fileInput = $('file');
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('over');
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) handleFile(f);
  });

  async function handleFile(file) {
    log.clear();
    state.file = file;
    state.files = state.sheets = state.repairedZip = state.zipStats = null;
    state.salvage = state.xlsRescue = null;
    $('results').innerHTML = '';
    $('strategies').classList.add('hidden');
    setStatus(`Analyzing "${file.name}"…`);

    let bytes = new Uint8Array(await file.arrayBuffer());
    log.info(`Loaded "${file.name}" (${fmtBytes(bytes.length)}).`);

    // FIRST STEP: the shared S2 File Identifier — reads magic numbers,
    // separates concatenated/embedded foreign files (each downloadable in the
    // panel), and hands back the bytes recovery should operate on.
    if (window.S2FileID) {
      try {
        const report = S2FileID.analyze(bytes, { programKey: 'excelrcvryaddin', fileName: file.name });
        S2FileID.renderPanel($('fileid'), report);
        log.info(`Identified: ${report.primary.description}`);
        if (report.mismatch) log.warn('This tool recovers Excel workbooks — the file appears to be a different type. See the recommendation above.');
        if (report.foreign.length) log.warn(`${report.foreign.length} embedded/concatenated file(s) of a different type separated — download them above before recovering the rest.`);
        if (report.multiple) log.info(`Split into ${report.segments.length} segments; recovering segment ${report.proceedSegment.index + 1} (${report.proceedSegment.ext}, ${fmtBytes(report.proceedSegment.length)}).`);
        bytes = report.proceedBytes;
      } catch (e) {
        log.warn(`File identification failed (${e.message || e}) — continuing with raw bytes.`);
      }
    }

    state.bytes = bytes;
    state.format = R.detectFormat(bytes, file.name);
    log.info(`Detected format: ${state.format}`);
    setStatus(`Ready — pick a strategy below, or run them all. Format: ${state.format.toUpperCase()}.`);

    renderStrategies();
    $('strategies').classList.remove('hidden');
  }

  /* ---------- strategies ---------- */
  const STRATEGIES = [
    { id: 'auto', name: 'Auto — Run Everything', tag: 'recommended',
      desc: 'ZIP repair, then data extraction, then salvage — in the order the original add-in recommended.',
      applies: () => true },
    { id: 'zip', name: 'ZIP Repair — Try', legacy: 'after Impl_ZipRepairTry (zip -FF)',
      desc: 'Scan for PK local headers, re-inflate every salvageable entry, repackage a fresh .xlsx.',
      applies: (f) => f !== 'xls' },
    { id: 'extract', name: 'Extract Data (non-MS)', legacy: 'after Impl_NonMSXlsxExtractData / II',
      desc: 'Parse shared strings and worksheet XML directly — CSV per sheet, no Excel required.',
      applies: (f) => f !== 'xls' },
    { id: 'salvage', name: 'Salvage Cells', legacy: 'ExcelFix-style extraction',
      desc: 'Regex-harvest cell text from partially readable XML when the sheets won’t parse.',
      applies: (f) => f !== 'xls' },
    { id: 'xlstext', name: '.xls Text Rescue', legacy: 'legacy BIFF best-effort',
      desc: 'Printable-string sweep (ASCII + UTF-16) of an OLE2 workbook nothing else will open.',
      applies: () => true },
  ];

  function renderStrategies() {
    const wrap = $('methods');
    wrap.innerHTML = '';
    for (const s of STRATEGIES) {
      const btn = document.createElement('button');
      btn.className = 'method';
      btn.disabled = !s.applies(state.format);
      btn.innerHTML = `<h3>${s.name}${s.tag ? ` <span class="tag">${s.tag}</span>` : ''}</h3>` +
        (s.legacy ? `<p class="legacy">${s.legacy}</p>` : '') + `<p>${s.desc}</p>`;
      btn.addEventListener('click', () => runStrategy(s.id));
      wrap.appendChild(btn);
    }
  }

  async function runStrategy(id) {
    if (!state.bytes) return;
    document.querySelectorAll('.method').forEach(b => { b.dataset.wasDisabled = b.disabled ? '1' : ''; b.disabled = true; });
    try {
      switch (id) {
        case 'auto': await runAuto(); break;
        case 'zip': await runZip(); break;
        case 'extract': await runExtract(); break;
        case 'salvage': await runSalvage(); break;
        case 'xlstext': await runXlsText(); break;
      }
    } catch (e) {
      log.err(`${id} failed: ${e.message || e}`);
      setStatus(`Strategy failed: ${e.message || e}`, 'error');
    } finally {
      document.querySelectorAll('.method').forEach(b => { b.disabled = b.dataset.wasDisabled === '1'; });
    }
  }

  function ensureUnzipped() {
    if (!state.files) {
      const { files, stats } = R.zipRepair(state.bytes, log);
      state.files = files;
      state.zipStats = stats;
    }
    return state.files;
  }

  async function runZip() {
    log.info('▶ ZIP Repair — Try');
    const files = ensureUnzipped();
    const n = Object.keys(files).length;
    if (!n) { setStatus('ZIP repair could not recover any entries.', 'error'); return; }
    state.repairedZip = await R.rebuildXlsx(files, log);
    setStatus(`ZIP repair recovered ${n} entr${n === 1 ? 'y' : 'ies'} — repaired .xlsx ready below.`, 'success');
    renderResults();
  }

  async function runExtract() {
    log.info('▶ Extract Data (non-MS)');
    const files = ensureUnzipped();
    const { sheets } = R.extractData(files, log);
    state.sheets = sheets.filter(s => s.cellCount > 0);
    if (!state.sheets.length) {
      setStatus('No parsable worksheet data found — try Salvage Cells.', 'error');
      state.sheets = null;
      return;
    }
    const cells = state.sheets.reduce((s, x) => s + x.cellCount, 0);
    setStatus(`Extracted ${cells} cell(s) across ${state.sheets.length} sheet(s) — CSV downloads below.`, 'success');
    renderResults();
  }

  async function runSalvage() {
    log.info('▶ Salvage Cells');
    const files = ensureUnzipped();
    const source = Object.keys(files).length ? files : state.bytes;
    const res = R.salvageCells(source, log);
    if (!res.rows.length && !res.loose.length) {
      setStatus('Salvage found no readable cell text.', 'error');
      return;
    }
    state.salvage = res;
    setStatus(`Salvaged ${res.cells.length} referenced cell(s) and ${res.loose.length} loose value(s).`, 'success');
    renderResults();
  }

  async function runXlsText() {
    log.info('▶ .xls Text Rescue');
    const res = R.xlsTextRescue(state.bytes, log);
    if (!res.strings.length) { setStatus('No readable text found.', 'error'); return; }
    state.xlsRescue = res;
    setStatus(`Rescued ${res.strings.length} text string(s) — download below.`, 'success');
    renderResults();
  }

  async function runAuto() {
    log.info('▶ Auto — Run Everything');
    if (state.format === 'xls') {
      await runXlsText();
      return;
    }
    const files = ensureUnzipped();
    if (Object.keys(files).length) {
      state.repairedZip = await R.rebuildXlsx(files, log);
      const { sheets } = R.extractData(files, log);
      state.sheets = sheets.filter(s => s.cellCount > 0);
    }
    if (!state.sheets || !state.sheets.length) {
      log.warn('Structured extraction found nothing — falling back to Salvage Cells.');
      const res = R.salvageCells(Object.keys(files).length ? files : state.bytes, log);
      if (res.rows.length || res.loose.length) state.salvage = res;
    }
    if (!state.repairedZip && !state.sheets && !state.salvage) {
      log.warn('ZIP paths exhausted — falling back to text rescue.');
      const res = R.xlsTextRescue(state.bytes, log);
      if (res.strings.length) state.xlsRescue = res;
    }
    const got = [];
    if (state.repairedZip) got.push('repaired .xlsx');
    if (state.sheets && state.sheets.length) got.push(`${state.sheets.length} CSV sheet(s)`);
    if (state.salvage) got.push('salvaged cells');
    if (state.xlsRescue) got.push('text rescue');
    if (got.length) setStatus(`Auto finished: ${got.join(', ')} — downloads below.`, 'success');
    else setStatus('Auto could not recover anything from this file.', 'error');
    renderResults();
  }

  /* ---------- results ---------- */
  function renderResults() {
    const wrap = $('results');
    wrap.innerHTML = '';
    const mk = (label, primary, fn) => {
      const b = document.createElement('button');
      b.className = 'btn' + (primary ? '' : ' secondary');
      b.textContent = label;
      b.addEventListener('click', fn);
      wrap.appendChild(b);
    };
    if (state.repairedZip) {
      mk('⬇ Repaired .xlsx', true, () => downloadBlob(
        new Blob([state.repairedZip], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        baseName() + '.repaired.xlsx'));
    }
    if (state.sheets) {
      for (const sheet of state.sheets) {
        mk(`⬇ ${sheet.name}.csv`, false, () => downloadBlob(
          new Blob([R.rowsToCsv(sheet.rows)], { type: 'text/csv;charset=utf-8' }),
          `${baseName()}.${sheet.name.replace(/[^\w\- ]+/g, '_')}.csv`));
      }
    }
    if (state.salvage) {
      if (state.salvage.rows.length) {
        mk('⬇ Salvaged cells .csv', false, () => downloadBlob(
          new Blob([R.rowsToCsv(state.salvage.rows)], { type: 'text/csv;charset=utf-8' }),
          baseName() + '.salvaged.csv'));
      }
      if (state.salvage.loose.length) {
        mk('⬇ Loose text .txt', false, () => downloadBlob(
          new Blob([state.salvage.loose.join('\n')], { type: 'text/plain;charset=utf-8' }),
          baseName() + '.salvaged.txt'));
      }
    }
    if (state.xlsRescue) {
      mk('⬇ Rescued text .txt', false, () => downloadBlob(
        new Blob([state.xlsRescue.strings.join('\n')], { type: 'text/plain;charset=utf-8' }),
        baseName() + '.rescued.txt'));
    }
  }

  /* ---------- PWA ---------- */
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch(() => { /* app still works */ });
    });
  }
})();
