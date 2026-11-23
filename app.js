/* =========================================================================
 * POH AED Duty Calendar Converter — app.js
 *
 * 100% client-side. Nothing is uploaded anywhere.
 *
 * The file is a tiny UMD-style module:
 *   - `POH.*` contains all pure logic (parsing, timing, ICS, Excel workbook).
 *     It has no DOM dependency and can be require()'d in Node for tests.
 *   - The UI wiring at the bottom only runs when a real `document` exists.
 *
 * Dependencies (loaded from CDN by index.html):
 *   - SheetJS  `XLSX`   : reads the source .xlsx
 *   - ExcelJS  `ExcelJS`: writes the styled .xlsx
 * ========================================================================= */

(function (root, factory) {
  'use strict';
  var api = factory(root);
  root.POH = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis, function (root) {
  'use strict';

  /* ------------------------------------------------------------------ *
   * 1. Constants
   * ------------------------------------------------------------------ */

  var MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
                     'July', 'August', 'September', 'October', 'November', 'December'];
  var MONTH_ABBR  = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                     'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  // Sunday-first week (as requested). Used by both the web grid and the Excel export.
  var DAY_NAMES   = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var DAY_ABBR    = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  /** Pastel palette. The exact same hex values are used by the web UI and ExcelJS. */
  var PALETTE = {
    A:     { bg: '#DCE6F1', fg: '#1E3A5F', label: 'A duties' },
    P:     { bg: '#D9EAD3', fg: '#1E4620', label: 'P duties' },
    O:     { bg: '#E4DFEC', fg: '#3B235A', label: 'O duties' },
    N:     { bg: '#F4CCCC', fg: '#661919', label: 'N' },
    VL:    { bg: '#FCE5CD', fg: '#663C00', label: 'VL' },
    OTHER: { bg: '#F4CCB8', fg: '#662900', label: 'Resp / Other' },
    OFF:   { bg: '#F3F4F6', fg: '#6B7280', label: 'Off' }
  };
  var CATEGORY_ORDER = ['A', 'P', 'O', 'N', 'VL', 'OTHER', 'OFF'];

  /** Short keys used for the CSS custom properties that drive the whole UI. */
  var CSS_KEY = { A: 'a', P: 'p', O: 'o', N: 'n', VL: 'v', OTHER: 'r', OFF: 'off' };

  /** Consultants are deliberately hidden from the staff picker. */
  var EXCLUDED_STAFF = ['wong cy', 'wong ht'];

  /** Keys for the two localStorage records (roster + user colours). */
  var LS_KEY = 'poh-duty-roster-v1';
  var COLOR_KEY = 'poh-duty-colors-v1';

  /** Fallback CDN mirrors, used only if a script tag in index.html failed. */
  var CDN = {
    xlsx: [
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
    ],
    exceljs: [
      'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
      'https://unpkg.com/exceljs@4.4.0/dist/exceljs.min.js'
    ]
  };

  /* ------------------------------------------------------------------ *
   * 2. Small helpers
   * ------------------------------------------------------------------ */

  function getXLSX() {
    var X = root.XLSX;
    if (!X) throw new Error('SheetJS (XLSX) is not loaded.');
    return X;
  }
  function pad2(n)  { return String(n).padStart(2, '0'); }
  function pad4(n)  { return String(n).padStart(4, '0'); }
  function norm(s)  { return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase(); }
  function esc(s)   { return String(s).replace(/\\/g, '\\\\').replace(/;/g, '\\;')
                                      .replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }
  function html(s)  { return String(s == null ? '' : s)
                        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
  function argb(hex) { return 'FF' + String(hex).replace('#', '').toUpperCase(); }
  function daysInMonth(year, month) { return new Date(Date.UTC(year, month, 0)).getUTCDate(); }
  /** Sunday-first leading blank cells for the 1st of the month (Sun = 0). */
  function leadingBlanks(year, month) {
    return new Date(Date.UTC(year, month - 1, 1)).getUTCDay();
  }
  function weekdayIndex(year, month, day) {
    return new Date(Date.UTC(year, month - 1, day)).getUTCDay();
  }
  function fmt(minutes) {
    return pad2(Math.floor(minutes / 60)) + ':' + pad2(minutes % 60);
  }
  function fmtEnd(minutes) {
    if (minutes === 1440) return '24:00';            // exactly midnight, keep "24:00"
    if (minutes > 1440) return fmt(minutes - 1440) + ' (+1)';
    return fmt(minutes);
  }

  /* -------------------- user-adjustable colour helpers ---------------- */

  function validHex(s) { return /^#[0-9a-f]{6}$/i.test(String(s || '')); }

  function clonePalette(src) {
    var out = {};
    CATEGORY_ORDER.forEach(function (cat) {
      var p = (src && src[cat]) || PALETTE[cat];
      out[cat] = { bg: validHex(p.bg) ? p.bg : PALETTE[cat].bg,
                   fg: validHex(p.fg) ? p.fg : PALETTE[cat].fg,
                   label: p.label || PALETTE[cat].label };
    });
    return out;
  }

  function hexToRgb(hex) {
    var h = String(hex).replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16) };
  }
  function rgbToHex(r, g, b) {
    function c(n) { return Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0'); }
    return '#' + c(r) + c(g) + c(b);
  }
  function luminance(hex) {
    var v = hexToRgb(hex);
    function lin(n) { n /= 255; return n <= 0.03928 ? n / 12.92 : Math.pow((n + 0.055) / 1.055, 2.4); }
    return 0.2126 * lin(v.r) + 0.7152 * lin(v.g) + 0.0722 * lin(v.b);
  }
  /**
   * Auto tone: pick a readable text colour for any background the user chooses.
   * Light backgrounds get a darkened tint of the same hue; dark ones get near-white.
   */
  function autoFg(bg) {
    if (!validHex(bg)) return '#1F2937';
    var v = hexToRgb(bg), L = luminance(bg);
    var ratio = function (fg) {
      var a = luminance(fg), b = L;
      return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
    };
    // Near-white works on dark backgrounds, but check contrast (pure red fails it).
    if (L < 0.28 && ratio('#F9FAFB') >= 4.5) return '#F9FAFB';
    // Otherwise darken the chosen background, preserving its hue, until contrast >= 4.5:1.
    for (var k = 0.55; k >= 0.10; k -= 0.05) {
      var tint = rgbToHex(v.r * k, v.g * k, v.b * k);
      if (ratio(tint) >= 4.5) return tint;
    }
    return '#1F2937';
  }

  function codeKey(code) {
    return String(code || '').replace(/\s*\([^)]*\)\s*/g, '').replace(/\s+/g, '').toUpperCase();
  }
  /** Resolve the effective colours for a duty: per-code override wins over category. */
  function colorsFor(code, palette, codeColors) {
    var cat = categoryOf(code);
    var base = (palette && palette[cat]) || PALETTE[cat];
    var ov = codeColors && codeColors[codeKey(code)];
    if (ov && validHex(ov.bg)) return { bg: ov.bg, fg: validHex(ov.fg) ? ov.fg : autoFg(ov.bg), label: base.label };
    return base;
  }
  /** Paired duties (e.g. A5 + FUC) are one duty: show "A5/FUC" in a single cell. */
  function joinLabels(duties) {
    return duties.map(function (d) { return d.label; }).join('/');
  }
  /** Every distinct duty code appearing in a roster, with a usage count. */
  function allDutyCodes(roster) {
    var counts = {};
    if (roster && roster.byStaff) {
      Object.keys(roster.byStaff).forEach(function (id) {
        var bucket = roster.byStaff[id];
        Object.keys(bucket).forEach(function (day) {
          bucket[day].forEach(function (d) {
            var k = codeKey(d.label);
            if (!counts[k]) counts[k] = { key: k, label: d.label, code: d.code, count: 0 };
            counts[k].count++;
          });
        });
      });
    }
    return Object.keys(counts).map(function (k) { return counts[k]; })
      .sort(function (a, b) {
        var ca = CATEGORY_ORDER.indexOf(categoryOf(a.code)), cb = CATEGORY_ORDER.indexOf(categoryOf(b.code));
        return ca - cb || a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: 'base' });
      });
  }

  /** Raw cell text addressed by 1-based row + 0-based column (A = 0). */
  function readCell(ws, row, col) {
    var XLSX = getXLSX();
    if (row < 1 || col < 0) return '';
    var cell = ws[XLSX.utils.encode_cell({ r: row - 1, c: col })];
    if (!cell) return '';
    var v = cell.v;
    if (v === undefined || v === null) v = (cell.w !== undefined ? cell.w : '');
    if (typeof v === 'number') return String(v);
    return String(v).trim();                          // kills "Wong CY " trailing space / " " filler
  }
  function usedRange(ws) {
    var XLSX = getXLSX();
    if (!ws['!ref']) return { maxRow: 0, maxCol: 0 };
    var r = XLSX.utils.decode_range(ws['!ref']);
    return { maxRow: r.e.r + 1, maxCol: r.e.c + 1 };
  }
  function colLetter(c) { return getXLSX().utils.encode_col(c); }

  /* ------------------------------------------------------------------ *
   * 3. Month / year and layout detection
   * ------------------------------------------------------------------ */

  /** 12_2026 -> December 2026; falls back to U1 text such as "Dec 2026". */
  function detectMonthYear(sheetName, ws) {
    var m = /^(\d{1,2})[_\-\s](\d{4})$/.exec(String(sheetName || '').trim());
    if (m) {
      var mo = parseInt(m[1], 10), yr = parseInt(m[2], 10);
      if (mo >= 1 && mo <= 12 && yr > 1900) return { month: mo, year: yr, source: 'sheet name' };
    }
    // Fallback: scan the first three rows for "Dec 2026" / "December 2026".
    var rng = usedRange(ws);
    for (var r = 1; r <= Math.min(3, rng.maxRow); r++) {
      for (var c = 0; c < Math.min(60, rng.maxCol); c++) {
        var t = readCell(ws, r, c);
        var mm = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{4})\b/i.exec(t);
        if (mm) {
          var idx = MONTH_ABBR.indexOf(mm[1].toLowerCase().slice(0, 3));
          if (idx >= 0) return { month: idx + 1, year: parseInt(mm[2], 10), source: 'cell ' + colLetter(c) + r };
        }
      }
    }
    var now = new Date();
    return { month: now.getMonth() + 1, year: now.getFullYear(), guessed: true, source: 'system date' };
  }

  /** Find an exact label ("Senior" / "MO") within a small band of rows. */
  function findLabel(ws, label, rows, maxCol) {
    var re = new RegExp('^' + label + '$', 'i');
    for (var i = 0; i < rows.length; i++) {
      for (var c = 0; c < maxCol; c++) {
        if (re.test(readCell(ws, rows[i], c))) return { row: rows[i], col: c };
      }
    }
    return null;
  }

  /** Columns holding a 1,2,3 day run followed by a 3-letter weekday column. */
  function findDayColumns(ws, maxRow, maxCol) {
    var out = [];
    for (var c = 0; c < Math.min(maxCol, 60); c++) {
      if (readCell(ws, 4, c) === '1' && readCell(ws, 5, c) === '2' && readCell(ws, 6, c) === '3' &&
          /^(mon|tue|wed|thu|fri|sat|sun)/i.test(readCell(ws, 4, c + 1))) {
        out.push(c);
      }
    }
    return out;
  }

  /** Collect shift headers on row 3. Skips blanks, "Remarks" and prose. */
  function collectHeaders(ws, startCol, endCol, headerRow, block) {
    var list = [];
    for (var c = startCol; c <= endCol; c++) {
      var h = readCell(ws, headerRow, c);
      if (!h) continue;
      if (/^remarks?$/i.test(h)) continue;             // MO "Remarks" notes column
      if (!/^[A-Za-z0-9]{1,8}$/.test(h)) continue;    // only short duty codes
      list.push({ col: c, code: h, block: block });
    }
    return list;
  }

  /**
   * Rows 36-38 hold sentences such as "A1# denotes duty from 8A-6P".
   * Build a map "A1#" -> that sentence so modifiers can carry their note.
   */
  function collectTimingNotes(ws, maxRow) {
    var notes = {};
    for (var r = 35; r <= Math.min(39, maxRow); r++) {
      for (var c = 0; c < 50; c++) {
        var t = readCell(ws, r, c);
        if (!t || !/[*^#+]/.test(t) || !/denotes/i.test(t)) continue;
        var toks = t.match(/[A-Za-z][A-Za-z0-9]{0,5}[*^#+]/g) || [];
        var desc = t.replace(/\s+/g, ' ').trim();
        for (var i = 0; i < toks.length; i++) {
          var key = toks[i].replace(/^\*/, '').toUpperCase();
          if (!notes[key]) notes[key] = desc;
        }
      }
    }
    return notes;
  }

  /**
   * Staff directory: rows 40-55, three two-column groups
   * (Senior A/B, Junior G/H, TMH-TSWH J/K). Labels are discovered in row 40
   * rather than hard-coded, so other months with the same layout still work.
   */
  function buildDirectory(ws, maxRow) {
    var groups = [], seen = {};
    for (var r = 38; r <= 42; r++) {
      for (var c = 0; c < 14; c++) {
        var t = readCell(ws, r, c);
        if (/^senior$/i.test(t) || /^junior$/i.test(t) || /^tmh[\/ ]?tswh$/i.test(t)) {
          if (!seen[c]) { seen[c] = true; groups.push({ col: c, row: r, group: t.trim() }); }
        }
      }
    }
    if (!groups.length) {
      groups = [{ col: 0, row: 40, group: 'Senior' },
                { col: 6, row: 40, group: 'Junior' },
                { col: 9, row: 40, group: 'TMH/TSWH' }];
    }
    var startRow = 41;
    for (var g = 0; g < groups.length; g++) startRow = Math.max(startRow, groups[g].row + 1);

    var dir = new Map();
    for (var gi = 0; gi < groups.length; gi++) {
      var grp = groups[gi];
      for (var rr = startRow; rr <= Math.min(maxRow, 60); rr++) {
        var key = readCell(ws, rr, grp.col);
        var name = readCell(ws, rr, grp.col + 1);
        if (!key || !name) continue;
        if (key.length > 16 || /\d{2}/.test(key)) continue;      // skip prose / counters
        if (!/[A-Za-z]/.test(key)) continue;
        if (/denotes|duty|list|activities/i.test(key)) continue;
        var nk = norm(key);
        if (dir.has(nk)) {
          var existing = dir.get(nk);
          if (existing.name !== name && existing.alt.indexOf(name) < 0) existing.alt.push(name);
        } else {
          dir.set(nk, { id: nk, key: key, name: name, group: grp.group, alt: [], discovered: false });
        }
      }
    }
    return dir;
  }

  /** Split one roster cell into staff tokens, preserving suffixes and (am)/(pm). */
  function splitTokens(text) {
    return String(text).split(/[,;]/).map(function (p) { return p.trim(); }).filter(Boolean);
  }
  function parseToken(raw) {
    var tok = String(raw).trim();
    var modMatch = tok.match(/([*^#+])$/);
    var modifier = modMatch ? modMatch[1] : '';
    var base = modMatch ? tok.slice(0, -1) : tok;
    var paren = /\((am|pm)\)/i.exec(base);
    var extra = paren ? paren[0] : '';
    if (paren) base = base.replace(/\s*\((am|pm)\)\s*/gi, ' ').trim();
    base = base.replace(/^[\s,;]+|[\s,;]+$/g, '').trim();
    return { base: base, modifier: modifier, extra: extra };
  }
  function isStaffToken(base) {
    if (!base || base.length > 20 || !/[A-Za-z]/.test(base)) return false;
    if (/^(off|nil|n\/a|nil|tbc|-|--)$/i.test(base)) return false;
    return true;
  }

  /* ------------------------------------------------------------------ *
   * 4. Workbook parser — the heart of the converter
   * ------------------------------------------------------------------ */

  /**
   * Parse a POH AED One-Page Duty .xlsx.
   * @param {ArrayBuffer|Uint8Array} data  file bytes
   * @param {string} fileName              used for diagnostics / filenames
   * @returns {object} roster (JSON-safe)
   */
  function parseWorkbook(data, fileName) {
    var XLSX = getXLSX();
    var wb = XLSX.read(data, { type: 'array', cellDates: false, cellFormula: false });
    if (!wb.SheetNames.length) throw new Error('The workbook has no worksheets.');

    // Prefer a sheet named MM_YYYY, otherwise use the first sheet.
    var sheetName = wb.SheetNames.filter(function (n) { return /^\d{1,2}[_\-\s]\d{4}$/.test(n.trim()); })[0]
                 || wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    var rng = usedRange(ws);
    var my = detectMonthYear(sheetName, ws);
    var days = daysInMonth(my.year, my.month);
    var warnings = [];

    if (my.guessed) warnings.push('Month/year not found in the sheet name or U1 — used the device date instead.');

    // --- Locate the two side-by-side blocks -----------------------------
    var seniorLabel = findLabel(ws, 'Senior', [2, 3], Math.min(rng.maxCol, 60));
    var moLabel     = findLabel(ws, 'MO',     [2, 3], Math.min(rng.maxCol, 60));
    var dayCols     = findDayColumns(ws, rng.maxRow, rng.maxCol);
    var seniorDayCol = dayCols.length ? dayCols[0] : 0;                       // A
    var moDayCol     = dayCols.length > 1 ? dayCols[1]
                     : (moLabel ? moLabel.col - 2 : seniorDayCol + 20);       // U

    var seniorStart = seniorLabel ? seniorLabel.col : seniorDayCol + 2;       // C
    var moStart     = moLabel ? moLabel.col : moDayCol + 2;                   // W
    var seniorHeaders = collectHeaders(ws, seniorStart, moDayCol - 1, 3, 'Senior');
    var moHeaders     = collectHeaders(ws, moStart, Math.min(rng.maxCol, 60) - 1, 3, 'MO');
    var headers = seniorHeaders.concat(moHeaders);

    if (!headers.length) throw new Error('No shift headers found on row 3. Is this the POH one-page duty format?');
    if (!seniorHeaders.length) warnings.push('Senior block not detected — check the row 2 "Senior" label.');
    if (!moHeaders.length)     warnings.push('MO block not detected — check the row 2 "MO" label.');

    // --- Row 4..34: day numbers ----------------------------------------
    var dayRows = [];
    for (var r = 4; r <= Math.min(rng.maxRow, 45); r++) {
      var raw = readCell(ws, r, seniorDayCol);
      var d = Number(raw);
      if (!Number.isInteger(d) || d < 1 || d > 31) {
        if (dayRows.length) break;
        continue;
      }
      dayRows.push({ day: d, row: r });
    }
    if (!dayRows.length) throw new Error('No day numbers (1-31) found in column ' + colLetter(seniorDayCol) + '.');
    if (dayRows.length < days) warnings.push('Only ' + dayRows.length + ' day rows found; ' + MONTH_NAMES[my.month - 1] + ' has ' + days + '.');

    var dir = buildDirectory(ws, rng.maxRow);
    var timingNotes = collectTimingNotes(ws, rng.maxRow);
    var byStaff = {};
    var discovered = [];

    // --- Walk every grid cell -------------------------------------------
    for (var i = 0; i < dayRows.length; i++) {
      var day = dayRows[i].day;
      if (day > days) continue;                       // e.g. 31 in a 30-day month
      var row = dayRows[i].row;
      for (var h = 0; h < headers.length; h++) {
        var hd = headers[h];
        var text = readCell(ws, row, hd.col);
        if (!text) continue;
        var toks = splitTokens(text);
        for (var t = 0; t < toks.length; t++) {
          var p = parseToken(toks[t]);
          if (!isStaffToken(p.base)) continue;
          var id = norm(p.base);
          if (EXCLUDED_STAFF.indexOf(id) >= 0) continue;      // consultants: hide everywhere
          var staff = dir.get(id);
          if (!staff) {
            staff = { id: id, key: p.base, name: p.base, group: 'Discovered',
                      alt: [], discovered: true };
            dir.set(id, staff);
            discovered.push(p.base);
          }
          var key = (hd.code + p.modifier).toUpperCase();
          var label = hd.code + p.modifier + (p.extra ? ' ' + p.extra : '');
          var duty = {
            code: hd.code,
            modifier: p.modifier,
            extra: p.extra,
            label: label,
            note: p.modifier ? (timingNotes[key] || '') : '',
            timing: getTiming(hd.code + p.modifier),
            block: hd.block,
            column: colLetter(hd.col),
            source: text
          };
          var bucket = byStaff[id] || (byStaff[id] = {});
          var list = bucket[day] || (bucket[day] = []);
          // A person can appear in several columns on the same day; keep unique labels.
          if (!list.some(function (x) { return x.label === label; })) list.push(duty);
        }
      }
    }

    // --- Staff list: only people with duties this month ------------------
    var staffList = [];
    dir.forEach(function (s) {
      if (EXCLUDED_STAFF.indexOf(s.id) >= 0) return;          // Wong CY / Wong HT
      var b = byStaff[s.id];
      var count = 0;
      if (b) Object.keys(b).forEach(function (d) { count += b[d].length; });
      if (!count) return;
      var name = s.name + (s.alt.length ? ' / ' + s.alt.join(' / ') : '');
      s.label = s.discovered ? (s.key + ' (not in legend)')
              : (name && name !== s.key ? s.key + ' — ' + name : s.key);
      s.dutyCount = count;
      staffList.push(s);
    });
    staffList.sort(function (a, b) { return a.key.localeCompare(b.key, undefined, { sensitivity: 'base' }); });

    if (discovered.length) {
      warnings.push(discovered.length + ' code(s) appeared in the grid but not in the legend and were auto-added: '
                    + Array.from(new Set(discovered)).slice(0, 12).join(', ')
                    + (discovered.length > 12 ? '…' : ''));
    }

    return {
      version: 1,
      fileName: fileName || 'duty.xlsx',
      sheetName: sheetName,
      month: my.month,
      year: my.year,
      monthName: MONTH_NAMES[my.month - 1],
      monthSource: my.source,
      daysInMonth: days,
      leadingBlanks: leadingBlanks(my.year, my.month),
      seniorHeaders: seniorHeaders.map(function (h) { return h.code; }),
      moHeaders: moHeaders.map(function (h) { return h.code; }),
      dayRowCount: dayRows.length,
      seniorDayCol: colLetter(seniorDayCol),
      moDayCol: colLetter(moDayCol),
      timingNotes: timingNotes,
      staff: staffList,
      byStaff: byStaff,
      warnings: warnings
    };
  }

  /* ------------------------------------------------------------------ *
   * 5. Duty classification + timings
   * ------------------------------------------------------------------ */

  function categoryOf(code) {
    var c = String(code || '').replace(/[*^#+].*$/, '').trim().toUpperCase();
    if (/^A\d?$/.test(c)) return 'A';
    if (/^P\d?$/.test(c)) return 'P';
    if (/^O\d?$/.test(c)) return 'O';
    if (c === 'N' || c === 'TSWN') return 'N';
    if (c === 'VL' || c === 'CCL') return 'VL';
    return 'OTHER';                                  // Resp / M / C / FUC / TSWM / unknown
  }

  /**
   * Default AED shift times (Asia/Hong_Kong) with the row 36-38 modifiers.
   * Returns minutes-from-midnight + a duration, so overnight shifts roll over.
   */
  function getTiming(label) {
    var clean = String(label || '').replace(/\s*\([^)]*\)\s*/g, '').trim();   // strip "(am)"
    var m = /^([A-Za-z]+)(\d+)?([*^#+])?$/.exec(clean);
    if (!m) return timed(9 * 60, 8 * 60, 'Unmapped code');
    var family = m[1].toUpperCase();
    var num = m[2] ? parseInt(m[2], 10) : null;
    var mod = m[3] || '';

    if (family === 'A') {
      if (mod === '#') return timed(8 * 60, 10 * 60);                 // 08:00-18:00
      if (num === 3 && mod === '*') return timed(8 * 60, 6 * 60);     // off at 14:00
      return timed(8 * 60, 9 * 60);                                   // 08:00-17:00 (incl ^, +)
    }
    if (family === 'P') {
      if (mod === '^' && num === 1) return timed(14 * 60, 9 * 60);    // 14:00-23:00
      if (mod === '^' && num === 3) return timed(15 * 60, 9 * 60);    // 15:00-24:00
      if (mod === '*') return timed(14 * 60, 10 * 60);                // 1 hr earlier -> 14:00-24:00
      return timed(15 * 60, 9 * 60);                                  // 15:00-24:00
    }
    if (family === 'O') return timed(13 * 60, 9 * 60);                // 13:00-22:00
    if (family === 'N' || family === 'TSWN') return timed(23 * 60 + 30, 9 * 60); // 23:30-08:30 (+1)
    if (family === 'M' || family === 'TSWM')
      return mod === '#' ? timed(11 * 60, 10 * 60) : timed(11 * 60, 9 * 60);    // 11:00-20:00 / 21:00
    if (family === 'RESP' || family === 'FUC' || family === 'C') return timed(9 * 60, 8 * 60);
    if (family === 'VL' || family === 'CCL')
      return { allDay: true, start: 0, duration: 0, label: 'All day', note: '' };
    return timed(9 * 60, 8 * 60);

    function timed(start, duration, note) {
      var end = start + duration;
      return {
        allDay: false, start: start, duration: duration, end: end,
        label: fmt(start) + ' – ' + fmtEnd(end),
        note: note || ''
      };
    }
  }

  function dutiesFor(roster, staffId, day) {
    var b = roster.byStaff[staffId];
    return (b && b[day]) || [];
  }
  function findStaff(roster, staffId) {
    for (var i = 0; i < roster.staff.length; i++) if (roster.staff[i].id === staffId) return roster.staff[i];
    return null;
  }
  function initialsFor(staff) {
    return String((staff && staff.key) || 'DUTY').replace(/[^A-Za-z0-9]/g, '') || 'DUTY';
  }

  /** Monthly badge counts (A / P / O / N / VL / OTHER / OFF / TOTAL). */
  function summary(roster, staffId) {
    var c = { A: 0, P: 0, O: 0, N: 0, VL: 0, OTHER: 0, OFF: 0, TOTAL: 0 };
    for (var d = 1; d <= roster.daysInMonth; d++) {
      var list = dutiesFor(roster, staffId, d);
      if (!list.length) { c.OFF++; continue; }
      for (var i = 0; i < list.length; i++) { c[categoryOf(list[i].code)]++; c.TOTAL++; }
    }
    return c;
  }

  /* ------------------------------------------------------------------ *
   * 6. .ics export (RFC 5545, Asia/Hong_Kong, all-day VL/CCL)
   * ------------------------------------------------------------------ */

  /** Local wall-clock date-time string from minutes since midnight on `day`. */
  function icsDateTime(year, month, day, minutes) {
    var d = new Date(Date.UTC(year, month - 1, 1));
    d.setUTCDate(d.getUTCDate() + day - 1 + Math.floor(minutes / 1440));
    var rest = ((minutes % 1440) + 1440) % 1440;
    return pad4(d.getUTCFullYear()) + pad2(d.getUTCMonth() + 1) + pad2(d.getUTCDate()) +
           'T' + pad2(Math.floor(rest / 60)) + pad2(rest % 60) + '00';
  }

  /** Fold a content line at 75 octets, UTF-8 aware. */
  function foldLine(line) {
    var enc = new TextEncoder();
    if (enc.encode(line).length <= 75) return line;
    var out = [], cur = '', len = 0, limit = 75;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i], bl = enc.encode(ch).length;
      if (len + bl > limit) { out.push(cur); cur = ' ' + ch; len = 1 + bl; limit = 74; }
      else { cur += ch; len += bl; }
    }
    out.push(cur);
    return out.join('\r\n');
  }

  function buildIcs(roster, staffId) {
    var staff = findStaff(roster, staffId);
    if (!staff) throw new Error('No doctor selected.');
    var initials = initialsFor(staff);
    var y = roster.year, mo = roster.month;
    var stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    var L = [];
    L.push('BEGIN:VCALENDAR');
    L.push('VERSION:2.0');
    L.push('PRODID:-//POH AED//Duty Calendar Converter//EN');
    L.push('CALSCALE:GREGORIAN');
    L.push('METHOD:PUBLISH');
    L.push('X-WR-CALNAME:' + esc(initials + ' Duty — ' + roster.monthName + ' ' + y));
    L.push('X-WR-TIMEZONE:Asia/Hong_Kong');
    L.push('BEGIN:VTIMEZONE');
    L.push('TZID:Asia/Hong_Kong');
    L.push('BEGIN:STANDARD');
    L.push('DTSTART:19790101T000000');
    L.push('TZOFFSETFROM:+0800');
    L.push('TZOFFSETTO:+0800');
    L.push('TZNAME:HKT');
    L.push('END:STANDARD');
    L.push('END:VTIMEZONE');

    var n = 0;
    for (var day = 1; day <= roster.daysInMonth; day++) {
      var list = dutiesFor(roster, staffId, day);
      for (var i = 0; i < list.length; i++) {         // "Off" days simply produce no event
        var duty = list[i], t = duty.timing || getTiming(duty.label);
        n++;
        L.push('BEGIN:VEVENT');
        L.push('UID:' + y + pad2(mo) + pad2(day) + '-' + n + '-' +
               initials.toLowerCase() + '@poh-aed-duty.local');
        L.push('DTSTAMP:' + stamp);
        if (t.allDay) {
          L.push('DTSTART;VALUE=DATE:' + icsDateTime(y, mo, day, 0).slice(0, 8));
          L.push('DTEND;VALUE=DATE:' + icsDateTime(y, mo, day + 1, 0).slice(0, 8));
        } else {
          L.push('DTSTART;TZID=Asia/Hong_Kong:' + icsDateTime(y, mo, day, t.start));
          L.push('DTEND;TZID=Asia/Hong_Kong:' + icsDateTime(y, mo, day, t.end));
        }
        L.push('SUMMARY:' + esc(initials + ' — ' + duty.label));
        var desc = [roster.monthName + ' ' + day + ', ' + y,
                    duty.block + ' block / column ' + duty.column,
                    t.allDay ? 'All-day duty' : 'Time: ' + t.label + ' (Asia/Hong_Kong)',
                    duty.note ? 'Roster note: ' + duty.note : '',
                    'Source: ' + roster.fileName].filter(Boolean).join(' | ');
        L.push('DESCRIPTION:' + esc(desc));
        L.push('LOCATION:POH AED');
        L.push('CATEGORIES:' + PALETTE[categoryOf(duty.code)].label);
        L.push('TRANSP:OPAQUE');
        L.push('END:VEVENT');
      }
    }
    L.push('END:VCALENDAR');
    return L.map(foldLine).join('\r\n') + '\r\n';
  }

  /* ------------------------------------------------------------------ *
   * 7. Styled .xlsx export (ExcelJS)
   * ------------------------------------------------------------------ */

  /**
   * Build (but do not download) the ExcelJS workbook. Kept DOM-free so it can
   * be unit-tested in Node with the `exceljs` npm package.
   */
  function buildExcelWorkbook(ExcelJS, roster, staffId, opts) {
    if (!ExcelJS) throw new Error('ExcelJS is not loaded.');
    opts = opts || {};
    var fillDateRow = opts.fillDateRow !== false;
    var palette = opts.palette ? clonePalette(opts.palette) : PALETTE;
    var codeColors = opts.codeColors || {};
    var staff = findStaff(roster, staffId);
    if (!staff) throw new Error('No doctor selected.');
    var initials = initialsFor(staff);
    var y = roster.year, mo = roster.month;

    var wb = new ExcelJS.Workbook();
    wb.creator = 'POH AED Duty Calendar Converter';
    wb.created = new Date();
    var ws = wb.addWorksheet('Duty Calendar', { views: [{ showGridLines: false }] });
    ws.columns = Array.from({ length: 7 }, function () { return { width: 18.5 }; });

    var thin = { style: 'thin', color: { argb: 'FFD1D5DB' } };
    var F = function (hex) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(hex) } }; };

    // Row 1 — title banner
    ws.mergeCells('A1:G1');
    ws.getRow(1).height = 34;
    for (var c = 1; c <= 7; c++) ws.getRow(1).getCell(c).fill = F('#1E3A5F');
    var title = ws.getCell('A1');
    title.value = roster.monthName + ' ' + y + ' | ' + initials + ' Duty Calendar';
    title.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    title.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    // Row 2 — subtitle
    ws.mergeCells('A2:G2');
    ws.getRow(2).height = 22;
    for (c = 1; c <= 7; c++) ws.getRow(2).getCell(c).fill = F('#F1F3F8');
    var sub = ws.getCell('A2');
    sub.value = 'MONTHLY ROSTER / SUNDAY TO SATURDAY';
    sub.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF6B7280' } };
    sub.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    // Row 3 — weekday headers
    ws.getRow(3).height = 25;
    for (c = 1; c <= 7; c++) {
      var hc = ws.getRow(3).getCell(c);
      hc.value = DAY_NAMES[c - 1];
      hc.fill = F('#2C4D75');
      hc.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      hc.alignment = { horizontal: 'center', vertical: 'middle' };
      hc.border = { bottom: thin };
    }

    // Calendar grid — each week = 1 date row + 1 shift row (Sunday first).
    // Compute from the date rather than trusting a cached value from an older build.
    var leading = leadingBlanks(y, mo);
    var weeks = Math.ceil((leading + roster.daysInMonth) / 7);
    var row = 4;
    for (var w = 0; w < weeks; w++) {
      var dateRow = ws.getRow(row), shiftRow = ws.getRow(row + 1);
      dateRow.height = 21;
      shiftRow.height = 32;
      for (var c2 = 1; c2 <= 7; c2++) {
        var day = w * 7 + c2 - leading;
        var inMonth = day >= 1 && day <= roster.daysInMonth;
        var duties = inMonth ? dutiesFor(roster, staffId, day) : [];
        var cat = duties.length ? categoryOf(duties[0].code) : 'OFF';
        var pal = duties.length ? colorsFor(duties[0].code, palette, codeColors) : palette.OFF;
        var dCell = dateRow.getCell(c2), sCell = shiftRow.getCell(c2);
        var text = !inMonth ? '' : (duties.length ? joinLabels(duties) : 'Off');

        dCell.value = inMonth ? day : null;
        dCell.fill = F(inMonth ? (fillDateRow ? pal.bg : '#FFFFFF') : '#F3F4F6');
        dCell.font = { name: 'Calibri', size: 11, bold: true,
                       color: { argb: argb(fillDateRow ? pal.fg : '#1F2937') } };
        dCell.alignment = { horizontal: 'left', vertical: 'top', indent: 1 };
        dCell.border = { top: thin, left: thin, right: thin };

        sCell.value = text || null;
        sCell.fill = F(inMonth ? pal.bg : '#F3F4F6');
        sCell.font = { name: 'Calibri', size: text.length > 9 ? 10 : 12, bold: true,
                       color: { argb: argb(cat === 'OFF' ? '#9CA3AF' : pal.fg) } };
        sCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        sCell.border = { bottom: thin, left: thin, right: thin };
      }
      row += 2;
    }

    // Bottom legend
    var titleRowNum = row + 1;
    ws.getRow(titleRowNum).height = 20;
    ws.mergeCells(titleRowNum, 1, titleRowNum, 7);
    var lt = ws.getRow(titleRowNum).getCell(1);
    lt.value = 'DUTY COLORS';
    lt.font = { name: 'Calibri', size: 10, bold: true, color: { argb: argb('#1E3A5F') } };
    lt.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    var legendRowNum = row + 2;
    ws.getRow(legendRowNum).height = 24;
    var legend = [['A duties', 'A'], ['P duties', 'P'], ['O duties', 'O'], ['N', 'N'],
                  ['VL', 'VL'], ['Resp / Other', 'OTHER'], ['Off', 'OFF']];
    for (var k = 0; k < 7; k++) {
      var lc = ws.getRow(legendRowNum).getCell(k + 1);
      var lp = palette[legend[k][1]];
      lc.value = legend[k][0];
      lc.fill = F(lp.bg);
      lc.font = { name: 'Calibri', size: 10, bold: true, color: { argb: argb(lp.fg) } };
      lc.alignment = { horizontal: 'center', vertical: 'middle' };
      lc.border = { top: thin, bottom: thin, left: thin, right: thin };
    }

    var noteRowNum = row + 3;
    ws.getRow(noteRowNum).height = 18;
    ws.mergeCells(noteRowNum, 1, noteRowNum, 7);
    var nc = ws.getRow(noteRowNum).getCell(1);
    nc.value = 'Source: ' + roster.fileName + '  •  Times for calendar export use Asia/Hong_Kong.';
    nc.font = { name: 'Calibri', size: 9, color: { argb: 'FF9CA3AF' } };
    nc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    // Print setup: landscape, fit to one page
    ws.pageSetup = {
      paperSize: 9,                       // A4
      orientation: 'landscape',
      fitToPage: true, fitToWidth: 1, fitToHeight: 1,
      horizontalCentered: true, verticalCentered: false,
      margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 }
    };
    ws.pageSetup.printArea = 'A1:G' + noteRowNum;
    return wb;
  }

  function excelFileName(roster, staffId) {
    return initialsFor(findStaff(roster, staffId)) + '_' + roster.year + '_' +
           pad2(roster.month) + '_Duty_Calendar.xlsx';
  }
  function icsFileName(roster, staffId) {
    return initialsFor(findStaff(roster, staffId)) + '_' + roster.year + '_' +
           pad2(roster.month) + '_Duty.ics';
  }

  /* ------------------------------------------------------------------ *
   * 8. Public API (pure logic)
   * ------------------------------------------------------------------ */
  var API = {
    PALETTE: PALETTE, CATEGORY_ORDER: CATEGORY_ORDER, CSS_KEY: CSS_KEY,
    EXCLUDED_STAFF: EXCLUDED_STAFF, COLOR_KEY: COLOR_KEY, CDN: CDN,
    MONTH_NAMES: MONTH_NAMES, DAY_NAMES: DAY_NAMES, DAY_ABBR: DAY_ABBR,
    clonePalette: clonePalette, colorsFor: colorsFor, autoFg: autoFg,
    allDutyCodes: allDutyCodes, codeKey: codeKey, joinLabels: joinLabels,
    parseWorkbook: parseWorkbook,
    categoryOf: categoryOf,
    getTiming: getTiming,
    summary: summary,
    dutiesFor: dutiesFor,
    findStaff: findStaff,
    initialsFor: initialsFor,
    leadingBlanks: leadingBlanks,
    weekdayIndex: weekdayIndex,
    buildIcs: buildIcs,
    buildExcelWorkbook: buildExcelWorkbook,
    excelFileName: excelFileName,
    icsFileName: icsFileName
  };

  /* ================================================================== *
   * 9. Browser UI — everything below is skipped in Node
   * ================================================================== */
  if (typeof document === 'undefined') return API;

  var $ = function (id) { return document.getElementById(id); };
  var state = {
    roster: null, staffId: null, view: 'calendar', filter: '', activeIndex: -1,
    palette: clonePalette(PALETTE), codeColors: {},
    autoText: { A: true, P: true, O: true, N: true, VL: true, OTHER: true, OFF: true }
  };

  function ready(fn) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fn);
    else fn();
  }
  function loadScript(urls) {
    return new Promise(function (resolve, reject) {
      function attempt(i) {
        if (i >= urls.length) return reject(new Error('Could not load a required library from any CDN.'));
        var s = document.createElement('script');
        s.src = urls[i];
        s.onload = resolve;
        s.onerror = function () { attempt(i + 1); };
        document.head.appendChild(s);
      }
      attempt(0);
    });
  }
  function ensureLibs() {
    return Promise.resolve()
      .then(function () { return root.XLSX ? null : loadScript(CDN.xlsx); })
      .then(function () { return root.ExcelJS ? null : loadScript(CDN.exceljs); });
  }

  var toastTimer;
  function toast(msg, kind) {
    var el = $('toast');
    el.textContent = msg;
    el.className = 'toast' + (kind ? ' is-' + kind : '');
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.hidden = true; }, 4200);
  }
  function setStatus(html, kind) {
    $('fileStatus').innerHTML = '<span class="status-' + (kind || 'idle') + '">' + html + '</span>';
  }

  function save() {
    try {
      if (!state.roster) { localStorage.removeItem(LS_KEY); return; }
      localStorage.setItem(LS_KEY, JSON.stringify({ roster: state.roster, staffId: state.staffId }));
    } catch (e) { /* storage full / private mode — keep working in memory */ }
  }
  function restore() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      if (!raw) return false;
      var o = JSON.parse(raw);
      if (o && o.roster && o.roster.version === 1 && o.roster.staff) {
        state.roster = o.roster;
        // Never auto-show a month: the doctor must be picked explicitly each session.
        state.staffId = null;
        // Recompute the leading blanks so older caches use the current Sunday-first layout.
        state.roster.leadingBlanks = leadingBlanks(state.roster.year, state.roster.month);
        return true;
      }
    } catch (e) { /* ignore corrupt cache */ }
    return false;
  }

  /* ----------------------- user colour preferences -------------------- */
  function loadColors() {
    try {
      var raw = localStorage.getItem(COLOR_KEY);
      if (!raw) return;
      var o = JSON.parse(raw);
      if (!o) return;
      if (o.palette) state.palette = clonePalette(o.palette);
      if (o.codeColors && typeof o.codeColors === 'object') {
        Object.keys(o.codeColors).forEach(function (k) {
          var v = o.codeColors[k];
          if (v && validHex(v.bg)) state.codeColors[k] = { bg: v.bg, fg: validHex(v.fg) ? v.fg : autoFg(v.bg) };
        });
      }
      CATEGORY_ORDER.forEach(function (cat) {
        if (o.autoText && typeof o.autoText[cat] === 'boolean') state.autoText[cat] = o.autoText[cat];
      });
    } catch (e) { /* ignore corrupt colour cache */ }
  }
  function saveColors() {
    try {
      localStorage.setItem(COLOR_KEY, JSON.stringify({
        palette: state.palette, codeColors: state.codeColors, autoText: state.autoText
      }));
    } catch (e) { /* private mode — keep working in memory */ }
  }
  /** Push the category palette into CSS variables used by every coloured element. */
  function applyCssPalette() {
    CATEGORY_ORDER.forEach(function (cat) {
      var p = state.palette[cat], k = CSS_KEY[cat], rootStyle = document.documentElement.style;
      rootStyle.setProperty('--' + k + '-bg', p.bg);
      rootStyle.setProperty('--' + k + '-fg', p.fg);
    });
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('Could not read the file.')); };
      fr.readAsArrayBuffer(file);
    });
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) { toast('Please choose an .xlsx file.', 'error'); return; }
    var dz = $('dropzone');
    dz.classList.add('is-busy');
    setStatus('Reading ' + html(file.name) + '…', 'warn');
    try {
      var data = await readFileAsArrayBuffer(file);
      var roster = API.parseWorkbook(data, file.name);
      state.roster = roster;
      state.staffId = null;                          // wait for an explicit pick in Step 2
      state.filter = '';
      $('staffSearch').value = '';
      save();
      renderAll(true);
      setStatus('Loaded <strong>' + html(roster.monthName + ' ' + roster.year) +
                '</strong> from ' + html(file.name) + ' &middot; ' + roster.staff.length +
                ' staff with duties. Pick a doctor in Step 2 to see their month.', 'ok');
      var wmsg = roster.warnings.length
        ? ' Parsed with ' + roster.warnings.length + ' note(s) — see Parse report.' : '';
      toast('Roster parsed successfully.' + wmsg, roster.warnings.length ? 'warn' : null);
    } catch (err) {
      console.error(err);
      setStatus('Could not parse this file: ' + html(err.message || String(err)), 'error');
      toast('Parse failed — see the message under Step 1.', 'error');
    } finally {
      dz.classList.remove('is-busy');
      $('fileInput').value = '';
    }
  }

  /* --------------------------- rendering ---------------------------- */
  function renderAll(fromUpload) {
    var r = state.roster;
    var hasRoster = !!r, hasStaff = !!(r && state.staffId);
    $('step2').classList.toggle('locked', !hasRoster);
    $('step3').classList.toggle('locked', !hasStaff);
    $('step4').classList.toggle('locked', !hasStaff);
    $('clearBtn').classList.toggle('hidden', !hasRoster);
    $('exportXlsx').disabled = !hasStaff;
    $('exportIcs').disabled = !hasStaff;
    $('calendarEmpty').hidden = hasStaff;
    $('calendarView').hidden = !hasStaff || state.view !== 'calendar';
    $('agendaView').hidden   = !hasStaff || state.view !== 'agenda';

    // Parse report
    var diag = $('diagnostics');
    diag.hidden = !r;
    if (r) {
      $('diagnosticsText').textContent =
        'Sheet name        : ' + r.sheetName + '\n' +
        'Month / year      : ' + r.monthName + ' ' + r.year + '  (from ' + r.monthSource + ')\n' +
        'Days in month     : ' + r.daysInMonth + '   |   day rows found: ' + r.dayRowCount + '\n' +
        'Day columns       : Senior ' + r.seniorDayCol + ' / MO ' + r.moDayCol + '\n' +
        'Senior headers    : ' + r.seniorHeaders.join(', ') + '\n' +
        'MO headers        : ' + r.moHeaders.join(', ') + '\n' +
        'Timing notes      : ' + Object.keys(r.timingNotes).sort().join(', ') + '\n' +
        'Staff with duties : ' + r.staff.length + '\n' +
        'Source file       : ' + r.fileName + '\n' +
        (r.warnings.length ? '\nNotes:\n - ' + r.warnings.join('\n - ') : '\nNo warnings.');
    }

    renderCombo();
    renderSummary();
    renderMonthTitle();
    renderCalendar();
    renderAgenda();
    renderLegend();
    renderColorUI();
  }

  function matchedStaff() {
    var r = state.roster;
    if (!r) return [];
    var q = norm(state.filter);
    return r.staff.filter(function (s) {
      return !q || norm(s.key).indexOf(q) >= 0 || norm(s.name).indexOf(q) >= 0 ||
             norm(s.label).indexOf(q) >= 0 || norm(s.group).indexOf(q) >= 0;
    });
  }
  function renderCombo() {
    var list = $('comboList');
    var items = matchedStaff();
    var selected = state.roster ? API.findStaff(state.roster, state.staffId) : null;
    $('selectedLine').textContent = selected
      ? selected.label + '  ·  ' + selected.dutyCount + ' duties in ' + state.roster.monthName
      : 'Select someone to see their month.';
    if (!items.length) { list.innerHTML = '<div class="combo-empty">No matching staff</div>'; return; }
    // Grouped exactly as in the lower-left staff directory of the workbook.
    var groups = {};
    items.forEach(function (s) {
      var g = s.group || 'Discovered';
      if (!groups[g]) groups[g] = [];
      groups[g].push(s);
    });
    var order = ['Senior', 'Junior', 'TMH/TSWH', 'Discovered'];
    var names = Object.keys(groups).sort(function (a, b) {
      var ia = order.indexOf(a), ib = order.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b);
    });
    list.innerHTML = names.map(function (g) {
      return '<div class="combo-group" data-group="' + html(g) + '"><span>' + html(g) +
             '</span><span class="combo-group-n">' + groups[g].length + '</span></div>' +
             groups[g].map(function (s) {
               return '<div class="combo-item' + (s.id === state.staffId ? ' is-selected' : '') +
                      (items.indexOf(s) === state.activeIndex ? ' is-active' : '') +
                      '" role="option" data-id="' + html(s.id) + '" data-group="' + html(g) +
                      '" aria-selected="' + (s.id === state.staffId) + '">' +
                      '<span class="combo-label">' + html(s.label) + '</span>' +
                      '<span class="combo-meta">' + s.dutyCount + ' duties</span></div>';
             }).join('');
    }).join('');
  }
  function openCombo(show) {
    $('comboList').hidden = !show;
    $('staffSearch').setAttribute('aria-expanded', String(!!show));
    state.activeIndex = -1;
    if (show) renderCombo();
  }
  function selectStaff(id) {
    state.staffId = id;
    state.filter = '';
    state.activeIndex = -1;
    $('staffSearch').value = '';
    openCombo(false);
    save();
    renderAll(false);
    setView(state.view);
  }

  function renderSummary() {
    var el = $('summary');
    if (!state.roster || !state.staffId) { el.innerHTML = ''; return; }
    var c = API.summary(state.roster, state.staffId);
    var items = [['Total duties', c.TOTAL, 'badge-total']];
    API.CATEGORY_ORDER.forEach(function (cat) {
      var key = cat === 'OTHER' ? 'Resp / Other' : (cat === 'OFF' ? 'Off days' : state.palette[cat].label);
      items.push([key, c[cat], 'badge-' + cat]);
    });
    el.innerHTML = items.map(function (it) {
      return '<span class="badge ' + it[2] + '">' + html(it[0]) +
             ' <span class="badge-n">' + it[1] + '</span></span>';
    }).join('');
    // Apply the current (possibly user-edited) palette to each badge.
    API.CATEGORY_ORDER.forEach(function (cat, i) {
      var node = el.children[i + 1];
      if (!node) return;
      var p = state.palette[cat];
      node.style.setProperty('--badge-bg', p.bg);
      node.style.setProperty('--badge-fg', p.fg);
      node.style.setProperty('--badge-line', p.bg);
    });
  }

  function renderMonthTitle() {
    var r = state.roster, s = r ? API.findStaff(r, state.staffId) : null;
    $('monthTitle').textContent = r
      ? r.monthName + ' ' + r.year + '  ·  ' + API.initialsFor(s) + '  ·  ' +
        (s ? s.dutyCount : 0) + ' duties'
      : '—';
  }

  var today = new Date();
  function renderCalendar() {
    var grid = $('calendarGrid');
    var r = state.roster;
    if (!r || !state.staffId) { grid.innerHTML = ''; return; }
    var cells = [];
    var leading = leadingBlanks(r.year, r.month);                 // Sunday first
    for (var i = 0; i < leading; i++) cells.push(null);
    for (var d = 1; d <= r.daysInMonth; d++) cells.push(d);
    while (cells.length % 7 !== 0) cells.push(null);

    grid.innerHTML = cells.map(function (day) {
      if (day === null) return '<div class="cal-cell cal-empty"></div>';
      var list = API.dutiesFor(r, state.staffId, day);
      var cat = list.length ? API.categoryOf(list[0].code) : 'OFF';
      var pal = list.length ? API.colorsFor(list[0].code, state.palette, state.codeColors)
                            : state.palette.OFF;
      var text = list.length ? API.joinLabels(list) : 'Off';      // e.g. A5/FUC
      var notes = list.map(function (x) {
        return x.label + ' ' + x.timing.label + (x.note ? ' — ' + x.note : '');
      }).join('\n');
      var isToday = r.year === today.getFullYear() && r.month === today.getMonth() + 1 &&
                    day === today.getDate();
      var longClass = text.length > 9 ? ' is-long' : '';
      return '<div class="cal-cell cat-' + cat + (isToday ? ' is-today' : '') +
             '" style="--cell-bg:' + pal.bg + ';--cell-fg:' + pal.fg + '"' +
             ' title="' + html(notes || 'No duty') + '">' +
             '<div class="cal-num">' + day + '</div>' +
             '<div class="cal-code' + (list.length ? longClass : ' is-off') + '">' +
             html(text) + '</div></div>';
    }).join('');
  }

  function renderAgenda() {
    var el = $('agendaView');
    var r = state.roster;
    if (!r || !state.staffId) { el.innerHTML = ''; return; }
    var rows = [];
    for (var d = 1; d <= r.daysInMonth; d++) {
      var list = API.dutiesFor(r, state.staffId, d);
      var pal = list.length ? API.colorsFor(list[0].code, state.palette, state.codeColors)
                            : state.palette.OFF;
      var wd = API.weekdayIndex(r.year, r.month, d);
      // Paired duties are shown as one chip: A5/FUC rather than two loose codes.
      var chipPal = list.length ? API.colorsFor(list[0].code, state.palette, state.codeColors)
                                : state.palette.OFF;
      var chips = '<span class="chip" style="--chip-bg:' + chipPal.bg + ';--chip-fg:' +
                  chipPal.fg + '">' + html(list.length ? API.joinLabels(list) : 'Off') + '</span>';
      var times = list.map(function (x) {
        return '<strong>' + html(x.label) + '</strong> ' + html(x.timing.label) +
               (x.note ? ' · ' + html(x.note) : '') +
               (x.extra ? ' · ' + html(x.extra) : '');
      }).join(' &nbsp;|&nbsp; ') || 'No duty assigned';
      rows.push('<div class="agenda-row' + (list.length ? '' : ' is-off') +
                '" style="--row-bg:' + pal.bg + ';--row-fg:' + pal.fg + '">' +
                '<div class="agenda-date"><span class="agenda-dow">' + DAY_ABBR[wd] +
                '</span><span class="agenda-d">' + d + '</span></div>' +
                '<div class="agenda-body"><div class="agenda-chips">' + chips + '</div>' +
                '<div class="agenda-time">' + times + '</div></div></div>');
    }
    el.innerHTML = rows.join('');
  }

  function renderLegend() {
    $('legend').innerHTML = API.CATEGORY_ORDER.map(function (cat) {
      var p = state.palette[cat];
      return '<div class="legend-item" style="--lg-bg:' + p.bg + ';--lg-fg:' + p.fg + '">' +
             html(p.label) + '</div>';
    }).join('');
  }

  /* ---------------------- interactive colour editor ------------------- */
  function refreshColouredViews() {
    applyCssPalette();
    renderSummary(); renderCalendar(); renderAgenda(); renderLegend();
  }

  function renderColorUI() {
    if (!$('colorGrid')) return;
    $('colorGrid').innerHTML = API.CATEGORY_ORDER.map(function (cat) {
      var p = state.palette[cat], auto = state.autoText[cat];
      return '<div class="color-row">' +
             '<div class="color-preview" data-preview="' + cat + '" style="background:' + p.bg +
               ';color:' + p.fg + '">Aa</div>' +
             '<div class="color-name">' + html(p.label) + '</div>' +
             '<label class="color-field">BG <input type="color" data-kind="bg" data-cat="' + cat +
               '" value="' + p.bg + '" aria-label="' + html(p.label) + ' background"></label>' +
             '<label class="color-field">Text <input type="color" data-kind="fg" data-cat="' + cat +
               '" value="' + p.fg + '"' + (auto ? ' disabled' : '') + ' aria-label="' + html(p.label) +
               ' text colour"></label>' +
             '<label class="color-auto"><input type="checkbox" data-kind="auto" data-cat="' + cat +
               '"' + (auto ? ' checked' : '') + '> Auto</label></div>';
    }).join('');

    var codes = state.roster ? API.allDutyCodes(state.roster) : [];
    $('codeColorCount').textContent = codes.length ? '(' + codes.length + ' in this roster)' : '';
    $('codeColorList').innerHTML = codes.map(function (x) {
      var inherited = !state.codeColors[x.key];
      var pal = API.colorsFor(x.code, state.palette, state.codeColors);
      return '<div class="color-row code-row">' +
             '<label class="color-inherit"><input type="checkbox" data-kind="inherit" data-code="' +
               html(x.key) + '"' + (inherited ? ' checked' : '') + '> Inherit</label>' +
             '<div class="color-name">' + html(x.label) + ' <small>&times;' + x.count + '</small></div>' +
             '<label class="color-field">BG <input type="color" data-kind="code-bg" data-code="' +
               html(x.key) + '" value="' + pal.bg + '"' + (inherited ? ' disabled' : '') +
               ' aria-label="' + html(x.label) + ' colour"></label></div>';
    }).join('');
  }

  function bindColorUI() {
    var panel = $('colorPanel');
    panel.addEventListener('input', function (e) {
      var t = e.target, kind = t.getAttribute('data-kind');
      if (kind === 'bg' || kind === 'fg') {
        var cat = t.getAttribute('data-cat');
        if (kind === 'bg') {
          state.palette[cat].bg = t.value;
          if (state.autoText[cat]) state.palette[cat].fg = autoFg(t.value);
        } else {
          state.palette[cat].fg = t.value;
        }
        var p = state.palette[cat];
        var prev = panel.querySelector('[data-preview="' + cat + '"]');
        if (prev) { prev.style.background = p.bg; prev.style.color = p.fg; }
        var fgInput = panel.querySelector('input[data-kind="fg"][data-cat="' + cat + '"]');
        if (fgInput) fgInput.value = p.fg;
        refreshColouredViews(); saveColors();
      } else if (kind === 'code-bg') {
        var code = t.getAttribute('data-code');
        state.codeColors[code] = { bg: t.value, fg: autoFg(t.value) };
        refreshColouredViews(); saveColors();
      }
    });
    panel.addEventListener('change', function (e) {
      var t = e.target, kind = t.getAttribute('data-kind');
      if (kind === 'auto') {
        var cat = t.getAttribute('data-cat');
        state.autoText[cat] = t.checked;
        if (t.checked) state.palette[cat].fg = autoFg(state.palette[cat].bg);
        renderColorUI(); refreshColouredViews(); saveColors();
      } else if (kind === 'inherit') {
        var code = t.getAttribute('data-code');
        if (t.checked) delete state.codeColors[code];
        else {
          var base = state.palette[API.categoryOf(code)];
          state.codeColors[code] = { bg: base.bg, fg: base.fg };
        }
        renderColorUI(); refreshColouredViews(); saveColors();
      }
    });
    $('resetColors').addEventListener('click', function () {
      state.palette = clonePalette(PALETTE);
      state.codeColors = {};
      CATEGORY_ORDER.forEach(function (cat) { state.autoText[cat] = true; });
      saveColors(); applyCssPalette(); renderColorUI(); refreshColouredViews();
      toast('Duty colours reset to the pastel defaults.');
    });
  }

  function setView(view) {
    state.view = view;
    var hasStaff = !!(state.roster && state.staffId);
    $('calendarView').hidden = !hasStaff || view !== 'calendar';
    $('agendaView').hidden   = !hasStaff || view !== 'agenda';
    $('viewCalendar').classList.toggle('is-active', view === 'calendar');
    $('viewAgenda').classList.toggle('is-active', view === 'agenda');
    $('viewCalendar').setAttribute('aria-selected', String(view === 'calendar'));
    $('viewAgenda').setAttribute('aria-selected', String(view === 'agenda'));
  }

  /* --------------------------- exports ------------------------------- */
  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename; a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
  }

  async function exportXlsx() {
    if (!state.roster || !state.staffId) return;
    var btn = $('exportXlsx');
    btn.disabled = true;
    try {
      var wb = API.buildExcelWorkbook(root.ExcelJS, state.roster, state.staffId, {
        fillDateRow: $('fillDateRow').checked,
        palette: state.palette,
        codeColors: state.codeColors
      });
      var buf = await wb.xlsx.writeBuffer();
      downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
                   API.excelFileName(state.roster, state.staffId));
      toast('Excel calendar downloaded.');
    } catch (e) {
      console.error(e);
      toast('Excel export failed: ' + (e.message || e), 'error');
    } finally { btn.disabled = false; }
  }

  function exportIcs() {
    if (!state.roster || !state.staffId) return;
    try {
      var ics = API.buildIcs(state.roster, state.staffId);
      downloadBlob(new Blob([ics], { type: 'text/calendar;charset=utf-8' }),
                   API.icsFileName(state.roster, state.staffId));
      toast('Calendar file downloaded — open it to add the shifts.');
    } catch (e) {
      console.error(e);
      toast('Calendar export failed: ' + (e.message || e), 'error');
    }
  }

  /* --------------------------- wiring -------------------------------- */
  function bindUI() {
    bindColorUI();
    var dz = $('dropzone'), input = $('fileInput');
    // A programmatic input.click() bubbles back up to the dropzone. Stop it here,
    // otherwise the dropzone handler would call input.click() again forever.
    input.addEventListener('click', function (e) { e.stopPropagation(); });
    $('chooseBtn').addEventListener('click', function (e) { e.stopPropagation(); input.click(); });
    dz.addEventListener('click', function (e) { if (e.target !== input) input.click(); });
    dz.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', function () { if (input.files[0]) handleFile(input.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('is-drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('is-drag'); });
    });
    dz.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    });
    // Stop the whole page from navigating when a file is dropped outside the zone.
    ['dragover', 'drop'].forEach(function (ev) {
      document.addEventListener(ev, function (e) { e.preventDefault(); });
    });

    $('clearBtn').addEventListener('click', function () {
      state.roster = null; state.staffId = null; state.filter = '';
      $('staffSearch').value = '';
      save(); renderAll(false);
      setStatus('Saved roster removed. Upload a .xlsx to start again.', 'idle');
      toast('Saved roster cleared.');
    });

    // Searchable combobox
    var search = $('staffSearch'), listEl = $('comboList');
    search.addEventListener('focus', function () { openCombo(true); });
    search.addEventListener('input', function () {
      state.filter = search.value; state.activeIndex = -1; renderCombo(); openCombo(true);
    });
    search.addEventListener('keydown', function (e) {
      var items = matchedStaff();
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault(); openCombo(true);
        state.activeIndex = Math.max(-1, Math.min(items.length - 1,
          state.activeIndex + (e.key === 'ArrowDown' ? 1 : -1)));
        renderCombo();
        var act = listEl.querySelector('.is-active');
        if (act) act.scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter') {
        e.preventDefault();
        var pick = items[state.activeIndex >= 0 ? state.activeIndex : 0];
        if (pick) selectStaff(pick.id);
      } else if (e.key === 'Escape') { openCombo(false); search.blur(); }
    });
    $('comboToggle').addEventListener('click', function (e) {
      e.stopPropagation(); openCombo($('comboList').hidden);
      if (!$('comboList').hidden) search.focus();
    });
    // `click` (not mousedown) so taps on iPhone/Android Chrome select reliably.
    listEl.addEventListener('click', function (e) {
      var item = e.target.closest ? e.target.closest('.combo-item') : null;
      if (item) { selectStaff(item.getAttribute('data-id')); search.blur(); }
    });
    document.addEventListener('click', function (e) {
      if (!$('combo').contains(e.target)) openCombo(false);
    });

    $('viewCalendar').addEventListener('click', function () { setView('calendar'); });
    $('viewAgenda').addEventListener('click', function () { setView('agenda'); });
    $('exportXlsx').addEventListener('click', exportXlsx);
    $('exportIcs').addEventListener('click', exportIcs);
  }

  function bootstrap() {
    bindUI();
    loadColors();
    applyCssPalette();
    var restored = restore();
    state.view = window.innerWidth < 640 ? 'agenda' : 'calendar';
    renderAll(false);
    setView(state.view);
    if (restored) {
      var r = state.roster;
      setStatus('Restored <strong>' + html(r.monthName + ' ' + r.year) +
                '</strong> from this device &middot; ' + html(r.fileName) +
                '. Pick a doctor in Step 2 to see the month.', 'ok');
      toast('Saved roster restored.');
    }
  }

  ready(function () {
    ensureLibs().then(bootstrap).catch(function (e) {
      console.error(e);
      setStatus('A required library could not load: ' + html(e.message) +
                '. Check your internet connection and reload.', 'error');
      toast('Library load failed — check your connection.', 'error');
    });
  });

  return API;
});
