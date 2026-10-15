/*
  Bulch Enpack for Adobe Illustrator - table/category common build
  Version: 3.1.0 (FIXED - enhanced fuzzy matching)
  
  Data source:
    build/layout-data.json generated from the address table.
  
  Output:
    One or more common sheet_###.ai files in the out folder.
    No per-address .ai files are created.
  
  Fixes applied:
    [NORM]  trimString: removes zero-width, NBSP, soft hyphen and other invisible chars
    [NORM]  normalizeName: NBSP->space, unified quotes, ё->е, collapse spaces
    [NORM]  looseNameKey: all dash types -> space, remove quotes, alnum+cyrillic only
    [FIND]  findArtboardByName: cascading search (exact→loose→compact→fuzzy→word-fallback)
    [FIND]  FUZZY_THRESHOLD lowered from 0.82 to 0.75
    [FIND]  Added word-fallback: matches if ≥60% of significant words (len≥3) overlap
    [FIND]  Better logging: shows match method and confidence %
*/

#target illustrator

(function () {

// ─────────────────────────────────────────────────────────────────────────────
//  SETTINGS
// ─────────────────────────────────────────────────────────────────────────────
var SETTINGS = {
    OUTPUT_FOLDER_NAME: 'out',
    OUTPUT_BASENAME:    'sheet_',
    
    DATA_JSON_RELATIVE: 'build/layout-data.json',
    DATA_JSON_FALLBACKS: [
        'C:/BULCH/build/layout-data.json'
    ],

    // The Excel file is selected at every launch and converted before layout.
    // A dedicated temporary JSON prevents a previous table from being reused.
    XLSX_CONVERTER_RELATIVE:    'build/convert-xlsx-to-layout-json.js',
    CURRENT_DATA_JSON_RELATIVE: 'build/layout-data.current.json',
    NODE_BIN:                   'node',
    
    // Sheet size — set by the settings dialog at launch.
    SHEET_WIDTH_MM:  3000,
    SHEET_HEIGHT_MM: 2000,
    SHEET_MARGIN_MM: 0,
    
    ARTBOARD_GAP_MM: 3,
    COLUMN_GAP_MM:   3,
    
    // Extra space applies only to the address-name header, not between blocks.
    LABEL_HEIGHT_MM: 22,
    LABEL_WIDTH_MM:  100,
    LABEL_FONT_MM:   10,
    LABEL_FRAME_PAD_MM: 2,
    LABEL_BLOCK_INSET_MM: 4,
    LABEL_FONT_NAMES: [
        'ArialNarrow-Bold',
        'Arial Narrow Bold',
        'Arial-BoldMT',
        'ArialMT'
    ],
    
    DRAW_ADDRESS_OUTLINE: true,
    OUTLINE_STROKE_MM:    1,
    
    ROTATION_THRESHOLD_MM: 140,
    CLIP_TO_ARTBOARD:      true,
    
    PDF_COMPATIBLE:  false,
    COMPRESSED_SAVE: false,
    
    DELETE_GUIDES:       true,
    CLEAN_ON_COPY_ERROR: true,
    
    TURBO_MODE:              true,
    REDRAW_EVERY_N_ADDRESSES: 0,
    RENDER_MAINTENANCE_EVERY_CELLS: 75,
    RENDER_MAINTENANCE_DELAY_MS:    100,
    RENDER_ART_CACHE:              true,
    PASTE_RETRY_COUNT:              3,
    PASTE_RETRY_DELAY_MS:           250,
    MATCH_REVIEW_ENABLED:           true,
    MATCH_REVIEW_THRESHOLD:         0.75,
    MATCH_REVIEW_AMBIGUITY_DELTA:   0.08,
    // Show every non-exact match for approval before any output is created.
    MATCH_REVIEW_NON_EXACT:         true,
    MISS_REVIEW_THRESHOLD:          0.75,
    MATCH_REVIEW_MAX_ITEMS:         80,
    
    // Fuzzy matching threshold (0.0–1.0). Lowered to 0.75 for better coverage.
    FUZZY_THRESHOLD: 0.75
};

// ─────────────────────────────────────────────────────────────────────────────
//  GLOBALS
// ─────────────────────────────────────────────────────────────────────────────
var PT = 72 / 25.4;   // mm -> pt
var MM = 25.4 / 72;   // pt -> mm
var MARKER_PFX = '__BEP__';
var OPERATIONAL_PFX = '__BEP_OPERATIONAL__';

var logs = [];
var errs = [];
var renderAttemptCounter = 0;
var lastArtboardMatchInfo = null;
var matchReviewSeen = {};

var runStats = {
    startedAt:          '',
    finishedAt:         '',
    startedMs:          0,
    dataFile:           '',
    sourceTableFile:    '',
    inputRows:          0,
    expectedCopies:     0,
    tableAddressOrder:  [],
    tableByAddress:     {},
    rowIssues:          [],
    matchReviews:       [],
    outputFiles:        [],
    fatalError:         '',
    renderAttempts:     0,
    pasteRetryCells:    0,
    pasteRetryAttempts: 0,
    pasteRetryFailures: 0,
    artCacheCopies:     0,
    artCachePlacements: 0,
    artCacheHits:       0,
    artCacheFallbacks:  0,
    artCacheDupFails:   0,
    maintenanceRuns:    0,
    timings: {
        loadMs:     0,
        scanMs:     0,
        layoutMs:   0,
        renderMs:   0,
        finalizeMs: 0,
        saveMs:     0,
        reportMs:   0,
        totalMs:    0
    }
};

// ─────────────────────────────────────────────────────────────────────────────
//  UTILITIES
// ─────────────────────────────────────────────────────────────────────────────
function mm(v)  { return Number(v) * PT; }
function pt(v)  { return Number(v) * MM; }
function r1(v)  { return Math.round(pt(v) * 10) / 10; }
function log(s) { logs.push(String(s)); }
function warn(s){ errs.push(String(s)); logs.push('WARN: ' + s); }

function nowMs()  { return (new Date()).getTime(); }
function isoNow() {
    var d = new Date();
    function p(v) { return v < 10 ? '0' + v : String(v); }
    return d.getFullYear() + '-' + p(d.getMonth()+1) + '-' + p(d.getDate()) + ' ' +
           p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}
function msText(v) {
    v = v || 0;
    return String(v) + ' ms (' + (Math.round(v / 100) / 10) + ' s)';
}
function safeGC()      { try { $.gc(); } catch(e) {} }
function safeRedraw()  { try { app.redraw(); } catch(e) {} }
function safeSleep(ms) { try { $.sleep(ms); } catch(e) {} }
function maybeRedraw(step, every) {
    if (SETTINGS.TURBO_MODE) return;
    if (!every || every < 1) return;
    if (step % every === 0) safeRedraw();
}
function scriptFolder() {
    return $.fileName ? File($.fileName).parent : Folder.current;
}
function ensureDir(f) {
    if (!f.exists && !f.create()) throw new Error('Cannot create: ' + f.fsName);
    return f;
}
function pad(v, sz) {
    var s = String(v);
    while (s.length < sz) s = '0' + s;
    return s;
}
function stripExt(n) { return String(n || '').replace(/\.ai$/i, ''); }
function decodedName(f) {
    var n = f.displayName || f.name;
    try { n = decodeURI(n); } catch(e) {}
    return n;
}

// ─────────────────────────────────────────────────────────────────────────────
//  STATISTICS
// ─────────────────────────────────────────────────────────────────────────────
function getTableAddressStats(label) {
    label = label || 'output';
    if (!runStats.tableByAddress[label]) {
        runStats.tableByAddress[label] = {
            label: label,
            expectedRows: 0,
            matchedRows: 0,
            placedRows: 0,
            renderFailedRows: 0,
            missingRows: 0,
            rows: []
        };
        runStats.tableAddressOrder.push(label);
    }
    return runStats.tableByAddress[label];
}
function recordTableRow(item) {
    var label = item && item.address ? item.address : 'output';
    var st = getTableAddressStats(label);
    st.expectedRows++;
    st.rows.push(item && item.rowNumber ? item.rowNumber : '?');
}
function recordMatchedRow(label)      { getTableAddressStats(label).matchedRows++; }
function recordPlacedRow(label)       { getTableAddressStats(label).placedRows++; }
function recordRenderFailedRow(label) { getTableAddressStats(label).renderFailedRows++; }
function recordIssue(kind, item, message) {
    var label = item && item.address ? item.address : 'output';
    getTableAddressStats(label).missingRows++;
    runStats.rowIssues.push({
        kind:           kind,
        rowNumber:      item && item.rowNumber      ? item.rowNumber      : '?',
        address:        label,
        group:          item && item.group          ? item.group          : '',
        product:        item && item.product        ? item.product        : '',
        sourceFileName: item && item.sourceFileName ? item.sourceFileName : '',
        message:        message
    });
}

function scorePct(score) {
    return Math.round((score || 0) * 100);
}
function isExactMatchMethod(method) {
    return method === 'exact' || method === 'loose' || method === 'compact';
}
function shouldReviewMatch(info) {
    if (!SETTINGS.MATCH_REVIEW_ENABLED || !info) return false;
    // Every miss must be shown, even if the suggested alternative is weak.
    if (info.idx < 0) return true;
    if (SETTINGS.MATCH_REVIEW_NON_EXACT && !isExactMatchMethod(info.method)) return true;
    if ((info.score || 0) < SETTINGS.MATCH_REVIEW_THRESHOLD) return true;
    if (info.secondScore && (info.score - info.secondScore) <= SETTINGS.MATCH_REVIEW_AMBIGUITY_DELTA)
        return true;
    return false;
}
function recordMatchReview(item, info) {
    if (!shouldReviewMatch(info)) return;
    var rowNumber = item && item.rowNumber ? item.rowNumber : '?';
    var address = item && item.address ? item.address : 'output';
    var product = item && item.product ? item.product : '';
    var sourceFileName = item && item.sourceFileName ? item.sourceFileName : '';
    var key = '$' + sourceFileName + '|' + product + '|' +
              (info.name || '') + '|' + (info.method || '') + '|' + String(info.idx);
    if (matchReviewSeen.hasOwnProperty(key)) {
        matchReviewSeen[key].count++;
        if (matchReviewSeen[key].rows.length < 120)
            matchReviewSeen[key].rows += ', ' + rowNumber;
        return;
    }
    var rec = {
        rowNumber:      item && item.rowNumber ? item.rowNumber : '?',
        rows:           String(rowNumber),
        count:          1,
        address:        address,
        product:        product,
        sourceFileName: sourceFileName,
        accepted:       info.idx >= 0,
        method:         info.method || '',
        score:          info.score || 0,
        matchedName:    info.name || '',
        secondName:     info.secondName || '',
        secondScore:    info.secondScore || 0
    };
    matchReviewSeen[key] = rec;
    runStats.matchReviews.push(rec);
}
function matchReviewLine(item, index) {
    var status = item.accepted ? 'AUTO' : 'MISS';
    var s = (index + 1) + '. [' + status + ' ' + scorePct(item.score) + '% ' + item.method + '] ' +
            'rows ' + item.rows + (item.count > 1 ? ' (x' + item.count + ')' : '') +
            ' | ' + item.sourceFileName + '\r' +
            '   table: ' + item.product + '\r' +
            '   artboard: ' + (item.matchedName || '(none)');
    return s;
}
function matchReviewText(items) {
    var limit = Math.min(items.length, SETTINGS.MATCH_REVIEW_MAX_ITEMS || 80);
    var lines = [
        'Проверка соответствий названий таблицы и монтажных областей перед рендером.',
        'Продолжайте только если эти соответствия ожидаемые.',
        ''
    ];
    for (var i = 0; i < limit; i++) {
        lines.push(matchReviewLine(items[i], i));
        lines.push('');
    }
    if (items.length > limit)
        lines.push('... and ' + (items.length - limit) + ' more items. See report.txt after run.');
    return lines.join('\r');
}
function showMatchReviewDialog(items) {
    if (!items || !items.length) return true;
    var text = matchReviewText(items);
    try {
        var w = new Window('dialog', 'Проверка соответствий');
        w.orientation = 'column';
        w.alignChildren = ['fill', 'top'];
        w.add('statictext', undefined, 'Подозрительные соответствия: ' + items.length);
        var box = w.add('edittext', undefined, text, { multiline: true, scrolling: true });
        box.preferredSize = [900, 520];
        var buttons = w.add('group');
        buttons.alignment = 'right';
        var cancelBtn = buttons.add('button', undefined, 'Отменить рендер');
        var okBtn = buttons.add('button', undefined, 'Продолжить');
        cancelBtn.onClick = function() { w.close(0); };
        okBtn.onClick = function() { w.close(1); };
        return w.show() === 1;
    } catch(e) {
        return confirm(text + '\r\rПродолжить рендер?');
    }
}

function preflightText() {
    var lines = [];
    var missingSources = countIssuesByKind('missing-source');
    var missingArtboards = countIssuesByKind('missing-artboard');
    lines.push('Pre-build check');
    lines.push('Table rows: ' + runStats.inputRows);
    lines.push('Uncertain matches or missing items: ' + runStats.matchReviews.length);
    lines.push('Missing source files: ' + missingSources + '; missing artboards: ' + missingArtboards);
    lines.push('');
    if (!runStats.matchReviews.length && !missingSources && !missingArtboards) {
        lines.push('All items were matched exactly.');
    } else {
        lines.push('Review the proposed replacements below.');
        for (var i = 0; i < runStats.matchReviews.length; i++) {
            var item = runStats.matchReviews[i];
            var suggested = item.matchedName || '(no suitable replacement found)';
            lines.push((i + 1) + '. Row ' + item.rows + ': ' + item.product);
            lines.push('   Suggested: ' + suggested + ' — ' + scorePct(item.score) + '% (' + item.method + ')');
        }
        for (var j = 0; j < runStats.rowIssues.length; j++) {
            var issue = runStats.rowIssues[j];
            if (issue.kind === 'missing-source')
                lines.push('Source file missing, row ' + issue.rowNumber + ': ' + issue.message);
        }
    }
    lines.push('');
    lines.push('Continue accepts the listed automatic matches and starts the build.');
    lines.push('Cancel stops the operation before layout and output files are created.');
    return lines.join('\r');
}
function showPreflightDialog() {
    var text = preflightText();
    try {
        var w = new Window('dialog', 'Pre-build check');
        w.orientation = 'column';
        w.alignChildren = ['fill', 'top'];
        w.add('statictext', undefined, 'Check matches before starting the build.');
        var box = w.add('edittext', undefined, text, { multiline: true, scrolling: true });
        box.preferredSize = [900, 560];
        var buttons = w.add('group');
        buttons.alignment = 'right';
        var cancelBtn = buttons.add('button', undefined, 'Cancel');
        var okBtn = buttons.add('button', undefined, 'Continue build');
        cancelBtn.onClick = function() { w.close(0); };
        okBtn.onClick = function() { w.close(1); };
        return w.show() === 1;
    } catch(e) {
        return confirm(text + '\r\rContinue build?');
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  I/O
// ─────────────────────────────────────────────────────────────────────────────
function readText(file) {
    file.encoding = 'UTF-8';
    if (!file.open('r')) throw new Error('Cannot open: ' + file.fsName);
    try { return file.read(); } finally { file.close(); }
}
function parseJson(text) {
    if (text.length && text.charCodeAt(0) === 0xFEFF) text = text.substring(1);
    if (typeof JSON !== 'undefined' && JSON.parse) return JSON.parse(text);
    return eval('(' + text + ')');
}
function resolveDataFile() {
    var selectedXlsx = File.openDialog('Выберите актуальную таблицу Excel', '*.xlsx;*.xls');
    if (!selectedXlsx || !selectedXlsx.exists)
        throw new Error('Выбор таблицы отменён. Сборка не запускалась.');

    var sf = scriptFolder();
    var converter = new File(sf.fsName + '/' + SETTINGS.XLSX_CONVERTER_RELATIVE);
    var output = new File(sf.fsName + '/' + SETTINGS.CURRENT_DATA_JSON_RELATIVE);
    if (!converter.exists) throw new Error('Не найден конвертер таблицы: ' + converter.fsName);

    // Do not allow an old conversion to be read when the current conversion fails.
    if (output.exists && !output.remove())
        throw new Error('Не удалось очистить временные данные: ' + output.fsName);

    var command = '"' + SETTINGS.NODE_BIN + '" "' + converter.fsName + '" "' +
                  selectedXlsx.fsName + '" "' + output.fsName + '"';
    var result = runTableConverter(command, sf);
    log('Selected table: ' + selectedXlsx.fsName);
    log('XLSX->JSON: ' + result);

    if (!output.exists)
        throw new Error('Не удалось создать данные из выбранной таблицы. ' + result);
    runStats.sourceTableFile = selectedXlsx.fsName;
    return output;
}


function callExternalCommand(command) {
    try {
        if (typeof system !== 'undefined' && system && typeof system.callSystem === 'function')
            return system.callSystem(command);
    } catch(e0) {}
    try {
        if (typeof app !== 'undefined' && app && typeof app.system === 'function')
            return app.system(command);
    } catch(e1) {}
    throw new Error('Illustrator не может запустить Node.js для чтения Excel.');
}

// Some Illustrator builds disable system.callSystem/app.system.  In that case
// launch a hidden WSH helper and wait for its status file before continuing.
function vbsEscape(value) { return String(value).replace(/"/g, '""'); }
function writePlainText(file, text) {
    file.encoding = 'UTF-8';
    if (!file.open('w')) throw new Error('Не удалось записать: ' + file.fsName);
    try { file.write(text); } finally { file.close(); }
}
function runTableConverter(command, scriptRoot) {
    try { return callExternalCommand(command); }
    catch(directError) {
        var buildPath = scriptRoot.fsName + '/build/';
        var runner = new File(buildPath + 'xlsx-convert-runner.vbs');
        var status = new File(buildPath + 'xlsx-convert-status.txt');
        var output = new File(buildPath + 'xlsx-convert-output.txt');
        try { if (status.exists) status.remove(); } catch(e0) {}
        try { if (output.exists) output.remove(); } catch(e1) {}

        var cmd = 'cmd.exe /d /s /c "' + command + ' > "' + output.fsName + '" 2>&1"';
        var vbs = 'Set shell = CreateObject("WScript.Shell")\r\n' +
                  'code = shell.Run("' + vbsEscape(cmd) + '", 0, True)\r\n' +
                  'Set fso = CreateObject("Scripting.FileSystemObject")\r\n' +
                  'Set f = fso.CreateTextFile("' + vbsEscape(status.fsName) + '", True, False)\r\n' +
                  'f.Write CStr(code)\r\n' +
                  'f.Close\r\n';
        writePlainText(runner, vbs);
        if (!runner.execute())
            throw new Error('Не удалось запустить резервный конвертер таблицы.');

        var until = nowMs() + 600000;
        while (!status.exists && nowMs() < until) safeSleep(200);
        if (!status.exists)
            throw new Error('Истекло время ожидания конвертации таблицы (10 минут).');
        var exitCode = trimString(readText(status));
        var details = output.exists ? readText(output) : '';
        if (exitCode !== '0')
            throw new Error('Конвертация таблицы завершилась с ошибкой ' + exitCode + ': ' + details);
        return details;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  DOCUMENT
// ─────────────────────────────────────────────────────────────────────────────
function setMM(doc) {
    try { app.preferences.setIntegerPreference('rulerType', RulerUnits.Millimeters); } catch(e)  {}
    try { doc.rulerUnits = RulerUnits.Millimeters; }  catch(e2) {}
    try { doc.rulerOrigin = [0, 0]; }                  catch(e3) {}
}
function activate(doc) {
    try { doc.activate(); } catch(e) { try { app.activeDocument = doc; } catch(e2) {} }
}
function unlockAll(doc) {
    function ul(layers) {
        for (var i = 0; i < layers.length; i++) {
            try { layers[i].locked  = false; } catch(e)  {}
            try { layers[i].visible = true;  } catch(e2) {}
            try { if (layers[i].layers) ul(layers[i].layers); } catch(e3) {}
        }
    }
    ul(doc.layers);
    for (var j = 0; j < doc.pageItems.length; j++) {
        try { doc.pageItems[j].locked = false; } catch(e4) {}
        try { doc.pageItems[j].hidden = false; } catch(e5) {}
    }
    try { doc.selection = null; } catch(e6) {}
}
function removeGuides(doc) {
    if (!SETTINGS.DELETE_GUIDES) return 0;
    var removed = 0;
    for (var i = doc.pageItems.length - 1; i >= 0; i--) {
        try {
            if (doc.pageItems[i].guides) {
                try { doc.pageItems[i].locked = false; } catch(e1) {}
                try { doc.pageItems[i].hidden = false; } catch(e2) {}
                doc.pageItems[i].remove();
                removed++;
            }
        } catch(e3) {}
    }
    return removed;
}
function prepareSourceDoc(doc) {
    unlockAll(doc);
    return removeGuides(doc);
}

// ─────────────────────────────────────────────────────────────────────────────
//  COLORS
// ─────────────────────────────────────────────────────────────────────────────
function makeCMYK(c, m, y, k) {
    var col = new CMYKColor();
    col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
    return col;
}
function black()     { return makeCMYK(0,   0,   0,   100); }
function darkBlue()  { return makeCMYK(100, 70,  15,  55);  }
function darkBrown() { return makeCMYK(25,  65,  100, 60);  }
function addressColor(addr) {
    return (addr.idx % 2 === 0) ? darkBlue() : darkBrown();
}

// ─────────────────────────────────────────────────────────────────────────────
//  STRING NORMALIZATION (FIXED)
// ─────────────────────────────────────────────────────────────────────────────

function isSpaceChar(ch) {
    var c = ch.charCodeAt(0);
    return ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n' ||
           c === 0x00A0 || c === 0x202F || c === 0x2007 ||
           c === 0x00AD || c === 0x200B || c === 0x200C ||
           c === 0x200D || c === 0x2060 || c === 0xFEFF;
}

function trimString(value) {
    var s = String(value == null ? '' : value);
    var a = 0;
    var b = s.length - 1;
    while (a <= b && isSpaceChar(s.charAt(a))) a++;
    while (b >= a && isSpaceChar(s.charAt(b))) b--;
    return a <= b ? s.substring(a, b + 1) : '';
}

function addNormalizedSpace(state) {
    if (!state.prevSpace && state.out.length) {
        state.out += ' ';
        state.prevSpace = true;
    }
}

function normalizeName(value) {
    var s = trimString(value).toLowerCase();
    var state = { out: '', prevSpace: false };
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        var code = ch.charCodeAt(0);
        if (isSpaceChar(ch)) {
            addNormalizedSpace(state);
            continue;
        }
        if (code === 0x0451 || code === 0x0401) ch = String.fromCharCode(0x0435);
        if (code === 0x00AB || code === 0x00BB ||
            code === 0x201C || code === 0x201D ||
            code === 0x201E || code === 0x201F ||
            code === 0x2018 || code === 0x2019) ch = '"';
        state.out += ch;
        state.prevSpace = false;
    }
    return trimString(state.out);
}

function isWordChar(ch) {
    var c = ch.charCodeAt(0);
    return (c >= 48 && c <= 57) ||
           (c >= 65 && c <= 90) ||
           (c >= 97 && c <= 122) ||
           (c >= 0x0400 && c <= 0x04FF);
}

function isQuoteChar(ch) {
    var c = ch.charCodeAt(0);
    return ch === '"' || ch === "'" || ch === '`' ||
           c === 0x00B4 || c === 0x2018 || c === 0x2019 ||
           c === 0x201C || c === 0x201D;
}

function looseNameKey(value) {
    var s = normalizeName(value);
    var state = { out: '', prevSpace: false };
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (isWordChar(ch)) {
            state.out += ch;
            state.prevSpace = false;
        } else if (!isQuoteChar(ch)) {
            addNormalizedSpace(state);
        }
    }
    return trimString(state.out);
}

function compactNameKey(value) {
    var s = looseNameKey(value);
    var out = '';
    for (var i = 0; i < s.length; i++) {
        if (s.charAt(i) !== ' ') out += s.charAt(i);
    }
    return out;
}

function nameWords(value) {
    var key = looseNameKey(value);
    var words = [];
    var cur = '';
    for (var i = 0; i < key.length; i++) {
        var ch = key.charAt(i);
        if (ch === ' ') {
            if (cur.length >= 2) words.push(cur);
            cur = '';
        } else {
            cur += ch;
        }
    }
    if (cur.length >= 2) words.push(cur);
    return words;
}

function isGenericMatchWord(word) {
    return word === '\u043f\u0438\u0440\u043e\u0436\u043d\u043e\u0435' || // пирожное
           word === '\u0442\u043e\u0440\u0442' ||                         // торт
           word === '\u043c\u0438\u043d\u0438' ||                         // мини
           word === '\u0432\u043a\u0443\u0441\u044b';                     // вкусы
}

function significantNameWords(value) {
    var src = nameWords(value);
    var result = [];
    for (var i = 0; i < src.length; i++) {
        if (!isGenericMatchWord(src[i])) result.push(src[i]);
    }
    return result.length ? result : src;
}

function compactWords(words) {
    var out = '';
    for (var i = 0; i < words.length; i++) out += words[i];
    return out;
}

function significantCompactKey(value) {
    return compactWords(significantNameWords(value));
}

function indexMapSet(map, key, entry) {
    if (!key) return;
    var k = '$' + key;
    if (!map.hasOwnProperty(k)) map[k] = entry;
}

function indexMapGet(map, key) {
    if (!key) return null;
    var k = '$' + key;
    return map.hasOwnProperty(k) ? map[k] : null;
}

function buildArtboardSearchIndex(doc) {
    var index = {
        entries:      [],
        byNorm:       {},
        byLoose:      {},
        byCompact:    {},
        bySigCompact: {}
    };
    for (var i = 0; i < doc.artboards.length; i++) {
        var n = doc.artboards[i].name;
        var words = nameWords(n);
        var sigWords = significantNameWords(n);
        var entry = {
            idx:        i,
            name:       n,
            norm:       normalizeName(n),
            loose:      looseNameKey(n),
            compact:    compactNameKey(n),
            words:      words,
            sigWords:   sigWords,
            sigCompact: compactWords(sigWords)
        };
        index.entries.push(entry);
        indexMapSet(index.byNorm, entry.norm, entry);
        indexMapSet(index.byLoose, entry.loose, entry);
        indexMapSet(index.byCompact, entry.compact, entry);
        indexMapSet(index.bySigCompact, entry.sigCompact, entry);
    }
    return index;
}

// ─────────────────────────────────────────────────────────────────────────────
//  FUZZY MATCHING (FIXED)
// ─────────────────────────────────────────────────────────────────────────────
function wordCoverageScore(targetWords, candidateWords) {
    if (!targetWords.length || !candidateWords.length) return 0;
    var common = 0;
    for (var i = 0; i < targetWords.length; i++) {
        for (var j = 0; j < candidateWords.length; j++) {
            if (targetWords[i] === candidateWords[j]) { common++; break; }
        }
    }
    return (common / targetWords.length) * 0.7 + (common / candidateWords.length) * 0.3;
}

function levenshteinDistance(a, b) {
    a = String(a || ''); b = String(b || '');
    if (a === b)      return 0;
    if (!a.length)    return b.length;
    if (!b.length)    return a.length;
    var row = [];
    for (var i = 0; i <= b.length; i++) row[i] = i;
    for (var x = 1; x <= a.length; x++) {
        var prev = row[0]; row[0] = x;
        for (var y = 1; y <= b.length; y++) {
            var old  = row[y];
            var cost = a.charAt(x-1) === b.charAt(y-1) ? 0 : 1;
            row[y]   = Math.min(row[y] + 1, row[y-1] + 1, prev + cost);
            prev = old;
        }
    }
    return row[b.length];
}

function similarityScore(a, b) {
    var maxLen = Math.max(a.length, b.length);
    return maxLen ? 1 - levenshteinDistance(a, b) / maxLen : 1;
}

// ─────────────────────────────────────────────────────────────────────────────
//  ARTBOARD SEARCH (indexed cascading matching with cache)
// ─────────────────────────────────────────────────────────────────────────────
findArtboardByName = function(doc, name, index, matchCache) {
    index = index || buildArtboardSearchIndex(doc);
    lastArtboardMatchInfo = null;
    
    var target        = normalizeName(name);
    var looseTarget   = looseNameKey(name);
    var compactTarget = compactNameKey(name);
    var targetWords   = nameWords(name);
    var sigWords      = significantNameWords(name);
    var sigCompact    = compactWords(sigWords);
    var targetHadGeneric = sigCompact !== compactTarget;
    var cacheKey      = '$' + compactTarget;
    var infoKey       = cacheKey + ':info';
    
    if (matchCache && matchCache.hasOwnProperty(cacheKey)) {
        lastArtboardMatchInfo = matchCache[infoKey] || {
            idx: matchCache[cacheKey], method: 'cache', score: 1, target: name,
            name: '', secondName: '', secondScore: 0
        };
        return matchCache[cacheKey];
    }
    function done(idx, info) {
        lastArtboardMatchInfo = info || {
            idx: idx, method: idx >= 0 ? 'unknown' : 'miss', score: idx >= 0 ? 1 : 0,
            target: name, name: '', secondName: '', secondScore: 0
        };
        if (matchCache) {
            matchCache[cacheKey] = idx;
            matchCache[infoKey] = lastArtboardMatchInfo;
        }
        return idx;
    }
    
    var hit = indexMapGet(index.byNorm, target);
    if (hit) return done(hit.idx, {
        idx: hit.idx, method: 'exact', score: 1, target: name,
        name: hit.name, secondName: '', secondScore: 0
    });
    
    hit = indexMapGet(index.byLoose, looseTarget);
    if (hit) return done(hit.idx, {
        idx: hit.idx, method: 'loose', score: 1, target: name,
        name: hit.name, secondName: '', secondScore: 0
    });
    
    hit = indexMapGet(index.byCompact, compactTarget);
    if (hit) return done(hit.idx, {
        idx: hit.idx, method: 'compact', score: 1, target: name,
        name: hit.name, secondName: '', secondScore: 0
    });
    
    hit = indexMapGet(index.bySigCompact, sigCompact);
    if (hit && sigCompact && sigCompact !== compactTarget) {
        log("GENERIC-MATCH: '" + name + "' -> '" + hit.name + "'");
        return done(hit.idx, {
            idx: hit.idx, method: 'generic', score: 0.94, target: name,
            name: hit.name, secondName: '', secondScore: 0
        });
    }
    
    var candidates = [];
    for (var n = 0; n < index.entries.length; n++) {
        var entry = index.entries[n];
        var entryHadGeneric = entry.sigCompact !== entry.compact;
        var score = 0;
        var method = '';
        
        if (compactTarget.length >= 5 && entry.compact.indexOf(compactTarget) === 0) {
            score = 0.95; method = 'prefix';
        } else if (entry.compact.length >= 5 && compactTarget.indexOf(entry.compact) >= 0 && entry.words.length >= 2) {
            score = 0.93; method = 'contains';
        } else if (compactTarget.length >= 5 && entry.compact.indexOf(compactTarget) >= 0 && targetWords.length >= 2) {
            score = 0.93; method = 'reverse-contains';
        } else if ((targetHadGeneric || entryHadGeneric) &&
                   sigCompact.length >= 5 && entry.sigCompact.length >= 3 &&
                   sigCompact.indexOf(entry.sigCompact) >= 0) {
            score = 0.94; method = 'generic-contains';
        } else if ((targetHadGeneric || entryHadGeneric) &&
                   sigCompact.length >= 3 && entry.sigCompact.length >= 5 &&
                   entry.sigCompact.indexOf(sigCompact) >= 0) {
            score = 0.94; method = 'generic-reverse-contains';
        } else {
            var ws1 = wordCoverageScore(targetWords, entry.words);
            var ws2 = wordCoverageScore(sigWords, entry.sigWords);
            var es1 = similarityScore(compactTarget, entry.compact);
            var es2 = similarityScore(sigCompact, entry.sigCompact);
            var ws = Math.max(ws1, ws2);
            var es = Math.max(es1, es2);
            score  = Math.max(ws, es);
            method = ws >= es ? 'words' : 'edit';
        }
        
        if (score > 0) candidates.push({ idx: entry.idx, score: score, method: method, name: entry.name });
    }
    
    candidates.sort(function(a, b) { return b.score - a.score; });
    
    if (candidates.length && candidates[0].score >= SETTINGS.FUZZY_THRESHOLD) {
        var best = candidates[0];
        var second = candidates.length > 1 ? candidates[1] : null;
        log("FUZZY [" + best.method + " " + Math.round(best.score*100) + "%]: '" + 
            name + "' -> '" + best.name + "'");
        return done(best.idx, {
            idx: best.idx, method: best.method, score: best.score, target: name,
            name: best.name,
            secondName: second ? second.name : '',
            secondScore: second ? second.score : 0
        });
    }
    
    if (sigWords.length >= 2) {
        for (var m = 0; m < index.entries.length; m++) {
            var cw = index.entries[m].sigWords;
            var matchCount = 0;
            for (var tw = 0; tw < sigWords.length; tw++) {
                if (sigWords[tw].length < 3) continue;
                for (var cwi = 0; cwi < cw.length; cwi++) {
                    if (cw[cwi].indexOf(sigWords[tw]) >= 0 || 
                        sigWords[tw].indexOf(cw[cwi]) >= 0) {
                        matchCount++;
                        break;
                    }
                }
            }
            var longWords = 0;
            for (var lw = 0; lw < sigWords.length; lw++) {
                if (sigWords[lw].length >= 3) longWords++;
            }
            if (longWords >= 2 && matchCount >= Math.max(2, Math.floor(longWords * 0.6))) {
                var wf = index.entries[m];
                log("WORD-FALLBACK [" + matchCount + "/" + longWords + "]: '" + 
                    name + "' -> '" + wf.name + "'");
                return done(wf.idx, {
                    idx: wf.idx, method: 'word-fallback', score: matchCount / longWords,
                    target: name, name: wf.name, secondName: '', secondScore: 0
                });
            }
        }
    }
    
    var topLog = candidates.length ? candidates.slice(0, 3) : [];
    var topStr = '';
    for (var ti = 0; ti < topLog.length; ti++) {
        topStr += "'" + topLog[ti].name + "' " + Math.round(topLog[ti].score*100) + "% ";
    }
    log("MISS: '" + name + "'" + (topStr ? " | best: " + topStr : " | no candidates"));
    
    return done(-1, {
        idx: -1, method: 'miss', score: topLog.length ? topLog[0].score : 0,
        target: name, name: topLog.length ? topLog[0].name : '',
        secondName: topLog.length > 1 ? topLog[1].name : '',
        secondScore: topLog.length > 1 ? topLog[1].score : 0
    });
};

// ─────────────────────────────────────────────────────────────────────────────
//  GEOMETRY
// ─────────────────────────────────────────────────────────────────────────────
function rectInfo(rect) {
    var ow = rect[2] - rect[0];
    var oh = rect[1] - rect[3];
    return { origW: ow, origH: oh, w: ow, h: oh };
}
function applyRotation(board) {
    if (board.origW > mm(SETTINGS.ROTATION_THRESHOLD_MM) + 0.01) {
        board.w        = board.origH;
        board.h        = board.origW;
        board.rotated90 = true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  LOAD DATA
// ─────────────────────────────────────────────────────────────────────────────
function loadTableData() {
    var f = resolveDataFile();
    var data = parseJson(readText(f));
    if (!data || !data.items || !data.items.length)
        throw new Error('No items in: ' + f.fsName);
    runStats.dataFile  = f.fsName;
    runStats.inputRows = data.items.length;
    log('Data: ' + f.fsName);
    log('Rows: ' + data.items.length);
    return data;
}

function dataAlreadyExpandedByQuantity(data) {
    return data && data.settings && data.settings.repeatByItog === true;
}

function parseCopyCount(value) {
    var text = trimString(value);
    if (!text) return 1;
    text = text.replace(',', '.');
    var number = parseFloat(text);
    if (isNaN(number)) return 1;
    if (number <= 0) return 0;
    number = Math.floor(number);
    return number < 1 ? 1 : number;
}

function copyCountForItem(data, item) {
    if (dataAlreadyExpandedByQuantity(data)) return 1;
    return parseCopyCount(item ? item.total : '');
}

function labelReserveWidth(label) {
    // Reserve sufficient space during layout; the rendered frame follows the
    // exact visible text bounds and is not constrained to this estimate.
    var estimatedMm = String(label || '').length * SETTINGS.LABEL_FONT_MM * 0.55 +
                      SETTINGS.LABEL_FRAME_PAD_MM * 2 +
                      SETTINGS.LABEL_BLOCK_INSET_MM * 2;
    return mm(Math.max(SETTINGS.LABEL_WIDTH_MM, estimatedMm));
}

// ─────────────────────────────────────────────────────────────────────────────
//  BUILD ADDRESSES (SCAN)
// ─────────────────────────────────────────────────────────────────────────────
function buildAddressesFromData(data) {
    var addrMap      = {};
    var addresses    = [];
    var sourceDocs   = {};
    var sourceDocList = [];
    
    function getAddress(label) {
        label = trimString(label) || 'output';
        if (!addrMap[label]) {
            var a = {
                idx:    addresses.length,
                label:  label,
                boards: [],
                stats:  { scanMs: 0, renderMs: 0, boardCount: 0 }
            };
            addrMap[label] = a;
            addresses.push(a);
        }
        return addrMap[label];
    }
    
    function getSourceEntry(path) {
        if (!sourceDocs[path]) {
            var file = new File(path);
            if (!file.exists) throw new Error('Missing source file: ' + path);
            var doc = app.open(file);
            setMM(doc);
            sourceDocs[path]  = {
                doc: doc,
                file: file,
                index: buildArtboardSearchIndex(doc),
                matchCache: {}
            };
            sourceDocList.push(sourceDocs[path]);
        }
        return sourceDocs[path];
    }
    
    var startMs = nowMs();
    
    try {
        var alreadyExpanded = dataAlreadyExpandedByQuantity(data);
        for (var i = 0; i < data.items.length; i++) {
            var item = data.items[i];
            var copyCount = copyCountForItem(data, item);
            if (copyCount < 1) continue;
            
            for (var rc = 0; rc < copyCount; rc++) recordTableRow(item);
            runStats.expectedCopies += copyCount;
            
            if (!item || !item.sourceFile) {
                var sourceHint = item && item.sourceSuggestion ?
                    ' Suggested source: ' + item.sourceSuggestion +
                    ' (' + scorePct(item.sourceSuggestionScore) + '% ' +
                    (item.sourceSuggestionMethod || '') + ')' : '';
                for (var ms = 0; ms < copyCount; ms++)
                    recordIssue('missing-source', item, 'Row skipped: no sourceFile.' + sourceHint);
                warn('No sourceFile: row ' + (item ? item.rowNumber : '?') +
                     (copyCount > 1 ? ' x' + copyCount : ''));
                continue;
            }
            
            try {
                var addr     = getAddress(item.address || 'output');
                var srcEntry = getSourceEntry(item.sourceFile);
                var srcDoc   = srcEntry.doc;
                
                var abIndex = findArtboardByName(srcDoc, item.product, srcEntry.index, srcEntry.matchCache);
                recordMatchReview(item, lastArtboardMatchInfo);
                
                if (abIndex < 0) {
                    for (var ma = 0; ma < copyCount; ma++)
                        recordIssue('missing-artboard', item, 'Artboard not found');
                    warn('NOT FOUND: row ' + item.rowNumber + 
                         ' | file: ' + (item.sourceFileName || item.sourceFile) + 
                         ' | product: ' + item.product +
                         (copyCount > 1 ? ' x' + copyCount : ''));
                    continue;
                }
                
                var rect = srcDoc.artboards[abIndex].artboardRect;
                var ri   = rectInfo(rect);
                for (var copy = 1; copy <= copyCount; copy++) {
                    var board = {
                        idx:            abIndex,
                        rect:           [rect[0], rect[1], rect[2], rect[3]],
                        origW:          ri.origW,
                        origH:          ri.origH,
                        w:              ri.w,
                        h:              ri.h,
                        rotated90:      false,
                        file:           new File(item.sourceFile),
                        fileKey:        item.sourceFile,
                        product:        item.product,
                        rowNumber:      item.rowNumber,
                        copyIndex:      alreadyExpanded ? (item.copyIndex || 1) : copy,
                        copyCount:      copyCount,
                        sourceFileName: item.sourceFileName || decodedName(new File(item.sourceFile))
                    };
                    applyRotation(board);
                    addr.boards.push(board);
                    addr.stats.boardCount = addr.boards.length;
                    recordMatchedRow(addr.label);
                }
                
            } catch(rowError) {
                // Retry once after GC and redraw — PARM errors can be transient.
                var retryOk = false;
                try {
                    safeGC(); safeRedraw(); safeSleep(500);
                    var addr2     = getAddress(item.address || 'output');
                    var srcEntry2 = getSourceEntry(item.sourceFile);
                    var srcDoc2   = srcEntry2.doc;
                    var abIndex2  = findArtboardByName(srcDoc2, item.product, srcEntry2.index, srcEntry2.matchCache);
                    if (abIndex2 >= 0) {
                        var rect2 = srcDoc2.artboards[abIndex2].artboardRect;
                        var ri2   = rectInfo(rect2);
                        for (var rc2 = 0; rc2 < copyCount; rc2++) {
                            var board2 = {
                                idx: abIndex2,
                                rect: [rect2[0], rect2[1], rect2[2], rect2[3]],
                                origW: ri2.origW, origH: ri2.origH,
                                w: ri2.w, h: ri2.h, rotated90: false,
                                file: new File(item.sourceFile),
                                fileKey: item.sourceFile,
                                product: item.product,
                                rowNumber: item.rowNumber,
                                copyIndex: alreadyExpanded ? (item.copyIndex || 1) : (rc2 + 1),
                                copyCount: copyCount,
                                sourceFileName: item.sourceFileName || decodedName(new File(item.sourceFile))
                            };
                            applyRotation(board2);
                            addr2.boards.push(board2);
                            addr2.stats.boardCount = addr2.boards.length;
                            recordMatchedRow(addr2.label);
                        }
                        recordMatchReview(item, lastArtboardMatchInfo);
                        retryOk = true;
                        log('RETRY OK: row ' + item.rowNumber + ' | ' + item.product);
                    }
                } catch(retryError) {
                    log('RETRY FAILED: row ' + (item ? item.rowNumber : '?') + ' | ' +
                        (retryError && retryError.message ? retryError.message : retryError));
                }
                if (!retryOk) {
                    for (var se = 0; se < copyCount; se++)
                        recordIssue('scan-error', item, 
                            rowError && rowError.message ? rowError.message : String(rowError));
                    warn('SCAN ERROR: row ' + (item ? item.rowNumber : '?') + ' | ' +
                         (item ? item.product : '') + 
                         (copyCount > 1 ? ' x' + copyCount : '') + ' | ' +
                         (rowError && rowError.message ? rowError.message : rowError));
                }
            }
        }
    } finally {
        for (var c = 0; c < sourceDocList.length; c++) {
            try { sourceDocList[c].doc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
        }
        safeGC();
    }
    
    var elapsed = nowMs() - startMs;
    for (var a = 0; a < addresses.length; a++) {
        addresses[a].stats.scanMs = elapsed;
    }
    log('Expected copies after quantity: ' + runStats.expectedCopies);
    return addresses;
}

// ─────────────────────────────────────────────────────────────────────────────
//  LAYOUT
// ─────────────────────────────────────────────────────────────────────────────
function buildLayout(addresses) {
    var mg     = mm(SETTINGS.SHEET_MARGIN_MM);
    var sheetW = mm(SETTINGS.SHEET_WIDTH_MM)  - mg * 2;
    var sheetH = mm(SETTINGS.SHEET_HEIGHT_MM) - mg * 2;
    var labelH = mm(SETTINGS.LABEL_HEIGHT_MM);
    var abGap  = mm(SETTINGS.ARTBOARD_GAP_MM);
    var colGap = mm(SETTINGS.COLUMN_GAP_MM);
    
    // Check: no artboard should be larger than sheet
    for (var a0 = 0; a0 < addresses.length; a0++) {
        for (var b0 = 0; b0 < addresses[a0].boards.length; b0++) {
            var brd = addresses[a0].boards[b0];
            var neededW = Math.max(brd.w, labelReserveWidth(addresses[a0].label));
            if (neededW > sheetW + 0.01)
                throw new Error('Board wider than sheet: ' + addresses[a0].label + ' row ' + brd.rowNumber +
                    ' (' + r1(neededW) + ' > ' + r1(sheetW) + ' mm)');
            if (brd.h + labelH > sheetH + 0.01)
                throw new Error('Board taller than sheet: ' + addresses[a0].label + ' row ' + brd.rowNumber +
                    ' (' + r1(brd.h) + '+' + r1(labelH) + ' > ' + r1(sheetH) + ' mm)');
        }
    }
    
    var sheets  = [];
    var addrMap = {};
    
    function newSheet() {
        var s = { cells: [], addrInfos: [], colX: mg, colY: mg, colW: 0 };
        sheets.push(s);
        return s;
    }
    function getOrCreateAddrInfo(sheet, addr) {
        if (!addrMap[addr.label]) {
            var ai = {
                addr:     addr,
                cells:    [],
                outlineL: 1e12, outlineT:  1e12,
                outlineR: -1e12, outlineB: -1e12,
                columns:  {}
            };
            sheet.addrInfos.push(ai);
            addrMap[addr.label] = ai;
        }
        return addrMap[addr.label];
    }
    function flushColumn(sheet, neededW) {
        var nextX = sheet.colX + sheet.colW + colGap;
        if (nextX + neededW > mg + sheetW + 0.01) return false;
        sheet.colX = nextX;
        sheet.colY = mg;
        sheet.colW = 0;
        return true;
    }
    var curSheet = newSheet();
    function startNewSheet() {
        curSheet = newSheet();
        addrMap  = {};
    }
    
    for (var a = 0; a < addresses.length; a++) {
        var addr = addresses[a];
        if (!addr.boards.length) continue;
        var addrLabelW = labelReserveWidth(addr.label);
        
        for (var b = 0; b < addr.boards.length; b++) {
            var board     = addr.boards[b];
            var needLabel = !addrMap[addr.label];
            var topPad    = needLabel ? labelH : 0;
            var layoutW   = needLabel ? Math.max(board.w, addrLabelW) : board.w;
            var cellH     = topPad + board.h;
            var fitsV     = (curSheet.colY + cellH <= mg + sheetH + 0.01);
            
            if (!fitsV) {
                var moved = flushColumn(curSheet, layoutW);
                if (moved) {
                    needLabel = !addrMap[addr.label];
                    topPad    = needLabel ? labelH : 0;
                    layoutW   = needLabel ? Math.max(board.w, addrLabelW) : board.w;
                    cellH     = topPad + board.h;
                    fitsV     = (curSheet.colY + cellH <= mg + sheetH + 0.01);
                }
                if (!fitsV || !moved) {
                    startNewSheet();
                    needLabel = true;
                    topPad    = labelH;
                    layoutW   = Math.max(board.w, addrLabelW);
                    cellH     = topPad + board.h;
                }
            }
            var colRightAfter = curSheet.colX + Math.max(curSheet.colW, layoutW);
            if (colRightAfter > mg + sheetW + 0.01) {
                startNewSheet();
                needLabel = true;
                topPad    = labelH;
                layoutW   = Math.max(board.w, addrLabelW);
                cellH     = topPad + board.h;
            }
            
            var cellX  = curSheet.colX;
            var cellY  = curSheet.colY + topPad;
            var labelY = needLabel ? curSheet.colY : -1;
            var cell   = {
                addr: addr, board: board,
                sheetX: cellX, sheetY: cellY, labelY: labelY,
                w: board.w, h: board.h,
                layoutW: layoutW,
                labelW: needLabel ? addrLabelW : 0
            };
            curSheet.cells.push(cell);
            curSheet.colY += cellH + abGap;
            if (layoutW > curSheet.colW) curSheet.colW = layoutW;
            
            var ai      = getOrCreateAddrInfo(curSheet, addr);
            ai.cells.push(cell);
            var topEdge = (labelY >= 0) ? labelY : cellY;
            if (cellX           < ai.outlineL) ai.outlineL = cellX;
            if (topEdge         < ai.outlineT) ai.outlineT = topEdge;
            if (cellX + layoutW > ai.outlineR) ai.outlineR = cellX + layoutW;
            if (cellY + board.h > ai.outlineB) ai.outlineB = cellY + board.h;
            
            var colKey = String(cellX);
            if (!ai.columns[colKey]) {
                ai.columns[colKey] = {
                    left: cellX, right: cellX + layoutW,
                    top: topEdge, bottom: cellY + board.h
                };
            } else {
                var col = ai.columns[colKey];
                if (cellX + layoutW > col.right)  col.right  = cellX + layoutW;
                if (topEdge         < col.top)     col.top    = topEdge;
                if (cellY + board.h > col.bottom)  col.bottom = cellY + board.h;
            }
        }
    }
    return sheets;
}

// ─────────────────────────────────────────────────────────────────────────────
//  COPY / PASTE
// ─────────────────────────────────────────────────────────────────────────────
function tryCopyBoard(srcDoc, board, mName) {
    srcDoc.selection = null;
    srcDoc.artboards.setActiveArtboardIndex(board.idx);
    var mk = srcDoc.pathItems.rectangle(
        board.rect[1], board.rect[0], board.origW, board.origH);
    mk.name    = mName;
    mk.filled  = false;
    mk.stroked = false;
    mk.selected = true;
    try { srcDoc.selectObjectsOnActiveArtboard(); } catch(e) {}
    mk.selected = true;
    if (!srcDoc.selection || srcDoc.selection.length <= 1) {
        try { mk.remove(); } catch(e2) {}
        srcDoc.selection = null;
        return false;
    }
    var operationalNames = tagOperationalPaths(srcDoc, board);
    app.executeMenuCommand('copy');
    restoreOperationalPathNames(operationalNames);
    try { mk.remove(); } catch(e3) {}
    srcDoc.selection = null;
    return true;
}

function intersectsBoard(item, rect) {
    try { var b = item.geometricBounds; return b[2] > rect[0] && b[0] < rect[2] && b[1] > rect[3] && b[3] < rect[1]; } catch(e) { return false; }
}
function sourceLayerKind(item) {
    var name = '';
    try { name = String(item.layer.name || '').toLowerCase(); } catch(e) {}
    if (name.indexOf('magnet') >= 0) return 'MAGNET';
    if (name.indexOf('cut') >= 0) return 'CUT';
    return '';
}
function isCircleMarkerPath(item) {
    try {
        if (!item.closed || item.filled || item.pathPoints.length !== 4) return false;
        var b = item.geometricBounds, w = Math.abs(b[2] - b[0]), h = Math.abs(b[1] - b[3]);
        return w > 0 && h > 0 && Math.abs(w - h) <= Math.max(w, h) * 0.08;
    } catch(e) { return false; }
}
function tagOperationalPaths(srcDoc, board) {
    var changed = [];
    function scan(container) {
        if (!container) return;
        var i, item, kind;
        try {
            for (i = 0; i < container.pathItems.length; i++) {
                item = container.pathItems[i];
                kind = sourceLayerKind(item);
                if (!kind && isCircleMarkerPath(item)) kind = 'MAGNET';
                if (!kind || !intersectsBoard(item, board.rect)) continue;
                try { changed.push({ item: item, name: item.name }); item.name = OPERATIONAL_PFX + kind + '__' + i; } catch(e) {}
            }
        } catch(e) {}
        try {
            for (i = 0; i < container.groupItems.length; i++) {
                scan(container.groupItems[i]);
            }
        } catch(e) {}
        try {
            for (i = 0; i < container.compoundPathItems.length; i++) {
                scan(container.compoundPathItems[i]);
            }
        } catch(e) {}
    }
    scan(srcDoc);
    return changed;
}
function restoreOperationalPathNames(changed) {
    for (var i = 0; i < changed.length; i++) try { changed[i].item.name = changed[i].name; } catch(e) {}
}


function copyBoard(srcDoc, board, mName, docState, label) {
    if (tryCopyBoard(srcDoc, board, mName)) return true;
    if (!SETTINGS.CLEAN_ON_COPY_ERROR || docState.cleaned) return false;
    docState.cleaned = true;
    var guideCount   = prepareSourceDoc(srcDoc);
    log('Retry copy' + (guideCount ? ' after removing ' + guideCount + ' guides' : '') + ': ' + label);
    return tryCopyBoard(srcDoc, board, mName);
}

function findByName(sel, name) {
    for (var i = 0; i < sel.length; i++) {
        try { if (sel[i].name === name) return sel[i]; } catch(e) {}
    }
    return null;
}
function findByNameDeepItem(item, name) {
    try { if (item.name === name) return item; } catch(e0) {}
    var collections = ['pageItems', 'groupItems', 'pathItems', 'compoundPathItems'];
    for (var c = 0; c < collections.length; c++) {
        var list = null;
        try { list = item[collections[c]]; } catch(e1) { list = null; }
        if (!list) continue;
        for (var i = 0; i < list.length; i++) {
            var hit = findByNameDeepItem(list[i], name);
            if (hit) return hit;
        }
    }
    return null;
}
function findByNameDeep(sel, name) {
    var hit = findByName(sel, name);
    if (hit) return hit;
    for (var i = 0; i < sel.length; i++) {
        hit = findByNameDeepItem(sel[i], name);
        if (hit) return hit;
    }
    return null;
}
function toArr(sel) {
    var a = [];
    for (var i = 0; i < sel.length; i++) a.push(sel[i]);
    return a;
}

function pasteIntoDestDoc(destDoc) {
    activate(destDoc);
    try { destDoc.activeLayer = destDoc.layers[0]; } catch(e0) {}
    try { app.executeMenuCommand('deselectall'); } catch(e1) {}
    destDoc.selection = null;
    try { app.executeMenuCommand('pasteInPlace'); } catch(e2) {}
    if (destDoc.selection && destDoc.selection.length) return true;
    try { app.executeMenuCommand('paste'); } catch(e3) {}
    return !!(destDoc.selection && destDoc.selection.length);
}

// ─────────────────────────────────────────────────────────────────────────────
//  RENDER CELL
// ─────────────────────────────────────────────────────────────────────────────
function rotateGroup90(grp, destL, destT) {
    try {
        grp.rotate(-90, true, true, true, true, Transformation.CENTER);
        var gb = grp.geometricBounds;
        grp.translate(destL - gb[0], destT - gb[1]);
    } catch(e) { warn('Rotate failed: ' + e.message); }
}

function renderCell(srcDoc, destDoc, cell, parentGroup, sheetDocTop, docState) {
    var board = cell.board;
    var addr  = cell.addr;
    var destL = cell.sheetX;
    var destT = sheetDocTop - cell.sheetY;
    var mName = '';
    var pasted = null;
    var mk = null;
    var mkB = null;
    var maxAttempts = Math.max(1, SETTINGS.PASTE_RETRY_COUNT || 1);
    var pasteHadRetry = false;
    
    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
        mName = MARKER_PFX + board.rowNumber + '_' + board.idx + '_' + nowMs() + '_' + attempt;
        activate(srcDoc);
        var copied = copyBoard(srcDoc, board, mName, docState, addr.label);
        if (!copied) {
            if (attempt === maxAttempts) {
                warn('Empty board (nothing to copy): ' + addr.label + 
                     ' row ' + board.rowNumber + ' | ' + board.product);
                return false;
            }
        } else if (pasteIntoDestDoc(destDoc)) {
            pasted = toArr(destDoc.selection);
            mk     = findByNameDeep(pasted, mName);
            mkB    = mk ? mk.geometricBounds 
                        : [board.rect[0], board.rect[1], board.rect[2], board.rect[3]];
            break;
        } else {
            if (!pasteHadRetry) {
                runStats.pasteRetryCells++;
                pasteHadRetry = true;
            }
            runStats.pasteRetryAttempts++;
            log('Paste retry ' + attempt + '/' + maxAttempts + ': ' + 
                addr.label + ' row ' + board.rowNumber);
        }
        
        if (attempt < maxAttempts) {
            safeGC();
            safeRedraw();
            safeSleep((SETTINGS.PASTE_RETRY_DELAY_MS || 0) * attempt);
        }
    }
    
    if (!pasted || !pasted.length) {
        runStats.pasteRetryFailures++;
        warn('Paste failed after retries: ' + addr.label + ' row ' + board.rowNumber);
        return false;
    }
    if (!mk) warn('Marker missing, using artboard rect as fallback: ' + 
                  addr.label + ' row ' + board.rowNumber);
    
    var artGrp = destDoc.groupItems.add();
    artGrp.name = addr.label + '/row' + board.rowNumber + '/' + board.product;
    
    for (var i = 0; i < pasted.length; i++) {
        if (pasted[i] === mk) continue;
        try { pasted[i].move(artGrp, ElementPlacement.PLACEATEND); }
        catch(e2) { warn('Move failed: ' + addr.label + ' row ' + board.rowNumber); }
    }
    if (mk) try { mk.remove(); } catch(e3) {}
    artGrp.translate(destL - mkB[0], destT - mkB[1]);
    if (board.rotated90) rotateGroup90(artGrp, destL, destT);
    
    if (SETTINGS.CLIP_TO_ARTBOARD) {
        var cg = destDoc.groupItems.add();
        cg.name = addr.label + '/clip/row' + board.rowNumber;
        var cr  = destDoc.pathItems.rectangle(destT, destL, board.w, board.h);
        cr.filled   = false;
        cr.stroked  = false;
        cr.clipping = true;
        cr.move(cg, ElementPlacement.PLACEATBEGINNING);
        artGrp.move(cg, ElementPlacement.PLACEATEND);
        cg.clipped = true;
        cg.move(parentGroup, ElementPlacement.PLACEATEND);
    } else {
        artGrp.move(parentGroup, ElementPlacement.PLACEATEND);
    }
    destDoc.selection = null;
    return true;
}

// ─────────────────────────────────────────────────────────────────────────────
//  LABELS AND OUTLINES
// ─────────────────────────────────────────────────────────────────────────────
function artCacheKey(board) {
    return '$' + board.fileKey + '#' + board.idx;
}

function getSheetArtCacheGroup(ctx) {
    if (!ctx.artCacheGroup) {
        ctx.artCacheGroup = ctx.doc.groupItems.add();
        ctx.artCacheGroup.name = '__BEP_ART_CACHE__';
    }
    return ctx.artCacheGroup;
}

function createCachedArtMaster(ctx, srcDoc, board, docState, label) {
    var destDoc = ctx.doc;
    var mName = '';
    var pasted = null;
    var mk = null;
    var mkB = null;
    var maxAttempts = Math.max(1, SETTINGS.PASTE_RETRY_COUNT || 1);
    var pasteHadRetry = false;
    
    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
        mName = MARKER_PFX + 'CACHE_' + board.rowNumber + '_' + board.idx + '_' + nowMs() + '_' + attempt;
        activate(srcDoc);
        var copied = copyBoard(srcDoc, board, mName, docState, label);
        if (!copied) {
            if (attempt === maxAttempts) {
                warn('Empty board (nothing to cache): ' + label +
                     ' row ' + board.rowNumber + ' | ' + board.product);
                return null;
            }
        } else if (pasteIntoDestDoc(destDoc)) {
            pasted = toArr(destDoc.selection);
            mk     = findByNameDeep(pasted, mName);
            mkB    = mk ? mk.geometricBounds
                        : [board.rect[0], board.rect[1], board.rect[2], board.rect[3]];
            break;
        } else {
            if (!pasteHadRetry) {
                runStats.pasteRetryCells++;
                pasteHadRetry = true;
            }
            runStats.pasteRetryAttempts++;
            log('Paste retry ' + attempt + '/' + maxAttempts + ' while caching: ' +
                label + ' row ' + board.rowNumber);
        }
        
        if (attempt < maxAttempts) {
            safeGC();
            safeRedraw();
            safeSleep((SETTINGS.PASTE_RETRY_DELAY_MS || 0) * attempt);
        }
    }
    
    if (!pasted || !pasted.length) {
        runStats.pasteRetryFailures++;
        warn('Paste failed after retries while caching: ' + label + ' row ' + board.rowNumber);
        return null;
    }
    if (!mk) warn('Marker missing, using artboard rect as fallback: ' +
                  label + ' row ' + board.rowNumber);
    
    var cacheGroup = getSheetArtCacheGroup(ctx);
    var master = destDoc.groupItems.add();
    master.name = '__BEP_CACHE_ART__' + board.idx + '/' + board.product;
    master.move(cacheGroup, ElementPlacement.PLACEATEND);
    
    for (var i = 0; i < pasted.length; i++) {
        if (pasted[i] === mk) continue;
        try { pasted[i].move(master, ElementPlacement.PLACEATEND); }
        catch(e2) { warn('Cache move failed: ' + label + ' row ' + board.rowNumber); }
    }
    if (mk) try { mk.remove(); } catch(e3) {}
    
    try {
        if (!master.pageItems.length) {
            master.remove();
            warn('Cached board is empty: ' + label + ' row ' + board.rowNumber);
            return null;
        }
    } catch(e4) {}
    
    return { group: master, anchor: [mkB[0], mkB[1], mkB[2], mkB[3]] };
}

function getCachedArtEntry(ctx, cell, cache) {
    if (!cell || !cell.board) return null;
    var board = cell.board;
    var key = artCacheKey(board);
    if (ctx.artCache.hasOwnProperty(key)) {
        runStats.artCacheHits++;
        return ctx.artCache[key];
    }
    
    var src = getRenderSourceDoc(cache, board);
    var entry = createCachedArtMaster(ctx, src.doc, board, src.state, cell.addr.label);
    if (entry) {
        ctx.artCache[key] = entry;
        runStats.artCacheCopies++;
    }
    return entry;
}

function placeCachedArtGroup(ctx, entry, cell, parentGroup) {
    var destDoc = ctx.doc;
    var board = cell.board;
    var destL = cell.sheetX;
    var destT = ctx.sheetDocTop - cell.sheetY;
    var artGrp = null;
    var clipGrp = null;
    
    try {
        artGrp = entry.group.duplicate(parentGroup, ElementPlacement.PLACEATEND);
        artGrp.name = cell.addr.label + '/row' + board.rowNumber + '/' + board.product;
        artGrp.translate(destL - entry.anchor[0], destT - entry.anchor[1]);
        if (board.rotated90) rotateGroup90(artGrp, destL, destT);
        
        if (SETTINGS.CLIP_TO_ARTBOARD) {
            clipGrp = destDoc.groupItems.add();
            clipGrp.name = cell.addr.label + '/clip/row' + board.rowNumber;
            var cr = destDoc.pathItems.rectangle(destT, destL, board.w, board.h);
            cr.filled = false;
            cr.stroked = false;
            cr.clipping = true;
            cr.move(clipGrp, ElementPlacement.PLACEATBEGINNING);
            artGrp.move(clipGrp, ElementPlacement.PLACEATEND);
            clipGrp.clipped = true;
            clipGrp.move(parentGroup, ElementPlacement.PLACEATEND);
        }
        
        destDoc.selection = null;
        return true;
    } catch(e) {
        runStats.artCacheDupFails++;
        try { if (clipGrp) clipGrp.remove(); } catch(e1) {}
        try { if (artGrp) artGrp.remove(); } catch(e2) {}
        log('Art cache duplicate failed: ' + cell.addr.label +
            ' row ' + board.rowNumber + ' | ' + (e && e.message ? e.message : e));
        return false;
    }
}

function renderCellFromArtCache(ctx, cell, cache, parentGroup) {
    if (!SETTINGS.RENDER_ART_CACHE || !ctx || !cell || !cache) return false;
    var entry = getCachedArtEntry(ctx, cell, cache);
    if (!entry) return false;
    if (!placeCachedArtGroup(ctx, entry, cell, parentGroup)) return false;
    runStats.artCachePlacements++;
    return true;
}

function clearSheetArtCache(ctx) {
    if (!ctx) return;
    try {
        if (ctx.artCacheGroup) ctx.artCacheGroup.remove();
    } catch(e) {}
    ctx.artCacheGroup = null;
    ctx.artCache = {};
}

var cachedLabelFont = null;
var cachedLabelFontChecked = false;

function getLabelFont() {
    if (cachedLabelFontChecked) return cachedLabelFont;
    cachedLabelFontChecked = true;
    for (var i = 0; i < SETTINGS.LABEL_FONT_NAMES.length; i++) {
        try {
            cachedLabelFont = app.textFonts.getByName(SETTINGS.LABEL_FONT_NAMES[i]);
            if (cachedLabelFont) return cachedLabelFont;
        } catch(e) {}
    }
    return null;
}

function drawLabel(destDoc, group, addr, cell, sheetDocTop) {
    if (!cell || cell.labelY < 0) return;
    if (!addr || !addr.label) return;
    var framePad = mm(SETTINGS.LABEL_FRAME_PAD_MM);
    var blockInset = mm(SETTINGS.LABEL_BLOCK_INSET_MM);
    var labelT = sheetDocTop - cell.labelY;
    var tf = destDoc.textFrames.add();
    tf.name     = 'Label:' + addr.label;
    tf.contents = addr.label;
    try {
        var ca = tf.textRange.characterAttributes;
        var font = getLabelFont();
        if (font) ca.textFont = font;
        ca.size      = mm(SETTINGS.LABEL_FONT_MM);
        ca.leading   = mm(SETTINGS.LABEL_FONT_MM * 1.05);
        ca.fillColor = addressColor(addr);
    } catch(e) {}
    try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch(e2) {}

    // Point text lets the frame follow the actual glyph bounds.
    tf.position = [cell.sheetX + blockInset + framePad, labelT - blockInset - framePad];
    var bounds = tf.visibleBounds;
    if (!bounds) { try { tf.remove(); } catch(e4) {} return; }
    tf.translate(cell.sheetX + blockInset + framePad - bounds[0],
                 labelT - blockInset - framePad - bounds[1]);
    bounds = tf.visibleBounds;
    var frame = destDoc.pathItems.rectangle(
        bounds[1] + framePad, bounds[0] - framePad,
        (bounds[2] - bounds[0]) + framePad * 2,
        (bounds[1] - bounds[3]) + framePad * 2);
    frame.name = 'Label frame:' + addr.label;
    frame.filled = false;
    frame.stroked = true;
    frame.strokeWidth = mm(0.5);
    frame.strokeColor = addressColor(addr);
    frame.move(group, ElementPlacement.PLACEATBEGINNING);
    tf.move(group, ElementPlacement.PLACEATEND);
}

function addressColumns(ai) {
    var cols = [];
    for (var key in ai.columns) {
        if (ai.columns.hasOwnProperty(key)) cols.push(ai.columns[key]);
    }
    cols.sort(function(a, b) { return a.left - b.left; });
    for (var i = 0; i < cols.length; i++) {
        cols[i].shapeRight = cols[i].right;
        if (i < cols.length - 1) cols[i].shapeRight = cols[i+1].left;
    }
    return cols;
}

function steppedOutlinePoints(cols) {
    var pts = [];
    if (!cols.length) return pts;
    pts.push([cols[0].left, cols[0].top]);
    pts.push([cols[0].shapeRight, cols[0].top]);
    for (var i = 1; i < cols.length; i++) {
        var x = cols[i].left;
        if (pts[pts.length-1][0] !== x || pts[pts.length-1][1] !== cols[i-1].top)
            pts.push([x, cols[i-1].top]);
        if (cols[i].top !== cols[i-1].top) pts.push([x, cols[i].top]);
        pts.push([cols[i].shapeRight, cols[i].top]);
    }
    pts.push([cols[cols.length-1].shapeRight, cols[cols.length-1].bottom]);
    pts.push([cols[cols.length-1].left,       cols[cols.length-1].bottom]);
    for (i = cols.length - 2; i >= 0; i--) {
        x = cols[i+1].left;
        if (pts[pts.length-1][0] !== x || pts[pts.length-1][1] !== cols[i+1].bottom)
            pts.push([x, cols[i+1].bottom]);
        if (cols[i].bottom !== cols[i+1].bottom) pts.push([x, cols[i].bottom]);
        pts.push([cols[i].left, cols[i].bottom]);
    }
    return pts;
}

function drawAddressOutline(destDoc, group, ai, sheetDocTop) {
    if (!SETTINGS.DRAW_ADDRESS_OUTLINE) return;
    var cols = addressColumns(ai);
    var pts  = steppedOutlinePoints(cols);
    if (pts.length < 4) return;
    var aiPts = [];
    for (var i = 0; i < pts.length; i++)
        aiPts.push([pts[i][0], sheetDocTop - pts[i][1]]);
    var p = destDoc.pathItems.add();
    p.name        = 'Outline:' + ai.addr.label;
    p.setEntirePath(aiPts);
    p.closed      = true;
    p.filled      = false;
    p.stroked     = true;
    p.strokeWidth = mm(SETTINGS.OUTLINE_STROKE_MM);
    p.strokeColor = addressColor(ai.addr);
    p.move(group, ElementPlacement.PLACEATBEGINNING);
}

// ─────────────────────────────────────────────────────────────────────────────
//  SHEET RENDER CONTEXT
// ─────────────────────────────────────────────────────────────────────────────
function makeTechnicalLayers(doc) {
    var artwork = doc.layers[0];
    artwork.name = 'artwork';
    var cut = doc.layers.add(); cut.name = 'cut';
    var magnet = doc.layers.add(); magnet.name = 'magnet';
    doc.activeLayer = cut;
    var cutPaths = doc.groupItems.add(); cutPaths.name = 'Cut paths';
    var controls = doc.groupItems.add(); controls.name = 'Address controls';
    doc.activeLayer = magnet;
    var magnets = doc.groupItems.add(); magnets.name = 'Magnets';
    doc.activeLayer = artwork;
    return { artwork: artwork, cutPaths: cutPaths, controls: controls, magnets: magnets };
}
function setTechnicalStroke(item, color, widthPt) {
    try { item.filled = false; item.stroked = true; item.strokeColor = color; item.strokeWidth = widthPt; } catch(e) {}
}
function routeTaggedOperationalPaths(tech, container) {
    function scan(parent) {
        if (!parent) return;
        var i, p, name;
        try {
            for (i = parent.pathItems.length - 1; i >= 0; i--) {
                p = parent.pathItems[i]; name = '';
                try { name = String(p.name || ''); } catch(e) {}
                if (name.indexOf(OPERATIONAL_PFX + 'MAGNET__') === 0) {
                    setTechnicalStroke(p, black(), 0.2835);
                    p.move(tech.magnets, ElementPlacement.PLACEATEND);
                } else if (name.indexOf(OPERATIONAL_PFX + 'CUT__') === 0) {
                    setTechnicalStroke(p, makeCMYK(100, 0, 0, 0), 0.284);
                    p.move(tech.cutPaths, ElementPlacement.PLACEATEND);
                }
            }
        } catch(e) {}
        try {
            for (i = parent.groupItems.length - 1; i >= 0; i--) {
                scan(parent.groupItems[i]);
            }
        } catch(e) {}
        try {
            for (i = parent.compoundPathItems.length - 1; i >= 0; i--) {
                scan(parent.compoundPathItems[i]);
            }
        } catch(e) {}
    }
    scan(container);
}
function copyAddressControls(ctx, tech) {
    for (var k = 0; k < ctx.sheetData.addrInfos.length; k++) {
        var ai = ctx.sheetData.addrInfos[k], group = ctx.addrGroups[ai.addr.label];
        if (!group) continue;
        for (var i = 0; i < group.pathItems.length; i++) {
            var p = group.pathItems[i], name = String(p.name || '');
            if (name.indexOf('Outline:') !== 0 && name.indexOf('Label frame:') !== 0) continue;
            var copy = p.duplicate(tech.controls, ElementPlacement.PLACEATEND);
            setTechnicalStroke(copy, p.strokeColor, name.indexOf('Outline:') === 0 ? 2.835 : 1.417);
        }
    }
}

function createSheetDoc() {
    var doc = app.documents.add(
        DocumentColorSpace.CMYK,
        mm(SETTINGS.SHEET_WIDTH_MM),
        mm(SETTINGS.SHEET_HEIGHT_MM));
    if (!doc) throw new Error('Failed to create sheet document.');
    setMM(doc);
    try { app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM; } catch(e) {}
    doc.artboards[0].artboardRect = [
        0, mm(SETTINGS.SHEET_HEIGHT_MM),
        mm(SETTINGS.SHEET_WIDTH_MM), 0];
    return doc;
}
function createSheetRenderContext(sheetData) {
    return {
        sheetData:    sheetData,
        doc:          createSheetDoc(),
        sheetDocTop:  mm(SETTINGS.SHEET_HEIGHT_MM),
        addrGroups:   {},
        artCache:     {},
        artCacheGroup: null
    };
}
function ensureAddressGroup(ctx, addr) {
    if (!ctx || !ctx.addrGroups || !addr || !addr.label) return null;
    if (!ctx.addrGroups[addr.label]) {
        var grp  = ctx.doc.groupItems.add();
        grp.name = 'Addr:' + addr.label;
        ctx.addrGroups[addr.label] = grp;
    }
    return ctx.addrGroups[addr.label];
}
function buildAddressPlacements(sheetContexts) {
    var map = {};
    for (var s = 0; s < sheetContexts.length; s++) {
        var cells = sheetContexts[s].sheetData.cells;
        for (var i = 0; i < cells.length; i++) {
            var label = cells[i].addr.label;
            if (!map[label]) map[label] = [];
            var buckets = map[label];
            var last    = buckets.length ? buckets[buckets.length-1] : null;
            if (!last || last.ctx !== sheetContexts[s]) {
                last = { ctx: sheetContexts[s], cells: [] };
                buckets.push(last);
            }
            last.cells.push(cells[i]);
        }
    }
    return map;
}

// ─────────────────────────────────────────────────────────────────────────────
//  RENDER CACHE
// ─────────────────────────────────────────────────────────────────────────────
function createRenderCache() { return { docs: {}, list: [] }; }
function getRenderSourceDoc(cache, board) {
    var key = board.fileKey;
    if (!cache.docs[key]) {
        var doc = app.open(board.file);
        setMM(doc);
        cache.docs[key] = { doc: doc, state: { cleaned: false } };
        cache.list.push(cache.docs[key]);
    }
    return cache.docs[key];
}
function closeRenderCache(cache) {
    for (var i = 0; i < cache.list.length; i++) {
        try { cache.list[i].doc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  RENDER ADDRESS
// ─────────────────────────────────────────────────────────────────────────────
function recordRenderFailureForCell(addr, cell) {
    addr.stats.renderFailedRows = (addr.stats.renderFailedRows || 0) + 1;
    recordRenderFailedRow(addr.label);
    runStats.rowIssues.push({
        kind:           'render-failed',
        rowNumber:      cell.board.rowNumber  || '?',
        address:        addr.label,
        group:          '',
        product:        cell.board.product    || '',
        sourceFileName: cell.board.sourceFileName || '',
        message:        'Copy/paste/render failed'
    });
}
function afterRenderAttemptMaintenance() {
    renderAttemptCounter++;
    runStats.renderAttempts = renderAttemptCounter;
    if (SETTINGS.RENDER_MAINTENANCE_EVERY_CELLS > 0 &&
        renderAttemptCounter % SETTINGS.RENDER_MAINTENANCE_EVERY_CELLS === 0) {
        runStats.maintenanceRuns++;
        safeGC();
        safeRedraw();
        safeSleep(SETTINGS.RENDER_MAINTENANCE_DELAY_MS || 0);
    }
}
function renderOneCell(ctx, cell, cache) {
    var addr  = cell.addr;
    var group = ensureAddressGroup(ctx, addr);
    if (group) drawLabel(ctx.doc, group, addr, cell, ctx.sheetDocTop);
    
    var rendered = false;
    if (SETTINGS.RENDER_ART_CACHE) {
        rendered = renderCellFromArtCache(ctx, cell, cache, group);
        if (!rendered) runStats.artCacheFallbacks++;
    }
    
    if (!rendered) {
        if (!cell.board) return;
        var src = getRenderSourceDoc(cache, cell.board);
        rendered = group ? renderCell(src.doc, ctx.doc, cell, group, ctx.sheetDocTop, src.state) : false;
    }
    
    if (rendered) {
        addr.stats.placedRows = (addr.stats.placedRows || 0) + 1;
        recordPlacedRow(addr.label);
    } else {
        recordRenderFailureForCell(addr, cell);
    }
    afterRenderAttemptMaintenance();
}
function renderAddressAcrossSheets(addr, placements, cache) {
    if (!placements.length) return;
    var startMs = nowMs();
    try {
        for (var p = 0; p < placements.length; p++) {
            var ctx   = placements[p].ctx;
            var cells = placements[p].cells;
            for (var j = 0; j < cells.length; j++) {
                renderOneCell(ctx, cells[j], cache);
            }
        }
    } finally {
        if (addr.stats) addr.stats.renderMs += (nowMs() - startMs);
    }
}

function renderSheetContext(ctx, cache, sheetIndex, sheetCount) {
    var cells     = ctx.sheetData.cells;
    var lastLabel = '';
    var addrStart = nowMs();
    var addr      = null;
    
    for (var i = 0; i < cells.length; i++) {
        var cell = cells[i];
        if (!cell || !cell.addr) continue;
        if (!addr || addr.label !== cell.addr.label) {
            if (addr && addr.stats) addr.stats.renderMs += (nowMs() - addrStart);
            addr      = cell.addr;
            addrStart = nowMs();
            lastLabel = addr.label;
            log('Render sheet ' + sheetIndex + '/' + sheetCount + ': ' + lastLabel);
        }
        try {
            renderOneCell(ctx, cell, cache);
        } catch(cellError) {
            runStats.renderAttempts = (runStats.renderAttempts || 0) + 1;
            warn('Cell render error: row ' + (cell.board ? cell.board.rowNumber : '?') +
                 ' | ' + (cell.addr ? cell.addr.label : '?') +
                 ' | ' + (cellError && cellError.message ? cellError.message : cellError));
            recordRenderFailureForCell(cell.addr, cell);
        }
    }
    if (addr && addr.stats) addr.stats.renderMs += (nowMs() - addrStart);
}

function finalizeSheetContext(ctx) {
    for (var k = 0; k < ctx.sheetData.addrInfos.length; k++) {
        var ai = ctx.sheetData.addrInfos[k];
        if (!ctx.addrGroups[ai.addr.label]) continue;
        drawAddressOutline(ctx.doc, ctx.addrGroups[ai.addr.label], ai, ctx.sheetDocTop);
    }
    var tech = makeTechnicalLayers(ctx.doc);
    routeTaggedOperationalPaths(tech, ctx.doc);
    copyAddressControls(ctx, tech);
    safeGC();
}

function finalizeSheetContexts(sheetContexts) {
    for (var s = 0; s < sheetContexts.length; s++) {
        finalizeSheetContext(sheetContexts[s]);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  COUNTERS
// ─────────────────────────────────────────────────────────────────────────────
function saveDoc(doc, outFile) {
    var opt = new IllustratorSaveOptions();
    opt.pdfCompatible = SETTINGS.PDF_COMPATIBLE;
    opt.compressed    = SETTINGS.COMPRESSED_SAVE;
    doc.saveAs(outFile, opt);
}
function countSheetCells(sheets) {
    var n = 0;
    for (var i = 0; i < sheets.length; i++) n += sheets[i].cells ? sheets[i].cells.length : 0;
    return n;
}
function countAddressBoards(addresses) {
    var n = 0;
    for (var i = 0; i < addresses.length; i++) n += addresses[i].boards ? addresses[i].boards.length : 0;
    return n;
}
function countAddressPlaced(addresses) {
    var n = 0;
    for (var i = 0; i < addresses.length; i++)
        n += addresses[i].stats && addresses[i].stats.placedRows ? addresses[i].stats.placedRows : 0;
    return n;
}
function countIssuesByKind(kind) {
    var n = 0;
    for (var i = 0; i < runStats.rowIssues.length; i++)
        if (runStats.rowIssues[i].kind === kind) n++;
    return n;
}
function cellsOverlap(a, b) {
    return a.sheetX     < b.sheetX + b.w - 0.01 &&
           a.sheetX + a.w > b.sheetX     + 0.01 &&
           a.sheetY     < b.sheetY + b.h - 0.01 &&
           a.sheetY + a.h > b.sheetY     + 0.01;
}
function countLayoutOverlaps(sheets) {
    var overlaps = 0;
    for (var s = 0; s < sheets.length; s++) {
        var cells = sheets[s].cells || [];
        for (var i = 0; i < cells.length; i++)
            for (var j = i+1; j < cells.length; j++)
                if (cellsOverlap(cells[i], cells[j])) overlaps++;
    }
    return overlaps;
}

// A spreadsheet-ready verification table.  It is deliberately tab-separated:
// Excel opens it directly and Russian addresses do not conflict with commas.
function verificationText(value) {
    return String(value == null ? '' : value)
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/^\s+|\s+$/g, '');
}
function addressSheetList(sheets) {
    var byAddress = {};
    for (var s = 0; s < sheets.length; s++) {
        var cells = sheets[s].cells || [];
        for (var i = 0; i < cells.length; i++) {
            var label = cells[i].addr.label;
            if (!byAddress[label]) byAddress[label] = {};
            byAddress[label][s + 1] = true;
        }
    }
    var result = {};
    for (var label in byAddress) {
        if (!byAddress.hasOwnProperty(label)) continue;
        var numbers = [];
        for (var n in byAddress[label])
            if (byAddress[label].hasOwnProperty(n)) numbers.push(Number(n));
        numbers.sort(function(a, b) { return a - b; });
        result[label] = numbers.join(', ');
    }
    return result;
}
function addressIssueRows(label) {
    var rows = [];
    for (var i = 0; i < runStats.rowIssues.length; i++) {
        var issue = runStats.rowIssues[i];
        if (issue.address === label)
            rows.push(issue.rowNumber + ' (' + issue.kind + ')');
    }
    return rows.join(', ');
}
function writeVerificationTable(outFolder, sheets, addresses) {
    sheets = sheets || [];
    addresses = addresses || [];
    var f = new File(outFolder.fsName + '/verification-table.tsv');
    f.encoding = 'UTF-8';
    if (!f.open('w')) return;

    try {
        f.write('\uFEFF'); // Preserve Cyrillic when the file is opened directly in Excel.
        var W = function(values) {
            var row = [];
            for (var i = 0; i < values.length; i++) row.push(verificationText(values[i]));
            f.writeln(row.join('\t'));
        };
        var expected = runStats.expectedCopies || runStats.inputRows;
        var matched = countAddressBoards(addresses);
        var layout = countSheetCells(sheets);
        var placed = countAddressPlaced(addresses);
        var missing = countIssuesByKind('missing-source') +
                      countIssuesByKind('missing-artboard') +
                      countIssuesByKind('scan-error');
        var overlaps = countLayoutOverlaps(sheets);
        var overall = expected === matched + missing && matched === layout &&
                      layout === placed && overlaps === 0 &&
                      runStats.outputFiles.length === sheets.length;
        var sheetMap = addressSheetList(sheets);

        W(['CHECK', 'STATUS', 'EXPECTED', 'MATCHED', 'LAYOUT', 'PLACED',
           'MISSING', 'RENDER FAILED', 'SHEETS', 'ISSUE ROWS']);
        W(['OVERALL', overall ? 'OK' : 'CHECK', expected, matched, layout, placed,
           missing, countIssuesByKind('render-failed'), sheets.length, 'overlaps=' + overlaps]);
        W([]);
        W(['ADDRESS', 'STATUS', 'EXPECTED', 'MATCHED', 'PLACED', 'MISSING',
           'RENDER FAILED', 'SHEETS', 'ISSUE ROWS', 'RENDER SEC']);

        for (var ao = 0; ao < runStats.tableAddressOrder.length; ao++) {
            var label = runStats.tableAddressOrder[ao];
            var st = runStats.tableByAddress[label];
            var status = st.expectedRows === st.matchedRows &&
                         st.expectedRows === st.placedRows &&
                         st.missingRows === 0 && st.renderFailedRows === 0 ? 'OK' : 'CHECK';
            W([label, status, st.expectedRows, st.matchedRows, st.placedRows,
               st.missingRows, st.renderFailedRows, sheetMap[label] || '',
               addressIssueRows(label), Math.round((st.renderMs || 0) / 100) / 10]);
        }
    } finally {
        f.close();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  REPORT
// ─────────────────────────────────────────────────────────────────────────────
function writeReport(outFolder, sheets, addresses) {
    var t0 = nowMs();
    sheets    = sheets    || [];
    addresses = addresses || [];
    runStats.finishedAt      = isoNow();
    runStats.timings.totalMs = runStats.startedMs ? (nowMs() - runStats.startedMs) : 0;
    
    var f = new File(outFolder.fsName + '/report.txt');
    f.encoding = 'UTF-8';
    if (!f.open('w')) return;
    
    function W(s) { f.writeln(s == null ? '' : String(s)); }
    
    W('Bulch Enpack common build report');
    W('Started:  ' + runStats.startedAt);
    W('Finished: ' + runStats.finishedAt);
    W('Total time: ' + msText(runStats.timings.totalMs));
    if (runStats.fatalError) W('FATAL: ' + runStats.fatalError);
    W('');
    W('Source data: ' + (runStats.dataFile || '(not loaded)'));
    if (runStats.sourceTableFile) W('Selected table: ' + runStats.sourceTableFile);
    W('Sheet: ' + SETTINGS.SHEET_WIDTH_MM + 'x' + SETTINGS.SHEET_HEIGHT_MM + ' mm');
    W('Fuzzy threshold: ' + SETTINGS.FUZZY_THRESHOLD);
    W('');
    W('Table rows: '        + runStats.inputRows);
    W('Expected copies: '   + (runStats.expectedCopies || runStats.inputRows));
    W('Table addresses: '   + runStats.tableAddressOrder.length);
    W('Built addresses: '   + addresses.length);
    W('Matched artboards: ' + countAddressBoards(addresses));
    W('Layout cells: '      + countSheetCells(sheets));
    W('Placed artboards: '  + countAddressPlaced(addresses));
    W('Output sheets: '     + sheets.length);
    W('Warnings/errors: '   + errs.length);
    W('');
    W('Timings:');
    W('  load:     ' + msText(runStats.timings.loadMs));
    W('  scan:     ' + msText(runStats.timings.scanMs));
    W('  layout:   ' + msText(runStats.timings.layoutMs));
    W('  render:   ' + msText(runStats.timings.renderMs));
    W('  finalize: ' + msText(runStats.timings.finalizeMs));
    W('  save:     ' + msText(runStats.timings.saveMs));
    W('');
    
    W('Render diagnostics:');
    W('  attempted cells: ' + runStats.renderAttempts);
    W('  paste retry cells: ' + runStats.pasteRetryCells);
    W('  paste retry attempts: ' + runStats.pasteRetryAttempts);
    W('  paste failures after retries: ' + runStats.pasteRetryFailures);
    W('  art cache copies: ' + runStats.artCacheCopies);
    W('  art cache placements: ' + runStats.artCachePlacements);
    W('  art cache hits: ' + runStats.artCacheHits);
    W('  art cache fallbacks: ' + runStats.artCacheFallbacks);
    W('  art cache duplicate failures: ' + runStats.artCacheDupFails);
    W('  maintenance runs: ' + runStats.maintenanceRuns +
        ' (every ' + SETTINGS.RENDER_MAINTENANCE_EVERY_CELLS + ' cells)');
    W('');
    
    W('Checks:');
    var expected    = runStats.expectedCopies || runStats.inputRows;
    var matched     = countAddressBoards(addresses);
    var placed      = countAddressPlaced(addresses);
    var scanMissing = countIssuesByKind('missing-source') + 
                      countIssuesByKind('missing-artboard') + 
                      countIssuesByKind('scan-error');
    var cellCount   = countSheetCells(sheets);
    var overlaps    = countLayoutOverlaps(sheets);
    
    W('  expected copies == matched + missing: ' + 
        (expected === matched + scanMissing ? 'OK' : 'MISMATCH') +
        ' (' + expected + ' vs ' + matched + ' + ' + scanMissing + ')');
    W('  matched == layout cells: ' + 
        (matched === cellCount ? 'OK' : 'MISMATCH') +
        ' (' + matched + ' vs ' + cellCount + ')');
    W('  layout cells == placed: ' + 
        (cellCount === placed ? 'OK' : 'MISMATCH') +
        ' (' + cellCount + ' vs ' + placed + ')');
    W('  layout overlaps: ' + (overlaps === 0 ? 'OK' : 'MISMATCH') + 
        ' (' + overlaps + ')');
    W('  output files == sheets: ' + 
        (runStats.outputFiles.length === sheets.length ? 'OK' : 'MISMATCH') +
        ' (' + runStats.outputFiles.length + ' vs ' + sheets.length + ')');
    W('');
    
    W('Address summary (Expected=copies after quantity, Matched=artboard found, Placed=rendered):');
    for (var ao = 0; ao < runStats.tableAddressOrder.length; ao++) {
        var label  = runStats.tableAddressOrder[ao];
        var st     = runStats.tableByAddress[label];
        var status = (st.expectedRows === st.placedRows) ? 'OK' : 'CHECK';
        W('  [' + status + '] ' + label + 
            ' | expected='    + st.expectedRows + 
            ' matched='       + st.matchedRows  + 
            ' placed='        + st.placedRows   + 
            ' missing='       + st.missingRows  + 
            ' renderFailed='  + st.renderFailedRows);
    }
    W('');
    
    for (var si = 0; si < sheets.length; si++) {
        W('=== Sheet ' + (si+1) + ' — ' + sheets[si].cells.length + 
          ' cells, ' + sheets[si].addrInfos.length + ' addresses ===');
        for (var k = 0; k < sheets[si].addrInfos.length; k++) {
            var ai = sheets[si].addrInfos[k];
            W('  ' + ai.addr.label + ' (' + ai.cells.length + ' cells)' +
              ' bbox=' + r1(ai.outlineL) + 'x' + r1(ai.outlineT) + ' ' +
              r1(ai.outlineR - ai.outlineL) + 'x' + r1(ai.outlineB - ai.outlineT) + ' mm');
        }
    }
    W('');
    
    W('Output files:');
    if (runStats.outputFiles.length) {
        for (var ofi = 0; ofi < runStats.outputFiles.length; ofi++)
            W('  ' + runStats.outputFiles[ofi]);
    } else { W('  (none)'); }
    W('');
    
    if (runStats.rowIssues.length) {
        W('Row issues (' + runStats.rowIssues.length + '):');
        for (var ri = 0; ri < runStats.rowIssues.length; ri++) {
            var issue = runStats.rowIssues[ri];
            W('  [' + issue.kind + '] row=' + issue.rowNumber +
              ' addr=' + issue.address +
              ' product=' + issue.product +
              ' source=' + issue.sourceFileName +
              ' | ' + issue.message);
        }
        W('');
    }
    
    if (runStats.matchReviews.length) {
        W('Match review (' + runStats.matchReviews.length + '):');
        for (var mi = 0; mi < runStats.matchReviews.length; mi++) {
            var mr = runStats.matchReviews[mi];
            W('  [' + (mr.accepted ? 'AUTO' : 'MISS') + ' ' +
              scorePct(mr.score) + '% ' + mr.method + '] rows=' + mr.rows +
              (mr.count > 1 ? ' x' + mr.count : '') +
              ' addr=' + mr.address +
              ' product=' + mr.product +
              ' source=' + mr.sourceFileName +
              ' -> ' + (mr.matchedName || '(none)'));
        }
        W('');
    }
    
    if (errs.length) {
        W('Warnings (' + errs.length + '):');
        for (var e = 0; e < errs.length; e++) W('  ' + errs[e]);
        W('');
    }
    
    W('Log:');
    for (var l = 0; l < logs.length; l++) W(logs[l]);
    
    f.close();
    runStats.timings.reportMs = nowMs() - t0;
}

// ─────────────────────────────────────────────────────────────────────────────
//  SETTINGS DIALOG
// ─────────────────────────────────────────────────────────────────────────────
function showSettingsDialog() {
    var presets = [
        { label: '3000 × 2000 (альбомный)', w: 3000, h: 2000 },
        { label: '2000 × 3000 (книжный)',   w: 2000, h: 3000 },
        { label: '3030 × 2010 (альбомный)',  w: 3030, h: 2010 },
        { label: '2010 × 3030 (книжный)',    w: 2010, h: 3030 }
    ];

    try {
        var w = new Window('dialog', 'Bulch Enpack — Настройки');
        w.orientation = 'column';
        w.alignChildren = ['fill', 'top'];

        w.add('statictext', undefined, 'Размер листа (мм):');

        var radioGroup = w.add('panel', undefined, 'Предустановки');
        radioGroup.orientation = 'column';
        radioGroup.alignChildren = ['left', 'top'];
        var radios = [];
        for (var i = 0; i < presets.length; i++) {
            var r = radioGroup.add('radiobutton', undefined, presets[i].label);
            if (i === 0) r.value = true;
            radios.push(r);
        }

        var customGroup = w.add('group');
        customGroup.add('statictext', undefined, 'Ширина:');
        var widthEdit = customGroup.add('edittext', undefined, String(SETTINGS.SHEET_WIDTH_MM));
        widthEdit.characters = 6;
        customGroup.add('statictext', undefined, 'Высота:');
        var heightEdit = customGroup.add('edittext', undefined, String(SETTINGS.SHEET_HEIGHT_MM));
        heightEdit.characters = 6;

        for (var ri = 0; ri < radios.length; ri++) {
            radios[ri].onClick = (function(idx) {
                return function() {
                    widthEdit.text  = String(presets[idx].w);
                    heightEdit.text = String(presets[idx].h);
                };
            })(ri);
        }

        var buttons = w.add('group');
        buttons.alignment = 'right';
        var cancelBtn = buttons.add('button', undefined, 'Отмена');
        var okBtn     = buttons.add('button', undefined, 'Начать сборку');
        cancelBtn.onClick = function() { w.close(0); };
        okBtn.onClick     = function() { w.close(1); };

        if (w.show() !== 1) return false;

        var newW = parseInt(widthEdit.text, 10);
        var newH = parseInt(heightEdit.text, 10);
        if (isNaN(newW) || isNaN(newH) || newW < 100 || newH < 100) {
            alert('Недопустимые размеры листа.');
            return false;
        }
        SETTINGS.SHEET_WIDTH_MM  = newW;
        SETTINGS.SHEET_HEIGHT_MM = newH;
        return true;
    } catch(e) {
        // Fallback to confirm dialog if Window API fails
        if (!confirm('Размер листа: ' + SETTINGS.SHEET_WIDTH_MM + ' × ' +
                     SETTINGS.SHEET_HEIGHT_MM + ' мм.\n\nНачать сборку?')) {
            return false;
        }
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────────────────────
function main() {
    // ── Show settings dialog ─────────────────────────────────────
    try { app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS; } catch(ui0) {}
    if (!showSettingsDialog()) {
        return;
    }
    
    var prevUI = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    
    var renderCache  = null;
    var outf         = null;
    var addresses    = [];
    var sheets       = [];
    var sheetContexts = [];
    var totalBoards  = 0;
    
    runStats.startedAt = isoNow();
    runStats.startedMs = nowMs();
    
    try {
        var sf = scriptFolder();
        outf   = ensureDir(new Folder(sf.fsName + '/' + SETTINGS.OUTPUT_FOLDER_NAME));
        
        // ── Load data ──────────────────────────────────────────────
        var t0 = nowMs();
        var data = loadTableData();
        runStats.timings.loadMs = nowMs() - t0;
        
        // ── Scan artboards ────────────────────────────────────────
        t0        = nowMs();
        addresses = buildAddressesFromData(data);
        runStats.timings.scanMs = nowMs() - t0;
        if (!addresses.length) throw new Error('No addresses built from data');
        
        totalBoards = countAddressBoards(addresses);
        log('Addresses: ' + addresses.length + ', artboards matched: ' + totalBoards);
        
        // No layout or output documents are created until this check is approved.
        {
            try { app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS; } catch(ui1) {}
            var reviewOk = showPreflightDialog();
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
            if (!reviewOk) throw new Error('Cancelled before build after preflight check');
        }
        
        // ── Layout ────────────────────────────────────────────────
        t0     = nowMs();
        sheets = buildLayout(addresses);
        log('Sheets needed: ' + sheets.length);
        
        runStats.timings.layoutMs = nowMs() - t0;
        
        // ── Render ───────────────────────────────────────────────
        for (var s = 0; s < sheets.length; s++) {
            if (!sheets[s].cells.length) continue;
            
            var ctx = null;
            var sheetNumber = sheetContexts.length + 1;
            try {
                ctx = createSheetRenderContext(sheets[s]);
                sheetContexts.push(ctx);
                
                t0 = nowMs();
                renderCache = createRenderCache();
                log('Render sheet ' + sheetNumber + '/' + sheets.length +
                    ': ' + sheets[s].cells.length + ' cells');
                renderSheetContext(ctx, renderCache, sheetNumber, sheets.length);
                closeRenderCache(renderCache);
                renderCache = null;
                runStats.timings.renderMs += nowMs() - t0;
                clearSheetArtCache(ctx);
                
                t0 = nowMs();
                finalizeSheetContext(ctx);
                safeRedraw();
                runStats.timings.finalizeMs += nowMs() - t0;
                
                t0 = nowMs();
                var outFile = new File(
                    outf.fsName + '/' + SETTINGS.OUTPUT_BASENAME + pad(sheetNumber, 3) + '.ai');
                log('Saving: ' + outFile.fsName);
                saveDoc(ctx.doc, outFile);
                runStats.outputFiles.push(outFile.fsName);
                runStats.timings.saveMs += nowMs() - t0;
            } finally {
                if (renderCache) {
                    closeRenderCache(renderCache);
                    renderCache = null;
                }
                if (ctx && ctx.doc) {
                    clearSheetArtCache(ctx);
                    try { ctx.doc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
                    ctx.doc = null;
                }
                safeGC();
            }
        }
        
        // ── Finalize (outlines) ───────────────────────────────────
        // ── Save ─────────────────────────────────────────────────
        writeReport(outf, sheets, addresses);
        writeVerificationTable(outf, sheets, addresses);
        
        // ── Alert with summary ───────────────────────────────────
        var expected = runStats.expectedCopies || runStats.inputRows;
        var matched  = countAddressBoards(addresses);
        var placed   = countAddressPlaced(addresses);
        var missing  = expected - matched;
        
        var msg = 'Bulch Enpack — Done!' +
            '\n\nTable rows      : ' + runStats.inputRows +
            '\nExpected copies : ' + expected +
            '\nArtboards found : ' + matched + ' / ' + expected +
            (missing > 0 ? '  (!!! ' + missing + ' not found)' : '') +
            '\nPlaced          : ' + placed +
            '\nAddresses       : ' + addresses.length +
            '\nOutput sheets   : ' + sheetContexts.length +
            '\n\nPhase times (s):' +
            '\n  scan=' + (Math.round(runStats.timings.scanMs/100)/10) +
            '  render=' + (Math.round(runStats.timings.renderMs/100)/10) +
            '  save=' + (Math.round(runStats.timings.saveMs/100)/10) +
            '\n  total=' + (Math.round(runStats.timings.totalMs/100)/10) +
            '\n\nFolder: ' + outf.fsName +
            '\nVerification table: verification-table.tsv';
        
        if (runStats.pasteRetryCells) {
            msg += '\nPaste retries : ' + runStats.pasteRetryCells +
                   ' cells / ' + runStats.pasteRetryAttempts + ' attempts';
        }
        if (SETTINGS.RENDER_ART_CACHE) {
            msg += '\nArt cache      : ' + runStats.artCachePlacements +
                   ' placed / ' + runStats.artCacheCopies + ' source copies';
            if (runStats.artCacheFallbacks)
                msg += ' / ' + runStats.artCacheFallbacks + ' fallbacks';
        }
        if (runStats.matchReviews.length) {
            msg += '\nMatch review   : ' + runStats.matchReviews.length + ' items';
        }
        
        if (errs.length) {
            msg += '\n\nWarnings: ' + errs.length + '\nSee report.txt';
        }
        alert(msg);
        
    } catch(fatal) {
        runStats.fatalError = fatal && fatal.message ? fatal.message : String(fatal);
        errs.push('FATAL: ' + runStats.fatalError);
        logs.push('FATAL: ' + runStats.fatalError);
        try {
            if (!outf) outf = ensureDir(
                new Folder(scriptFolder().fsName + '/' + SETTINGS.OUTPUT_FOLDER_NAME));
            writeReport(outf, sheets, addresses);
            writeVerificationTable(outf, sheets, addresses);
        } catch(reportError) {}
        alert('Bulch Enpack — ERROR\n\n' + runStats.fatalError + 
              '\n\nSee report.txt');
    } finally {
        if (renderCache) closeRenderCache(renderCache);
        app.userInteractionLevel = prevUI;
    }
}

main();

}());
