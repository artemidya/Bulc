/*
  Bulch Enpack for Adobe Illustrator - table/category common build

  Data source:
    build/layout-data.json generated from the address table.

  Output:
    One or more common sheet_###.ai files in the out folder.
    No per-address .ai files are created.
*/

#target illustrator

(function () {

var SETTINGS = {
    OUTPUT_FOLDER_NAME: 'out',
    OUTPUT_BASENAME: 'sheet_',

    DATA_JSON_RELATIVE: 'build/layout-data.json',
    DATA_JSON_FALLBACKS: [
        'C:/BULCH/build/layout-data.json'
    ],

    SHEET_WIDTH_MM: 3030,
    SHEET_HEIGHT_MM: 2010,
    SHEET_MARGIN_MM: 0,

    ARTBOARD_GAP_MM: 3,
    COLUMN_GAP_MM: 3,

    LABEL_HEIGHT_MM: 7,
    LABEL_FONT_MM: 4,
    LABEL_INSET_MM: 1,

    DRAW_ADDRESS_OUTLINE: true,
    OUTLINE_STROKE_MM: 1,

    ROTATION_THRESHOLD_MM: 140,
    CLIP_TO_ARTBOARD: true,

    PDF_COMPATIBLE: false,
    COMPRESSED_SAVE: false,

    DELETE_GUIDES: true,
    CLEAN_ON_COPY_ERROR: true,

    TURBO_MODE: true,
    REDRAW_EVERY_N_ADDRESSES: 0
};

var PT = 72 / 25.4;
var MM = 25.4 / 72;
var MARKER_PFX = '__BEP__';
var logs = [];
var errs = [];

function mm(v) { return Number(v) * PT; }
function pt(v) { return Number(v) * MM; }
function r1(v) { return Math.round(pt(v) * 10) / 10; }
function log(s) { logs.push(String(s)); }
function warn(s) { errs.push(String(s)); logs.push('WARN: ' + s); }
function nowMs() { return (new Date()).getTime(); }
function safeGC() { try { $.gc(); } catch(e) {} }
function safeRedraw() { try { app.redraw(); } catch(e) {} }
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
function stripExt(n) {
    return String(n || '').replace(/\.ai$/i, '');
}
function decodedName(f) {
    var n = f.displayName || f.name;
    try { n = decodeURI(n); } catch(e) {}
    return n;
}
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
    var sf = scriptFolder();
    var local = new File(sf.fsName + '/' + SETTINGS.DATA_JSON_RELATIVE);
    if (local.exists) return local;
    for (var i=0; i<SETTINGS.DATA_JSON_FALLBACKS.length; i++) {
        var f = new File(SETTINGS.DATA_JSON_FALLBACKS[i]);
        if (f.exists) return f;
    }
    var chosen = File.openDialog('Select layout-data.json', '*.json');
    if (chosen && chosen.exists) return chosen;
    throw new Error('layout-data.json not found');
}

function setMM(doc) {
    try { app.preferences.setIntegerPreference('rulerType', RulerUnits.Millimeters); } catch(e) {}
    try { doc.rulerUnits = RulerUnits.Millimeters; } catch(e2) {}
    try { doc.rulerOrigin = [0, 0]; } catch(e3) {}
}
function activate(doc) {
    try { doc.activate(); } catch(e) { try { app.activeDocument = doc; } catch(e2) {} }
}
function unlockAll(doc) {
    function ul(layers) {
        for (var i=0; i<layers.length; i++) {
            try { layers[i].locked = false; } catch(e) {}
            try { layers[i].visible = true; } catch(e2) {}
            try { if (layers[i].layers) ul(layers[i].layers); } catch(e3) {}
        }
    }
    ul(doc.layers);
    for (var j=0; j<doc.pageItems.length; j++) {
        try { doc.pageItems[j].locked = false; } catch(e4) {}
        try { doc.pageItems[j].hidden = false; } catch(e5) {}
    }
    try { doc.selection = null; } catch(e6) {}
}
function removeGuides(doc) {
    if (!SETTINGS.DELETE_GUIDES) return 0;
    var removed = 0;
    for (var i=doc.pageItems.length-1; i>=0; i--) {
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

function black() {
    var c = new CMYKColor();
    c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
    return c;
}
function darkBlue() {
    var c = new CMYKColor();
    c.cyan = 100; c.magenta = 70; c.yellow = 15; c.black = 55;
    return c;
}
function darkBrown() {
    var c = new CMYKColor();
    c.cyan = 25; c.magenta = 65; c.yellow = 100; c.black = 60;
    return c;
}
function addressColor(addr) {
    return (addr.idx % 2 === 0) ? darkBlue() : darkBrown();
}

function trimString(value) {
    return String(value == null ? '' : value).replace(/^\s+|\s+$/g, '');
}
function normalizeName(value) {
    return trimString(value).toLowerCase()
        .replace(/ё/g, 'е')
        .replace(/[«»“”„]/g, '"')
        .replace(/\s+/g, ' ');
}
function looseNameKey(value) {
    return normalizeName(value)
        .replace(/["'`´’‘]/g, '')
        .replace(/[‐‑‒–—―-]/g, ' ')
        .replace(/[^\wа-яА-ЯёЁ0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .replace(/^\s+|\s+$/g, '');
}
function compactNameKey(value) {
    return looseNameKey(value).replace(/\s+/g, '');
}
function nameWords(value) {
    var key = looseNameKey(value);
    return key ? key.split(/\s+/) : [];
}
function wordCoverageScore(targetWords, candidateWords) {
    if (!targetWords.length || !candidateWords.length) return 0;
    var common = 0;
    for (var i=0; i<targetWords.length; i++) {
        for (var j=0; j<candidateWords.length; j++) {
            if (targetWords[i] === candidateWords[j]) { common++; break; }
        }
    }
    return (common / targetWords.length) * 0.7 + (common / candidateWords.length) * 0.3;
}
function levenshteinDistance(a, b) {
    a = String(a || ''); b = String(b || '');
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var row = [];
    for (var i=0; i<=b.length; i++) row[i] = i;
    for (var x=1; x<=a.length; x++) {
        var prev = row[0];
        row[0] = x;
        for (var y=1; y<=b.length; y++) {
            var old = row[y];
            var cost = a.charAt(x-1) === b.charAt(y-1) ? 0 : 1;
            row[y] = Math.min(row[y] + 1, row[y-1] + 1, prev + cost);
            prev = old;
        }
    }
    return row[b.length];
}
function similarityScore(a, b) {
    var maxLength = Math.max(a.length, b.length);
    return maxLength ? 1 - (levenshteinDistance(a, b) / maxLength) : 1;
}
function findArtboardByName(doc, name) {
    var target = normalizeName(name);
    for (var i=0; i<doc.artboards.length; i++) {
        if (normalizeName(doc.artboards[i].name) === target) return i;
    }
    var looseTarget = looseNameKey(name);
    for (var j=0; j<doc.artboards.length; j++) {
        if (looseNameKey(doc.artboards[j].name) === looseTarget) return j;
    }
    var compactTarget = compactNameKey(name);
    for (var k=0; k<doc.artboards.length; k++) {
        if (compactNameKey(doc.artboards[k].name) === compactTarget) return k;
    }

    var targetWords = nameWords(name);
    var bestIndex = -1, bestScore = 0, bestMethod = '';
    for (var n=0; n<doc.artboards.length; n++) {
        var candidateName = doc.artboards[n].name;
        var compactCandidate = compactNameKey(candidateName);
        var candidateWords = nameWords(candidateName);
        var score = 0, method = '';
        if (compactTarget.length >= 7 && compactCandidate.indexOf(compactTarget) === 0) {
            score = 0.92; method = 'prefix';
        } else if (compactCandidate.length >= 7 && compactTarget.indexOf(compactCandidate) >= 0 && candidateWords.length >= 2) {
            score = 0.90; method = 'contains';
        } else {
            var ws = wordCoverageScore(targetWords, candidateWords);
            var es = similarityScore(compactTarget, compactCandidate);
            score = Math.max(ws, es);
            method = ws >= es ? 'words' : 'edit';
        }
        if (score > bestScore) { bestScore = score; bestIndex = n; bestMethod = method; }
    }
    if (bestIndex >= 0 && bestScore >= 0.82) {
        log("FUZZY: '" + name + "' -> '" + doc.artboards[bestIndex].name + "' (" + bestMethod + ', ' + Math.round(bestScore * 100) + '%)');
        return bestIndex;
    }
    return -1;
}
function artboardNameList(doc) {
    var names = [];
    for (var i=0; i<doc.artboards.length; i++) names.push(doc.artboards[i].name);
    return names.join(' | ');
}

function rectInfo(rect) {
    var ow = rect[2] - rect[0];
    var oh = rect[1] - rect[3];
    return { origW: ow, origH: oh, w: ow, h: oh };
}
function applyRotation(board) {
    if (board.origW > mm(SETTINGS.ROTATION_THRESHOLD_MM) + 0.01) {
        board.w = board.origH;
        board.h = board.origW;
        board.rotated90 = true;
    }
}

function loadTableData() {
    var f = resolveDataFile();
    var data = parseJson(readText(f));
    if (!data || !data.items || !data.items.length) throw new Error('No items in: ' + f.fsName);
    log('Data: ' + f.fsName);
    log('Rows: ' + data.items.length);
    return data;
}

function buildAddressesFromData(data) {
    var addrMap = {};
    var addresses = [];
    var sourceDocs = {};
    var sourceDocList = [];

    function getAddress(label) {
        if (!addrMap[label]) {
            var a = {
                idx: addresses.length,
                label: label,
                boards: [],
                stats: { scanMs: 0, renderMs: 0, boardCount: 0 }
            };
            addrMap[label] = a;
            addresses.push(a);
        }
        return addrMap[label];
    }
    function getSourceDoc(path) {
        if (!sourceDocs[path]) {
            var file = new File(path);
            if (!file.exists) throw new Error('Missing source file: ' + path);
            var doc = app.open(file);
            setMM(doc);
            sourceDocs[path] = { doc: doc, file: file };
            sourceDocList.push(sourceDocs[path]);
        }
        return sourceDocs[path].doc;
    }

    var startMs = nowMs();
    try {
        for (var i=0; i<data.items.length; i++) {
            var item = data.items[i];
            if (!item || !item.sourceFile) {
                warn('Row skipped without sourceFile: ' + (item ? item.rowNumber : '?'));
                continue;
            }
            var addr = getAddress(item.address || 'output');
            var srcDoc = getSourceDoc(item.sourceFile);
            var abIndex = findArtboardByName(srcDoc, item.product);
            if (abIndex < 0) {
                warn('Artboard not found: row ' + item.rowNumber + ' | ' + item.sourceFileName + ' | ' + item.product);
                log('  ARTBOARDS: ' + artboardNameList(srcDoc));
                continue;
            }
            var rect = srcDoc.artboards[abIndex].artboardRect;
            var ri = rectInfo(rect);
            var board = {
                idx: abIndex,
                rect: [rect[0], rect[1], rect[2], rect[3]],
                origW: ri.origW,
                origH: ri.origH,
                w: ri.w,
                h: ri.h,
                rotated90: false,
                file: new File(item.sourceFile),
                fileKey: item.sourceFile,
                product: item.product,
                rowNumber: item.rowNumber,
                sourceFileName: item.sourceFileName || decodedName(new File(item.sourceFile))
            };
            applyRotation(board);
            addr.boards.push(board);
            addr.stats.boardCount = addr.boards.length;
        }
    } finally {
        for (var c=0; c<sourceDocList.length; c++) {
            try { sourceDocList[c].doc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
        }
        safeGC();
    }
    for (var a=0; a<addresses.length; a++) addresses[a].stats.scanMs = nowMs() - startMs;
    return addresses;
}

function buildLayout(addresses) {
    var mg = mm(SETTINGS.SHEET_MARGIN_MM);
    var sheetW = mm(SETTINGS.SHEET_WIDTH_MM) - mg * 2;
    var sheetH = mm(SETTINGS.SHEET_HEIGHT_MM) - mg * 2;
    var labelH = mm(SETTINGS.LABEL_HEIGHT_MM);
    var abGap = mm(SETTINGS.ARTBOARD_GAP_MM);
    var colGap = mm(SETTINGS.COLUMN_GAP_MM);

    for (var a0=0; a0<addresses.length; a0++) {
        for (var b0=0; b0<addresses[a0].boards.length; b0++) {
            var brd = addresses[a0].boards[b0];
            if (brd.w > sheetW + 0.01) throw new Error('Board wider than sheet: ' + addresses[a0].label + ' row ' + brd.rowNumber);
            if (brd.h + labelH > sheetH + 0.01) throw new Error('Board taller than sheet: ' + addresses[a0].label + ' row ' + brd.rowNumber);
        }
    }

    var sheets = [];
    var addrMap = {};
    function newSheet() {
        var s = { cells: [], addrInfos: [], colX: mg, colY: mg, colW: 0 };
        sheets.push(s);
        return s;
    }
    function getOrCreateAddrInfo(sheet, addr) {
        if (!addrMap[addr.label]) {
            var ai = {
                addr: addr,
                cells: [],
                outlineL: 1e12,
                outlineT: 1e12,
                outlineR: -1e12,
                outlineB: -1e12,
                columns: {}
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
        addrMap = {};
    }

    for (var a=0; a<addresses.length; a++) {
        var addr = addresses[a];
        if (!addr.boards.length) continue;
        for (var b=0; b<addr.boards.length; b++) {
            var board = addr.boards[b];
            var needLabel = !addrMap[addr.label];
            var topPad = needLabel ? labelH : 0;
            var cellH = topPad + board.h;
            var fitsV = (curSheet.colY + cellH <= mg + sheetH + 0.01);
            if (!fitsV) {
                var moved = flushColumn(curSheet, board.w);
                if (moved) {
                    needLabel = !addrMap[addr.label];
                    topPad = needLabel ? labelH : 0;
                    cellH = topPad + board.h;
                    fitsV = (curSheet.colY + cellH <= mg + sheetH + 0.01);
                }
                if (!fitsV || !moved) {
                    startNewSheet();
                    needLabel = true;
                    topPad = labelH;
                    cellH = topPad + board.h;
                }
            }
            var colRightAfter = curSheet.colX + Math.max(curSheet.colW, board.w);
            if (colRightAfter > mg + sheetW + 0.01) {
                startNewSheet();
                needLabel = true;
                topPad = labelH;
                cellH = topPad + board.h;
            }

            var cellX = curSheet.colX;
            var cellY = curSheet.colY + topPad;
            var labelY = needLabel ? curSheet.colY : -1;
            var cell = { addr: addr, board: board, sheetX: cellX, sheetY: cellY, labelY: labelY, w: board.w, h: board.h };
            curSheet.cells.push(cell);
            curSheet.colY += cellH + abGap;
            if (board.w > curSheet.colW) curSheet.colW = board.w;

            var ai = getOrCreateAddrInfo(curSheet, addr);
            ai.cells.push(cell);
            var topEdge = (labelY >= 0) ? labelY : cellY;
            if (cellX < ai.outlineL) ai.outlineL = cellX;
            if (topEdge < ai.outlineT) ai.outlineT = topEdge;
            if (cellX + board.w > ai.outlineR) ai.outlineR = cellX + board.w;
            if (cellY + board.h > ai.outlineB) ai.outlineB = cellY + board.h;
            var colKey = String(cellX);
            if (!ai.columns[colKey]) {
                ai.columns[colKey] = { left: cellX, right: cellX + board.w, top: topEdge, bottom: cellY + board.h };
            } else {
                var col = ai.columns[colKey];
                if (cellX + board.w > col.right) col.right = cellX + board.w;
                if (topEdge < col.top) col.top = topEdge;
                if (cellY + board.h > col.bottom) col.bottom = cellY + board.h;
            }
        }
    }
    return sheets;
}

function tryCopyBoard(srcDoc, board, mName) {
    srcDoc.selection = null;
    srcDoc.artboards.setActiveArtboardIndex(board.idx);
    var mk = srcDoc.pathItems.rectangle(board.rect[1], board.rect[0], board.origW, board.origH);
    mk.name = mName;
    mk.filled = false;
    mk.stroked = false;
    mk.selected = true;
    try { srcDoc.selectObjectsOnActiveArtboard(); } catch(e) {}
    mk.selected = true;
    if (!srcDoc.selection || srcDoc.selection.length <= 1) {
        try { mk.remove(); } catch(e2) {}
        srcDoc.selection = null;
        return false;
    }
    app.executeMenuCommand('copy');
    try { mk.remove(); } catch(e3) {}
    srcDoc.selection = null;
    return true;
}
function copyBoard(srcDoc, board, mName, docState, label) {
    if (tryCopyBoard(srcDoc, board, mName)) return true;
    if (!SETTINGS.CLEAN_ON_COPY_ERROR || docState.cleaned) return false;
    docState.cleaned = true;
    var guideCount = prepareSourceDoc(srcDoc);
    if (guideCount > 0) log('Removed guides on retry: ' + label + ' (' + guideCount + ')');
    else log('Retried copy after source cleanup: ' + label);
    return tryCopyBoard(srcDoc, board, mName);
}
function findByName(sel, name) {
    for (var i=0; i<sel.length; i++) {
        try { if (sel[i].name === name) return sel[i]; } catch(e) {}
    }
    return null;
}
function toArr(sel) {
    var a = [];
    for (var i=0; i<sel.length; i++) a.push(sel[i]);
    return a;
}
function rotateGroup90(grp, destL, destT) {
    try {
        grp.rotate(-90, true, true, true, true, Transformation.CENTER);
        var gb = grp.geometricBounds;
        grp.translate(destL - gb[0], destT - gb[1]);
    } catch(e) { warn('Rotate failed: ' + e.message); }
}

function renderCell(srcDoc, destDoc, cell, parentGroup, sheetDocTop, docState) {
    var board = cell.board;
    var addr = cell.addr;
    var destL = cell.sheetX;
    var destT = sheetDocTop - cell.sheetY;
    var mName = MARKER_PFX + board.rowNumber + '_' + board.idx + '_' + nowMs();

    activate(srcDoc);
    if (!copyBoard(srcDoc, board, mName, docState, addr.label)) {
        warn('Empty board skipped: ' + addr.label + ' row ' + board.rowNumber + ' ' + board.product);
        return;
    }

    activate(destDoc);
    destDoc.selection = null;
    try { app.executeMenuCommand('pasteInPlace'); }
    catch(e) { app.executeMenuCommand('paste'); }
    if (!destDoc.selection || !destDoc.selection.length) {
        warn('Paste failed: ' + addr.label + ' row ' + board.rowNumber);
        return;
    }

    var pasted = toArr(destDoc.selection);
    var mk = findByName(pasted, mName);
    var mkB = mk ? mk.geometricBounds : [board.rect[0], board.rect[1], board.rect[2], board.rect[3]];
    if (!mk) warn('Marker missing fallback: ' + addr.label + ' row ' + board.rowNumber);

    var artGrp = destDoc.groupItems.add();
    artGrp.name = addr.label + '/row' + board.rowNumber + '/' + board.product;
    for (var i=0; i<pasted.length; i++) {
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
        var cr = destDoc.pathItems.rectangle(destT, destL, board.w, board.h);
        cr.filled = false;
        cr.stroked = false;
        cr.clipping = true;
        cr.move(cg, ElementPlacement.PLACEATBEGINNING);
        artGrp.move(cg, ElementPlacement.PLACEATEND);
        cg.clipped = true;
        cg.move(parentGroup, ElementPlacement.PLACEATEND);
    } else {
        artGrp.move(parentGroup, ElementPlacement.PLACEATEND);
    }
    destDoc.selection = null;
}

function drawLabel(destDoc, group, addr, cell, sheetDocTop) {
    if (cell.labelY < 0) return;
    var inset = mm(SETTINGS.LABEL_INSET_MM);
    var labelT = sheetDocTop - cell.labelY;
    var boxH = mm(SETTINGS.LABEL_HEIGHT_MM) - inset * 2;
    if (boxH < mm(3)) boxH = mm(3);
    var box = destDoc.pathItems.rectangle(labelT - inset, cell.sheetX + inset, cell.w - inset * 2, boxH);
    box.filled = false;
    box.stroked = false;
    var tf = destDoc.textFrames.areaText(box);
    tf.name = 'Label:' + addr.label;
    tf.contents = addr.label;
    try {
        tf.textRange.characterAttributes.size = mm(SETTINGS.LABEL_FONT_MM);
        tf.textRange.characterAttributes.fillColor = addressColor(addr);
    } catch(e) {}
    try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch(e2) {}
    tf.move(group, ElementPlacement.PLACEATEND);
}
function addressColumns(ai) {
    var cols = [];
    for (var key in ai.columns) if (ai.columns.hasOwnProperty(key)) cols.push(ai.columns[key]);
    cols.sort(function(a,b){ return a.left - b.left; });
    for (var i=0; i<cols.length; i++) {
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
    for (var i=1; i<cols.length; i++) {
        var x = cols[i].left;
        if (pts[pts.length-1][0] !== x || pts[pts.length-1][1] !== cols[i-1].top) pts.push([x, cols[i-1].top]);
        if (cols[i].top !== cols[i-1].top) pts.push([x, cols[i].top]);
        pts.push([cols[i].shapeRight, cols[i].top]);
    }
    pts.push([cols[cols.length-1].shapeRight, cols[cols.length-1].bottom]);
    pts.push([cols[cols.length-1].left, cols[cols.length-1].bottom]);
    for (i=cols.length-2; i>=0; i--) {
        x = cols[i+1].left;
        if (pts[pts.length-1][0] !== x || pts[pts.length-1][1] !== cols[i+1].bottom) pts.push([x, cols[i+1].bottom]);
        if (cols[i].bottom !== cols[i+1].bottom) pts.push([x, cols[i].bottom]);
        pts.push([cols[i].left, cols[i].bottom]);
    }
    return pts;
}
function drawAddressOutline(destDoc, group, ai, sheetDocTop) {
    if (!SETTINGS.DRAW_ADDRESS_OUTLINE) return;
    var cols = addressColumns(ai);
    var pts = steppedOutlinePoints(cols);
    if (pts.length < 4) return;
    var aiPts = [];
    for (var i=0; i<pts.length; i++) aiPts.push([pts[i][0], sheetDocTop - pts[i][1]]);
    var p = destDoc.pathItems.add();
    p.name = 'Outline:' + ai.addr.label;
    p.setEntirePath(aiPts);
    p.closed = true;
    p.filled = false;
    p.stroked = true;
    p.strokeWidth = mm(SETTINGS.OUTLINE_STROKE_MM);
    p.strokeColor = addressColor(ai.addr);
    p.move(group, ElementPlacement.PLACEATBEGINNING);
}

function createSheetDoc() {
    var doc = app.documents.add(DocumentColorSpace.CMYK, mm(SETTINGS.SHEET_WIDTH_MM), mm(SETTINGS.SHEET_HEIGHT_MM));
    setMM(doc);
    try { app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM; } catch(e) {}
    doc.artboards[0].artboardRect = [0, mm(SETTINGS.SHEET_HEIGHT_MM), mm(SETTINGS.SHEET_WIDTH_MM), 0];
    return doc;
}
function createSheetRenderContext(sheetData) {
    return { sheetData: sheetData, doc: createSheetDoc(), sheetDocTop: mm(SETTINGS.SHEET_HEIGHT_MM), addrGroups: {} };
}
function ensureAddressGroup(ctx, addr) {
    if (!ctx.addrGroups[addr.label]) {
        var grp = ctx.doc.groupItems.add();
        grp.name = 'Addr:' + addr.label;
        ctx.addrGroups[addr.label] = grp;
    }
    return ctx.addrGroups[addr.label];
}
function buildAddressPlacements(sheetContexts) {
    var map = {};
    for (var s=0; s<sheetContexts.length; s++) {
        var cells = sheetContexts[s].sheetData.cells;
        for (var i=0; i<cells.length; i++) {
            var label = cells[i].addr.label;
            if (!map[label]) map[label] = [];
            var buckets = map[label];
            var last = buckets.length ? buckets[buckets.length-1] : null;
            if (!last || last.ctx !== sheetContexts[s]) {
                last = { ctx: sheetContexts[s], cells: [] };
                buckets.push(last);
            }
            last.cells.push(cells[i]);
        }
    }
    return map;
}

function createRenderCache() {
    return { docs: {}, list: [] };
}
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
    for (var i=0; i<cache.list.length; i++) {
        try { cache.list[i].doc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
    }
}
function renderAddressAcrossSheets(addr, placements, cache) {
    if (!placements.length) return;
    var startMs = nowMs();
    try {
        for (var p=0; p<placements.length; p++) {
            var ctx = placements[p].ctx;
            var group = ensureAddressGroup(ctx, addr);
            var cells = placements[p].cells;
            for (var j=0; j<cells.length; j++) {
                var cell = cells[j];
                drawLabel(ctx.doc, group, addr, cell, ctx.sheetDocTop);
                var src = getRenderSourceDoc(cache, cell.board);
                renderCell(src.doc, ctx.doc, cell, group, ctx.sheetDocTop, src.state);
            }
        }
    } finally {
        if (addr.stats) addr.stats.renderMs += (nowMs() - startMs);
    }
}
function finalizeSheetContexts(sheetContexts) {
    for (var s=0; s<sheetContexts.length; s++) {
        var ctx = sheetContexts[s];
        for (var k=0; k<ctx.sheetData.addrInfos.length; k++) {
            var ai = ctx.sheetData.addrInfos[k];
            if (!ctx.addrGroups[ai.addr.label]) continue;
            drawAddressOutline(ctx.doc, ctx.addrGroups[ai.addr.label], ai, ctx.sheetDocTop);
        }
        safeGC();
    }
}
function saveDoc(doc, outFile) {
    var opt = new IllustratorSaveOptions();
    opt.pdfCompatible = SETTINGS.PDF_COMPATIBLE;
    opt.compressed = SETTINGS.COMPRESSED_SAVE;
    doc.saveAs(outFile, opt);
}
function writeReport(outFolder, sheets, addresses) {
    var f = new File(outFolder.fsName + '/report.txt');
    f.encoding = 'UTF-8';
    if (!f.open('w')) return;
    f.writeln('Bulch Enpack common build report');
    f.writeln('Sheet: ' + SETTINGS.SHEET_WIDTH_MM + 'x' + SETTINGS.SHEET_HEIGHT_MM + ' mm');
    f.writeln('Addresses: ' + addresses.length + '   Sheets: ' + sheets.length);
    f.writeln('');
    for (var s=0; s<sheets.length; s++) {
        f.writeln('=== Sheet ' + (s+1) + ' - ' + sheets[s].cells.length + ' cells, ' + sheets[s].addrInfos.length + ' addresses ===');
        for (var k=0; k<sheets[s].addrInfos.length; k++) {
            var ai = sheets[s].addrInfos[k];
            f.writeln('  ' + ai.addr.label + ' (' + ai.cells.length + ' cells)' +
                ' bbox=' + r1(ai.outlineL) + 'x' + r1(ai.outlineT) +
                ' ' + r1(ai.outlineR - ai.outlineL) + 'x' + r1(ai.outlineB - ai.outlineT) + ' mm');
        }
    }
    if (errs.length) {
        f.writeln('');
        f.writeln('Warnings (' + errs.length + '):');
        for (var e=0; e<errs.length; e++) f.writeln('  ' + errs[e]);
    }
    f.writeln('');
    f.writeln('Log:');
    for (var l=0; l<logs.length; l++) f.writeln(logs[l]);
    f.close();
}

function main() {
    var prevUI = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    var renderCache = null;
    try {
        var sf = scriptFolder();
        var outf = ensureDir(new Folder(sf.fsName + '/' + SETTINGS.OUTPUT_FOLDER_NAME));

        var data = loadTableData();
        var addresses = buildAddressesFromData(data);
        if (!addresses.length) throw new Error('No addresses built from data');

        log('Addresses: ' + addresses.length);
        var totalBoards = 0;
        for (var ai=0; ai<addresses.length; ai++) totalBoards += addresses[ai].boards.length;
        log('Selected artboards: ' + totalBoards);

        var sheets = buildLayout(addresses);
        log('Sheets needed: ' + sheets.length);

        var sheetContexts = [];
        for (var s=0; s<sheets.length; s++) {
            if (sheets[s].cells.length) sheetContexts.push(createSheetRenderContext(sheets[s]));
        }

        var placementsByAddress = buildAddressPlacements(sheetContexts);
        renderCache = createRenderCache();
        for (var a=0; a<addresses.length; a++) {
            log('Rendering address ' + (a+1) + '/' + addresses.length + ': ' + addresses[a].label);
            renderAddressAcrossSheets(addresses[a], placementsByAddress[addresses[a].label] || [], renderCache);
            if ((a + 1) % 5 === 0) safeGC();
            maybeRedraw(a + 1, SETTINGS.REDRAW_EVERY_N_ADDRESSES);
        }
        closeRenderCache(renderCache);
        renderCache = null;

        finalizeSheetContexts(sheetContexts);
        safeRedraw();

        for (var s2=0; s2<sheetContexts.length; s2++) {
            var sheetDoc = null;
            try {
                sheetDoc = sheetContexts[s2].doc;
                var outFile = new File(outf.fsName + '/' + SETTINGS.OUTPUT_BASENAME + pad(s2+1, 3) + '.ai');
                log('Saving: ' + outFile.fsName);
                saveDoc(sheetDoc, outFile);
            } finally {
                if (sheetDoc) try { sheetDoc.close(SaveOptions.DONOTSAVECHANGES); } catch(e) {}
                safeGC();
            }
        }
        writeReport(outf, sheets, addresses);

        var msg = 'Done!\n\n' +
            'Addresses     : ' + addresses.length + '\n' +
            'Artboards     : ' + totalBoards + '\n' +
            'Output sheets : ' + sheetContexts.length + '\n' +
            'Output folder : ' + outf.fsName;
        if (errs.length) msg += '\n\nWarnings: ' + errs.length + '\nSee report.txt';
        alert(msg);
    } finally {
        if (renderCache) closeRenderCache(renderCache);
        app.userInteractionLevel = prevUI;
    }
}

main();

}());
