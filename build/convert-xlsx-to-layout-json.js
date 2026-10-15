#!/usr/bin/env node
/*
  convert-xlsx-to-layout-json.js
  ─────────────────────────────────────────────────────────────────────────────
  Converts a Bulch Enpack Excel table into layout-data.json for B_Enpack.jsx.

  Usage:
    node convert-xlsx-to-layout-json.js <input.xlsx> <output.json> [sourceFolder]

  Arguments:
    input.xlsx     — Excel workbook (first sheet is used)
    output.json    — destination JSON file (layout-data.json)
    sourceFolder   — optional; folder with .ai source files (default: ../M)

  Excel column mapping (first row = header):
    A (col 1) = Address   — delivery address / output file grouping key
    B (col 2) = Group     — product group (matched to .ai source file)
    C (col 3) = Product   — artboard name inside the .ai source file
    D (col 4) = Total     — quantity; 0 or empty rows are skipped

  Output format:
    {
      "generatedAt": "ISO timestamp",
      "settings": { "repeatByItog": false, ... },
      "items": [
        {
          "address": "...",
          "group": "...",
          "product": "...",
          "total": "...",
          "sourceFile": "absolute path to .ai",
          "sourceFileName": "filename.ai",
          "outputFileName": "Address.ai",
          "rowNumber": 2,
          "copyIndex": 1
        }, ...
      ]
    }
*/

'use strict';

// ─────────────────────────────────────────────────────────────────────────────
//  DEPENDENCIES
// ─────────────────────────────────────────────────────────────────────────────
var path = require('path');
var fs   = require('fs');

var XLSX;
try {
    XLSX = require('xlsx');
} catch (e) {
    console.error('ERROR: xlsx module not found.');
    console.error('Run "npm install" in the build/ folder first:');
    console.error('  cd ' + __dirname);
    console.error('  npm install');
    process.exit(1);
}

// ─────────────────────────────────────────────────────────────────────────────
//  ARGUMENTS
// ─────────────────────────────────────────────────────────────────────────────
var args = process.argv.slice(2);
if (args.length < 2) {
    console.error('Usage: node convert-xlsx-to-layout-json.js <input.xlsx> <output.json> [sourceFolder]');
    process.exit(1);
}

var inputPath    = path.resolve(args[0]);
var outputPath   = path.resolve(args[1]);
var scriptDir    = __dirname;                                   // build/
var projectRoot  = path.resolve(scriptDir, '..');               // project root
var sourceFolder = args[2]
    ? path.resolve(args[2])
    : path.resolve(projectRoot, 'M');                           // default: ../M
var aliasPath    = path.resolve(projectRoot, 'group-file-aliases.tsv');

// Validate input
if (!fs.existsSync(inputPath)) {
    console.error('ERROR: Input file not found: ' + inputPath);
    process.exit(1);
}

// Ensure output directory exists
var outputDir = path.dirname(outputPath);
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function trim(v) {
    if (v == null) return '';
    return String(v).replace(/^\s+|\s+$/g, '');
}

function toNumber(v) {
    var s = trim(v).replace(',', '.');
    if (!s) return 0;
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

function isNonZeroTotal(v) {
    var s = trim(v);
    if (!s) return false;
    var n = toNumber(s);
    if (n !== 0) return true;
    // Non-numeric text counts as "included" (same as PowerShell version)
    return isNaN(n);
}

function safeFileName(name) {
    var s = trim(name);
    s = s.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_');
    s = s.replace(/\s+/g, ' ').replace(/^[\s.]+|[\s.]+$/g, '');
    return s || 'output';
}

function normalizeKey(text) {
    return trim(text).toLowerCase()
        .replace(/[\s\p{P}\p{S}_-]+/gu, '');
}

function wordKey(text) {
    var words = trim(text).toLowerCase()
        .split(/[\s\p{P}\p{S}_-]+/gu)
        .filter(function (w) { return w.length > 0; });
    return words.sort().join('|');
}

// ─────────────────────────────────────────────────────────────────────────────
//  READ ALIASES (group-file-aliases.tsv)
// ─────────────────────────────────────────────────────────────────────────────
function readAliases(filePath) {
    var aliases = {};
    if (!fs.existsSync(filePath)) return aliases;

    var content = fs.readFileSync(filePath, 'utf-8');
    // Strip BOM
    if (content.charCodeAt(0) === 0xFEFF) content = content.substring(1);

    var lines = content.split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line || line.charAt(0) === '#') continue;
        var parts = lines[i].split('\t');
        if (parts.length < 2) continue;
        var group = trim(parts[0]);
        var file  = trim(parts[1]);
        if (group && file) aliases[group] = file;
    }
    return aliases;
}

// ─────────────────────────────────────────────────────────────────────────────
//  BUILD AI FILE INDEX
// ─────────────────────────────────────────────────────────────────────────────
function buildAiFileIndex(folderPath) {
    var byName   = {};
    var byBase   = {};
    var byNorm   = {};
    var byWords  = {};
    var files    = [];

    if (!fs.existsSync(folderPath)) {
        console.error('WARNING: Source folder not found: ' + folderPath);
        return { byName: byName, byBase: byBase, byNorm: byNorm, byWords: byWords, files: files };
    }

    var entries = fs.readdirSync(folderPath);
    for (var i = 0; i < entries.length; i++) {
        var name = entries[i];
        if (!/\.ai$/i.test(name)) continue;
        var fullPath = path.resolve(folderPath, name);
        var stat;
        try { stat = fs.statSync(fullPath); } catch (e) { continue; }
        if (!stat.isFile()) continue;

        var baseName = path.basename(name, '.ai');
        var entry = { name: name, baseName: baseName, fullPath: fullPath };
        files.push(entry);
        byName[name] = entry;
        byBase[baseName] = entry;
        byNorm[normalizeKey(baseName)] = entry;
        var wk = wordKey(baseName);
        if (!byWords[wk]) byWords[wk] = entry;
    }
    return { byName: byName, byBase: byBase, byNorm: byNorm, byWords: byWords, files: files };
}

// ─────────────────────────────────────────────────────────────────────────────
//  RESOLVE AI FILE FOR A GROUP
// ─────────────────────────────────────────────────────────────────────────────
function resolveAiFile(group, aliases, index) {
    var candidates = [];

    // 1. Explicit alias
    if (aliases[group]) candidates.push(aliases[group]);

    // 2. Direct name and name.ai
    candidates.push(group);
    candidates.push(group + '.ai');

    for (var c = 0; c < candidates.length; c++) {
        var candidate = candidates[c];
        if (!candidate) continue;

        // Absolute path
        if (path.isAbsolute(candidate) && fs.existsSync(candidate)) {
            return candidate;
        }

        var name = path.basename(candidate);
        var base = path.basename(candidate, '.ai');

        if (index.byName[name])   return index.byName[name].fullPath;
        if (index.byBase[base])   return index.byBase[base].fullPath;

        // Relative to source folder
        var relPath = path.resolve(sourceFolder, name);
        if (fs.existsSync(relPath)) return relPath;
    }

    // 3. Normalized key
    var normKey = normalizeKey(group);
    if (index.byNorm[normKey]) return index.byNorm[normKey].fullPath;

    // 4. Word key
    var wk = wordKey(group);
    if (index.byWords[wk]) return index.byWords[wk].fullPath;

    return null;
}

// ─────────────────────────────────────────────────────────────────────────────
//  READ EXCEL
// ─────────────────────────────────────────────────────────────────────────────
function readExcel(filePath) {
    var workbook = XLSX.readFile(filePath, { cellDates: true, raw: true });

    if (!workbook.SheetNames || !workbook.SheetNames.length) {
        throw new Error('Workbook has no sheets.');
    }

    var sheetName = workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: true,
        blankrows: true
    });

    return rows;
}

// ─────────────────────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────────────────────
function main() {
    console.log('Input:   ' + inputPath);
    console.log('Output:  ' + outputPath);
    console.log('Sources: ' + sourceFolder);
    console.log('Aliases: ' + aliasPath);

    // Load aliases and AI file index
    var aliases = readAliases(aliasPath);
    var index   = buildAiFileIndex(sourceFolder);
    console.log('AI files found: ' + index.files.length);
    console.log('Aliases loaded: ' + Object.keys(aliases).length);

    // Read Excel
    var rawRows = readExcel(inputPath);
    if (rawRows.length < 2) {
        throw new Error('Workbook does not contain data rows.');
    }

    // First non-empty row = header (we skip it)
    var headerRowIndex = 0;
    for (var h = 0; h < rawRows.length; h++) {
        var row = rawRows[h];
        if (row && row.length && row.some(function (cell) { return trim(cell) !== ''; })) {
            headerRowIndex = h;
            break;
        }
    }

    var headerRow  = rawRows[headerRowIndex] || [];
    var headerPreview = [];
    for (var hp = 0; hp < 4; hp++) {
        headerPreview.push(trim(headerRow[hp] || ''));
    }
    console.log('Headers A-D: ' + headerPreview.join(' | '));

    // Process data rows
    var items = [];
    var warnings = [];
    var outputNameByAddress = {};
    var usedOutputNames = {};
    var dataRows = rawRows.slice(headerRowIndex + 1);

    for (var r = 0; r < dataRows.length; r++) {
        var dataRow = dataRows[r];
        if (!dataRow) continue;

        var address = trim(dataRow[0]);
        var group   = trim(dataRow[1]);
        var product = trim(dataRow[2]);
        var total   = trim(dataRow[3]);
        var rowNumber = headerRowIndex + 1 + r + 1;   // 1-based, header is row 1

        // Skip fully empty rows
        if (!address && !group && !product && !total) continue;

        // Skip rows with zero total
        if (!isNonZeroTotal(total)) continue;

        // Validate required fields
        if (!address || !group || !product) {
            warnings.push('Row ' + rowNumber + ': skipped — address, group, or product is empty.');
            continue;
        }

        // Determine output filename for this address
        if (!outputNameByAddress[address]) {
            var baseName = safeFileName(address);
            var fileName = baseName + '.ai';
            var suffix = 2;
            while (usedOutputNames[fileName]) {
                fileName = baseName + '_' + suffix + '.ai';
                suffix++;
            }
            outputNameByAddress[address] = fileName;
            usedOutputNames[fileName] = address;
        }

        // Resolve source AI file
        var resolvedFile = resolveAiFile(group, aliases, index);

        items.push({
            address:        address,
            group:          group,
            product:        product,
            total:          total,
            sourceFile:     resolvedFile || null,
            sourceFileName: resolvedFile ? path.basename(resolvedFile) : null,
            outputFileName: outputNameByAddress[address],
            rowNumber:      rowNumber,
            copyIndex:      1
        });
    }

    // Build output JSON
    var data = {
        generatedAt: new Date().toISOString(),
        settings: {
            workbookPath:  inputPath,
            sourceFolder:  sourceFolder,
            outputFolder:  path.resolve(projectRoot, 'OUT'),
            maxRowWidthMm: 1000,
            gapMm:         20,
            safeMaxArtboardWidthMm:  205,
            safeMaxArtboardHeightMm: 115,
            repeatByItog: false
        },
        items: items
    };

    // Write JSON
    var json = JSON.stringify(data, null, 2);
    fs.writeFileSync(outputPath, json, 'utf-8');

    // Print summary
    var addressSet = {};
    var groupSet   = {};
    var missingCount = 0;
    for (var i = 0; i < items.length; i++) {
        addressSet[items[i].address] = true;
        groupSet[items[i].group] = true;
        if (!items[i].sourceFile) missingCount++;
    }

    console.log('');
    console.log('=== Summary ===');
    console.log('Items prepared:    ' + items.length);
    console.log('Unique addresses:  ' + Object.keys(addressSet).length);
    console.log('Unique groups:     ' + Object.keys(groupSet).length);
    console.log('Missing AI files:  ' + missingCount);

    if (warnings.length) {
        console.log('');
        console.log('Warnings (' + warnings.length + '):');
        for (var w = 0; w < warnings.length; w++) {
            console.log('  ' + warnings[w]);
        }
    }

    if (missingCount > 0) {
        console.log('');
        console.log('Missing source files:');
        for (var m = 0; m < items.length; m++) {
            if (!items[m].sourceFile) {
                console.log('  Row ' + items[m].rowNumber + ': group \'' +
                    items[m].group + '\', product \'' + items[m].product + '\'');
            }
        }
    }

    console.log('');
    console.log('Data written to: ' + outputPath);
}

try {
    main();
} catch (err) {
    console.error('FATAL: ' + (err.message || err));
    process.exit(1);
}
