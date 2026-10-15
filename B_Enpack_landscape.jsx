#target illustrator

// Landscape launcher: 3000 x 2000 mm.
(function () {
    var primary = new File(File($.fileName).parent.fsName + '/B_Enpack.jsx');
    if (!primary.exists) throw new Error('B_Enpack.jsx not found: ' + primary.fsName);
    var hadPrevious = typeof $.global.__BEP_SHEET_ORIENTATION !== 'undefined';
    var previous = $.global.__BEP_SHEET_ORIENTATION;
    try {
        $.global.__BEP_SHEET_ORIENTATION = 'landscape';
        $.evalFile(primary);
    } finally {
        if (hadPrevious) $.global.__BEP_SHEET_ORIENTATION = previous;
        else delete $.global.__BEP_SHEET_ORIENTATION;
    }
}());
