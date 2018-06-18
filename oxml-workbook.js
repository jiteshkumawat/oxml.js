(function (window) {
    var getLastSheet = function (_workBook) {
        if (_workBook._rels.relations.length) {
            return _workBook._rels.relations[_workBook._rels.relations.length - 1];
        }
        return {};
    };

    var generateContent = function (_workBook, file) {
        // Create Workbood
        var workBook = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"> 
        <sheets>
        `;
        for (index = 0; index < _workBook.sheets.length; index++) {
            workBook += `<sheet name="${_workBook.sheets[index]._sheet.sheetName}" sheetId="${_workBook.sheets[index]._sheet.sheetId}" r:id="${_workBook.sheets[index]._sheet.rId}" /> 
            `;
            if (file) {
                _workBook.sheets[index].attach(file);
            }
        }
        workBook += `</sheets> 
        </workbook>`;
        return workBook;
    };

    var addSheet = function (_workBook, sheetName) {
        var lastSheetRel = getLastSheet(_workBook);
        var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
        sheetName = sheetName || "sheet" + nextSheetRelId;
        _workBook._rels.addRelation(
            "rId" + nextSheetRelId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "sheets/sheet" + nextSheetRelId + ".xml");
        // Update Sheets
        var sheet = oxml.createSheet(sheetName, nextSheetRelId, "rId" + nextSheetRelId);
        _workBook.sheets.push(sheet);
        return sheet;
    };

    var attach = function (file, _workBook) {
        // Add REL
        _workBook._rels.attach(file);
        var workBook = generateContent(_workBook, file);
        file.addFile(workBook, "workbook.xml", "workbook");
    };

    var createWorkbook = function () {
        var _workBook = {
            sheets: []
        };
        _workBook._rels = oxml.createRelation('workbook.xml.rels', "workbook/_rels");

        return {
            _workBook: _workBook,
            addSheet: function (sheetName) {
                return addSheet(_workBook, sheetName);
            },
            generateContent: function () {
                return generateContent(_workBook);
            },
            attach: function (file) {
                attach(file, _workBook);
            }
        };
    };

    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createWorkbook = createWorkbook;
})(window);