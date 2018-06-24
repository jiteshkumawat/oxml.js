define(['oxml_content_types', 'oxml_rels', 'oxml_sheet', 'oxml_xlsx_styles'], function (oxmlContentTypes, oxmlRels, oxmlSheet, oxmlXlsxStyles) {
    'use strict';

    var getLastSheet = function (_workBook) {
        if (_workBook._rels.relations.length) {
            return _workBook._rels.relations[_workBook._rels.relations.length - 1];
        }
        return {};
    };

    var generateContent = function (_workBook, file) {
        // Create Workbood
        var index, workBook = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
        for (index = 0; index < _workBook.sheets.length; index++) {
            workBook += '<sheet name="' + _workBook.sheets[index]._sheet.sheetName + '" sheetId="' + _workBook.sheets[index]._sheet.sheetId +
                '" r:id="' + _workBook.sheets[index]._sheet.rId + '" />';
            if (file) {
                _workBook.sheets[index].attach(file);
            }
        }
        workBook += '</sheets></workbook>';
        return workBook;
    };

    var generateSharedStrings = function (_workBook) {
        if (_workBook._sharedStrings && Object.keys(_workBook._sharedStrings).length) {
            var sharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
            for (var key in _workBook._sharedStrings) {
                sharedStrings += '<si><t>' + key + '</t></si>';
            }
            sharedStrings += '</sst>';
            return sharedStrings;
        }
        return '';
    }

    var addSheet = function (_workBook, sheetName) {
        var lastSheetRel = getLastSheet(_workBook);
        var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
        sheetName = sheetName || "sheet" + nextSheetRelId;
        _workBook._rels.addRelation(
            "rId" + nextSheetRelId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "sheets/sheet" + nextSheetRelId + ".xml");
        // Update Sheets
        var sheet = oxmlSheet.createSheet(sheetName, nextSheetRelId, "rId" + nextSheetRelId, _workBook);
        _workBook.sheets.push(sheet);
        // Add Content Type
        _workBook.xlsxContentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", {
            PartName: "/workbook/sheets/sheet" + nextSheetRelId + ".xml"
        });

        return sheet;
    };

    var attach = function (file, _workBook) {
        // Add REL
        _workBook._rels.attach(file);
        var workBook = generateContent(_workBook, file);
        file.addFile(workBook, "workbook.xml", "workbook");
        var sharedStrings = generateSharedStrings(_workBook);
        if (sharedStrings) {
            file.addFile(sharedStrings, "sharedstrings.xml", "workbook");
        }
        if (_workBook._styles) {
            _workBook._styles.attach(file);
        }
    };

    var createSharedString = function (str, _workBook) {
        if (!str) {
            return;
        }
        if (!_workBook._sharedStrings) {
            _workBook._sharedStrings = {};

            // Add Content Types
            _workBook.xlsxContentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", {
                PartName: "/workbook/sharedstrings.xml"
            });

            // Add REL
            var lastSheetRel = getLastSheet(_workBook);
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
            _workBook._rels.addRelation("rId" + nextSheetRelId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "sharedstrings.xml");
        }

        if (isNaN(_workBook._sharedStrings[str])) {
            _workBook._sharedStrings[str] = Object.keys(_workBook._sharedStrings).length || 0;
        }

        return _workBook._sharedStrings[str];
    };

    var getSharedString = function (str, _workBook) {
        return _workBook._sharedStrings[str];
    };

    var destroy = function (_workBook) {
        var index;
        delete _workBook.xlsxContentTypes;
        delete _workBook.xlsxRels;
        _workBook._rels.destroy();
        _workBook._rels.destroy = null;
        delete _workBook._rels.destroy;
        _workBook._rels = null;
        delete _workBook._rels;
        _workBook.generateContent = null;
        delete _workBook.generateContent;
        _workBook.attach = null;
        delete _workBook.attach;
        _workBook.addSheet = null;
        delete _workBook.addSheet;
        _workBook.createSharedString = null;
        delete _workBook.createSharedString;
        _workBook.getSharedString = null;
        delete _workBook.getSharedString;
        _workBook._sharedStrings = null;
        delete _workBook._sharedStrings;
        for (index = 0; index < _workBook.sheets.length; index++) {
            _workBook.sheets[index].destroy();
            _workBook.sheets[index].destroy = null;
            delete _workBook.sheets[index].destroy;
            _workBook.sheets[index] = null;
        }
        _workBook.sheets = null;
        delete _workBook.sheets;
    };

    var createWorkbook = function (xlsxContentTypes, xlsxRels) {
        var _workBook = {
            sheets: [],
            xlsxContentTypes: xlsxContentTypes,
            xlsxRels: xlsxRels
        };
        _workBook._rels = oxmlRels.createRelation('workbook.xml.rels', "workbook/_rels");

        _workBook.addSheet = function (sheetName) {
            return addSheet(_workBook, sheetName);
        };
        _workBook.generateContent = function () {
            return generateContent(_workBook);
        };
        _workBook.attach = function (file) {
            attach(file, _workBook);
        };
        _workBook.createSharedString = function (str) {
            return createSharedString(str, _workBook);
        };
        _workBook.getSharedString = function (str) {
            return getSharedString(str, _workBook);
        };
        _workBook.destroy = function () {
            destroy(_workBook);
        };
        _workBook.createStyles = function () {
            if (!_workBook._styles)
                _workBook._styles = oxmlXlsxStyles.createStyle(_workBook, _workBook._rels, xlsxContentTypes);
            return _workBook._styles;
        };
        _workBook.getLastSheet = function () {
            return getLastSheet(_workBook);
        };
        return _workBook;
    };

    return { createWorkbook: createWorkbook };
});