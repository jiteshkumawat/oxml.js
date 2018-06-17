"use strict";

(function (window) {
    var downloadFile = function (fileName, _xlsx) {
        var index = 0, file = window.oxml.util.createFile();

        // Create Content Types
        _xlsx.contentTypes.attach(file);

        // Create RELS
        _xlsx._rels.attach(file);

        // Create WorkBook RELS
        var rels = _xlsx.workBook._rels.generateContent();
        file.addFile(rels, "workbook.xml.rels", "workbook/_rels");

        // Create Workbood
        var workbook = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"> 
        <sheets>
        `;
        for(index = 0; index < _xlsx.workBook.sheets.length; index++){
            workbook += `<sheet name="${_xlsx.workBook.sheets[index].name}" sheetId="${_xlsx.workBook.sheets[index].sheetId}" r:id="${_xlsx.workBook.sheets[index].rId}" /> 
            `;
            var sheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row>
                    <c>
                        <v>1234</v>
                    </c>
                </row>
            </sheetData>
        </worksheet>`;
        file.addFile(sheet, "sheet" + _xlsx.workBook.sheets[index].sheetId + ".xml", "workbook/sheets");
        }
        workbook += `</sheets> 
        </workbook>`;
        file.addFile(workbook, 'workbook.xml', 'workbook');
        file.saveFile(fileName);
    }
    var createXLSX = function () {
        var _xlsx = {
            contentTypes: oxml.createContentType(),
            workBook: {
                _rels: [],
                _workBook: {},
                sheets: []
            }
        };

        _xlsx.contentTypes.addContentType("Default", "application/vnd.openxmlformats-package.relationships+xml", {
            Extension: "rels"
        });
        _xlsx.contentTypes.addContentType("Default", "application/xml", {
            Extension: "xml"
        });
        _xlsx.contentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", {
            PartName: "/workbook/workbook.xml"
        });
        _xlsx.contentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", {
            PartName: "/workbook/sheets/sheet1.xml"
        });

        _xlsx._rels = oxml.createRelation('.rels', '_rels');
        _xlsx._rels.addRelation(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "workbook/workbook.xml"
        );

        var getLastSheet = function () {
            if (_xlsx.workBook._rels.length) {
                return _xlsx.workBook._rels[_xlsx.workBook._rels.length - 1];
            }
            return {};
        }

        var addSheet = function (sheetName) {
            var lastSheetRel = getLastSheet();
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;

            sheetName = sheetName || "sheet" + nextSheetRelId;
            // Update REL
            _xlsx.workBook._rels = oxml.createRelation('workbook.xml.rels', "workbook/_rels");
            _xlsx.workBook._rels.addRelation(
                "rId" + nextSheetRelId, 
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", 
                "sheets/sheet" + nextSheetRelId + ".xml");

            // Update Workbook
            // Update Sheets
            _xlsx.workBook.sheets.push({
                name: sheetName,
                sheetId: nextSheetRelId,
                rId: "rId" + nextSheetRelId
            });
        }

        var download = function (fileName) {
            if (!JSZip) {
                return;
            }
            downloadFile(fileName, _xlsx);
        }

        return {
            _xlsx: _xlsx,
            addSheet: addSheet,
            download: download
        };
    };

    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createXLSX = createXLSX;
})(window);