"use strict";

(function (window) {
    var downloadFile = function (fileName, _xlsx) {
        var index = 0, file = window.oxml.util.createFile();

        // Create Content Types
        var contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        `;
        for (index = 0; index < _xlsx.ContentTypes.length; index++) {
            contentTypes += `<${_xlsx.ContentTypes[index].name} ContentType="${_xlsx.ContentTypes[index].ContentType}" `;
            if (_xlsx.ContentTypes[index].Extension) {
                contentTypes += `Extension="${_xlsx.ContentTypes[index].Extension}" `;
            }
            if (_xlsx.ContentTypes[index].PartName) {
                contentTypes += `PartName="${_xlsx.ContentTypes[index].PartName}" `;
            }
            contentTypes += '/>\n';
        }
        contentTypes += `</Types>`;
        file.addFile(contentTypes, "[Content_Types].xml");

        // Create RELS
        var rels = _xlsx._rels.generateContent();
        file.addFile(rels, ".rels", "_rels");

        // Create WorkBook RELS
        rels = _xlsx.workBook._rels.generateContent();
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
            ContentTypes: [{
                name: "Default",
                Extension: "rels",
                ContentType: "application/vnd.openxmlformats-package.relationships+xml",
            }, {
                name: "Default",
                Extension: "xml",
                ContentType: "application/xml",
            }, {
                name: "Override",
                PartName: "/workbook/workbook.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            }, {
                name: "Override",
                PartName: "/workbook/sheets/sheet1.xml",
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            }],
            workBook: {
                _rels: [],
                _workBook: {},
                sheets: []
            }
        };

        _xlsx._rels = oxml.createRelation('.rels');
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
            _xlsx.workBook._rels = oxml.createRelation('workbook.xml.rels');
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