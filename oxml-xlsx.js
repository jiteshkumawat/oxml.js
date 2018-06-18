"use strict";

(function (window) {
    var downloadFile = function (fileName, _xlsx) {
        var file = window.oxml.util.createFile();

        // Attach Content Types
        _xlsx.contentTypes.attach(file);

        // Attach RELS
        _xlsx._rels.attach(file);

        // Attach WorkBook
        _xlsx.workBook.attach(file);

        file.saveFile(fileName);
    };
    
    var createXLSX = function () {
        var _xlsx = {};
        _xlsx.workBook = oxml.createWorkbook();
        _xlsx.contentTypes = oxml.createContentType();

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

        var download = function (fileName) {
            if (!JSZip) {
                return;
            }
            downloadFile(fileName, _xlsx);
        }

        return {
            _xlsx: _xlsx,
            addSheet: _xlsx.workBook.addSheet,
            download: download
        };
    };

    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createXLSX = createXLSX;
})(window);