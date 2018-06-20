define(['fileHandler', 'oxml_content_types', 'oxml_rels', 'oxml_workbook'], function (fileHandler, oxmlContentTypes, oxmlRels, oxmlWorkBook) {
    'use strict';

    var oxml = {};

    var downloadFile = function (fileName, _xlsx) {
        var file = fileHandler.createFile();

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
        _xlsx.contentTypes = oxmlContentTypes.createContentType();

        _xlsx.contentTypes.addContentType("Default", "application/vnd.openxmlformats-package.relationships+xml", {
            Extension: "rels"
        });
        _xlsx.contentTypes.addContentType("Default", "application/xml", {
            Extension: "xml"
        });
        _xlsx.contentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", {
            PartName: "/workbook/workbook.xml"
        });

        _xlsx._rels = oxmlRels.createRelation('.rels', '_rels');
        _xlsx._rels.addRelation(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "workbook/workbook.xml"
        );

        _xlsx.workBook = oxmlWorkBook.createWorkbook(_xlsx.contentTypes, _xlsx._rels);

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
    
    oxml.createXLSX = createXLSX;

    if (!window.oxml) {
        window.oxml = oxml;
    }

    return oxml;
});