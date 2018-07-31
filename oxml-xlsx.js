define(['fileHandler', 'oxml_content_types', 'oxml_rels', 'oxml_workbook'], function (fileHandler, oxmlContentTypes, oxmlRels, oxmlWorkBook) {
    'use strict';

    var oxml = {};

    var downloadFile = function (fileName, callback, _xlsx) {
        try {
            var jsZip, fs;
            if (typeof window === "undefined") {
                jsZip = _jsZip();
                fs = _fs;
            } 
            /* istanbul ignore next */
            else if (typeof JSZip === "undefined") {
                if (callback) {
                    callback('Err: JSZip reference not found.');
                } else if (typeof Promise !== "undefined") {
                    return new Promise(function (resolve, reject) {
                        reject("Err: JSZip reference not found.");
                    });
                }
            } 
            /* istanbul ignore next */
            else jsZip = new JSZip();

            var file = fileHandler.createFile(jsZip, fs);
            _xlsx.contentTypes.attach(file);
            _xlsx._rels.attach(file);
            _xlsx.workBook.attach(file);
            return file.saveFile(fileName, callback);
        } catch (err) {
            /* istanbul ignore next */
            if (callback) {
                callback('Err: Not able to create Workbook.');
            } 
            /* istanbul ignore next */
            else if (typeof Promise !== "undefined") {
                return new Promise(function (resolve, reject) {
                    reject("Err: Not able to create Workbook.");
                });
            }
        }
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

        var download = function (fileName, callback) {
            return downloadFile(fileName, callback, _xlsx);
        };

        return {
            _xlsx: _xlsx,
            sheet: _xlsx.workBook.addSheet,
            download: download
        };
    };

    oxml.xlsx = createXLSX;
    /* istanbul ignore next */
    if (typeof window !== "undefined" && !window.oxml) {
        window.oxml = oxml;
    }

    return oxml;
});