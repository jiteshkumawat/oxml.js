(function (root, factory) {
        /* istanbul ignore next */
        if (typeof define === 'function' && define.amd) { // AMD
                define([], factory);
        } 
        /* istanbul ignore next */
        else if (typeof define === 'function'){
                define(factory);
        } else if (typeof module !== 'undefined' && module.exports) { // Node.js
                var jsZip = require('jszip');
                var fs = require('fs');
                module.exports = factory(jsZip, fs);
        } 
        /* istanbul ignore next */
        else {
                root.oxml = factory();
        }
}(this, function (_jsZip, _fs) {
var requirejs,require,define;
/* istanbul ignore next */
!function(e){function n(e,n){return m.call(e,n)}function r(e,n){var r,i,t,o,u,f,l,c,s,p,a,d=n&&n.split("/"),h=g.map,m=h&&h["*"]||{};if(e){for(u=(e=e.split("/")).length-1,g.nodeIdCompat&&y.test(e[u])&&(e[u]=e[u].replace(y,"")),"."===e[0].charAt(0)&&d&&(e=d.slice(0,d.length-1).concat(e)),s=0;s<e.length;s++)if("."===(a=e[s]))e.splice(s,1),s-=1;else if(".."===a){if(0===s||1===s&&".."===e[2]||".."===e[s-1])continue;s>0&&(e.splice(s-1,2),s-=2)}e=e.join("/")}if((d||m)&&h){for(s=(r=e.split("/")).length;s>0;s-=1){if(i=r.slice(0,s).join("/"),d)for(p=d.length;p>0;p-=1)if((t=h[d.slice(0,p).join("/")])&&(t=t[i])){o=t,f=s;break}if(o)break;!l&&m&&m[i]&&(l=m[i],c=s)}!o&&l&&(o=l,f=c),o&&(r.splice(0,f,o),e=r.join("/"))}return e}function i(n,r){return function(){var i=v.call(arguments,0);return"string"!=typeof i[0]&&1===i.length&&i.push(null),c.apply(e,i.concat([n,r]))}}function t(e){return function(n){a[e]=n}}function o(r){if(n(d,r)){var i=d[r];delete d[r],h[r]=!0,l.apply(e,i)}if(!n(a,r)&&!n(h,r))throw new Error("No "+r);return a[r]}function u(e){var n,r=e?e.indexOf("!"):-1;return r>-1&&(n=e.substring(0,r),e=e.substring(r+1,e.length)),[n,e]}function f(e){return e?u(e):[]}var l,c,s,p,a={},d={},g={},h={},m=Object.prototype.hasOwnProperty,v=[].slice,y=/\.js$/;s=function(e,n){var i,t=u(e),f=t[0],l=n[1];return e=t[1],f&&(i=o(f=r(f,l))),f?e=i&&i.normalize?i.normalize(e,function(e){return function(n){return r(n,e)}}(l)):r(e,l):(f=(t=u(e=r(e,l)))[0],e=t[1],f&&(i=o(f))),{f:f?f+"!"+e:e,n:e,pr:f,p:i}},p={require:function(e){return i(e)},exports:function(e){var n=a[e];return void 0!==n?n:a[e]={}},module:function(e){return{id:e,uri:"",exports:a[e],config:function(e){return function(){return g&&g.config&&g.config[e]||{}}}(e)}}},l=function(r,u,l,c){var g,m,v,y,j,q,x,b=[],w=typeof l;if(c=c||r,q=f(c),"undefined"===w||"function"===w){for(u=!u.length&&l.length?["require","exports","module"]:u,j=0;j<u.length;j+=1)if(y=s(u[j],q),"require"===(m=y.f))b[j]=p.require(r);else if("exports"===m)b[j]=p.exports(r),x=!0;else if("module"===m)g=b[j]=p.module(r);else if(n(a,m)||n(d,m)||n(h,m))b[j]=o(m);else{if(!y.p)throw new Error(r+" missing "+m);y.p.load(y.n,i(c,!0),t(m),{}),b[j]=a[m]}v=l?l.apply(a[r],b):void 0,r&&(g&&g.exports!==e&&g.exports!==a[r]?a[r]=g.exports:v===e&&x||(a[r]=v))}else r&&(a[r]=l)},requirejs=require=c=function(n,r,i,t,u){if("string"==typeof n)return p[n]?p[n](r):o(s(n,f(r)).f);if(!n.splice){if((g=n).deps&&c(g.deps,g.callback),!r)return;r.splice?(n=r,r=i,i=null):n=e}return r=r||function(){},"function"==typeof i&&(i=t,t=u),t?l(e,n,r,i):setTimeout(function(){l(e,n,r,i)},4),c},c.config=function(e){return c(e)},requirejs._defined=a,(define=function(e,r,i){if("string"!=typeof e)throw new Error("See almond README: incorrect module build, no module name");r.splice||(i=r,r=[]),n(a,e)||n(d,e)||(d[e]=[e,r,i])}).amd={jQuery:!0}}();

define("build/almond", function(){});

define('fileHandler',[], function () {
    "use strict";

    var createCompressedFile = function (jsZip, fs) {

        var zip = jsZip;

        var compressedFile = {
            _zip: zip
        };

        compressedFile.addFile = function (content, fileName, filePath) {
            var path;
            if (filePath) {
                path = zip.folder(filePath);
            }
            (path || zip).file(fileName, content);
        };

        compressedFile.saveFile = function (fileName, callback) {
            /* istanbul ignore if  */
            if (typeof window !== "undefined") {
                return zip.generateAsync({ type: "blob" })
                    .then(function (content) {
                        try {
                            if (window.saveAs) {
                                return saveAs(content, fileName);
                            }
                            var url = window.URL.createObjectURL(content);
                            var element = document.createElement('a');
                            element.setAttribute('href', url);
                            element.setAttribute('download', fileName);

                            element.style.display = 'none';
                            document.body.appendChild(element);
                            element.click();
                            document.body.removeChild(element);

                            if (callback) {
                                callback(zip);
                            } else if (typeof Promise !== "undefined") {
                                return new Promise(function (resolve, reject) {
                                    resolve(zip);
                                });
                            }
                        } catch (err) {
                            if (callback) {
                                callback(null, "Err: Not able to create file object.");
                            } else if (typeof Promise !== "undefined") {
                                return new Promise(function (resolve, reject) {
                                    reject("Err: Not able to create file object.");
                                });
                            }
                        }
                    });
            } else {
                if(callback){
                    zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(fileName));
                    callback(zip);
                    return;
                }
                return new Promise(function (resolve, reject) {
                    zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(fileName));
                    resolve(zip);
                });
            }
        };

        return compressedFile;
    };

    return {
        createFile: createCompressedFile
    };
});
define('oxml_rels',[], function () {
    'use strict';

    // Add Relation
    var addRelation = function (id, type, target, _rels) {
        _rels.relations.push({
            Id: id,
            Type: type,
            Target: target
        });
    };

    var generateContent = function (_rels) {
        // Create RELS
        var index, rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for (index = 0; index < _rels.relations.length; index++) {
            var relation = _rels.relations[index];
            rels += '<Relationship Id="' + relation.Id + '" Type="' + relation.Type + '" Target="' + relation.Target + '"/>';
        }
        rels += '</Relationships>';
        return rels;
    };

    var attach = function (file, _rels) {
        var rels = generateContent(_rels);
        file.addFile(rels, _rels.fileName, _rels.folderName);
    };

    var destroy = function (_rels) {
        _rels.relations = null;
        delete _rels.relations;
        _rels.addRelation = null;
        delete _rels.addRelation;
        _rels.generateContent = null;
        delete _rels.generateContent;
        _rels.attach = null;
        delete _rels.attach;
    };

    // Create Relation
    var createRelation = function (fileName, folderName) {
        var _rels = {
            relations: [],
            fileName: fileName,
            folderName: folderName
        };
        _rels.addRelation = function (id, type, target) {
            addRelation(id, type, target, _rels);
        };
        _rels.generateContent = function () {
            return generateContent(_rels);
        };
        _rels.attach = function (file) {
            return attach(file, _rels);
        };
        _rels.destroy = function () {
            destroy(_rels);
        };
        return _rels;
    };

    return { createRelation: createRelation };
});
define('oxml_content_types',[], function () {
    'use strict';

    var generateContent = function (_contentType) {
        var index, contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        for (index = 0; index < _contentType.contentTypes.length; index++) {
            var contentType = _contentType.contentTypes[index];
            contentTypes += '<' + contentType.name + ' ContentType="' + contentType.contentType + '" ';
            if (contentType.Extension) {
                contentTypes += 'Extension="' + contentType.Extension + '" ';
            }
            if (contentType.PartName) {
                contentTypes += 'PartName="' + contentType.PartName + '" ';
            }
            contentTypes += '/>\n';
        }
        contentTypes += '</Types>';
        return contentTypes;
    };

    var attach = function (file, _contentType) {
        var contentTypes = generateContent(_contentType);
        file.addFile(contentTypes, "[Content_Types].xml");
    };

    var addContentType = function (name, contentType, attributes, _contentType) {
        var content = {
            name: name,
            contentType: contentType
        };

        for (var key in attributes) {
            content[key] = attributes[key];
        }
        _contentType.contentTypes.push(content);
    };

    var destroy = function (_contentType) {
        _contentType.addContentType = null;
        delete _contentType.addContentType;
        _contentType.generateContent = null;
        delete _contentType.generateContent;
        _contentType.attach = null;
        delete _contentType.attach;
        _contentType.contentTypes = null;
        delete _contentType.contentTypes;
    };

    var createContentType = function () {
        var _contentType = {
            contentTypes: []
        };
        _contentType.addContentType = function (name, contentType, attributes) {
            addContentType(name, contentType, attributes, _contentType);
        };
        _contentType.generateContent = function () {
            return generateContent(_contentType);
        };
        _contentType.attach = function (file) {
            return attach(file, _contentType);
        };
        _contentType.destroy = function () {
            destroy(_contentType);
        };
        return _contentType;
    };

    return {
        createContentType: createContentType
    };
});
define('oxml_table',[], function () {
    var generateContent = function (_table) {
        var tableContent = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table totalsRowShown="0" ref="' + _table.fromCell +
            ':' + _table.toCell + '" displayName="' + _table.displayName + '" name="' + _table.displayName +
            '" id="' + _table.rid + '" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        if (_table.filters) {
            tableContent += '<autoFilter ref="' + _table.fromCell +
                ':' + _table.toCell + '">';
            for (var index = 0; index < _table.filters.length; index++) {
                tableContent += '<filterColumn colId="' + _table.filters[index].column + '">';
                if (_table.filters[index].values[0].type === "default") {
                    tableContent += '<filters>';
                    for (var index2 = 0; index2 < _table.filters[index].values.length; index2++) {
                        tableContent += '<filter val="' + _table.filters[index].values[index2].value + '"/>';
                    }
                    tableContent += '</filters>';
                } else if (_table.filters[index].values[0].type === "custom") {
                    tableContent += '<customFilters and="' + (_table.filters[index].values[0].and ? '1' : '0') + '">';
                    for (var index2 = 0; index2 < _table.filters[index].values.length; index2++) {
                        tableContent += '<customFilter operator="' + _table.filters[index].values[index2].operator + '" val="' + _table.filters[index].values[index2].value + '"/>';
                    }
                    tableContent += '</customFilters>';
                }
                tableContent += '</filterColumn>';
            }
            tableContent += '</autoFilter>';
        }
        if (_table.sort) {
            tableContent += '<sortState caseSensitive="' + (_table.sort.caseSensitive ? '1' : '0') + '" ref="' + _table.sort.dataRange + '">';
            tableContent += '<sortCondition' + (_table.sort.direction === "ascending" ? "" : ' descending="1"') + ' ref="' + _table.sort.range + '"/>';
            tableContent += '</sortState>';
        }
        tableContent += '<tableColumns count="' + _table.columns.length + '">';
        for (var index = 0; index < _table.columns.length; index++)
            tableContent += '<tableColumn name="' + _table.columns[index] + '" id="' + (index + 1) + '"/>';
        var tableStyle = '';
        if (_table.tableStyle)
            tableStyle = '<tableStyleInfo name="' + _table.tableStyle.name
                + '" showColumnStripes="' + (_table.tableStyle.showColumnStripes ? '1' : '0')
                + '" showRowStripes="' + (_table.tableStyle.showRowStripes ? '1' : '0') + '" showLastColumn="'
                + (_table.tableStyle.showLastColumn ? '1' : '0') + '" showFirstColumn="'
                + (_table.tableStyle.showFirstColumn ? '1' : '0') + '"/>';
        tableContent += '</tableColumns>' + tableStyle + '</table>';
        return tableContent;
    };

    var attach = function (_table, file) {
        var content = generateContent(_table);
        file.addFile(content, "table" + _table.rid + ".xml", "workbook/tables");
    };

    var applyFilter = function (options, _table, _sheet) {
        if (!_table.filters) _table.filters = [];
        if (typeof options.filters === "object" && options.filters.length) {
            for (var index = 0; index < options.filters.length; index++) {
                var values;
                if (options.filters[index].value) {
                    values = [];
                    values.push({
                        value: options.filters[index].value,
                        type: options.filters[index].type || "default",
                        operator: options.filters[index].operator || null,
                        and: options.filters[index].and !== false
                    });
                    hideRows(_sheet, options.filters[index].column - 1, [options.filters[index].value], _table.fromCell, _table.toCell, options.filters[index].operator);
                } else if (options.filters[index].values && typeof options.filters[index].values === "object" && options.filters[index].values.length) {
                    values = [];
                    for (var index2 = 0; index2 < options.filters[index].values.length; index2++) {
                        values.push({
                            value: options.filters[index].values[index2],
                            type: options.filters[index].type || "default",
                            operator: options.filters[index].operator || null,
                            and: options.filters[index].and !== false
                        });
                    }
                    hideRows(_sheet, options.filters[index].column - 1, options.filters[index].values, _table.fromCell, _table.toCell, options.filters[index].operator);
                }
                _table.filters.push({
                    column: options.filters[index].column - 1,
                    values: values
                });
            }
        }
    };

    var applySort = function (options, _table) {
        var sort = {};
        var columnFrom = _table.fromColumn - 1;
        var columnTo = _table.toColumn - 1;
        var rowFrom = _table.fromRow;
        var rowTo = _table.toRow;
        if (typeof options.sort === "number") {
            columnFrom = String.fromCharCode(options.sort + columnFrom + 64);
            sort.direction = "ascending";
            sort.caseSensitive = true;
        } else if (typeof options.sort === "object" && !options.sort.length && options.sort.column) {
            columnFrom = String.fromCharCode(options.sort.column + columnFrom + 64);
            sort.direction = options.sort.direction || "ascending";
            sort.caseSensitive = options.sort.caseSensitive !== undefined ? options.sort.caseSensitive : true;
        }
        sort.range = columnFrom + (rowFrom + 1) + ":" + columnFrom + rowTo;
        sort.dataRange = String.fromCharCode(_table.fromColumn + 64) + (rowFrom + 1) + ":" + String.fromCharCode(columnTo + 65) + rowTo;
        _table.sort = sort;
    };

    var applyOptions = function (options, _table, _sheet) {
        if (options) {
            if (options.filters) {
                applyFilter(options, _table, _sheet);
            }
            if (options.sort) {
                applySort(options, _table);
            }
        }
    };

    var hideRows = function (_sheet, column, values, fromCell, toCell, operator) {
        var rowFrom = parseInt(fromCell.match(/\d+/)[0], 10);
        var rowTo = parseInt(toCell.match(/\d+/)[0], 10);
        var columnFrom = fromCell.match(/\D+/)[0].toUpperCase().charCodeAt() - 65;
        column += columnFrom;
        for (var index = rowFrom; index < rowTo; index++) {
            var isHidden = false;
            if (_sheet.values[index]) {
                if (operator && (operator === "greaterThan" || operator === "greaterThanOrEqual" || operator === "lessThan" || operator === "lessThanOrEqual")) {
                    isHidden = true;
                    for (var index2 = 0; index2 < values.length; index2++) {
                        switch (operator) {
                            case "greaterThan":
                                if (_sheet.values[index][column].value > values[index2])
                                    isHidden = false;
                                break;
                            case "greaterThanOrEqual":
                                if (_sheet.values[index][column].value >= values[index2])
                                    isHidden = false;
                                break;
                            case "lessThan":
                                if (_sheet.values[index][column].value >= values[index2])
                                    isHidden = false;
                                break;
                            case "lessThanOrEqual":
                                if (_sheet.values[index][column].value >= values[index2])
                                    isHidden = false;
                                break;
                        };
                        if (!isHidden) break;
                    }
                } else if (operator === "notEqual") {
                    isHidden = values.indexOf(_sheet.values[index][column].value) > -1;
                } else {
                    isHidden = !(values.indexOf(_sheet.values[index][column].value) > -1);
                }
            }
            if (isHidden) {
                _sheet.values[index].hidden = true;
            }
        }
    };

    var style = function (options, tableStyleName, _sheet, _table) {
        var _styles = _sheet._workBook.createStyles();
        _styles.addTableStyle(options, tableStyleName);
        _table.tableStyle = {
            name: tableStyleName,
            showColumnStripes: !!(options.evenColumn || options.oddColumn),
            showRowStripes: !!(options.evenRow || options.oddRow),
            showLastColumn: !!options.lastColumn,
            showFirstColumn: !!options.firstColumn
        };
    };

    var tableOptions = function (_table, _sheet, relId) {
        return {
            _table: _table,
            set: function (options) {
                applyOptions(options, _table, _sheet);
                return tableOptions(_table, _sheet, relId);
            },
            style: function (options, tableStyleName) {
                if (!tableStyleName) tableStyleName = 'customTableStyle' + relId;
                style(options, tableStyleName, _sheet, _table);
                return tableOptions(_table, _sheet, relId);
            }
        };
    };

    var addTable = function (displayName, fromCell, toCell, columns, relId, options, _sheet) {
        var rowFrom = parseInt(fromCell.match(/\d+/)[0], 10);
        var rowTo = parseInt(toCell.match(/\d+/)[0], 10);
        var columnFrom = fromCell.match(/\D+/)[0].toUpperCase().charCodeAt() - 64;
        var columnTo = toCell.match(/\D+/)[0].toUpperCase().charCodeAt() - 64;
        var _table = {
            rid: relId,
            displayName: displayName,
            columns: columns,
            fromCell: fromCell,
            toCell: toCell,
            fromRow: rowFrom,
            toRow: rowTo,
            fromColumn: columnFrom,
            toColumn: columnTo
        };
        applyOptions(options, _table, _sheet);
        return {
            _table: _table,
            rid: relId,
            generateContent: function () {
                generateContent(_table);
            },
            attach: function (file) {
                attach(_table, file);
            },
            tableOptions: function () {
                return tableOptions(_table, _sheet, relId);
            }
        };
    };

    return { addTable: addTable };
});
define('oxml_sheet',['oxml_table', 'oxml_rels'], function (oxmlTable, oxmlRels) {
    'use strict';

    var cellString = function (value, cellIndex, rowIndex, columnIndex) {
        if (!value) {
            return '';
        }
        var styleString = value.styleIndex ? ' s="' + value.styleIndex + '" ' : '';
        if (value.type === 'numeric') {
            return '<c r="' + cellIndex + '" ' + styleString + '><v>' + value.value + '</v></c>';
        } else if (value.type === 'sharedString') {
            return '<c r="' + cellIndex + '" t="s" ' + styleString + '><v>' + value.value + '</v></c>';
        } else if (value.type === "sharedFormula") {
            var v = "";
            if (value.value && typeof value.value === "function") {
                v = '<v>' + value.value(rowIndex + 1, columnIndex + 1) + '</v>';
            } else if (value.value !== undefined && value.value !== null) {
                v = '<v>' + value.value + '</v>';
            }
            if (value.formula) {
                return '<c  r="' + cellIndex + '" ' + styleString + '><f t="shared" ref="' + value.range + '" si="' + value.si + '">' + value.formula + '</f>' + v + '</c>';
            }
            return '<c  r="' + cellIndex + '" ' + styleString + '><f t="shared" si="' + value.si + '"></f>' + v + '</c>';
        } else if (value.type === 'string') {
            return '<c r="' + cellIndex + '" t="inlineStr" ' + styleString + '><is><t>' + value.value + '</t></is></c>';
        } else if (value.type === 'formula') {
            var v = "";
            if (value.value && typeof value.value === "function") {
                v = '<v>' + value.value(rowIndex + 1, columnIndex + 1) + '</v>';
            } else if (value.value !== undefined && value.value !== null) {
                v = '<v>' + value.value + '</v>';
            }
            return '<c r="' + cellIndex + '" ' + styleString + '><f>' + value.formula + '</f>' + v + '</c>';
        }
    };

    var generateContent = function (_sheet, file) {
        var rowIndex = 0, sheetValues = '';
        if (_sheet.values && _sheet.values.length) {
            for (rowIndex = 0; rowIndex < _sheet.values.length; rowIndex++) {
                if (_sheet.values[rowIndex] && _sheet.values[rowIndex].length > 0) {
                    var hidden = _sheet.values[rowIndex].hidden ? ' hidden="1" ' : '';
                    sheetValues += '<row r="' + (rowIndex + 1) + '"' + hidden + '>\n';

                    var columnIndex = 0;
                    for (columnIndex = 0; columnIndex < _sheet.values[rowIndex].length; columnIndex++) {
                        var columnChar = String.fromCharCode(65 + columnIndex);
                        var value = _sheet.values[rowIndex][columnIndex];
                        var cellIndex = columnChar + (rowIndex + 1);
                        var cellStr = cellString(value, cellIndex, rowIndex, columnIndex);
                        sheetValues += cellStr;
                    }

                    sheetValues += '</row>';
                }
            }
        }
        var tables = '';
        if (_sheet.tables && _sheet.tables.length) {
            tables += '<tableParts count="' + _sheet.tables.length + '">';
            for (var index = 0; index < _sheet.tables.length; index++) {
                tables += '<tablePart r:id="rId' + _sheet.tables[index].rid + '"/>';
                if (file) _sheet.tables[index].attach(file);
            }
            tables += '</tableParts>';
        }
        var sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>' +
            sheetValues + '</sheetData>' + tables + '</worksheet>';
        return sheet;
    };

    var attach = function (_sheet, file) {
        var content = generateContent(_sheet, file);
        file.addFile(content, "sheet" + _sheet.sheetId + ".xml", "workbook/sheets");
        if (_sheet._rels) {
            _sheet._rels.attach(file);
        }
    };

    var sanitizeValue = function (value, _sheet) {
        if (value === undefined || value === null) {
            return null;
        }
        if (typeof value === "object") {
            // Object type of value
            if ((value.value === undefined || value.value === null) && !(value.formula && value.type === "formula")) {
                return null;
            }
            if (!value.type) {
                value = value.value;
            } else {
                if (value.type === "sharedString") {
                    if (typeof value.value === "number") {
                        return {
                            type: value.type,
                            value: value.value
                        };
                    } else {
                        // Add shared string
                        value.value = _sheet._workBook.createSharedString(value.value, _sheet._workBook);
                        return {
                            type: value.type,
                            value: value.value
                        };
                    }
                } else if (value.type === "formula") {
                    return {
                        type: value.type,
                        formula: value.formula,
                        value: value.value
                    };
                } else if (value.type === "numeric" || value.type === "string") {
                    return {
                        type: value.type,
                        value: value.value
                    };
                }
                return null;
            }
        }

        if (typeof value === "number") {
            return {
                type: "numeric",
                value: value
            };
        } else if (typeof value === "string") {
            return {
                type: "string",
                value: value
            };
        } else if (typeof value === "boolean") {
            return {
                type: "string",
                value: value + ""
            };
        }
        return null;
    };

    var updateValuesInMatrix = function (values, _sheet, rowIndex, columnIndex, options, cellIndices) {
        var index, styleIndex;
        if (options) {
            options.cellIndices = cellIndices;
            var _styles = _sheet._workBook.createStyles();
            styleIndex = _styles.addStyles(options);
        }
        for (index = 0; index < values.length; index++) {
            if (values[index] !== undefined || values[index] !== null) {
                var sheetRowIndex = index + rowIndex - 1;
                if (!_sheet.values[sheetRowIndex]) {
                    _sheet.values[sheetRowIndex] = [];
                }
                if (typeof values[index] === "object" && values[index].length >= 0) {
                    var index2;
                    for (index2 = 0; index2 < values[index].length; index2++) {
                        var value = sanitizeValue(values[index][index2], _sheet);
                        if (value) {
                            if (styleIndex) {
                                value.styleIndex = styleIndex.index;
                            }
                            if (styleIndex || !_sheet.values[sheetRowIndex][index2 + columnIndex - 1])
                                _sheet.values[sheetRowIndex][index2 + columnIndex - 1] = value;
                            else {
                                _sheet.values[sheetRowIndex][index2 + columnIndex - 1].value = value.value;
                                _sheet.values[sheetRowIndex][index2 + columnIndex - 1].type = value.type;
                            }
                        }
                    }
                } else {
                    var value = sanitizeValue(values[index], _sheet);
                    if (value) {
                        if (styleIndex) {
                            value.styleIndex = styleIndex.index;
                        }
                        if (styleIndex || !_sheet.values[sheetRowIndex][columnIndex - 1])
                            _sheet.values[sheetRowIndex][columnIndex - 1] = value;
                        else {
                            _sheet.values[sheetRowIndex][columnIndex - 1].value = value.value;
                            _sheet.values[sheetRowIndex][columnIndex - 1].type = value.type;
                        }
                    }
                }
            }
        }
    };

    var updateValueInCell = function (value, _sheet, rowIndex, columnIndex, options) {
        var cellIndex = String.fromCharCode(65 + columnIndex - 1) + rowIndex;
        var sheetRowIndex = rowIndex - 1;
        if (value !== undefined && value !== null) {
            if (!_sheet.values[sheetRowIndex]) {
                _sheet.values[sheetRowIndex] = [];
            }
            value = sanitizeValue(value, _sheet);
            if (options) {
                options.cellIndex = cellIndex;
                var _styles = _sheet._workBook.createStyles();
                value.styleIndex = _styles.addStyles(options).index;
            }
            if (value && !value.styleIndex && _sheet.values[sheetRowIndex][columnIndex - 1])
                value.styleIndex = _sheet.values[sheetRowIndex][columnIndex - 1].styleIndex;
            _sheet.values[sheetRowIndex][columnIndex - 1] = value;
        }
        return getCellAttributes(_sheet, cellIndex, sheetRowIndex, columnIndex - 1);
    };

    var getCellAttributes = function (_sheet, cellIndex, rowIndex, columnIndex) {
        return {
            style: function (options) {
                options.cellIndex = cellIndex;
                updateSingleStyle(_sheet, options, rowIndex, columnIndex);
                return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
            },
            value: _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null,
            type: _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null,
            cellIndex: cellIndex,
            rowIndex: rowIndex + 1,
            columnIndex: String.fromCharCode(65 + columnIndex),
            formula: _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].formula : undefined,
            set: function (value, options) {
                cell(_sheet, rowIndex + 1, columnIndex + 1, value, options);
                this.value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
                this.formula = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].formula : undefined;
                this.type = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null;
                return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
            }
        };
    };

    var getCellRangeAttributes = function (_sheet, cellIndices, _cells, cRowIndex, cColumnIndex, totalRows, totalColumns, isRow, isColumn) {
        var cellRange = [], index;
        for (index = 0; index < _cells.length; index++) {
            var rowIndex = _cells[index].rowIndex, columnIndex = _cells[index].columnIndex;
            var value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
            var formula = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].formula : undefined;
            var type = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null;
            var cellIndex = String.fromCharCode(65 + columnIndex) + (rowIndex + 1);
            var sharedFormulaIndex = type === "sharedFormula" ? _sheet.values[rowIndex][columnIndex].si : undefined;
            cellRange.push({
                style: function (options) {
                    options.cellIndex = cellIndex;
                    updateSingleStyle(_sheet, options, rowIndex, columnIndex);
                    return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
                },
                rowIndex: rowIndex + 1,
                columnIndex: String.fromCharCode(65 + columnIndex),
                value: value,
                formula: formula,
                cellIndex: cellIndex,
                sharedFormulaIndex: sharedFormulaIndex,
                type: type,
                set: function (value, options) {
                    cell(_sheet, rowIndex + 1, columnIndex + 1, value, options);
                    this.value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
                    this.formula = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].formula : undefined;
                    this.type = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null;
                    return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
                }
            });
        }
        return {
            style: function (options) {
                options.cellIndices = cellIndices;
                updateRangeStyle(_sheet, options, _cells);
                return getCellRangeAttributes(_sheet, cellIndices, _cells, cRowIndex, cColumnIndex, totalRows, totalColumns);
            },
            cellIndices: cellIndices,
            cells: cellRange,
            set: function (values, options) {
                var tVal = [];
                if (!values || !values.length || !totalColumns || !totalRows)
                    return getCellRangeAttributes(_sheet, cellIndices, _cells, cRowIndex, cColumnIndex, totalRows, totalColumns);
                if (isRow || isColumn) {
                    for (var index = 0; index < this.cells.length && index < values.length; index++) {
                        tVal.push(values[index]);
                    }
                }
                if (isRow) {
                    values = [tVal];
                } else if (isColumn) {
                    values = tVal;
                } else {
                    var tVal = [];
                    for (var index = 0; index < values.length && index < totalRows; index++) {
                        if (values[index]) {
                            var tVal2 = [];
                            if (values[index].length > totalColumns) {
                                for (var index2 = 0; index2 < values[index].length && index2 < totalColumns; index2++) {
                                    tVal2.push(values[index][index2]);
                                }
                                values[index] = tVal2;
                            }
                            tVal.push(values[index]);
                        } else tVal.push([]);
                    }
                    values = tVal;
                }
                cells(_sheet, cRowIndex, cColumnIndex, totalRows, totalColumns, values, options, false, isRow, isColumn);
                var cellsAttributes = getCellRangeAttributes(_sheet, cellIndices, _cells, cRowIndex, cColumnIndex, totalRows, totalColumns);
                this.cells = cellsAttributes.cells;
                return cellsAttributes;
            }
        };
    };

    var updateSingleStyle = function (_sheet, options, rowIndex, columnIndex) {
        var _styles = _sheet._workBook.createStyles();
        var styleIndex = _styles.addStyles(options).index;
        if (!_sheet.values[rowIndex]) {
            _sheet.values[rowIndex] = [];
        }
        if (!_sheet.values[rowIndex][columnIndex]) {
            _sheet.values[rowIndex][columnIndex] = {
                value: '',
                type: 'string'
            };
        }
        _sheet.values[rowIndex][columnIndex].styleIndex = styleIndex;
    };

    var updateRangeStyle = function (_sheet, options, cells) {
        var _styles = _sheet._workBook.createStyles();
        var styleIndex = _styles.addStyles(options).index;
        for (var cellIndex = 0; cellIndex < cells.length; cellIndex++) {
            if (!_sheet.values[cells[cellIndex].rowIndex]) {
                _sheet.values[cells[cellIndex].rowIndex] = [];
            }
            if (!_sheet.values[cells[cellIndex].rowIndex][cells[cellIndex].columnIndex]) {
                _sheet.values[cells[cellIndex].rowIndex][cells[cellIndex].columnIndex] = {
                    type: "string",
                    value: ""
                };
            }
            _sheet.values[cells[cellIndex].rowIndex][cells[cellIndex].columnIndex].styleIndex = styleIndex;
        }
    };

    var updateSharedFormula = function (_sheet, formula, fromCell, toCell, options) {
        var nextId, cellIndices = [], cells = [];

        if (!fromCell || !formula || !toCell) {
            return;
        }

        var val;
        if (typeof formula === "object" && !formula.length) {
            if (!formula.formula) return;
            val = formula.value;
            formula = formula.formula;
        }

        var fromCellChar = fromCell.match(/\D+/)[0];
        var fromCellNum = fromCell.match(/\d+/)[0];

        if (!_sheet._sharedFormula) {
            _sheet._sharedFormula = [];
            nextId = 0;
        } else {
            var lastSharedFormula = _sheet._sharedFormula[_sheet._sharedFormula.length - 1];
            nextId = lastSharedFormula.si + 1;
        }

        _sheet._sharedFormula.push({
            si: nextId,
            fromCell: fromCell,
            toCell: toCell,
            formula: formula
        });

        // Update from Cell
        var columIndex = fromCellChar.toUpperCase().charCodeAt() - 65;
        var rowIndex = parseInt(fromCellNum, 10);
        if (!_sheet.values[rowIndex - 1]) {
            _sheet.values[rowIndex - 1] = [];
        }
        _sheet.values[rowIndex - 1][columIndex] = {
            type: "sharedFormula",
            si: nextId,
            formula: formula,
            range: fromCell + ":" + toCell,
            value: val
        };
        cellIndices.push(String.fromCharCode(65 + columIndex) + rowIndex);
        cells.push({ rowIndex: rowIndex - 1, columnIndex: columIndex });

        var toCellChar, toCellNum;
        if (toCell) {
            toCellChar = toCell.match(/\D+/)[0];
            toCellNum = toCell.match(/\d+/)[0];
        }

        // Update all cell in row
        if (toCellNum === fromCellNum) {
            var toColumnIndex = toCellChar.toUpperCase().charCodeAt() - 65;
            for (columIndex++; columIndex <= toColumnIndex; columIndex++) {
                _sheet.values[rowIndex - 1][columIndex] = {
                    type: "sharedFormula",
                    si: nextId,
                    value: val
                };
                cellIndices.push(String.fromCharCode(65 + columIndex) + rowIndex);
                cells.push({ rowIndex: rowIndex - 1, columnIndex: columIndex });
            }
        } else if (toCellChar === fromCellChar) {
            var torowIndex = parseInt(toCellNum, 10);
            for (rowIndex++; rowIndex <= torowIndex; rowIndex++) {
                if (!_sheet.values[rowIndex - 1]) {
                    _sheet.values[rowIndex - 1] = [];
                }
                cellIndices.push(String.fromCharCode(65 + columIndex) + rowIndex);
                cells.push({ rowIndex: rowIndex - 1, columnIndex: columIndex });

                _sheet.values[rowIndex - 1][columIndex] = {
                    type: "sharedFormula",
                    si: nextId,
                    value: val
                };
            }
        }
        if (options) {
            options.cellIndices = cellIndices;
            var styleIndex, _styles = _sheet._workBook.createStyles();
            styleIndex = _styles.addStyles(options);
            for (var cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                _sheet.values[cells[cellIndex].rowIndex][cells[cellIndex].columnIndex].styleIndex = styleIndex.index;
            }
        }
        return getCellRangeAttributes(_sheet, cellIndices, cells);
    };

    var destroy = function (_sheet) {
        delete _sheet.sheetName;
        delete _sheet.sheetId;
        delete _sheet.rId;
        delete _sheet.values;
        delete _sheet._workBook;
        delete _sheet._sharedFormula;
    };

    var cell = function (_sheet, rowIndex, columnIndex, value, options) {
        if (!rowIndex || !columnIndex || typeof rowIndex !== "number" || typeof columnIndex !== "number")
            return;
        if (!options && (value === undefined || value === null)) {
            var cellIndex = String.fromCharCode(65 + columnIndex - 1) + rowIndex;
            return getCellAttributes(_sheet, cellIndex, rowIndex - 1, columnIndex - 1);
        } else if (!options && typeof value === "object" && !value.type && (value.value === undefined || value.value === null)) {
            return cellStyle(_sheet, rowIndex, columnIndex, value);
        } else if (value === undefined || value === null) {
            return cellStyle(_sheet, rowIndex, columnIndex, options);
        } else {
            return updateValueInCell(value, _sheet, rowIndex, columnIndex, options);
        }
    };

    var cells = function (_sheet, rowIndex, columnIndex, totalRows, totalColumns, values, options, isReturn, isRow, isColumn) {
        if (!rowIndex || !columnIndex || typeof rowIndex !== "number" || typeof columnIndex !== "number" || typeof totalRows !== "number" || typeof totalColumns !== "number")
            return;
        var cells = [], cellIndices = [], tmpRows = totalRows;
        for (var index = rowIndex - 1; tmpRows > 0; index++ , tmpRows--) {
            var tmpColumns = totalColumns;
            if (!_sheet.values[index])
                _sheet.values[index] = [];
            for (var index2 = columnIndex - 1; tmpColumns > 0; index2++ , tmpColumns--) {
                cellIndices.push(String.fromCharCode(65 + index2) + (index + 1));
                cells.push({ rowIndex: index, columnIndex: index2 });
            }
        }
        if (!options && !values) {
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow, isColumn);
        } else if (!options && values && !values.length) {
            values.cellIndices = cellIndices;
            updateRangeStyle(_sheet, values, cells);
            if (!isReturn) return;
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow, isColumn);
        } else if (values === undefined || values === null) {
            options.cellIndices = cellIndices;
            updateRangeStyle(_sheet, options, cells);
            if (!isReturn) return;
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow, isColumn);
        } else {
            updateValuesInMatrix(values, _sheet, rowIndex, columnIndex, options, cellIndices, cells, totalRows, totalColumns);
            if (!isReturn) return;
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow, isColumn);
        }
    };

    var cellStyle = function (_sheet, rowIndex, columnIndex, options) {
        var cellIndex = String.fromCharCode(65 + columnIndex - 1) + rowIndex;
        options.cellIndex = cellIndex;
        updateSingleStyle(_sheet, options, rowIndex - 1, columnIndex - 1);
        return getCellAttributes(_sheet, cellIndex, rowIndex - 1, columnIndex - 1);
    };

    var addTable = function (_sheet, xlsxContentTypes, tableName, fromCell, toCell, options) {
        var titles = [];
        var fromColumIndex = fromCell.match(/\D+/)[0].toUpperCase().charCodeAt() - 65;
        var fromRowIndex = parseInt(fromCell.match(/\d+/)[0], 10);
        var toColumnIndex = toCell.match(/\D+/)[0].toUpperCase().charCodeAt() - 65;
        var titleRow = _sheet.values[fromRowIndex - 1];
        for (var index = fromColumIndex; index <= toColumnIndex; index++) {
            titles.push(titleRow[index].value || '');
        }

        if (!_sheet._rels) {
            _sheet._rels = oxmlRels.createRelation("sheet" + _sheet.sheetId + ".xml.rels", "workbook/sheets/_rels");
        }
        var relId = getNextRelation(_sheet);
        _sheet._rels.addRelation("rId" + relId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            "../tables/table" + relId + ".xml");

        xlsxContentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", {
            PartName: "/workbook/tables/table" + relId + ".xml"
        });

        if (!_sheet.tables) _sheet.tables = [];
        var table = oxmlTable.addTable(tableName, fromCell, toCell, titles, relId, options, _sheet);
        _sheet.tables.push(table);
        return table;
    };

    var getNextRelation = function (_sheet) {
        var lastSheetRel = {};
        if (_sheet._rels.relations.length) {
            lastSheetRel = _sheet._rels.relations[_sheet._rels.relations.length - 1];
        }
        var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
        return nextSheetRelId;
    };

    var createSheet = function (sheetName, sheetId, rId, workBook, xlsxContentTypes) {
        var _sheet = {
            sheetName: sheetName,
            sheetId: sheetId,
            rId: rId,
            values: [],
            _workBook: workBook
        };
        return {
            _sheet: _sheet,
            generateContent: function () {
                generateContent(_sheet);
            },
            attach: function (file) {
                attach(_sheet, file);
            },
            sharedFormula: function (fromCell, toCell, formula, options) {
                return updateSharedFormula(_sheet, formula, fromCell, toCell, options);
            },
            cell: function (rowIndex, columnIndex, value, options) {
                return cell(_sheet, rowIndex, columnIndex, value, options);
            },
            row: function (rowIndex, columnIndex, values, options) {
                var totalColumns = _sheet.values[rowIndex - 1] ? _sheet.values[rowIndex - 1].length - columnIndex + 1 : 0;
                return cells(_sheet, rowIndex, columnIndex, 1, values ? values.length : totalColumns, values ? [values] : null, options, true, true, false);
            },
            column: function (rowIndex, columnIndex, values, options) {
                var totalRows = 0;
                if (!values || !values.length) {
                    totalRows = _sheet.values && _sheet.values.length && _sheet.values.length > rowIndex ? _sheet.values.length - rowIndex + 1 : 0;
                }
                return cells(_sheet, rowIndex, columnIndex, values ? values.length : totalRows, 1, values, options, true, false, true);
            },
            grid: function (rowIndex, columnIndex, values, options) {
                var index, totalRows = 0, totalColumns = 0;
                if (values && values.length) {
                    totalRows = values.length;
                    for (index = 0; index < values.length; index++) {
                        if (values[index] && values[index].length)
                            totalColumns = totalColumns < values[index].length ? values[index].length : totalColumns;
                    }
                } else if (_sheet.values && _sheet.values.length) {
                    totalRows = _sheet.values.length - rowIndex + 1;
                    for (var index = 0; index < _sheet.values.length; index++) {
                        if (_sheet.values[index] && _sheet.values[index].length) {
                            totalColumns = totalColumns < _sheet.values[index].length - columnIndex + 1 ? _sheet.values[index].length - columnIndex + 1 : totalColumns;
                        }
                    }
                }

                return cells(_sheet, rowIndex, columnIndex, totalRows, totalColumns, values, options, true, false);
            },
            table: function (tableName, fromCell, toCell, options) {
                var _table = addTable(_sheet, xlsxContentTypes, tableName, fromCell, toCell, options);
                return _table.tableOptions();
            },
            destroy: function () {
                return destroy(_sheet);
            }
        };
    };

    return { createSheet: createSheet };
});
define('utils',[], function () {
    sortObject = function (obj) {
        if (typeof obj !== 'object' || obj.length)
            return obj;
        var temp = {};
        var keys = [];
        for (var key in obj)
            keys.push(key);
        keys.sort();
        for (var index in keys)
            temp[keys[index]] = sortObject(obj[keys[index]]);
        return temp;
    };

    var stringify = function (obj) {
        return JSON.stringify(sortObject(obj));
    };

    return {
        stringify: stringify
    };
});
define('oxml_xlsx_font',['utils'], function (utils) {
    var generateContent = function (_styles) {
        var stylesString = '', fontKey;
        stylesString += '<fonts count="' + _styles._fontsCount + '">';
        for (fontKey in _styles._fonts) {
            var font = JSON.parse(fontKey);
            stylesString += generateSingleContent(font);
        }
        stylesString += '</fonts>';
        return stylesString;
    };

    var generateSingleContent = function (font) {
        var stylesString = '<font>';
        stylesString += font.strike ? '<strike/>' : '';
        stylesString += font.italic ? '<i/>' : '';
        stylesString += font.bold ? '<b/>' : '';
        stylesString += font.underline ? '<u/>' : '';
        stylesString += font.size ? '<sz val="' + font.size + '"/>' : '';
        stylesString += font.color ? '<color rgb="'
            + font.color + '"/>' : '';
        stylesString += font.name ? '<name val="' + font.name + '"/>' : '';
        stylesString += font.family ? '<family val="'
            + font.family + '"/>' : '';
        stylesString += font.scheme ? '<scheme val="'
            + font.scheme + '"/>' : '';
        stylesString += '</font>';
        return stylesString;
    };

    var createFont = function (options) {
        var font = {};
        font.bold = !!options.bold;
        font.italic = !!options.italic;
        font.underline = !!options.underline;
        font.size = options.fontSize || false;
        font.color = options.fontColor || false;
        font.name = options.fontName || false;
        font.family = options.fontFamily || false;
        font.scheme = options.scheme || false;
        font.strike = !!options.strike;
        return font;
    };

    var searchFont = function (font, _styles) {
        return _styles._fonts[utils.stringify(font)];
    };

    var searchSavedFontsForUpdate = function (_styles, cellIndices) {
        var index = 0, fontCount = 0;
        var cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined
                    || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    fontCount++;
                    if (Object.keys(cellStyle.cellIndices).length
                        !== cellIndices.length
                        || fontCount > 1 || cellStyle._font === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._font;
    };

    var addFont = function (font, _styles) {
        if (!_styles._fonts) {
            _styles._fonts = {};
            _styles._fontsCount = 0;
        }
        var index = _styles._fontsCount++;
        _styles._fonts[utils.stringify(font)] = "" + index;
        return _styles._fonts[utils.stringify(font)];
    };

    var updateFont = function (font, savedFont, _styles) {
        mergeFont(font, savedFont, _styles, true);
        _styles._fonts[utils.stringify(font)] = savedFont;
        return savedFont;
    };

    var mergeFont = function (font, savedFont, _styles, deleteSavedFont) {
        var savedFontDetails;
        for (var key in _styles._fonts) {
            if (_styles._fonts[key] === savedFont) {
                savedFontDetails = JSON.parse(key);
                break;
            }
        }
        if (deleteSavedFont) {
            delete _styles._fonts[utils.stringify(savedFontDetails)];
        }
        for (var key in font) {
            if (font[key])
                savedFontDetails[key] = font[key];
            font[key] = savedFontDetails[key];
        }
        return font;
    };

    var getFontCounts = function (font, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._font === font) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    var getFontForCells = function (_styles, options) {
        var newStyleCreated = false, font = createFont(options), fontIndex;
        var savedFont = _styles._fonts ? searchFont(font, _styles) : null;
        if (savedFont !== undefined && savedFont !== null) {
            fontIndex = savedFont;
        } else {
            newStyleCreated = true;
        }
        if (!fontIndex) {
            savedFont = searchSavedFontsForUpdate(_styles, options.cellIndices);
            if (savedFont !== false) {
                fontIndex = updateFont(font, savedFont, _styles);
            }
        }
        if (fontIndex === undefined || fontIndex === null)
            fontIndex = addFont(font, _styles);
        return {
            font: font,
            fontIndex: fontIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getFontForCell = function (_styles, options, cellStyle) {
        // Create Font Object
        var newStyleCreated = false, font = createFont(options), fontIndex;

        var savedFont = _styles._fonts ? searchFont(font, _styles) : null;

        if (savedFont !== undefined && savedFont !== null) {
            fontIndex = savedFont;
        } else {
            // Check if font can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._font) {
            var fontCount = getFontCounts(cellStyle._font, _styles);
            if (fontCount <= 1) {
                // Update font
                fontIndex = updateFont(font, cellStyle._font, _styles);
            } else {
                // Merge Font
                font = mergeFont(font, cellStyle._font, _styles);
                savedFont = searchFont(font, _styles);
                if (savedFont !== undefined && savedFont !== null) {
                    fontIndex = savedFont;
                } else {
                    fontIndex = null;
                    newStyleCreated = true;
                }
            }
        }
        if (fontIndex === undefined || fontIndex === null)
            fontIndex = addFont(font, _styles);

        return {
            font: font,
            fontIndex: fontIndex,
            newStyleCreated: newStyleCreated
        };
    };

    return {
        createFont: createFont,
        getFontForCell: getFontForCell,
        getFontForCells: getFontForCells,
        generateContent: generateContent,
        generateSingleContent: generateSingleContent
    };
});
define('oxml_xlsx_num_format',['utils'], function (utils) {
    var generateContent = function (_styles) {
        var stylesString = '';
        if (_styles._numFormats) {
            stylesString += '<numFmts count="' + Object.keys(_styles._numFormats).length + '">';
            var numFormatKey;
            for (numFormatKey in _styles._numFormats) {
                var numFormat = JSON.parse(numFormatKey);
                stylesString += '<numFmt numFmtId="' + _styles._numFormats[numFormatKey] + '" formatCode="' + numFormat.formatString + '"/>';
            }
            stylesString += '</numFmts>';
        }
        return stylesString;
    };

    var createNumFormat = function (options) {
        var numberFormat = {};
        numberFormat.formatString = options.numberFormat;
        return numberFormat;
    };

    var searchNumFormat = function (numFormat, _styles) {
        return _styles._numFormats[utils.stringify(numFormat)];
    };

    var addNumFormat = function (numFormat, _styles) {
        if (!_styles._numFormats) {
            _styles._numFormats = {};
            _styles._numFormatsCount = 200;
        }
        var index = _styles._numFormatsCount++;
        _styles._numFormats[utils.stringify(numFormat)] = "" + index;
        return _styles._numFormats[utils.stringify(numFormat)];
    };

    var updateNumFormat = function (numFormat, savedNumFormat, _styles) {
        var savedNumFormatDetails;
        for (var key in _styles._numFormats) {
            if (_styles._numFormats[key] === savedNumFormat) {
                savedNumFormatDetails = JSON.parse(key);
                break;
            }
        }
        if (numFormat.formatString) {
            delete _styles._numFormats[utils.stringify(savedNumFormatDetails)];
            _styles._numFormats[utils.stringify(numFormat)] = savedNumFormat;
        } else {
            numFormat.formatString = savedNumFormatDetails.formatString;
        }
        return savedNumFormat;
    };

    var getNumFormatCounts = function (numFormat, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._numFormat === numFormat) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    var searchSavedNumFormatForUpdate = function (_styles, cellIndices) {
        var index = 0, numFormatCount = 0;
        var cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    numFormatCount++;
                    if (Object.keys(cellStyle.cellIndices).length !== cellIndices.length || numFormatCount > 1) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._numFormat;
    };

    var getNumFormatForCells = function (_styles, options) {
        var newStyleCreated = false, numFormat = createNumFormat(options), numFormatIndex;
        var savedNumFormat = _styles._numFormats ? searchNumFormat(numFormat, _styles) : null;
        if (savedNumFormat !== undefined && savedNumFormat !== null) {
            numFormatIndex = savedNumFormat;
        } else if (numFormat.formatString) {
            newStyleCreated = true;
        }
        if (!numFormatIndex) {
            savedNumFormat = searchSavedNumFormatForUpdate(_styles, options.cellIndices);
            if (savedNumFormat !== false) {
                numFormatIndex = updateNumFormat(numFormat, savedNumFormat, _styles);
            }
        }
        if ((numFormatIndex === undefined || numFormatIndex === null) && numFormat.formatString)
            numFormatIndex = addNumFormat(numFormat, _styles);

        return {
            numFormat: numFormat,
            numFormatIndex: numFormatIndex || false,
            newStyleCreated: newStyleCreated
        };
    };

    var getNumFormatForCell = function (_styles, options, cellStyle) {
        // Create Num Format Object
        var newStyleCreated = false, numFormat = createNumFormat(options), numFormatIndex;
        var savedNumFormat = _styles._numFormats ? searchNumFormat(numFormat, _styles) : null;
        if (savedNumFormat !== undefined && savedNumFormat !== null) {
            numFormatIndex = savedNumFormat;
        } else if (numFormat.formatString) {
            // Check if num format can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._numFormat) {
            var numFormatCount = getNumFormatCounts(cellStyle._numFormat, _styles);
            if (numFormatCount <= 1 && _styles._numFormats) {
                // Update num format
                numFormatIndex = updateNumFormat(numFormat, cellStyle._numFormat, _styles);
            }
        }
        if ((numFormatIndex === undefined || numFormatIndex === null) && numFormat.formatString)
            numFormatIndex = addNumFormat(numFormat, _styles);


        return {
            numFormat: numFormat,
            numFormatIndex: numFormatIndex || false,
            newStyleCreated: newStyleCreated
        };
    };

    return {
        generateContent: generateContent,
        getNumFormatForCell: getNumFormatForCell,
        getNumFormatForCells: getNumFormatForCells
    };
});
define('oxml_xlsx_border',['utils'], function (utils) {
    var getBorderString = function (border, borderType) {
        if (border.style || border.color) {
            var borderStyleString = border.style ? ' style="' + border.style + '" ' : '';
            var borderColorString = border.color ? '<color rgb="' + border.color + '"/>' : '';
            return '<' + borderType + ' ' + borderStyleString + '>' + borderColorString + '</' + borderType + '>';
        }
        return '<' + borderType + '/>';
    };

    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '';
        var borderKey;
        stylesString += '<borders count="' + _styles._bordersCount + '">';
        for (borderKey in _styles._borders) {
            var border = JSON.parse(borderKey);
            if (border) {
                stylesString += generateSingleContent(border);
            } else {
                stylesString += '<border />';
            }
        }
        stylesString += '</borders>';
        return stylesString;
    };

    var generateSingleContent = function (border) {
        var stylesString = '<border>';
        stylesString += getBorderString(border.left, 'left');
        stylesString += getBorderString(border.right, 'right');
        stylesString += getBorderString(border.top, 'top');
        stylesString += getBorderString(border.bottom, 'bottom');
        stylesString += getBorderString(border.diagonal, 'diagonal');
        stylesString += '</border>';
        return stylesString;
    };

    var createBorder = function (options) {
        var border = {};
        if (options.border) {
            border.left = {
                color: (options.border.left ? options.border.left.color : false) || options.border.color || false,
                style: (options.border.left ? options.border.left.style : false) || options.border.style || false
            };
            border.top = {
                color: (options.border.top ? options.border.top.color : false) || options.border.color || false,
                style: (options.border.top ? options.border.top.style : false) || options.border.style || false
            };
            border.bottom = {
                color: (options.border.bottom ? options.border.bottom.color : false) || options.border.color || false,
                style: (options.border.bottom ? options.border.bottom.style : false) || options.border.style || false
            };
            border.right = {
                color: (options.border.right ? options.border.right.color : false) || options.border.color || false,
                style: (options.border.right ? options.border.right.style : false) || options.border.style || false
            };
            border.diagonal = {
                color: (options.border.diagonal ? options.border.diagonal.color : false) || options.border.color || false,
                style: (options.border.diagonal ? options.border.diagonal.style : false) || options.border.style || false
            };
            border.isValid = border.left.style || border.top.style || border.bottom.style || border.right.style || border.diagonal.style
                || border.left.color || border.top.color || border.bottom.color || border.right.color || border.diagonal.color;
            return border.isValid ? border : false;
        }
        return false;
    };

    var searchBorder = function (border, _styles) {
        return _styles._borders[utils.stringify(border)];
    };

    var addBorder = function (border, _styles) {
        if (!_styles._borders) {
            _styles._borders = {};
            _styles._bordersCount = 0;
        }
        var index = _styles._bordersCount++;
        _styles._borders[utils.stringify(border)] = "" + index;
        return _styles._borders[utils.stringify(border)];
    };

    var searchSavedBordersForUpdate = function (_styles, cellIndices) {
        var index = 0, borderCount = 0;
        var cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    borderCount++;
                    if (Object.keys(cellStyle.cellIndices).length !== cellIndices.length || borderCount > 1 || cellStyle._border === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._border;
    };

    var updateBorder = function (border, savedBorder, _styles) {
        mergeBorder(border, savedBorder, _styles, true);
        _styles._borders[utils.stringify(border)] = savedBorder;
        return savedBorder;
    };

    var mergeBorder = function (border, savedBorder, _styles, deleteSavedBorder) {
        var savedBorderDetails;
        for (var key in _styles._borders) {
            if (_styles._borders[key] === savedBorder) {
                savedBorderDetails = JSON.parse(key);
                break;
            }
        }
        if (deleteSavedBorder) {
            delete _styles._borders[utils.stringify(savedBorderDetails)];
        }
        if (border) {
            for (var key in border) {
                if (border[key].color || border[key].style)
                    savedBorderDetails[key] = border[key];
                border[key] = savedBorderDetails[key];
            }
        } else
            border = savedBorderDetails;
        return border;
    };

    var getBorderCounts = function (border, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._border === border) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    var getBorderForCell = function (_styles, options, cellStyle) {
        // Create Border Object
        var newStyleCreated = false, border = createBorder(options), borderIndex;

        var savedBorder = _styles._borders ? searchBorder(border, _styles) : null;

        if (savedBorder !== undefined && savedBorder !== null) {
            borderIndex = savedBorder;
        } else {
            // Check if border can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._border && cellStyle._border !== savedBorder) {
            var borderCount = getBorderCounts(cellStyle._border, _styles);
            if (borderCount <= 1) {
                // Update border
                borderIndex = updateBorder(border, cellStyle._border, _styles);
            }
        }
        if (borderIndex === undefined || borderIndex === null) {
            if (cellStyle && cellStyle._border) {
                border = mergeBorder(border, cellStyle._border, _styles, false);
            }
            borderIndex = addBorder(border, _styles);
        }

        return {
            border: border,
            borderIndex: borderIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getBorderForCells = function (_styles, options) {
        var newStyleCreated = false, border = createBorder(options), borderIndex;
        var savedBorder = _styles._borders ? searchBorder(border, _styles) : null;
        if (savedBorder !== undefined && savedBorder !== null) {
            borderIndex = savedBorder;
        } else {
            newStyleCreated = true;
        }
        savedBorder = searchSavedBordersForUpdate(_styles, options.cellIndices);
        if (savedBorder !== false) {
            borderIndex = updateBorder(border, savedBorder, _styles);
        }
        if (borderIndex === undefined || borderIndex === null)
            borderIndex = addBorder(border, _styles);
        return {
            border: border,
            borderIndex: borderIndex,
            newStyleCreated: newStyleCreated
        };
    };

    return {
        createBorder: createBorder,
        getBorderForCell: getBorderForCell,
        getBorderForCells: getBorderForCells,
        generateContent: generateContent,
        generateSingleContent: generateSingleContent
    };
});
define('oxml_xlsx_fill',['utils'], function (utils) {
    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '';
        var fillKey;
        stylesString += '<fills count="' + _styles._fillsCount + '">';
        for (fillKey in _styles._fills) {
            var fill = JSON.parse(fillKey);
            if (fill) {
                stylesString += getFillString(fill);
            } else {
                stylesString += '<fill/>';
            }
        }
        stylesString += '</fills>';
        return stylesString;
    };

    var getFillString = function (fill) {
        var fillString = '<fill>';
        if (fill.pattern) {
            var patternType = fill.pattern || 'none';
            var colorString =
                fill.fgColor ? '<fgColor rgb="' + fill.fgColor + '"/>' : '';
            colorString +=
                fill.bgColor ? '<bgColor rgb="' + fill.bgColor + '"/>' : '';
            fillString += '<patternFill patternType="'
                + patternType + '">' + colorString + '</patternFill>';
        } else if (fill.gradient) {
            var typeString =
                fill.gradient.type ? ' type="' + fill.gradient.type + '" ' : '';
            var leftString =
                fill.gradient.left ? ' left="' + fill.gradient.left + '" ' : '';
            var rightString =
                fill.gradient.right ? ' right="'
                    + fill.gradient.right + '"' : '';
            var topString =
                fill.gradient.top ? ' top="' + fill.gradient.top + '" ' : '';
            var bottomString =
                fill.gradient.bottom ? ' bottom="'
                    + fill.gradient.bottom + '" ' : '';
            var degreeString =
                fill.gradient.degree ? ' degree="'
                    + fill.gradient.degree + '" ' : '';
            fillString += '<gradientFill '
                + typeString + leftString + rightString
                + topString + bottomString + degreeString + ' >';
            for (var stopIndex = 0;
                stopIndex < fill.gradient.stops.length;
                stopIndex++) {
                fillString += '<stop position="'
                    + fill.gradient.stops[stopIndex].position + '">';
                fillString += '<color rgb="'
                    + fill.gradient.stops[stopIndex].color + '" /></stop>';
            }
            fillString += '</gradientFill>';
        } else {
            return '<fill/>';
        }
        return fillString + '</fill>';
    };


    var createFill = function (options) {
        if (options.fill && options.fill.pattern) {
            var fill = {};
            fill.pattern = options.fill.pattern;
            fill.fgColor = options.fill.foreColor || false;
            fill.bgColor = options.fill.backColor
                || options.fill.color || false;
            return fill;
        } else if (options.fill
            && options.fill.gradient
            && options.fill.gradient.stops) {
            var fill = {};
            fill.gradient = {};
            fill.gradient.degree = options.fill.gradient.degree || false;
            fill.gradient.bottom = options.fill.gradient.bottom || false;
            fill.gradient.left = options.fill.gradient.left || false;
            fill.gradient.right = options.fill.gradient.right || false;
            fill.gradient.top = options.fill.gradient.top || false;
            fill.gradient.type = options.fill.gradient.type || false;
            fill.gradient.stops = [];
            for (var stopIndex = 0;
                stopIndex < options.fill.gradient.stops.length;
                stopIndex++) {
                var stop = {};
                stop.position =
                    options.fill.gradient.stops[stopIndex].position || 0;
                stop.color =
                    options.fill.gradient.stops[stopIndex].color || false;
                fill.gradient.stops.push(stop);
            }
            return fill;
        }
        return false;
    };

    var searchFill = function (fill, _styles) {
        return _styles._fills[utils.stringify(fill)];
    };

    var addFill = function (fill, _styles) {
        if (!_styles._fills) {
            _styles._fills = {};
            _styles._fillsCount = 0;
        }
        var index = _styles._fillsCount++;
        _styles._fills[utils.stringify(fill)] = "" + index;
        return _styles._fills[utils.stringify(fill)];
    };

    var searchSavedFillsForUpdate = function (_styles, cellIndices) {
        var index = 0, fillCount = 0, cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined
                    || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    fillCount++;
                    if (Object.keys(cellStyle.cellIndices).length
                        !== cellIndices.length
                        || fillCount > 1 || cellStyle._fill === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._fill;
    };

    var updateFill = function (fill, savedFill, _styles) {
        mergeFill(fill, savedFill, _styles, true);
        _styles._fills[utils.stringify(fill)] = savedFill;
        return savedFill;
    };

    var mergeFill = function (fill, savedFill, _styles, deleteExisting) {
        var savedFillDetails;
        for (var key in _styles._fills) {
            if (_styles._fills[key] === savedFill) {
                savedFillDetails = JSON.parse(key);
                break;
            }
        }
        if (deleteExisting)
            delete _styles._fills[utils.stringify(savedFillDetails)];
        for (var key in fill) {
            if (fill[key])
                savedFillDetails[key] = fill[key];
            fill[key] = savedFillDetails[key];
        }
        return fill;
    };

    var getFillForCells = function (_styles, options) {
        var newStyleCreated = false, fill = createFill(options), fillIndex;
        var savedFill = _styles._fills ? searchFill(fill, _styles) : null;
        if (savedFill !== undefined && savedFill !== null) {
            fillIndex = savedFill;
        } else {
            newStyleCreated = true;
        }
        if (!fillIndex) {
            savedFill = searchSavedFillsForUpdate(_styles, options.cellIndices);
            if (savedFill !== false) {
                fillIndex = updateFill(fill, savedFill, _styles);
            }
        }
        if (fillIndex === undefined || fillIndex === null)
            fillIndex = addFill(fill, _styles);
        return {
            fill: fill,
            fillIndex: fillIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getFillForCell = function (_styles, options, cellStyle) {
        // Create Fill Object
        var newStyleCreated = false, fill = createFill(options), fillIndex;

        var savedFill = _styles._fills ? searchFill(fill, _styles) : null;

        if (savedFill !== undefined && savedFill !== null) {
            fillIndex = savedFill;
        } else {
            // Check if fill can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._fill) {
            var fillCount = getFillCounts(cellStyle._fill, _styles);
            if (fillCount <= 1) {
                // Update fill
                fillIndex = updateFill(fill, cellStyle._fill, _styles);
            }
        }
        if (fillIndex === undefined || fillIndex === null) {
            if (cellStyle && cellStyle._fill) {
                fill = mergeFill(fill, cellStyle._fill, _styles, false);
            }
            fillIndex = addFill(fill, _styles);
        }
        if (!fillIndex && cellStyle && cellStyle._fill)
            fillIndex = cellStyle._fill;

        return {
            fill: fill,
            fillIndex: fillIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getFillCounts = function (fill, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._fill === fill) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    return {
        createFill: createFill,
        generateContent: generateContent,
        getFillForCell: getFillForCell,
        getFillForCells: getFillForCells,
        generateSingleContent: getFillString
    };
});
define('oxml_xlsx_styles',['utils',
    'oxml_xlsx_font',
    'oxml_xlsx_num_format',
    'oxml_xlsx_border',
    'oxml_xlsx_fill'],
    function (utils,
        oxmlXlsxFont,
        oxmlXlsxNumFormat,
        oxmlXlsxBorder,
        oxmlXlsxFill) {
        var generateContent = function (_styles) {
            // Create Styles
            var stylesString = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
            var index = 0;
            stylesString += oxmlXlsxNumFormat.generateContent(_styles);
            stylesString += oxmlXlsxFont.generateContent(_styles);
            stylesString += oxmlXlsxFill.generateContent(_styles);
            stylesString += oxmlXlsxBorder.generateContent(_styles);
            stylesString += '<cellStyleXfs count="1"><xf /></cellStyleXfs>';

            stylesString += '<cellXfs count="'
                + (parseInt(_styles.styles.length, 10) + 1) + '"><xf />';
            for (index = 0; index < _styles.styles.length; index++) {
                var numFormatString = _styles.styles[index]._numFormat
                    ? ' numFmtId="'
                    + _styles.styles[index]._numFormat + '" ' : '';
                var borderString = _styles.styles[index]._border
                    ? ' borderId="' + _styles.styles[index]._border + '" ' : '';
                var fillString = _styles.styles[index]._fill
                    ? ' fillId="' + _styles.styles[index]._fill + '" ' : '';
                stylesString += '<xf fontId="' + _styles.styles[index]._font
                    + '" ' + numFormatString
                    + borderString + fillString + ' />';
            }
            var tableStyles = '';
            if (_styles.tableStyles && _styles.tableStyles.length) tableStyles = generateTableContent(_styles);
            stylesString += '</cellXfs>' + tableStyles + '</styleSheet>';
            return stylesString;
        };

        var generateTableContent = function (_styles) {
            var dxfsCount = 0, tableStyles = '', tableStylesCount = 0, tableStyleCount;
            for (var index = 0; index < _styles.tableStyles.length; index++) {
                var dxfString = '', individualTalbeStyle;
                dxfsCount++;
                tableStylesCount++;
                tableStyleCount = 1;
                dxfString += generateDxfContent(_styles.tableStyles[index]);
                individualTalbeStyle = '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="wholeTable"/>';
                if (_styles.tableStyles[index].evenColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="secondColumnStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].evenColumn);
                }
                if (_styles.tableStyles[index].oddColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstColumnStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].oddColumn);
                }
                if (_styles.tableStyles[index].evenRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="secondRowStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].evenRow);
                }
                if (_styles.tableStyles[index].oddRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstRowStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].oddRow);
                }
                if (_styles.tableStyles[index].firstColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstColumn"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstColumn);
                }
                if (_styles.tableStyles[index].lastColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastColumn"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastColumn);
                }
                if (_styles.tableStyles[index].firstRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="headerRow"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRow);
                }
                if (_styles.tableStyles[index].lastRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="totalRow"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRow);
                }
                if (_styles.tableStyles[index].firstRowFirstCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstHeaderCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRowFirstCell);
                }
                if (_styles.tableStyles[index].firstRowLastCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastHeaderCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRowLastCell);
                }
                if (_styles.tableStyles[index].lastRowFirstCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstTotalCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRowFirstCell);
                }
                if (_styles.tableStyles[index].lastRowLastCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastTotalCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRowLastCell);
                }
                tableStyles += '<tableStyle count="' + tableStyleCount + '" name="' + _styles.tableStyles[index].name + '">' + individualTalbeStyle + '</tableStyle>';
            }
            tableStyles = '<tableStyles count="' + tableStylesCount + '">' + tableStyles + '</tableStyles>';
            dxfString = '<dxfs count="' + dxfsCount + '">' + dxfString + '</dxfs>' + tableStyles;
            return dxfString;
        };

        var generateDxfContent = function (style) {
            dxfString = '<dxf>';
            if (style.font) {
                dxfString += oxmlXlsxFont.generateSingleContent(style.font);
            }
            if (style.fill) {
                dxfString += oxmlXlsxFill.generateSingleContent(style.fill);
            }
            if (style.border) {
                dxfString += oxmlXlsxBorder.generateSingleContent(style.border);
            }
            dxfString += '</dxf>';
            return dxfString;
        };

        var attach = function (file, _styles) {
            var styles = generateContent(_styles);
            file.addFile(styles, _styles.fileName, "workbook");
        };

        var createStyleParts = function (_workBook, _rel, _contentType) {
            var lastSheetRel = _workBook.getLastSheet();
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0")
                .replace("rId", ""), 10) + 1;
            _contentType
                .addContentType("Override",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                    {
                        PartName: "/workbook/style" + nextSheetRelId + ".xml"
                    });

            _rel.addRelation(
                "rId" + nextSheetRelId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "style" + nextSheetRelId + ".xml"
            );

            return nextSheetRelId;
        };

        var searchStyleForCell = function (_styles, cellIndex) {
            var index = 0;
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined
                    && _styles.styles[index].cellIndices[cellIndex] !== null) {
                    return _styles.styles[index];
                }
            }
            return null;
        };

        var searchSimilarStyle = function (_styles, style) {
            var index = 0;
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index]._font === style._font
                    && _styles.styles[index]._numFormat === style._numFormat) {
                    return _styles.styles[index];
                }
            }
            return null;
        };

        var addStyles = function (options, _styles) {
            if (options.cellIndex || options.cellIndices) {
                var newStyleCreated = false, cellStyle,
                    saveFont, saveNumFormat, saveBorder, saveFill;
                if (options.cellIndex) {
                    cellStyle = searchStyleForCell(_styles, options.cellIndex);
                    saveFont = oxmlXlsxFont
                        .getFontForCell(_styles, options, cellStyle);
                    saveNumFormat = oxmlXlsxNumFormat
                        .getNumFormatForCell(_styles, options, cellStyle);
                    saveBorder = oxmlXlsxBorder
                        .getBorderForCell(_styles, options, cellStyle);
                    saveFill = oxmlXlsxFill
                        .getFillForCell(_styles, options, cellStyle);
                    newStyleCreated = newStyleCreated
                        || saveFont.newStyleCreated
                        || saveBorder.newStyleCreated
                        || saveFill.newStyleCreated;
                    if (cellStyle) {
                        if (cellStyle._font === saveFont.fontIndex
                            && cellStyle._numFormat === saveNumFormat.numFormatIndex
                            && cellStyle._border === saveBorder.borderIndex
                            && cellStyle._fill === saveFill.fillIndex) {
                            return cellStyle;
                        }
                        var totalCellApplied = Object.keys(cellStyle.cellIndices).length;
                        if (totalCellApplied === 1) {
                            cellStyle._font = saveFont.fontIndex;
                            cellStyle._numFormat = saveNumFormat.numFormatIndex;
                            cellStyle._border = saveBorder.borderIndex;
                            cellStyle._fill = saveFill.fillIndex;
                            return cellStyle;
                        }
                        delete cellStyle.cellIndices[options.cellIndex];
                    }

                    if (!newStyleCreated) {
                        cellStyle = searchSimilarStyle(_styles, {
                            _font: saveFont.fontIndex,
                            _numFormat: saveNumFormat.numFormatIndex,
                            _border: saveBorder.borderIndex,
                            _fill: saveFill.fillIndex
                        });
                        if (cellStyle) {
                            cellStyle.cellIndices[options.cellIndex] = Object.keys(cellStyle.cellIndices).length;
                            return cellStyle;
                        }
                    }

                    cellStyle = {
                        _font: saveFont.fontIndex,
                        _numFormat: saveNumFormat.numFormatIndex || false,
                        _border: saveBorder.borderIndex,
                        _fill: saveFill.fillIndex,
                        cellIndices: {}
                    };

                    cellStyle.cellIndices[options.cellIndex] = 0;
                    cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);
                    _styles.styles.push(cellStyle);
                    return cellStyle;
                }

                if (options.cellIndices) {
                    var index;
                    saveFont = oxmlXlsxFont.getFontForCells(_styles, options);
                    saveNumFormat = oxmlXlsxNumFormat.getNumFormatForCells(_styles, options);
                    saveBorder = oxmlXlsxBorder.getBorderForCells(_styles, options);
                    saveFill = oxmlXlsxFill.getFillForCells(_styles, options);
                    newStyleCreated = newStyleCreated || saveFont.newStyleCreated
                        || saveNumFormat.newStyleCreated || saveBorder.newStyleCreated || saveFill.newStyleCreated;
                    if (!newStyleCreated) {
                        cellStyle = searchSimilarStyle(_styles, {
                            _font: saveFont.fontIndex,
                            _numFormat: saveNumFormat.numFormatIndex,
                            _border: saveBorder.borderIndex,
                            _fill: saveFill.fillIndex
                        });

                        if (cellStyle) {
                            for (index = 0; index < options.cellIndices.length; index++) {
                                var cellIndex = options.cellIndices[index];
                                var savedCellStyle = searchStyleForCell(_styles, cellIndex);
                                delete savedCellStyle.cellIndices[cellIndex];
                                cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                            }
                            return cellStyle;
                        }
                    }
                    // Maintain old styling

                    cellStyle = {
                        _font: saveFont.fontIndex,
                        _numFormat: saveNumFormat.numFormatIndex,
                        _border: saveBorder.borderIndex,
                        _fill: saveFill.fillIndex,
                        cellIndices: {}
                    };

                    for (index = 0; index < options.cellIndices.length; index++) {
                        var cellIndex = options.cellIndices[index];
                        var savedCellStyle = searchStyleForCell(_styles, cellIndex);
                        if (savedCellStyle) {
                            delete savedCellStyle.cellIndices[cellIndex];
                        }
                        cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                        if (savedCellStyle && !Object.keys(savedCellStyle.cellIndices).length) {
                            for (key in cellStyle) {
                                savedCellStyle[key] = cellStyle[key];
                            }
                            return savedCellStyle;
                        }
                    }
                    cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);
                    _styles.styles.push(cellStyle);
                    return cellStyle;
                }
            }
        };

        var addTableStyle = function (options, tableStyleName, _styles) {
            if (!_styles.tableStyles) _styles.tableStyles = [];
            var tableStyle = prepareTableStyleObj(options);
            tableStyle.name = tableStyleName;
            if (options.evenColumn) tableStyle.evenColumn = prepareTableStyleObj(options.evenColumn);
            if (options.oddColumn) tableStyle.oddColumn = prepareTableStyleObj(options.oddColumn);
            if (options.evenRow) tableStyle.evenRow = prepareTableStyleObj(options.evenRow);
            if (options.oddRow) tableStyle.oddRow = prepareTableStyleObj(options.oddRow);
            if (options.firstColumn) tableStyle.firstColumn = prepareTableStyleObj(options.firstColumn);
            if (options.lastColumn) tableStyle.lastColumn = prepareTableStyleObj(options.lastColumn);
            if (options.firstRow) tableStyle.firstRow = prepareTableStyleObj(options.firstRow);
            if (options.lastRow) tableStyle.lastRow = prepareTableStyleObj(options.lastRow);
            if (options.firstRowFirstCell) tableStyle.firstRowFirstCell = prepareTableStyleObj(options.firstRowFirstCell);
            if (options.firstRowLastCell) tableStyle.firstRowLastCell = prepareTableStyleObj(options.firstRowLastCell);
            if (options.lastRowFirstCell) tableStyle.lastRowFirstCell = prepareTableStyleObj(options.lastRowFirstCell);
            if (options.lastRowLastCell) tableStyle.lastRowLastCell = prepareTableStyleObj(options.lastRowLastCell);
            _styles.tableStyles.push(tableStyle);
            return tableStyle;
        };

        var prepareTableStyleObj = function (options) {
            var tableStyle = {};
            tableStyle.font = oxmlXlsxFont.createFont(options);
            tableStyle.fill = oxmlXlsxFill.createFill(options);
            tableStyle.border = oxmlXlsxBorder.createBorder(options);
            return tableStyle;
        };

        return {
            createStyle: function (_workbook, _rel, _contentType) {
                var sheetId = createStyleParts(_workbook, _rel, _contentType);
                var _styles = {
                    sheetId: sheetId,
                    fileName: "style" + sheetId + ".xml",
                    styles: [],
                    _fonts: {},
                    _borders: {},
                    _fills: {},
                    _fontsCount: 1,
                    _bordersCount: 1,
                    _fillsCount: 2
                };
                var font = {
                    bold: false,
                    italic: false,
                    underline: false,
                    size: false,
                    color: false,
                    strike: false,
                    family: false,
                    name: false,
                    scheme: false
                };
                _styles._fonts[utils.stringify(font)] = 0;
                _styles._borders[false] = 0;
                _styles._fills[false] = 0;
                var fillGray125 = {
                    pattern: 'gray125',
                    fgColor: false,
                    bgColor: false
                };
                _styles._fills[utils.stringify(fillGray125)] = 1;
                return {
                    _styles: _styles,
                    generateContent: function () {
                        return generateContent(_styles);
                    },
                    attach: function (file) {
                        attach(file, _styles);
                    },
                    addStyles: function (options) {
                        return addStyles(options, _styles);
                    },
                    addTableStyle: function (options, tableStyleName) {
                        return addTableStyle(options, tableStyleName, _styles);
                    }
                };
            }
        };
    });
define('oxml_workbook',['oxml_content_types', 'oxml_rels', 'oxml_sheet', 'oxml_xlsx_styles'], function (oxmlContentTypes, oxmlRels, oxmlSheet, oxmlXlsxStyles) {
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
    };

    var addSheet = function (_workBook, sheetName, xlsxContentTypes) {
        var lastSheetRel = getLastSheet(_workBook);
        var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
        sheetName = sheetName || "sheet" + nextSheetRelId;
        _workBook._rels.addRelation(
            "rId" + nextSheetRelId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "sheets/sheet" + nextSheetRelId + ".xml");
        // Update Sheets
        var sheet = oxmlSheet.createSheet(sheetName, nextSheetRelId, "rId" + nextSheetRelId, _workBook, xlsxContentTypes);
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
            return addSheet(_workBook, sheetName, xlsxContentTypes);
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
define('oxml_xlsx',['fileHandler', 'oxml_content_types', 'oxml_rels', 'oxml_workbook'], function (fileHandler, oxmlContentTypes, oxmlRels, oxmlWorkBook) {
    'use strict';

    var oxml = {};

    var downloadFile = function (fileName, callback, _xlsx) {
        try {
            var jsZip, fs;
            if (typeof window === "undefined") {
                jsZip = _jsZip();
                fs = _fs;
            } else if (typeof JSZip === "undefined") {
                if (callback) {
                    callback('Err: JSZip reference not found.');
                } else if (typeof Promise !== "undefined") {
                    return new Promise(function (resolve, reject) {
                        reject("Err: JSZip reference not found.");
                    });
                }
            } else jsZip = new JSZip();

            var file = fileHandler.createFile(jsZip, fs);

            // Attach Content Types
            _xlsx.contentTypes.attach(file);

            // Attach RELS
            _xlsx._rels.attach(file);

            // Attach WorkBook
            _xlsx.workBook.attach(file);

            return file.saveFile(fileName, callback);
        } catch (err) {
            if (callback) {
                callback('Err: Not able to create Workbook.');
            } else if (typeof Promise !== "undefined") {
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

        var destroy = function () {
            _xlsx.workBook.destroy();
            delete _xlsx.workBook;
            _xlsx._rels.destroy();
            delete _xlsx._rels;
            _xlsx.contentTypes.destroy();
            delete _xlsx.contentTypes;
            _xlsx = null;
        };

        return {
            _xlsx: _xlsx,
            sheet: _xlsx.workBook.addSheet,
            download: download,
            destroy: destroy
        };
    };

    oxml.xlsx = createXLSX;

    if (typeof window !== "undefined" && !window.oxml) {
        window.oxml = oxml;
    }

    return oxml;
});
    return require('oxml_xlsx');
}));
