(function (root, factory) {
        root.oxml = factory();
}(this, function () {


var requirejs,require,define;!function(e){function n(e,n){return m.call(e,n)}function r(e,n){var r,i,t,o,u,f,l,c,s,p,a,d=n&&n.split("/"),h=g.map,m=h&&h["*"]||{};if(e){for(u=(e=e.split("/")).length-1,g.nodeIdCompat&&y.test(e[u])&&(e[u]=e[u].replace(y,"")),"."===e[0].charAt(0)&&d&&(e=d.slice(0,d.length-1).concat(e)),s=0;s<e.length;s++)if("."===(a=e[s]))e.splice(s,1),s-=1;else if(".."===a){if(0===s||1===s&&".."===e[2]||".."===e[s-1])continue;s>0&&(e.splice(s-1,2),s-=2)}e=e.join("/")}if((d||m)&&h){for(s=(r=e.split("/")).length;s>0;s-=1){if(i=r.slice(0,s).join("/"),d)for(p=d.length;p>0;p-=1)if((t=h[d.slice(0,p).join("/")])&&(t=t[i])){o=t,f=s;break}if(o)break;!l&&m&&m[i]&&(l=m[i],c=s)}!o&&l&&(o=l,f=c),o&&(r.splice(0,f,o),e=r.join("/"))}return e}function i(n,r){return function(){var i=v.call(arguments,0);return"string"!=typeof i[0]&&1===i.length&&i.push(null),c.apply(e,i.concat([n,r]))}}function t(e){return function(n){a[e]=n}}function o(r){if(n(d,r)){var i=d[r];delete d[r],h[r]=!0,l.apply(e,i)}if(!n(a,r)&&!n(h,r))throw new Error("No "+r);return a[r]}function u(e){var n,r=e?e.indexOf("!"):-1;return r>-1&&(n=e.substring(0,r),e=e.substring(r+1,e.length)),[n,e]}function f(e){return e?u(e):[]}var l,c,s,p,a={},d={},g={},h={},m=Object.prototype.hasOwnProperty,v=[].slice,y=/\.js$/;s=function(e,n){var i,t=u(e),f=t[0],l=n[1];return e=t[1],f&&(i=o(f=r(f,l))),f?e=i&&i.normalize?i.normalize(e,function(e){return function(n){return r(n,e)}}(l)):r(e,l):(f=(t=u(e=r(e,l)))[0],e=t[1],f&&(i=o(f))),{f:f?f+"!"+e:e,n:e,pr:f,p:i}},p={require:function(e){return i(e)},exports:function(e){var n=a[e];return void 0!==n?n:a[e]={}},module:function(e){return{id:e,uri:"",exports:a[e],config:function(e){return function(){return g&&g.config&&g.config[e]||{}}}(e)}}},l=function(r,u,l,c){var g,m,v,y,j,q,x,b=[],w=typeof l;if(c=c||r,q=f(c),"undefined"===w||"function"===w){for(u=!u.length&&l.length?["require","exports","module"]:u,j=0;j<u.length;j+=1)if(y=s(u[j],q),"require"===(m=y.f))b[j]=p.require(r);else if("exports"===m)b[j]=p.exports(r),x=!0;else if("module"===m)g=b[j]=p.module(r);else if(n(a,m)||n(d,m)||n(h,m))b[j]=o(m);else{if(!y.p)throw new Error(r+" missing "+m);y.p.load(y.n,i(c,!0),t(m),{}),b[j]=a[m]}v=l?l.apply(a[r],b):void 0,r&&(g&&g.exports!==e&&g.exports!==a[r]?a[r]=g.exports:v===e&&x||(a[r]=v))}else r&&(a[r]=l)},requirejs=require=c=function(n,r,i,t,u){if("string"==typeof n)return p[n]?p[n](r):o(s(n,f(r)).f);if(!n.splice){if((g=n).deps&&c(g.deps,g.callback),!r)return;r.splice?(n=r,r=i,i=null):n=e}return r=r||function(){},"function"==typeof i&&(i=t,t=u),t?l(e,n,r,i):setTimeout(function(){l(e,n,r,i)},4),c},c.config=function(e){return c(e)},requirejs._defined=a,(define=function(e,r,i){if("string"!=typeof e)throw new Error("See almond README: incorrect module build, no module name");r.splice||(i=r,r=[]),n(a,e)||n(d,e)||(d[e]=[e,r,i])}).amd={jQuery:!0}}();

define("build/almond", function(){});

define('fileHandler',[], function () {
    "use strict";

    var createCompressedFile = function () {
        if (!window.JSZip) {
            return;
        }
        var zip = new JSZip();

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
            return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
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
                            callback();
                        } else if (typeof Promise !== "undefined") {
                            return new Promise(function (resolve, reject) {
                                resolve();
                            });
                        }
                    }
                    catch (err){
                        if (callback) {
                            callback("Err: Not able to create file object.");
                        } else if (typeof Promise !== "undefined") {
                            return new Promise(function (resolve, reject) {
                                reject("Err: Not able to create file object.");
                            });
                        }
                    }
                });
        };

        return compressedFile;
    };

    return {
        createFile: createCompressedFile
    }
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
    }

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
        }
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
define('oxml_sheet',[], function () {
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
            return '<c r="' + cellIndex + '" t="inlineStr" ' + styleString + '><is><t>' + value.value.replace(/&/g, '&amp;') + '</t></is></c>';
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

    var generateContent = function (_sheet) {
        var rowIndex = 0, sheetValues = '';
        if (_sheet.values && _sheet.values.length) {
            for (rowIndex = 0; rowIndex < _sheet.values.length; rowIndex++) {
                if (_sheet.values[rowIndex] && _sheet.values[rowIndex].length > 0) {
                    sheetValues += '<row r="' + (rowIndex + 1) + '">\n';

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
        var sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + sheetValues + '</sheetData></worksheet>';
        return sheet;
    };

    var attach = function (_sheet, file) {
        var content = generateContent(_sheet);
        file.addFile(content, "sheet" + _sheet.sheetId + ".xml", "workbook/sheets");
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
            if (!value.styleIndex && _sheet.values[sheetRowIndex][columnIndex - 1])
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
            set: function (value, options) {
                cell(_sheet, rowIndex + 1, columnIndex + 1, value, options);
                this.value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
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
            var type = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null;
            var cellIndex = String.fromCharCode(65 + columnIndex) + (rowIndex + 1);
            cellRange.push({
                style: function (options) {
                    options.cellIndex = cellIndex;
                    updateSingleStyle(_sheet, options, rowIndex, columnIndex);
                    return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
                },
                rowIndex: rowIndex,
                columnIndex: String.fromCharCode(65 + columnIndex),
                value: value,
                cellIndex: cellIndex,
                type: type,
                set: function (value, options) {
                    cell(_sheet, rowIndex + 1, columnIndex + 1, value, options);
                    this.value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
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
                        var tVal2 = [];
                        if (values[index].length > totalColumns) {
                            for (var index2 = 0; index2 < values[index].length && index2 < totalColumns; index2++) {
                                tVal2.push(values[index][index2]);
                            }
                            values[index] = tVal2;
                        }
                        tVal.push(values[index]);
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
            if (formula.type !== "formula") return;
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

    var createSheet = function (sheetName, sheetId, rId, workBook) {
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
            destroy: function () {
                return destroy(_sheet);
            }
        };
    };

    return { createSheet: createSheet };
});
define('oxml_xlsx_styles',[],
    function () {
        return {
            createStyle: function (_workbook, _rel, _contentType) {
                return {
                    _styles: null,
                    generateContent: function () {
                        return '';
                    },
                    attach: function (file) {
                    },
                    addStyles: function (options) {
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
                sharedStrings += '<si><t>' + key.replace(/&/g, '&amp;') + '</t></si>';
            }
            sharedStrings += '</sst>';
            return sharedStrings;
        }
        return '';
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
define('oxml_xlsx',['fileHandler', 'oxml_content_types', 'oxml_rels', 'oxml_workbook'], function (fileHandler, oxmlContentTypes, oxmlRels, oxmlWorkBook) {
    'use strict';

    var oxml = {};

    var downloadFile = function (fileName, callback, _xlsx) {
        try {
            var file = fileHandler.createFile();

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
            if (typeof JSZip === "undefined") {
                if (callback) {
                    callback('Err: JSZip reference not found.');
                } else if (typeof Promise !== "undefined") {
                    return new Promise(function (resolve, reject) {
                        reject("Err: JSZip reference not found.");
                    });
                }
            }
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

    if (!window.oxml) {
        window.oxml = oxml;
    }

    return oxml;
});
    return require('oxml_xlsx');
}));
