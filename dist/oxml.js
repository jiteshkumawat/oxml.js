//Copyright 2012, etc.

(function (root, factory) {
        // Browser globals
        root.oxml = factory();
}(this, function () {


/**
 * almond 0.1.2 Copyright (c) 2011, The Dojo Foundation All Rights Reserved.
 * Available via the MIT or new BSD license.
 * see: http://github.com/jrburke/almond for details
 */
//Going sloppy to avoid 'use strict' string cost, but strict practices should
//be followed.
/*jslint sloppy: true */
/*global setTimeout: false */

var requirejs, require, define;
(function (undef) {
    var defined = {},
        waiting = {},
        config = {},
        defining = {},
        aps = [].slice,
        main, req;

    /**
     * Given a relative module name, like ./something, normalize it to
     * a real name that can be mapped to a path.
     * @param {String} name the relative name
     * @param {String} baseName a real name that the name arg is relative
     * to.
     * @returns {String} normalized name
     */
    function normalize(name, baseName) {
        var baseParts = baseName && baseName.split("/"),
            map = config.map,
            starMap = (map && map['*']) || {},
            nameParts, nameSegment, mapValue, foundMap,
            foundI, foundStarMap, starI, i, j, part;

        //Adjust any relative paths.
        if (name && name.charAt(0) === ".") {
            //If have a base name, try to normalize against it,
            //otherwise, assume it is a top-level require that will
            //be relative to baseUrl in the end.
            if (baseName) {
                //Convert baseName to array, and lop off the last part,
                //so that . matches that "directory" and not name of the baseName's
                //module. For instance, baseName of "one/two/three", maps to
                //"one/two/three.js", but we want the directory, "one/two" for
                //this normalization.
                baseParts = baseParts.slice(0, baseParts.length - 1);

                name = baseParts.concat(name.split("/"));

                //start trimDots
                for (i = 0; (part = name[i]); i++) {
                    if (part === ".") {
                        name.splice(i, 1);
                        i -= 1;
                    } else if (part === "..") {
                        if (i === 1 && (name[2] === '..' || name[0] === '..')) {
                            //End of the line. Keep at least one non-dot
                            //path segment at the front so it can be mapped
                            //correctly to disk. Otherwise, there is likely
                            //no path mapping for a path starting with '..'.
                            //This can still fail, but catches the most reasonable
                            //uses of ..
                            return true;
                        } else if (i > 0) {
                            name.splice(i - 1, 2);
                            i -= 2;
                        }
                    }
                }
                //end trimDots

                name = name.join("/");
            }
        }

        //Apply map config if available.
        if ((baseParts || starMap) && map) {
            nameParts = name.split('/');

            for (i = nameParts.length; i > 0; i -= 1) {
                nameSegment = nameParts.slice(0, i).join("/");

                if (baseParts) {
                    //Find the longest baseName segment match in the config.
                    //So, do joins on the biggest to smallest lengths of baseParts.
                    for (j = baseParts.length; j > 0; j -= 1) {
                        mapValue = map[baseParts.slice(0, j).join('/')];

                        //baseName segment has  config, find if it has one for
                        //this name.
                        if (mapValue) {
                            mapValue = mapValue[nameSegment];
                            if (mapValue) {
                                //Match, update name to the new value.
                                foundMap = mapValue;
                                foundI = i;
                                break;
                            }
                        }
                    }
                }

                if (foundMap) {
                    break;
                }

                //Check for a star map match, but just hold on to it,
                //if there is a shorter segment match later in a matching
                //config, then favor over this star map.
                if (!foundStarMap && starMap && starMap[nameSegment]) {
                    foundStarMap = starMap[nameSegment];
                    starI = i;
                }
            }

            if (!foundMap && foundStarMap) {
                foundMap = foundStarMap;
                foundI = starI;
            }

            if (foundMap) {
                nameParts.splice(0, foundI, foundMap);
                name = nameParts.join('/');
            }
        }

        return name;
    }

    function makeRequire(relName, forceSync) {
        return function () {
            //A version of a require function that passes a moduleName
            //value for items that may need to
            //look up paths relative to the moduleName
            return req.apply(undef, aps.call(arguments, 0).concat([relName, forceSync]));
        };
    }

    function makeNormalize(relName) {
        return function (name) {
            return normalize(name, relName);
        };
    }

    function makeLoad(depName) {
        return function (value) {
            defined[depName] = value;
        };
    }

    function callDep(name) {
        if (waiting.hasOwnProperty(name)) {
            var args = waiting[name];
            delete waiting[name];
            defining[name] = true;
            main.apply(undef, args);
        }

        if (!defined.hasOwnProperty(name)) {
            throw new Error('No ' + name);
        }
        return defined[name];
    }

    /**
     * Makes a name map, normalizing the name, and using a plugin
     * for normalization if necessary. Grabs a ref to plugin
     * too, as an optimization.
     */
    function makeMap(name, relName) {
        var prefix, plugin,
            index = name.indexOf('!');

        if (index !== -1) {
            prefix = normalize(name.slice(0, index), relName);
            name = name.slice(index + 1);
            plugin = callDep(prefix);

            //Normalize according
            if (plugin && plugin.normalize) {
                name = plugin.normalize(name, makeNormalize(relName));
            } else {
                name = normalize(name, relName);
            }
        } else {
            name = normalize(name, relName);
        }

        //Using ridiculous property names for space reasons
        return {
            f: prefix ? prefix + '!' + name : name, //fullName
            n: name,
            p: plugin
        };
    }

    function makeConfig(name) {
        return function () {
            return (config && config.config && config.config[name]) || {};
        };
    }

    main = function (name, deps, callback, relName) {
        var args = [],
            usingExports,
            cjsModule, depName, ret, map, i;

        //Use name if no relName
        relName = relName || name;

        //Call the callback to define the module, if necessary.
        if (typeof callback === 'function') {

            //Pull out the defined dependencies and pass the ordered
            //values to the callback.
            //Default to [require, exports, module] if no deps
            deps = !deps.length && callback.length ? ['require', 'exports', 'module'] : deps;
            for (i = 0; i < deps.length; i++) {
                map = makeMap(deps[i], relName);
                depName = map.f;

                //Fast path CommonJS standard dependencies.
                if (depName === "require") {
                    args[i] = makeRequire(name);
                } else if (depName === "exports") {
                    //CommonJS module spec 1.1
                    args[i] = defined[name] = {};
                    usingExports = true;
                } else if (depName === "module") {
                    //CommonJS module spec 1.1
                    cjsModule = args[i] = {
                        id: name,
                        uri: '',
                        exports: defined[name],
                        config: makeConfig(name)
                    };
                } else if (defined.hasOwnProperty(depName) || waiting.hasOwnProperty(depName)) {
                    args[i] = callDep(depName);
                } else if (map.p) {
                    map.p.load(map.n, makeRequire(relName, true), makeLoad(depName), {});
                    args[i] = defined[depName];
                } else if (!defining[depName]) {
                    throw new Error(name + ' missing ' + depName);
                }
            }

            ret = callback.apply(defined[name], args);

            if (name) {
                //If setting exports via "module" is in play,
                //favor that over return value and exports. After that,
                //favor a non-undefined return value over exports use.
                if (cjsModule && cjsModule.exports !== undef &&
                    cjsModule.exports !== defined[name]) {
                    defined[name] = cjsModule.exports;
                } else if (ret !== undef || !usingExports) {
                    //Use the return value from the function.
                    defined[name] = ret;
                }
            }
        } else if (name) {
            //May just be an object definition for the module. Only
            //worry about defining if have a module name.
            defined[name] = callback;
        }
    };

    requirejs = require = req = function (deps, callback, relName, forceSync) {
        if (typeof deps === "string") {
            //Just return the module wanted. In this scenario, the
            //deps arg is the module name, and second arg (if passed)
            //is just the relName.
            //Normalize module name, if it contains . or ..
            return callDep(makeMap(deps, callback).f);
        } else if (!deps.splice) {
            //deps is a config object, not an array.
            config = deps;
            if (callback.splice) {
                //callback is an array, which means it is a dependency list.
                //Adjust args if there are dependencies
                deps = callback;
                callback = relName;
                relName = null;
            } else {
                deps = undef;
            }
        }

        //Support require(['a'])
        callback = callback || function () {};

        //Simulate async callback;
        if (forceSync) {
            main(undef, deps, callback, relName);
        } else {
            setTimeout(function () {
                main(undef, deps, callback, relName);
            }, 15);
        }

        return req;
    };

    /**
     * Just drops the config on the floor, but returns req in case
     * the config return value is used.
     */
    req.config = function (cfg) {
        config = cfg;
        return req;
    };

    define = function (name, deps, callback) {

        //This module may not have dependencies
        if (!deps.splice) {
            //deps is not an array, so probably means
            //an object literal or factory function for
            //the value. Adjust args.
            callback = deps;
            deps = [];
        }

        waiting[name] = [name, deps, callback];
    };

    define.amd = {
        jQuery: true
    };
}());

define("build/almond", function(){});

define('fileHandler',[], function(){
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

        compressedFile.saveFile = function (fileName) {
            zip.generateAsync({type: "blob"})
                .then(function(content){
                    if(window.saveAs){
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
            rels += '<Relationship Id="' + _rels.relations[index].Id + '" Type="' + _rels.relations[index].Type + '" Target="' + _rels.relations[index].Target + '"/>';
        }
        rels += '</Relationships>';
        return rels;
    };

    var attach = function (file, _rels) {
        var rels = generateContent(_rels);
        file.addFile(rels, _rels.fileName, _rels.folderName);
    };

    // Create Relation
    var createRelation = function (fileName, folderName) {
        var _rels = {
            relations: []
        };
        _rels.fileName = fileName;
        _rels.folderName = folderName;
        _rels.addRelation = function (id, type, target) {
            addRelation(id, type, target, _rels);
        };
        _rels.generateContent = function () {
            return generateContent(_rels);
        };
        _rels.attach = function (file) {
            return attach(file, _rels);
        }
        return _rels;
    };

    return { createRelation: createRelation };

});
define('oxml_content_types',[], function () {
    'use strict';
    
        var generateContent = function (_contentType) {
            var index, contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
            for (index = 0; index < _contentType.length; index++) {
                contentTypes += '<' + _contentType[index].name + ' ContentType="' + _contentType[index].contentType + '" ';
                if (_contentType[index].Extension) {
                    contentTypes += 'Extension="' + _contentType[index].Extension + '" ';
                }
                if (_contentType[index].PartName) {
                    contentTypes += 'PartName="' + _contentType[index].PartName + '" ';
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
            }
            for (var key in attributes) {
                content[key] = attributes[key];
            }
            _contentType.push(content);
        };

        var createContentType = function () {
            var _contentType = [];
            _contentType.addContentType = function (name, contentType, attributes) {
                addContentType(name, contentType, attributes, _contentType);
            };
            _contentType.generateContent = function () {
                return generateContent(_contentType);
            };
            _contentType.attach = function (file) {
                return attach(file, _contentType);
            }
            return _contentType;
        };

        return {
            createContentType: createContentType
        };

});
define('oxml_sheet',[], function () {
    'use strict';

    var cellString = function (value, cellIndex) {
        if (!value) {
            return '';
        }
        if (value.type === 'numeric') {
            return '<c r="' + cellIndex + '"><v>' + value.value + '</v></c>';
        } else if (value.type === 'sharedString') {
            return '<c r="' + cellIndex + '" t="s"><v>' + value.value + '</v></c>';
        } else if (value.type === "sharedFormula") {
            if (value.formula) {
                return '<c  r="' + cellIndex + '"><f t="shared" ref="' + value.range + '" si="' + value.si + '">' + value.formula + '</f></c>';
            }
            return '<c  r="' + cellIndex + '"><f t="shared" si="' + value.si + '"></f></c>';
        } else if (value.type === 'string') {
            return '<c r="' + cellIndex + '" t="inlineStr"><is><t>' + value.value + '</t></is></c>';
        } else if (value.type === 'formula') {
            var v = (value.value !== null && value.value !== undefined) ? '<v>' + value.value + '</v>' : '';
            return '<c r="' + cellIndex + '"><f>' + value.formula + '</f>' + v + '</c>';
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
                        var cellStr = cellString(value, cellIndex);
                        sheetValues += cellStr;
                    }
                }

                sheetValues += '</row>'
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
                    }
                    else {
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
            }
        }
        return null;
    };

    var updateValuesInRow = function (rowId, values, _sheet) {
        if (!_sheet.values[rowId - 1]) {
            _sheet.values[rowId - 1] = [];
        }
        var index = 0;
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet);
            if (value) {
                _sheet.values[rowId - 1][index] = value;
            }
        }
    };

    var updateValuesInColumn = function (columnId, values, _sheet) {
        var index = 0;
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet);
            if (!_sheet.values[index]) {
                _sheet.values[index] = [];
            }
            if (value) {
                _sheet.values[index][columnId - 1] = value;
            }
        }
    };

    var updateValuesInMatrix = function (values, _sheet) {
        var rowIndex = 0;
        for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
            if (values[rowIndex] !== undefined || values[rowIndex] !== null) {
                if (!_sheet.values[rowIndex]) {
                    _sheet.values[rowIndex] = [];
                }
                if (typeof values[rowIndex] === "object" && values[rowIndex].length >= 0) {
                    var columnIndex = 0;
                    for (columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
                        var value = sanitizeValue(values[rowIndex][columnIndex], _sheet);
                        if (value) {
                            _sheet.values[rowIndex][columnIndex] = value;
                        }
                    }
                }
                else {
                    var value = sanitizeValue(values[rowIndex], _sheet);
                    if (value) {
                        _sheet.values[rowIndex][0] = value;
                    }
                }
            }
        }
    };

    var updateSharedFormula = function (_sheet, formula, fromCell, toCell) {
        var nextId;

        if (!fromCell || !formula || !toCell) {
            return;
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
            range: fromCell + ":" + toCell
        };

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
                    si: nextId
                };
            }
        }

        // Update all cell in column
        else if (toCellChar === fromCellChar) {
            var torowIndex = parseInt(toCellNum, 10);
            for (rowIndex++; rowIndex <= torowIndex; rowIndex++) {
                if (!_sheet.values[rowIndex - 1]) {
                    _sheet.values[rowIndex - 1] = [];
                }

                _sheet.values[rowIndex - 1][columIndex] = {
                    type: "sharedFormula",
                    si: nextId
                };
            }
        }

        return nextId;
    };

    var createSheet = function (sheetName, sheetId, rId, workBook) {
        var _sheet = {
            sheetName: sheetName,
            sheetId: sheetId,
            rId: rId,
            values: [],
            _workBook: workBook
        }
        return {
            _sheet: _sheet,
            generateContent: function () {
                generateContent(_sheet)
            },
            attach: function (file) {
                attach(_sheet, file);
            },
            updateValuesInRow: function (rowId, values) {
                if (!isNaN(rowId) && rowId >= 0 && values && values.length) {
                    updateValuesInRow(rowId, values, _sheet);
                }
            },
            updateValuesInColumn: function (columnId, values) {
                if (!isNaN(columnId) && columnId >= 0 && values && values.length) {
                    updateValuesInColumn(columnId, values, _sheet);
                }
            },
            updateValuesInMatrix: function (values) {
                if (values && values.length) {
                    updateValuesInMatrix(values, _sheet);
                }
            },
            updateSharedFormula: function (formula, fromCell, toCell) {
                return updateSharedFormula(_sheet, formula, fromCell, toCell);
            }
        };
    }
    
    // window.oxml.createSheet = createSheet;
    return { createSheet: createSheet };
});
define('oxml_workbook',['oxml_content_types', 'oxml_rels', 'oxml_sheet'], function (oxmlContentTypes, oxmlRels, oxmlSheet) {
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
        // Add Content Type (TODO)
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
    }

    var createWorkbook = function (xlsxContentTypes, xlsxRels) {
        var _workBook = {
            sheets: [],
            xlsxContentTypes: xlsxContentTypes,
            xlsxRels: xlsxRels,
            createSharedString: createSharedString
        };
        _workBook._rels = oxmlRels.createRelation('workbook.xml.rels', "workbook/_rels");

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
            },
            createSharedString: function (str) {
                return createSharedString(str, _workBook);
            },
            getSharedString: function (str) {
                return getSharedString(str, _workBook);
            }
        };
    };

    return { createWorkbook: createWorkbook };
});
define('oxml_xlsx',['fileHandler', 'oxml_content_types', 'oxml_rels', 'oxml_workbook'], function (fileHandler, oxmlContentTypes, oxmlRels, oxmlWorkBook) {
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
    //Register in the values from the outer closure for common dependencies
    //as local almond modules

    //Use almond's special top-level, synchronous require to trigger factory
    //functions, get the final module value, and export it as the public
    //value.
    return require('oxml_xlsx');
}));
