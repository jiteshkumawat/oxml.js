define([], function () {
    'use strict';

    var cellString = function (value, cellIndex) {
        if (!value) {
            return '';
        }
        var styleString = value.styleIndex ? ' s="' + value.styleIndex + '" ' : '';
        if (value.type === 'numeric') {
            return '<c r="' + cellIndex + '" ' + styleString + '><v>' + value.value + '</v></c>';
        } else if (value.type === 'sharedString') {
            return '<c r="' + cellIndex + '" t="s" ' + styleString + '><v>' + value.value + '</v></c>';
        } else if (value.type === "sharedFormula") {
            if (value.formula) {
                return '<c  r="' + cellIndex + '" ' + styleString + '><f t="shared" ref="' + value.range + '" si="' + value.si + '">' + value.formula + '</f></c>';
            }
            return '<c  r="' + cellIndex + '" ' + styleString + '><f t="shared" si="' + value.si + '"></f></c>';
        } else if (value.type === 'string') {
            return '<c r="' + cellIndex + '" t="inlineStr" ' + styleString + '><is><t>' + value.value + '</t></is></c>';
        } else if (value.type === 'formula') {
            var v = (value.value !== null && value.value !== undefined) ? '<v>' + value.value + '</v>' : '';
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
                        var cellStr = cellString(value, cellIndex);
                        sheetValues += cellStr;
                    }

                    sheetValues += '</row>'
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

    var updateValuesInRow = function (values, _sheet, rowIndex, columnIndex, options) {
        rowIndex--;
        if (!_sheet.values[rowIndex]) {
            _sheet.values[rowIndex] = [];
        }
        var index = 0;
        if (typeof values !== "object" || !values.length) {
            values = [values];
        }
        var styleIndex;
        if (options) {
            options.cellIndices = [];
            for (index = 0; index < values.length; index++) {
                options.cellIndices.push(String.fromCharCode(65 + index + columnIndex - 1) + (rowIndex + 1));
            }
            var _styles = _sheet._workBook.createStyles();
            styleIndex = _styles.addStyles(options);
        }
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet);
            if (value) {
                if (styleIndex) {
                    value.styleIndex = styleIndex.index;
                }
                _sheet.values[rowIndex][index + columnIndex - 1] = value;
            }
        }
    };

    var updateValuesInColumn = function (values, _sheet, rowIndex, columnIndex, options) {
        var index = 0;
        if (typeof values !== "object" || !values.length) {
            values = [values];
        }
        var styleIndex;
        if (options) {
            options.cellIndices = [];
            for (index = 0; index < values.length; index++) {
                options.cellIndices.push(String.fromCharCode(65 + columnIndex - 1) + (rowIndex + index));
            }
            var _styles = _sheet._workBook.createStyles();
            styleIndex = _styles.addStyles(options);
        }
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet), sheetRowIndex = index + rowIndex - 1;
            if (!_sheet.values[sheetRowIndex]) {
                _sheet.values[sheetRowIndex] = [];
            }
            if (value) {
                if (styleIndex) {
                    value.styleIndex = styleIndex.index;
                }
                _sheet.values[sheetRowIndex][columnIndex - 1] = value;
            }
        }
    };

    var updateValuesInMatrix = function (values, _sheet, rowIndex, columnIndex, options) {
        var index, styleIndex;
        if (options) {
            options.cellIndices = [];
            for (index = 0; index < values.length; index++) {
                var index2;
                    for (index2 = 0; index2 < values[index].length; index2++) {
                options.cellIndices.push(String.fromCharCode(65 + columnIndex + index2 - 1) + (rowIndex + index));
                    }
            }
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
                            _sheet.values[sheetRowIndex][index2 + columnIndex - 1] = value;
                        }
                    }
                }
                else {
                    var value = sanitizeValue(values[index], _sheet);
                    if (value) {
                        if (styleIndex) {
                            value.styleIndex = styleIndex.index;
                        }
                        _sheet.values[sheetRowIndex][columnIndex - 1] = value;
                    }
                }
            }
        }
    };

    var updateValueInCell = function (value, _sheet, rowIndex, columnIndex, options) {
        if (value !== undefined && value !== null) {
            var sheetRowIndex = rowIndex - 1;
            if (!_sheet.values[sheetRowIndex]) {
                _sheet.values[sheetRowIndex] = [];
            }
            value = sanitizeValue(value, _sheet);
            if (options) {
                options.cellIndex = String.fromCharCode(65 + columnIndex - 1) + rowIndex;
                var _styles = _sheet._workBook.createStyles();
                value.styleIndex = _styles.addStyles(options).index;
            }
            _sheet.values[sheetRowIndex][columnIndex - 1] = value;
        }
    }

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

    var destroy = function (_sheet) {
        delete _sheet.sheetName;
        delete _sheet.sheetId;
        delete _sheet.rId;
        delete _sheet.values;
        delete _sheet._workBook;
        delete _sheet._sharedFormula;
    };

    var validateIndex = function (index) {
        if (isNaN(index) || index <= 0) {
            return 1;
        }
        return index;
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
            updateValuesInRow: function (values, rowIndex, columnIndex, options) {
                rowIndex = validateIndex(rowIndex);
                columnIndex = validateIndex(columnIndex);
                if (values) {
                    updateValuesInRow(values, _sheet, rowIndex, columnIndex, options);
                }
            },
            updateValuesInColumn: function (values, rowIndex, columnIndex, options) {
                rowIndex = validateIndex(rowIndex);
                columnIndex = validateIndex(columnIndex);
                if (values) {
                    updateValuesInColumn(values, _sheet, rowIndex, columnIndex, options);
                }
            },
            updateValuesInMatrix: function (values, rowIndex, columnIndex, options) {
                rowIndex = validateIndex(rowIndex);
                columnIndex = validateIndex(columnIndex);
                if (values && values.length) {
                    updateValuesInMatrix(values, _sheet, rowIndex, columnIndex, options);
                }
            },
            updateSharedFormula: function (formula, fromCell, toCell) {
                return updateSharedFormula(_sheet, formula, fromCell, toCell);
            },
            updateValueInCell: function (value, rowIndex, columnIndex, options) {
                rowIndex = validateIndex(rowIndex);
                columnIndex = validateIndex(columnIndex);
                return updateValueInCell(value, _sheet, rowIndex, columnIndex, options);
            },
            destroy: function () {
                return destroy(_sheet);
            }
        };
    }

    // window.oxml.createSheet = createSheet;
    return { createSheet: createSheet };
});