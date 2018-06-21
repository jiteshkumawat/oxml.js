define([], function () {
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
        if (typeof values !== "object" || !values.length) {
            values = [values];
        }
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet);
            if (value) {
                _sheet.values[rowId - 1][index] = value;
            }
        }
    };

    var updateValuesInColumn = function (columnId, values, _sheet) {
        var index = 0;
        if (typeof values !== "object" || !values.length) {
            values = [values];
        }
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

    var updateValueInCell = function (_sheet, value, rowIndex, columnIndex) {
        if (value !== undefined && value !== null) {
            if (!rowIndex) rowIndex = 1;
            if (!columnIndex) columnIndex = 1;
            if (!_sheet.values[rowIndex - 1]) {
                _sheet.values[rowIndex - 1] = [];
            }
            value = sanitizeValue(value, _sheet);
            _sheet.values[rowIndex - 1][columnIndex - 1] = value;
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

    var destroy = function(_sheet){
        delete _sheet.sheetName;
        delete _sheet.sheetId;
        delete _sheet.rId;
        delete _sheet.values;
        delete _sheet._workBook;
        delete _sheet._sharedFormula;
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
                if (!isNaN(rowId) && rowId > 0 && values) {
                    updateValuesInRow(rowId, values, _sheet);
                }
            },
            updateValuesInColumn: function (columnId, values) {
                if (!isNaN(columnId) && columnId > 0 && values) {
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
            },
            updateValueInCell: function (rowIndex, columnIndex, value) {
                return updateValueInCell(_sheet, value, rowIndex, columnIndex);
            },
            destroy: function(){
                return destroy(_sheet);
            }
        };
    }

    // window.oxml.createSheet = createSheet;
    return { createSheet: createSheet };
});