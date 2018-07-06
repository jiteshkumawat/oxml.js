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
            columnIndex: columnIndex + 1,
            set: function (value, options) {
                cell(_sheet, rowIndex + 1, columnIndex + 1, value, options);
                this.value = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].value : null;
                this.type = _sheet.values[rowIndex] && _sheet.values[rowIndex][columnIndex] ? _sheet.values[rowIndex][columnIndex].type : null;
                return getCellAttributes(_sheet, cellIndex, rowIndex, columnIndex);
            }
        };
    };

    var getCellRangeAttributes = function (_sheet, cellIndices, _cells, cRowIndex, cColumnIndex, totalRows, totalColumns, isRow) {
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
                columnIndex: columnIndex,
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
                if (!values || !values.length) return;
                if (isRow) {
                    totalColumns = totalColumns > values.length ? totalColumns : values.length;
                    values = [values];
                } else {
                    totalRows = totalRows > values.length ? totalRows : values.length;
                }
                return cells(_sheet, cRowIndex, cColumnIndex, totalRows, totalColumns, values, options, isRow);
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
                    si: nextId
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
                    si: nextId
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

    var validateIndex = function (index) {
        if (isNaN(index) || index <= 0) {
            return 1;
        }
        return index;
    };

    var cell = function (_sheet, rowIndex, columnIndex, value, options) {
        if (typeof rowIndex !== "number" || typeof columnIndex !== "number")
            return;
        if (!options && (value === undefined || value === null)) {
            var cellIndex = String.fromCharCode(65 + columnIndex - 1) + rowIndex;
            return getCellAttributes(_sheet, cellIndex, rowIndex - 1, columnIndex - 1);
        } else if (!options && typeof value === "object" && !value.type && (value.value === undefined || value.value === null)) {
            return cellStyle(_sheet, rowIndex, columnIndex, value);
        } else if (value === undefined || value === null) {
            return cellStyle(_sheet, rowIndex, columnIndex, options);
        } else {
            rowIndex = validateIndex(rowIndex);
            columnIndex = validateIndex(columnIndex);
            return updateValueInCell(value, _sheet, rowIndex, columnIndex, options);
        }
    };

    var cells = function (_sheet, rowIndex, columnIndex, totalRows, totalColumns, values, options, isRow) {
        if (typeof rowIndex !== "number" || typeof columnIndex !== "number" || typeof totalRows !== "number" || typeof totalColumns !== "number")
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
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow);
        } else if (!options && values && !values.length) {
            values.cellIndices = cellIndices;
            updateRangeStyle(_sheet, values, cells);
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow);
        } else if (values === undefined || values === null) {
            options.cellIndices = cellIndices;
            updateRangeStyle(_sheet, options, cells);
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow);
        } else {
            rowIndex = validateIndex(rowIndex);
            columnIndex = validateIndex(columnIndex);
            updateValuesInMatrix(values, _sheet, rowIndex, columnIndex, options, cellIndices, cells, totalRows, totalColumns);
            return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow);
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
            updateValuesInMatrix: function (values, rowIndex, columnIndex, options) {
                rowIndex = validateIndex(rowIndex);
                columnIndex = validateIndex(columnIndex);
                if (values && values.length) {
                    return updateValuesInMatrix(values, _sheet, rowIndex, columnIndex, options);
                }
            },
            updateSharedFormula: function (formula, fromCell, toCell, options) {
                return updateSharedFormula(_sheet, formula, fromCell, toCell, options);
            },
            cell: function (rowIndex, columnIndex, value, options) {
                return cell(_sheet, rowIndex, columnIndex, value, options);
            },
            row: function (rowIndex, columnIndex, values, options) {
                var totalColumns = _sheet.values[rowIndex - 1] ? _sheet.values[rowIndex - 1].length : 0;
                return cells(_sheet, rowIndex, columnIndex, 1, values ? values.length : totalColumns, values ? [values] : null, options, true);
            },
            column: function (rowIndex, columnIndex, values, options) {
                var index, totalRows = 0;
                if (!values || !values.length) {
                    for (index = 0; index < rowIndex; index++) {
                        if (_sheet.values[index])
                            totalRows = totalRows < _sheet.values[index].length ? _sheet.values[index] : totalRows;
                    }
                }
                return cells(_sheet, rowIndex, columnIndex, values ? values.length : totalRows, 1, values, options, false);
            },
            getRange: function (rowIndex, columnIndex, totalRows, totalColumns) {
                var cellIndices = [], cells = [];
                for (var index = 0; index < totalRows; index++) {
                    var index2;
                    for (index2 = 0; index2 < totalColumns; index2++) {
                        cellIndices.push(String.fromCharCode(65 + columnIndex + index2 - 1) + (rowIndex + index));
                        cells.push({ rowIndex: index + rowIndex - 1, columnIndex: index2 + columnIndex - 1 });
                    }
                }
                return getCellRangeAttributes(_sheet, cellIndices, cells);
            },
            destroy: function () {
                return destroy(_sheet);
            }
        };
    };

    return { createSheet: createSheet };
});