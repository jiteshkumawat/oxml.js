(function (window) {
    var cellString = function (value, cellIndex) {
        if (!value) {
            return '';
        }
        if (value.type === 'numeric') {
            return '<c r="' + cellIndex + '"><v>' + value.value + '</v></c>';
        } else if (value.type === 'sharedString') {
            return '<c r="' + cellIndex + '" t="s"><v>' + value.value + '</v></c>';
        } else if (value.type === 'string') {
            return '<c r="'+ cellIndex + '" t="inlineStr"><is><t>' + value.value + '</t></is></c>';
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
            if (value.value === undefined || value.value === null) {
                return null;
            }
            if (!value.type) {
                value = value.value;
            } else {
                if (value.type === "sharedString") {
                    if(typeof value.value === "number"){
                        return {
                            type: value.type,
                            value: value.value
                        }
                    }
                    else{
                        // Add shared string
                        value.value = _sheet._workBook.createSharedString(value.value, _sheet._workBook);
                        return {
                            type: value.type,
                            value: value.value
                        }
                    }
                }
                else if (value.type === "numeric" || value.type === "string") {
                    return {
                        type: value.type,
                        value: value.value
                    }
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
        if (!_sheet.values[rowId]) {
            _sheet.values[rowId] = [];
        }
        var index = 0;
        for (index = 0; index < values.length; index++) {
            var value = sanitizeValue(values[index], _sheet);
            if (value) {
                _sheet.values[rowId][index] = value;
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
                _sheet.values[index][columnId] = value;
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
                if (values[rowIndex].length) {
                    var columnIndex = 0;
                    for (columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
                        var value = sanitizeValue(values[rowIndex][columnIndex], _sheet);
                        if (value) {
                            _sheet.values[rowIndex][columnIndex] = value;
                        }
                    }
                }
                else if (values[rowIndex].length === undefined || values[rowIndex].length === null) {
                    var value = sanitizeValue(values[rowIndex], _sheet);
                    if (value) {
                        _sheet.values[rowIndex][0] = value;
                    }
                }
            }
        }
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
            }
        };
    }
    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createSheet = createSheet;
})(window);