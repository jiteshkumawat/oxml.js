(function (window) {
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
                        var isNumber = !isNaN(value);
                        if (value !== null && value !== undefined) {
                            var cellIndex = columnChar + (rowIndex + 1);
                            if (isNumber) {
                                sheetValues += `<c r="${cellIndex}">
                            <v>${value}</v>
                        </c>
                        `;
                            }
                            else if (value.type === 'sharedString'){
                                sheetValues += `<c r="${cellIndex}" t="s">
                            <v>${value.value}</v>
                        </c>
                        `;
                            }
                            else {
                                sheetValues += `<c r="${cellIndex}" t="inlineStr">
                            <is><t>${value}</t></is>
                        </c>
                        `;
                            }
                        }
                    }
                }

                sheetValues += '</row>'
            }
        }
        var sheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                ${sheetValues}
            </sheetData>
        </worksheet>`
        return sheet;
    };

    var attach = function (_sheet, file) {
        var content = generateContent(_sheet);
        file.addFile(content, "sheet" + _sheet.sheetId + ".xml", "workbook/sheets");
    };

    var updateValuesInRow = function (rowId, values, _sheet) {
        if (!_sheet.values[rowId]) {
            _sheet.values[rowId] = [];
        }
        var index = 0;
        for (index = 0; index < values.length; index++) {
            if (values[index] !== undefined && values[index] !== null) {
                if (values[index].type === 'sharedString' || values[index].type === 'sharedstring') {
                    if (values[index].value !== undefined && values[index].value !== null) {
                        _sheet.values[rowId][index] = {
                            type: "sharedString",
                            value: values[index].value
                        }
                    }
                } else {
                    _sheet.values[rowId][index] = values[index];
                }
            }
        }
    };

    var updateValuesInColumn = function (columnId, values, _sheet) {
        var index = 0;
        for (index = 0; index < values.length; index++) {
            if (values[index] !== undefined || values[index] !== null) {
                if (!_sheet.values[index]) {
                    _sheet.values[index] = [];
                }
                if (values[index].type === 'sharedString' || values[index].type === 'sharedstring') {
                    if (values[index].value !== undefined && values[index].value !== null) {
                        _sheet.values[index][columnId] = {
                            type: "sharedString",
                            value: values[index].value
                        }
                    }
                }
                else {
                    _sheet.values[index][columnId] = values[index];
                }
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
                        var value = values[rowIndex][columnIndex];
                        if (value !== undefined && value !== null) {
                            if (value.type === 'sharedString' || value.type === 'sharedstring') {
                                if (value.value !== undefined && value.value !== null) {
                                    _sheet.values[rowIndex][columnIndex] = {
                                        type: "sharedString",
                                        value: value.value
                                    }
                                }
                            }
                            else
                            {
                            _sheet.values[rowIndex][columnIndex] = value;
                            }
                        }
                    }
                }
                else if (values[rowIndex].length === undefined || values[rowIndex].length === null) {
                    _sheet.values[rowIndex][0] = values[rowIndex];
                }
            }
        }
    };

    var createSheet = function (sheetName, sheetId, rId) {
        var _sheet = {
            sheetName: sheetName,
            sheetId: sheetId,
            rId: rId,
            values: []
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