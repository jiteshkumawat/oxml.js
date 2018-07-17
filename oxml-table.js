define([], function () {
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
        tableContent += '<tableColumns count="' + _table.columns.length + '">';
        for (var index = 0; index < _table.columns.length; index++)
            tableContent += '<tableColumn name="' + _table.columns[index] + '" id="' + (index + 1) + '"/>';
        tableContent += '</tableColumns>' + '</table>';
        return tableContent;
    };

    var attach = function (_table, file) {
        var content = generateContent(_table);
        file.addFile(content, "table" + _table.rid + ".xml", "workbook/tables");
    };

    var applyOptions = function (options, _table, _sheet) {
        if (options) {
            if (options.filters) {
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

    var addTable = function (displayName, fromCell, toCell, columns, relId, options, _sheet) {
        var _table = {
            rid: relId,
            displayName: displayName,
            columns: columns,
            fromCell: fromCell,
            toCell: toCell
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
            set: function (options) {
                applyOptions(options, _table, _sheet);
            }
        };
    };

    return { addTable: addTable };
});