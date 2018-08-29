define(['xmlContentString', 'contentFile'],
    function (
        XMLContentString,
        ContentFile
    ) {
        'use strict';

        var Table = function (displayName, fromCell, toCell, columns, relId, options, _sheet) {
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
            this.className = "Table";
            this.fileName = "table" + relId + ".xml";
            this.folderName = "workbook/tables";
            this._table = _table;
            this.rid = relId;
            this._sheet = _sheet;
        };
        Table.prototype = Object.create(ContentFile.prototype);

        Table.prototype.generateContent = function () {
            var template = new XMLContentString({
                rootNode: 'table',
                nameSpaces: ["http://schemas.openxmlformats.org/spreadsheetml/2006/main"],
                attributes: {
                    totalsRowShown: "0",
                    ref: this._table.fromCell + ":" + this._table.toCell,
                    displayName: this._table.displayName,
                    name: this._table.displayName,
                    id: this._table.rid
                }
            });

            var tableContent = '';
            if (this._table.filters) {
                tableContent += '<autoFilter ref="' + this._table.fromCell +
                    ':' + this._table.toCell + '">';
                for (var index = 0; index < this._table.filters.length; index++) {
                    tableContent += '<filterColumn colId="' + this._table.filters[index].column + '">';
                    if (this._table.filters[index].values[0].type === "default") {
                        tableContent += '<filters>';
                        for (var index2 = 0; index2 < this._table.filters[index].values.length; index2++) {
                            tableContent += '<filter val="' + this._table.filters[index].values[index2].value + '"/>';
                        }
                        tableContent += '</filters>';
                    } else if (this._table.filters[index].values[0].type === "custom") {
                        tableContent += '<customFilters and="' + (this._table.filters[index].values[0].and ? '1' : '0') + '">';
                        for (var index2 = 0; index2 < this._table.filters[index].values.length; index2++) {
                            tableContent += '<customFilter operator="' + this._table.filters[index].values[index2].operator + '" val="' + this._table.filters[index].values[index2].value + '"/>';
                        }
                        tableContent += '</customFilters>';
                    }
                    tableContent += '</filterColumn>';
                }
                tableContent += '</autoFilter>';
            }
            if (this._table.sort) {
                tableContent += '<sortState caseSensitive="' + (this._table.sort.caseSensitive ? '1' : '0') + '" ref="' + this._table.sort.dataRange + '">';
                tableContent += '<sortCondition' + (this._table.sort.direction === "ascending" ? "" : ' descending="1"') + ' ref="' + this._table.sort.range + '"/>';
                tableContent += '</sortState>';
            }
            tableContent += '<tableColumns count="' + this._table.columns.length + '">';
            for (var index = 0; index < this._table.columns.length; index++)
                tableContent += '<tableColumn name="' + this._table.columns[index] + '" id="' + (index + 1) + '"/>';
            var tableStyle = '';
            if (this._table.tableStyle)
                tableStyle = '<tableStyleInfo name="' + this._table.tableStyle.name
                    + '" showColumnStripes="' + (this._table.tableStyle.showColumnStripes ? '1' : '0')
                    + '" showRowStripes="' + (this._table.tableStyle.showRowStripes ? '1' : '0') + '" showLastColumn="'
                    + (this._table.tableStyle.showLastColumn ? '1' : '0') + '" showFirstColumn="'
                    + (this._table.tableStyle.showFirstColumn ? '1' : '0') + '"/>';
            tableContent += '</tableColumns>' + tableStyle;
            return template.format(tableContent);
        };

        Table.prototype.tableOptions = function () {
            return tableOptions(this._table, this._sheet, this.relId);
        }

        var applyFilter = function (options, _table, _sheet) {
            if (!_table.filters) _table.filters = [];
            if (typeof options.filters === "object" && options.filters.length) {
                for (var index = 0; index < options.filters.length; index++) {
                    if (options.filters[index].column) {
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
            var savedStyle = _styles.addTableStyle(options, tableStyleName, _table);
            _table.tableStyle = {
                name: tableStyleName,
                showColumnStripes: !!(savedStyle.evenColumn || savedStyle.oddColumn),
                showRowStripes: !!(savedStyle.evenRow || savedStyle.oddRow),
                showLastColumn: !!savedStyle.lastColumn,
                showFirstColumn: !!savedStyle.firstColumn
            };
        };

        var tableOptions = function (_table, _sheet, relId) {
            return {
                name: _table.displayName,
                columns: _table.columns,
                fromCell: _table.fromCell,
                toCell: _table.toCell,
                sort: _table.sort,
                filters: _table.filters,
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

        return {
            addTable: function (displayName, fromCell, toCell, columns, relId, options, _sheet) {
                return new Table(displayName, fromCell, toCell, columns, relId, options, _sheet);
            }
        };
    });