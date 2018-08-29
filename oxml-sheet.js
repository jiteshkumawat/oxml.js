define(['oxml_table', 'oxml_rels', 'xmlContentString', 'contentFile'],
    function (
        oxmlTable,
        OxmlRels,
        XMLContentString,
        ContentFile) {
        'use strict';

        var Sheet = function (sheetName, sheetId, rId, workBook, xlsxContentTypes) {
            this.className = "Worksheet";
            this.fileName = "sheet" + sheetId + ".xml";
            this.folderName = "workbook/sheets";
            this.xlsxContentTypes = xlsxContentTypes;
            this._sheet = {
                sheetName: sheetName,
                sheetId: sheetId,
                rId: rId,
                values: [],
                _workBook: workBook
            };
        };
        Sheet.prototype = Object.create(ContentFile.prototype);

        var cellString = function (value, cellIndex, rowIndex, columnIndex) {
            if (!value) {
                return '';
            }
            var styleString = value.styleIndex ? 's="' + value.styleIndex + '"' : '';
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
        
        var getNextRelation = function (_sheet) {
            var lastSheetRel = {};
            if (_sheet._rels.relations.length) {
                lastSheetRel = _sheet._rels.relations[_sheet._rels.relations.length - 1];
            }
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
            return nextSheetRelId;
        };

        Sheet.prototype.generateContent = function (file) {
            var rowIndex = 0, sheetValues = '<sheetData>';
            if (this._sheet.values && this._sheet.values.length) {
                for (rowIndex = 0; rowIndex < this._sheet.values.length; rowIndex++) {
                    if (this._sheet.values[rowIndex] && this._sheet.values[rowIndex].length > 0) {
                        var hidden = this._sheet.values[rowIndex].hidden ? ' hidden="1" ' : '';
                        sheetValues += '<row r="' + (rowIndex + 1) + '"' + hidden + '>';

                        var columnIndex = 0;
                        for (columnIndex = 0; columnIndex < this._sheet.values[rowIndex].length; columnIndex++) {
                            var columnChar = String.fromCharCode(65 + columnIndex);
                            var value = this._sheet.values[rowIndex][columnIndex];
                            var cellIndex = columnChar + (rowIndex + 1);
                            var cellStr = cellString(value, cellIndex, rowIndex, columnIndex);
                            sheetValues += cellStr;
                        }

                        sheetValues += '</row>';
                    }
                }
            }
            sheetValues += '</sheetData>';
            var tables = '';
            if (this._sheet.tables && this._sheet.tables.length) {
                tables += '<tableParts count="' + this._sheet.tables.length + '">';
                for (var index = 0; index < this._sheet.tables.length; index++) {
                    tables += '<tablePart r:id="rId' + this._sheet.tables[index].rid + '"/>';
                    if (file) this._sheet.tables[index].attach(file);
                }
                tables += '</tableParts>';
            }
            var sheetTemplate = new XMLContentString({
                rootNode: 'worksheet',
                nameSpaces: [
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    {
                        value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        attribute: "r"
                    }
                ]
            });
            return sheetTemplate.format(sheetValues + tables);
        };

        Sheet.prototype.attachChild = function (file) {
            if (this._sheet._rels) {
                this._sheet._rels.attach(file);
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
                        options.cellIndex = this.cellIndex;
                        updateSingleStyle(_sheet, options, (this.rowIndex - 1), (this.columnIndex.toUpperCase().charCodeAt() - 65));
                        return getCellAttributes(_sheet, this.cellIndex, (this.rowIndex - 1), (this.columnIndex.toUpperCase().charCodeAt() - 65));
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

        Sheet.prototype.sharedFormula = function (fromCell, toCell, formula, options) {
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

            if (!this._sheet._sharedFormula) {
                this._sheet._sharedFormula = [];
                nextId = 0;
            } else {
                var lastSharedFormula = this._sheet._sharedFormula[this._sheet._sharedFormula.length - 1];
                nextId = lastSharedFormula.si + 1;
            }

            this._sheet._sharedFormula.push({
                si: nextId,
                fromCell: fromCell,
                toCell: toCell,
                formula: formula
            });

            // Update from Cell
            var columIndex = fromCellChar.toUpperCase().charCodeAt() - 65;
            var rowIndex = parseInt(fromCellNum, 10);
            if (!this._sheet.values[rowIndex - 1]) {
                this._sheet.values[rowIndex - 1] = [];
            }
            this._sheet.values[rowIndex - 1][columIndex] = {
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
                    this._sheet.values[rowIndex - 1][columIndex] = {
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
                    if (!this._sheet.values[rowIndex - 1]) {
                        this._sheet.values[rowIndex - 1] = [];
                    }
                    cellIndices.push(String.fromCharCode(65 + columIndex) + rowIndex);
                    cells.push({ rowIndex: rowIndex - 1, columnIndex: columIndex });

                    this._sheet.values[rowIndex - 1][columIndex] = {
                        type: "sharedFormula",
                        si: nextId,
                        value: val
                    };
                }
            }
            if (options) {
                options.cellIndices = cellIndices;
                var styleIndex, _styles = this._sheet._workBook.createStyles();
                styleIndex = _styles.addStyles(options);
                for (var cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                    this._sheet.values[cells[cellIndex].rowIndex][cells[cellIndex].columnIndex].styleIndex = styleIndex.index;
                }
            }
            return getCellRangeAttributes(this._sheet, cellIndices, cells);
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

        Sheet.prototype.cell = function (rowIndex, columnIndex, value, options) {
            return cell(this._sheet, rowIndex, columnIndex, value, options);
        };

        Sheet.prototype.row = function (rowIndex, columnIndex, values, options) {
            var totalColumns = this._sheet.values[rowIndex - 1] ? this._sheet.values[rowIndex - 1].length - columnIndex + 1 : 0;
            return cells(this._sheet, rowIndex, columnIndex, 1, ((values && Array.isArray(values)) ? values.length : (values ? 1 : totalColumns)), values ? [values] : null, options, true, true, false);
        };

        Sheet.prototype.column = function (rowIndex, columnIndex, values, options) {
            var totalRows = 0;
            if (!values || !values.length) {
                totalRows = this._sheet.values && this._sheet.values.length && this._sheet.values.length > rowIndex ? this._sheet.values.length - rowIndex + 1 : 0;
            }
            return cells(this._sheet, rowIndex, columnIndex, ((values && Array.isArray(values)) ? values.length : (values ? 1 : totalRows)), 1, values, options, true, false, true);
        };

        Sheet.prototype.grid = function (rowIndex, columnIndex, values, options) {
            var index, totalRows = 0, totalColumns = 0;
            if (values && values.length) {
                totalRows = values.length;
                for (index = 0; index < values.length; index++) {
                    if (values[index] && values[index].length)
                        totalColumns = totalColumns < values[index].length ? values[index].length : totalColumns;
                }
            } else if (this._sheet.values && this._sheet.values.length) {
                totalRows = this._sheet.values.length - rowIndex + 1;
                for (var index = 0; index < this._sheet.values.length; index++) {
                    if (this._sheet.values[index] && this._sheet.values[index].length) {
                        totalColumns = totalColumns < this._sheet.values[index].length - columnIndex + 1 ? this._sheet.values[index].length - columnIndex + 1 : totalColumns;
                    }
                }
            }

            return cells(this._sheet, rowIndex, columnIndex, totalRows, totalColumns, values, options, true, false);
        };

        Sheet.prototype.table = function (tableName, fromCell, toCell, options) {
            var _table = addTable(this._sheet, this.xlsxContentTypes, tableName, fromCell, toCell, options);
            return (_table ? _table.tableOptions() : undefined);
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
            }
            // ---Deprecated---
            // else if (!options && values && !values.length) {
            //     values.cellIndices = cellIndices;
            //     updateRangeStyle(_sheet, values, cells);
            //     if (!isReturn) return;
            //     return getCellRangeAttributes(_sheet, cellIndices, cells, rowIndex, columnIndex, totalRows, totalColumns, isRow, isColumn);
            // } 
            else if (values === undefined || values === null) {
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
            if (!titleRow) return;
            for (var index = fromColumIndex; index <= toColumnIndex; index++) {
                titles.push(titleRow[index] && titleRow[index].value ? titleRow[index].value : '');
            }
            var relId = addFile("table", xlsxContentTypes, _sheet);

            if (!_sheet.tables) _sheet.tables = [];
            var table = oxmlTable.addTable(tableName, fromCell, toCell, titles, relId, options, _sheet);
            _sheet.tables.push(table);
            return table;
        };

        var addFile = function (fileType, xlsxContentTypes, _sheet) {
            if (!_sheet._rels) _sheet._rels = OxmlRels.createRelation("sheet" + _sheet.sheetId + ".xml.rels", "workbook/sheets/_rels");
            var relId = getNextRelation(_sheet), type, target, contentType, partName;

            switch (fileType) {
                case "table":
                    type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
                    target = "../tables/table" + relId + ".xml";
                    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
                    partName = "/workbook/tables/table" + relId + ".xml";
                    break;
                case "drawing":
                    type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
                    target = "../drawings/drawing" + relId + ".xml";
                    contentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
                    partName = "/workbook/drawings/drawing" + relId + ".xml";
                    break;
            }
            _sheet._rels.addRelation("rId" + relId, type, target);
            xlsxContentTypes.addContentType("Override", contentType, { PartName: partName });
            return relId;
        };

        return {
            createSheet: function (sheetName, sheetId, rId, workBook, xlsxContentTypes) {
                return new Sheet(sheetName, sheetId, rId, workBook, xlsxContentTypes);
            }
        };
    });