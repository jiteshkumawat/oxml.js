define(['utils',
    'oxml_xlsx_font',
    'oxml_xlsx_num_format',
    'oxml_xlsx_border',
    'oxml_xlsx_fill',
    'xmlContentString',
    'contentFile'],
    function (utils,
        oxmlXlsxFont,
        oxmlXlsxNumFormat,
        oxmlXlsxBorder,
        oxmlXlsxFill,
        XMLContentString,
        ContentFile) {

        "use strict";
        var Styles = function (_workbook, _rel, _contentType) {
            var sheetId = createStyleParts(_workbook, _rel, _contentType);
            var _styles = {
                sheetId: sheetId,
                fileName: "style" + sheetId + ".xml",
                styles: [],
                _fonts: {},
                _borders: {},
                _fills: {},
                _fontsCount: 1,
                _bordersCount: 1,
                _fillsCount: 2
            };
            var font = {
                bold: false,
                italic: false,
                underline: false,
                size: false,
                color: false,
                strike: false,
                family: false,
                name: false,
                scheme: false
            };
            _styles._fonts[utils.stringify(font)] = 0;
            _styles._borders[false] = 0;
            _styles._fills[false] = 0;
            var fillGray125 = {
                pattern: 'gray125',
                fgColor: false,
                bgColor: false
            };
            _styles._fills[utils.stringify(fillGray125)] = 1;
            this.className = "Style";
            this.fileName = _styles.fileName;
            this.folderName = "workbook";
            this._styles = _styles;
        };
        Styles.prototype = Object.create(ContentFile.prototype);

        Styles.prototype.generateContent = function () {
            // Create Styles
            var template = new XMLContentString({
                rootNode: "styleSheet",
                nameSpaces: ["http://schemas.openxmlformats.org/spreadsheetml/2006/main"]
            });
            var stylesString = '';
            var index = 0;
            stylesString += oxmlXlsxNumFormat.generateContent(this._styles);
            stylesString += oxmlXlsxFont.generateContent(this._styles);
            stylesString += oxmlXlsxFill.generateContent(this._styles);
            stylesString += oxmlXlsxBorder.generateContent(this._styles);
            stylesString += '<cellStyleXfs count="1"><xf /></cellStyleXfs>';

            stylesString += '<cellXfs count="'
                + (parseInt(this._styles.styles.length, 10) + 1) + '"><xf />';
            for (index = 0; index < this._styles.styles.length; index++) {
                var numFormatString = this._styles.styles[index]._numFormat
                    ? ' numFmtId="'
                    + this._styles.styles[index]._numFormat + '" ' : '';
                var borderString = this._styles.styles[index]._border
                    ? ' borderId="' + this._styles.styles[index]._border + '" ' : '';
                var fillString = this._styles.styles[index]._fill
                    ? ' fillId="' + this._styles.styles[index]._fill + '" ' : '';
                stylesString += '<xf fontId="' + this._styles.styles[index]._font
                    + '" ' + numFormatString
                    + borderString + fillString + ' />';
            }
            var tableStyles = '';
            if (this._styles.tableStyles && this._styles.tableStyles.length) tableStyles = generateTableContent(this._styles);
            stylesString += '</cellXfs>' + tableStyles;
            return template.format(stylesString);
        };

        var generateTableContent = function (_styles) {
            var dxfsCount = 0, tableStyles = '', tableStylesCount = 0, tableStyleCount, dxfString = '';
            for (var index = 0; index < _styles.tableStyles.length; index++) {
                var individualTalbeStyle;
                dxfsCount++;
                tableStylesCount++;
                tableStyleCount = 1;
                dxfString += generateDxfContent(_styles.tableStyles[index]);
                individualTalbeStyle = '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="wholeTable"/>';
                if (_styles.tableStyles[index].evenColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="secondColumnStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].evenColumn);
                }
                if (_styles.tableStyles[index].oddColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstColumnStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].oddColumn);
                }
                if (_styles.tableStyles[index].evenRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="secondRowStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].evenRow);
                }
                if (_styles.tableStyles[index].oddRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstRowStripe"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].oddRow);
                }
                if (_styles.tableStyles[index].firstColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstColumn"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstColumn);
                }
                if (_styles.tableStyles[index].lastColumn) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastColumn"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastColumn);
                }
                if (_styles.tableStyles[index].firstRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="headerRow"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRow);
                }
                if (_styles.tableStyles[index].lastRow) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="totalRow"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRow);
                }
                if (_styles.tableStyles[index].firstRowFirstCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstHeaderCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRowFirstCell);
                }
                if (_styles.tableStyles[index].firstRowLastCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastHeaderCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].firstRowLastCell);
                }
                if (_styles.tableStyles[index].lastRowFirstCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="firstTotalCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRowFirstCell);
                }
                if (_styles.tableStyles[index].lastRowLastCell) {
                    dxfsCount++;
                    tableStyleCount++;
                    individualTalbeStyle += '<tableStyleElement dxfId="' + (dxfsCount - 1) + '" type="lastTotalCell"/>';
                    dxfString += generateDxfContent(_styles.tableStyles[index].lastRowLastCell);
                }
                tableStyles += '<tableStyle count="' + tableStyleCount + '" name="' + _styles.tableStyles[index].name + '">' + individualTalbeStyle + '</tableStyle>';
            }
            tableStyles = '<tableStyles count="' + tableStylesCount + '">' + tableStyles + '</tableStyles>';
            dxfString = '<dxfs count="' + dxfsCount + '">' + dxfString + '</dxfs>' + tableStyles;
            return dxfString;
        };

        var generateDxfContent = function (style) {
            var dxfString = '<dxf>';
            if (style.font) {
                dxfString += oxmlXlsxFont.generateSingleContent(style.font);
            }
            if (style.fill) {
                dxfString += oxmlXlsxFill.generateSingleContent(style.fill);
            }
            if (style.border) {
                dxfString += oxmlXlsxBorder.generateSingleContent(style.border);
            }
            dxfString += '</dxf>';
            return dxfString;
        };

        var createStyleParts = function (_workBook, _rel, _contentType) {
            var lastSheetRel = _workBook.getLastSheet();
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0")
                .replace("rId", ""), 10) + 1;
            _contentType
                .addContentType("Override",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                    {
                        PartName: "/workbook/style" + nextSheetRelId + ".xml"
                    });

            _rel.addRelation(
                "rId" + nextSheetRelId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "style" + nextSheetRelId + ".xml"
            );

            return nextSheetRelId;
        };

        var searchStyleForCell = function (_styles, cellIndex) {
            var index = 0;
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined
                    && _styles.styles[index].cellIndices[cellIndex] !== null) {
                    return _styles.styles[index];
                }
            }
            return null;
        };

        var searchSimilarStyle = function (_styles, style) {
            var index = 0;
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index]._font === style._font
                    && _styles.styles[index]._numFormat === style._numFormat) {
                    return _styles.styles[index];
                }
            }
            return null;
        };

        Styles.prototype.addStyles = function (options) {
            if (options.cellIndex || options.cellIndices) {
                var newStyleCreated = false, cellStyle,
                    saveFont, saveNumFormat, saveBorder, saveFill;
                if (options.cellIndex) {
                    cellStyle = searchStyleForCell(this._styles, options.cellIndex);
                    saveFont = oxmlXlsxFont
                        .getFontForCell(this._styles, options, cellStyle);
                    saveNumFormat = oxmlXlsxNumFormat
                        .getNumFormatForCell(this._styles, options, cellStyle);
                    saveBorder = oxmlXlsxBorder
                        .getBorderForCell(this._styles, options, cellStyle);
                    saveFill = oxmlXlsxFill
                        .getFillForCell(this._styles, options, cellStyle);
                    newStyleCreated = newStyleCreated
                        || saveFont.newStyleCreated
                        || saveBorder.newStyleCreated
                        || saveFill.newStyleCreated;
                    if (cellStyle) {
                        if (cellStyle._font === saveFont.fontIndex
                            && cellStyle._numFormat === saveNumFormat.numFormatIndex
                            && cellStyle._border === saveBorder.borderIndex
                            && cellStyle._fill === saveFill.fillIndex) {
                            return cellStyle;
                        }
                        var totalCellApplied = Object.keys(cellStyle.cellIndices).length;
                        if (totalCellApplied === 1) {
                            cellStyle._font = saveFont.fontIndex;
                            cellStyle._numFormat = saveNumFormat.numFormatIndex;
                            cellStyle._border = saveBorder.borderIndex;
                            cellStyle._fill = saveFill.fillIndex;
                            return cellStyle;
                        }
                        delete cellStyle.cellIndices[options.cellIndex];
                    }

                    if (!newStyleCreated) {
                        cellStyle = searchSimilarStyle(this._styles, {
                            _font: saveFont.fontIndex,
                            _numFormat: saveNumFormat.numFormatIndex,
                            _border: saveBorder.borderIndex,
                            _fill: saveFill.fillIndex
                        });
                        if (cellStyle) {
                            cellStyle.cellIndices[options.cellIndex] = Object.keys(cellStyle.cellIndices).length;
                            return cellStyle;
                        }
                    }

                    cellStyle = {
                        _font: saveFont.fontIndex,
                        _numFormat: saveNumFormat.numFormatIndex || false,
                        _border: saveBorder.borderIndex,
                        _fill: saveFill.fillIndex,
                        cellIndices: {}
                    };

                    cellStyle.cellIndices[options.cellIndex] = 0;
                    cellStyle.index = "" + (parseInt(this._styles.styles.length, 10) + 1);
                    this._styles.styles.push(cellStyle);
                    return cellStyle;
                }

                if (options.cellIndices) {
                    var index;
                    saveFont = oxmlXlsxFont.getFontForCells(this._styles, options);
                    saveNumFormat = oxmlXlsxNumFormat.getNumFormatForCells(this._styles, options);
                    saveBorder = oxmlXlsxBorder.getBorderForCells(this._styles, options);
                    saveFill = oxmlXlsxFill.getFillForCells(this._styles, options);
                    newStyleCreated = newStyleCreated || saveFont.newStyleCreated
                        || saveNumFormat.newStyleCreated || saveBorder.newStyleCreated || saveFill.newStyleCreated;
                    if (!newStyleCreated) {
                        cellStyle = searchSimilarStyle(this._styles, {
                            _font: saveFont.fontIndex,
                            _numFormat: saveNumFormat.numFormatIndex,
                            _border: saveBorder.borderIndex,
                            _fill: saveFill.fillIndex
                        });

                        if (cellStyle) {
                            for (index = 0; index < options.cellIndices.length; index++) {
                                var cellIndex = options.cellIndices[index];
                                var savedCellStyle = searchStyleForCell(this._styles, cellIndex);
                                if (savedCellStyle) delete savedCellStyle.cellIndices[cellIndex];
                                cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                            }
                            return cellStyle;
                        }
                    }
                    // Maintain old styling

                    cellStyle = {
                        _font: saveFont.fontIndex,
                        _numFormat: saveNumFormat.numFormatIndex,
                        _border: saveBorder.borderIndex,
                        _fill: saveFill.fillIndex,
                        cellIndices: {}
                    };

                    for (index = 0; index < options.cellIndices.length; index++) {
                        var cellIndex = options.cellIndices[index];
                        var savedCellStyle = searchStyleForCell(this._styles, cellIndex);
                        if (savedCellStyle) {
                            delete savedCellStyle.cellIndices[cellIndex];
                        }
                        cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                        if (savedCellStyle && !Object.keys(savedCellStyle.cellIndices).length) {
                            for (var key in cellStyle) {
                                savedCellStyle[key] = cellStyle[key];
                            }
                            return savedCellStyle;
                        }
                    }
                    cellStyle.index = "" + (parseInt(this._styles.styles.length, 10) + 1);
                    this._styles.styles.push(cellStyle);
                    return cellStyle;
                }
            }
        };

        Styles.prototype.addTableStyle = function (options, tableStyleName, _table) {
            if (!this._styles.tableStyles) this._styles.tableStyles = [];

            // Find existing table style
            var existingTableStyle;
            if (_table && _table.tableStyle) {
                for (var index = 0; index < this._styles.tableStyles.length; index++) {
                    if (_table.tableStyle.name === this._styles.tableStyles[index].name) {
                        existingTableStyle = this._styles.tableStyles[index];
                        break;
                    }
                }
            }

            var tableStyle = prepareTableStyleObj(options, existingTableStyle);
            tableStyle.name = tableStyleName;
            if (options.evenColumn) tableStyle.evenColumn = prepareTableStyleObj(options.evenColumn, existingTableStyle && existingTableStyle.evenColumn)
            else if (existingTableStyle && existingTableStyle.evenColumn) tableStyle.evenColumn = existingTableStyle.evenColumn;
            if (options.oddColumn) tableStyle.oddColumn = prepareTableStyleObj(options.oddColumn, existingTableStyle && existingTableStyle.oddColumn);
            else if (existingTableStyle && existingTableStyle.oddColumn) tableStyle.oddColumn = existingTableStyle.oddColumn;
            if (options.evenRow) tableStyle.evenRow = prepareTableStyleObj(options.evenRow, existingTableStyle && existingTableStyle.evenRow);
            else if (existingTableStyle && existingTableStyle.evenRow) tableStyle.evenRow = existingTableStyle.evenRow;
            if (options.oddRow) tableStyle.oddRow = prepareTableStyleObj(options.oddRow, existingTableStyle && existingTableStyle.oddRow);
            else if (existingTableStyle && existingTableStyle.oddRow) tableStyle.oddRow = existingTableStyle.oddRow;
            if (options.firstColumn) tableStyle.firstColumn = prepareTableStyleObj(options.firstColumn, existingTableStyle && existingTableStyle.firstColumn);
            else if (existingTableStyle && existingTableStyle.firstColumn) tableStyle.firstColumn = existingTableStyle.firstColumn;
            if (options.lastColumn) tableStyle.lastColumn = prepareTableStyleObj(options.lastColumn, existingTableStyle && existingTableStyle.lastColumn);
            else if (existingTableStyle && existingTableStyle.lastColumn) tableStyle.lastColumn = existingTableStyle.lastColumn;
            if (options.firstRow) tableStyle.firstRow = prepareTableStyleObj(options.firstRow, existingTableStyle && existingTableStyle.firstRow);
            else if (existingTableStyle && existingTableStyle.firstRow) tableStyle.firstRow = existingTableStyle.firstRow;
            if (options.lastRow) tableStyle.lastRow = prepareTableStyleObj(options.lastRow, existingTableStyle && existingTableStyle.lastRow);
            else if (existingTableStyle && existingTableStyle.lastRow) tableStyle.lastRow = existingTableStyle.lastRow;
            if (options.firstRowFirstCell) tableStyle.firstRowFirstCell = prepareTableStyleObj(options.firstRowFirstCell, existingTableStyle && existingTableStyle.firstRowFirstCell);
            else if (existingTableStyle && existingTableStyle.firstRowFirstCell) tableStyle.firstRowFirstCell = existingTableStyle.firstRowFirstCell;
            if (options.firstRowLastCell) tableStyle.firstRowLastCell = prepareTableStyleObj(options.firstRowLastCell, existingTableStyle && existingTableStyle.firstRowLastCell);
            else if (existingTableStyle && existingTableStyle.firstRowLastCell) tableStyle.firstRowLastCell = existingTableStyle.firstRowLastCell;
            if (options.lastRowFirstCell) tableStyle.lastRowFirstCell = prepareTableStyleObj(options.lastRowFirstCell, existingTableStyle && existingTableStyle.lastRowFirstCell);
            else if (existingTableStyle && existingTableStyle.lastRowFirstCell) tableStyle.lastRowFirstCell = existingTableStyle.lastRowFirstCell;
            if (options.lastRowLastCell) tableStyle.lastRowLastCell = prepareTableStyleObj(options.lastRowLastCell, existingTableStyle && existingTableStyle.lastRowLastCell);
            else if (existingTableStyle && existingTableStyle.lastRowLastCell) tableStyle.lastRowLastCell = existingTableStyle.lastRowLastCell;
            if (!existingTableStyle)
                this._styles.tableStyles.push(tableStyle);
            else {
                for (var index = 0; index < this._styles.tableStyles.length; index++) {
                    if (_table.tableStyle.name === this._styles.tableStyles[index].name) {
                        this._styles.tableStyles[index] = tableStyle;
                        break;
                    }
                }
            }
            return tableStyle;
        };

        var prepareTableStyleObj = function (options, existingTableStyle) {
            var tableStyle = {};
            tableStyle.font = oxmlXlsxFont.createTableFont(options, (existingTableStyle ? existingTableStyle.font : null));
            tableStyle.fill = oxmlXlsxFill.createTableFill(options, (existingTableStyle ? existingTableStyle.fill : null));
            tableStyle.border = oxmlXlsxBorder.createTableBorder(options);
            return tableStyle;
        };

        return {
            createStyle: function (_workbook, _rel, _contentType) {
                return new Styles(_workbook, _rel, _contentType);
            }
        };
    });