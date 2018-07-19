define(['utils',
    'oxml_xlsx_font',
    'oxml_xlsx_num_format',
    'oxml_xlsx_border',
    'oxml_xlsx_fill'],
    function (utils,
        oxmlXlsxFont,
        oxmlXlsxNumFormat,
        oxmlXlsxBorder,
        oxmlXlsxFill) {
        var generateContent = function (_styles) {
            // Create Styles
            var stylesString = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
            var index = 0;
            stylesString += oxmlXlsxNumFormat.generateContent(_styles);
            stylesString += oxmlXlsxFont.generateContent(_styles);
            stylesString += oxmlXlsxFill.generateContent(_styles);
            stylesString += oxmlXlsxBorder.generateContent(_styles);
            stylesString += '<cellStyleXfs count="1"><xf /></cellStyleXfs>';

            stylesString += '<cellXfs count="'
                + (parseInt(_styles.styles.length, 10) + 1) + '"><xf />';
            for (index = 0; index < _styles.styles.length; index++) {
                var numFormatString = _styles.styles[index]._numFormat
                    ? ' numFmtId="'
                    + _styles.styles[index]._numFormat + '" ' : '';
                var borderString = _styles.styles[index]._border
                    ? ' borderId="' + _styles.styles[index]._border + '" ' : '';
                var fillString = _styles.styles[index]._fill
                    ? ' fillId="' + _styles.styles[index]._fill + '" ' : '';
                stylesString += '<xf fontId="' + _styles.styles[index]._font
                    + '" ' + numFormatString
                    + borderString + fillString + ' />';
            }
            var tableStyles = '';
            if (_styles.tableStyles && _styles.tableStyles.length) tableStyles = generateTableContent(_styles);
            stylesString += '</cellXfs>' + tableStyles + '</styleSheet>';
            return stylesString;
        };

        var generateTableContent = function (_styles) {
            var dxfsCount = 0, tableStyles = '', tableStylesCount = 0, tableStyleCount;
            for (var index = 0; index < _styles.tableStyles.length; index++) {
                var dxfString = '', individualTalbeStyle;
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
            dxfString = '<dxf>';
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

        var attach = function (file, _styles) {
            var styles = generateContent(_styles);
            file.addFile(styles, _styles.fileName, "workbook");
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

        var addStyles = function (options, _styles) {
            if (options.cellIndex || options.cellIndices) {
                var newStyleCreated = false, cellStyle,
                    saveFont, saveNumFormat, saveBorder, saveFill;
                if (options.cellIndex) {
                    cellStyle = searchStyleForCell(_styles, options.cellIndex);
                    saveFont = oxmlXlsxFont
                        .getFontForCell(_styles, options, cellStyle);
                    saveNumFormat = oxmlXlsxNumFormat
                        .getNumFormatForCell(_styles, options, cellStyle);
                    saveBorder = oxmlXlsxBorder
                        .getBorderForCell(_styles, options, cellStyle);
                    saveFill = oxmlXlsxFill
                        .getFillForCell(_styles, options, cellStyle);
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
                        cellStyle = searchSimilarStyle(_styles, {
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
                    cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);
                    _styles.styles.push(cellStyle);
                    return cellStyle;
                }

                if (options.cellIndices) {
                    var index;
                    saveFont = oxmlXlsxFont.getFontForCells(_styles, options);
                    saveNumFormat = oxmlXlsxNumFormat.getNumFormatForCells(_styles, options);
                    saveBorder = oxmlXlsxBorder.getBorderForCells(_styles, options);
                    saveFill = oxmlXlsxFill.getFillForCells(_styles, options);
                    newStyleCreated = newStyleCreated || saveFont.newStyleCreated
                        || saveNumFormat.newStyleCreated || saveBorder.newStyleCreated || saveFill.newStyleCreated;
                    if (!newStyleCreated) {
                        cellStyle = searchSimilarStyle(_styles, {
                            _font: saveFont.fontIndex,
                            _numFormat: saveNumFormat.numFormatIndex,
                            _border: saveBorder.borderIndex,
                            _fill: saveFill.fillIndex
                        });

                        if (cellStyle) {
                            for (index = 0; index < options.cellIndices.length; index++) {
                                var cellIndex = options.cellIndices[index];
                                var savedCellStyle = searchStyleForCell(_styles, cellIndex);
                                delete savedCellStyle.cellIndices[cellIndex];
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
                        var savedCellStyle = searchStyleForCell(_styles, cellIndex);
                        if (savedCellStyle) {
                            delete savedCellStyle.cellIndices[cellIndex];
                        }
                        cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                        if (savedCellStyle && !Object.keys(savedCellStyle.cellIndices).length) {
                            for (key in cellStyle) {
                                savedCellStyle[key] = cellStyle[key];
                            }
                            return savedCellStyle;
                        }
                    }
                    cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);
                    _styles.styles.push(cellStyle);
                    return cellStyle;
                }
            }
        };

        var addTableStyle = function (options, tableStyleName, _styles) {
            if (!_styles.tableStyles) _styles.tableStyles = [];
            var tableStyle = prepareTableStyleObj(options);
            tableStyle.name = tableStyleName;
            if (options.evenColumn) tableStyle.evenColumn = prepareTableStyleObj(options.evenColumn);
            if (options.oddColumn) tableStyle.oddColumn = prepareTableStyleObj(options.oddColumn);
            if (options.evenRow) tableStyle.evenRow = prepareTableStyleObj(options.evenRow);
            if (options.oddRow) tableStyle.oddRow = prepareTableStyleObj(options.oddRow);
            if (options.firstColumn) tableStyle.firstColumn = prepareTableStyleObj(options.firstColumn);
            if (options.lastColumn) tableStyle.lastColumn = prepareTableStyleObj(options.lastColumn);
            if (options.firstRow) tableStyle.firstRow = prepareTableStyleObj(options.firstRow);
            if (options.lastRow) tableStyle.lastRow = prepareTableStyleObj(options.lastRow);
            if (options.firstRowFirstCell) tableStyle.firstRowFirstCell = prepareTableStyleObj(options.firstRowFirstCell);
            if (options.firstRowLastCell) tableStyle.firstRowLastCell = prepareTableStyleObj(options.firstRowLastCell);
            if (options.lastRowFirstCell) tableStyle.lastRowFirstCell = prepareTableStyleObj(options.lastRowFirstCell);
            if (options.lastRowLastCell) tableStyle.lastRowLastCell = prepareTableStyleObj(options.lastRowLastCell);
            _styles.tableStyles.push(tableStyle);
            return tableStyle;
        };

        var prepareTableStyleObj = function (options) {
            var tableStyle = {};
            tableStyle.font = oxmlXlsxFont.createFont(options);
            tableStyle.fill = oxmlXlsxFill.createFill(options);
            tableStyle.border = oxmlXlsxBorder.createBorder(options);
            return tableStyle;
        };

        return {
            createStyle: function (_workbook, _rel, _contentType) {
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
                return {
                    _styles: _styles,
                    generateContent: function () {
                        return generateContent(_styles);
                    },
                    attach: function (file) {
                        attach(file, _styles);
                    },
                    addStyles: function (options) {
                        return addStyles(options, _styles);
                    },
                    addTableStyle: function (options, tableStyleName) {
                        return addTableStyle(options, tableStyleName, _styles);
                    }
                };
            }
        };
    });