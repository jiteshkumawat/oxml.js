define(['utils', 'oxml_xlsx_font', 'oxml_xlsx_num_format', 'oxml_xlsx_border', 'oxml_xlsx_fill'], function (utils, oxmlXlsxFont, oxmlXlsxNumFormat, oxmlXlsxBorder, oxmlXlsxFill) {
    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        var index = 0, fillKey;
        stylesString += oxmlXlsxNumFormat.generateContent(_styles);
        stylesString += oxmlXlsxFont.generateContent(_styles);
        stylesString += oxmlXlsxFill.generateContent(_styles);
        stylesString += oxmlXlsxBorder.generateContent(_styles);
        stylesString += '<cellStyleXfs count="1"><xf /></cellStyleXfs>';

        stylesString += '<cellXfs count="' + (parseInt(_styles.styles.length, 10) + 1) + '"><xf />';
        for (index = 0; index < _styles.styles.length; index++) {
            var numFormatString = _styles.styles[index]._numFormat ? ' numFmtId="' + _styles.styles[index]._numFormat + '" ' : '';
            var borderString = _styles.styles[index]._border ? ' borderId="' + _styles.styles[index]._border + '" ' : '';
            var fillString = _styles.styles[index]._fill ? ' fillId="' + _styles.styles[index]._fill + '" ' : '';
            stylesString += '<xf fontId="' + _styles.styles[index]._font + '" ' + numFormatString + borderString + fillString + ' />';
        }
        stylesString += '</cellXfs></styleSheet>';
        return stylesString;
    };

    var attach = function (file, _styles) {
        // Add REL
        var styles = generateContent(_styles);
        file.addFile(styles, _styles.fileName, "workbook");
    };

    var createStyleParts = function (_workBook, _rel, _contentType) {
        var lastSheetRel = _workBook.getLastSheet();
        var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
        _contentType.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", {
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
            if (_styles.styles[index].cellIndices[cellIndex] !== undefined && _styles.styles[index].cellIndices[cellIndex] !== null) {
                return _styles.styles[index];
            }
        }
        return null;
    };

    var searchSimilarStyle = function (_styles, style) {
        var index = 0;
        for (; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._font === style._font && _styles.styles[index]._numFormat === style._numFormat) {
                return _styles.styles[index];
            }
        }
        return null;
    };

    var addStyles = function (options, _styles) {
        if (options.cellIndex || options.cellIndices) {
            var newStyleCreated = false, cellStyle, saveFont, saveNumFormat, saveBorder, saveFill;
            if (options.cellIndex) {
                cellStyle = searchStyleForCell(_styles, options.cellIndex);
                saveFont = oxmlXlsxFont.getFontForCell(_styles, options, cellStyle);
                saveNumFormat = oxmlXlsxNumFormat.getNumFormatForCell(_styles, options, cellStyle);
                saveBorder = oxmlXlsxBorder.getBorderForCell(_styles, options, cellStyle);
                saveFill = oxmlXlsxFill.getFillForCell(_styles, options, cellStyle);
                newStyleCreated = newStyleCreated || saveFont.newStyleCreated || saveBorder.newStyleCreated || saveFill.newStyleCreated;
                if (cellStyle) {
                    if (cellStyle._font === saveFont.fontIndex && cellStyle._numFormat === saveNumFormat.numFormatIndex && cellStyle._border === saveBorder.borderIndex && cellStyle._fill === saveFill.fillIndex) {
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
                newStyleCreated = newStyleCreated || saveFont.newStyleCreated || saveNumFormat.newStyleCreated || saveBorder.newStyleCreated || saveFill.newStyleCreated;
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
                    if(savedCellStyle && !Object.keys(savedCellStyle.cellIndices).length){
                        for(key in cellStyle){
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
                _fillsCount: 1
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
                }
            };
        }
    };
});