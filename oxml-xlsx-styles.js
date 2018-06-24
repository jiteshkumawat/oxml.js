define([], function () {
    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        var index = 0, fontKey, borderKey, fillKey;
        if (_styles._numFormats) {
            stylesString += '<numFmts count="' + Object.keys(_styles._numFormats).length + '">';
            var numFormatKey;
            for (numFormatKey in _styles._numFormats) {
                var numFormat = JSON.parse(numFormatKey);
                stylesString += '<numFmt numFmtId="' + _styles._numFormats[numFormatKey] + '" formatCode="' + numFormat.formatString + '"/>';
            }
            stylesString += '</numFmts>';
        }
        stylesString += '<fonts count="' + _styles._fontsCount + '">';
        for (fontKey in _styles._fonts) {
            // Excel requires the child elements to be in the following sequence: i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, name, family, charset, scheme.
            var font = JSON.parse(fontKey);
            stylesString += '<font>';
            stylesString += font.strike ? '<strike/>' : '';
            stylesString += font.italic ? '<i/>' : '';
            stylesString += font.bold ? '<b/>' : '';
            stylesString += font.underline ? '<u/>' : '';
            stylesString += font.size ? '<sz val="' + font.size + '"/>' : '';
            stylesString += font.color ? '<color rgb="' + font.color + '"/>' : '';
            stylesString += font.name ? '<name val="' + font.name + '"/>' : '';
            stylesString += font.family ? '<family val="' + font.family + '"/>' : '';
            stylesString += font.scheme ? '<scheme val="' + font.scheme + '"/>' : '';
            stylesString += '</font>';
        }
        stylesString += '</fonts>';
        stylesString += '<fills count="' + _styles._fillsCount + '">';
        for (fillKey in _styles._fills) {
            var fill = JSON.parse(fillKey);
            if (fill) {
                stylesString += getFillString(fill);
            } else {
                stylesString += '<fill/>';
            }
        }
        stylesString += '</fills>';
        stylesString += '<borders count="' + _styles._bordersCount + '">';
        for (borderKey in _styles._borders) {
            var border = JSON.parse(borderKey);
            if (border) {
                stylesString += '<border>';
                stylesString += getBorderString(border.left, 'left');
                stylesString += getBorderString(border.right, 'right');
                stylesString += getBorderString(border.top, 'top');
                stylesString += getBorderString(border.bottom, 'bottom');
                stylesString += getBorderString(border.diagonal, 'diagonal');
                stylesString += '</border>'
            } else {
                stylesString += '<border />';
            }
        }
        stylesString += '</borders>' + '<cellStyleXfs count="1"><xf /></cellStyleXfs>';

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

    var getBorderString = function (border, borderType) {
        if (border.style || border.color) {
            var borderStyleString = border.style ? ' style="' + border.style + '" ' : '';
            var borderColorString = border.color ? '<color rgb="' + border.color + '"/>' : '';
            return '<' + borderType + ' ' + borderStyleString + '>' + borderColorString + '</' + borderType + '>';
        }
        return '<' + borderType + '/>';
    };

    var getFillString = function (fill) {
        var fillString = '<fill>';
        if (fill.pattern) {
            var patternType = fill.pattern || 'none';
            var colorString = fill.fgColor ? '<fgColor rgb="' + fill.fgColor + '"/>' : '';
            colorString += fill.bgColor ? '<bgColor rgb="' + fill.bgColor + '"/>' : '';
            fillString += '<patternFill patternType="' + patternType + '">' + colorString + '</patternFill>';
        }
        else if (fill.gradient) {
            var typeString = fill.gradient.type ? ' type="' + fill.gradient.type + '" ' : '';
            var leftString = fill.gradient.left ? ' left="' + fill.gradient.left + '" ' : '';
            var rightString = fill.gradient.right ? ' right="' + fill.gradient.right + '"' : '';
            var topString = fill.gradient.top ? ' top="' + fill.gradient.top + '" ' : '';
            var bottomString = fill.gradient.bottom ? ' bottom="' + fill.gradient.bottom + '" ' : '';
            var degreeString = fill.gradient.degree ? ' degree="' + fill.gradient.degree + '" ' : '';
            fillString += '<gradientFill ' + typeString + leftString + rightString + topString + bottomString + degreeString + ' >';
            for (var stopIndex = 0; stopIndex < fill.gradient.stops.length; stopIndex++) {
                fillString += '<stop position="' + fill.gradient.stops[stopIndex].position + '">';
                fillString += '<color rgb="' + fill.gradient.stops[stopIndex].color + '" /></stop>';
            }
            fillString += '</gradientFill>';
        } else {
            return '<fill/>';
        }
        return fillString + '</fill>';
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

    var createFont = function (options) {
        var font = {};
        font.bold = !!options.bold;
        font.italic = !!options.italic;
        font.underline = !!options.underline;
        font.size = options.fontSize || false;
        font.color = options.fontColor || false;
        font.name = options.fontName || false;
        font.family = options.fontFamily || false;
        font.scheme = options.scheme || false;
        font.strike = !!options.strike;
        return font;
    }

    var searchFont = function (font, _styles) {
        return _styles._fonts[JSON.stringify(font)];
    };

    var addFont = function (font, _styles) {
        if (!_styles._fonts) {
            _styles._fonts = {};
            _styles._fontsCount = 0;
        }
        var index = _styles._fontsCount++;
        _styles._fonts[JSON.stringify(font)] = "" + index;
        return _styles._fonts[JSON.stringify(font)];
    };

    var createFill = function (options) {
        if (options.fill && options.fill.pattern) {
            var fill = {};
            fill.pattern = options.fill.pattern;
            fill.fgColor = options.fill.foreColor || false;
            fill.bgColor = options.fill.backColor || options.fill.color || false;
            return fill;
        } else if (options.fill && options.fill.gradient && options.fill.gradient.stop) {
            var fill = {};
            fill.gradient = {};
            fill.gradient.degree = options.fill.gradient.degree || false;
            fill.gradient.bottom = options.fill.gradient.bottom || false;
            fill.gradient.left = options.fill.gradient.left || false;
            fill.gradient.right = options.fill.gradient.right || false;
            fill.gradient.top = options.fill.gradient.top || false;
            fill.gradient.type = options.fill.gradient.type || false;
            fill.gradient.stops = [];
            for (var stopIndex = 0; stopIndex < options.fill.gradient.stop.length; stopIndex++) {
                var stop = {};
                stop.position = options.fill.gradient.stop[stopIndex].position || 0;
                stop.color = options.fill.gradient.stop[stopIndex].color || false;
                fill.gradient.stops.push(stop);
            }
            return fill;
        }
        return false;
    };

    var searchFill = function (fill, _styles) {
        return _styles._fills[JSON.stringify(fill)];
    };

    var addFill = function (fill, _styles) {
        if (!_styles._fills) {
            _styles._fills = {};
            _styles._fillsCount = 0;
        }
        var index = _styles._fillsCount++;
        _styles._fills[JSON.stringify(fill)] = "" + index;
        return _styles._fills[JSON.stringify(fill)];
    };

    var createNumFormat = function (options) {
        var numberFormat = {};
        numberFormat.formatString = options.numberFormat;
        return numberFormat;
    };

    var searchNumFormat = function (numFormat, _styles) {
        return _styles._numFormats[JSON.stringify(numFormat)];
    };

    var addNumFormat = function (numFormat, _styles) {
        if (!_styles._numFormats) {
            _styles._numFormats = {};
            _styles._numFormatsCount = 200;
        }
        var index = _styles._numFormatsCount++;
        _styles._numFormats[JSON.stringify(numFormat)] = "" + index;
        return _styles._numFormats[JSON.stringify(numFormat)];
    };

    var createBorder = function (options) {
        var border = {};
        if (options.border) {
            border.left = {
                color: (options.border.left ? options.border.left.color : false) || options.border.color || false,
                style: (options.border.left ? options.border.left.style : false) || options.border.style || false
            };
            border.top = {
                color: (options.border.top ? options.border.top.color : false) || options.border.color || false,
                style: (options.border.top ? options.border.top.style : false) || options.border.style || false
            };
            border.bottom = {
                color: (options.border.bottom ? options.border.bottom.color : false) || options.border.color || false,
                style: (options.border.bottom ? options.border.bottom.style : false) || options.border.style || false
            };
            border.right = {
                color: (options.border.right ? options.border.right.color : false) || options.border.color || false,
                style: (options.border.right ? options.border.right.style : false) || options.border.style || false
            };
            border.diagonal = {
                color: (options.border.diagonal ? options.border.diagonal.color : false) || options.border.color || false,
                style: (options.border.diagonal ? options.border.diagonal.style : false) || options.border.style || false
            };
            border.isValid = border.left.style || border.top.style || border.bottom.style || border.right.style || border.diagonal.style
                || border.left.color || border.top.color || border.bottom.color || border.right.color || border.diagonal.color;
            return border.isValid ? border : false;
        }
        return false;
    };

    var searchBorder = function (border, _styles) {
        return _styles._borders[JSON.stringify(border)];
    };

    var addBorder = function (border, _styles) {
        if (!_styles._borders) {
            _styles._borders = {};
            _styles._bordersCount = 0;
        }
        var index = _styles._bordersCount++;
        _styles._borders[JSON.stringify(border)] = "" + index;
        return _styles._borders[JSON.stringify(border)];
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
            var newStyleCreated = false;

            var font = createFont(options);
            var numFormat = options.numberFormat ? createNumFormat(options) : null;
            var border = createBorder(options);
            var fill = createFill(options);

            var savedFont = _styles._fonts ? searchFont(font, _styles) : null;
            var savedNumFormat = options.numberFormat && _styles._numFormats ? searchNumFormat(numFormat, _styles) : null;
            var savedBorder = border && _styles._borders ? searchBorder(border, _styles) : null;
            var savedFill = fill && _styles._fills ? searchFill(fill, _styles) : null;

            if (savedFont) {
                font = savedFont;
            } else {
                newStyleCreated = true;
                font = addFont(font, _styles);
            }
            if (savedNumFormat) {
                numFormat = savedNumFormat;
            } else if (options.numberFormat) {
                newStyleCreated = true;
                numFormat = addNumFormat(numFormat, _styles);
            }
            if (savedBorder) {
                border = savedBorder;
            } else if (border) {
                newStyleCreated = true;
                border = addBorder(border, _styles);
            }
            if (savedFill) {
                fill = savedFill;
            } else if (fill) {
                newStyleCreated = true;
                fill = addFill(fill, _styles);
            }

            if (options.cellIndex) {
                var cellStyle = searchStyleForCell(_styles, options.cellIndex);
                if (cellStyle) {
                    if (cellStyle._font === font && cellStyle._numFormat === numFormat && cellStyle._border === border && cellStyle._fill === fill) {
                        return cellStyle;
                    }
                    var totalCellApplied = Object.keys(cellStyle.cellIndices).length;
                    if (totalCellApplied === 1) {
                        cellStyle._font = font;
                        cellStyle._numFormat = numFormat || false;
                        cellStyle._border = border;
                        cellStyle._fill = fill;
                        return cellStyle;
                    }
                    delete cellStyle.cellIndices[options.cellIndex];
                }

                if (!newStyleCreated) {
                    cellStyle = searchSimilarStyle(_styles, {
                        _font: font,
                        _numFormat: numFormat || false,
                        _border: border,
                        _fill: fill
                    });
                    if (cellStyle) {
                        cellStyle.cellIndices[options.cellIndex] = Object.keys(cellStyle.cellIndices).length;
                        return cellStyle;
                    }
                }

                cellStyle = {
                    _font: font,
                    _numFormat: numFormat || false,
                    _border: border,
                    _fill: fill,
                    cellIndices: {}
                };

                cellStyle.cellIndices[options.cellIndex] = 0;
                cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);
                _styles.styles.push(cellStyle);
                return cellStyle;
            }

            if (options.cellIndices) {
                var cellStyle, index;
                if (!newStyleCreated) {
                    cellStyle = searchSimilarStyle(_styles, {
                        _font: font,
                        _numFormat: numFormat || false,
                        _border: border,
                        _fill: fill
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
                cellStyle = {
                    _font: font,
                    _numFormat: numFormat || false,
                    _border: border,
                    _fill: fill,
                    cellIndices: {}
                };
                cellStyle.index = "" + (parseInt(_styles.styles.length, 10) + 1);

                for (index = 0; index < options.cellIndices.length; index++) {
                    var cellIndex = options.cellIndices[index];
                    var savedCellStyle = searchStyleForCell(_styles, cellIndex);
                    if (savedCellStyle) {
                        delete savedCellStyle.cellIndices[cellIndex];
                    }
                    cellStyle.cellIndices[cellIndex] = Object.keys(cellStyle.cellIndices).length;
                }
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
            _styles._fonts[JSON.stringify({
                bold: false,
                italic: false,
                underline: false,
                size: false,
                color: false
            })] = 0;
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