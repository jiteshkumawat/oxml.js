define([], function () {
    var generateContent = function (_styles) {
        var stylesString = '', fontKey;
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
        return stylesString;
    }

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
    };

    var searchFont = function (font, _styles) {
        return _styles._fonts[JSON.stringify(font, Object.keys(font).sort())];
    };

    var searchSavedFontsForUpdate = function (_styles, cellIndices) {
        var index = 0, fontCount = 0;
        var cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    fontCount++;
                    if (Object.keys(cellStyle.cellIndices).length !== cellIndices.length || fontCount > 1 || cellStyle._font === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._font;
    };

    var addFont = function (font, _styles) {
        if (!_styles._fonts) {
            _styles._fonts = {};
            _styles._fontsCount = 0;
        }
        var index = _styles._fontsCount++;
        _styles._fonts[JSON.stringify(font, Object.keys(font).sort())] = "" + index;
        return _styles._fonts[JSON.stringify(font, Object.keys(font).sort())];
    };

    var updateFont = function (font, savedFont, _styles) {
        var savedFontDetails;
        for (var key in _styles._fonts) {
            if (_styles._fonts[key] === savedFont) {
                savedFontDetails = JSON.parse(key);
                break;
            }
        }
        delete _styles._fonts[JSON.stringify(savedFontDetails, Object.keys(savedFontDetails).sort())];
        for (var key in font) {
            if (font[key])
                savedFontDetails[key] = font[key];
            font[key] = savedFontDetails[key];
        }
        _styles._fonts[JSON.stringify(savedFontDetails, Object.keys(savedFontDetails).sort())] = savedFont;
        return savedFont;
    };

    var getFontCounts = function (font, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._font === font) {
                count++;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    var getFontForCells = function (_styles, options) {
        var newStyleCreated = false, font = createFont(options), fontIndex;
        var savedFont = _styles._fonts ? searchFont(font, _styles) : null;
        if (savedFont !== undefined && savedFont !== null) {
            fontIndex = savedFont;
        } else {
            newStyleCreated = true;
        }
        if (!fontIndex) {
            savedFont = searchSavedFontsForUpdate(_styles, options.cellIndices);
            if (savedFont !== false) {
                fontIndex = updateFont(font, savedFont, _styles);
            }
        }
        if (fontIndex === undefined || fontIndex === null)
            fontIndex = addFont(font, _styles);
        return {
            font: font,
            fontIndex: fontIndex,
            newStyleCreated: newStyleCreated
        };
    }

    var getFontForCell = function (_styles, options, cellStyle) {
        // Create Font Object
        var newStyleCreated = false, font = createFont(options), fontIndex;

        var savedFont = _styles._fonts ? searchFont(font, _styles) : null;

        if (savedFont !== undefined && savedFont !== null) {
            fontIndex = savedFont;
        } else {
            // Check if font can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._font) {
            var fontCount = getFontCounts(cellStyle._font, _styles);
            if (fontCount <= 1) {
                // Update font
                fontIndex = updateFont(font, cellStyle._font, _styles);
            }
        }
        if (fontIndex === undefined || fontIndex === null)
            fontIndex = addFont(font, _styles);

        return {
            font: font,
            fontIndex: fontIndex,
            newStyleCreated: newStyleCreated
        };
    };

    return {
        getFontForCell: getFontForCell,
        getFontForCells: getFontForCells,
        generateContent: generateContent
    };
});