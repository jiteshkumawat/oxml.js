define(['utils'], function (utils) {
    var getBorderString = function (border, borderType) {
        if (border.style || border.color) {
            var borderStyleString = border.style ? ' style="' + border.style + '" ' : '';
            var borderColorString = border.color ? '<color rgb="' + border.color + '"/>' : '';
            return '<' + borderType + ' ' + borderStyleString + '>' + borderColorString + '</' + borderType + '>';
        }
        return '<' + borderType + '/>';
    };

    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '';
        var borderKey;
        stylesString += '<borders count="' + _styles._bordersCount + '">';
        for (borderKey in _styles._borders) {
            var border = JSON.parse(borderKey);
            if (border) {
                stylesString += generateSingleContent(border);
            } else {
                stylesString += '<border />';
            }
        }
        stylesString += '</borders>';
        return stylesString;
    };

    var generateSingleContent = function (border) {
        var stylesString = '<border>';
        stylesString += getBorderString(border.left, 'left');
        stylesString += getBorderString(border.right, 'right');
        stylesString += getBorderString(border.top, 'top');
        stylesString += getBorderString(border.bottom, 'bottom');
        stylesString += getBorderString(border.diagonal, 'diagonal');
        stylesString += '</border>';
        return stylesString;
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
        return _styles._borders[utils.stringify(border)];
    };

    var addBorder = function (border, _styles) {
        if (!_styles._borders) {
            _styles._borders = {};
            _styles._bordersCount = 0;
        }
        var index = _styles._bordersCount++;
        _styles._borders[utils.stringify(border)] = "" + index;
        return _styles._borders[utils.stringify(border)];
    };

    var searchSavedBordersForUpdate = function (_styles, cellIndices) {
        var index = 0, borderCount = 0;
        var cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    borderCount++;
                    if (Object.keys(cellStyle.cellIndices).length !== cellIndices.length || borderCount > 1 || cellStyle._border === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._border;
    };

    var updateBorder = function (border, savedBorder, _styles) {
        mergeBorder(border, savedBorder, _styles, true);
        _styles._borders[utils.stringify(border)] = savedBorder;
        return savedBorder;
    };

    var mergeBorder = function (border, savedBorder, _styles, deleteSavedBorder) {
        var savedBorderDetails;
        for (var key in _styles._borders) {
            if (_styles._borders[key] === savedBorder) {
                savedBorderDetails = JSON.parse(key);
                break;
            }
        }
        if (deleteSavedBorder) {
            delete _styles._borders[utils.stringify(savedBorderDetails)];
        }
        if (border) {
            for (var key in border) {
                if (border[key].color || border[key].style)
                    savedBorderDetails[key] = border[key];
                border[key] = savedBorderDetails[key];
            }
        } else
            border = savedBorderDetails;
        return border;
    };

    var getBorderCounts = function (border, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._border === border) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    var getBorderForCell = function (_styles, options, cellStyle) {
        // Create Border Object
        var newStyleCreated = false, border = createBorder(options), borderIndex;

        var savedBorder = _styles._borders ? searchBorder(border, _styles) : null;

        if (savedBorder !== undefined && savedBorder !== null) {
            borderIndex = savedBorder;
        } else {
            // Check if border can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._border && cellStyle._border !== savedBorder) {
            var borderCount = getBorderCounts(cellStyle._border, _styles);
            if (borderCount <= 1) {
                // Update border
                borderIndex = updateBorder(border, cellStyle._border, _styles);
            }
        }
        if (borderIndex === undefined || borderIndex === null) {
            if (cellStyle && cellStyle._border) {
                border = mergeBorder(border, cellStyle._border, _styles, false);
            }
            borderIndex = addBorder(border, _styles);
        }

        return {
            border: border,
            borderIndex: borderIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getBorderForCells = function (_styles, options) {
        var newStyleCreated = false, border = createBorder(options), borderIndex;
        var savedBorder = _styles._borders ? searchBorder(border, _styles) : null;
        if (savedBorder !== undefined && savedBorder !== null) {
            borderIndex = savedBorder;
        } else {
            newStyleCreated = true;
        }
        savedBorder = searchSavedBordersForUpdate(_styles, options.cellIndices);
        if (savedBorder !== false) {
            borderIndex = updateBorder(border, savedBorder, _styles);
        }
        if (borderIndex === undefined || borderIndex === null)
            borderIndex = addBorder(border, _styles);
        return {
            border: border,
            borderIndex: borderIndex,
            newStyleCreated: newStyleCreated
        };
    };

    return {
        createBorder: createBorder,
        getBorderForCell: getBorderForCell,
        getBorderForCells: getBorderForCells,
        generateContent: generateContent,
        generateSingleContent: generateSingleContent
    };
});