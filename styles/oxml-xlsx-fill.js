define(['utils'], function (utils) {
    var generateContent = function (_styles) {
        // Create Styles
        var stylesString = '';
        var fillKey;
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
        return stylesString;
    };

    var getFillString = function (fill) {
        var fillString = '<fill>';
        if (fill.pattern) {
            var patternType = fill.pattern || 'none';
            var colorString =
                fill.fgColor ? '<fgColor rgb="' + fill.fgColor + '"/>' : '';
            colorString +=
                fill.bgColor ? '<bgColor rgb="' + fill.bgColor + '"/>' : '';
            fillString += '<patternFill patternType="'
                + patternType + '">' + colorString + '</patternFill>';
        } else if (fill.gradient) {
            var typeString =
                fill.gradient.type ? ' type="' + fill.gradient.type + '" ' : '';
            var leftString =
                fill.gradient.left ? ' left="' + fill.gradient.left + '" ' : '';
            var rightString =
                fill.gradient.right ? ' right="'
                    + fill.gradient.right + '"' : '';
            var topString =
                fill.gradient.top ? ' top="' + fill.gradient.top + '" ' : '';
            var bottomString =
                fill.gradient.bottom ? ' bottom="'
                    + fill.gradient.bottom + '" ' : '';
            var degreeString =
                fill.gradient.degree ? ' degree="'
                    + fill.gradient.degree + '" ' : '';
            fillString += '<gradientFill '
                + typeString + leftString + rightString
                + topString + bottomString + degreeString + ' >';
            for (var stopIndex = 0;
                stopIndex < fill.gradient.stops.length;
                stopIndex++) {
                fillString += '<stop position="'
                    + fill.gradient.stops[stopIndex].position + '">';
                fillString += '<color rgb="'
                    + fill.gradient.stops[stopIndex].color + '" /></stop>';
            }
            fillString += '</gradientFill>';
        } else {
            return '<fill/>';
        }
        return fillString + '</fill>';
    };


    var createFill = function (options) {
        if (options.fill && options.fill.pattern) {
            var fill = {};
            fill.pattern = options.fill.pattern;
            fill.fgColor = options.fill.foreColor || false;
            fill.bgColor = options.fill.backColor
                || options.fill.color || false;
            return fill;
        } else if (options.fill
            && options.fill.gradient
            && options.fill.gradient.stop) {
            var fill = {};
            fill.gradient = {};
            fill.gradient.degree = options.fill.gradient.degree || false;
            fill.gradient.bottom = options.fill.gradient.bottom || false;
            fill.gradient.left = options.fill.gradient.left || false;
            fill.gradient.right = options.fill.gradient.right || false;
            fill.gradient.top = options.fill.gradient.top || false;
            fill.gradient.type = options.fill.gradient.type || false;
            fill.gradient.stops = [];
            for (var stopIndex = 0;
                stopIndex < options.fill.gradient.stop.length;
                stopIndex++) {
                var stop = {};
                stop.position =
                    options.fill.gradient.stop[stopIndex].position || 0;
                stop.color =
                    options.fill.gradient.stop[stopIndex].color || false;
                fill.gradient.stops.push(stop);
            }
            return fill;
        }
        return false;
    };

    var searchFill = function (fill, _styles) {
        return _styles._fills[utils.stringify(fill)];
    };

    var addFill = function (fill, _styles) {
        if (!_styles._fills) {
            _styles._fills = {};
            _styles._fillsCount = 0;
        }
        var index = _styles._fillsCount++;
        _styles._fills[utils.stringify(fill)] = "" + index;
        return _styles._fills[utils.stringify(fill)];
    };

    var searchSavedFillsForUpdate = function (_styles, cellIndices) {
        var index = 0, fillCount = 0, cellStyle;
        for (var index2 = 0; index2 < cellIndices.length; index2++) {
            var cellIndex = cellIndices[0];
            for (; index < _styles.styles.length; index++) {
                if (_styles.styles[index].cellIndices[cellIndex] !== undefined
                    || _styles.styles[index].cellIndices[cellIndex] !== null) {
                    cellStyle = _styles.styles[index];
                    fillCount++;
                    if (Object.keys(cellStyle.cellIndices).length
                        !== cellIndices.length
                        || fillCount > 1 || cellStyle._fill === 0) {
                        return false;
                    }
                }
            }
        }
        if (!cellStyle)
            return false;
        return cellStyle._fill;
    };

    var updateFill = function (fill, savedFill, _styles) {
        mergeFill(fill, savedFill, _styles, true);
        _styles._fills[utils.stringify(fill)] = savedFill;
        return savedFill;
    };

    var mergeFill = function (fill, savedFill, _styles, deleteExisting) {
        var savedFillDetails;
        for (var key in _styles._fills) {
            if (_styles._fills[key] === savedFill) {
                savedFillDetails = JSON.parse(key);
                break;
            }
        }
        if (deleteExisting)
            delete _styles._fills[utils.stringify(savedFillDetails)];
        for (var key in fill) {
            if (fill[key])
                savedFillDetails[key] = fill[key];
            fill[key] = savedFillDetails[key];
        }
        return fill;
    };

    var getFillForCells = function (_styles, options) {
        var newStyleCreated = false, fill = createFill(options), fillIndex;
        var savedFill = _styles._fills ? searchFill(fill, _styles) : null;
        if (savedFill !== undefined && savedFill !== null) {
            fillIndex = savedFill;
        } else {
            newStyleCreated = true;
        }
        if (!fillIndex) {
            savedFill = searchSavedFillsForUpdate(_styles, options.cellIndices);
            if (savedFill !== false) {
                fillIndex = updateFill(fill, savedFill, _styles);
            }
        }
        if (fillIndex === undefined || fillIndex === null)
            fillIndex = addFill(fill, _styles);
        return {
            fill: fill,
            fillIndex: fillIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getFillForCell = function (_styles, options, cellStyle) {
        // Create Fill Object
        var newStyleCreated = false, fill = createFill(options), fillIndex;

        var savedFill = _styles._fills ? searchFill(fill, _styles) : null;

        if (savedFill !== undefined && savedFill !== null) {
            fillIndex = savedFill;
        } else {
            // Check if fill can be updated
            newStyleCreated = true;
        }
        if (cellStyle && cellStyle._fill) {
            var fillCount = getFillCounts(cellStyle._fill, _styles);
            if (fillCount <= 1) {
                // Update fill
                fillIndex = updateFill(fill, cellStyle._fill, _styles);
            }
        }
        if (fillIndex === undefined || fillIndex === null) {
            if (cellStyle && cellStyle._fill) {
                fill = mergeFill(fill, cellStyle._fill, _styles, false);
            }
            fillIndex = addFill(fill, _styles);
        }
        if (!fillIndex && cellStyle && cellStyle._fill)
            fillIndex = cellStyle._fill;

        return {
            fill: fill,
            fillIndex: fillIndex,
            newStyleCreated: newStyleCreated
        };
    };

    var getFillCounts = function (fill, _styles) {
        var count = 0, index;
        for (index = 0; index < _styles.styles.length; index++) {
            if (_styles.styles[index]._fill === fill) {
                count += Object.keys(_styles.styles[index].cellIndices).length;
                if (count > 1)
                    return count;
            }
        }
        return count;
    };

    return {
        generateContent: generateContent,
        getFillForCell: getFillForCell,
        getFillForCells: getFillForCells
    };
});