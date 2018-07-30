define([], function () {
    'use strict';

    var generateContent = function (_contentType) {
        var index, contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        for (index = 0; index < _contentType.contentTypes.length; index++) {
            var contentType = _contentType.contentTypes[index];
            contentTypes += '<' + contentType.name + ' ContentType="' + contentType.contentType + '" ';
            if (contentType.Extension) {
                contentTypes += 'Extension="' + contentType.Extension + '" ';
            }
            if (contentType.PartName) {
                contentTypes += 'PartName="' + contentType.PartName + '" ';
            }
            contentTypes += '/>\n';
        }
        contentTypes += '</Types>';
        return contentTypes;
    };

    var attach = function (file, _contentType) {
        var contentTypes = generateContent(_contentType);
        file.addFile(contentTypes, "[Content_Types].xml");
    };

    var addContentType = function (name, contentType, attributes, _contentType) {
        var content = {
            name: name,
            contentType: contentType
        };

        for (var key in attributes) {
            content[key] = attributes[key];
        }
        _contentType.contentTypes.push(content);
    };

    var createContentType = function () {
        var _contentType = {
            contentTypes: []
        };
        _contentType.addContentType = function (name, contentType, attributes) {
            addContentType(name, contentType, attributes, _contentType);
        };
        _contentType.attach = function (file) {
            return attach(file, _contentType);
        };
        return _contentType;
    };

    return {
        createContentType: createContentType
    };
});