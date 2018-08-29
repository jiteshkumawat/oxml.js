define(['contentFile', 'contentString', 'xmlContentString'], function (ContentFile, ContentString, XMLContentString) {
    'use strict';

    var ContentTypes = function () {
        this.className = "ContentTypes";
        this._contentType = {
            contentTypes: []
        };
        this.fileName = "[Content_Types].xml";
    };
    ContentTypes.prototype = Object.create(ContentFile.prototype);
    ContentTypes.prototype.generateContent = function () {
        var template = new XMLContentString({
            rootNode : 'Types',
            nameSpaces: ['http://schemas.openxmlformats.org/package/2006/content-types']
        });
        var index, contentTypes = '';
        for (index = 0; index < this._contentType.contentTypes.length; index++) {
            var contentTypeTemplate = new ContentString('<{0} ContentType="{1}"{2}{3}/>');
            var contentType = this._contentType.contentTypes[index];
            contentTypes += contentTypeTemplate.format(
                contentType.name,
                contentType.contentType,
                (contentType.Extension ? (' Extension="' + contentType.Extension + '"') : ''),
                (contentType.PartName ? (' PartName="' + contentType.PartName + '"') : ''));
        }
        return template.format(contentTypes);
    };

    ContentTypes.prototype.addContentType = function (name, contentType, attributes) {
        var content = {
            name: name,
            contentType: contentType
        };

        for (var key in attributes) {
            content[key] = attributes[key];
        }
        this._contentType.contentTypes.push(content);
    };

    return {
        createContentType: function () { return new ContentTypes(); }
    };
});