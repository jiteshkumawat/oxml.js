define(['contentFile', 'contentString', 'xmlContentString'], function (ContentFile, ContentString, XMLContentString) {
    'use strict';

    // Add Relation
    var Relation = function (fileName, folderName) {
        this.className = "Relation";
        this.fileName = fileName;
        this.folderName = folderName;
        this.relations = [];
    };
    Relation.prototype = Object.create(ContentFile.prototype);

    Relation.prototype.addRelation = function (id, type, target) {
        this.relations.push({
            Id: id,
            Type: type,
            Target: target
        });
    };

    Relation.prototype.generateContent = function () {
        // Create RELS
        var template = new XMLContentString({
            rootNode: 'Relationships',
            nameSpaces: ['http://schemas.openxmlformats.org/package/2006/relationships']
        });
        var index, rels = '';
        for (index = 0; index < this.relations.length; index++) {
            var relationTemplate = new ContentString('<Relationship Id="{0}" Type="{1}" Target="{2}"/>');
            rels += relationTemplate.format(this.relations[index].Id, this.relations[index].Type, this.relations[index].Target);
        }
        return template.format(rels);
    };

    return {
        createRelation: function (fileName, folderName) {
            return new Relation(fileName, folderName);
        }
    };
});