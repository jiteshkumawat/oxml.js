define([], function () {
    'use strict';

    // Add Relation
    var addRelation = function (id, type, target, _rels) {
        _rels.relations.push({
            Id: id,
            Type: type,
            Target: target
        });
    };

    var generateContent = function (_rels) {
        // Create RELS
        var index, rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for (index = 0; index < _rels.relations.length; index++) {
            rels += '<Relationship Id="' + _rels.relations[index].Id + '" Type="' + _rels.relations[index].Type + '" Target="' + _rels.relations[index].Target + '"/>';
        }
        rels += '</Relationships>';
        return rels;
    };

    var attach = function (file, _rels) {
        var rels = generateContent(_rels);
        file.addFile(rels, _rels.fileName, _rels.folderName);
    };

    // Create Relation
    var createRelation = function (fileName, folderName) {
        var _rels = {
            relations: []
        };
        _rels.fileName = fileName;
        _rels.folderName = folderName;
        _rels.addRelation = function (id, type, target) {
            addRelation(id, type, target, _rels);
        };
        _rels.generateContent = function () {
            return generateContent(_rels);
        };
        _rels.attach = function (file) {
            return attach(file, _rels);
        }
        return _rels;
    };

    return { createRelation: createRelation };

});