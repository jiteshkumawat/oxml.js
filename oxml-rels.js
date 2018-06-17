(function (window) {
    // Add Relation
    var addRelation = function(id, type, target, _rels){
        _rels.relations.push({
            Id: id,
            Type: type,
            Target: target
        });
    };

    var generateContent = function(_rels){
        // Create RELS
        var index, rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        `;
        for (index = 0; index < _rels.relations.length; index++) {
            rels += `<Relationship Id="${_rels.relations[index].Id}" Type="${_rels.relations[index].Type}" Target="${_rels.relations[index].Target}"/>
            `;
        }
        rels += `</Relationships>`;
        return rels;
    };

    // Create Relation
    var createRelation = function(fileName){
        var _rels = {
            relations: []
        };
        _rels.fileName = fileName;
        _rels.addRelation = function(id, type, target){
            addRelation(id, type, target, _rels);
        };
        _rels.generateContent = function(){
            return generateContent(_rels);
        };
        return _rels;
    };

    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createRelation = createRelation;
})(window);