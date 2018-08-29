define(['contentString'], function (ContentString) {
    var XMLContentString = function (options) {
        this.className = "XMLContentString";
        this.rootNode = options.rootNode;
        this.nameSpaces = options.nameSpaces;
        this.template = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        var contentTemplate = new ContentString('<{0}{1}{3}>{2}</{0}>'), nameSpaceContent = '';
        var attributes = "";

        for (var index = 0; index < this.nameSpaces.length; index++) {
            var nameSpaceTemplate = new ContentString(' xmlns{0}="{1}"');
            if (typeof this.nameSpaces[index] === "string") {
                nameSpaceContent += nameSpaceTemplate.format('', this.nameSpaces[index]);
            } else {
                nameSpaceContent += nameSpaceTemplate.format(
                    this.nameSpaces[index].attribute ? (':' + this.nameSpaces[index].attribute) : '',
                    this.nameSpaces[index].value
                );
            }
        }
        if (options.attributes) {
            for (var attr in options.attributes) {
                var attributesContent = new ContentString(' {0}="{1}"');
                attributes += attributesContent.format(attr, options.attributes[attr]);
            }
        }

        this.template += contentTemplate.format(this.rootNode, nameSpaceContent, "{0}", attributes);
    };
    XMLContentString.prototype = Object.create(ContentString.prototype);

    return XMLContentString;
});