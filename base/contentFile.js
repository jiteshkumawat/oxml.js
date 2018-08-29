define([], function () {
    var ContentFile = function (file) { this.className = "ContentFile"; this._file = file; };

    ContentFile.prototype = function () {
        var attach = function (zipFile) {
            var fileContent = this.generateContent(zipFile);
            zipFile.addFile(fileContent, this.fileName, this.folderName);
            if (this.attachChild) this.attachChild(zipFile);
        };

        return { attach: attach };
    }();

    return ContentFile;
});