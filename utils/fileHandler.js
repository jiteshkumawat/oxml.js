"use strict";

(function (window) {
    var createCompressedFile = function () {
        if (!window.JSZip) {
            return;
        }
        var zip = new JSZip();

        var compressedFile = {
            _zip: zip
        };

        compressedFile.addFile = function (content, fileName, filePath) {
            var path;
            if (filePath) {
                path = zip.folder(filePath);
            }
            (path || zip).file(fileName, content);
        };

        compressedFile.saveFile = function (fileName) {
            zip.generateAsync({type: "blob"})
                .then(function(content){
                    if(window.saveAs){
                        return saveAs(content, fileName);
                    }
                    var url = window.URL.createObjectURL(content);
                    var element = document.createElement('a');
                    element.setAttribute('href', url);
                    element.setAttribute('download', fileName);
        
                    element.style.display = 'none';
                    document.body.appendChild(element);
                    element.click();
                    document.body.removeChild(element);
                });
        };

        return compressedFile;
    };

    if (!window.oxml) {
        window.oxml = {};
    }

    if (!window.oxml.util) {
        window.oxml.util = {};
    }

    window.oxml.util.createFile = createCompressedFile;
})(window);