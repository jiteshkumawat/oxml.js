define([], function () {
    "use strict";

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

        compressedFile.saveFile = function (fileName, callback) {
            return zip.generateAsync({ type: "blob" })
                .then(function (content) {
                    try {
                        if (window.saveAs) {
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

                        if (callback) {
                            callback();
                        } else if (typeof Promise !== "undefined") {
                            return new Promise(function (resolve, reject) {
                                resolve();
                            });
                        }
                    } catch (err) {
                        if (callback) {
                            callback("Err: Not able to create file object.");
                        } else if (typeof Promise !== "undefined") {
                            return new Promise(function (resolve, reject) {
                                reject("Err: Not able to create file object.");
                            });
                        }
                    }
                });
        };

        return compressedFile;
    };

    return {
        createFile: createCompressedFile
    };
});