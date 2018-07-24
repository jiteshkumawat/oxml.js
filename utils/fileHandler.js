define([], function () {
    "use strict";

    var createCompressedFile = function (jsZip, fs) {

        var zip = jsZip;

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
            if (typeof window !== "undefined") {
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
                                callback(zip);
                            } else if (typeof Promise !== "undefined") {
                                return new Promise(function (resolve, reject) {
                                    resolve(zip);
                                });
                            }
                        } catch (err) {
                            if (callback) {
                                callback(null, "Err: Not able to create file object.");
                            } else if (typeof Promise !== "undefined") {
                                return new Promise(function (resolve, reject) {
                                    reject("Err: Not able to create file object.");
                                });
                            }
                        }
                    });
            } else {
                if(callback){
                    zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(fileName));
                    callback(zip);
                    return;
                }
                return new Promise(function (resolve, reject) {
                    zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
                    .pipe(fs.createWriteStream(fileName));
                    resolve(zip);
                });
            }
        };

        return compressedFile;
    };

    return {
        createFile: createCompressedFile
    };
});