define([], function () {
    "use strict";

    var FileHandler = function () { this.className = "FileHandler"; };
    FileHandler.prototype = function () {
        var CompressedFile = function (zip, fs) { this.className = "CompressedFile"; this._zip = zip; this._fs = fs; };
        CompressedFile.prototype = function () {
            var addFile = function (content, fileName, filePath, options) {
                var path;
                if (filePath) {
                    path = this._zip.folder(filePath);
                }
                if (!options) (path || this._zip).file(fileName, content);
                else (path || this._zip).file(fileName, content, options);
            };

            var saveFile = function (fileName, callback) {
                var zip = this._zip;
                var fs = this._fs;
                /* istanbul ignore if  */
                if (typeof window !== "undefined") {
                    return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
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
                    if (callback) {
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

            return {
                addFile: addFile,
                saveFile: saveFile
            };
        }();

        return {
            createFile: function (zip, fs) {
                return new CompressedFile(zip, fs);
            }
        };
    }();

    return FileHandler;
});