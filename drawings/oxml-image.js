define([], function () {

    var addBlobImage = function (_zip, fileName, file) {
        _zip.addFile(file, fileName, "workbook/media");
    };

    var addBase64Image = function (_zip, fileName, file) {
        _zip.addFile(file, fileName, "workbook/media", { base64: true });
    };

    var addImage = function (fileName, _sheet, image, _zip) {
        if (typeof window !== "undefined") {
            if (image instanceof Blob) addBlobImage(_zip, fileName, image);
            else if (typeof image === "string") addBase64Image(_zip, fileName, image);
            else return;
        } else {
            _zip.addFile(file, fileName, "workbook/media");
        }
    };

    return {
        addImage: addImage
    };
});