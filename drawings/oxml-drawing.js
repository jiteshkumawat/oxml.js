define(['oxml_rels'], function (oxmlRels) {

    var generateContent = function (_drawing) {
        var content = '<?xml version="1.0" encoding="UTF-8" standalone="true"?><xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">';
        switch (drawing.position) {
            case "absolute":
                content += '<xdr:absoluteAnchor>';
                content += '<xdr:pos x="' + drawing.x + '" y="' + drawing.y + '"/>';
                content += '<xdr:ext cx="' + drawing.width + '" cy="' + drawing.height + '"/>';
                content += '<xdr:clientData/></xdr:absoluteAnchor>';
                break;
            case "oneCell":
                content += '<xdr:oneCellAnchor>';
                content += ''
                break;
        }
        content += '</xdr:wsDr>';
        return content;
    };
    var attach = function (_drawing, file) {

    };

    var addImage = function (_sheet, option) {
        if (!_sheet._workBook.drawings) _sheet._workBook.drawings = [];
        var image = {}, drawing = {};
        if (!_sheet._workBook.media) _sheet._workBook.media = [];
        _sheet._workBook.media.push(image);
        drawing.id = _sheet.addFile("drawing");
        if (!option.position || (option.position !== "absolute" && option.position !== "oneCell" && option.position !== "twoCell"))
            option.position = "absolute";
        drawing.position = option.position;
        drawing.height = option.height;
        drawing.width = option.width;
        drawing.x = option.x;
        drawing.y = option.y;
        drawing.fromColumn = option.fromColumn;

        drawing.attach = function (file) {
            attach(_sheet._workBook.drawings, file);
        }
        _sheet._workBook.drawings.push(drawing);
        return drawing;
    };

    return {
        addImage: addImage
    };
});