define([], function () {
    var generateContent = function (_table) {
        var tableContent = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table totalsRowShown="0" ref="' + _table.fromCell +
            ':' + _table.toCell + '" displayName="' + _table.displayName + '" name="' + _table.displayName +
            '" id="' + _table.rid + '" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<tableColumns count="' + _table.columns.length + '">';
        for (var index = 0; index < _table.columns.length; index++)
            tableContent += '<tableColumn name="' + _table.columns[index] + '" id="' + (index + 1) + '"/>';
        tableContent += '</tableColumns></table>';
        return tableContent;
    };

    var attach = function (_table, file) {
        var content = generateContent(_table);
        file.addFile(content, "table" + _table.rid + ".xml", "workbook/tables");
    };

    var addTable = function (displayName, fromCell, toCell, columns, relId) {
        var _table = {
            rid: relId,
            displayName: displayName,
            columns: columns,
            fromCell: fromCell,
            toCell: toCell,
        };
        return {
            _table: _table,
            rid: relId,
            generateContent: function () {
                generateContent(_table);
            },
            attach: function (file) {
                attach(_table, file);
            }
        };
    };

    return { addTable: addTable };
});