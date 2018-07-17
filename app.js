/* eslint-disable */
requirejs.config({
    baseUrl: '..',
    paths: {
        fileHandler: 'utils/fileHandler',
        jsZip: 'lib/jsZip.min',
        fileSaver: 'lib/fileSaver.min',
        oxml_content_types: 'oxml-content-types',
        oxml_rels: 'oxml-rels',
        oxml_sheet: 'oxml-sheet',
        oxml_workbook: 'oxml-workbook',
        oxml_xlsx: 'oxml-xlsx',
        oxml_xlsx_styles: 'styles/oxml-xlsx-styles',
        oxml_xlsx_font: 'styles/oxml-xlsx-font',
        oxml_xlsx_fill: 'styles/oxml-xlsx-fill',
        oxml_xlsx_border: 'styles/oxml-xlsx-border',
        oxml_xlsx_num_format: 'styles/oxml-xlsx-num-format',
        oxml_table: 'oxml-table',
        utils: 'utils/utils'
    }
});

requirejs(['fileSaver', 'oxml-xlsx'],
    function (fileSaver, oxmlXLSX) {
        // Example 2
        (function () {
            var workbook = oxml.xlsx();
            var worksheet = workbook.sheet('sheet1');
            // var cell = worksheet.cell(2, 2, { bold: true });
            // cell.set('Updated Value');
            // console.log(cell);
            // workbook.download('tmp.xlsx');
            debugger;
            worksheet.grid(2, 3, [
                ['Cost', 'Sales', 'Profit'],
                [10, 12],
                [9, 12],
                [11, 12],
                ['Total']
            ]);
            worksheet.sharedFormula('E3', 'E5', {
                type: 'formula', formula: '(D3 - C3)', value: function (rowIndex, columnIndex) {
                    var sale = worksheet.cell(rowIndex, columnIndex - 1).value;
                    var cost = worksheet.cell(rowIndex, columnIndex - 2).value;
                    return sale - cost;
                }
            });
            worksheet.sharedFormula('C6', 'D6', {
                type: 'formula', formula: 'SUM(C3:C5)', value: function (rowIndex, columnIndex) {
                    var column = worksheet.column(3, 3), sum = 0;
                    for (var index = 0; index < column.cells.length; index++) {
                        if (column.cells[index].value && typeof column.cells[index].value === "number") {
                            sum += column.cells[index].value;
                        }
                    }
                    return sum;
                }
            });
            worksheet.table('Table1', 'C2', 'E6', {
                sort: 1
            }).set({sort: {
                direction: "descending",
                caseSensitive: false,
                column: 2
            }});
            workbook.download('tmp.xlsx');
        })();
    });