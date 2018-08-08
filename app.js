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
            worksheet.grid(1,1,[
                ['Title1', 'Title2', 'Title3'],
                [1,2,3],
                [2,5,3]
            ]);
            var table = worksheet.table('Table1', 'A1','C3');
    
            // ACT
            table.style({
                fontSize: 12
            }).style({
                firstRow: {
                    bold: true,
                    fontColor: 'ffffff',
                    fill: {
                        pattern: 'solid',
                        color: '000000'
                    }
                }
            }).style({
                evenRow: {
                    fill: {
                        pattern: 'solid',
                        color: 'aaaaaa'
                    }
                }
            }).style({
                oddRow: {
                    fill: {
                        pattern: 'solid',
                        color: 'eeeeee'
                    }
                }
            }, 'tableStyle1');
            workbook.download('tmp.xlsx');
        })();
    });