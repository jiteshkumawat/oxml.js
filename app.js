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

            // ACT
            worksheet.cell(1, 1).style({
                fill: {
                    gradient: {
                        degree: 90,
                        stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                    }
                }
            });
            worksheet.cell(1, 1).style({
                bold: true
            });
            
            // Assert
            workbook.download('/demo.xlsx').then(function (zip) {
                expect(zip.files["workbook/style2.xml"]).toBeDefined();
                zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                    console.log(data);
                    var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                    expect(index1).toBeGreaterThan(-1);
                    done();
                });
            }).catch(function () {
                done.fail();
            });
        })();
    });