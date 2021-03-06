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
        utils: 'utils/utils',
        contentFile: 'base/contentFile',
        contentString: 'base/contentString',
        xmlContentString: 'base/xmlContentString'
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

            var a = worksheet.row(1, 1, ["test"]);
            worksheet.merge("A1:B2");
            worksheet.row(1, 1, null, {
                bold: true,
                italic: true,
                hAlignment: "center",
                vAlignment: "center"
            });

            // Assert
            workbook.download('/demo.xlsx').then(function (zip) {
                expect(zip.files["workbook/style2.xml"]).toBeDefined();
                zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                    var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                    expect(index1).toBeGreaterThan(-1);
                    done();
                });
            }).catch(function () {
                done.fail();
            });

            workbook.download('demo.xlsx').then(function (zip) {
                // ASSERT
                expect(zip.files["workbook/tables/table1.xml"]).toBeDefined();
                zip.file("workbook/tables/table1.xml").async('string').then((data) => {
                    var index = data.indexOf('<tableStyleInfo name="tableStyle1" showColumnStripes="0" showRowStripes="0" showLastColumn="0" showFirstColumn="0"/>');
                    expect(index).toBeGreaterThan(-1);
                });
                zip.file('workbook/style2.xml').async('string').then(function (data) {
                    var index1 = data.indexOf('<dxfs count="2"><dxf><font></font></dxf><dxf><font><b/><color rgb="ffffff"/></font><fill><patternFill patternType="solid"><bgColor rgb="000000"/></patternFill></fill></dxf></dxfs>');
                    var index2 = data.indexOf('<tableStyles count="1"><tableStyle count="2" name="tableStyle1"><tableStyleElement dxfId="0" type="wholeTable"/><tableStyleElement dxfId="1" type="headerRow"/></tableStyle></tableStyles>');
                    expect(index1).toBeGreaterThan(-1);
                    expect(index2).toBeGreaterThan(-1);
                    done();
                });
            }).catch(function () {
                done.fail();
            });
        })();
    });