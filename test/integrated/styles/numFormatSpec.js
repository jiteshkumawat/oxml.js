const oxml = require('../../../dist/oxml.js');
describe('numFormat style', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('numformat defined with default values', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            numberFormat: "#,##0;(#,##0)"
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0;(#,##0)"/>');
                expect(indexNumFormat).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('numformat update with values', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            numberFormat: "#,##0;(#,##0)"
        });
        cell.style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0.00;(#,##0.00)"/>');
                var indexOldNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0;(#,##0)"/>');
                expect(indexNumFormat).toBeGreaterThan(-1);
                expect(indexOldNumFormat).toBe(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('numformat reused among different cells', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });
        worksheet.cell(1, 2).style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0.00;(#,##0.00)"/>');
                var indexNumFormatCount = data.indexOf('<numFmts count="1">');
                expect(indexNumFormat).toBeGreaterThan(-1);
                expect(indexNumFormatCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('numformat defined for collection of cells', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0.00;(#,##0.00)"/>');
                expect(indexNumFormat).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('numformat reused among different rows', (done) => {
        // Act
        worksheet.row(1, 1, [1, 2, 3]).style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });
        worksheet.row(2, 1, [1, 2, 3]).style({
            numberFormat: "#,##0.00;(#,##0.00)"
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexNumFormat = data.indexOf('<numFmt numFmtId="200" formatCode="#,##0.00;(#,##0.00)"/>');
                var indexNumFormatCount = data.indexOf('<numFmts count="1">');
                expect(indexNumFormat).toBeGreaterThan(-1);
                expect(indexNumFormatCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});