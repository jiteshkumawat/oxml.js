const oxml = require('../../../dist/oxml.js');
describe('font style', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('default font defined at first index for all cells', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            font: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<fonts count="1"><font></font></fonts>');
                expect(indexFont).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font defined with JSON values', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            bold: true,
            italic: true,
            underline: true,
            fontColor: '00ff00',
            fontSize: 14,
            fontName: 'Arial',
            strike: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><strike/><i/><b/><u/><sz val="14"/><color rgb="00ff00"/><name val="Arial"/></font>');
                expect(indexFont).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font reused among different cells', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            bold: true,
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 2).style({
            bold: true,
            fontColor: '00ff00',
            fontSize: 14
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                var indexFontCount = data.indexOf('<fonts count="2">');
                expect(indexFont).toBeGreaterThan(-1);
                expect(indexFontCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font updated if only one defined', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 1).style({
            bold: true,
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                var indexFontCount = data.indexOf('<fonts count="2">');
                expect(indexFont).toBeGreaterThan(-1);
                expect(indexFontCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font not updated if used by other cell', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 2).style({
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 1).style({
            bold: true,
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                var indexFontCount = data.indexOf('<fonts count="3">');
                expect(indexFont).toBeGreaterThan(-1);
                expect(indexFontCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font fetched if other style updated', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 2).style({
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                } 
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><sz val="14"/><color rgb="00ff00"/></font>');
                var indexFontCount = data.indexOf('<fonts count="2">');
                expect(indexFont).toBeGreaterThan(-1);
                expect(indexFontCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font merged in single cell', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            bold: true
        });
        cell.style({
            fontColor: '00ff00'
        }).style({
            fontSize: 14
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                expect(indexFont).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font defined with JSON values for collection of cells', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            bold: true,
            italic: true,
            underline: true,
            fontColor: '00ff00',
            fontSize: 14,
            fontName: 'Arial',
            strike: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><strike/><i/><b/><u/><sz val="14"/><color rgb="00ff00"/><name val="Arial"/></font>');
                expect(indexFont).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font merged in a row', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            bold: true
        });
        row.style({
            fontColor: '00ff00'
        }).style({
            fontSize: 14
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                expect(indexFont).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('font reused among different rows', (done) => {
        // Act
        worksheet.row(1, 1, [1, 2, 3]).style({
            bold: true,
            fontColor: '00ff00',
            fontSize: 14
        });
        worksheet.row(2, 1, [1, 2, 3]).style({
            bold: true,
            fontColor: '00ff00',
            fontSize: 14
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFont = data.indexOf('<font><b/><sz val="14"/><color rgb="00ff00"/></font>');
                var indexFontCount = data.indexOf('<fonts count="2">');
                expect(indexFont).toBeGreaterThan(-1);
                expect(indexFontCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});