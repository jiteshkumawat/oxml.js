const oxml = require('../../dist/oxml.js');
describe('workbook basic tests', function () {
    var workbook;

    beforeEach(function () {
        workbook = oxml.xlsx();
    });

    afterEach(function(){
        workbook = null;
    });

    it('workbook attributes', function () {
        // ASSERT
        expect(workbook).toBeDefined();
        expect(workbook.sheet).toBeDefined();
        expect(typeof workbook.sheet).toBe("function");
        expect(workbook.download).toBeDefined();
        expect(typeof workbook.download).toBe('function');
    });

    it('sheet is added to workbook', function () {
        // ACT
        var worksheet = workbook.sheet('sheet1');

        // ASSERT
        expect(worksheet).toBeDefined();
        expect(worksheet.sharedFormula).toBeDefined();
        expect(typeof worksheet.sharedFormula).toBe("function");
        expect(worksheet.cell).toBeDefined();
        expect(typeof worksheet.cell).toBe("function");
        expect(worksheet.row).toBeDefined();
        expect(typeof worksheet.row).toBe("function");
        expect(worksheet.column).toBeDefined();
        expect(typeof worksheet.column).toBe("function");
        expect(worksheet.grid).toBeDefined();
        expect(typeof worksheet.grid).toBe("function");
        expect(worksheet.table).toBeDefined();
        expect(typeof worksheet.table).toBe("function");
    });

    it('workbook is downloaded with promise', function (done) {
        // ARRANGE
        workbook.sheet('sheet1');

        // ACT
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["[Content_Types].xml"]).toBeDefined();
            expect(zip.files["_rels/"]).toBeDefined();
            expect(zip.files["_rels/.rels"]).toBeDefined();
            expect(zip.files["workbook/"]).toBeDefined();
            expect(zip.files["workbook/_rels/"]).toBeDefined();
            expect(zip.files["workbook/_rels/workbook.xml.rels"]).toBeDefined();
            expect(zip.files["workbook/sheets/"]).toBeDefined();
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            expect(zip.files["workbook/workbook.xml"]).toBeDefined();
            done();
        }).catch(function () {
            done.fail();
        });
    });

    it('workbook is downloaded with callback', function (done) {
        // ARRANGE
        workbook.sheet('sheet1');

        // ACT
        workbook.download(__dirname + '/demo.xlsx', function (zip) {
            // ASSERT
            expect(zip.files["[Content_Types].xml"]).toBeDefined();
            expect(zip.files["_rels/"]).toBeDefined();
            expect(zip.files["_rels/.rels"]).toBeDefined();
            expect(zip.files["workbook/"]).toBeDefined();
            expect(zip.files["workbook/_rels/"]).toBeDefined();
            expect(zip.files["workbook/_rels/workbook.xml.rels"]).toBeDefined();
            expect(zip.files["workbook/sheets/"]).toBeDefined();
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            expect(zip.files["workbook/workbook.xml"]).toBeDefined();
            done();
        });
    });

    it('multiple sheet is supported', function(done){
        // ACT
        var worksheet1 = workbook.sheet('sheet1');
        var worksheet2 = workbook.sheet();
        var worksheet3 = workbook.sheet();

        // ASSERT
        expect(worksheet1).toBeDefined();
        expect(worksheet2).toBeDefined();
        expect(worksheet3).toBeDefined();
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["[Content_Types].xml"]).toBeDefined();
            expect(zip.files["_rels/"]).toBeDefined();
            expect(zip.files["_rels/.rels"]).toBeDefined();
            expect(zip.files["workbook/"]).toBeDefined();
            expect(zip.files["workbook/_rels/"]).toBeDefined();
            expect(zip.files["workbook/_rels/workbook.xml.rels"]).toBeDefined();
            expect(zip.files["workbook/sheets/"]).toBeDefined();
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            expect(zip.files["workbook/sheets/sheet2.xml"]).toBeDefined();
            expect(zip.files["workbook/sheets/sheet3.xml"]).toBeDefined();
            expect(zip.files["workbook/workbook.xml"]).toBeDefined();
            done();
        }).catch(function () {
            done.fail();
        });
    });
});