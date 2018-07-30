const oxml = require('../../dist/oxml.js');
describe('formula values', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('defined with formula type of values', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { type: 'formula', formula: '(A1+A2)' });

        // ASSERT
        expect(cell.value).toBeUndefined();
        expect(cell.type).toBe('formula');
        expect(cell.formula).toBe('(A1+A2)');
    });

    it('store cached values', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { type: 'formula', formula: '(A1+A2)', value: 5 });

        // ASSERT
        expect(cell.value).toBe(5);
        expect(cell.type).toBe('formula');
        expect(cell.formula).toBe('(A1+A2)');
    });

    it('workbook with formula values can be downloaded', function (done) {
        // ACT
        var cell = worksheet.cell(2, 2, { type: 'formula', formula: '(A1+A2)', value: 'dummyValue' });
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            zip.file('workbook/sheets/sheet1.xml').async('string').then(function(data){
                var index = data.indexOf('<v>dummyValue</v>');
                expect(index).toBeGreaterThan(-1);
                done();
            });
        });

        // ASSERT
        expect(cell.value).toBeDefined();
        expect(cell.type).toBe('formula');
        expect(cell.formula).toBe('(A1+A2)');
    });

    it('store cached values with function', function (done) {
        // ACT
        var cell = worksheet.cell(2, 2, { type: 'formula', formula: '(A1+A2)', value: (rowIndex, columnIndex) => { return 2; } });
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            zip.file('workbook/sheets/sheet1.xml').async('string').then(function(data){
                var index = data.indexOf('<v>2</v>');
                expect(index).toBeGreaterThan(-1);
                done();
            });
        });

        // ASSERT
        expect(cell.value).toBeDefined();
        expect(typeof cell.value).toBe("function");
        expect(cell.type).toBe('formula');
        expect(cell.formula).toBe('(A1+A2)');
    });
});