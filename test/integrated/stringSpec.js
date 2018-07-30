const oxml = require('../../dist/oxml.js');
describe('string values', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('can be determined with string literals', function () {
        // ACT
        var cell = worksheet.cell(1, 1, 'DummyString');

        // ASSERT
        expect(cell.value).toBe('DummyString');
        expect(cell.type).toBe('string');
    });

    it('can be determined with JSON without type', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 'DummyString' });

        // ASSERT
        expect(cell.value).toBe('DummyString');
        expect(cell.type).toBe('string');
    });

    it('can be determined with JSON and type', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 'DummyString', type: "string" });

        // ASSERT
        expect(cell.value).toBe('DummyString');
        expect(cell.type).toBe('string');
    });

    it('can download workbook with string values', (done) => {
        // ACT
        worksheet.cell(1, 1, { value: "dummyString" });

        // ACT
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["workbook/sheets/sheet1.xml"]).toBeDefined();
            zip.file('workbook/sheets/sheet1.xml').async('string').then(function(data){
                var index = data.indexOf('<c r="A1" t="inlineStr" ><is><t>dummyString</t></is></c>');
                expect(index).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});