const oxml = require('../../dist/oxml.js');
describe('shared string values', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('can be determined with JSON', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 'DummyString', type: 'sharedString' });

        // ASSERT
        expect(cell.value).toBe(0);
        expect(cell.type).toBe('sharedString');
    });

    it('can be set with value of already saved shared string', function () {
        // ACT
        var cell1 = worksheet.cell(1, 1, { value: 'DummyString', type: 'sharedString' });
        var cell2 = worksheet.cell(1, 2, { value: cell1.value, type: 'sharedString' });
        // ASSERT
        expect(cell1.value).toBe(0);
        expect(cell1.type).toBe('sharedString');
        expect(cell2.value).toBe(0);
        expect(cell2.type).toBe('sharedString');
    });

    it('shares value with similar strings', function () {
        // ACT
        var cell1 = worksheet.cell(1, 1, { value: 'DummyString', type: 'sharedString' });
        var cell2 = worksheet.cell(1, 2, { value: 'DummyString', type: 'sharedString' });

        // ASSERT
        expect(cell1.value).toBe(0);
        expect(cell2.value).toBe(0);
    });

    it('creates a shared string file', function (done) {
        // ACT
        var cell1 = worksheet.cell(1, 1, { value: 'DummyString', type: 'sharedString' });
        var cell2 = worksheet.cell(1, 2, { value: 'DummyString', type: 'sharedString' });

        // ASSERT
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/sharedstrings.xml"]).toBeDefined();
            zip.file('workbook/sharedstrings.xml').async('string').then(function (data) {
                var index = data.indexOf('<si><t>DummyString</t></si>');
                var index2 = data.lastIndexOf('<si><t>DummyString</t></si>');
                expect(index).toBeGreaterThan(-1);
                expect(index).toBe(index2);
                done();
            });
        });
    });
});