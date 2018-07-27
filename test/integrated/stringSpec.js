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
});