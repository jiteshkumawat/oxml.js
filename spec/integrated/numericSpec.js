const oxml = require('../../dist/oxml.js');
describe('numeric values', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('can be determined with numeric literals', function () {
        // ACT
        var cell = worksheet.cell(1, 1, 1);

        // ASSERT
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
    });

    it('can be determined with JSON without type', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 1 });

        // ASSERT
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
    });

    it('can be determined with JSON and type', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 1, type: "numeric" });

        // ASSERT
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
    });
});