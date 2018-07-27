const oxml = require('../../../dist/oxml.js');
describe('row method', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('creates new instance of row with values', () => {
        // ACT
        var row = worksheet.row(1,1,['Some','Dummy','Value']);

        // ASSERT
        expect(row.cellIndices).toMatch(['A1', 'B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Some');
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('B');
        expect(row.cells[1].value).toBe('Dummy');
        expect(row.cells[1].cellIndex).toBe('B1');
        expect(row.cells[1].type).toBe('string');
        expect(row.cells[2].rowIndex).toBe(1);
        expect(row.cells[2].columnIndex).toBe('C');
        expect(row.cells[2].value).toBe('Value');
        expect(row.cells[2].cellIndex).toBe('C1');
        expect(row.cells[2].type).toBe('string');
    });

    it('updates value and return updated with set method', () => {
        // ACT
        var row = worksheet.row(1,1,[1,2,3]).set(['Some','Dummy','Value']);

        // ASSERT
        expect(row.cellIndices).toMatch(['A1', 'B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Some');
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('B');
        expect(row.cells[1].value).toBe('Dummy');
        expect(row.cells[1].cellIndex).toBe('B1');
        expect(row.cells[1].type).toBe('string');
        expect(row.cells[2].rowIndex).toBe(1);
        expect(row.cells[2].columnIndex).toBe('C');
        expect(row.cells[2].value).toBe('Value');
        expect(row.cells[2].cellIndex).toBe('C1');
        expect(row.cells[2].type).toBe('string');
    });

    it('updates value of instance with set method', () => {
        // ACT
        var row = worksheet.row(1,1,[1,2,3]);
        row.set(['Some','Dummy','Value']);

        // ASSERT
        expect(row.cellIndices).toMatch(['A1', 'B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Some');
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('B');
        expect(row.cells[1].value).toBe('Dummy');
        expect(row.cells[1].cellIndex).toBe('B1');
        expect(row.cells[1].type).toBe('string');
        expect(row.cells[2].rowIndex).toBe(1);
        expect(row.cells[2].columnIndex).toBe('C');
        expect(row.cells[2].value).toBe('Value');
        expect(row.cells[2].cellIndex).toBe('C1');
        expect(row.cells[2].type).toBe('string');
    });

    it('do not change collection of cells with set', () => {
        // ACT
        var row = worksheet.row(1,1,[1,2,3]);
        row.set(['Some']);

        // ASSERT
        expect(row.cellIndices).toMatch(['A1', 'B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Some');
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('B');
        expect(row.cells[1].value).toBe(2);
        expect(row.cells[1].cellIndex).toBe('B1');
        expect(row.cells[1].type).toBe('numeric');
        expect(row.cells[2].rowIndex).toBe(1);
        expect(row.cells[2].columnIndex).toBe('C');
        expect(row.cells[2].value).toBe(3);
        expect(row.cells[2].cellIndex).toBe('C1');
        expect(row.cells[2].type).toBe('numeric');
    });

    it('do not update cells out of range', () => {
        // ACT
        var row = worksheet.row(1,1,[1,2,3]);
        row.set(['Some', 'Dummy', 'Text', 'Updated']);

        // ASSERT
        expect(row.cellIndices).toMatch(['A1', 'B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Some');
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('B');
        expect(row.cells[1].value).toBe('Dummy');
        expect(row.cells[1].cellIndex).toBe('B1');
        expect(row.cells[1].type).toBe('string');
        expect(row.cells[2].rowIndex).toBe(1);
        expect(row.cells[2].columnIndex).toBe('C');
        expect(row.cells[2].value).toBe('Text');
        expect(row.cells[2].cellIndex).toBe('C1');
        expect(row.cells[2].type).toBe('string');
        expect(row.cells.length).toBe(3);
        expect(worksheet.cell(1,4).value).toBeFalsy();
    });

    it('returns all the values starting from rowIndex, columnIndex in a row', () => {
        // ACT
        worksheet.row(1,1,['Some', 'Dummy', 'Text']);
        row = worksheet.row(1,2);

        // ASSERT
        expect(row.cellIndices).toMatch(['B1', 'C1']);
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('B');
        expect(row.cells[0].value).toBe('Dummy');
        expect(row.cells[0].cellIndex).toBe('B1');
        expect(row.cells[0].type).toBe('string');
        expect(row.cells[1].rowIndex).toBe(1);
        expect(row.cells[1].columnIndex).toBe('C');
        expect(row.cells[1].value).toBe('Text');
        expect(row.cells[1].cellIndex).toBe('C1');
        expect(row.cells[1].type).toBe('string');
        expect(row.cells.length).toBe(2);
    });

    it('do not return any cell, if not defined starting from rowIndex, columnIndex in a row', () => {
        // ACT
        var row = worksheet.row(1,2);

        // ASSERT
        expect(row.cells.length).toBe(0);
        expect(row.cellIndices.length).toBe(0);
    });
});