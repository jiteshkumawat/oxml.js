const oxml = require('../../../dist/oxml.js');
describe('grid method', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('creates new instance of grid with values', () => {
        // ACT
        var grid = worksheet.grid(1, 1, [
            ['Col1', 'Col2', 'Col3'],
            [5, 2, 7],
            [8, 4, 8]
        ]);

        // ASSERT
        expect(grid.cellIndices).toMatch(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(1);
        expect(grid.cells[0].columnIndex).toBe('A');
        expect(grid.cells[0].value).toBe('Col1');
        expect(grid.cells[0].cellIndex).toBe('A1');
        expect(grid.cells[0].type).toBe('string');
        expect(grid.cells[1].rowIndex).toBe(1);
        expect(grid.cells[1].columnIndex).toBe('B');
        expect(grid.cells[1].value).toBe('Col2');
        expect(grid.cells[1].cellIndex).toBe('B1');
        expect(grid.cells[1].type).toBe('string');
        expect(grid.cells[2].rowIndex).toBe(1);
        expect(grid.cells[2].columnIndex).toBe('C');
        expect(grid.cells[2].value).toBe('Col3');
        expect(grid.cells[2].cellIndex).toBe('C1');
        expect(grid.cells[2].type).toBe('string');

        expect(grid.cells[3].rowIndex).toBe(2);
        expect(grid.cells[3].columnIndex).toBe('A');
        expect(grid.cells[3].value).toBe(5);
        expect(grid.cells[3].cellIndex).toBe('A2');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells[4].rowIndex).toBe(2);
        expect(grid.cells[4].columnIndex).toBe('B');
        expect(grid.cells[4].value).toBe(2);
        expect(grid.cells[4].cellIndex).toBe('B2');
        expect(grid.cells[4].type).toBe('numeric');
        expect(grid.cells[5].rowIndex).toBe(2);
        expect(grid.cells[5].columnIndex).toBe('C');
        expect(grid.cells[5].value).toBe(7);
        expect(grid.cells[5].cellIndex).toBe('C2');
        expect(grid.cells[5].type).toBe('numeric');

        expect(grid.cells[6].rowIndex).toBe(3);
        expect(grid.cells[6].columnIndex).toBe('A');
        expect(grid.cells[6].value).toBe(8);
        expect(grid.cells[6].cellIndex).toBe('A3');
        expect(grid.cells[6].type).toBe('numeric');
        expect(grid.cells[7].rowIndex).toBe(3);
        expect(grid.cells[7].columnIndex).toBe('B');
        expect(grid.cells[7].value).toBe(4);
        expect(grid.cells[7].cellIndex).toBe('B3');
        expect(grid.cells[7].type).toBe('numeric');
        expect(grid.cells[8].rowIndex).toBe(3);
        expect(grid.cells[8].columnIndex).toBe('C');
        expect(grid.cells[8].value).toBe(8);
        expect(grid.cells[8].cellIndex).toBe('C3');
        expect(grid.cells[8].type).toBe('numeric');
    });

    it('updates value and return updated with set method', () => {
        // ACT
        var grid = worksheet.grid(1, 1, [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]).set([
            ['Col1', 'Col2', 'Col3'],
            [5, 2, 7],
            [8, 4, 8]
        ]);

        // ASSERT
        expect(grid.cellIndices).toMatch(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(1);
        expect(grid.cells[0].columnIndex).toBe('A');
        expect(grid.cells[0].value).toBe('Col1');
        expect(grid.cells[0].cellIndex).toBe('A1');
        expect(grid.cells[0].type).toBe('string');
        expect(grid.cells[1].rowIndex).toBe(1);
        expect(grid.cells[1].columnIndex).toBe('B');
        expect(grid.cells[1].value).toBe('Col2');
        expect(grid.cells[1].cellIndex).toBe('B1');
        expect(grid.cells[1].type).toBe('string');
        expect(grid.cells[2].rowIndex).toBe(1);
        expect(grid.cells[2].columnIndex).toBe('C');
        expect(grid.cells[2].value).toBe('Col3');
        expect(grid.cells[2].cellIndex).toBe('C1');
        expect(grid.cells[2].type).toBe('string');

        expect(grid.cells[3].rowIndex).toBe(2);
        expect(grid.cells[3].columnIndex).toBe('A');
        expect(grid.cells[3].value).toBe(5);
        expect(grid.cells[3].cellIndex).toBe('A2');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells[4].rowIndex).toBe(2);
        expect(grid.cells[4].columnIndex).toBe('B');
        expect(grid.cells[4].value).toBe(2);
        expect(grid.cells[4].cellIndex).toBe('B2');
        expect(grid.cells[4].type).toBe('numeric');
        expect(grid.cells[5].rowIndex).toBe(2);
        expect(grid.cells[5].columnIndex).toBe('C');
        expect(grid.cells[5].value).toBe(7);
        expect(grid.cells[5].cellIndex).toBe('C2');
        expect(grid.cells[5].type).toBe('numeric');

        expect(grid.cells[6].rowIndex).toBe(3);
        expect(grid.cells[6].columnIndex).toBe('A');
        expect(grid.cells[6].value).toBe(8);
        expect(grid.cells[6].cellIndex).toBe('A3');
        expect(grid.cells[6].type).toBe('numeric');
        expect(grid.cells[7].rowIndex).toBe(3);
        expect(grid.cells[7].columnIndex).toBe('B');
        expect(grid.cells[7].value).toBe(4);
        expect(grid.cells[7].cellIndex).toBe('B3');
        expect(grid.cells[7].type).toBe('numeric');
        expect(grid.cells[8].rowIndex).toBe(3);
        expect(grid.cells[8].columnIndex).toBe('C');
        expect(grid.cells[8].value).toBe(8);
        expect(grid.cells[8].cellIndex).toBe('C3');
        expect(grid.cells[8].type).toBe('numeric');
    });

    it('updates value of instance with set method', () => {
        // ACT
        var grid = worksheet.grid(1, 1, [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]);
        grid.set([
            ['Col1', 'Col2', 'Col3'],
            [5, 2, 7],
            [8, 4, 8]
        ]);

        // ASSERT
        expect(grid.cellIndices).toMatch(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(1);
        expect(grid.cells[0].columnIndex).toBe('A');
        expect(grid.cells[0].value).toBe('Col1');
        expect(grid.cells[0].cellIndex).toBe('A1');
        expect(grid.cells[0].type).toBe('string');
        expect(grid.cells[1].rowIndex).toBe(1);
        expect(grid.cells[1].columnIndex).toBe('B');
        expect(grid.cells[1].value).toBe('Col2');
        expect(grid.cells[1].cellIndex).toBe('B1');
        expect(grid.cells[1].type).toBe('string');
        expect(grid.cells[2].rowIndex).toBe(1);
        expect(grid.cells[2].columnIndex).toBe('C');
        expect(grid.cells[2].value).toBe('Col3');
        expect(grid.cells[2].cellIndex).toBe('C1');
        expect(grid.cells[2].type).toBe('string');

        expect(grid.cells[3].rowIndex).toBe(2);
        expect(grid.cells[3].columnIndex).toBe('A');
        expect(grid.cells[3].value).toBe(5);
        expect(grid.cells[3].cellIndex).toBe('A2');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells[4].rowIndex).toBe(2);
        expect(grid.cells[4].columnIndex).toBe('B');
        expect(grid.cells[4].value).toBe(2);
        expect(grid.cells[4].cellIndex).toBe('B2');
        expect(grid.cells[4].type).toBe('numeric');
        expect(grid.cells[5].rowIndex).toBe(2);
        expect(grid.cells[5].columnIndex).toBe('C');
        expect(grid.cells[5].value).toBe(7);
        expect(grid.cells[5].cellIndex).toBe('C2');
        expect(grid.cells[5].type).toBe('numeric');

        expect(grid.cells[6].rowIndex).toBe(3);
        expect(grid.cells[6].columnIndex).toBe('A');
        expect(grid.cells[6].value).toBe(8);
        expect(grid.cells[6].cellIndex).toBe('A3');
        expect(grid.cells[6].type).toBe('numeric');
        expect(grid.cells[7].rowIndex).toBe(3);
        expect(grid.cells[7].columnIndex).toBe('B');
        expect(grid.cells[7].value).toBe(4);
        expect(grid.cells[7].cellIndex).toBe('B3');
        expect(grid.cells[7].type).toBe('numeric');
        expect(grid.cells[8].rowIndex).toBe(3);
        expect(grid.cells[8].columnIndex).toBe('C');
        expect(grid.cells[8].value).toBe(8);
        expect(grid.cells[8].cellIndex).toBe('C3');
        expect(grid.cells[8].type).toBe('numeric');
    });

    it('do not change collection of cells with set', () => {
        // ACT
        var grid = worksheet.grid(1, 1, [
            [1, 2, 'Col3'],
            [3, 2, 7],
            [8, 4, 8]
        ]);
        grid.set([
            ['Col1', 'Col2'],
            [5]
        ]);

        // ASSERT
        expect(grid.cellIndices).toMatch(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(1);
        expect(grid.cells[0].columnIndex).toBe('A');
        expect(grid.cells[0].value).toBe('Col1');
        expect(grid.cells[0].cellIndex).toBe('A1');
        expect(grid.cells[0].type).toBe('string');
        expect(grid.cells[1].rowIndex).toBe(1);
        expect(grid.cells[1].columnIndex).toBe('B');
        expect(grid.cells[1].value).toBe('Col2');
        expect(grid.cells[1].cellIndex).toBe('B1');
        expect(grid.cells[1].type).toBe('string');
        expect(grid.cells[2].rowIndex).toBe(1);
        expect(grid.cells[2].columnIndex).toBe('C');
        expect(grid.cells[2].value).toBe('Col3');
        expect(grid.cells[2].cellIndex).toBe('C1');
        expect(grid.cells[2].type).toBe('string');

        expect(grid.cells[3].rowIndex).toBe(2);
        expect(grid.cells[3].columnIndex).toBe('A');
        expect(grid.cells[3].value).toBe(5);
        expect(grid.cells[3].cellIndex).toBe('A2');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells[4].rowIndex).toBe(2);
        expect(grid.cells[4].columnIndex).toBe('B');
        expect(grid.cells[4].value).toBe(2);
        expect(grid.cells[4].cellIndex).toBe('B2');
        expect(grid.cells[4].type).toBe('numeric');
        expect(grid.cells[5].rowIndex).toBe(2);
        expect(grid.cells[5].columnIndex).toBe('C');
        expect(grid.cells[5].value).toBe(7);
        expect(grid.cells[5].cellIndex).toBe('C2');
        expect(grid.cells[5].type).toBe('numeric');

        expect(grid.cells[6].rowIndex).toBe(3);
        expect(grid.cells[6].columnIndex).toBe('A');
        expect(grid.cells[6].value).toBe(8);
        expect(grid.cells[6].cellIndex).toBe('A3');
        expect(grid.cells[6].type).toBe('numeric');
        expect(grid.cells[7].rowIndex).toBe(3);
        expect(grid.cells[7].columnIndex).toBe('B');
        expect(grid.cells[7].value).toBe(4);
        expect(grid.cells[7].cellIndex).toBe('B3');
        expect(grid.cells[7].type).toBe('numeric');
        expect(grid.cells[8].rowIndex).toBe(3);
        expect(grid.cells[8].columnIndex).toBe('C');
        expect(grid.cells[8].value).toBe(8);
        expect(grid.cells[8].cellIndex).toBe('C3');
        expect(grid.cells[8].type).toBe('numeric');
    });

    it('do not update cells out of range', () => {
        // ACT
        var grid = worksheet.grid(1, 1, [
            [1, 2, 3],
            [3, 2, 7],
            [8, 4, 8]
        ]);
        grid.set([
            ['Col1', 'Col2', 'Col3', 'Col4'],
            [5],
            null,
            [1, 2, 3]
        ]);

        // ASSERT
        expect(grid.cellIndices).toMatch(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(1);
        expect(grid.cells[0].columnIndex).toBe('A');
        expect(grid.cells[0].value).toBe('Col1');
        expect(grid.cells[0].cellIndex).toBe('A1');
        expect(grid.cells[0].type).toBe('string');
        expect(grid.cells[1].rowIndex).toBe(1);
        expect(grid.cells[1].columnIndex).toBe('B');
        expect(grid.cells[1].value).toBe('Col2');
        expect(grid.cells[1].cellIndex).toBe('B1');
        expect(grid.cells[1].type).toBe('string');
        expect(grid.cells[2].rowIndex).toBe(1);
        expect(grid.cells[2].columnIndex).toBe('C');
        expect(grid.cells[2].value).toBe('Col3');
        expect(grid.cells[2].cellIndex).toBe('C1');
        expect(grid.cells[2].type).toBe('string');

        expect(grid.cells[3].rowIndex).toBe(2);
        expect(grid.cells[3].columnIndex).toBe('A');
        expect(grid.cells[3].value).toBe(5);
        expect(grid.cells[3].cellIndex).toBe('A2');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells[4].rowIndex).toBe(2);
        expect(grid.cells[4].columnIndex).toBe('B');
        expect(grid.cells[4].value).toBe(2);
        expect(grid.cells[4].cellIndex).toBe('B2');
        expect(grid.cells[4].type).toBe('numeric');
        expect(grid.cells[5].rowIndex).toBe(2);
        expect(grid.cells[5].columnIndex).toBe('C');
        expect(grid.cells[5].value).toBe(7);
        expect(grid.cells[5].cellIndex).toBe('C2');
        expect(grid.cells[5].type).toBe('numeric');

        expect(grid.cells[6].rowIndex).toBe(3);
        expect(grid.cells[6].columnIndex).toBe('A');
        expect(grid.cells[6].value).toBe(8);
        expect(grid.cells[6].cellIndex).toBe('A3');
        expect(grid.cells[6].type).toBe('numeric');
        expect(grid.cells[7].rowIndex).toBe(3);
        expect(grid.cells[7].columnIndex).toBe('B');
        expect(grid.cells[7].value).toBe(4);
        expect(grid.cells[7].cellIndex).toBe('B3');
        expect(grid.cells[7].type).toBe('numeric');
        expect(grid.cells[8].rowIndex).toBe(3);
        expect(grid.cells[8].columnIndex).toBe('C');
        expect(grid.cells[8].value).toBe(8);
        expect(grid.cells[8].cellIndex).toBe('C3');
        expect(grid.cells[8].type).toBe('numeric');
        expect(grid.cells.length).toBe(9);
        expect(worksheet.cell(1, 4).value).toBeFalsy();
        expect(worksheet.cell(4, 1).value).toBeFalsy();
    });

    it('returns all the values starting from rowIndex, columnIndex in a grid', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Col1', 'Col2', 'Col3'],
            [5, 2, 7],
            [8, 4, 8]
        ]);
        var grid = worksheet.grid(2, 2);

        // ASSERT
        expect(grid.cellIndices).toMatch(['B2', 'C2', 'B3', 'C3']);
        expect(grid.cells[0].rowIndex).toBe(2);
        expect(grid.cells[0].columnIndex).toBe('B');
        expect(grid.cells[0].value).toBe(2);
        expect(grid.cells[0].cellIndex).toBe('B2');
        expect(grid.cells[0].type).toBe('numeric');
        expect(grid.cells[1].rowIndex).toBe(2);
        expect(grid.cells[1].columnIndex).toBe('C');
        expect(grid.cells[1].value).toBe(7);
        expect(grid.cells[1].cellIndex).toBe('C2');
        expect(grid.cells[1].type).toBe('numeric');
        expect(grid.cells[2].rowIndex).toBe(3);
        expect(grid.cells[2].columnIndex).toBe('B');
        expect(grid.cells[2].value).toBe(4);
        expect(grid.cells[2].cellIndex).toBe('B3');
        expect(grid.cells[2].type).toBe('numeric');
        expect(grid.cells[3].rowIndex).toBe(3);
        expect(grid.cells[3].columnIndex).toBe('C');
        expect(grid.cells[3].value).toBe(8);
        expect(grid.cells[3].cellIndex).toBe('C3');
        expect(grid.cells[3].type).toBe('numeric');
        expect(grid.cells.length).toBe(4);
    });

    it('do not return any cell, if not defined starting from rowIndex, columnIndex in a grid', () => {
        // ACT
        var grid = worksheet.grid(1, 2);

        // ASSERT
        expect(grid.cells.length).toBe(0);
        expect(grid.cellIndices.length).toBe(0);
    });
});