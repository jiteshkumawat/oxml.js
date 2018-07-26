const oxml = require('../../dist/oxml.js');
describe('shared formula values', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('can be determined with JSON in column', function () {
        // ACT
        var cells = worksheet.sharedFormula('C1', 'C3', '(A1+B1)');

        // ASSERT
        expect(cells.cellIndices).toMatch(['C1', 'C2', 'C3']);
        expect(cells.cells[0].formula).toBe('(A1+B1)');
        expect(cells.cells[0].rowIndex).toBe(1);
        expect(cells.cells[0].columnIndex).toBe('C');
        expect(cells.cells[0].cellIndex).toBe('C1');
        expect(cells.cells[0].type).toBe('sharedFormula');
        expect(cells.cells[0].sharedFormulaIndex).toBe(0);
        expect(cells.cells[1].formula).toBeUndefined();
        expect(cells.cells[1].rowIndex).toBe(2);
        expect(cells.cells[1].columnIndex).toBe('C');
        expect(cells.cells[1].cellIndex).toBe('C2');
        expect(cells.cells[1].type).toBe('sharedFormula');
        expect(cells.cells[1].sharedFormulaIndex).toBe(0);
        expect(cells.cells[2].formula).toBeUndefined();
        expect(cells.cells[2].rowIndex).toBe(3);
        expect(cells.cells[2].columnIndex).toBe('C');
        expect(cells.cells[2].cellIndex).toBe('C3');
        expect(cells.cells[2].type).toBe('sharedFormula');
        expect(cells.cells[2].sharedFormulaIndex).toBe(0);
    });

    it('can be determined with JSON in row', function () {
        // ACT
        var cells = worksheet.sharedFormula('A3', 'C3', '(A1+A2)');

        // ASSERT
        expect(cells.cellIndices).toMatch(['A3', 'B3', 'C3']);
        expect(cells.cells[0].formula).toBe('(A1+A2)');
        expect(cells.cells[0].rowIndex).toBe(3);
        expect(cells.cells[0].columnIndex).toBe('A');
        expect(cells.cells[0].cellIndex).toBe('A3');
        expect(cells.cells[0].type).toBe('sharedFormula');
        expect(cells.cells[0].sharedFormulaIndex).toBe(0);
        expect(cells.cells[1].formula).toBeUndefined();
        expect(cells.cells[1].rowIndex).toBe(3);
        expect(cells.cells[1].columnIndex).toBe('B');
        expect(cells.cells[1].cellIndex).toBe('B3');
        expect(cells.cells[1].type).toBe('sharedFormula');
        expect(cells.cells[1].sharedFormulaIndex).toBe(0);
        expect(cells.cells[2].formula).toBeUndefined();
        expect(cells.cells[2].rowIndex).toBe(3);
        expect(cells.cells[2].columnIndex).toBe('C');
        expect(cells.cells[2].cellIndex).toBe('C3');
        expect(cells.cells[2].type).toBe('sharedFormula');
        expect(cells.cells[2].sharedFormulaIndex).toBe(0);
    });

    it('not defined without formula', function () {
        // ACT
        var cells = worksheet.sharedFormula('C1', 'C2', undefined);

        // ASSERT
        expect(cells).toBeUndefined();
    });

    it('not defined without Cell Indices', function () {
        // ACT
        var cells1 = worksheet.sharedFormula(undefined, 'C2', '(A1+B1)');
        var cells2 = worksheet.sharedFormula('C1', undefined, '(A1+B1)');

        // ASSERT
        expect(cells1).toBeUndefined();
        expect(cells2).toBeUndefined();
    });

    it('formula can be defined as JSON', function () {
        // ACT
        var cells = worksheet.sharedFormula('C1', 'C2', { formula: '(A1+B1)' });

        // ASSERT
        expect(cells).toBeDefined();
        expect(cells.cellIndices).toMatch(['C1', 'C2']);
        expect(cells.cells[0].formula).toBe('(A1+B1)');
        expect(cells.cells[0].rowIndex).toBe(1);
        expect(cells.cells[0].columnIndex).toBe('C');
        expect(cells.cells[0].cellIndex).toBe('C1');
        expect(cells.cells[0].type).toBe('sharedFormula');
        expect(cells.cells[0].sharedFormulaIndex).toBe(0);
        expect(cells.cells[1].formula).toBeUndefined();
        expect(cells.cells[1].rowIndex).toBe(2);
        expect(cells.cells[1].columnIndex).toBe('C');
        expect(cells.cells[1].cellIndex).toBe('C2');
        expect(cells.cells[1].type).toBe('sharedFormula');
        expect(cells.cells[1].sharedFormulaIndex).toBe(0);
    });

    it('formula can be defined as JSON with cached value', function () {
        // ACT
        var cells = worksheet.sharedFormula('C1', 'C2', { formula: '(A1+B1)', value: 'dummyValue' });

        // ASSERT
        expect(cells).toBeDefined();
        expect(cells.cellIndices).toMatch(['C1', 'C2']);
        expect(cells.cells[0].formula).toBe('(A1+B1)');
        expect(cells.cells[0].rowIndex).toBe(1);
        expect(cells.cells[0].columnIndex).toBe('C');
        expect(cells.cells[0].cellIndex).toBe('C1');
        expect(cells.cells[0].type).toBe('sharedFormula');
        expect(cells.cells[0].sharedFormulaIndex).toBe(0);
        expect(cells.cells[0].value).toBe('dummyValue');
        expect(cells.cells[1].formula).toBeUndefined();
        expect(cells.cells[1].rowIndex).toBe(2);
        expect(cells.cells[1].columnIndex).toBe('C');
        expect(cells.cells[1].cellIndex).toBe('C2');
        expect(cells.cells[1].type).toBe('sharedFormula');
        expect(cells.cells[1].value).toBe('dummyValue');
        expect(cells.cells[1].sharedFormulaIndex).toBe(0);
    });

    it('formula can be defined as JSON with cached value function', function () {
        // ACT
        var cells = worksheet.sharedFormula('C1', 'C2', { formula: '(A1+B1)', value: (rowIndex, columnIndex) => { return 2; } });

        // ASSERT
        expect(cells).toBeDefined();
        expect(cells.cellIndices).toMatch(['C1', 'C2']);
        expect(cells.cells[0].formula).toBe('(A1+B1)');
        expect(cells.cells[0].rowIndex).toBe(1);
        expect(cells.cells[0].columnIndex).toBe('C');
        expect(cells.cells[0].cellIndex).toBe('C1');
        expect(cells.cells[0].type).toBe('sharedFormula');
        expect(cells.cells[0].sharedFormulaIndex).toBe(0);
        expect(cells.cells[0].value).toBeDefined();
        expect(cells.cells[1].formula).toBeUndefined();
        expect(cells.cells[1].rowIndex).toBe(2);
        expect(cells.cells[1].columnIndex).toBe('C');
        expect(cells.cells[1].cellIndex).toBe('C2');
        expect(cells.cells[1].type).toBe('sharedFormula');
        expect(cells.cells[1].value).toBeDefined();
        expect(cells.cells[1].sharedFormulaIndex).toBe(0);
    });
});