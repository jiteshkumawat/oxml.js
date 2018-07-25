const oxml = require('../../../dist/oxml.js');
describe('cell method', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('save string data in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, 'DummyValue');

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON string data with type defined in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 'DummyValue', type: 'string' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON string data with type not defined in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 'DummyValue' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save string data in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set('DummyValue');

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON string data with type defined in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set({ value: 'DummyValue', type: 'string' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON string data with type not defined in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set({ value: 'DummyValue' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('DummyValue');
        expect(cell.type).toBe('string');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save numeric data in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, 1);

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON numeric data without type in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 1 });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON numeric data with type in cell as parameter', function () {
        // ACT
        var cell = worksheet.cell(1, 1, { value: 1, type: "numeric" });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save numeric data in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set(1);

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON numeric data without type in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set({ value: 1 });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save JSON numeric data with type in cell using set', function () {
        // ACT
        var cell = worksheet.cell(1, 1).set({ value: 1, type: "numeric" });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(1);
        expect(cell.type).toBe('numeric');
        expect(cell.cellIndex).toBe('A1');
        expect(cell.rowIndex).toBe(1);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save formula data in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 2, { type: 'formula', formula: '(A1+A2)' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBeUndefined();
        expect(cell.type).toBe('formula');
        expect(cell.cellIndex).toBe('B2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('B');
        expect(cell.set).toBeDefined();
    });

    it('save formula data in cell using set', function () {
        // ACT
        var cell = worksheet.cell(2, 2).set({ type: 'formula', formula: '(A1+A2)' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBeUndefined();
        expect(cell.type).toBe('formula');
        expect(cell.cellIndex).toBe('B2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('B');
        expect(cell.set).toBeDefined();
    });

    it('save cached formula data in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 2, { type: 'formula', formula: '(A1+A2)', value: 3 });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(3);
        expect(cell.type).toBe('formula');
        expect(cell.cellIndex).toBe('B2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('B');
        expect(cell.set).toBeDefined();
    });

    it('save cached formula data in cell using set', function () {
        // ACT
        var cell = worksheet.cell(2, 2).set({ type: 'formula', formula: '(A1+A2)', value: 3 });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(3);
        expect(cell.type).toBe('formula');
        expect(cell.cellIndex).toBe('B2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('B');
        expect(cell.set).toBeDefined();
    });

    it('save shared string data in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 1, { type: 'sharedString', value: 'DummyString' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(0);
        expect(cell.type).toBe('sharedString');
        expect(cell.cellIndex).toBe('A2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('save shared string data in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 1).set({ type: 'sharedString', value: 'DummyString' });

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe(0);
        expect(cell.type).toBe('sharedString');
        expect(cell.cellIndex).toBe('A2');
        expect(cell.rowIndex).toBe(2);
        expect(cell.columnIndex).toBe('A');
        expect(cell.set).toBeDefined();
    });

    it('set updates value in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 1, 'Value1');
        cell.set('Value2');

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('Value2');
    });

    it('set updates type in cell', function () {
        // ACT
        var cell = worksheet.cell(2, 1, 'Value1');
        cell.set(1);

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.type).toBe('numeric');
    });

    it('set returns updated value', function () {
        // ACT
        var cell = worksheet.cell(2, 1, 'Value1').set('Value2');

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.value).toBe('Value2');
    });

    it('set returns updated type', function () {
        // ACT
        var cell = worksheet.cell(2, 1, 'Value1').set(1);

        // ASSERT
        expect(cell).toBeDefined();
        expect(cell.type).toBe('numeric');
    });

    it('requires rowId and columnId', function () {
        // ACT
        var cell1 = worksheet.cell();
        var cell2 = worksheet.cell(undefined, 1);
        var cell3 = worksheet.cell(1);
        var cell4 = worksheet.cell('1', '2');
        var cell5 = worksheet.cell('1', 1);
        var cell6 = worksheet.cell(1, '1');
        var cell7 = worksheet.cell({ value: 1 }, 1);
        var cell8 = worksheet.cell([1], 1);

        // ASSERT
        expect(cell1).toBeUndefined();
        expect(cell2).toBeUndefined();
        expect(cell3).toBeUndefined();
        expect(cell4).toBeUndefined();
        expect(cell5).toBeUndefined();
        expect(cell6).toBeUndefined();
        expect(cell7).toBeUndefined();
        expect(cell8).toBeUndefined();
    });

    it('save data in multiple cells', function () {
        // ACT
        var cell1 = worksheet.cell(1, 1, 'DummyValue');
        var cell2 = worksheet.cell(1, 2, { value: 'DummyValue', type: 'string' });
        var cell3 = worksheet.cell(1, 3, { value: 'DummyValue' });
        var cell4 = worksheet.cell(1, 4).set('DummyValue');
        var cell5 = worksheet.cell(1, 5).set({ value: 'DummyValue', type: 'string' });
        var cell6 = worksheet.cell(1, 6).set({ value: 'DummyValue' });
        var cell7 = worksheet.cell(1, 7, 1);
        var cell8 = worksheet.cell(1, 8, { value: 1 });
        var cell9 = worksheet.cell(1, 9, { value: 1, type: "numeric" });
        var cell10 = worksheet.cell(1, 10).set(1);
        var cell11 = worksheet.cell(1, 11).set({ value: 1 });
        var cell12 = worksheet.cell(1, 12).set({ value: 1, type: "numeric" });
        var cell13 = worksheet.cell(2, 1, { type: 'formula', formula: 'SUM(A7:A12)' });
        var cell14 = worksheet.cell(2, 2, { type: 'formula', formula: 'SUM(A7:A12)', value: 6 });
        var cell15 = worksheet.cell(2, 3).set({ type: 'formula', formula: 'SUM(A7:A12)' });
        var cell16 = worksheet.cell(2, 4).set({ type: 'formula', formula: 'SUM(A7:A12)', value: 6 });
        var cell17 = worksheet.cell(2, 5, { type: 'sharedString', value: 'DummyString' });
        var cell18 = worksheet.cell(2, 6).set({ type: 'sharedString', value: 'DummyString' });

        // ASSERT
        expect(cell1.value).toBe('DummyValue');
        expect(cell2.value).toBe('DummyValue');
        expect(cell3.value).toBe('DummyValue');
        expect(cell4.value).toBe('DummyValue');
        expect(cell5.value).toBe('DummyValue');
        expect(cell6.value).toBe('DummyValue');
        expect(cell7.value).toBe(1);
        expect(cell8.value).toBe(1);
        expect(cell9.value).toBe(1);
        expect(cell10.value).toBe(1);
        expect(cell11.value).toBe(1);
        expect(cell12.value).toBe(1);
        expect(cell13.value).toBeUndefined();
        expect(cell14.value).toBe(6);
        expect(cell15.value).toBeUndefined();
        expect(cell16.value).toBe(6);
        expect(cell17.value).toBe(0);
        expect(cell18.value).toBe(0);
    });
});