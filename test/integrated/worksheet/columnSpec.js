const oxml = require('../../../dist/oxml.js');
describe('column method', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('creates new instance of column with values', () => {
        // ACT
        var column = worksheet.column(1, 1, ['Some', 'Dummy', 'Value']);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Value');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
    });

    it('updates value and return updated with set method', () => {
        // ACT
        var column = worksheet.column(1, 1, [1, 2, 3]).set(['Some', 'Dummy', 'Value']);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Value');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
    });

    it('updates value of instance with set method', () => {
        // ACT
        var column = worksheet.column(1, 1, [1, 2, 3]);
        column.set(['Some', 'Dummy', 'Value']);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Value');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
    });

    it('do not change collection of cells with set', () => {
        // ACT
        var column = worksheet.column(1, 1, [1, 2, 3]);
        column.set(['Some']);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe(2);
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('numeric');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe(3);
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('numeric');
    });

    it('do not update cells out of range', () => {
        // ACT
        var column = worksheet.column(1, 1, [1, 2, 3]);
        column.set(['Some', 'Dummy', 'Text', 'Updated']);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Text');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
        expect(column.cells.length).toBe(3);
        expect(worksheet.cell(4, 1).value).toBeFalsy();
    });

    it('returns all the values starting from rowIndex, columnIndex in a column', () => {
        // ACT
        worksheet.column(1, 1, ['Some', 'Dummy', 'Text']);
        column = worksheet.column(2, 1);

        // ASSERT
        expect(column.cellIndices).toMatch(['A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(2);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Dummy');
        expect(column.cells[0].cellIndex).toBe('A2');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(3);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Text');
        expect(column.cells[1].cellIndex).toBe('A3');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells.length).toBe(2);
    });

    it('set method requires value to be defined', () => {
        // ACT
        var column = worksheet.column(1, 1, ['Some', 'Dummy', 'Text']);
        column.set();
        column.set([]);

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Some');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Text');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
        expect(column.cells.length).toBe(3);
        expect(worksheet.cell(4, 1).value).toBeFalsy();
    });

    it('update single cell using cells array', () => {
        // ACT
        var column = worksheet.column(1, 1, ['Some', 'Dummy', 'Text']);
        column.cells[0].set('Updated');

        // ASSERT
        expect(column.cellIndices).toMatch(['A1', 'A2', 'A3']);
        expect(column.cells[0].rowIndex).toBe(1);
        expect(column.cells[0].columnIndex).toBe('A');
        expect(column.cells[0].value).toBe('Updated');
        expect(column.cells[0].cellIndex).toBe('A1');
        expect(column.cells[0].type).toBe('string');
        expect(column.cells[1].rowIndex).toBe(2);
        expect(column.cells[1].columnIndex).toBe('A');
        expect(column.cells[1].value).toBe('Dummy');
        expect(column.cells[1].cellIndex).toBe('A2');
        expect(column.cells[1].type).toBe('string');
        expect(column.cells[2].rowIndex).toBe(3);
        expect(column.cells[2].columnIndex).toBe('A');
        expect(column.cells[2].value).toBe('Text');
        expect(column.cells[2].cellIndex).toBe('A3');
        expect(column.cells[2].type).toBe('string');
        expect(column.cells.length).toBe(3);
        expect(worksheet.cell(4, 1).value).toBeFalsy();
    });

    it('do not return any cell, if not defined starting from rowIndex, columnIndex in a column', () => {
        // ACT
        var column = worksheet.column(1, 2);

        // ASSERT
        expect(column.cells.length).toBe(0);
        expect(column.cellIndices.length).toBe(0);
    });

    it('style with optional parameter', function (done) {
        worksheet.column(2, 1, [1, 2, 3], {
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                var index2 = data.indexOf('<c r="A3" s="1"><v>2</v></c>');
                expect(index2).toBeGreaterThan(-1);
                var index3 = data.indexOf('<c r="A4" s="1"><v>3</v></c>');
                expect(index3).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('style with style method', function (done) {
        worksheet.column(2, 1, [1, 2, 3]).style({
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                var index2 = data.indexOf('<c r="A3" s="1"><v>2</v></c>');
                expect(index2).toBeGreaterThan(-1);
                var index3 = data.indexOf('<c r="A4" s="1"><v>3</v></c>');
                expect(index3).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});