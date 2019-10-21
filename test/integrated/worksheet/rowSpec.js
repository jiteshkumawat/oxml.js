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
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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

    it('do not define without row index', () => {
        // ACT
        var row = worksheet.row(null,1,['Some','Dummy','Value']);

        // ASSERT
        expect(row).toBe(undefined);
    });

    it('do not define without column index', () => {
        // ACT
        var row = worksheet.row(1,null,['Some','Dummy','Value']);

        // ASSERT
        expect(row).toBe(undefined);
    });

    it('do not define without numeric row index', () => {
        // ACT
        var row = worksheet.row("1",1,['Some','Dummy','Value']);

        // ASSERT
        expect(row).toBe(undefined);
    });

    it('do not define without column index', () => {
        // ACT
        var row = worksheet.row(1,"1",['Some','Dummy','Value']);

        // ASSERT
        expect(row).toBe(undefined);
    });

    it('creates new instance of row with single Value', () => {
        // ACT
        var row = worksheet.row(1,1,5);

        // ASSERT
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe(5);
        expect(row.cells[0].cellIndex).toBe('A1');
        expect(row.cells[0].type).toBe('numeric');
    });

    it('updates value and return updated with set method', () => {
        // ACT
        var row = worksheet.row(1,1,[1,2,3]).set(['Some','Dummy','Value']);

        // ASSERT
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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
        var row = worksheet.row(1,2);

        // ASSERT
        expect(row.cellIndices[0]).toBe('B1');
        expect(row.cellIndices[1]).toBe('C1');
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

    it('set method requires values attribute', () => {
        // ACT
        var row = worksheet.row(1,1,['Some', 'Dummy', 'Text']);
        row.set();
        row.set([]);

        // ASSERT
        expect(row.cellIndices[0]).toBe('A1');
        expect(row.cellIndices[1]).toBe('B1');
        expect(row.cellIndices[2]).toBe('C1');
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
    });

    it('update single value using cells array', () => {
        // ACT
        var row = worksheet.row(1,1,['Some', 'Dummy', 'Text']);
        row.cells[0].set('Updated');

        // ASSERT
        expect(row.cellIndices[0]).toMatch('A1');
        expect(row.cellIndices[1]).toMatch('B1');
        expect(row.cellIndices[2]).toMatch('C1');
        expect(row.cells[0].rowIndex).toBe(1);
        expect(row.cells[0].columnIndex).toBe('A');
        expect(row.cells[0].value).toBe('Updated');
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
    });

    it('do not return any cell, if not defined starting from rowIndex, columnIndex in a row', () => {
        // ACT
        var row = worksheet.row(1,2);

        // ASSERT
        expect(row.cells.length).toBe(0);
        expect(row.cellIndices.length).toBe(0);
    });
    
    it('style with optional parameter', function (done) {
        worksheet.row(2, 1, [1, 2, 3], {
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                var index2 = data.indexOf('<c r="B2" s="1"><v>2</v></c>');
                expect(index2).toBeGreaterThan(-1);
                var index3 = data.indexOf('<c r="C2" s="1"><v>3</v></c>');
                expect(index3).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('style with style method', function (done) {
        worksheet.row(2, 1, [1, 2, 3]).style({
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                var index2 = data.indexOf('<c r="B2" s="1"><v>2</v></c>');
                expect(index2).toBeGreaterThan(-1);
                var index3 = data.indexOf('<c r="C2" s="1"><v>3</v></c>');
                expect(index3).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('style individual cell', function (done) {
        var a = worksheet.row(2, 1, [1, 2, 3]);
        a.cells[0].style({
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('style without values', function (done) {
        var a = worksheet.row(2, 1, [1]);
        worksheet.row(2, 1, null, {
            bold: true,
            italic: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/sheets/sheet1.xml").async('string').then(function (data) {
                var index1 = data.indexOf('<c r="A2" s="1"><v>1</v></c>');
                expect(index1).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});