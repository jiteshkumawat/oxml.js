const oxml = require('../../dist/oxml.js');
describe('table method', () => {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('table not defined without title row', () => {
        // ACT
        var table = worksheet.table('Table1', 'A1', 'C3');

        //Assert
        expect(table).toBeUndefined();
    });

    it('defines table with from cell and end cell', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3');

        //Assert
        expect(table).toBeDefined();
        expect(table.name).toBe('Table1');
        expect(table.fromCell).toBe('A1');
        expect(table.toCell).toBe('C3');
        expect(table.columns).toMatch(['Title1', 'Title2', 'Title3']);
    });

    it('defines blank column for row values not defined', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3');

        //Assert
        expect(table).toBeDefined();
        expect(table.name).toBe('Table1');
        expect(table.fromCell).toBe('A1');
        expect(table.toCell).toBe('C3');
        expect(table.columns).toMatch(['Title1', 'Title2', '']);
    });

    it('defines table with sorting (numeric value)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            sort: 1
        });

        //Assert
        expect(table.sort).toBeDefined();
        expect(table.sort.direction).toBe('ascending');
        expect(table.sort.caseSensitive).toBe(true);
        expect(table.sort.range).toBe('A2:A3');
        expect(table.sort.dataRange).toBe('A2:C3');
    });

    it('defines table with sorting (JSON value)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            sort: {
                column: 2,
                direction: 'descending',
                caseSensitive: false
            }
        });

        //Assert
        expect(table.sort).toBeDefined();
        expect(table.sort.direction).toBe('descending');
        expect(table.sort.caseSensitive).toBe(false);
        expect(table.sort.range).toBe('B2:B3');
        expect(table.sort.dataRange).toBe('A2:C3');
    });

    it('set sorting (numeric value)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3').set({
            sort: 1
        });

        //Assert
        expect(table.sort).toBeDefined();
        expect(table.sort.direction).toBe('ascending');
        expect(table.sort.caseSensitive).toBe(true);
        expect(table.sort.range).toBe('A2:A3');
        expect(table.sort.dataRange).toBe('A2:C3');
    });

    it('set sorting (JSON value)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3').set({
            sort: {
                column: 2,
                direction: 'descending',
                caseSensitive: false
            }
        });

        //Assert
        expect(table.sort).toBeDefined();
        expect(table.sort.direction).toBe('descending');
        expect(table.sort.caseSensitive).toBe(false);
        expect(table.sort.range).toBe('B2:B3');
        expect(table.sort.dataRange).toBe('A2:C3');
    });

    it('defines table with filter value (default)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 1,
                column: 1
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].type).toBe('default');
    });

    it('defines table with filter value (custom)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 1,
                column: 1,
                type: "custom",
                operator: "greaterThan",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('greaterThan');
    });

    it('defines table with custom filter value - greaterThan', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 1,
                column: 1,
                type: "custom",
                operator: "greaterThan",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('greaterThan');
    });

    it('defines table with custom filter value - greaterThanOrEqual', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 1,
                column: 1,
                type: "custom",
                operator: "greaterThanOrEqual",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('greaterThanOrEqual');
    });

    it('defines table with custom filter value - lessThan', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 2,
                column: 1,
                type: "custom",
                operator: "lessThan",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(2);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('lessThan');
    });

    it('defines table with custom filter value - lessThanOrEqual', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 2,
                column: 1,
                type: "custom",
                operator: "lessThanOrEqual",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(2);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('lessThanOrEqual');
    });

    it('defines table with custom filter value - notEqual', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 2,
                column: 1,
                type: "custom",
                operator: "notEqual",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(2);
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].operator).toBe('notEqual');
    });

    it('defines table with filter values (default)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                values: [1, 2],
                column: 1,
                type: "custom",
                operator: "equal",
                and: false
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].type).toBe('custom');
        expect(table.filters[0].values[0].and).toBe(false);
        expect(table.filters[0].values[0].operator).toBe('equal');
        expect(table.filters[0].values[1].value).toBe(2);
        expect(table.filters[0].values[1].type).toBe('custom');
        expect(table.filters[0].values[1].and).toBe(false);
        expect(table.filters[0].values[0].operator).toBe('equal');
    });

    it('defines table with filter values (custom)', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                values: [1, 2],
                column: 1
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(1);
        expect(table.filters[0].column).toBe(0);
        expect(table.filters[0].values[0].value).toBe(1);
        expect(table.filters[0].values[0].type).toBe('default');
        expect(table.filters[0].values[1].value).toBe(2);
        expect(table.filters[0].values[1].type).toBe('default');
    });

    it('filters not defined without column', () => {
        // ACT
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                value: 1
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(0);
    });

    it('styling table basic', (done) => {
        // ARRANGE
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3');

        // ACT
        table.style({
            fontSize: 12,
            firstRow: {
                bold: true,
                fontColor: 'ffffff',
                fill: {
                    pattern: 'solid',
                    color: '000000'
                }
            },
            evenRow: {
                fill: {
                    pattern: 'solid',
                    color: 'aaaaaa'
                }
            },
            oddRow: {
                fill: {
                    pattern: 'solid',
                    color: 'eeeeee'
                }
            }
        }, 'tableStyle1');

        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["workbook/tables/table1.xml"]).toBeDefined();
            zip.file("workbook/tables/table1.xml").async('string').then((data) => {
                var index = data.indexOf('<tableStyleInfo name="tableStyle1" showColumnStripes="0" showRowStripes="1" showLastColumn="0" showFirstColumn="0"/>');
                expect(index).toBeGreaterThan(-1);
            });
            zip.file('workbook/style2.xml').async('string').then(function (data) {
                var index1 = data.indexOf('<dxfs count="4"><dxf><font><sz val="12"/></font></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="aaaaaa"/></patternFill></fill></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="eeeeee"/></patternFill></fill></dxf><dxf><font><b/><color rgb="ffffff"/></font><fill><patternFill patternType="solid"><bgColor rgb="000000"/></patternFill></fill></dxf></dxfs>');
                var index2 = data.indexOf('<tableStyles count="1"><tableStyle count="4" name="tableStyle1"><tableStyleElement dxfId="0" type="wholeTable"/><tableStyleElement dxfId="1" type="secondRowStripe"/><tableStyleElement dxfId="2" type="firstRowStripe"/><tableStyleElement dxfId="3" type="headerRow"/></tableStyle></tableStyles>');
                expect(index1).toBeGreaterThan(-1);
                expect(index2).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('styling table with filter and sort', (done) => {
        // ARRANGE
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3', {
            filters: [{
                values: [1, 2],
                column: 1,
                type: "custom",
                operator: "equal",
                and: false
            }],
            sort: {
                column: 2,
                direction: 'descending',
                caseSensitive: false
            }
        });

        // ACT
        table.style({
            fontSize: 12,
            firstRow: {
                bold: true,
                fontColor: 'ffffff',
                fill: {
                    pattern: 'solid',
                    color: '000000'
                }
            },
            evenRow: {
                fill: {
                    pattern: 'solid',
                    color: 'aaaaaa'
                }
            },
            oddRow: {
                fill: {
                    pattern: 'solid',
                    color: 'eeeeee'
                }
            }
        }, 'tableStyle1');

        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["workbook/tables/table1.xml"]).toBeDefined();
            zip.file("workbook/tables/table1.xml").async('string').then((data) => {
                var index = data.indexOf('<tableStyleInfo name="tableStyle1" showColumnStripes="0" showRowStripes="1" showLastColumn="0" showFirstColumn="0"/>');
                expect(index).toBeGreaterThan(-1);
            });
            zip.file('workbook/style2.xml').async('string').then(function (data) {
                var index1 = data.indexOf('<dxfs count="4"><dxf><font><sz val="12"/></font></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="aaaaaa"/></patternFill></fill></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="eeeeee"/></patternFill></fill></dxf><dxf><font><b/><color rgb="ffffff"/></font><fill><patternFill patternType="solid"><bgColor rgb="000000"/></patternFill></fill></dxf></dxfs>');
                var index2 = data.indexOf('<tableStyles count="1"><tableStyle count="4" name="tableStyle1"><tableStyleElement dxfId="0" type="wholeTable"/><tableStyleElement dxfId="1" type="secondRowStripe"/><tableStyleElement dxfId="2" type="firstRowStripe"/><tableStyleElement dxfId="3" type="headerRow"/></tableStyle></tableStyles>');
                expect(index1).toBeGreaterThan(-1);
                expect(index2).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('styling table merge different styling region', (done) => {
        // ARRANGE
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3');

        // ACT
        table.style({
            fontSize: 12
        }).style({
            firstRow: {
                bold: true,
                fontColor: 'ffffff',
                fill: {
                    pattern: 'solid',
                    color: '000000'
                }
            }
        }).style({
            evenRow: {
                fill: {
                    pattern: 'solid',
                    color: 'aaaaaa'
                }
            }
        }).style({
            oddRow: {
                fill: {
                    pattern: 'solid',
                    color: 'eeeeee'
                }
            }
        }, 'tableStyle1');

        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["workbook/tables/table1.xml"]).toBeDefined();
            zip.file("workbook/tables/table1.xml").async('string').then((data) => {
                var index = data.indexOf('<tableStyleInfo name="tableStyle1" showColumnStripes="0" showRowStripes="1" showLastColumn="0" showFirstColumn="0"/>');
                expect(index).toBeGreaterThan(-1);
            });
            zip.file('workbook/style2.xml').async('string').then(function (data) {
                var index1 = data.indexOf('<dxfs count="4"><dxf><font><sz val="12"/></font></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="aaaaaa"/></patternFill></fill></dxf><dxf><font></font><fill><patternFill patternType="solid"><bgColor rgb="eeeeee"/></patternFill></fill></dxf><dxf><font><b/><color rgb="ffffff"/></font><fill><patternFill patternType="solid"><bgColor rgb="000000"/></patternFill></fill></dxf></dxfs>');
                var index2 = data.indexOf('<tableStyles count="1"><tableStyle count="4" name="tableStyle1"><tableStyleElement dxfId="0" type="wholeTable"/><tableStyleElement dxfId="1" type="secondRowStripe"/><tableStyleElement dxfId="2" type="firstRowStripe"/><tableStyleElement dxfId="3" type="headerRow"/></tableStyle></tableStyles>');
                expect(index1).toBeGreaterThan(-1);
                expect(index2).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('styling table merge different style in one region', (done) => {
        // ARRANGE
        worksheet.grid(1, 1, [
            ['Title1', 'Title2', 'Title3'],
            [1, 2, 3],
            [2, 5, 3]
        ]);
        var table = worksheet.table('Table1', 'A1', 'C3');

        // ACT
        table.style({
            firstRow: {
                bold: true
            }
        }).style({
            firstRow: {
                fontColor: 'ffffff',
            }
        }, 'tableStyle1').style({
            firstRow: {
                fill: {
                    pattern: 'solid',
                    color: '000000'
                }
            }
        }, 'tableStyle1');

        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            // ASSERT
            expect(zip.files["workbook/tables/table1.xml"]).toBeDefined();
            zip.file("workbook/tables/table1.xml").async('string').then((data) => {
                var index = data.indexOf('<tableStyleInfo name="tableStyle1" showColumnStripes="0" showRowStripes="0" showLastColumn="0" showFirstColumn="0"/>');
                expect(index).toBeGreaterThan(-1);
            });
            zip.file('workbook/style2.xml').async('string').then(function (data) {
                var index1 = data.indexOf('<dxfs count="2"><dxf><font></font></dxf><dxf><font><b/><color rgb="ffffff"/></font><fill><patternFill patternType="solid"><bgColor rgb="000000"/></patternFill></fill></dxf></dxfs>');
                var index2 = data.indexOf('<tableStyles count="1"><tableStyle count="2" name="tableStyle1"><tableStyleElement dxfId="0" type="wholeTable"/><tableStyleElement dxfId="1" type="headerRow"/></tableStyle></tableStyles>');
                expect(index1).toBeGreaterThan(-1);
                expect(index2).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});