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
        var table = worksheet.table('Table1', 'A1','C3');

        //Assert
        expect(table).toBeUndefined();
    });

    it('defines table with from cell and end cell', () => {
        // ACT
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3');

        //Assert
        expect(table).toBeDefined();
        expect(table.name).toBe('Table1');
        expect(table.fromCell).toBe('A1');
        expect(table.toCell).toBe('C3');
        expect(table.columns).toMatch(['Title1', 'Title2', 'Title3']);
    });

    it('defines blank column for row values not defined', () => {
        // ACT
        worksheet.grid(1,1,[
            ['Title1', 'Title2'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3');

        //Assert
        expect(table).toBeDefined();
        expect(table.name).toBe('Table1');
        expect(table.fromCell).toBe('A1');
        expect(table.toCell).toBe('C3');
        expect(table.columns).toMatch(['Title1', 'Title2', '']);
    });

    it('defines table with sorting (numeric value)', () => {
        // ACT
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3').set({
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3').set({
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
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
        worksheet.grid(1,1,[
            ['Title1', 'Title2', 'Title3'],
            [1,2,3],
            [2,5,3]
        ]);
        var table = worksheet.table('Table1', 'A1','C3', {
            filters: [{
                value: 1
            }]
        });

        // ASSERT
        expect(table.filters.length).toBe(0);
    });
});