const oxml = require('../../../dist/oxml.js');
describe('border style', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('default border defined at first index for all cells', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<borders count="1"><border /></borders>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined with default values', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                color: '000000',
                style: 'thick'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left style="thick"><color rgb="000000"/></left><right style="thick"><color rgb="000000"/></right><top style="thick"><color rgb="000000"/></top><bottom style="thick"><color rgb="000000"/></bottom><diagonal style="thick"><color rgb="000000"/></diagonal></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined for left value', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                left: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left style="thick"><color rgb="000000"/></left><right/><top/><bottom/><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined for right value', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                right: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right style="thick"><color rgb="000000"/></right><top/><bottom/><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined for top value', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                top: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top style="thick"><color rgb="000000"/></top><bottom/><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined for bottom value', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top/><bottom style="thick"><color rgb="000000"/></bottom><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined for diagonal value', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                diagonal: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top/><bottom/><diagonal style="thick"><color rgb="000000"/></diagonal></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border reused among different cells', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });
        worksheet.cell(1, 2).style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top/><bottom style="thick"><color rgb="000000"/></bottom><diagonal/></border>');
                var indexBorderCount = data.indexOf('<borders count="2">');
                expect(indexBorder).toBeGreaterThan(-1);
                expect(indexBorderCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border merged for two directions', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });
        cell.style({
            border: {
                top: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top style="thick"><color rgb="000000"/></top><bottom style="thick"><color rgb="000000"/></bottom><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined with default values for collection of cells', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            border: {
                color: '000000',
                style: 'thick'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left style="thick"><color rgb="000000"/></left><right style="thick"><color rgb="000000"/></right><top style="thick"><color rgb="000000"/></top><bottom style="thick"><color rgb="000000"/></bottom><diagonal style="thick"><color rgb="000000"/></diagonal></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border merged for two directions with row', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });
        row.style({
            border: {
                top: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top style="thick"><color rgb="000000"/></top><bottom style="thick"><color rgb="000000"/></bottom><diagonal/></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border defined with default values using row', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            border: {
                color: '000000',
                style: 'thick'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left style="thick"><color rgb="000000"/></left><right style="thick"><color rgb="000000"/></right><top style="thick"><color rgb="000000"/></top><bottom style="thick"><color rgb="000000"/></bottom><diagonal style="thick"><color rgb="000000"/></diagonal></border>');
                expect(indexBorder).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('border reused among different rows', (done) => {
        // Act
        worksheet.row(1, 1, [1, 2, 3]).style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });
        worksheet.row(2, 1, [1, 2, 3]).style({
            border: {
                bottom: {
                    color: '000000',
                    style: 'thick'
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexBorder = data.indexOf('<border><left/><right/><top/><bottom style="thick"><color rgb="000000"/></bottom><diagonal/></border>');
                var indexBorderCount = data.indexOf('<borders count="2">');
                expect(indexBorder).toBeGreaterThan(-1);
                expect(indexBorderCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});