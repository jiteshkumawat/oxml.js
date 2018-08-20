const oxml = require('../../../dist/oxml.js');
describe('fill style', function () {
    var workbook, worksheet;
    beforeEach(function () {
        workbook = oxml.xlsx();
        worksheet = workbook.sheet();
    });

    afterEach(function () {
        worksheet = null;
        workbook = null;
    });

    it('default fill defined at first index for all cells', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            fill: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fills count="2"><fill/><fill><patternFill patternType="gray125"></patternFill></fill></fills>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with pattern values - default', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            fill: {
                pattern: 'darkGrid',
                color: '0000ff'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><patternFill patternType="darkGrid"><bgColor rgb="0000ff"/></patternFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with pattern values with forecolor and backcolor', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            fill: {
                pattern: 'darkGrid',
                foreColor: '0000ff',
                backColor: '00ff00'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><patternFill patternType="darkGrid"><fgColor rgb="0000ff"/><bgColor rgb="00ff00"/></patternFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with gradient values - default', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with gradient values - custom', (done) => {
        // Act
        var cell = worksheet.cell(1, 1);
        cell.style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }],
                    type: 'linear',
                    left: 0.2,
                    right: 0.2,
                    top: 0.2,
                    bottom: 0.2
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill type="linear" left="0.2" right="0.2" top="0.2" bottom="0.2" degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill reused among different cells', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });
        worksheet.cell(1, 2).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                var indexFillCount = data.indexOf('<fills count="3">');
                expect(indexFill).toBeGreaterThan(-1);
                expect(indexFillCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill updated of cell', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });
        worksheet.cell(1, 2).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 60,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill1 = data.indexOf('<fill><gradientFill degree="60"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                var indexFill2 = data.indexOf('<fill><gradientFill degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                var indexFillCount = data.indexOf('<fills count="4">');
                expect(indexFill1).toBeGreaterThan(-1);
                expect(indexFill2).toBeGreaterThan(-1);
                expect(indexFillCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill not updated if used by other cell', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 60,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill degree="60"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                var indexFillCount = data.indexOf('<fills count="3">');
                expect(indexFill).toBeGreaterThan(-1);
                expect(indexFillCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill not impacted by other style updates on cell', (done) => {
        // Act
        worksheet.cell(1, 1).style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }]
                }
            }
        });
        worksheet.cell(1, 1).style({
            bold: true
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                var indexFillCount = data.indexOf('<fills count="3">');
                expect(indexFill).toBeGreaterThan(-1);
                expect(indexFillCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with pattern for collection of cells', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            fill: {
                pattern: 'solid',
                color: '0000ff'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><patternFill patternType="solid"><bgColor rgb="0000ff"/></patternFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill defined with gradient for collection of cells', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }],
                    type: 'linear',
                    left: 0.2,
                    right: 0.2,
                    top: 0.2,
                    bottom: 0.2
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill type="linear" left="0.2" right="0.2" top="0.2" bottom="0.2" degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill updated for row', (done) => {
        // Act
        var row = worksheet.row(1, 1, [1, 2, 3]);
        row.style({
            fill: {
                pattern: 'solid',
                color: '0000ff'
            }
        });
        row.style({
            fill: {
                gradient: {
                    degree: 90,
                    stops: [{ position: 0, color: "FF92D050" }, { position: 1, color: "FF0070C0" }],
                    type: 'linear',
                    left: 0.2,
                    right: 0.2,
                    top: 0.2,
                    bottom: 0.2
                }
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><gradientFill type="linear" left="0.2" right="0.2" top="0.2" bottom="0.2" degree="90"><stop position="0"><color rgb="FF92D050"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill>');
                expect(indexFill).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });

    it('fill reused among different rows', (done) => {
        // Act
        worksheet.row(1, 1, [1, 2, 3]).style({
            fill: {
                pattern: 'solid',
                color: '0000ff'
            }
        });
        worksheet.row(2, 1, [1, 2, 3]).style({
            fill: {
                pattern: 'solid',
                color: '0000ff'
            }
        });

        // Assert
        workbook.download(__dirname + '/demo.xlsx').then(function (zip) {
            expect(zip.files["workbook/style2.xml"]).toBeDefined();
            zip.file("workbook/style2.xml").async('string').then(function (data) {
                var indexFill = data.indexOf('<fill><patternFill patternType="solid"><bgColor rgb="0000ff"/></patternFill></fill>');
                var indexFillCount = data.indexOf('<fills count="3">');
                expect(indexFill).toBeGreaterThan(-1);
                expect(indexFillCount).toBeGreaterThan(-1);
                done();
            });
        }).catch(function () {
            done.fail();
        });
    });
});