requirejs.config({
    baseUrl: '..',
    paths: {
        fileHandler: 'utils/fileHandler',
        jsZip: 'lib/jsZip.min',
        fileSaver: 'lib/fileSaver.min',
        oxml_content_types: 'oxml-content-types',
        oxml_rels: 'oxml-rels',
        oxml_sheet: 'oxml-sheet',
        oxml_workbook: 'oxml-workbook',
        oxml_xlsx: 'oxml-xlsx',
        oxml_xlsx_styles: 'oxml-xlsx-styles'
    }
});

requirejs(['fileSaver', 'oxml-xlsx'],
    function (fileSaver, oxmlXLSX) {
        // var file = oxml.createXLSX();
        // var sharedString = file._xlsx.workBook.createSharedString("Hello");
        // var sheet2 = file.addSheet('sheet1');
        // sheet2.updateValuesInMatrix([
        //     [null, 'Sale Price', 'Cost Price', 'Profit', 'Profit%'],
        //     [null, 10, 14], //{type: "formula", formula: "(B2 - A2)", value: 4}, {type: "formula", formula: "(C2 / A2) * 100", value: 40}]
        //     [null, 11, 14],
        //     'Total'
        // ]);
        // sheet2.updateSharedFormula('(C2 - B2)', 'D2', 'D3');
        // sheet2.updateSharedFormula('(B2 + B3)', 'B4', 'D4');
        // sheet2.updateSharedFormula('(D2 / B2) * 100', 'E2', 'E4');
        // sheet2.updateValueInCell(40, 5, 1);
        // var sheet = file.addSheet('tst1');
        // sheet.updateValuesInRow(0, [null, 22, undefined, 23, 28]);
        // file.download('t.xlsx');
        // file.destroy();

        var workbook = oxml.createXLSX();
        var worksheet = workbook.addSheet('sheet1');
        // worksheet.updateValuesInColumn(["Dealer", "Cost Price", "Sale Price", "Profit", "Profit %"], 1, 1, {
        //     bold: true,
        //     italic: true,
        //     border: {
        //         color: "00ff00",
        //         style: "thick"
        //     },
        //     fill: {
        //         pattern: 'solid',
        //         color: 'FFFF00'
        //     }
        // });
        // worksheet.updateValueInCell('Test Bold2', 2, 1, {
        //     bold: true,
        //     italic: true,
        //     underline: true,
        //     fontName: "Arial",
        //     fontSize: 9,
        //     fontFamily: 2,
        //     fontColor: 'ff0000',
        //     strike: true,
        //     border: {
        //         left: {
        //             color: "0000ff",
        //             style: "dashDot" // dashDot,dashDotDot,dashed,dotted,double,hair,medium,mediumDashDot,mediumDashDotDot,mediumDashed,none,slantDashDot,thick,thin
        //         },
        //         bottom: {
        //             color: "0000ff",
        //             style: "double"
        //         },
        //         fill: {
        //             pattern: 'lightGrid',
        //             backColor: 'FFFF00' // none,solid,mediumGray,darkGray,lightGray,darkHorizontal,darkVertical,darkDown,darkUp,darkGrid,darkTrellis,lightHorizontal,lightVertical,lightDown,lightUp,lightGrid,lightTrellis,gray125,gray0625
        //         }
        //     }
        // });
        // worksheet.updateValueInCell(0.094, 2, 2, {
        //     bold: true,
        //     italic: true,
        //     underline: true,
        //     fontName: "Arial",
        //     fontSize: 9,
        //     fontFamily: 2,
        //     // color: 'Red',
        //     numberFormat: "$ #,##0.00;$ #,##0.00;-",
        //     fill: {
        //         gradient: {
        //             degree: 90,
        //             stop: [{
        //                 position: 0,
        //                 color: "FF92D050"
        //             }, {
        //                 position: 1,
        //                 color: "FF0070C0"
        //             }]
        //         }
        //     }
        // });

        worksheet.updateValuesInMatrix([
            [null, 'Sale Price', 'Cost Price', 'Profit', 'Profit%'],
            [null, 10, 14],
            [null, 11, 14],
            'Total'
        ], 4, 4, {
                bold: true,
                italic: true,
                underline: true,
                fontName: "Arial",
                fontSize: 9,
                fontFamily: 2,
                fontColor: 'ff0000',
                strike: true,
                border: {
                    left: {
                        color: "0000ff",
                        style: "dashDot" // dashDot,dashDotDot,dashed,dotted,double,hair,medium,mediumDashDot,mediumDashDotDot,mediumDashed,none,slantDashDot,thick,thin
                    },
                    bottom: {
                        color: "0000ff",
                        style: "double"
                    },
                    fill: {
                        pattern: 'lightGrid',
                        backColor: 'FFFF00' // none,solid,mediumGray,darkGray,lightGray,darkHorizontal,darkVertical,darkDown,darkUp,darkGrid,darkTrellis,lightHorizontal,lightVertical,lightDown,lightUp,lightGrid,lightTrellis,gray125,gray0625
                    }
                }
            });
        // worksheet.updateSharedFormula('(F5 - E5)', 'G5', 'G6');
        // worksheet.updateSharedFormula('(E5 + E6)', 'E7', 'F7');
        // worksheet.updateSharedFormula('(G5 / E5) * 100', 'H5', 'H6');
        workbook.download('workbook.xlsx');
    });