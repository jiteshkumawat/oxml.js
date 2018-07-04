/* eslint-disable */
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
        oxml_xlsx_styles: 'styles/oxml-xlsx-styles',
        oxml_xlsx_font: 'styles/oxml-xlsx-font',
        oxml_xlsx_fill: 'styles/oxml-xlsx-fill',
        oxml_xlsx_border: 'styles/oxml-xlsx-border',
        oxml_xlsx_num_format: 'styles/oxml-xlsx-num-format',
        utils: 'utils/utils'
    }
});

requirejs(['fileSaver', 'oxml-xlsx'],
    function (fileSaver, oxmlXLSX) {

        // // Example 1
        // (function () {
        //     var workbook = oxml.createXLSX();
        //     var worksheet = workbook.addSheet('sheet1');
        //     var normalStyle = {
        //         fontName: 'Arial',
        //         fontSize: 10,
        //         numberFormat: '#,##0.00'
        //     };
        //     var headerStyle = {
        //         fill: {
        //             pattern: 'solid',
        //             foreColor: '090A0C'
        //         },
        //         fontColor: 'ffffff',
        //         bold: true,
        //         border: {
        //             color: 'ffff00',
        //             style: 'double',
        //             right: {
        //                 style: 'none'
        //             },
        //             left: {
        //                 style: 'none'
        //             }
        //         }
        //     };
        //     var totalCellsStyle = {
        //         fill: {
        //             pattern: 'solid',
        //             foreColor: '090A0C'
        //         },
        //         fontColor: 'ffffff',
        //         border: {
        //             color: 'ffff00',
        //             style: 'double',
        //             top: {
        //                 style: 'none'
        //             },
        //             bottom: {
        //                 style: 'none'
        //             }
        //         }
        //     }
        //     worksheet.updateValueInCell(null, 1, 1);
        //     worksheet.updateValueInCell({ type: 'string', value: 'Data 1' }, 1, 2, headerStyle);
        //     worksheet.updateValueInCell('Data 2', 1, 3, headerStyle);
        //     worksheet.updateValueInCell({ type: 'sharedString', value: 'Total' }, 1, 4, headerStyle);
        //     worksheet.updateValueInCell(undefined, 2, 1);
        //     worksheet.updateValueInCell({ type: "numeric", value: 5 }, 2, 2, normalStyle);
        //     worksheet.updateValueInCell(9, 2, 3, normalStyle);
        //     worksheet.updateValueInCell({ type: 'formula', formula: '(B2 + C2)', value: 14 }, 2, 4, normalStyle);
        //     worksheet.updateValueInCell(null, 3, 1);
        //     worksheet.updateValueInCell(7, 3, 2, normalStyle);
        //     worksheet.updateValueInCell(3, 3, 3, normalStyle);
        //     worksheet.updateValueInCell({ type: 'formula', formula: '(B3 + C3)', value: 10 }, 3, 4, normalStyle);
        //     worksheet.updateValueInCell({ type: 'sharedString', value: 'Total' }, 4, 1, {
        //         fill: {
        //             pattern: 'solid',
        //             foreColor: '090A0C'
        //         },
        //         fontColor: 'ffffff',
        //         bold: true,
        //         border: {
        //             color: 'ffff00',
        //             style: 'double',
        //             right: {
        //                 style: 'none'
        //             }
        //         }
        //     });
        //     worksheet.updateSharedFormula('(B2 + B3)', 'B4', 'D4');
        //     var totalCells = worksheet.getRow(4, 2, 3);
        //     totalCells.style({
        //         fill: {
        //             pattern: 'solid',
        //             foreColor: '090A0C'
        //         },
        //         fontColor: 'ffffff',
        //         border: {
        //             color: 'ffff00',
        //             style: 'double',
        //             right: {
        //                 style: 'none'
        //             },
        //             left: {
        //                 style: 'none'
        //             }
        //         },
        //         numberFormat: '#,##0.00'
        //     });
        //     worksheet.getCell(4, 4)
        //         .style({
        //             border: {
        //                 right: {
        //                     color: 'ffff00',
        //                     style: 'double'
        //                 },
        //                 top: {
        //                     style: 'none'
        //                 }
        //             }
        //         });
        //     worksheet.getRange(2, 4, 2, 1)
        //         .style(totalCellsStyle);
        //     worksheet.getCell(1, 2)
        //         .style({
        //             border: {
        //                 left: {
        //                     color: 'ffff00',
        //                     style: 'double'
        //                 }
        //             }
        //         });
        //     worksheet.getCell(1, 4)
        //         .style({
        //             border: {
        //                 right: {
        //                     color: 'ffff00',
        //                     style: 'double'
        //                 },
        //                 bottom: {
        //                     style: 'none'
        //                 }
        //             }
        //         });
        //     workbook.download('workbook.xlsx');
        // })();

        // Example 2
        (function () {
            var workbook = oxml.createXLSX();
            var worksheet = workbook.addSheet('sheet1');
            worksheet
					.getRange(2, 2, 3, 1)
					.style({
						fill: {
							pattern: 'solid',
							foreColor: '090A0C'
						}
					});
            workbook.download('workbook.xlsx');
        })();
    });