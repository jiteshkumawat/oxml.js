define(['oxml_rels', 'oxml_sheet', 'oxml_xlsx_styles', 'xmlContentString', 'contentString', 'contentFile'],
    function (
        OxmlRels,
        oxmlSheet,
        oxmlXlsxStyles,
        XMLContentString,
        ContentString,
        ContentFile) {
        'use strict';

        var Workbook = function (xlsxContentTypes, xlsxRels) {
            this.className = "Workbook";
            this.fileName = "workbook.xml";
            this.folderName = "workbook";
            this._workBook = {
                sheets: [],
                xlsxContentTypes: xlsxContentTypes,
                xlsxRels: xlsxRels,
                _rels: OxmlRels.createRelation('workbook.xml.rels', "workbook/_rels")
            };
        };
        Workbook.prototype = Object.create(ContentFile.prototype);

        var generateSharedStrings = function (_workBook) {
            if (_workBook._sharedStrings && Object.keys(_workBook._sharedStrings).length) {
                var template = new XMLContentString({
                    rootNode: "sst",
                    nameSpaces: ["http://schemas.openxmlformats.org/spreadsheetml/2006/main"]
                });
                var sharedStrings = '';
                for (var key in _workBook._sharedStrings) {
                    sharedStrings += '<si><t>' + key.replace(/&/g, '&amp;') + '</t></si>';
                }
                return template.format(sharedStrings);
            }
            return '';
        };

        Workbook.prototype.generateContent = function (file) {
            // Create Workbood
            var template = new XMLContentString({
                rootNode: 'workbook',
                nameSpaces: [
                    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    {value: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', attribute: 'r'}]
            });
            var index, workSheets = '<sheets>';
            for (index = 0; index < this._workBook.sheets.length; index++) {
                var worksheetTemplate = new ContentString('<sheet name="{0}" sheetId="{1}" r:id="{2}"/>');
                workSheets += worksheetTemplate.format(this._workBook.sheets[index]._sheet.sheetName, this._workBook.sheets[index]._sheet.sheetId, this._workBook.sheets[index]._sheet.rId);
                if (file) {
                    this._workBook.sheets[index].attach(file);
                }
            }
            workSheets += '</sheets>';
            return template.format(workSheets);
        };

        Workbook.prototype.getLastSheet = function () {
            if (this._workBook._rels.relations.length) {
                return this._workBook._rels.relations[this._workBook._rels.relations.length - 1];
            }
            return {};
        };

        Workbook.prototype.addSheet = function (sheetName) {
            var lastSheetRel = this.getLastSheet();
            var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
            sheetName = sheetName || "sheet" + nextSheetRelId;
            this._workBook._rels.addRelation(
                "rId" + nextSheetRelId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "sheets/sheet" + nextSheetRelId + ".xml");
            // Update Sheets
            var sheet = oxmlSheet.createSheet(sheetName, nextSheetRelId, "rId" + nextSheetRelId, this, this._workBook.xlsxContentTypes);
            this._workBook.sheets.push(sheet);
            // Add Content Type
            this._workBook.xlsxContentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", {
                PartName: "/workbook/sheets/sheet" + nextSheetRelId + ".xml"
            });

            return sheet;
        };

        Workbook.prototype.attachChild = function (file) {
            this._workBook._rels.attach(file);
            var sharedStrings = generateSharedStrings(this._workBook);
            if (sharedStrings) {
                file.addFile(sharedStrings, "sharedstrings.xml", "workbook");
            }
            if (this._workBook._styles) {
                this._workBook._styles.attach(file);
            }
        }

        Workbook.prototype.createSharedString = function (str) {
            if (!str) {
                return;
            }
            if (!this._workBook._sharedStrings) {
                this._workBook._sharedStrings = {};

                // Add Content Types
                this._workBook.xlsxContentTypes.addContentType("Override", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", {
                    PartName: "/workbook/sharedstrings.xml"
                });

                // Add REL
                var lastSheetRel = this.getLastSheet();
                var nextSheetRelId = parseInt((lastSheetRel.Id || "rId0").replace("rId", ""), 10) + 1;
                this._workBook._rels.addRelation("rId" + nextSheetRelId,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                    "sharedstrings.xml");
            }

            if (isNaN(this._workBook._sharedStrings[str])) {
                this._workBook._sharedStrings[str] = Object.keys(this._workBook._sharedStrings).length || 0;
            }

            return this._workBook._sharedStrings[str];
        };

        Workbook.prototype.getSharedString = function (str) { return this._workBook._sharedStrings[str]; };

        Workbook.prototype.createStyles = function () {
            if (!this._workBook._styles) this._workBook._styles = oxmlXlsxStyles.createStyle(this, this._workBook._rels, this._workBook.xlsxContentTypes);
            return this._workBook._styles;
        }

        return {
            createWorkbook: function (xlsxContentTypes, xlsxRels) {
                return new Workbook(xlsxContentTypes, xlsxRels);
            }
        };
    });