{
    "baseUrl": "..",
    "paths": {
        "fileHandler": 'utils/fileHandler',
        "jsZip": 'lib/jsZip.min',
        "fileSaver": 'lib/fileSaver.min',
        "oxml_content_types": 'oxml-content-types',
        "oxml_rels": 'oxml-rels',
        "oxml_sheet": 'oxml-sheet',
        "oxml_workbook": 'oxml-workbook',
        "oxml_xlsx": 'oxml-xlsx'
    },
    "include": ["build/almond", "fileHandler", "oxml_rels", "oxml_content_types", "oxml_sheet", "oxml_workbook", "oxml_xlsx"],
    "exclude": ["jsZip", "fileSaver"],
    "out": "../dist/oxml.min.js",
    "wrap": {
        "startFile": "wrap.start",
        "endFile": "wrap.end"
    }
}
