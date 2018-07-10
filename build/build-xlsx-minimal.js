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
        "oxml_xlsx": 'oxml-xlsx',
        "oxml_xlsx_styles": 'build/helper/no_styles/oxml-xlsx-styles',
        "utils": 'utils/utils'
    },
    "include": ["build/almond", "fileHandler", "oxml_rels", "oxml_content_types", "oxml_sheet", "oxml_workbook", "oxml_xlsx"],
    "exclude": ["jsZip", "fileSaver"],
    "out": "../dist/oxml-xlsx-minimal.js",
    "optimize": "none",
    "wrap": {
        "startFile": "wrap.start",
        "endFile": "wrap.end"
    }
}
