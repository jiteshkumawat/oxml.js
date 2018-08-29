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
        "oxml_xlsx_styles": 'styles/oxml-xlsx-styles',
        "oxml_xlsx_font": 'styles/oxml-xlsx-font',
        "oxml_xlsx_fill": 'styles/oxml-xlsx-fill',
        "oxml_xlsx_border": 'styles/oxml-xlsx-border',
        "oxml_xlsx_num_format": 'styles/oxml-xlsx-num-format',
        "oxml_table": "oxml-table",
        "utils": 'utils/utils',
        "contentFile": 'base/contentFile',
        "contentString": 'base/contentString',
        "xmlContentString": 'base/xmlContentString'
    },
    "include": ["build/almond", "fileHandler", "oxml_rels", "oxml_content_types", "oxml_sheet", "oxml_workbook", "oxml_xlsx", "oxml_table"],
    "exclude": ["jsZip", "fileSaver"],
    "out": "../dist/oxml.min.js",
    "wrap": {
        "startFile": "wrap.start",
        "endFile": "wrap.end"
    }
}
