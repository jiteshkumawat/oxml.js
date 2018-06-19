# oxml.js
## Javascript library for creating and manipulating Open XML Documents like docx, xlsx, etc.
This is a small javascript library, which can let you download / export data in open xml documents like docx and xlsx which can be opened in any desktop or online document processing application such as Microsoft Excel, Google Sheets, etc. User can export data from charts or grid without making any server calls and thus saving processing units.

Open XML documents are just a ziped collection of XML files, thus only library required for oxml.js to work is JSZip, for compressing data in zip format and downloading.

## Dependencies

* [JSZIP](https://stuk.github.io/jszip/)
* [File Saver](https://github.com/eligrey/FileSaver.js/)
  *- But not required. oxml.js can work on modern browsers without File Saver. Please see browser supported section for more details.*
