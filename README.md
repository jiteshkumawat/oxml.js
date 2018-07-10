# oxml.js
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Open Source Love](https://badges.frapsoft.com/os/v1/open-source.svg?v=102)](https://github.com/ellerbrock/open-source-badge/)
## Javascript library for creating and manipulating Open XML Documents like docx, xlsx, etc.
This is a small javascript library, which can let you download / export data in open xml documents like docx and xlsx which can be opened in any desktop or online document processing application such as Microsoft Excel, Google Sheets, etc. User can export data from charts or grid without making any server calls and thus saving processing units.

Open XML documents are just a ziped collection of XML files, thus only library required for oxml.js to work is JSZip, for compressing data in zip format and downloading.

## Dependencies

* [JSZIP](https://stuk.github.io/jszip/)
* [File Saver](https://github.com/eligrey/FileSaver.js/)
  *- But not required. oxml.js can work on modern browsers without File Saver. Please see browser supported section for more details.*

## Installation

Oxml.js can be installed by referring development js file, or minified js file in dist directory. This small library (~30KB) is based on javascript and require only dependency JSZip for creating compressed files. File Saver can help in achieving browser compatibility, however is not a required library.

## Usage

Simple refer scripts file from dist folder and start working on workbook document.

```html
<script src="scripts/oxml.min.js"></script>
<script>
  var workbook = oxml.xlsx();
  var worksheet = oxml.sheet('Sheet1');
  worksheet.cell(1,1, 'First Workbook');
  workbook.download('workbook.xlsx');
</script>
```

## Documentation

Oxml.js is simple to use and does not require background knowledge of Open XML. Refer [openxml.js examples](https://jiteshkumawat.github.io/oxml.js-examples/index.html)
for detailed documentation and example of all the methods available in oxml.js.
For more detailed examples on open-xlsx.js, refer examples repository [openxml.js examples](https://jiteshkumawat.github.io/oxml.js-examples/index.html).

## Builds

Oxml.js comes in two different build files, i.e., oxml.min.js and oxml.js. They are available in distributions (dist) folder. This project is configured to build with require.js and almond.js, files for the same are kept in build folder. Use build-full.js and build-min.js configuration files with r.js for generating full and minified files respectively.

## License
This project is licensed under MIT license. See [LICENSE](https://github.com/jiteshkumawat/oxml.js/blob/master/LICENSE) for more details.
