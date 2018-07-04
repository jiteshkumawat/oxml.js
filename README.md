# oxml.js
## Javascript library for creating and manipulating Open XML Documents like docx, xlsx, etc.
This is a small javascript library, which can let you download / export data in open xml documents like docx and xlsx which can be opened in any desktop or online document processing application such as Microsoft Excel, Google Sheets, etc. User can export data from charts or grid without making any server calls and thus saving processing units.

Open XML documents are just a ziped collection of XML files, thus only library required for oxml.js to work is JSZip, for compressing data in zip format and downloading.

## Dependencies

* [JSZIP](https://stuk.github.io/jszip/)
* [File Saver](https://github.com/eligrey/FileSaver.js/)
  *- But not required. oxml.js can work on modern browsers without File Saver. Please see browser supported section for more details.*

## Installation

Openxml.js can be installed by referring development js file, or minified js file in dist directory. This small library (~30KB) is based on javascript and require only dependency JSZip for creating compressed files. File Saver can help in achieving browser compatibility, however is not a required library.

## Methods

### XLSX Documents

#### Creating, removing or exporting Workbook / Worksheets

Below methods are available for creating, adding and manipulating workbook or worksheets.

##### createXLSX()

 This method lets user create a new instance of workbook. createXLSX is available in global object oxml defined after including the library. It do not require any parameter and sets default values required for any workbook document.
 ```javascript
 Syntax: oxml.createXLSX();
 Example: var workBook = oxml.createXLSX();
 ```
 
##### addSheet()
 
 This method lets user create a new instance of a worksheet inside workbook. addSheet is available to instance of workbook document. It require an only parameter of sheet name. Default sheet name if not provided by user will be "Sheet" appended with a dynamic number, which in most cases is the relation id of open xml sheet document.
 ```javascript
 Syntax: [workBook].addSheet([sheetName])
 Example:
 var workBook = oxml.createXLSX();
 var workSheet = workBook.addSheet('Sheet1');
 ```
 
##### download()
 
 This method lets user download or export created xlsx document. download is available to instance of workbook document. It requires a file name as a parameter for download.
  ```javascript
  Syntax: [workBook].download([fileName])
  Example:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workBook.download('workbook.xlsx');
  ```
  
#### Update data in Worksheets

Below methods are available for pushing data into worksheets. They use value argument to fill sheet data, this value argument can be numeric, string or a JSON object depending on following scenarios:

**_Numeric Values:_** Numeric values will update a cell with numeric data. They can be passed directly or using number type.

**_String Values:_** String values will update a cell with inline string data. They can be passed directly or using string type.

**_Shared String Values:_** Shared string are optimized way to update a string value in cell. They are shared among different cells, thus otimizing space complexity of xlsx document.

**_Formula:_** Formula type of values will perform calculations to assign value in cell. They can be written as equation for formula bar of any workbook editor like Microsoft Excel. Formula can be assigned to a cell with or without cached value.

**_Shared Formula:_** Shared Frmula are smarter type of values assigned in xlsx documents. They can not be directly assigned along with values in column, row or matrix but can be assigned seperately as described later in this section.

Below are few of the examples for creating values:
 ```javascript
 // This will create Numeric values
 var numericValues = [1,2,3,{type: "number", value: 4}];

 // This will create String values
 var stringValues = ["a","b",{type: "string", value: "c"},"d"];
 
 // This will create Shared string values
 var sharedStringValues = [{type: "sharedString", value: "Hello"}, {type: "sharedString", value: "World"}];
 
 // This will create Formula type cell with cached value. This will calculate sum of A1 and A2 cell, where 'A' defined column and numeric value define row. Value can be provided if this has to be cached.
 var formulaWithValue = [{type: "formula", value: 10, formula: "(A1 + A2)"}]
 
 // Below will create empty values
 var nullValues = [null, undefined, {type: "number"}, {value: null}];
 ```

##### updateValueInCell()

 This method is used to push data in a cell of worksheet. updateValueInCell method is available to instance of worksheet, and it takes three parameters:
 1. value: First parameter is value to be filled. User can pass value in any format except shared formula, described above.
 1. rowIndex: Second parameter passed in method is index of row starting from 1.
 1. columnIndex: Thirt parameter passed in method is index of column starting from 1.
  ```javascript
  Syntax: [workSheet].updateValueInCell([value], [rowIndex], [columnIndex])
  Example:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workSheet.updateValuesInRow(["Data"], 1, 3);
  // This will update data in following format:
  // --- --- --- Data
  ```

##### updateValuesInRow()

 This method is used to push data in consecutive cells of a row. updateValuesInRow method is available to instance of worksheet. updateValuesInRow takes three parameters:
 1. valuesArray: First parameter is an array list of values to be filled in row. User can pass value in any format except shared formula, described above.
 1. rowIndex: Second parameter passed in method is index of row starting from 1.
 1. columnIndex: Thirt parameter passed in method is index of column starting from 1.
  ```javascript
  Syntax: [workSheet].updateValuesInRow([values], [rowIndex], [columnIndex])
  Example:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workSheet.updateValuesInRow([1,2,null,4,5], 2, 1);
  // This will update data in following format:
  // -----------------------
  // 1 --- 2 --- --- 4 --- 5
  ```
  
##### updateValuesInColumn()

 This method is used to push data in consecutive cells of a column. updateValuesInColumn method is available to instance of worksheet. updateValuesInColumn takes three parameters:
  1. valuesArray: First parameter is an array list of values to be filled in column. User can pass value in any format except shared formula, described above.
  1. rowIndex: Second parameter passed in method is index of row starting from 1.
  1. columnIndex: Third parameter passed in method is index of column starting from 1.
  ```javascript
  Syntax: [workSheet].updateValuesInColumn([values], [rowIndex], [columnIndex])
  Example:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workSheet.updateValuesInColumn(['Data1',1,2,3], 1, 1);
  // This will update data in following format:
  // Data1
  // 1
  // 2
  // 3
  ```

##### updateValuesInMatrix()
 
 This method is used to push data in row and column format. updateValuesInMatrix is available to instance of worksheet. updateValuesInMatrix again takes following three arguments
  1. valuesArray: First parameter is a two dimensional array list of values to be filled in rows and column column. User can pass value in any format except shared formula, described above.
  1. rowIndex: Second parameter passed in method is index of row starting from 1.
  1. columnIndex: Third parameter passed in method is index of column starting from 1.
  ```javascript
  Syntax: [workSheet].updateValuesInMatrix([values], [rowIndex], [columnIndex])
  Example1:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workSheet.updateValuesInMatrix([
      [1, 3, {type: "sharedString", value: "Hello"}],
      ['a','b'],
      [4, {type: "sharedString", value: "Hello"}, 9]
  ], 1, 1);
  // This will update data in following format:
  // 1 ---   3   --- Hello
  // a ---   b
  // 4 --- Hello ---  9
  
  Example2:
  var workBook = oxml.createXLSX();
  var workSheet = workBook.addSheet('Sheet1');
  workSheet.updateValuesInMatrix([
      ['Sale Price', 'Cost Price', 'Profit', 'Profit%'],
      [10, 14, {type: "formula", formula: "(B2 - A2)", value: 4}, {type: "formula", formula: "(C2 / A2) * 100", value: 40}]
  ], 2, 1);
  // This will update data in following format:
  // ------------------------------------------------
  // Sale Price --- Cost Price --- Profit --- Profit%
  //     10     ---     14     ---    4   ---   40
  ```

For more detailed examples on open-xlsx.js, refer examples repository [openxml.js examples](https://jiteshkumawat.github.io/oxml.js-examples/index.html).

## License
This project is licensed under MIT license. See [LICENSE](https://github.com/jiteshkumawat/oxml.js/blob/master/LICENSE) for more details.
