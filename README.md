# oxml.js
### Javascript library for creating and manipulating Open XML Documents like docx, xlsx, etc. <br/>
User can export grid data or images to open xml documents using this library.<br/>

## Dependencies

* <a href="https://stuk.github.io/jszip/">JSZip</a>
* <a href="https://github.com/eligrey/FileSaver.js/">File Saver</a> - But not a required. Can work on modern browsers without File Saver. Please see browser supported section for more details.

## Commands

* <b>oxml.createXLSX()</b> <br/>
Creates an open xml document in .xlsx format. This document can be used to store worksheets and charts. A convinient way to export grids and charts from html.<br/>
```
Syntax: oxml.createXLSX()
Example: var workbook = oxml.createXLSX();
Returns: XLSX document object
```
* <b>addSheet()</b> <br/>
This method is present on xlsx object obtained from createXLSX method. It facilitates adding sheets into a workbook for storing data.<br/>
```
Syntax: <workbook>.addSheet(<sheetName>)
Example: workbook.addSheet('Sheet1');
Arguments: sheetName
```
* <b>download()</b><br/>
This method is present on xlsx object obtained from createXLSX method. It facilitates downloading created workbook.<br/>
```
Syntax: <workbook>.download(<fileName>)
Example: workbook.download('workbook1.xlsx');
Arguments: fileName
```

-- Work In Progress --<br/>
I welcome contributers to help me build this library.