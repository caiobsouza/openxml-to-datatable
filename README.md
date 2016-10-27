# sheet-datatables
Converts OpenXML Spreadsheets files to DataTable objects and vice versa.

## Dependencies
* WindowsBase.dll
* DocumentFormat.OpenXml.dll

## Usage
### Read DataTable to local .xlsx file
```
var dataTable = new DataTable();

var spreadsheetMLParser = new DataTablesSML.SpreadsheetMLParser();
//Returns the generated file's path.
string filePath = spreadsheetMLParser.ExportSpreadsheet(dataTable, @"C:\Temp");
```

#### Optional #1: Set file name. If you don't (or pass null), the file name will be DateTime.Now
```
string filePath = spreadsheetMLParser.ExportSpreadsheet(dataTable, @"C:\Temp", "mySheet");
```
#### Optional #2: Set sheet name
```
string filePath = spreadsheetMLParser.ExportSpreadsheet(dataTable, @"C:\Temp", null, "Sheet 01");
```

#### Complete Usage
```
string filePath = spreadsheetMLParser.ExportSpreadsheet(dataTable, @"C:\Temp", "MySheetFile", "Sheet 01");
```

### Read DataTable to MemoryStream
Useful when working with ASP.NET

```
string filename = string.Format("{0:yyyyMMddhhmmss}.xlsx");

Response.Clear();
Response.Buffer = true;
Response.AddHeader("content-disposition", $"attachment;filename={filename}");
Response.ContentEncoding = Encoding.GetEncoding("Windows-1252");
Response.Charset = "UTF-8";
Response.ContentType = "application/vnd.ms-excel";

MemoryStream stream = new MemoryStream();

var spreadsheetMLParser = new DataTablesSML.SpreadsheetMLParser();
spreadsheetMLParser.ExportSpreadsheet(dtCsv, ref stream, "Sheet1");

stream.WriteTo(Response.OutputStream);
stream.Dispose();

Response.Flush();
Response.End();
```

### Read OpenXML Spreadsheet file (xslx) to DataTable

```
 //<asp:FileUpload ID="fileUploadControl" ... />
 Stream fileStream = new MemoryStream(fileUploadControl.FileBytes);
 
 //true parameter indicates if spreadsheet has headers at first row
 DataTable xlsDataTable = new DataTablesSML.SpreadsheetMLParser().ImportSpreadsheet(fileStream, true);
```

### More information
#### About OpenXML SDK
* MSDN Article: https://msdn.microsoft.com/en-us/library/office/bb448854.aspx
* Download: https://www.microsoft.com/en-us/download/details.aspx?id=30425

