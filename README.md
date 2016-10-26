# sheet-datatables
Parse OpenXML Spreadsheets files to DataTable objects and vice versa.

## Dependencies
* WindowsBase.dll
* DocumentFormat.OpenXml.dll

## Usage
### Read DataTable to local .xlsx file
```
var dataTable = new DataTable();

var spreadsheetMLParser = new DataTablesSML.SpreadsheetMLParser();
//Returns the generated file's path.
string filePath = spreadsheetMLParser.ExportSpreadsheet(dataTable, @"C:\Temp", "Sheet1");
```

### Read DataTable to MemoryStream
Useful when working with ASP.NET

```
string filename = DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";

Response.Clear();
Response.Buffer = true;
Response.AddHeader("content-disposition", "attachment;filename=" + filename);
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
