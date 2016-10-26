using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using static DataTablesSML.SMLHelper;

namespace DataTablesSML
{
    /// <summary>
    /// OpenXML Spreadsheet parser. 
    /// </summary>
    public class SpreadsheetMLParser
    {
        /// <summary>
        /// DataTable that's being manipulated
        /// </summary>
        public DataTable CurrentDataTable { get; private set; }

        /// <summary>
        /// File Path
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Current WorkbookPart
        /// </summary>
        public WorkbookPart WorkbookPart { get; private set; }

        /// <summary>
        /// Current WorksheetPart
        /// </summary>
        public WorksheetPart WorksheetPart { get; private set; }

        /// <summary>
        /// Exports DataTable to an OpenXML Spreadsheet
        /// </summary>
        /// <param name="dataTable">DataTable that will be parsed</param>
        /// <param name="destFolder">Destination folder</param>
        /// <param name="workSheetName">Name of sheet</param>
        /// <returns>Returns the path of the generated .xlsx file</returns>
        public string ExportSpreadsheet(DataTable dataTable, string destFolder, string workSheetName = "Sheet 01")
        {
            CurrentDataTable = dataTable;

            string fileName = string.Format("{0:yyyyMMdd_HHmmss}", DateTime.Now);

            if (!destFolder.EndsWith("\\"))
                destFolder += '\\';

            FilePath = $"{destFolder}{fileName}.xlsx";
            var spreadsheetDocument = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart = spreadsheetDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            var sheetData = new SheetData();

            WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet(sheetData);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(WorksheetPart), SheetId = 1, Name = "LançamentoValores" };
            sheets.Append(sheet);

            var columns = WriteSheetHeader(sheetData);
            WriteSheetBody(sheetData, columns);

            WorksheetPart.Worksheet.Save();
            WorkbookPart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
            spreadsheetDocument.Dispose();

            return FilePath;
        }

        public void ExportSpreadsheet(DataTable dataTable, ref MemoryStream stream, string workSheetName = "Sheet 01")
        {
            CurrentDataTable = dataTable;

            string fileName = string.Format("{0:yyyyMMdd_HHmmss}", DateTime.Now);

            FilePath = nameof(MemoryStream);
            
            var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart = spreadsheetDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            var sheetData = new SheetData();

            WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet(sheetData);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(WorksheetPart), SheetId = 1, Name = "LançamentoValores" };
            sheets.Append(sheet);

            var columns = WriteSheetHeader(sheetData);
            WriteSheetBody(sheetData, columns);

            WorksheetPart.Worksheet.Save();
            WorkbookPart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
            spreadsheetDocument.Dispose();

            
        }

        /// <summary>
        /// Imports spreadsheet and returns a DataTable
        /// </summary>
        /// <param name="filePathOrStream">Stream object or string of file to be parsed to a DataTable</param>
        /// <param name="hasHeaders">If true, the first line will be the columns</param>
        /// <returns>Returns a DataTable filled with spreadsheet values.</returns>
        public DataTable ImportSpreadsheet(object filePathOrStream, bool hasHeaders)
        {
            var dataTable = new DataTable();

            SpreadsheetDocument document;

            if (filePathOrStream is Stream)
                document = SpreadsheetDocument.Open(filePathOrStream as Stream, false);
            else if (!string.IsNullOrEmpty(filePathOrStream?.ToString()))
            {
                string path = filePathOrStream.ToString();
                document = SpreadsheetDocument.Open(Path.GetFullPath(path), false);
                FilePath = path;                
            }
            else
                throw new ArgumentException("Argument must be Stream or Path string.");


            this.WorkbookPart = document.WorkbookPart;
            Sheet sheet = WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
            this.WorksheetPart = this.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;            
            Worksheet worksheet = this.WorksheetPart.Worksheet;
            
            var rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

            var headers = new List<string>();
            int count = 0;
            foreach (Row row in rows)
            {
                count++;
                //Read the first row as header
                if (count == 1)
                {
                    var j = 1;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        var colunmName = hasHeaders ? GetCellValue(document, cell) : "Field" + j++;
                        Console.WriteLine(colunmName);
                        headers.Add(colunmName);
                        dataTable.Columns.Add(colunmName);
                    }
                }
                else
                {
                    dataTable.Rows.Add();
                    int i = 0;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1][i] = GetCellValue(document, cell);
                        i++;
                    }
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Writes header's data into the spreadsheet
        /// </summary>
        /// <param name="sheetData">OpenXML Sheetdata object</param>
        /// <returns>Return a list of string with the name of the columns</returns>
        private List<string> WriteSheetHeader(SheetData sheetData)
        {
            Row headerRow = new Row();

            List<string> columns = new List<string>();
            foreach (DataColumn dscolumn in CurrentDataTable.Columns)
            {
                columns.Add(dscolumn.ColumnName);

                Cell cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(dscolumn.ColumnName);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);

            return columns;
        }

        /// <summary>
        /// Writes the datatable's content into the spreadsheet
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="columns"></param>
        private void WriteSheetBody(SheetData sheetData, List<string> columns)
        {
            foreach (DataRow dsrow in CurrentDataTable.Rows)
            {
                Row newRow = new Row();
                foreach (string col in columns)
                {
                    Cell cell = new Cell();

                    if (dsrow[col] is double)
                    {
                        double cellValue = double.Parse(dsrow[col].ToString());

                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(cellValue.ToString(System.Globalization.CultureInfo.InvariantC‌​ulture));
                    }
                    else if (dsrow[col] is int)
                    {
                        int cellValue = int.Parse(dsrow[col].ToString());

                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(cellValue.ToString(System.Globalization.CultureInfo.InvariantC‌​ulture));
                    }
                    else
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                    }

                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }

        }
    }
}
