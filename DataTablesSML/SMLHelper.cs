using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace DataTablesSML
{
    /// <summary>
    /// Helper methods
    /// </summary>
    class SMLHelper
    {
        /// <summary>
        /// Alphabet used to generate column names
        /// </summary>
        public const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVXWYZ";

        /// <summary>
        /// Inserts a cell into a row
        /// </summary>
        /// <param name="row">OpenXML's Row object</param>
        /// <param name="column">Column position (from Alphabet)</param>
        /// <returns>Returns the generated Cell</returns>
        public static Cell InsertCell(Row row, string column)
        {
            
            string cellReference = $"{column}{row.RowIndex}";

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                return newCell;
            }            

        }

        /// <summary>
        /// Inserts a row into a Worksheet
        /// </summary>
        /// <param name="worksheetPart">OpenXML's WorksheetPart object</param>
        /// <returns>Returns the generated Row</returns>
        public static Row InsertRow(WorksheetPart worksheetPart)
        {
            Row created = null;

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row lastRow = sheetData.Elements<Row>().LastOrDefault();

            if (lastRow != null)
            {
                created = sheetData.InsertAfter(new Row() { RowIndex = (lastRow.RowIndex + 1) }, lastRow);
            }
            else
            {
                created = new Row() { RowIndex = 1 };
                sheetData.Append(created);
                return created;
            }

            return created;
        }

        /// <summary>
        /// Gets the cell value
        /// </summary>
        /// <param name="document">OpenXML's SpreadsheetDocument object</param>
        /// <param name="cell">OpenXML's Cell object</param>
        /// <returns>Returns the cell value as string</returns>
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
