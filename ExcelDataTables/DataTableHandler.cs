using System;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using SysDataTable = System.Data.DataTable;

namespace ExcelDataTables
{
    public class DataTableHandler
    {
      
        public SysDataTable CurrentDataTable { get; private set; }
        public string FilePath { get; private set; }

        private Excel._Application application;
        private Excel._Worksheet workSheet;
        private Excel._Workbook workBook;
        private Excel.Range cellRange;

        public SysDataTable ImportSpreadsheet()
        {
            throw new NotImplementedException();
        }

        public string ExportSpreadsheet(SysDataTable dataTable, string temporaryFolder, string workSheetName = "Default")
        {
            CurrentDataTable = dataTable;

            application = new Excel.Application();
            
            application.Visible = false;
            application.DisplayAlerts = false;

            workBook = application.Workbooks.Add(Type.Missing);

            workSheet = (Excel.Worksheet)application.ActiveSheet;
            if (!string.IsNullOrEmpty(workSheetName))
                workSheet.Name = workSheetName;

            cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[CurrentDataTable.Rows.Count, CurrentDataTable.Columns.Count]];

            WriteSheetHeader();
            WriteSheetLines();

            cellRange.EntireColumn.AutoFit();

            return SaveSpreadSheet(temporaryFolder);
        }

        private string SaveSpreadSheet(string tempFolder)
        {
            string fileName = string.Format("{0:yyyyMMdd_HHmmss}", DateTime.Now);

            if (!tempFolder.EndsWith("\\"))
                tempFolder += '\\';

            string FilePath = $"{tempFolder}{fileName}";

            workBook.SaveAs(FilePath);
            workBook.Close();
            application.Quit();

            return FilePath;
        }

        private void WriteSheetLines()
        {
            foreach (DataRow row in CurrentDataTable.Rows)
            {
                int index = CurrentDataTable.Rows.IndexOf(row);

                if (index == 0)
                    continue;

                ++index;
                foreach (DataColumn column in CurrentDataTable.Columns)
                {
                    
                    int xsColumnIndex = CurrentDataTable.Columns.IndexOf(column) + 1;

                    workSheet.Cells[index, xsColumnIndex] = row[column.ColumnName];
                    if (row[column.ColumnName].GetType().Equals(typeof(double)))
                    {
                        ((Excel.Range)workSheet.Cells[index, xsColumnIndex]).NumberFormat = "#,###.00";
                    }
                }
            }
        }

        private void WriteSheetHeader()
        {
            foreach (DataColumn column in CurrentDataTable.Columns)
            {
                int xsColumnIndex = CurrentDataTable.Columns.IndexOf(column) + 1;
                workSheet.Cells[1, xsColumnIndex] = column.ColumnName;
            }
        }


    }
}
