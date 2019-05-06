using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;

namespace AutoTool
{
    public class InteropExcel
    {
        private Application excelApp = new Application();
        private Workbook wb = null;
        private Worksheet ws = null;
        private Range range = null;

        private object missValue;
        private Process[] excelProcs;

        private const string filePath = "D:\\Stabilization Status.xlsx";

        public void InsertNewReport(string sheetName)
        {
            CleanExcelProcess();
            LoadExcelSheet(filePath, sheetName);

            // Insert copied Range.
            range = GetRange("Columns", "F:I");
            range.Copy();
            range.Insert();

            // Update Date value.
            range = GetCell(1, 6);
            range.Value = String.Format("{0:m}", DateTime.Today);

            range = GetCell(7, 6);
            range.Value = String.Format("{0:m}", DateTime.Today);

            // Clear old data. Values below are fixed to Excel template.
            var startCellRow = 10;
            var startCellCol = 6;
            var endCellRow = 56;
            var endCellCol = 7;
            range = ws.Range[ws.Cells[startCellRow, startCellCol], ws.Cells[endCellRow, endCellCol]];

            for (int i = startCellCol; i < range.Columns.Count + startCellCol; i++)
            {
                for (int j = startCellRow; j < range.Rows.Count + startCellRow; j++)
                {
                    Range cell = ws.Cells[j, i];
                    if (i == 6)
                    {
                        cell.ClearContents();
                    }
                    else
                    {
                        // use for merged cells in template
                        cell.MergeArea.ClearContents(); 
                    }
                }
            }
            
            // Save changes then quit.
            wb.Save();
            QuitExcelApplication();
            CleanExcelProcess();
        }
        
        private Range GetRange(string rangeType, object rangeDefine)
        {
            /*
             * From cell to cell ---
             * Range range = ws.Range[ws.Cells[2, 2], ws.Cells[5, 6]];
             * 
             * For Rows ---
             * Range range = ws.Rows[6];
             * Range range = ws.Rows["2:4"];
             * 
             * For Columns ---
             * Range range = ws.Columns["B"];
             * Range range = ws.Columns[3];
             * Range range = ws.Columns["F:I"];
             * 
             * For Multiple range
             * Range range = ws.Range["7:9,12:12,14:14"];
             * Range range = ws.Range["B:C,E:E,G:G"];
            */

            range = ws.Cells[1, 1]; // default range sets to cell A1

            try
            {
                switch (rangeType)
                {
                    case "Range": range = ws.Range[rangeDefine]; break;
                    case "Columns": range = ws.Columns[rangeDefine]; break;
                    case "Rows": range = ws.Rows[rangeDefine]; break;
                    default: Console.WriteLine("Invalid input!"); break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return range;
        }

        private Range GetCell (int rowIndex, int colIndex)
        {
            return ws.Cells[rowIndex, colIndex];
        }

        private void LoadExcelSheet(string filePath, string sheetName)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    missValue = System.Reflection.Missing.Value;
                    wb = excelApp.Workbooks.Open(filePath,
                        missValue, missValue, missValue, missValue, missValue, missValue, missValue,
                        missValue, missValue, missValue, missValue, missValue, missValue, missValue);
                    if (sheetName.Trim() == "")
                        ws = wb.Worksheets[1];
                    else
                        ws = wb.Worksheets[sheetName];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                QuitExcelApplication();
                CleanExcelProcess();
            }
        }

        private void CleanExcelProcess()
        {
            excelProcs = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in excelProcs)
            {
                proc.Kill();
            }
        }

        private void QuitExcelApplication()
        {
            excelApp.Application.Quit();
            excelApp.Quit();
        }


    }
}
