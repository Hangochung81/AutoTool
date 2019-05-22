using AutoTool.Entity;
using AutoTool.Utility;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace AutoTool.ReportManager
{
    public class ReportUtils
    {
        Microsoft.Office.Interop.Excel.Application oXL = null;
        _Workbook oWB = null;
        _Worksheet oSheet = null;

        public void UpdateExcel(ReportInfo report, TemplateInfo template)
        {
            try
            {
                // Stop process if there is no sub title was chosen to fill data
                if (template.SubTitleColumnIndexList.Count == 0)
                {
                    throw new Exception("There is no column was chosen to fill data. Please check template info again.");
                }

                // Create Report object
                var collectedReport = ReportFactory.GetReport(report.ReportType);
                var listResult = collectedReport.CollectDataFromFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);

                // Open Excel file
                oWB = oXL.Workbooks.Open(report.TargetPath);

                oSheet = String.IsNullOrEmpty(report.SheetName) ? (_Worksheet)oWB.ActiveSheet : (_Worksheet)oWB.Worksheets[report.SheetName];

                // Get location of target date area, if target date not exist, new target date area will be generated
                int firstSubTitleIndex = GenerateReportDate(report, template);

                int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                int testCaseColumn = Helper.TitleToNumber(template.TestCaseColumnName);

                // Loop all test case to fill data into chosen sub titles
                for (int i = template.FillableRowStartIndex; i <= lastRow; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[i, testCaseColumn].Text)) continue;

                    string tcVal = oSheet.Cells[i, testCaseColumn].Text.Trim();

                    // Only fill data if the test case belong to target test case list
                    if (IsTargetTestCase(report.TestCaseList, tcVal) && listResult.ContainsKey(tcVal))
                    {
                        // Loop chosen sub title list and fill data
                        for (int j = 0; j < template.SubTitleColumnIndexList.Count; j++)
                        {
                            int subTitleType = template.SubTitleColumnIndexList[j].Key;
                            int subTitleIndex = template.SubTitleColumnIndexList[j].Value;
                            oSheet.Cells[i, firstSubTitleIndex + subTitleIndex - 1] = listResult[tcVal][subTitleType];
                        }
                    }
                }

                oWB.Save();
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close();
                    oWB = null;
                }
            }
        }

        public List<KeyValuePair<string, string[]>> GetCollectedData(ReportInfo report)
        {
            // Create Report object
            var collectedReport = ReportFactory.GetReport(report.ReportType);
            // Get all data from reports
            var listResult = collectedReport.CollectHistoryDataFromFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);

            // sort list by key (test case id) and return
            return listResult.OrderBy(x => x.Key).ToList();
        }

        private bool IsTargetTestCase(string[] targetList, string testName)
        {
            // Check if the test case belong to list or not
            if (targetList != null && !targetList.Contains(testName))
            {
                return false;
            }

            return true;
        }

        private int GenerateReportDate(ReportInfo report, TemplateInfo template)
        {
            // Find location (column name) of target date
            int targetDateCol = -1;
            int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int lastColumn = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int fillableColumnStartIndex = Helper.TitleToNumber(template.FillableColumnStartName);

            for (int i = fillableColumnStartIndex; i <= lastColumn; i++)
            {
                if (String.IsNullOrEmpty(oSheet.Cells[template.DateRowIndex, i].Text)) continue;
                DateTime rawValue = DateTime.FromOADate(oSheet.Cells[template.DateRowIndex, i].Value2);

                if (rawValue.Date == report.ReportDate.Date)
                {
                    targetDateCol = i;
                    break;
                }
            }

            // If target date was not found, new date area will be generated for it
            if (targetDateCol == -1)
            {
                // Insert copied Range.
                string startColumnName = template.FillableColumnStartName;
                string endColumnName = Helper.NumberToTitle(fillableColumnStartIndex + template.ColumnNumberPerDate - 1);
                Range range = oSheet.Columns[startColumnName + ":" + endColumnName];
                // Copy and paste all template (with value and format) from an old date area
                range.Copy();
                range.Insert();

                // Update Date value.
                range = oSheet.Cells[template.DateRowIndex, fillableColumnStartIndex];
                range.Value2 = report.ReportDate.Date.ToOADate();

                // Clear old data. Values below are based on Excel template.
                range = oSheet.Range[oSheet.Cells[template.FillableRowStartIndex, fillableColumnStartIndex], oSheet.Cells[lastRow, fillableColumnStartIndex + template.ColumnNumberPerDate - 1]];
                range.Value = "";

                targetDateCol = fillableColumnStartIndex;
            }

            return targetDateCol;
        }

        public List<string> GetSheetName(string excelPath)
        {
            // Get all sheet names of an Excel file
            var sheetNames = new List<string>();

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(excelPath);
                
                foreach (Worksheet worksheet in oXL.Worksheets)
                {
                    if (worksheet.Visible == XlSheetVisibility.xlSheetVisible)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                }
            }
            catch (Exception)
            {
                sheetNames = new List<string>();
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close();
                    oWB = null;
                }
            }

            return sheetNames;
        }

        public string GetCurrentTestCase(string excelPath, string sheetName, string testCaseColumnName, int testCaseStartIndex)
        {
            // Get all current test case in specific sheet of an Excel file
            string tcList = "";
            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(excelPath);
                oSheet = (_Worksheet)oWB.Worksheets[sheetName];
                int testCaseColumn = Helper.TitleToNumber(testCaseColumnName);
                int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                for (int i = testCaseStartIndex; i <= lastRow; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[i, testCaseColumn].Text)) continue;

                    string tcVal = oSheet.Cells[i, testCaseColumn].Text.Trim();
                    tcList += tcVal + "\r\n";
                }
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close();
                    oWB = null;
                }
            }

            return tcList;
        }

        public void OpenExcel(string excelPath, string sheetName)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(excelPath);
            excelApp.Worksheets[sheetName].Activate();
        }
    }
}
