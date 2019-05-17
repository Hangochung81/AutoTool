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
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                if (template.FillStatus == false && template.FillDetail == false)
                {
                    throw new Exception("There is no column was chosen to fill data. Please check template info again.");
                }

                // Create Report object
                var collectedReport = ReportFactory.GetReport(report.ReportType);
                var listResult = collectedReport.CollectDataFromFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile, report.DateTimeFormat);

                oWB = oXL.Workbooks.Open(report.TargetPath);

                oSheet = String.IsNullOrEmpty(report.SheetName) ? (_Worksheet)oWB.ActiveSheet : (_Worksheet)oWB.Worksheets[report.SheetName];

                var columnList = GenerateReportDate(report, template);
                int statusColumn = columnList[0];
                int detailColumn = columnList[1];

                int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                int testCaseColumn = Helper.TitleToNumber(template.TestCaseColumnName);

                for (int i = template.FillableRowStartIndex; i <= lastRow; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[i, testCaseColumn].Text)) continue;

                    string tcVal = oSheet.Cells[i, testCaseColumn].Text.Trim();

                    if (IsTargetTestCase(report.TestCaseList, tcVal) && listResult.ContainsKey(tcVal))
                    {
                        if (template.FillStatus)
                        {
                            oSheet.Cells[i, statusColumn] = listResult[tcVal][2];
                        }
                        if (template.FillDetail)
                        {
                            oSheet.Cells[i, detailColumn] = listResult[tcVal][3];
                        }
                    }
                }

                oWB.Save();

                stopwatch.Stop();

                MessageBox.Show(String.Format("Completed in {0} seconds", stopwatch.ElapsedMilliseconds/1000), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            var listResult = collectedReport.CollectHistoryDataFromFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);

            return listResult.OrderBy(x => x.Key).ToList();
        }

        private bool IsTargetTestCase(string[] targetList, string testName)
        {
            if (targetList != null && !targetList.Contains(testName))
            {
                return false;
            }

            return true;
        }

        private int[] GenerateReportDate(ReportInfo report, TemplateInfo template)
        {
            int targetDateCol = -1;
            int targetStatusCol = -1;
            int targetDetailCol = -1;
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
                    targetStatusCol = i + template.StatusColumnIndexPerDate - 1;
                    targetDetailCol = i + template.DetailColumnIndexPerDate - 1;
                    break;
                }
            }

            if (targetDateCol == -1)
            {
                // Insert copied Range.
                string startColumnName = template.FillableColumnStartName;
                string endColumnName = Helper.NumberToTitle(fillableColumnStartIndex + template.ColumnNumberPerDate - 1);
                Range range = oSheet.Columns[startColumnName + ":" + endColumnName];
                range.Copy();
                range.Insert();

                // Update Date value.
                range = oSheet.Cells[template.DateRowIndex, fillableColumnStartIndex];
                range.Value2 = report.ReportDate.Date.ToOADate();

                // Clear old data. Values below are based on Excel template.
                range = oSheet.Range[oSheet.Cells[template.FillableRowStartIndex, fillableColumnStartIndex], oSheet.Cells[lastRow, fillableColumnStartIndex + template.ColumnNumberPerDate - 1]];
                range.Value = "";

                targetStatusCol = fillableColumnStartIndex + template.StatusColumnIndexPerDate - 1;
                targetDetailCol = fillableColumnStartIndex + template.DetailColumnIndexPerDate - 1;
            }

            return new int[] { targetStatusCol, targetDetailCol };
        }

        public List<string> GetSheetName(string excelPath)
        {
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
}
