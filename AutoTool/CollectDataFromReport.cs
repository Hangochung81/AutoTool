using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoTool
{
    public class CollectDataFromReport
    {
        Microsoft.Office.Interop.Excel.Application oXL = null;
        Microsoft.Office.Interop.Excel._Workbook oWB = null;
        Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
        public string dateFormat = "M/d/yyyy h:mm:ss tt";

        private Dictionary<string, string[]> CollectDataFromHtmlFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                Dictionary<string, string[]> results = new Dictionary<string, string[]>();

                WebClient webClient = new WebClient();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.html");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no html file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    var filePath = "file:///" + resultPath.Replace("\\", "/") + "/" + file.Name;
                    string page = webClient.DownloadString(filePath);
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    var tables = doc.DocumentNode.SelectNodes("//div[@class='test-heading']").ToList();
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        var table = doc.DocumentNode.SelectNodes($"(//div[@class='test-heading'])[{i}]").ToList();
                        if (table[0].ChildNodes[1].InnerText.Trim() == "Pre-condition") continue;
                        if (table[0].ChildNodes[1].InnerText.Trim() == "Test Summary") break;
                        string testID = table[0].ChildNodes[1].InnerText.Trim();
                        string executeDate = table[0].ChildNodes[3].InnerText.Trim();
                        string executeResult = UpperFirstCharacter(table[0].ChildNodes[5].InnerText.Trim()) + "ed";

                        if (!Array.Exists(ignoreList, E => E == testID))
                        {
                            if (results.ContainsKey(testID))
                            {
                                DateTime storedDate = DateTime.ParseExact(results[testID][0].ToString(), dateFormat, null);
                                DateTime currentDate = DateTime.ParseExact(executeDate, dateFormat, null);
                                if (storedDate < currentDate)
                                {
                                    results[testID] = new string[] { executeDate, executeResult };
                                }
                            }
                            else
                            {
                                results[testID] = new string[] { executeDate, executeResult };
                            }
                        }
                    }

                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private List<KeyValuePair<string, string[]>> CollectRawDataFromHtmlFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                var results = new List<KeyValuePair<string, string[]>>();

                WebClient webClient = new WebClient();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.html");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no html file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    var filePath = "file:///" + resultPath.Replace("\\", "/") + "/" + file.Name;
                    string page = webClient.DownloadString(filePath);
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    var tables = doc.DocumentNode.SelectNodes("//div[@class='test-heading']").ToList();
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        var table = doc.DocumentNode.SelectNodes($"(//div[@class='test-heading'])[{i}]").ToList();
                        if (table[0].ChildNodes[1].InnerText.Trim() == "Pre-condition") continue;
                        if (table[0].ChildNodes[1].InnerText.Trim() == "Test Summary") break;
                        string testID = table[0].ChildNodes[1].InnerText.Trim();
                        string executeDate = table[0].ChildNodes[3].InnerText.Trim();
                        string executeResult = UpperFirstCharacter(table[0].ChildNodes[5].InnerText.Trim()) + "ed";

                        results.Add(new KeyValuePair<string, string[]>(testID, new string[] { executeDate, executeResult }));
                    }

                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private Dictionary<string, string[]> ColectDataFromXmlFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                Dictionary<string, string[]> results = new Dictionary<string, string[]>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.xml");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no xml file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    DataSet ds = ConvertXMLtoDataset(file.FullName);
                    if (ds.Tables[0].TableName == "testng-results")
                    {
                        for (int i = 0; i < ds.Tables[4].Rows.Count; i++)
                        {
                            string testID = ds.Tables[4].Rows[i].ItemArray[4].ToString().Trim();
                            string executeDate = ds.Tables[4].Rows[i].ItemArray[8].ToString().Trim();
                            string executeResult = UpperFirstCharacter(ds.Tables[4].Rows[i].ItemArray[2].ToString().Trim()) + "ed";

                            if (!Array.Exists(ignoreList, E => E == testID))
                            {
                                if (results.ContainsKey(testID))
                                {
                                    DateTime storedDate = DateTime.ParseExact(results[testID][0].ToString(), dateFormat, null);
                                    DateTime currentDate = DateTime.ParseExact(executeDate, dateFormat, null);
                                    if (storedDate < currentDate)
                                    {
                                        results[testID] = new string[] { executeDate, executeResult };
                                    }
                                }
                                else
                                {
                                    results[testID] = new string[] { executeDate, executeResult };
                                }
                            }
                        }
                    }
                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        private List<KeyValuePair<string, string[]>> ColectRawDataFromXmlFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                var results = new List<KeyValuePair<string, string[]>>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.xml");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no xml file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    DataSet ds = ConvertXMLtoDataset(file.FullName);
                    if (ds.Tables[0].TableName == "testng-results")
                    {
                        for (int i = 0; i < ds.Tables[4].Rows.Count; i++)
                        {
                            string testID = ds.Tables[4].Rows[i].ItemArray[4].ToString().Trim();
                            string executeDate = ds.Tables[4].Rows[i].ItemArray[8].ToString().Trim();
                            string executeResult = UpperFirstCharacter(ds.Tables[4].Rows[i].ItemArray[2].ToString().Trim()) + "ed";

                            if (!Array.Exists(ignoreList, E => E == testID))
                                results.Add(new KeyValuePair<string, string[]>(testID, new string[] { executeDate, executeResult }));
                        }
                    }
                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private Dictionary<string, string[]> ColectDataFromJsonFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                Dictionary<string, string[]> results = new Dictionary<string, string[]>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.json");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no JsonFile file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    StreamReader rd = new StreamReader(file.FullName);
                    var json = rd.ReadToEnd();
                    var jobj = JObject.Parse(json);
                    string testID = (string)jobj["name"];
                    string executeDate = (string)jobj["start"];
                    string executeResult = UpperFirstCharacter((string)jobj["status"]);
                    if (!Array.Exists(ignoreList, E => E == testID))
                    {
                        if (results.ContainsKey(testID))
                        {
                            DateTime storedDate = ConvertTimestamp(results[testID][0]);
                            DateTime currentDate = ConvertTimestamp(executeDate);
                            if (storedDate < currentDate)
                            {
                                results[testID] = new string[] { executeDate, executeResult };
                            }
                        }
                        else
                        {
                            results[testID] = new string[] { executeDate, executeResult };
                        }
                    }
                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private List<KeyValuePair<string, string[]>> ColectRawDataFromJsonFile(string resultPath, string[] ignoreList, string filterFile = null)
        {
            try
            {
                List<KeyValuePair<string, string[]>> results = new List<KeyValuePair<string, string[]>>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] files = dir.GetFiles($"*{filterFile}.json");

                if (files.Count() == 0)
                {
                    throw new Exception("There is no JsonFile file in folder: " + resultPath);
                }

                foreach (FileInfo file in files)
                {
                    StreamReader rd = new StreamReader(file.FullName);
                    var json = rd.ReadToEnd();
                    var jobj = JObject.Parse(json);
                    string testID = (string)jobj["name"];
                    string executeDate = ConvertTimestamp((string)jobj["start"]).ToString();
                    string executeResult = UpperFirstCharacter((string)jobj["status"]);
                    
                    if (!Array.Exists(ignoreList, E => E == testID))
                        results.Add(new KeyValuePair<string, string[]>(testID, new string[] { executeDate, executeResult }));
                }
                return results;
            }
            catch (Exception)
            {
                throw new Exception("There is no JsonFile match with filter name" + filterFile);
            }
        }



        private DateTime ConvertTimestamp(string timestamp)
        {
            
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            dateTime = dateTime.AddMilliseconds(double.Parse(timestamp));
            return dateTime;
        }

        private DataSet ConvertXMLtoDataset(string xmlPathFile)
        {
            //string filePath = string.Empty;
            DataSet objDataSet = new DataSet();
            objDataSet.ReadXml(xmlPathFile);
            return objDataSet;
        }

        public void UpdateExcel(ReportInfo report, TemplateInfo template)
        {
            try
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                dateFormat = report.DateTimeFormat;

                Dictionary<string, string[]> listResult;
                if (report.ReportType == "html")
                {
                    listResult = CollectDataFromHtmlFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
                }
                else if (report.ReportType == "json")
                {
                    listResult = ColectDataFromJsonFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
                }
                else
                {
                    listResult = ColectDataFromXmlFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
                }

                oWB = oXL.Workbooks.Open(report.TargetPath);

                oSheet = String.IsNullOrEmpty(report.SheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[report.SheetName];

                int statusColumn = GenerateReportDate(report, template);
                int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                int testCaseColumn = TitleToNumber(template.TestCaseColumnName);

                for (int i = template.FillableRowStartIndex; i <= lastRow; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[i, testCaseColumn].Text)) continue;

                    string tcVal = oSheet.Cells[i, testCaseColumn].Text.Trim();

                    if (IsTargetTestCase(report.TestCaseList, tcVal) && listResult.ContainsKey(tcVal))
                    {
                        oSheet.Cells[i, statusColumn] = listResult[tcVal][1];
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
            dateFormat = report.DateTimeFormat;

            List<KeyValuePair<string, string[]>> listResult;
            if (report.ReportType == "html")
            {
                listResult = CollectRawDataFromHtmlFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
            }
            else if (report.ReportType == "json")
            {
                listResult = ColectRawDataFromJsonFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
            }
            else
            {
                listResult = ColectRawDataFromXmlFile(report.ResultPath, report.IgnoreTestCaseList, report.FilterFile);
            }

            return listResult.OrderBy(x => x.Key).ToList();
        }

        private Boolean IsTargetTestCase(string[] targetList, string testName)
        {
            if (targetList != null && !targetList.Contains(testName))
            {
                return false;
            }

            return true;
        }

        private int GenerateReportDate(ReportInfo report, TemplateInfo template)
        {
            int targetDateCol = -1;
            int targetStatusCol = -1;
            int lastRow = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int lastColumn = oSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int fillableColumnStartIndex = TitleToNumber(template.FillableColumnStartName);

            for (int i = fillableColumnStartIndex; i <= lastColumn; i++)
            {
                if (String.IsNullOrEmpty(oSheet.Cells[template.DateRowIndex, i].Text)) continue;
                DateTime rawValue = DateTime.FromOADate(oSheet.Cells[template.DateRowIndex, i].Value2);

                if (rawValue.Date == report.ReportDate.Date)
                {
                    targetDateCol = i;
                    targetStatusCol = i + template.StatusColumnIndexPerDate - 1;
                    break;
                }
            }

            if (targetDateCol == -1)
            {
                // Insert copied Range.
                string startColumnName = template.FillableColumnStartName;
                string endColumnName = NumberToTitle(fillableColumnStartIndex + template.ColumnNumberPerDate - 1);
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
            }

            if (targetStatusCol == -1)
            {
                throw new Exception("Cannot find Status column in Date : " + report.ReportDate.Date.ToString());
            }

            return targetStatusCol;
        }

        private int TitleToNumber(string columnName)
        {
            int result = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                result += columnName[i] - 'A' + 1;
            }
            return result;
        }

        private string NumberToTitle(int columnNumber)
        {
            // To store result (Excel column name) 
            StringBuilder columnName = new StringBuilder();

            while (columnNumber > 0)
            {
                // Find remainder 
                int rem = columnNumber % 26;

                // If remainder is 0, then a  
                // 'Z' must be there in output 
                if (rem == 0)
                {
                    columnName.Append("Z");
                    columnNumber = (columnNumber / 26) - 1;
                }
                else // If remainder is non-zero 
                {
                    columnName.Append((char)((rem - 1) + 'A'));
                    columnNumber = columnNumber / 26;
                }
            }

            return columnName.ToString();
        }

        public List<string> GetSheetName(string excelPath)
        {
            var sheetNames = new List<string>();

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(excelPath);
                
                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in oXL.Worksheets)
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
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];
                int testCaseColumn = TitleToNumber(testCaseColumnName);
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

        public string UpperFirstCharacter(string value)
        {
            return value.First().ToString().ToUpper() + value.Substring(1).ToLower();
        }
    }
}
