using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoTool
{
    public class CollectDataFromHTML
    {
        Microsoft.Office.Interop.Excel.Application oXL = null;
        Microsoft.Office.Interop.Excel._Workbook oWB = null;
        Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

        
        //}

        public Dictionary<string, string[]> collectData(string resultPath)
        {

            try
            {

              
                Dictionary<string, string[]> results = new Dictionary<string, string[]>();

                WebClient webClient = new WebClient();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] Files = dir.GetFiles("*.html");

                foreach (FileInfo file in Files)
                {
                    var filePath = "file:///" + resultPath.Replace("\\", "/") + "/" + file.Name;
                    string page = webClient.DownloadString(filePath);
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    var table = doc.DocumentNode.SelectNodes("(//div[@class='test-heading'])[2]").ToList();
                    string testID = table[0].ChildNodes[1].InnerText.Trim();
                    string executeDate = table[0].ChildNodes[3].InnerText.Trim();
                    string executeResult = table[0].ChildNodes[5].InnerText.Trim() + "ed";

                    if (results.ContainsKey(testID))
                    {
                        if (results[testID][0].ToString().CompareTo(executeDate) == -1)
                        {
                            results[testID] = new string[] { executeDate, executeResult };
                        }
                    }
                    else
                    {
                        results[testID] = new string[] { executeDate, executeResult };
                    }
                }
                return results;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public List<string> GetSheetName(string excelPath)
        {
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open(excelPath);
            var sheetNames = new List<string>();
            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in oXL.Worksheets)
            {
                if (worksheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    sheetNames.Add(worksheet.Name);
                }
            }
            oWB.Close();
            return sheetNames;
        }

        public void UpdateExcel(string sheetName, string resultPath, string excelPath, DateTime date)
        {
            oWB = oXL.Workbooks.Open(excelPath);
            Dictionary<string, string[]> listResult = collectData(resultPath);
            int dateRow = 7;
            int dateCol = 6;
            int statusCol = 0;
            int tcNameRow = 10;
            int tcNameCol = 2;
            try
            {
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];
               
                for (int i = dateCol; i < oSheet.Columns.Count; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[dateRow, i].Text)) continue;
                    DateTime rawValue = DateTime.FromOADate(oSheet.Cells[dateRow, i].Value2);

                    if (rawValue.Date == date.Date)
                    {
                        statusCol = i;
                        break;
                    }
                    
                }
                for (int i = tcNameRow; i < oSheet.Rows.Count; i++)
                {
                    if (String.IsNullOrEmpty(oSheet.Cells[i, tcNameCol].Text))
                    {
                        break;
                    }
                    string tcVal = oSheet.Cells[i, tcNameCol].Text.Trim();
                    if (listResult.ContainsKey(tcVal))
                    {
                        oSheet.Cells[i, statusCol] = listResult[tcVal][1];
                    }

                }
               
                oWB.Save();

                MessageBox.Show("Done!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                    oWB.Close();
            }
        }

        public void InputDataFromExcelFile()
        {
          
        }
    }
}
