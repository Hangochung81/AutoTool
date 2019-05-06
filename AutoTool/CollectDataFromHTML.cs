using System;
using System.Collections.Generic;
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
        public List<string[]> collectData(string resultPath)
        {
            //var filePath = "D:\test.xlsx";
            //var helper = (ExcelHelper)ExcelDriver.getExcelHelper(filePath);
            //helper.UpdateCellValue(filePath, "test", );

            try
            {

                List<string[]> results = new List<string[]>();

                WebClient webClient = new WebClient();

                DirectoryInfo dir = new DirectoryInfo(resultPath);
                FileInfo[] Files = dir.GetFiles("*.html");

                foreach (FileInfo file in Files)
                {
                    var filePath = "file:///C:/temp/testresults/" + file.Name;
                    var fileDemo = "file:///" + resultPath.Replace("\\","/") +"/" + file.Name;
                    string page = webClient.DownloadString(fileDemo);
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    var table = doc.DocumentNode.SelectNodes("(//div[@class='test-heading'])[2]").ToList();
                    string testID = table[0].ChildNodes[1].InnerText.Trim();
                    string executeDate = table[0].ChildNodes[3].InnerText.Trim();
                    string executeResult = table[0].ChildNodes[5].InnerText.Trim();

                    results.Add(new string[3] { testID, executeDate, executeResult });
                }
                return results;
            }
            catch (Exception e)
            {

                throw e;
            }


            //StreamWriter sw = new StreamWriter("D:\\TempResult.txt");

            //foreach (string[] item in results)
            //{
            //    try
            //    {
            //        var line = "";
            //        for (int i = 0; i < item.Count(); i++)
            //        {
            //            line += item[i] + " | ";
            //        }
            //        sw.WriteLine(line);
            //    }
            //    catch (Exception e)
            //    {
            //        throw;
            //    }
            //}

            //sw.Close();
        }

        public void UpdateExcel(string sheetName, string resultPath, string excelPath, string [] columnName)
        {
            List<string[]> listResult = collectData(resultPath);
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(excelPath);
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                int i;
                for (i = 0 ; i < columnName.Length; i++)
                {
                    int col = 0;
                    int row = 1;
                    for (row = 1; row <= oSheet.Rows.Count; row++)
                    {
                        for (int j = 1; j <= 100; j++)
                        {
                            if (oSheet.Cells[row, j].Text == columnName[i])
                            {
                                col += j;
                                break;
                            }
                        }
                        if (col != 0)
                        {
                            break;
                        }
                    } 
                    for (int j = 0; j < listResult.Count; j++)
                    {
                        row ++;
                        oSheet.Cells[row, col] = listResult[j][i];
                    }
                   
                }

                //int a = oSheet.Columns.Count;
                //string abc = oSheet.Cells[1,1].Text;
                //int nameRow = 1;
                //int nameCol = 0;
                //int dateRow = 1;
                //int dateCol = 0;
                //int resultRow = 1;
                //int resultCol = 0;
                ////string columnname = oSheet.Columns[1].Text;
                //for (int i = 1; i <= oSheet.Columns.Count; i++)
                //{
                //    if (oSheet.Cells[1, i].Text == "Name")
                //    {
                //        nameCol += i;
                //        break;
                //    }
                //}
                //for (int i = 0; i < listResult.Count; i++)
                //{
                //    nameRow += 1;
                //    oSheet.Cells[nameRow, nameCol] = listResult[i][0];
                //}

                //for (int i = 1; i <= oSheet.Columns.Count; i++)
                //{
                //    if (oSheet.Cells[1, i].Text == "Date")
                //    {
                //        dateCol += i;
                //        break;
                //    }
                //}
                //for (int i = 0; i < listResult.Count; i++)
                //{
                //    dateRow += 1;
                //    oSheet.Cells[dateRow, dateCol] = listResult[i][1];
                //}

                //for (int i = 1; i <= oSheet.Columns.Count; i++)
                //{
                //    if (oSheet.Cells[1, i].Text == "Result")
                //    {
                //        resultCol += i;
                //        break;
                //    }
                //}
                //for (int i = 0; i < listResult.Count; i++)
                //{
                //    resultRow += 1;
                //    oSheet.Cells[resultRow, resultCol] = listResult[i][2];
                //}


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
