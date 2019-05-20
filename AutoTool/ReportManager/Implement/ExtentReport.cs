using AutoTool.ReportManager.Interface;
using AutoTool.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace AutoTool.ReportManager.Implement
{
    public class ExtentReport : IReport
    {
        private const string REPORT_DATE_FORMAT = "M/d/yyyy h:mm:ss tt";

        public Dictionary<string, string[]> CollectDataFromFile(string resultPath, string[] ignoreList, string filterFile)
        {
            try
            {
                dynamic results = new Dictionary<string, string[]>();

                AnalyzeContent(ref results, resultPath, ignoreList, filterFile);

                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<KeyValuePair<string, string[]>> CollectHistoryDataFromFile(string resultPath, string[] ignoreList, string filterFile)
        {
            try
            {
                dynamic results = new List<KeyValuePair<string, string[]>>();

                AnalyzeContent(ref results, resultPath, ignoreList, filterFile);

                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void AnalyzeContent(ref dynamic results, string resultPath, string[] ignoreList, string filterFile)
        {
            WebClient webClient = new WebClient();

            DirectoryInfo dir = new DirectoryInfo(resultPath);

            if (String.IsNullOrEmpty(filterFile))
            {
                filterFile = "*";
            }

            FileInfo[] files = dir.GetFiles($"{filterFile}.html");

            if (files.Count() == 0)
            {
                throw new Exception($"There is no file with name \"{filterFile}.html\" in folder \"{resultPath}\"");
            }

            foreach (FileInfo file in files)
            {
                var filePath = "file:///" + resultPath.Replace("\\", "/") + "/" + file.Name;
                string page = webClient.DownloadString(filePath);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(page);

                var tables = doc.DocumentNode.SelectNodes("//ul[@id='test-collection']/li").ToList();
                for (int i = 1; i <= tables.Count; i++)
                {
                    var testNode = doc.DocumentNode.SelectNodes($"(//ul[@id='test-collection']/li)[{i}]").Single();
                    string testID = testNode.SelectNodes(".//span[@class='test-name']").Single().InnerText.Trim();
                    string startTime = testNode.SelectNodes(".//span[@class='label start-time']").Single().InnerText.Trim();
                    string endTime = testNode.SelectNodes(".//span[@class='label end-time']").Single().InnerText.Trim();
                    string status = Helper.UpperFirstCharacter(testNode.SelectNodes(".//div[@class='test-heading']/span[contains(@class,'test-status')]").Single().InnerText.Trim()) + "ed";
                    string message = "";
                    if (status != "Passed")
                    {
                        var lastNode = testNode.SelectNodes("(.//li[contains(@class,'node')])[last()]").Single();
                        if (lastNode.Attributes["status"].Value != "pass")
                        {
                            var lastLog = lastNode.SelectNodes("(.//tr[@class='log'])[last()]").Single();
                            if (lastLog.Attributes["status"].Value != "info" && lastLog.Attributes["status"].Value != "pass")
                            {
                                message = lastLog.SelectNodes(".//td[@class='step-details']").Single().InnerText.Trim();
                            }
                        }
                    }

                    if (!Array.Exists(ignoreList, E => E == testID))
                    {
                        if (results.GetType() == typeof(Dictionary<string, string[]>))
                        {
                            if (results.ContainsKey(testID))
                            {
                                DateTime storedDate = DateTime.ParseExact(results[testID][Constant.TEST_END_TIME_INDEX].ToString(), REPORT_DATE_FORMAT, null);
                                DateTime currentDate = DateTime.ParseExact(endTime, REPORT_DATE_FORMAT, null);
                                if (storedDate < currentDate)
                                {
                                    results[testID] = new string[] { startTime, endTime, status, message };
                                }
                            }
                            else
                            {
                                results[testID] = new string[] { startTime, endTime, status, message };
                            }
                        }
                        else
                        {
                            results.Add(new KeyValuePair<string, string[]>(testID, new string[] { startTime, endTime, status, message }));
                        }
                    }
                }
            }
        }
    }
}
