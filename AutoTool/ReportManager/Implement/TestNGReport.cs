using AutoTool.ReportManager.Interface;
using AutoTool.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace AutoTool.ReportManager.Implement
{
    public class TestNGReport : IReport
    {
        private const string REPORT_DATE_FORMAT = "yyyy-MM-dd'T'HH:mm:ss'Z'";

        public Dictionary<string, string[]> CollectDataFromFile(string resultPath, string[] ignoreList, string filterFile, string filterTestSuite)
        {
            try
            {
                dynamic results = new Dictionary<string, string[]>();

                AnalyzeContent(ref results, resultPath, ignoreList, filterFile, filterTestSuite);

                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<KeyValuePair<string, string[]>> CollectHistoryDataFromFile(string resultPath, string[] ignoreList, string filterFile, string filterTestSuite)
        {
            try
            {
                dynamic results = new List<KeyValuePair<string, string[]>>();

                AnalyzeContent(ref results, resultPath, ignoreList, filterFile, filterTestSuite);

                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void AnalyzeContent(ref dynamic results, string resultPath, string[] ignoreList, string filterFile, string filterTestSuite)
        {
            DirectoryInfo dir = new DirectoryInfo(resultPath);

            if (String.IsNullOrEmpty(filterFile))
            {
                filterFile = "*";
            }

            // Get all files in directory with matched prefix name
            FileInfo[] files = dir.GetFiles($"{filterFile}.xml");

            if (files.Count() == 0)
            {
                throw new Exception($"There is no file with name \"{filterFile}.xml\" in folder \"{resultPath}\"");
            }

            XmlDocument doc = new XmlDocument();

            // Loop all files and get data base on defined keywords
            foreach (FileInfo file in files)
            {
                doc.Load(file.FullName);
                XmlNode root = doc.DocumentElement;

                if (root.Name == "testng-results")
                {
                    var nodeList = root.SelectNodes($".//test[contains(@name, '{filterTestSuite}')]//test-method");

                    for (int i = 0; i < nodeList.Count; i++)
                    {
                        if (nodeList[i].Attributes["is-config"] == null)
                        {
                            string testID = nodeList[i].Attributes["name"].Value;
                            string startTime = nodeList[i].Attributes["started-at"].Value;
                            string endTime = nodeList[i].Attributes["finished-at"].Value;
                            string status = Helper.UpperFirstCharacter(nodeList[i].Attributes["status"].Value) + "ed";
                            string message = "";

                            // Get error/warning message if test case is not pass
                            if (status != "Passed")
                            {
                                var exception = nodeList[i].SelectNodes(".//exception");
                                if (exception.Count != 0)
                                {
                                    message += exception[0].Attributes["class"].Value;
                                    if (exception[0].SelectNodes(".//message").Count > 0)
                                    {
                                        message += " : " + exception[0].SelectNodes(".//message")[0].InnerText;
                                    }
                                }
                            }

                            // Ignore all test cases in ignore list
                            if (!Array.Exists(ignoreList, E => E == testID))
                            {
                                // If get only one data for each test case, compare and get data has latest end time
                                // Else get all raw data for each test case (one test case id can be run many times with different test result)
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
    }
}
