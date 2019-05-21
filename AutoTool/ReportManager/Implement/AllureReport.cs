using AutoTool.ReportManager.Interface;
using AutoTool.Utility;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoTool.ReportManager.Implement
{
    public class AllureReport : IReport
    {
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
            DirectoryInfo dir = new DirectoryInfo(resultPath);

            if (String.IsNullOrEmpty(filterFile))
            {
                filterFile = "*";
            }

            // Get all files in directory with matched prefix name
            FileInfo[] files = dir.GetFiles($"{filterFile}.json");

            if (files.Count() == 0)
            {
                throw new Exception($"There is no file with name \"{filterFile}.json\" in folder \"{resultPath}\"");
            }

            // Loop all files and get data base on defined keywords
            foreach (FileInfo file in files)
            {
                StreamReader rd = new StreamReader(file.FullName);
                var json = rd.ReadToEnd();
                var jobj = JObject.Parse(json);
                string testID = (string)jobj["name"];
                string startTime = (string)jobj["start"];
                string endTime = (string)jobj["stop"];
                string status = Helper.UpperFirstCharacter((string)jobj["status"]);
                string message = "";

                // Get error/warning message if test case is not pass
                if (status != "Passed")
                {
                    if (jobj["statusDetails"] != null && jobj["statusDetails"]["message"] != null)
                    {
                        message = (string)jobj["statusDetails"]["message"];
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
                            DateTime storedDate = Helper.ConvertTimestamp(results[testID][Constant.TEST_END_TIME_INDEX]);
                            DateTime currentDate = Helper.ConvertTimestamp(endTime);
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
                        results.Add(new KeyValuePair<string, string[]>(testID, new string[] { Helper.ConvertTimestamp(startTime).ToString(), Helper.ConvertTimestamp(endTime).ToString(), status, message }));
                    }
                }
            }
        }
    }
}
