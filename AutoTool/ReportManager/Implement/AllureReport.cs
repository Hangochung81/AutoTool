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
        public Dictionary<string, string[]> CollectDataFromFile(string resultPath, string[] ignoreList, string filterFile, string dateFormat)
        {
            try
            {
                Dictionary<string, string[]> results = new Dictionary<string, string[]>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);

                if (String.IsNullOrEmpty(filterFile))
                {
                    filterFile = "*";
                }

                FileInfo[] files = dir.GetFiles($"{filterFile}.json");

                if (files.Count() == 0)
                {
                    throw new Exception($"There is no file with name \"{filterFile}.json\" in folder \"{resultPath}\"");
                }

                foreach (FileInfo file in files)
                {
                    StreamReader rd = new StreamReader(file.FullName);
                    var json = rd.ReadToEnd();
                    var jobj = JObject.Parse(json);
                    string testID = (string)jobj["name"];
                    string startTime = (string)jobj["start"];
                    string endTime = (string)jobj["stop"];
                    string message = "";
                    string status = Helper.UpperFirstCharacter((string)jobj["status"]);
                    if (status != "Passed")
                    {
                        message = (string)jobj["statusDetails"]["message"];
                    }
                    if (!Array.Exists(ignoreList, E => E == testID))
                    {
                        if (results.ContainsKey(testID))
                        {
                            DateTime storedDate = Helper.ConvertTimestamp(results[testID][1]);
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
                }
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
                List<KeyValuePair<string, string[]>> results = new List<KeyValuePair<string, string[]>>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);

                if (String.IsNullOrEmpty(filterFile))
                {
                    filterFile = "*";
                }

                FileInfo[] files = dir.GetFiles($"{filterFile}.json");

                if (files.Count() == 0)
                {
                    throw new Exception($"There is no file with name \"{filterFile}.json\" in folder \"{resultPath}\"");
                }

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
                    if (status != "Passed")
                    {
                        message = (string)jobj["statusDetails"]["message"];
                    }

                    if (!Array.Exists(ignoreList, E => E == testID))
                        results.Add(new KeyValuePair<string, string[]>(testID, new string[] { Helper.ConvertTimestamp(startTime).ToString(), Helper.ConvertTimestamp(endTime).ToString(), status, message }));
                }
                return results;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
