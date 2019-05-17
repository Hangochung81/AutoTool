using AutoTool.ReportManager.Interface;
using AutoTool.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace AutoTool.ReportManager.Implement
{
    public class TestNGReport : IReport
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

                FileInfo[] files = dir.GetFiles($"{filterFile}.xml");

                if (files.Count() == 0)
                {
                    throw new Exception($"There is no file with name \"{filterFile}.xml\" in folder \"{resultPath}\"");
                }

                foreach (FileInfo file in files)
                {
                    DataSet ds = Helper.ConvertXMLtoDataset(file.FullName);
                    if (ds.Tables[0].TableName == "testng-results")
                    {
                        for (int i = 0; i < ds.Tables[4].Rows.Count; i++)
                        {
                            string testID = ds.Tables[4].Rows[i].ItemArray[4].ToString().Trim();
                            string executeDate = ds.Tables[4].Rows[i].ItemArray[8].ToString().Trim();
                            string executeResult = Helper.UpperFirstCharacter(ds.Tables[4].Rows[i].ItemArray[2].ToString().Trim()) + "ed";

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

        public List<KeyValuePair<string, string[]>> CollectHistoryDataFromFile(string resultPath, string[] ignoreList, string filterFile)
        {
            try
            {
                var results = new List<KeyValuePair<string, string[]>>();

                DirectoryInfo dir = new DirectoryInfo(resultPath);

                if (String.IsNullOrEmpty(filterFile))
                {
                    filterFile = "*";
                }

                FileInfo[] files = dir.GetFiles($"{filterFile}.xml");

                if (files.Count() == 0)
                {
                    throw new Exception($"There is no file with name \"{filterFile}.xml\" in folder \"{resultPath}\"");
                }

                foreach (FileInfo file in files)
                {
                    DataSet ds = Helper.ConvertXMLtoDataset(file.FullName);
                    if (ds.Tables[0].TableName == "testng-results")
                    {
                        for (int i = 0; i < ds.Tables[4].Rows.Count; i++)
                        {
                            string testID = ds.Tables[4].Rows[i].ItemArray[4].ToString().Trim();
                            string executeDate = ds.Tables[4].Rows[i].ItemArray[8].ToString().Trim();
                            string executeResult = Helper.UpperFirstCharacter(ds.Tables[4].Rows[i].ItemArray[2].ToString().Trim()) + "ed";

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
    }
}
