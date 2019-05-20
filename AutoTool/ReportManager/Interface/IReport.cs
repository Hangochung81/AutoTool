﻿using System.Collections.Generic;

namespace AutoTool.ReportManager.Interface
{
    public interface IReport
    {
        Dictionary<string, string[]> CollectDataFromFile(string resultPath, string[] ignoreList, string filterFile);

        List<KeyValuePair<string, string[]>> CollectHistoryDataFromFile(string resultPath, string[] ignoreList, string filterFile);
    }
}
