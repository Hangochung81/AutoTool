using System.Collections.Generic;

namespace AutoTool.Entity
{
    public class TemplateInfo
    {
        public int DateRowIndex { get; set; }
        public string TestCaseColumnName { get; set; }
        public string FillableColumnStartName { get; set; }
        public int FillableRowStartIndex { get; set; }
        public int ColumnNumberPerDate { get; set; }
        public List<KeyValuePair<int, int>> SubTitleColumnIndexList { get; set; }
    }
}
