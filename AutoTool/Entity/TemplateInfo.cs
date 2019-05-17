namespace AutoTool.Entity
{
    public class TemplateInfo
    {
        public int DateRowIndex { get; set; }
        public string TestCaseColumnName { get; set; }
        public string FillableColumnStartName { get; set; }
        public int FillableRowStartIndex { get; set; }
        public int StatusColumnIndexPerDate { get; set; }
        public int DetailColumnIndexPerDate { get; set; }
        public int ColumnNumberPerDate { get; set; }
        public bool FillStatus { get; set; }
        public bool FillDetail { get; set; }
    }
}
