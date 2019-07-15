using System;

namespace AutoTool.Entity
{
    public class ReportInfo
    {
        public string SheetName { get; set; }
        public string ResultPath { get; set; }
        public string ReportType { get; set; }
        public string TestSuiteName { get; set; }
        public string FilterFile { get; set; }
        public string TargetPath { get; set; }
        public DateTime ReportDate { get; set; }
        public string[] TestCaseList { get; set; }
        public string[] IgnoreTestCaseList { get; set; }
    }
}
