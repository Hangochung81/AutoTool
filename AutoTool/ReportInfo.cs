using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTool
{
    public class ReportInfo
    {
        public string SheetName { get; set; }
        public string ResultPath { get; set; }
        public string ReportType { get; set; }
        public string TargetPath { get; set; }
        public DateTime ReportDate { get; set; }
        public string[] TestCaseList { get; set; }
    }
}
