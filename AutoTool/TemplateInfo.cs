using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTool
{
    public class TemplateInfo
    {
        public int DateRowIndex { get; set; }
        public string TestCaseColumnName { get; set; }
        public string FillableColumnStartName { get; set; }
        public int FillableRowStartIndex { get; set; }
        public int StatusColumnIndexPerDate { get; set; }
        public int ColumnNumberPerDate { get; set; }
        public string DateTimeFormat { get; set; }
    }
}
