using AutoTool.ReportManager.Implement;
using AutoTool.ReportManager.Interface;

namespace AutoTool.ReportManager
{
    public class ReportFactory
    {
        public static IReport GetReport(string reportType)
        {
            IReport report;

            switch(reportType)
            {
                case "ExtentReport":
                    report = new ExtentReport();
                    break;
                case "TestNGReport":
                    report = new TestNGReport();
                    break;
                case "AllureReport":
                    report = new AllureReport();
                    break;
                default:
                    report = null;
                    break;
            }

            return report;
        }
    }
}
