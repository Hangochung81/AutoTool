using AutoTool.ReportManager.Implement;
using AutoTool.ReportManager.Interface;
using AutoTool.Utility;

namespace AutoTool.ReportManager
{
    public class ReportFactory
    {
        public static IReport GetReport(string reportType)
        {
            if (reportType == Constant.REPORT_TYPE_EXTENT)
            {
                return new ExtentReport();
            }
            else if (reportType == Constant.REPORT_TYPE_TESTNG)
            {
                return new TestNGReport();
            }
            else if (reportType == Constant.REPORT_TYPE_ALLURE)
            {
                return new AllureReport();
            }
            else return null;
        }
    }
}
