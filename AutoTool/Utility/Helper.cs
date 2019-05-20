using System;
using System.Data;
using System.Linq;
using System.Text;

namespace AutoTool.Utility
{
    public class Helper
    {
        public static string UpperFirstCharacter(string value)
        {
            if (!String.IsNullOrEmpty(value))
            {
                return value.First().ToString().ToUpper() + value.Substring(1).ToLower();
            }
            else return "";
        }

        public static DataSet ConvertXMLtoDataset(string xmlPathFile)
        {
            //string filePath = string.Empty;
            DataSet objDataSet = new DataSet();
            objDataSet.ReadXml(xmlPathFile);
            return objDataSet;
        }

        public static int TitleToNumber(string columnName)
        {
            int result = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                result += columnName[i] - 'A' + 1;
            }
            return result;
        }

        public static string NumberToTitle(int columnNumber)
        {
            // To store result (Excel column name) 
            StringBuilder columnName = new StringBuilder();

            while (columnNumber > 0)
            {
                // Find remainder 
                int rem = columnNumber % 26;

                // If remainder is 0, then a  
                // 'Z' must be there in output 
                if (rem == 0)
                {
                    columnName.Append("Z");
                    columnNumber = (columnNumber / 26) - 1;
                }
                else // If remainder is non-zero 
                {
                    columnName.Append((char)((rem - 1) + 'A'));
                    columnNumber = columnNumber / 26;
                }
            }

            return columnName.ToString();
        }

        public static DateTime ConvertTimestamp(string timestamp)
        {

            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            dateTime = dateTime.AddMilliseconds(double.Parse(timestamp));
            return dateTime;
        }
    }
}
