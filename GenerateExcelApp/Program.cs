using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GenerateExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateExcel(100);
        }
        static void GenerateExcel(int quantity)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("FullName");
            dt.Columns.Add("DateOfBirth");
            dt.Columns.Add("DateCreated");
            dt.Columns.Add("Amount");
            dt.Columns.Add("Reference");

            for (int i = 1; i <= quantity; i++)
            {
                int Ref = RandomNumber(1, 200000);
                int Amount = RandomNumber(9, 10000);
                string FullName = RandomName();
                DateTime DateCreated = DateTime.Now;
                DateTime DateOfBirth = DateTime.Now.AddYears(-18 - RandomNumber(0, 60)).AddMonths(-RandomNumber(0, 12)).AddDays(-RandomNumber(0, 28));
                
                dt.Rows.Add(FullName, 
                    DateOfBirth.ToString("dd/MM/yyyy"), 
                    DateCreated.ToString(),
                    Amount.ToString(),
                    Ref.ToString());
            }
            ds.Tables.Add(dt);
          
            string fileName = @"C:\temp\Customers_" + quantity.ToString() + ".xlsx";
            //FileStream fs = new FileStream(@"C:\temp\" + fileName, FileMode.OpenOrCreate, FileAccess.Write);
            
            using (XLWorkbook wb = new XLWorkbook())
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    wb.Worksheets.Add(ds.Tables[i], ds.Tables[i].TableName);
                }
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(fileName);
            }
            Console.WriteLine("Done!");
        }

        public static void ExportDataSetToExcel(DataSet ds)
        {
            string AppLocation = "";
            AppLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            AppLocation = AppLocation.Replace("file:\\", "");
            string date = DateTime.Now.ToShortDateString();
            date = date.Replace("/", "_");
            string filepath = AppLocation + "\\ExcelFiles\\" + "RECEIPTS_COMPARISON_" + date + ".xlsx";

            using (XLWorkbook wb = new XLWorkbook())
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    wb.Worksheets.Add(ds.Tables[i], ds.Tables[i].TableName);
                }
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(filepath);
            }
        }


        static string RandomName()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(RandomString(1, false));
            builder.Append(RandomString(9, true));
            return builder.ToString();
        }

        // Generate a random string of a given size
        static string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random(Guid.NewGuid().GetHashCode());
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            if (lowerCase)
                return builder.ToString().ToLower();
            return builder.ToString();
        }
        // Generate a random number between two numbers
        static int RandomNumber(int min, int max)
        {
            Random random = new Random(Guid.NewGuid().GetHashCode());
            return random.Next(min, max);
        }
    }
}
