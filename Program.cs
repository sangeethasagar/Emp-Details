using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Emp Details.xlsx";

            string filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString() + "\\files\\" + fileName;
            Excel.Application xlApp = new Excel.Application();

            //Open the Emp Details Excel  workbook
            Excel.Workbook Wkb = xlApp.Workbooks.Open(filePath);

            //Set the Excel sheet name to Emp details
            Excel.Worksheet WsSht = Wkb.Sheets["Emp Details"];

            //Set the range to used range in emp details sheet
            Excel.Range rng = WsSht.UsedRange;

            int RowCount = rng.Rows.Count;
            int ColumnCount = rng.Columns.Count;

            int rCount;

            string format = "{0,5} {1,22} {2,25}" + Environment.NewLine;
            var stringBuilder = new StringBuilder().AppendFormat(format, "", "", "");

            //Loop through each data in the excel sheet
            for (rCount = 1; rCount <= RowCount; rCount++)
            {
                stringBuilder.AppendFormat(format, rng.Cells[rCount, 1].Value2.ToString(), rng.Cells[rCount, 2].Value2.ToString(), rng.Cells[rCount, 3].Value2.ToString());
            }

            Console.WriteLine(stringBuilder.ToString());
            Console.ReadKey();

            //Close the Emp[ details workbook
            Wkb.Close();

            //Quit the excel application
            xlApp.Quit();

        }//End of function
    }//End of class
}//End of namespace
