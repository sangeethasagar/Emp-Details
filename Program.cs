using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Emp Details.xlsx";
            string filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString() + "\\files\\"+fileName;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook Wkb = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet WsSht = Wkb.Sheets[1]; // assume it is the first sheet
            Excel.Range rng = WsSht.UsedRange;
    
          
            int RowCount = rng.Rows.Count;
            int ColumnCount = rng.Columns.Count;

            int rCount, cCount;

            //Loop through each data in the excel sheet

            for (rCount = 1; rCount <= RowCount; rCount++)
            {
                for (cCount = 1; cCount <= ColumnCount; cCount++)
                {
                    Console.WriteLine("/n Coulmn Number: " + cCount + "--> " + (rng.Cells[rCount, cCount] as Excel.Range).Value2);
                }
            }
            Console.ReadKey();

        }
    }
}
