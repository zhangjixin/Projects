using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace SIRegressionReports
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\jixzha\Documents\Visual Studio 2015\Projects\2016-11-10";
            string excelFileName = @"SIRegressionTestReport.xlsx";
            var testLST = myXMLReader.xmlReader(filePath);
            
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                Console.WriteLine("ERROR: Excel is not available for now.");
                Console.Read();
                return;
            }
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
            int i = 1, j = 1;
            xlWorkSheet.Cells[1, 1] = "testName";
            xlWorkSheet.Cells[1, 2] = "outCome";
            xlWorkSheet.Cells[1, 3] = "startTime";
            xlWorkSheet.Cells[1, 4] = "endTime";
            xlWorkSheet.Cells[1, 5] = "duration";
            foreach (var xlData in testLST)
            {
                j = 1;
                i += 1;
                xlWorkSheet.Cells[i, j++] = xlData.testName;
                xlWorkSheet.Cells[i, j++] = xlData.outcom;
                xlWorkSheet.Cells[i, j++] = xlData.startTime;
                xlWorkSheet.Cells[i, j++] = xlData.endTime;
                xlWorkSheet.Cells[i, j++] = xlData.duration;
            }
            xlWorkBook.SaveAs(excelFileName);
            xlWorkBook.Close();
            xlApp.Quit();
            Console.WriteLine("Press any key to exit...");
            Console.Read();            
        }
    }
}
