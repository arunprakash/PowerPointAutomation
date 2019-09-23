using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace PowerPointAutomation
{
    class ExcelReusable
    {
        Application xlApp;
        Workbook xlWorkbook;
        _Worksheet xlWorksheet;
        Range xlRange;

        public Range ReadExcel(string fileName)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = new Application();
            xlWorkbook = xlApp.Workbooks.Open(fileName);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            return xlRange;
        }

        public void closeExcel()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
