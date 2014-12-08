using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Senbazuru.HirarchicalExtraction
{
    public class ExcelReaderInterop
    {
        Application excelapp;
        public ExcelReaderInterop()
        {
            excelapp = new Application();
 
        }

        public void OpenExcel(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;

            Workbook workbook = excelapp.Workbooks.Open(strFileName);
            ExcelScanIntenal(workbook);

            //
            // Clean up.
            //
            workbook.Close(false, strFileName, null);
            Marshal.ReleaseComObject(workbook);

        }

        /// <summary>
        /// Scan the selected Excel workbook and store the information in the cells
        /// for this workbook in an object[,] array. Then, call another method
        /// to process the data.
        /// </summary>
        private void ExcelScanIntenal(Workbook workBookIn)
        {
            //
            // Get sheet Count and store the number of sheets.
            //
            int numSheets = workBookIn.Sheets.Count;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Worksheet sheet = (Worksheet)workBookIn.Sheets[sheetNum];

                
                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                Range excelRange = sheet.UsedRange;
                
                object[,] valueArray = (object[,])excelRange.get_Value(
                    XlRangeValueDataType.xlRangeValueDefault);
                
                
                //
                // Do something with the data in the array with a custom method.
                //
                for (int i = 1; i < valueArray.Length; i++)
                {
                    Range cell = sheet.Cells[i, 1];

                    if (cell.Font.Bold)
                    {
                        Console.WriteLine("Bold!");
                    }
                    else
                    {
                        Console.WriteLine("Not Bold!");
                    }
                    Console.WriteLine(valueArray[i,1].ToString());
                }
            }
        }
    }
}
