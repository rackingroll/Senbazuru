using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    class FrameFinderRunner
    {
        static void Main(string[] args)
        {
            /*
            Application app = new Application();
            Workbook wb = app.Workbooks.Open(@"D:\dev\test.xlsx");
            Worksheet ws = wb.Sheets[1];
            Range cell = ws.Cells[1, 1];
            Console.WriteLine(cell.NumberFormat);
            wb.Close();
            */
            FrameFinder.ProcessEachTable();
            Console.ReadKey();
        }
    }
}
